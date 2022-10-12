/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: t; c-basic-offset: 4 -*- */
/* libwps
 * Version: MPL 2.0 / LGPLv2.1+
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *
 * Major Contributor(s):
 * Copyright (C) 2006, 2007 Andrew Ziem
 * Copyright (C) 2004 Marc Maurer (uwog@uwog.net)
 * Copyright (C) 2004-2006 Fridrich Strba (fridrich.strba@bluewin.ch)
 *
 * For minor contributions see the git repository.
 *
 * Alternatively, the contents of this file may be used under the terms
 * of the GNU Lesser General Public License Version 2.1 or later
 * (LGPLv2.1+), in which case the provisions of the LGPLv2.1+ are
 * applicable instead of those above.
 */

#include <stdlib.h>
#include <string.h>

#include <cmath>
#include <sstream>
#include <limits>
#include <map>
#include <stack>
#include <utility>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WKSChart.h"
#include "WKSContentListener.h"
#include "WPSEntry.h"
#include "WPSFont.h"
#include "WPSPosition.h"
#include "WPSStream.h"

#include "Lotus.h"
#include "LotusStyleManager.h"

#include "LotusChart.h"

namespace LotusChartInternal
{

///////////////////////////////////////////////////////////////////
//! the chart of a Lotus Pro Dos
class Chart final : public WKSChart
{
public:
	//! constructor
	explicit Chart(LotusChart &parser, std::shared_ptr<WPSStream> const &stream)
		: WKSChart()
		, m_fileType(-1)
		, m_hasLegend(false)
		, m_fileSerieStyles(false)
		, m_parser(parser)
		, m_stream(stream)
	{
	}
	//! send the zone content (called when the zone is of text type)
	void sendContent(TextZone const &zone, WPSListenerPtr &listener) const final;
	//! the chart type
	int m_fileType;
	//! a flag to know if we have some legend
	bool m_hasLegend;
	//! a flag to know if we have seen some serie style
	bool m_fileSerieStyles;
	//! wk3 serie format
	struct SerieFormat
	{
		//! constructor
		SerieFormat()
			: m_color(0)
			, m_hash(0)
			, m_yAxis(1)
			, m_format(0)
			, m_align(0)
		{
		}
		//! the color
		int m_color;
		//! the hash
		int m_hash;
		//! the y axis
		int m_yAxis;
		//! the format
		int m_format;
		//! the label alignement
		int m_align;
	};

	//! the series format
	SerieFormat m_serieFormats[6];
protected:
	//! the parser
	LotusChart &m_parser;
	//! the input stream
	std::shared_ptr<WPSStream> m_stream;
};

void Chart::sendContent(TextZone const &zone, WPSListenerPtr &listener) const
{
	if (!listener.get())
	{
		WPS_DEBUG_MSG(("LotusChartInternal::Chart::sendContent: no listener\n"));
		return;
	}
	long pos = m_stream->m_input->tell();
	listener->setFont(zone.m_font);
	bool sendText=false;
	for (auto const &e : zone.m_textEntryList)
	{
		if (!e.valid()) continue;
		if (sendText) listener->insertEOL(true);
		m_parser.sendText(m_stream, e);
		sendText=true;
	}
	m_stream->m_input->seek(pos, librevenge::RVNG_SEEK_SET);
}

//! the state of LotusChart
struct State
{
	//! constructor
	State()
		: m_version(-1)
		, m_idChartMap()
		, m_chartId(-1)
	{
	}
	/** returns a chart corresponding to an id, create it if needed.

		\note almost always a chart definition appears before the
		other chart's structures, but this is not always true...
	 */
	std::shared_ptr<Chart> getChart(int id, LotusChart &parser, std::shared_ptr<WPSStream> stream)
	{
		if (m_idChartMap.find(id)!=m_idChartMap.end())
			return m_idChartMap.find(id)->second;
		std::shared_ptr<Chart> newChart(new LotusChartInternal::Chart(parser, stream));
		if (id>=0)
			m_idChartMap[id]=newChart;
		else
		{
			WPS_DEBUG_MSG(("LotusChartInternal::State::getChart: call with id=%d, create temporary chart\n", id));
		}
		return newChart;

	}
	//! the file version
	int m_version;
	//! list of chart
	std::map<int,std::shared_ptr<Chart> > m_idChartMap;
	//! the current chart id(wps3Mac)
	int m_chartId;
};

}

// constructor, destructor
LotusChart::LotusChart(LotusParser &parser)
	: m_listener()
	, m_mainParser(parser)
	, m_styleManager(parser.m_styleManager)
	, m_state(new LotusChartInternal::State)
{
}

LotusChart::~LotusChart()
{
}

void LotusChart::cleanState()
{
	m_state.reset(new LotusChartInternal::State);
}

std::map<std::string,int> LotusChart::getNameToChartIdMap() const
{
	std::map<std::string,int> res;
	for (auto it : m_state->m_idChartMap)
	{
		auto const &chart = it.second;
		if (chart || !chart->m_name.empty())
			res[chart->m_name.cstr()]=it.first;
	}
	return res;
}

void LotusChart::updateState()
{
	std::set<int> toRemoveSet;
	/* In pc's wk3 files, the current chart is unamed while the others
	   have a name.  So if we have more than one chart, suppose that
	   the creator has named all its charts and remove the unamed
	   chart to avoid dupplication...

	   In mac's wk3 files, all used files are named, so this must be ok
	 */
	bool removeNoName=version()==1 && m_state->m_idChartMap.size()>1;
	for (auto &it : m_state->m_idChartMap)
	{
		auto id=it.first;
		auto chart=it.second;
		if (!chart || (removeNoName && chart->m_name.empty()))
		{
			toRemoveSet.insert(id);
			continue;
		}
		updateChart(*chart, id);
		// check if the chart contain at least one serie
		bool findSomeSerie=false;
		for (auto cIt : chart->getIdSerieMap())
		{
			if (cIt.second.m_ranges[0].valid(cIt.second.m_ranges[1]))
			{
				findSomeSerie=true;
				break;
			}
		}
		if (!findSomeSerie)
			toRemoveSet.insert(id);
	}
	for (auto id : toRemoveSet)
		m_state->m_idChartMap.erase(id);
}

int LotusChart::getNumCharts() const
{
	return int(m_state->m_idChartMap.size());
}

int LotusChart::version() const
{
	if (m_state->m_version<0)
		m_state->m_version=m_mainParser.version();
	return m_state->m_version;
}

////////////////////////////////////////////////////////////
// low level

////////////////////////////////////////////////////////////
// general
////////////////////////////////////////////////////////////

bool LotusChart::readChart(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = long(libwps::read16(input));
	if (type != 0x11)
	{
		WPS_DEBUG_MSG(("LotusChart::readChart: not a chart name\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	f << "Entries(ChartDef):sz=" << sz << ",";
	if (sz < 0xB2) // find b2 or b3
	{
		WPS_DEBUG_MSG(("LotusChart::readChart: chart name is too short\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto id=int(libwps::readU8(input));
	auto chart=m_state->getChart(id, *this, stream);
	f << "id=" << id << ",";
	std::string name("");
	auto fontType=m_mainParser.getDefaultFontType();
	for (int i=0; i<16; ++i)
	{
		unsigned char c = libwps::readU8(input);
		if (c == '\0') break;
		name+=char(c);
	}
	if (!name.empty())
	{
		chart->m_name=libwps_tools_win::Font::unicodeString(name, fontType);
		f << name << ",";
	}
	input->seek(pos+4+17, librevenge::RVNG_SEEK_SET);
	// group 0: title, 1: axis name, 2: axis data+note+subtitle+legend
	int val;
	for (int i=0; i<3; ++i) // useme
	{
		val=int(libwps::read8(input));
		if (val) f << "font[group" << i << "]=" << val << ",";
	}
	for (int i=0; i<6; ++i)
	{
		chart->m_serieFormats[i].m_color=val=int(libwps::readU8(input));
		if (!val) continue; // auto
		if (val==255)
			f << "serie" << i << "[color]=range,";
		else if (val==254)
			f << "serie" << i << "[color]=hidden,";
		else
			f << "serie" << i << "[color]=" << val << ",";
	}
	val=int(libwps::read8(input));
	if (val) f << "f0=" << val << ",";
	for (int i=0; i<6; ++i)
	{
		chart->m_serieFormats[i].m_hash=val=int(libwps::readU8(input));
		if (!val) continue;
		if (val==255)
			f << "hash[serie" << i << "]=range,";
		else
			f << "hash[serie" << i << "]=" << val << ",";
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	pos=input->tell();
	f.str("");
	f << "ChartDef-A:";
	for (int i=0; i<6; ++i)   // small number
	{
		val=int(libwps::read8(input));
		if (!val) continue;
		f << "f" << i << "=" << val << ",";
	}
	for (int i=0; i<3; ++i) // useme
	{
		val=int(libwps::read8(input));
		if (val) f << "fSize[group" << i << "]=" << val << ",";
	}
	val=int(libwps::readU8(input));
	bool showGridY=(val&2);
	if (val&1) // checkme
		f << "X,";
	chart->getAxis(0).m_showGrid=(val&1);
	if (val&2)
		f << "Y,";
	chart->getAxis(1).m_showGrid=showGridY;
	val &= 0xfc;
	if (val) f << "##grid" <<  "=" << val << ",";
	val=int(libwps::readU8(input)); // can be 0,1,2
	if (val&1) f << "color=bw,";
	val &= 0xfe;
	if (val) f << "##color" <<  "=" << val << ",";
	auto const &chartType=chart->m_fileType=int(libwps::readU8(input));
	chart->m_type=WKSChart::Serie::S_Bar;
	switch (chartType)
	{
	case 0: // line
		chart->m_type=WKSChart::Serie::S_Line;
		break;
	case 1: // bar
		break;
	case 2: // XY
		chart->m_type=WKSChart::Serie::S_Scatter;
		break;
	case 3: // bar stacked
		chart->m_dataStacked=true;
		break;
	case 4: // pie
		chart->m_type=WKSChart::Serie::S_Circle;
		break;
	case 5: // min-max
		chart->m_type=WKSChart::Serie::S_Stock;
		break;
	case 6: // radar (checkme)
		chart->m_type=WKSChart::Serie::S_Radar;
		break;
	case 7: // mixed
		break;
	default:
		WPS_DEBUG_MSG(("LotusChart::readChart: unknown chart type\n"));
		f << "###";
	}
	f << "type=" << chartType << ",";

	char const *axisNames[]= {"X","Y","YSecond"};
	for (int i=0; i<3; ++i)
	{
		val=int(libwps::read8(input));
		// 0=auto
		if (val==-1)
		{
			chart->getAxis(i).m_automaticScaling=false;
			f << "scale[" << axisNames[i] << "]=manual,";
		}
		else if (val) f << "###scale[" << axisNames[i] << "]=" << val << ",";
	}
	for (auto axisName : axisNames) // useme
	{
		val=int(libwps::read8(input));
		// 0=auto
		if (val==-1) f << "setExponent[" << axisName << "]=manual,";
		else if (val) f << "###setExponent[" << axisName << "]=" << val << ",";
	}
	for (auto axisName : axisNames) // useme
	{
		val=int(libwps::read8(input));
		// 0=auto
		if (val==-1) f << "legend[" << axisName << "]=manual,";
		else if (val==1) f << "legend[" << axisName << "]=none,";
		else if (val) f << "###legend[" << axisName << "]=" << val << ",";
	}
	for (int i=0; i<3; ++i)
	{
		val=int(libwps::read8(input));
		// 0=normal
		auto &axis=chart->getAxis(i);
		if (val==1)
		{
			f << "axis[" << axisNames[i] << "]=log,";
			axis.m_type= WKSChart::Axis::A_Logarithmic;
		}
		else
		{
			axis.m_type= WKSChart::Axis::A_Numeric;
			if (val) f << "###axis[" << axisNames[i] << "]=" << val << ",";
		}
	}
	for (auto axisName : axisNames)
	{
		val=int(libwps::read8(input));
		// 0=auto
		if (val==-1) f << "setWidth[" << axisName << "]=manual,";
		else if (val) f << "###setWidth[" << axisName << "]=" << val << ",";
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	pos=input->tell();
	f.str("");
	f << "ChartDef-B:";
	for (int i=0; i<6; ++i)
	{
		chart->m_serieFormats[i].m_yAxis=val=int(libwps::read8(input));
		if (val==1) continue; // primary axis
		if (val==2)
			f << "serie" << i << "[axis]=secondary,";
		else
			f << "##serie" << i << "[axis]=" << val << ",";
	}
	for (int i=0; i<6; ++i)   // small number
	{
		chart->m_serieFormats[i].m_format=val=int(libwps::readU8(input));
		if (!val) continue;
		if (val<5)
		{
			char const *wh[]= {"both", "lines", "symbols", "neither", "area"};
			f << "serie" << i << "[format]=" << wh[val] << ",";
		}
		else
			f << "##serie" << i << "[format]=" << val << ",";
	}
	for (int i=0; i<6; ++i)   // small number
	{
		chart->m_serieFormats[i].m_align=val=int(libwps::readU8(input));
		if (!val) continue;
		if (val<5)
		{
			char const *wh[]= {"center", "right", "below", "left", "above"};
			f << "serie" << i << "[align]=" << wh[val] << ",";
		}
		else
			f << "##serie" << i << "[align]=" << val << ",";
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	pos=input->tell();
	f.str("");
	f << "ChartDef-C:";
	for (int i=0; i<7; ++i)   // small number expect f18=0|4|64
	{
		val=int(libwps::read8(input));
		if (i==0)
		{
			chart->getAxis(1).m_showGrid=false;
			chart->getAxis(2).m_showGrid=false;
			if (val==0)   // 0: means y primary
			{
				if (showGridY) chart->getAxis(1).m_showGrid=true;
			}
			else if (val==1)
			{
				if (showGridY) chart->getAxis(2).m_showGrid=true;
				f << "grid[hori]=ysecond,";
			}
			else if (val==2)
			{
				if (showGridY)
				{
					chart->getAxis(1).m_showGrid=true;
					chart->getAxis(2).m_showGrid=true;
				}
				f << "grid[hori]=y+ysecond,";
			}
			else
				f << "##grid[hori]=" << val << ",";
			continue;
		}
		else if (i==1)
		{
			if (val&1)
			{
				chart->m_dataVertical=true;
				f << "swapXY,"; // Y is horizontal, X vertical
			}
			val &= 0xfe;
		}
		else if (i==3)
		{
			if (val&1)
			{
				chart->m_dataPercentStacked=true;
				f << "percentage,";
			}
			val &= 0xfe;
		}
		else if (i==4)
		{
			if (val&1)
			{
				chart->m_dataStacked=true;
				f << "stacked,";
			}
			val &= 0xfe;
		}
		else if (i==5)
		{
			if (val&1)
			{
				chart->m_is3D=true;;
				f << "drop[shadow],"; // ie. 2.5D
			}
			if (val&2)
			{
				chart->m_is3D=true;;
				chart->m_is3DDeep=true;;
				f << "3d[range],";  // ie 3D
			}
			if (val&4) f << "show[table],"; // show data table below the chart
			// plot area border
			if (val&0x10) f << "noBorder[L],";
			if (val&0x20) f << "noBorder[R],";
			if (val&0x40) f << "noBorder[T],";
			if (val&0x80) f << "noBorder[B],";
			val &= 0x8;
		}
		if (val) f << "f" << i+6 << "=" << val << ",";
	}
	for (int i=0; i<3; ++i)
	{
		// useme
		val=int(libwps::read8(input));
		if (val) f << "color[group" << i << "]=" << val << ",";
	}
	val=int(libwps::read16(input));
	if (val!=1) f << "ticks=" << val << ",";
	for (auto axisName : axisNames)
	{
		val=int(libwps::read16(input));
		if (val!=14) f << "width" << axisName << "=" << val << ",";;
	}
	for (int i=0; i<2; ++i)   // small number expect g2=g3=0|1
	{
		val=int(libwps::read16(input));
		if (val) f << "g" << i+1 << "=" << val << ",";
	}
	for (auto axisName : axisNames)
	{
		val=int(libwps::read16(input));
		if (val) f << "exp[manual," << axisName << "]=" << val << ",";;
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	pos=input->tell();
	f.str("");
	f << "ChartDef-D:";
	for (auto axisName : axisNames)
	{
		// useme
		f << axisName << "=[";
		val=int(libwps::readU8(input)); // or 2 bytes
		if (val!=0x71) f << "fmt=" << std::hex << val << std::dec << ",";
		for (int j=0; j<3; ++j)
		{
			val=int(libwps::readU8(input));
			if (val) f << "f" << j << "=" << val << ",";
		}
		f << "],";
	}
	ascFile.addDelimiter(input->tell(),'|');
	for (int i=0; i<3; ++i)
	{
		bool isNan;
		double value;
		if (!libwps::readDouble10(input, value, isNan))
			f << "##min" << axisNames[i] << ",";
		else
		{
			chart->getAxis(i).m_scaling[0]=float(value);
			if (value<0 || value>0)
				f << "min" << axisNames[i] << "=" << value << ",";
		}
	}
	for (int i=0; i<3; ++i)
	{
		bool isNan;
		double value;
		if (!libwps::readDouble10(input, value, isNan))
			f << "##max" << axisNames[i] << ",";
		else
		{
			chart->getAxis(i).m_scaling[1]=float(value);
			if (value<0 || value>0)
				f << "max" << axisNames[i] << "=" << value << ",";
		}
	}
	if (input->tell()!=endPos)
	{
		ascFile.addDelimiter(input->tell(),'|');
		input->seek(endPos, librevenge::RVNG_SEEK_SET);
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusChart::readChartName(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = long(libwps::read16(input));
	if (type != 0x12)
	{
		WPS_DEBUG_MSG(("LotusChart::readChartName: not a chart name\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	f << "Entries(ChartName):";
	if (sz < 3)
	{
		WPS_DEBUG_MSG(("LotusChart::readChartName: chart name is too short\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}

	auto cId=int(libwps::readU8(input));
	f << "chart[id]=" << cId << ",";
	auto chart=m_state->getChart(cId, *this, stream);
	auto id=int(libwps::readU8(input));
	f << "data[id]=" << id << ",";
	std::string name("");
	auto fontType=m_mainParser.getDefaultFontType();
	for (long i = 0; i < sz-2; i++)
	{
		unsigned char c = libwps::readU8(input);
		if (c == '\0') break;
		name+=char(c);
	}
	f << name << ",";
	if (!name.empty())
	{
		auto uniName=libwps_tools_win::Font::unicodeString(name, fontType);
		switch (id)
		{
		case 0:
		case 1:
		case 2:
		case 3:
		case 4:
		case 5:
		{
			auto *serie=chart->getSerie(id, true);
			serie->m_legendText=uniName;
			chart->m_hasLegend=true;
			break;
		}
		case 6:
		case 7:
		case 8:
			chart->getAxis(id-6).m_title=uniName;
			break;
		case 9:
		case 10:
		case 11:
			chart->getAxis(id-9).m_subTitle=uniName;
			break;
		case 12:
		case 13:
		case 14: // note1
		case 15: // note2
		{
			auto wh=id==12 ? WKSChart::TextZone::T_Title :
			        id==13 ? WKSChart::TextZone::T_SubTitle : WKSChart::TextZone::T_Footer;
			WPSEntry entry;
			entry.setBegin(pos+6);
			entry.setEnd(input->tell());
			auto *textZone=chart->getTextZone(wh, true);
			textZone->m_contentType=WKSChart::TextZone::C_Text;
			textZone->m_textEntryList.push_back(entry);
			break;
		}
		default:
			break;
		}
	}
	if (input->tell()!=pos+4+sz && input->tell()+1!=pos+4+sz)
	{
		WPS_DEBUG_MSG(("LotusChart::readChartName: the zone seems too short\n"));
		f << "##";
		ascFile.addDelimiter(input->tell(), '|');
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

////////////////////////////////////////////////////////////
// wk3mac
////////////////////////////////////////////////////////////
bool LotusChart::readMacHeader(std::shared_ptr<WPSStream> stream, long endPos, int &chartId)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	long sz=endPos-pos;
	f << "Entries(ChartMac):";
	if (sz<12)
	{
		WPS_DEBUG_MSG(("LotusChart::readChartMac: Oops the zone seems too short\n"));
		f << "###";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		m_state->m_chartId=chartId=-1;
		return true;
	}

	m_state->m_chartId=chartId=int(libwps::read16(input));
	f << "chart[id]=" << chartId << ",";
	auto chart=m_state->getChart(chartId,*this,stream);
	for (int i=0; i<5; ++i)   // f1=[4c][02]5[19], f2=50, f3=14
	{
		auto val=int(libwps::read16(input));
		if (!val) continue;
		if (i==1)
		{
			if (val&0x20)
			{
				f << "area[stacked],";
				chart->m_dataStacked=true;
			}
			val&=0xffdf;
		}
		if (val)
			f << "f" << i << "=" << std::hex << val << std::dec << ",";
	}
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	input->seek(endPos, librevenge::RVNG_SEEK_SET);
	return true;
}

bool LotusChart::readMacAxis(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	long sz=endPos-pos;

	f << "Entries(ChartAxis):id=" << m_state->m_chartId << ",";
	if (sz!=56)
	{
		WPS_DEBUG_MSG(("LotusChart::readMacAxis: the size seems bad\n"));
		f << "##sz";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto chart=m_state->getChart(m_state->m_chartId,*this,stream);
	auto id=int(libwps::readU8(input));
	if (id<0||id>=3)
	{
		WPS_DEBUG_MSG(("LotusChart::readMacAxis: the id seems bad\n"));
		f << "###";
	}
	auto &axis=chart->getAxis((id<0||id>=3) ? 4 : id);
	f << "id[axis]=" << id << ",";
	auto format=int(libwps::readU8(input));
	if ((format&0x20)==0)
	{
		f << "hidden[name],";
		axis.m_showTitle=false;
	}
	ascFile.addDelimiter(input->tell(),'|');
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	input->seek(endPos, librevenge::RVNG_SEEK_SET);
	return true;
}

bool LotusChart::readMacSerie(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	long sz=endPos-pos;

	f << "Entries(ChartSerie):id=" << m_state->m_chartId << ",";
	if (sz!=28)
	{
		WPS_DEBUG_MSG(("LotusChart::readMacSerie: the size seems bad\n"));
		f << "##sz";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto chart=m_state->getChart(m_state->m_chartId,*this,stream);
	auto id=int(libwps::readU8(input));
	f << "id[serie]=" << id << ",";
	chart->m_fileSerieStyles=true;
	auto *serie=chart->getSerie(id, true);
	serie->m_type=chart->m_type;
	auto format=int(libwps::readU8(input));
	if (id>=0 && id<6)
	{
		auto const &sFormat=chart->m_serieFormats[id];
		if (sFormat.m_yAxis==2)
			serie->m_useSecondaryY=true;
		if (chart->m_fileType<=3 || chart->m_fileType==7)
		{
			switch (sFormat.m_format)
			{
			case 0: // both
				if (chart->m_fileType==7 && (format&3)==1) // special case, mixed means last serie as line
					serie->m_type = WKSChart::Serie::S_Line;
				serie->m_pointType=WKSChart::Serie::P_Automatic;
				break;
			case 1: // lines
				serie->m_type = WKSChart::Serie::S_Line;
				break;
			case 2: // symbol
				serie->m_pointType=WKSChart::Serie::P_Automatic;
				serie->m_style.m_lineWidth=0;
				break;
			case 3:
				serie->m_style.m_lineWidth=0;
				break;
			case 4:
				serie->m_type = WKSChart::Serie::S_Area;
				break;
			default:
				break;
			}
		}
	}
	else
	{
		switch (format&3)
		{
		case 1:
			f << "line,";
			if (chart->m_fileType==7)
				serie->m_type=serie->S_Line;
			serie->m_pointType=WKSChart::Serie::P_Automatic;
			break;
		case 2: // bar
			break;
		default:
			f << "##format[low]=" << (format&3) << ",";
		}
		if (format&4)
		{
			if (chart->m_fileType<=3 || chart->m_fileType==7) // force area
				serie->m_type=serie->S_Area;
			f << "area,";
		}
	}
	if (format&0xf8) f << "##format[high]=" << (format>>5) << ",";
	auto val=int(libwps::readU16(input));
	if ((val>>8)==0x10)
		f << "L" << (val&0xff) << "[select],";
	else
		f << "##L[select]=" << std::hex << val << std::dec << ",";
	val=int(libwps::readU16(input));
	if ((val>>8)==0x20)
	{
		f << "C" << (val&0xff) << ",";
		m_styleManager->updateSurfaceStyle(val&0xff, serie->m_style);
	}
	else
		f << "##C=" << std::hex << val << std::dec << ",";
	for (int i=0; i<2; ++i)   // i==0: ? , i==1: surface border
	{
		val=int(libwps::readU16(input));
		if ((val>>8)==0x10)
		{
			f << "L" << (val&0xff) << (i==1 ? "[1]" : "") << ",";
			if (i==0)
				m_styleManager->updateLineStyle(val&0xff, serie->m_style);
		}
		else
			f << "##L"  << (i==1 ? "[1]" : "") << "=" << std::hex << val << std::dec << ",";
	}
	val=int(libwps::readU16(input));
	if ((val>>8)==0x20)   // surface external: used for pattern, ...
	{
		f << "C" << (val&0xff) << "[ext],";
		m_styleManager->updateSurfaceStyle(val&0xff, serie->m_style);
	}
	else
		f << "##Cext=" << std::hex << val << std::dec << ",";
	val=int(libwps::readU16(input));
	if (val!=id) f << "P" << val << ",";
	if (serie->m_pointType != WKSChart::Serie::P_None)
	{
		switch (val)
		{
		case 0: // square
		case 3: // hollow square
			serie->m_pointType=WKSChart::Serie::P_Square;
			break;
		case 1: // diamond
		case 4: // hollow diamond
			serie->m_pointType=WKSChart::Serie::P_Diamond;
			break;
		case 2: // triangle
		case 5: // hollow triangle
			serie->m_pointType=WKSChart::Serie::P_Arrow_Up;
			break;
		case 8: // triangle inverted
			serie->m_pointType=WKSChart::Serie::P_Arrow_Down;
			break;
		case 6: // circle
			serie->m_pointType=WKSChart::Serie::P_Circle;
			break;
		case 7: // star
			serie->m_pointType=WKSChart::Serie::P_Star;
			break;
		case 12: // x
			serie->m_pointType=WKSChart::Serie::P_X;
			break;
		case 14: // *
			serie->m_pointType=WKSChart::Serie::P_Asterisk;
			break;
		case 16: // +
			serie->m_pointType=WKSChart::Serie::P_Plus;
			break;
		case 18: // Y inverted
			serie->m_pointType=WKSChart::Serie::P_Bow_Tie;
			break;
		case 19: // -
			serie->m_pointType=WKSChart::Serie::P_Horizontal_Bar;
			break;
		case 20: // |
			serie->m_pointType=WKSChart::Serie::P_Vertical_Bar;
			break;
		default:
			break;
		}
	}

	for (int i=0; i<7; ++i)   // 0
	{
		val=int(libwps::readU16(input));
		if (val)
			f << "f" << i+1 << "=" << val << ",";
	}
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusChart::readMacPlacement(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	long sz=endPos-pos;

	f << "Entries(ChartPlacement):id=" << m_state->m_chartId << ",";
	if (sz!=8)
	{
		WPS_DEBUG_MSG(("LotusChart::readMacPlacement: the size seems bad\n"));
		f << "##sz";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto chart=m_state->getChart(m_state->m_chartId,*this,stream);
	auto val=int(libwps::readU8(input));
	if ((val&0x10)==0)
		f << "hidden,";
	switch (val&3)
	{
	case 1:
		f << "title,";
		if ((val&0x10)==0)
		{
			WKSChart::TextZone::Type allTypes[]= {WKSChart::TextZone::T_Title, WKSChart::TextZone::T_SubTitle};
			for (auto type : allTypes)
				chart->getTextZone(type,true)->m_show=false;
		}
		break;
	case 2:
		f << "note,";
		if ((val&0x10)==0) chart->getTextZone(WKSChart::TextZone::T_Footer,true)->m_show=false;
		break;
	default:
		f << "##wh=" << (val&3) << ",";
		break;
	}
	val &= 0xec;
	if (val) f << "fl0=" << std::hex << val << std::dec << ",";
	val=int(libwps::readU8(input));
	if (val&0x10)
		f << "manual,";
	else if (val!=1)
		f << "pos=" << val << ",";
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusChart::readMacFloor(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	long sz=endPos-pos;

	f << "Entries(ChartFloor):id=" << m_state->m_chartId << ",";
	if (sz!=17)
	{
		WPS_DEBUG_MSG(("LotusChart::readMacFloor: the size seems bad\n"));
		f << "##sz";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto chart=m_state->getChart(m_state->m_chartId,*this,stream);
	for (int i=0; i<4; ++i)
	{
		auto val=int(libwps::readU8(input));
		int const expected[]= {0xf,0x1e,0x12,0};
		if (val!=expected[i])
			f << "f" << i << "=" << val << ",";
	}
	for (int i=0; i<5; ++i)   // i=4: floor style
	{
		auto val=int(libwps::readU16(input));
		if ((val>>8)==0x20)
		{
			f << "C" << (val&0xff) << "[" << i << "],";
			if (i==4)
				m_styleManager->updateSurfaceStyle(val&0xff,chart->m_floorStyle);
		}
		else
			f << "##C=" << std::hex << val << std::dec << "[" << i << "],";
	}
	auto val=int(libwps::readU16(input));
	if ((val>>8)==0x10)
		f << "L" << (val&0xff) << ",";
	else
		f << "##L=" << std::hex << val << std::dec << ",";
	val=int(libwps::readU8(input)); // 0
	if (val) f << "f4=" << val << ",";
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusChart::readMacLegend(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	long sz=endPos-pos;

	f << "Entries(ChartLegend):id=" << m_state->m_chartId << ",";
	if (sz!=7)
	{
		WPS_DEBUG_MSG(("LotusChart::readMacLegend: the size seems bad\n"));
		f << "##sz";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto chart=m_state->getChart(m_state->m_chartId,*this,stream);
	auto val=int(libwps::readU8(input));
	if (val&0x10)
		f << "manual,";
	val&=0xef;
	if (val!=4) f << "f0=" << std::hex << val << std::dec << ",";
	val=int(libwps::readU8(input));
	if ((val&0x1)==0)
	{
		f << "hidden,";
		chart->m_hasLegend=false;
	}
	val&=0xfe;
	if (val!=2) f << "f1=" << std::hex << val << std::dec << ",";
	val=int(libwps::readU16(input));
	if ((val>>8)==0x40)
		f << "G" << (val&0xff) << ",";
	else
		f << "##G=" << std::hex << val << std::dec << ",";
	val=int(libwps::readU16(input));
	if (val!=2)
		f << "f2=" << val << ",";
	val=int(libwps::readU8(input)); // vertical/horizontal?
	if (val) f << "f3=" << val << ",";
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusChart::readMacPlotArea(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	long sz=endPos-pos;

	f << "Entries(ChartPlotArea):id=" << m_state->m_chartId << ",";
	if (sz!=18)
	{
		WPS_DEBUG_MSG(("LotusChart::readMacPlotArea: the size seems bad\n"));
		f << "##sz";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto chart=m_state->getChart(m_state->m_chartId,*this,stream);
	auto val=int(libwps::readU16(input));
	if ((val>>8)==0x10)
	{
		f << "L" << (val&0xff) << ",";
		m_styleManager->updateLineStyle(val&0xff, chart->m_wallStyle);
	}
	else
		f << "##L[select]=" << std::hex << val << std::dec << ",";
	val=int(libwps::readU16(input));
	if ((val>>8)==0x20)
	{
		f << "C" << (val&0xff) << ",";
		m_styleManager->updateSurfaceStyle(val&0xff, chart->m_wallStyle);
	}
	else
		f << "##C=" << std::hex << val << std::dec << ",";
	val=int(libwps::readU8(input));
	if (val&0x10)
		f << "manual,";
	val&=0xef;
	if (val) f << "f0=" << std::hex << val << std::dec << ",";
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusChart::readMacPosition(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	long sz=endPos-pos;

	f << "Entries(ChartPosition):id=" << m_state->m_chartId << ",";
	if (sz!=9)
	{
		WPS_DEBUG_MSG(("LotusChart::readMacPosition: the size seems bad\n"));
		f << "##sz";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	int dim[4];
	for (auto &d : dim) d=int(libwps::read16(input));
	if (dim[2]||dim[3])
	{
		// USEME
		float const scale=1.f/65536.f;
		f << "pos=" << Vec2f(scale*float(dim[0]),scale*float(dim[3]))
		  << "<->" << Vec2f(scale*float(dim[2]),scale*float(dim[1])) << "%,";
	}
	auto val=int(libwps::readU8(input));
	if (val) f << "f0=" << val << ",";
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	return true;
}
////////////////////////////////////////////////////////////
// Windows wk3 and wk4 files
////////////////////////////////////////////////////////////
bool LotusChart::readSerie(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	long sz=endPos-pos;

	f << "Entries(ChartSerie):";
	if (sz!=22)
	{
		WPS_DEBUG_MSG(("LotusChart::readSerie: the size seems bad\n"));
		f << "##sz";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto cId=int(libwps::readU8(input));
	f << "id[chart]=" << cId  << ",";
	auto chart=m_state->getChart(cId,*this,stream);
	chart->m_fileSerieStyles=true;
	for (int i=0; i<3; ++i)   // f0=0|b4
	{
		auto val=int(libwps::readU8(input));
		if (val)
			f << "f" << i << "=" << val << ",";
	}
	auto id=int(libwps::readU8(input));
	f << "id[serie]=" << id << ",";
	auto *serie=chart->getSerie(id, true);
	serie->m_type = chart->m_type;
	auto format=int(libwps::readU8(input));
	if (format==2)
	{
		serie->m_useSecondaryY=true;
		f << "secondary[y],";
	}
	else if (format!=1)
		f << "##yAxis=" << format << ",";
	format=int(libwps::readU8(input));
	if (format&8)
		f << "bar[force],";
	else if (chart->m_fileType==7)
		serie->m_type = WKSChart::Serie::S_Line;
	format &= 0xf7;
	serie->m_style.m_lineWidth=1;
	if (format>=0 && format<5)
	{
		char const *wh[]= {"both", "lines", "symbols", "neither", "area"};
		f << "format=" << wh[format] << ",";
		if (chart->m_fileType<=3 || chart->m_fileType==7)
		{
			switch (format)
			{
			case 0: // both
				//if (chart->m_fileType==7) // special case, mixed means last serie as line
				//	serie->m_type = WKSChart::Serie::S_Line;
				serie->m_pointType=WKSChart::Serie::P_Automatic;
				break;
			case 1: // lines
				if (chart->m_fileType==7)
					serie->m_type = WKSChart::Serie::S_Line;
				break;
			case 2: // symbol
				serie->m_pointType=WKSChart::Serie::P_Automatic;
				serie->m_style.m_lineWidth=0;
				break;
			case 3:
				serie->m_style.m_lineWidth=0;
				break;
			case 4:
				if (chart->m_fileType==0)
					chart->m_dataStacked=true;
				serie->m_type = WKSChart::Serie::S_Area;
				break;
			default:
				break;
			}
		}
	}
	else
		f << "###format=" << format << ",";
	int val;
	for (int i=0; i<2; ++i)
	{
		val=int(libwps::readU8(input));
		if (val)
			f << "f" << i+3 << "=" << std::hex << val << std::dec << ",";
	}
	auto col=int(libwps::readU8(input));  // 32|44|46|7c|81|a8|ff: checkme classic color 256 ?
	WPSColor color[3]= {WPSColor(255,0,0), WPSColor::white(), WPSColor::black()};
	if (m_styleManager->getColor256(col, color[0]))
		f << "color=" << color[0] << ",";
	else
		f << "##color=" << col << ",";
	for (int i=0; i<6; ++i)   // g0=1|3,g1=0-3,g3=0|-1,g4=0|1
	{
		val=int(libwps::read8(input));
		if (i==0)
		{
			if (val!=1)
				f << "line[style]=" << val << ",";
			switch (val)
			{
			case 0: // hidden
				serie->m_style.m_lineWidth=0;
				break;
			case 2: // long dash
				serie->m_style.m_lineDashWidth.push_back(7);
				serie->m_style.m_lineDashWidth.push_back(3);
				break;
			case 3: // dotted line
			case 6: // chain dotted line
				serie->m_style.m_lineDashWidth.push_back(1);
				serie->m_style.m_lineDashWidth.push_back(3);
				break;
			case 4: // chain dash line ?
			case 7: // dash line
				serie->m_style.m_lineDashWidth.push_back(3);
				serie->m_style.m_lineDashWidth.push_back(3);
				break;
			case 5: // long dotted
				serie->m_style.m_lineDashWidth.push_back(2);
				serie->m_style.m_lineDashWidth.push_back(3);
				break;
			case 1: // solid
			default:
				break;
			}
			continue;
		}
		else if (i==2)
		{
			if (val!=1)
				f << "symbol=" << val << ",";
			if (serie->m_pointType != WKSChart::Serie::P_None)
			{
				switch (val)
				{
				case 1: // square
				case 4: // hollow square
					serie->m_pointType=WKSChart::Serie::P_Square;
					break;
				case 2: // diamond
				case 5: // hollow diamond
					serie->m_pointType=WKSChart::Serie::P_Diamond;
					break;
				case 3: // triangle
				case 6: // hollow triangle
					serie->m_pointType=WKSChart::Serie::P_Circle;
					break;
				case 13: // x in square
				case 16:
					serie->m_pointType=WKSChart::Serie::P_X;
					break;
				case 14: // + in diamond
				case 17:
					serie->m_pointType=WKSChart::Serie::P_Plus;
					break;
				case 19: // Y inverted in triangle
				case 22: // Y inverted
					serie->m_pointType=WKSChart::Serie::P_Bow_Tie;
					break;
				default:
					break;
				}
			}
			continue;
		}
		else if (i==3)
		{
			if (m_styleManager->getColor256(uint8_t(val), color[2]))
			{
				if (!color[2].isBlack())
					f << "color[line]=" << color[2] << ",";
			}
			else
				f << "##color[line]=" << uint8_t(val) << ",";
			continue;
		}
		else if (i==4)
		{
			// 0: none, 1: normal, 2: long dash, 3: dash
			if (val!=1) f << "line[style]=" << val << ",";
			continue;
		}
		else if (i==5)
		{
			if (!val) continue;
			if (val>0 && val<8 && serie->m_style.m_lineWidth>0)
				serie->m_style.m_lineWidth=float(val+1);
			f << "line[width]=" << val+1 << ",";
			continue;
		}
		if (!val) continue;
		f << "g" << i << "=" << val << ",";
	}
	col=int(libwps::readU8(input));
	if (m_styleManager->getColor256(col, color[1]))
	{
		if (!color[1].isWhite())
			f << "color[surf]=" << color[1] << ",";
	}
	else
		f << "##color[surf2]=" << col << ",";
	int patternId=1;
	for (int i=0; i<5; ++i)   // h0=f
	{
		val=int(libwps::read8(input));
		if (i==1)
		{
			patternId=val;
			if (val!=1) f << "pattern[id]=" << val << ",";
			continue;
		}
		else if (i==4)   // see chartDef
		{
			// useme: ie. probably better than to use the pattern id

			// 1: solid, 2: fine cross, 3: fine double, 4: fine triple
			// 5: coarse cross, 6: coarse double, 7: coarse single, 8: hollow
			if (val==-1) f << "hash[id]=range,";
			else if (val!=1) f << "hash[id]=" << val << ",";
			continue;
		}
		if (!val) continue;
		f << "h" << i << "=" << val << ",";
	}
	WPSGraphicStyle::Pattern pattern;
	bool has0D=serie->m_pointType!=WKSChart::Serie::P_None;
	bool has1D=serie->is1DStyle() || (chart->m_fileType==2 && serie->m_style.m_lineWidth>0);
	bool has2D=!serie->is1DStyle() || (has1D && chart->m_is3D);
	if (patternId>0 && m_styleManager->getPattern64(patternId,pattern))
	{
		if (version()>=3)
		{
			pattern.m_colors[0]=color[0];
			pattern.m_colors[1]=color[1];
		}
		else
		{
			pattern.m_colors[0]=WPSColor::white();
			pattern.m_colors[1]=color[0];
		}
		WPSColor finalColor;
		if (has0D || has2D)
		{
			if (pattern.getUniqueColor(finalColor))
				serie->m_style.setSurfaceColor(finalColor);
			else
				serie->m_style.setPattern(pattern);
		}
		if (has1D && pattern.getAverageColor(finalColor))
			serie->m_style.m_lineColor=finalColor;
	}
	else
	{
		if (has1D || patternId==0)
			serie->m_style.m_lineColor=color[0];
		if (has0D || (has2D && patternId!=0))
			serie->m_style.setSurfaceColor(color[0]);
	}
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusChart::readSerieName(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	long sz=endPos-pos;

	f << "Entries(ChartSerName):";
	if (sz<6)
	{
		WPS_DEBUG_MSG(("LotusChart::readSerieName: the size seems bad\n"));
		f << "##sz";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto cId=int(libwps::readU8(input));
	f << "id[chart]=" << cId  << ",";
	auto chart=m_state->getChart(cId,*this,stream);
	for (int i=0; i<3; ++i)   // 0
	{
		auto val=int(libwps::readU8(input));
		if (val)
			f << "f" << i << "=" << val << ",";
	}
	auto id=int(libwps::readU8(input));
	f << "id[serie]=" << id << ",";
	std::string name("");
	for (long i = 0; i < sz-5; i++)
	{
		unsigned char c = libwps::readU8(input);
		if (c == '\0') break;
		name+=char(c);
	}
	if (!name.empty())
	{
		f << name << ",";
		auto *serie=chart->getSerie(id, true);
		serie->m_legendText=libwps_tools_win::Font::unicodeString(name, m_mainParser.getDefaultFontType());
		chart->m_hasLegend=true;
	}
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusChart::readSerieWidth(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	long sz=endPos-pos;

	f << "Entries(ChartSerWidth):";
	if (sz!=8)
	{
		WPS_DEBUG_MSG(("LotusChart::readSerieWidth: the size seems bad\n"));
		f << "##sz";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << "id[chart]=" << int(libwps::readU8(input)) << ",";
	for (int i=0; i<3; ++i)   // 0
	{
		auto val=int(libwps::readU8(input));
		if (val)
			f << "f" << i << "=" << val << ",";
	}
	f << "id[serie]=" << int(libwps::readU8(input)) << ",";
	// checkme
	auto val=int(libwps::readU8(input)); // 0
	if (val) f << "f3=" << val << ",";
	f << "w[inv]=" << int(libwps::readU16(input)) << ",";
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusChart::readPlotArea(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	long sz=endPos-pos;

	f << "Entries(ChartPlotArea):";
	if (sz!=111)
	{
		WPS_DEBUG_MSG(("LotusChart::readPlotArea: the size seems bad\n"));
		f << "##sz";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto cId=int(libwps::readU8(input));
	f << "id[chart]=" << cId  << ",";
	auto chart=m_state->getChart(cId,*this,stream);
	for (int i=0; i<3; ++i)   // f0=0|c,f,3a
	{
		auto val=int(libwps::readU8(input));
		if (val) f << "f" << i << "=" << std::hex << val << std::dec << ",";
	}
	bool isNan;
	double value;
	for (int i=0; i<6; ++i)
	{
		if (!libwps::readDouble10(input, value, isNan))
			f << "##value,";
		else if (value<0 || value>0)
			f << "v" << i << "=" << value << ",";
	}
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());

	pos=input->tell();
	f.str("");
	f << "ChartPlotArea-A:";
	for (int i=0; i<3; ++i)   // f0=0|1, f1=0|1|5, f2=0|1
	{
		auto val=int(libwps::read16(input));
		if (val) f << "f" << i << "=" << val << ",";
	}
	char const *zonesName[]= {"title", "note", "serie,legend", "plot"};
	for (int i=0; i<4; ++i)
	{
		int dim[4]; // in percent * 65536
		for (auto &d : dim) d=int(libwps::readU16(input));
		if (dim[0]==0 && dim[1]==0 && dim[2]==0 && dim[3]==0)
			continue;
		float const scale=1.f/65536.f;
		WPSBox2f box(Vec2f(scale*float(dim[0]),1.f-scale*float(dim[1])), Vec2f(scale*float(dim[2]),1.f-scale*float(dim[3])));
		f << "pos[" << zonesName[i] << "]=" << box << "%,";
		if (i==2)
		{
			auto &legend=chart->getLegend();
			legend.m_autoPosition=false;
			chart->m_legendPosition=box;
		}
		else if (i==3)
			chart->m_plotAreaPosition=box;
	}
	for (int i=0; i<4; ++i) // UseME
	{
		auto val=int(libwps::readU8(input));
		if (!val) continue;
		f << "pos[" << zonesName[i] << "]=[";
		if (val&0x10)
			f << "manual,";
		if (i<2)
		{
			if ((val&0x10)==0)
			{
				if (val&1) f << "left,";
				if (val&2) f << "center,";
				if (val&4) f << "right,";
			}
			val &= 0xf8;
		}
		else if (i==2)
		{
			if ((val&0x10)==0)
			{
				if (val&4) f << "right,";
				if (val&8) f << "below,";
			}
			val &= 0xf3;
		}
		val &= 0xef;
		if (val)
			f << "fl=" << std::hex << val << std::dec << ",";
		f << "],";
	}
	char const *axisNames[]= {"X","Y","YSecond"};
	for (auto axisName : axisNames) // useme
	{
		auto val=int(libwps::readU8(input));
		if (val==0x10) continue;
		f << axisName << "[";
		if ((val&0x10)==0) f << "not10,";
		if (val&0x40) f << "major,";
		if (val&0x80) f << "minor,";
		val &= 0x2f;
		if (val)
			f << "fl=" << std::hex << val << std::dec;
		f << "],";
	}
	auto val=int(libwps::readU8(input));
	if (val) f << "fl=" << val << ",";
	val=int(libwps::readU8(input));
	if (val)
	{
		f << "type=" << val << ",";
		if (val==8)
			chart->m_type=WKSChart::Serie::S_Radar;
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	return true;
}

bool LotusChart::readFontsStyle(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	long sz=endPos-pos;

	f << "Entries(ChartFontsStyle):";
	if (sz!=38)
	{
		WPS_DEBUG_MSG(("LotusChart::readFontsStyle: the size seems bad\n"));
		f << "##sz";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << "id[chart]=" << int(libwps::readU8(input)) << ",";
	for (int i=0; i<3; ++i)   // 0
	{
		auto val=int(libwps::readU8(input));
		if (val)
			f << "f" << i << "=" << val << ",";
	}
	int prev=-1;
	f << "val=[";
	for (int i=0; i<17; ++i)   // 20-3e, 57, ...
	{
		auto val=int(libwps::readU16(input));
		if (val==prev)
			f << "=,";
		else
			f << "F" << val << ",";
		prev=val;
	}
	f << "],";
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusChart::readFramesStyle(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	long sz=endPos-pos;

	f << "Entries(ChartFramesStyle):";
	if (sz!=102)
	{
		WPS_DEBUG_MSG(("LotusChart::readFramesStyle: the size seems bad\n"));
		f << "##sz";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto cId=int(libwps::readU8(input));
	f << "id[chart]=" << cId << ",";
	auto chart=m_state->getChart(cId,*this,stream);
	for (int i=0; i<3; ++i)   // 0
	{
		auto val=int(libwps::readU8(input));
		if (val)
			f << "f" << i << "=" << val << ",";
	}
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());

	for (int i=0; i<4; ++i)
	{
		pos=input->tell();
		char const *zonesName[]= {"title", "serie,legend", "note", "plot"};
		f.str("");
		f << "ChartFramesStyle-" << zonesName[i] << ":";
		WPSColor color[4]= {WPSColor::black(), WPSColor::white(), WPSColor::black(), WPSColor::black()};
		auto val=int(libwps::readU8(input));
		WPSGraphicStyle style;
		if (!m_styleManager->getColor256(val, color[2]))
			f << "col[lineId]=###" << val << ",";
		else if (!color[2].isBlack())
		{
			f << "col[line]=" << color[2] << ",";
			style.m_lineColor=color[2];
		}
		val=int(libwps::readU8(input));
		if (val!=1)
		{
			f << "line[style]=" << val << ",";
			if (val==0) style.m_lineWidth=0;
		}
		val=int(libwps::readU8(input));
		if (val)
		{
			f << "line[width]=" << val+1 << ",";
			if (style.m_lineWidth>0)
				style.m_lineWidth=float(val+1);
		}
		for (int j=0; j<2; ++j)
		{
			val=int(libwps::readU8(input));
			if (!m_styleManager->getColor256(val, color[j]))
				f << "col[surf" << j << "]=###" << val << ",";
		}
		auto patId=int(libwps::readU8(input));
		WPSGraphicStyle::Pattern pattern;
		if (patId>0 && m_styleManager->getPattern64(patId,pattern))
		{
			pattern.m_colors[0]=color[1];
			pattern.m_colors[1]=color[0];

			WPSColor finalColor;
			if (!pattern.getUniqueColor(finalColor))
			{
				style.setPattern(pattern);
				f << pattern << ",";
			}
			else
			{
				style.setSurfaceColor(finalColor);
				if (!finalColor.isWhite())
					f << "surf=" << finalColor << ",";
			}
		}
		else
			f << "pattern[id]=##" << patId << ",";
		val=int(libwps::readU8(input));
		if (!m_styleManager->getColor256(val, color[3]))
			f << "frame[colId]=###" << val << ",";
		else if (!color[3].isBlack())
			f << "col[frame]=" << color[3] << ",";
		val=int(libwps::readU8(input));
		if ((i!=3 && val!=2) || (i==3 && val)) f << "type[frame]" << val << ",";
		if (i==0)
		{
			auto title=chart->getTextZone(WKSChart::TextZone::T_Title, true);
			title->m_style=style;
		}
		else if (i==1)
			chart->getLegend().m_style = style;
		else if (i==3)
			chart->m_wallStyle = chart->m_floorStyle = style;
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
	}
	pos=input->tell();
	f.str("");
	f << "ChartFramesStyle-A:";
	int val;
	f << "plot[line1]=[";
	val=int(libwps::readU8(input));
	WPSColor lineColor;
	if (!m_styleManager->getColor256(val, lineColor))
		f << "colId=###" << val << ",";
	else if (!lineColor.isBlack())
		f << lineColor << ",";
	val=int(libwps::readU8(input));
	if (val!=1)
	{
		f << "style=" << val << ",";
		if (val==0) chart->m_floorStyle.m_lineWidth=0;
	}
	val=int(libwps::readU8(input));
	if (val)
		f << "width=" << val+1 << ",";
	f << "],";
	ascFile.addDelimiter(input->tell(),'|');
	input->seek(pos+15, librevenge::RVNG_SEEK_SET);
	f << "plot[line2]=[";
	val=int(libwps::readU8(input));
	if (!m_styleManager->getColor256(val, lineColor))
		f << "colId=###" << val << ",";
	else if (!lineColor.isBlack())
		f << lineColor << ",";
	val=int(libwps::readU8(input));
	if (val!=1)
	{
		f << "style=" << val << ",";
		if (val==0) chart->m_floorStyle.m_lineWidth=0;
	}
	val=int(libwps::readU8(input));
	if (val)
		f << "width=" << val+1 << ",";
	f << "],";
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	pos=input->tell();
	f.str("");
	f << "ChartFramesStyle-B:";
	input->seek(pos+12, librevenge::RVNG_SEEK_SET);
	ascFile.addDelimiter(input->tell(),'|');
	f << "plot[line3]=[";
	val=int(libwps::readU8(input));
	if (!m_styleManager->getColor256(val, lineColor))
		f << "colId=###" << val << ",";
	else if (!lineColor.isBlack())
		f << lineColor << ",";
	val=int(libwps::readU8(input));
	if (val!=1)
	{
		f << "style=" << val << ",";
		if (val==0) chart->m_wallStyle.m_lineWidth=0;
	}
	val=int(libwps::readU8(input));
	if (val)
		f << "width=" << val+1 << ",";
	f << "],";
	ascFile.addDelimiter(input->tell(),'|');
	input->seek(pos+24, librevenge::RVNG_SEEK_SET);
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	pos=input->tell();
	f.str("");
	f << "ChartFramesStyle-C:";
	input->seek(pos+24, librevenge::RVNG_SEEK_SET);
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	return true;
}

////////////////////////////////////////////////////////////
// send data
////////////////////////////////////////////////////////////
bool LotusChart::convert(LotusParser::Link const &link, WKSChart::Position(&positions)[2]) const
{
	for (int i=0; i<2; ++i)
	{
		auto const &c = link.m_cells[i];
		positions[i].m_pos=Vec2i(c[0], c[1]);
		positions[i].m_sheetName = m_mainParser.getSheetName(c[2]);
	}
	return positions[0].valid(positions[1]);
}

void LotusChart::updateChart(LotusChartInternal::Chart &chart, int id)
{
	int const vers=version();
	// .wk4 pie chart does no have legend
	if (chart.m_hasLegend && (vers<3 || chart.m_fileType!=4))
	{
		chart.getLegend().m_show=true;
		chart.getLegend().m_autoPosition=true;
		chart.getLegend().m_relativePosition=WPSBorder::RightBit;
	}
	else
		chart.getLegend().m_show=false;
	auto links=m_mainParser.getLinksList(id);
	std::map<std::string, LotusParser::Link const &> linkMap;
	for (auto const &l : links)
		linkMap.emplace(l.m_name, l);

	// G[39-3e]: data series 0, 1, ...
	// G[40-45]: legend serie 0->5
	if (!chart.m_fileSerieStyles)   // wks3 dos pc
	{
		// we must create the serie, if there have data, ...
		bool findSerie=false;
		for (int i=5; i>=0; --i)
		{
			std::string dataName("G");
			dataName+=char(0x39+i);
			WKSChart::Position ranges[2];
			if (linkMap.find(dataName)==linkMap.end() || !convert(linkMap.find(dataName)->second,ranges))
				continue;
			auto *serie=chart.getSerie(i, true);
			for (int r=0; r<2; ++r) serie->m_ranges[r]=ranges[r];

			// check label
			std::string labelName("G");
			labelName+=char(0x40+i);
			if (linkMap.find(labelName)!=linkMap.end() && convert(linkMap.find(labelName)->second,ranges))
			{
				for (int r=0; r<2; ++r) serie->m_labelRanges[r]=ranges[r];
			}
			// now update the style
			auto const &format=chart.m_serieFormats[i];
			if (format.m_yAxis==2)
				serie->m_useSecondaryY=true;
			serie->m_type=chart.m_type;
			serie->m_style.m_lineWidth=1;
			if (chart.m_fileType==0 || chart.m_fileType==2 || chart.m_fileType==3 || chart.m_fileType==7)
			{
				switch (format.m_format)
				{
				case 0: // both
					if (chart.m_fileType==7 && !findSerie) // special case, mixed means last serie as line
						serie->m_type = WKSChart::Serie::S_Line;
					serie->m_pointType=WKSChart::Serie::P_Automatic;
					break;
				case 1: // lines
					serie->m_type = WKSChart::Serie::S_Line;
					break;
				case 2: // symbol
					serie->m_pointType=WKSChart::Serie::P_Automatic;
					serie->m_style.m_lineWidth=0;
					break;
				case 3:
					serie->m_style.m_lineWidth=0;
					break;
				case 4:
					serie->m_type = WKSChart::Serie::S_Area;
					break;
				default:
					break;
				}
			}
			findSerie=true;
			uint32_t const defColor[6] = {0xff0000, 0x00ff00, 0x0000ff, 0xffff00, 0x00ffff, 0xff00ff};
			WPSColor color=WPSColor(defColor[i]);
			if (format.m_color)
				m_styleManager->getColor256(format.m_color, color);
			// useme m_hash: H8->P2, H3->P3, H4->P4, H5->P12, H7->P6
			bool has0D=serie->m_pointType!=WKSChart::Serie::P_None;
			bool has1D=serie->is1DStyle();
			bool has2D=!serie->is1DStyle();
			if (has1D || format.m_hash==0)
				serie->m_style.m_lineColor=color;
			if (has0D || has2D)
				serie->m_style.setSurfaceColor(color);
		}
	}
	else
	{
		// G[47][22,27,2c,31,36,3b,40,45,4a,4f,54,59,5e]: data serie 6-18 (+1 label)
		// G[48][23,28,2d,32]: serie 19-22 (+1 label)
		for (auto it=chart.getIdSerieMap().begin(); it!=chart.getIdSerieMap().end(); ++it)
		{
			int sId=it->first;
			if (sId<0 || sId>22)
			{
				WPS_DEBUG_MSG(("LotusChart::updateChart: find unexpected id=%d\n", sId));
				continue;
			}
			std::string dataName("G"), labelName("G");
			if (sId<6)
			{
				dataName+=char(0x39+sId);
				labelName+=char(0x40+sId);
			}
			else if (sId<=18)
			{
				dataName+=char(0x47);
				dataName+=char(0x22+5*(sId-6));
				labelName+=char(0x47);
				labelName+=char(0x23+5*(sId-6));
			}
			else
			{
				dataName+=char(0x48);
				dataName+=char(0x23+5*(sId-19));
				labelName+=char(0x48);
				labelName+=char(0x24+5*(sId-19));
			}
			WKSChart::Position ranges[2];
			if (linkMap.find(dataName)==linkMap.end() || !convert(linkMap.find(dataName)->second,ranges))
			{
				if (vers>1)
				{
					WPS_DEBUG_MSG(("LotusChart::updateChart: can find data for serie %d in chart %d\n", sId, id));
				}
				continue;
			}
			auto *serie = chart.getSerie(sId, true);
			for (int i=0; i<2; ++i) serie->m_ranges[i]=ranges[i];
			if (linkMap.find(labelName)==linkMap.end() || !convert(linkMap.find(labelName)->second,ranges))
				continue;
			for (int i=0; i<2; ++i) serie->m_labelRanges[i]=ranges[i];
		}
	}
	for (int i=0; i<7; ++i)
	{
		// G[4f-51]: label axis x,y,ysecond
		// G[52-53]: title, subtile
		// G[54-55]: note1, note2
		std::string name("G");
		name+=char(0x4f+i);
		WKSChart::Position ranges[2];
		if (linkMap.find(name)==linkMap.end() || !convert(linkMap.find(name)->second,ranges))
			continue;
		if (i<3)
			chart.getAxis(i).m_titleRange=ranges[0];
		else
		{
			auto *zone=chart.getTextZone(i==3 ? WKSChart::TextZone::T_Title :
			                             i==4 ? WKSChart::TextZone::T_SubTitle : WKSChart::TextZone::T_Footer, true);
			zone->m_contentType=zone->C_Cell;
			zone->m_cell=ranges[0];
		}
	}
	// G[3f]: axis 0
	std::string name("G");
	name+=char(0x3f);
	WKSChart::Position ranges[2];
	if (linkMap.find(name)!=linkMap.end() && convert(linkMap.find(name)->second,ranges))
	{
		auto &axis = chart.getAxis(0);
		for (int r=0; r<2; ++r)	axis.m_labelRanges[r]=ranges[r];
	}
	else if (chart.m_fileType==2)
	{
		// if chart is a scatter, the first serie can store the xaxis data...
		auto *serie = chart.getSerie(0, false);
		if (serie)
		{
			auto &axis=chart.getAxis(0);
			for (int i=0; i<2; ++i)
			{
				axis.m_labelRanges[i]=serie->m_ranges[i];
				serie->m_ranges[i]=WKSChart::Position();
			}
		}
	}
	// G[23-28] color series 0->5
	// G[2a-2f] hatch series 0->5
	// G[4c-4e]: unit axis x,y,ysecond
}

bool LotusChart::sendText(std::shared_ptr<WPSStream> stream, WPSEntry const &entry)
{
	if (m_listener.get() == nullptr)
	{
		WPS_DEBUG_MSG(("LotusChart::sendText: I can not find the listener\n"));
		return false;
	}
	if (stream.get() == nullptr || !entry.valid())
		return true;
	RVNGInputStreamPtr &input = stream->m_input;
	input->seek(entry.begin(), librevenge::RVNG_SEEK_SET);
	m_listener->insertUnicodeString(libwps_tools_win::Font::unicodeString(input.get(), static_cast<unsigned long>(entry.length()), m_mainParser.getDefaultFontType()));
	return true;
}

bool LotusChart::sendCharts()
{
	if (m_listener.get() == nullptr)
	{
		WPS_DEBUG_MSG(("LotusChart::sendCharts: I can not find the listener\n"));
		return false;
	}
	Vec2i actPos(0,0);
	int actSquare=0;
	WPSGraphicStyle emptyStyle(WPSGraphicStyle::emptyStyle());
	for (auto &it : m_state->m_idChartMap)
	{
		if (!it.second || it.second->getIdSerieMap().empty()) continue;
		WPSPosition pos(Vec2f(float(512*actPos[0]),float(350*actPos[1])), Vec2f(512,350), librevenge::RVNG_POINT);
		pos.m_anchorTo = WPSPosition::Page;
		it.second->m_dimension=Vec2f(512,350); // set basic dimension
		sendChart(it.first, pos, emptyStyle);
		if (actPos[0]<actSquare)
			actPos[0]+=1;
		else if (actPos[1]<actSquare)
		{
			actPos[1]+=1;
			actPos[0]=actPos[1]==actSquare ? 0 : actSquare;
		}
		else
			actPos=Vec2i(++actSquare, 0);
	}
	return true;
}

bool LotusChart::sendChart(int cId, WPSPosition const &pos, WPSGraphicStyle const &style)
{
	if (m_listener.get() == nullptr)
	{
		WPS_DEBUG_MSG(("LotusChart::sendChart: I can not find the listener\n"));
		return false;
	}
	auto it=m_state->m_idChartMap.find(cId);
	if (it==m_state->m_idChartMap.end() || !it->second)
	{
		WPS_DEBUG_MSG(("LotusChart::sendChart: I can not find the chart with id=%d\n", cId));
		return false;
	}
	if ((it->second->m_dimension[0]<=0 || it->second->m_dimension[1]<=0) &&
	        pos.size()[0]>0 && pos.size()[1]>0)   // set basic dimension
	{
		float factor=WPSPosition::getScaleFactor(pos.unit(), librevenge::RVNG_POINT);
		it->second->m_dimension=factor*pos.size();
	}
	it->second->m_style=style;
	m_listener->insertChart(pos, *it->second);
	return true;
}

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
