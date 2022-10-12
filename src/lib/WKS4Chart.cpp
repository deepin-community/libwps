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

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WKSChart.h"
#include "WKSContentListener.h"
#include "WPSEntry.h"
#include "WPSFont.h"
#include "WPSPosition.h"

#include "WKS4.h"

#include "WKS4Chart.h"

namespace WKS4ChartInternal
{

///////////////////////////////////////////////////////////////////
//! the chart of a WKS4 Pro Dos
class Chart final : public WKSChart
{
public:
	//! constructor
	explicit Chart(WKS4Chart &parser, RVNGInputStreamPtr const &input)
		: WKSChart()
		, m_fileType(-1)
		, m_axisZoneFound(false)
		, m_use2D(false)
		, m_seriesStylesZoneFound(false)
		, m_parser(parser)
		, m_input(input)
	{
	}
	//! send the zone content (called when the zone is of text type)
	void sendContent(TextZone const &zone, WPSListenerPtr &listener) const final;
	//! check if the chart has no serie
	bool empty() const
	{
		for (int i=0; i<6; ++i)
		{
			if (const_cast<Chart *>(this)->getSerie(i,false))
				return false;
		}
		return true;
	}
	//! the chart type
	int m_fileType;
	//! flag to know if we have found the axis zone
	bool m_axisZoneFound;
	//! flag to know if we print line/surface data
	bool m_use2D;
	//! flag to know if we have found the series zone
	bool m_seriesStylesZoneFound;

	//! small struct used to defined the serie data
	struct SerieData
	{
		//! constructor
		SerieData()
			: m_type(-1)
		{
			for (auto &c : m_ids) c=-1;
		}
		//! the serie secondary type(used to swap line<->bar)
		int m_type;
		//! the serie color
		int m_ids[3];
	};

	//! the serie local data
	SerieData m_seriesData[6];
protected:
	//! the parser
	WKS4Chart &m_parser;
	//! the input
	RVNGInputStreamPtr m_input;
};

void Chart::sendContent(TextZone const &zone, WPSListenerPtr &listener) const
{
	if (!listener.get())
	{
		WPS_DEBUG_MSG(("WKS4ChartInternal::Chart::sendContent: no listener\n"));
		return;
	}
	long pos = m_input->tell();
	listener->setFont(zone.m_font);
	bool sendText=false;
	for (auto const &e : zone.m_textEntryList)
	{
		if (!e.valid()) continue;
		if (sendText) listener->insertEOL(true);
		m_parser.sendText(e);
		sendText=true;
	}
	m_input->seek(pos, librevenge::RVNG_SEEK_SET);
}

//! the state of WKS4Chart
struct State
{
	//! constructor
	State()
		: m_eof(-1)
		, m_version(-1)
		, m_idToColorMap()
		, m_chartList()
	{
	}
	//! returns a color corresponding to an id
	bool getColor(int id, WPSColor &color) const;
	//! returns the pattern corresponding to a pattern id between 0 and 15
	static bool getPattern(int id, WPSGraphicStyle::Pattern &pattern);
	//! the last file position
	long m_eof;
	//! the file version
	int m_version;
	//! a map id -> color
	std::map<int, WPSColor> m_idToColorMap;
	//! list of chart
	std::vector<std::shared_ptr<Chart> > m_chartList;
};

bool State::getColor(int id, WPSColor &color) const
{
	if (m_idToColorMap.empty())
	{
		// FIXME: find the complete table
		static const uint32_t colorMap[]=
		{
			0, 0,
			27, 0x00007B,
			15, 0x007B00,
			21, 0x007B7B,
			3, 0x7B0000,
			33, 0x7B007B,
			9, 0x7B7B00,
			38, 0x7B7B7B,
			39, 0x393939,
			26, 0x0000FF,
			14, 0x00FF00,
			20, 0x00FFFF,
			2, 0xFF0000,
			32, 0xFF00FF,
			8, 0xFFFF00
		};
		for (size_t i=0; i+1<WPS_N_ELEMENTS(colorMap); i+=2)
			const_cast<State *>(this)->m_idToColorMap[int(colorMap[i])]=WPSColor(colorMap[i+1]);
	};
	if (m_idToColorMap.find(id)==m_idToColorMap.end())
	{
		WPS_DEBUG_MSG(("WKS4ChartInternal::State::getColor(): unknown color id: %d\n",id));
		return false;
	}
	color = m_idToColorMap.find(id)->second;
	return true;
}

bool State::getPattern(int id, WPSGraphicStyle::Pattern &pat)
{
	if (id<0 || id>14)
	{
		WPS_DEBUG_MSG(("WKS4ChartInternal::State::getPattern(): unknown pattern id: %d\n",id));
		return false;
	}
	/* none(lineColor), solid, dense, medium
	   sparse, ==, ||, \\,
	   //, ++, XX, // dark,
	   // light, \\ dark, \\light
	 */
	static const uint16_t  patterns[]=
	{
		0xffff, 0xffff, 0xffff, 0xffff, 0x0000, 0x0000, 0x0000, 0x0000, 0x0001, 0x0100, 0x0001, 0x0100, 0x8855, 0x2255, 0x8855, 0x2255, // 0-3
		0xff77, 0xffdd, 0xff77, 0xffdd, 0x00ff, 0xff00, 0x00ff, 0xff00, 0xcc00, 0xcc00, 0xcc00, 0xcc00, 0xf1f8, 0x7c3e, 0x1f8f, 0xc7e3, // 4-7
		0xf8f1, 0xe3c7, 0x8f1f, 0x3e7c, 0xefef, 0xefef, 0xefef, 0x0000, 0xeef5, 0xfbf5, 0xee5f, 0xbf5f, 0xf0e1, 0xc387, 0x0f1e, 0x3c78, // 8-11
		0xefdf, 0xbf7f, 0xfefd, 0xfbf7, 0xf078, 0x3c1e, 0x0f87, 0xc3e1, 0xdfef, 0xf7fb, 0xfdfe, 0x7fbf // 12-14
	};
	pat.m_dim=Vec2i(8,8);
	uint16_t const *ptr=&patterns[4*id];
	pat.m_data.resize(8);
	for (size_t i=0; i < 8; i+=2)
	{
		uint16_t val=*(ptr++);
		pat.m_data[i]=static_cast<unsigned char>((val>>8) & 0xFF);
		pat.m_data[i+1]=static_cast<unsigned char>(val & 0xFF);
	}
	return true;
}
}

// constructor, destructor
WKS4Chart::WKS4Chart(WKS4Parser &parser)
	: m_input(parser.getInput())
	, m_listener()
	, m_mainParser(parser)
	, m_state(new WKS4ChartInternal::State)
	, m_asciiFile(parser.ascii())
{
}

WKS4Chart::~WKS4Chart()
{
}

int WKS4Chart::getNumCharts() const
{
	int num=0;
	for (auto chart : m_state->m_chartList)
	{
		if (chart && !chart->empty())
			++num;
	}
	return num;
}

void WKS4Chart::resetInput(RVNGInputStreamPtr const &newInput)
{
	m_input=newInput;
}

int WKS4Chart::version() const
{
	if (m_state->m_version<0)
		m_state->m_version=m_mainParser.version();
	return m_state->m_version;
}

bool WKS4Chart::checkFilePosition(long pos)
{
	if (m_state->m_eof < 0)
	{
		long actPos = m_input->tell();
		m_input->seek(0, librevenge::RVNG_SEEK_END);
		m_state->m_eof=m_input->tell();
		m_input->seek(actPos, librevenge::RVNG_SEEK_SET);
	}
	return pos <= m_state->m_eof;
}

void WKS4Chart::updateChart(WKS4ChartInternal::Chart &chart)
{
	int const vers=version();
	auto const creator=m_mainParser.creator();
	if (!chart.m_axisZoneFound)
	{
		// old msworks/lotus file
		for (int i=0; i<2; ++i) // show axis
			chart.getAxis(i).m_type= WKSChart::Axis::A_Numeric;
	}
	//! update the chart type and serie type
	if (chart.m_fileType==3 || chart.m_fileType==5) chart.m_dataStacked=true;
	if (chart.m_fileType==8) chart.m_dataVertical=true;
	auto serieType = chart.m_type;
	auto pointType = WKSChart::Serie::P_None;
	if (serieType==WKSChart::Serie::S_Scatter)
		pointType = vers>=3 ? WKSChart::Serie::P_Circle : WKSChart::Serie::P_Automatic;
	for (int i=0; i<6; ++i)
	{
		auto *serie=chart.getSerie(i, false);
		if (!serie) continue;
		serie->m_type=serieType;
		serie->m_pointType=pointType;
		serie->m_style.m_lineWidth=1;

		auto const &serieData=chart.m_seriesData[i];
		if (vers>=3 && (chart.m_fileType==3 || chart.m_fileType==4) && chart.m_use2D)
			chart.m_type = serie->m_type=WKSChart::Serie::S_Area;
		if (serieData.m_type!=-1)
		{
			if ((serieData.m_type&1)==0 && serieType==WKSChart::Serie::S_Bar)
				serie->m_type=WKSChart::Serie::S_Line;
			if (serieData.m_type&2)
				serie->m_useSecondaryY=true;
		}
		if (creator==libwps::WPS_MSWORKS && serie->m_type==WKSChart::Serie::S_Line)
			serie->m_pointType=WKSChart::Serie::P_Automatic;
	}

	//! times to update the color
	for (int i=0; i<6; ++i)
	{
		auto *serie= chart.getSerie(i, false);
		if (!serie) continue;
		auto &serieData = chart.m_seriesData[i];

		if (serieData.m_ids[2]>=0 && (serie->m_type==WKSChart::Serie::S_Line || serie->m_type==WKSChart::Serie::S_Radar || serie->m_type==WKSChart::Serie::S_Scatter))
		{
			if (serieData.m_ids[2]<10)
			{
				WKSChart::Serie::PointType const fPointType[]=
				{
					WKSChart::Serie::P_None, WKSChart::Serie::P_Circle, WKSChart::Serie::P_Square, WKSChart::Serie::P_Diamond,
					WKSChart::Serie::P_Asterisk, WKSChart::Serie::P_Circle/*hollow*/, WKSChart::Serie::P_Square/*hollow*/, WKSChart::Serie::P_Diamond/*hollow*/,
					WKSChart::Serie::P_Plus/*small full circle*/, WKSChart::Serie::P_Horizontal_Bar
				};
				serie->m_pointType=fPointType[serieData.m_ids[2]];
				// fixme: make hollow circle, ... and correct the dot
			}
			else
			{
				WPS_DEBUG_MSG(("WKS4Chart::sendChart: find unknown point type %d\n", serieData.m_ids[2]));
			}
		}
		bool has0D=serie->m_pointType!=WKSChart::Serie::P_None;
		bool has1D=serie->is1DStyle();
		bool has2D=!serie->is1DStyle();
		if (serieData.m_ids[1]>=0)
			has1D=serie->m_type!=WKSChart::Serie::S_Line || serieData.m_ids[1]!=0;
		serie->m_style.m_lineWidth=has1D ? 1 : 0;

		int col=serieData.m_ids[0];
		if (col<0)
		{
			if (vers<=2)
			{
				// old msworks/lotus file: set the series' color to default wk3 color
				// (even if there are in fact grey color)
				int const defColor[6]= {26,14,2,20,32,8};
				col= defColor[i];
			}
			else
			{
				// msworks3. windows file
				int const defColor[6]= {2,14,26,8,20,32};
				col= defColor[i];
			}
		}
		WPSColor color=WPSColor(0,0,255);
		m_state->getColor(col, color);
		WPSGraphicStyle::Pattern pattern;
		if (serieData.m_ids[1]>0 && m_state->getPattern(serieData.m_ids[1],pattern))
		{
			pattern.m_colors[0]=color;
			pattern.m_colors[1]=WPSColor::white();
			if (has0D || has2D)
				serie->setPrimaryPattern(pattern);
			WPSColor finalColor;
			if (has1D && pattern.getUniqueColor(finalColor))
				serie->m_style.m_lineColor=finalColor;
			continue;
		}
		if (serieData.m_ids[1]>0)
		{
			WPS_DEBUG_MSG(("QuattroDosChart::sendCharts: oops, can not find pattern %d\n", serieData.m_ids[1]));
		}
		if (has1D || serieData.m_ids[1]==0)
			serie->m_style.m_lineColor=color;
		if (has0D || (has2D && serieData.m_ids[1]!=0))
			serie->m_style.setSurfaceColor(color);
	}
}

bool WKS4Chart::sendCharts()
{
	if (m_listener.get() == nullptr)
	{
		WPS_DEBUG_MSG(("WKS4Chart::sendChart: I can not find the listener\n"));
		return false;
	}
	Vec2i actPos(0,0);
	int actSquare=0;
	for (auto chart : m_state->m_chartList)
	{
		if (!chart || chart->empty()) continue;
		updateChart(*chart);
		WPSPosition pos(Vec2f(float(512*actPos[0]),float(350*actPos[1])), Vec2f(512,350), librevenge::RVNG_POINT);
		pos.m_anchorTo = WPSPosition::Page;
		// the chart is in fact a 1024*7?? windows scaled back on the chosen cells
		chart->m_dimension=Vec2f(512,350);
		m_listener->insertChart(pos, *chart);
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

bool WKS4Chart::sendText(WPSEntry const &entry)
{
	if (m_listener.get() == nullptr)
	{
		WPS_DEBUG_MSG(("WKS4Chart::sendText: I can not find the listener\n"));
		return false;
	}
	if (!entry.valid())
		return true;
	m_input->seek(entry.begin(), librevenge::RVNG_SEEK_SET);
	m_listener->insertUnicodeString(libwps_tools_win::Font::unicodeString(m_input.get(), static_cast<unsigned long>(entry.length()), m_mainParser.getDefaultFontType()));
	return true;
}
////////////////////////////////////////////////////////////
// low level

////////////////////////////////////////////////////////////
// general
////////////////////////////////////////////////////////////

bool WKS4Chart::readChart()
{
	libwps::DebugStream f;

	long pos = m_input->tell();
	auto type = long(libwps::readU16(m_input));
	if (type != 0x2D && type != 0x2e)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChart: not a chart definition\n"));
		return false;
	}
	auto sz = long(libwps::readU16(m_input));
	int normalSz = (type == 0x2D) ? 0x1b5 : 0x1c5;
	if (sz < normalSz)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChart: chart definition too short\n"));
		ascii().addPos(pos);
		ascii().addNote("Entries(ChartDef):###");
		return true;
	}

	f << "Entries(ChartDef):sz=" << sz << ",";
	std::shared_ptr<WKS4ChartInternal::Chart> chart(new WKS4ChartInternal::Chart(*this, m_input));
	if (type == 0x2e)
	{
		librevenge::RVNGString name;
		if (!m_mainParser.readCString(name,16))
			f << "##sSz,";
		else if (!name.empty())
		{
			chart->m_name=name;
			f << "name=" << name.cstr() << ',';
		}
		m_input->seek(pos+4+16, librevenge::RVNG_SEEK_SET);
	}
	bool hasLegend=false;
	for (int i = 0; i < 13; i++) // 0: X-serie, 1-6: data[series%d], 7-12:label serie%d
	{
		WKSChart::Position ranges[2];
		for (auto &range : ranges)
		{
			int dim[2];
			for (auto &d : dim) d = int(libwps::read16(m_input));
			if (dim[0]==-1)
				continue;
			range = WKSChart::Position(Vec2i(dim[0],dim[1]),m_mainParser.getSheetName(0));
		}
		if (ranges[0].valid(ranges[1]))
		{
			f << "z" << i << "=" << ranges[0] << ":" << ranges[1] << ",";
			if (i==0)
			{
				auto &axis=chart->getAxis(0);
				axis.m_labelRanges[0]=ranges[0];
				axis.m_labelRanges[1]=ranges[1];
			}
			else if (i<=6)
			{
				auto *serie=chart->getSerie(i-1, true);
				serie->m_ranges[0]=ranges[0];
				serie->m_ranges[1]=ranges[1];
			}
			else
			{
				auto *serie=chart->getSerie(i-7, false);
				if (serie)
				{
					serie->m_labelRanges[0]=ranges[0];
					serie->m_labelRanges[1]=ranges[1];
				}
			}
		}
	}
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	pos=m_input->tell();
	f.str("");
	f << "ChartDef-A:"; // GraphDef0
	char const *axisNames[]= {"X", "Y", "Y2"};
	auto const &chartType=chart->m_fileType=int(libwps::readU8(m_input));
	WKSChart::Serie::Type serieType=WKSChart::Serie::S_Bar;
	switch (chartType)
	{
	case 0: // XY
		serieType=WKSChart::Serie::S_Scatter;
		break;
	case 1: // basic column
		break;
	case 2: // pie
		serieType=WKSChart::Serie::S_Circle;
		break;
	case 6: // radar (checkme: or circle)
		serieType=WKSChart::Serie::S_Radar;
		break;
	case 3: // line/area stacked
		serieType=WKSChart::Serie::S_Line;
		break;
	case 4: // line/area
		serieType=WKSChart::Serie::S_Line;
		break;
	case 5: // basic column stacked
		break;
	case 7: // min-max
		serieType=WKSChart::Serie::S_Stock;
		break;
	case 8: // basic bar
		break;
	default:
		WPS_DEBUG_MSG(("QuattroDosChart::readChart: unknown chart type\n"));
		f << "###";
	}
	if (chartType) f << "type=" << chartType << ",";

	chart->m_type=serieType;
	auto val=int(libwps::readU8(m_input));
	f << "grid=";
	if (val&1)
		f << "X";
	else
		chart->getAxis(0).m_showGrid=false;
	if (val&2)
		f << "Y";
	else
		chart->getAxis(1).m_showGrid=false;
	if (val&0xFC)
		f << "[##" << std::hex << (val&0xFC) << std::dec;
	f << ",";
	val=int(libwps::readU8(m_input));
	// bool useColor=val!=0; USEME
	if (val==0)
		f << "use[color]=no,";
	else if (val!=0xFF)
		f << "use[color]=###"<< val << ",";
	f << "unkn=[";
	for (int i=0; i<6; ++i)
	{
		val=int(libwps::readU8(m_input));
		if (val)
			f << val << ",";
		else
			f << "_,";
	}
	f << "],";
	f << "align[serie]=[";
	for (int i=0; i<6; ++i) // serie, label position, safe to ignore ?
	{
		val=int(libwps::readU8(m_input));
		if (val<6)
		{
			char const *wh[]= {"center", "left", "above", "right", "below", "none"};
			f << wh[val] << ",";
		}
		else
			f << "##" << val << ",";
	}
	f << "],";
	for (int i=0; i<2; ++i)
	{
		auto &axis=chart->getAxis(i);
		f << "axis" << axisNames[i] << "=[";
		val=int(libwps::readU8(m_input));
		if (val==1)
		{
			f << "scale[manual],";
			axis.m_automaticScaling=false;
		}
		else if (val)
			f << "scale=##" << val << ",";
		for (int j=0; j<2; ++j)
		{
			long actPos=m_input->tell();
			double value;
			bool isNaN;
			if (!libwps::readDouble8(m_input, value, isNaN))
			{
				m_input->seek(actPos+8, librevenge::RVNG_SEEK_SET);
				f << "##value,";
			}
			else
			{
				if (value<0 || value>0)
					f << (j==0 ? "low" : "high") << "=" << value << ",";
				axis.m_scaling[j]=float(value);
			}
		}
		f << "],";
	}
	m_input->seek(pos+49, librevenge::RVNG_SEEK_SET);
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	pos=m_input->tell();
	f.str("");
	f << "ChartDef-names:";
	for (int i = 0; i < 10; i++)
	{
		long actPos = m_input->tell();
		int dataSz = i < 4 ? 40 : 20;
		librevenge::RVNGString name("");
		if (!m_mainParser.readCString(name,dataSz))
			f << "##sSz,";
		if (name.empty())
		{
			m_input->seek(actPos+dataSz, librevenge::RVNG_SEEK_SET);
			continue;
		}
		if (i<2)
		{
			WPSEntry entry;
			entry.setBegin(actPos);
			entry.setEnd(m_input->tell());
			auto *textZone=chart->getTextZone(i==0 ? WKSChart::TextZone::T_Title : WKSChart::TextZone::T_SubTitle, true);
			textZone->m_contentType=WKSChart::TextZone::C_Text;
			textZone->m_textEntryList.push_back(entry);
			f << (i==0 ? "title" : "subTitle") << "=" << name.cstr() << ",";
		}
		else if (i<4)
		{
			chart->getAxis(i-2).m_title=name;
			f << (i==2 ? "x" : "y") << "Title=" << name.cstr() << ",";
		}
		else
		{
			auto *serie=chart->getSerie(i-4, false);
			if (serie)
			{
				serie->m_legendText=name;
				hasLegend=true;
			}
			f << "serie" << i-4 <<  "=" << name.cstr() << ",";
		}
		m_input->seek(actPos+dataSz, librevenge::RVNG_SEEK_SET);
	}
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	pos=m_input->tell();
	f.str("");
	f << "ChartDef-B:";
	for (int i=0; i<2; ++i) // axisX, axisY: format: USEME
	{
		val=int(libwps::readU8(m_input));
		if (val) f << "fmt" << axisNames[i] << "=" << val << ",";
	}
	for (int i=0; i<2; ++i) // axisX, axisY, number of tick: USEME
	{
		val=int(libwps::readU8(m_input));
		if (val) f << "num[tick" << axisNames[i] << "]=" << val << ",";
	}
	if (sz != normalSz) ascii().addDelimiter(m_input->tell(),'|');
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	// time to update the legend
	if (hasLegend)
	{
		chart->getLegend().m_show=true;
		chart->getLegend().m_autoPosition=true;
		chart->getLegend().m_relativePosition=WPSBorder::RightBit;
	}
	m_state->m_chartList.push_back(chart);
	return true;
}

bool WKS4Chart::readChartName()
{
	libwps::DebugStream f;

	long pos = m_input->tell();
	auto type = long(libwps::read16(m_input));
	if (type != 0x41)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartName: not a chart name\n"));
		return false;
	}
	auto sz = long(libwps::readU16(m_input));

	if (sz < 0x10)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartName: chart name is too short\n"));
		ascii().addPos(pos);
		ascii().addNote("Entries(ChartName):###");
		return true;
	}

	f << "Entries(ChartName):";
	librevenge::RVNGString name;
	if (!m_mainParser.readCString(name,16))
		f << "##name,";
	else if (!name.empty())
		f << name.cstr() << ",";
	if (m_state->m_chartList.empty())
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartName: can not find the current chart\n"));
	}
	else
		m_state->m_chartList.back()->m_name = name;
	if (sz != 0x10) ascii().addDelimiter(pos+4+sz, '#');

	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

bool WKS4Chart::readChartDim()
{
	libwps::DebugStream f;

	long pos = m_input->tell();
	auto type = long(libwps::read16(m_input));
	if (type != 0x5435)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartDim: not a chart dim\n"));
		return false;
	}
	auto sz = long(libwps::readU16(m_input));

	if (sz != 0xc)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartDim: chart dim is too short\n"));
		ascii().addPos(pos);
		ascii().addNote("Entries(ChartDim):###");
		return true;
	}

	f << "Entries(ChartDim):";
	std::shared_ptr<WKS4ChartInternal::Chart> chart;
	if (m_state->m_chartList.empty())
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartDim: can not find the current chart\n"));
	}
	else
		chart = m_state->m_chartList.back();
	int val;
	for (int i=0; i<2; ++i)   // fl0=42|49, fl1=1|9
	{
		val=int(libwps::readU8(m_input));
		// checkme unsure
		if (i==0)
		{
			if (val&1) f << "display[value],";
			val &= 0xfe;
		}
		else
		{
			if (val&2)
			{
				f << "area,";
				if (chart) chart->m_use2D=true;
			}
			if (val&4) f << "gridY,";
			if (val&8) f << "stackX,"; // USEME: display the sum of values in each series
			if (val&0x40) f << "display[serie,name],";
			val &= 0xb1;
		}
		if (!val) continue;
		f << "fl" << i << "=" << std::hex << val << std::dec << ",";
	}
	val=int(libwps::read16(m_input)); // 0
	if (val) f << "f0=" << val << ",";
	for (int i=0; i<2; ++i)
	{
		int dim[2];
		for (auto &d : dim) d=int(libwps::read16(m_input));
		if (dim[0]!=0 || dim[1]!=0) // maybe some position
			f << "pos" << i << "=" << Vec2i(dim[0],dim[1]) << ",";
	}
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

bool WKS4Chart::readChartFont()
{
	libwps::DebugStream f;

	long pos = m_input->tell();
	auto type = long(libwps::read16(m_input));
	if (type != 0x5440)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartFont: not a chart font\n"));
		return false;
	}
	auto sz = long(libwps::readU16(m_input));
	long endPos = pos+4+sz;
	if (sz < 0x22)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartFont: chart font is too short\n"));
		ascii().addPos(pos);
		ascii().addNote("Entries(ChartFont):###");
		return true;
	}
	f << "Entries(ChartFont):";
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	auto N=int(sz/0x22);

	for (int i=0; i<N; ++i)
	{
		pos=m_input->tell();
		f.str("");
		f << "ChartFont-" << i << ":";
		auto fl = int(libwps::readU8(m_input));
		if (fl!=0x20)
			f << "flag=" << std::hex << fl << std::dec << ",";
		librevenge::RVNGString name;
		if (!m_mainParser.readCString(name,32))
			f << "##name,";
		else if (!name.empty())
			f << name.cstr() << ",";
		m_input->seek(pos+33, librevenge::RVNG_SEEK_SET);
		auto fl2 = int(libwps::readU8(m_input)); // always 0
		if (fl2) f << "flag2=" << std::hex << fl2 << std::dec << ",";
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
	}
	if (m_input->tell() != endPos)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartFont: find extra data\n"));
		ascii().addPos(m_input->tell());
		ascii().addNote("ChartFont:###extra");
	}
	return true;
}

bool WKS4Chart::readChart3D()
{
	libwps::DebugStream f;

	long pos = m_input->tell();
	auto type = long(libwps::read16(m_input));
	if (type != 0x5444)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChart3d: not a chart 3d\n"));
		return false;
	}
	auto sz = long(libwps::readU16(m_input));
	if (sz != 4)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChart3d: chart 3d size is unknown\n"));
		ascii().addPos(pos);
		ascii().addNote("Entries(Chart3d):###");
		return true;
	}
	f << "Entries(Chart3D):";
	std::shared_ptr<WKS4ChartInternal::Chart> chart;
	if (m_state->m_chartList.empty())
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChart3D: can not find the current chart\n"));
	}
	else
		chart = m_state->m_chartList.back();
	int dim[2]; // 0,0: 2d or f,14: 3d
	for (auto &d: dim) d=int(libwps::read16(m_input));
	if (dim[0]!=dim[1])
	{
		f << "dim=" << Vec2i(dim[0],dim[1]) << ",";
		if (chart) chart->m_is3D = true;
	}
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

bool WKS4Chart::readChart2Font()
{
	libwps::DebugStream f;

	long pos = m_input->tell();
	auto type = long(libwps::read16(m_input));
	if (type != 0x5484)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChart2Font: not a chart2 font\n"));
		return false;
	}
	auto sz = long(libwps::readU16(m_input));
	long endPos = pos+4+sz;
	if (sz < 0x23)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChart2Font: chart2 font is too short\n"));
		ascii().addPos(pos);
		ascii().addNote("Entries(ChartFont):###");
		return true;
	}
	f << "Entries(Chart2Font):";
	// USEME
	auto nbElt = int(sz/0x23);
	for (int i = 0; i < nbElt; i++)
	{
		long actPos = m_input->tell();
		f << "ft" << i << "=[";
		auto fl = int(libwps::readU8(m_input));
		f << "flag=" << std::hex << fl << std::dec << ",";
		librevenge::RVNGString name;
		if (!m_mainParser.readCString(name,32))
			f << "##name,";
		else if (!name.empty())
			f << name.cstr() << ",";
		m_input->seek(actPos+33, librevenge::RVNG_SEEK_SET);
		auto fl2 = int(libwps::readU8(m_input));
		if (fl2) f << ",#flag2=" << std::hex << fl2 << std::dec;
		auto fl3 = int(libwps::readU8(m_input)); // check me ?
		if (fl3) f << ",sz="  << fl3/2;
		f << "],";
	}
	if (long(m_input->tell()) != endPos)
		ascii().addDelimiter(m_input->tell(),'#');
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

bool WKS4Chart::readChartLimit()
{
	libwps::DebugStream f;

	long pos = m_input->tell();
	auto type = long(libwps::read16(m_input));
	// USEME
	if (type == 0x5480) f << "Entries(ChartBegin)";
	else if (type == 0x5481) f << "Entries(ChartEnd)";
	else
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartLimit: not a chart limit\n"));
		return false;
	}
	auto sz = long(libwps::readU16(m_input));
	if (sz != 0) ascii().addDelimiter(pos+4,'#');

	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	return true;
}

bool WKS4Chart::readChartAxis()
{
	libwps::DebugStream f;

	long pos = m_input->tell();
	auto type = long(libwps::read16(m_input));
	if (type != 0x5414)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartAxis: not a chart ???\n"));
		return false;
	}
	auto sz = long(libwps::readU16(m_input));
	long endPos = pos+4+sz;
	if (sz < 0x8d)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartAxis: chart axis zone is too short\n"));
		ascii().addPos(pos);
		ascii().addNote("Entries(ChartAxis):###");
		return true;
	}
	std::shared_ptr<WKS4ChartInternal::Chart> chart;
	if (m_state->m_chartList.empty() || m_state->m_chartList.back()->m_axisZoneFound)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartAxis: can not find the current chart\n"));
	}
	else
	{
		chart = m_state->m_chartList.back();
		chart->m_axisZoneFound = true;
	}
	f << "Entries(ChartAxis):";
	auto val= int(libwps::readU16(m_input)); //f0={00|20}{23|30|32|34|38|60|70}
	if (val & 0x10)
	{
		if (chart)
		{
			chart->getLegend().m_show=true;
			chart->getLegend().m_autoPosition=true;
			chart->getLegend().m_relativePosition=WPSBorder::RightBit;
		}
		f << "legend[show],";
	}
	if (val & 0x40)
		f << "border[show],";
	val &= 0xFFAF;
	if (val) f << "f0=" << std::hex << val << std::dec << ",";
	for (int i = 1; i < 9; i++)
	{
		/* I find  f5=7 */
		val = int(libwps::readU16(m_input));
		if (i==5)
		{
			f << "X=[";
			if ((val&1)==0)
				f << "min=manual,";
			if ((val&2)==0)
				f << "max=manual,";
			if ((val&4)==0)
				f << "increment=manual,";

			if (val&8)
				f << "log,";
			if (val&0x10)
				f << "grid,";
			f << "],";
			if (chart)
			{
				auto &axis=chart->getAxis(0);
				//axis.m_automaticScaling=((val&7)==0);
				if (val&0x8)
					axis.m_type= WKSChart::Axis::A_Logarithmic;
				else
					axis.m_type= WKSChart::Axis::A_Numeric;
			}
			val &= 0xFFE0;
		}
		if (val) f << "f" << i << "=" << std::hex << val << std::dec << ",";
	}
	val=int(libwps::readU8(m_input));
	if (val) f << "f9=" << std::hex << val << std::dec << ",";
	val=int(libwps::readU8(m_input));
	f << "Y=[";
	if ((val&1)==0)
		f << "min=manual,";
	if ((val&2)==0)
		f << "max=manual,";
	if ((val&4)==0)
		f << "increment=manual,";
	if (val&8)
		f << "log,";
	if (val&0x10)
		f << "grid,";
	if (chart)
	{
		auto &axis=chart->getAxis(1);
		//axis.m_automaticScaling=((val&7)==0);
		if (val&0x8)
			axis.m_type= WKSChart::Axis::A_Logarithmic;
		else
			axis.m_type= WKSChart::Axis::A_Numeric;
	}
	val &= 0xE0;
	if (val) f << "#" << std::hex << val << std::dec << ",";
	f << "],";
	for (int i=0; i<2; ++i)
	{
		char const *wh[]= {"Y", "Ysecond"};
		val = int(libwps::read16(m_input)); // 0,1,-1
		if (val==0)
			f << "type[" << wh[i] << "]=normal,";
		else if (val==1)
		{
			f << "type[" << wh[i] << "]=stacked,";
			if (chart && i==0)  chart->m_dataStacked=true;
		}
		else if (val==2)
		{
			f << "type[" << wh[i] << "]=100%,";
			if (chart && i==0)  chart->m_dataPercentStacked=true;
		}
		else if (val==3)
			f << "type[" << wh[i] << "]=hiLo,";
		else if (val==4)
		{
			f << "type[" << wh[i] << "]=3Dpers,";
			if (chart && i==0)
				chart->m_is3DDeep=true;;
		}
		else if (val!=-1) // -1: unused?
			f << "##type[" << wh[i] << "]=" << val << ",";
	}
	f << "YSecond=[";
	for (int i=0; i<3; ++i)  // Ysecond: min, max, increment
	{
		double value;
		bool isNaN;
		long actPos = m_input->tell();
		if (!libwps::readDouble8(m_input, value, isNaN))
		{
			m_input->seek(actPos+8, librevenge::RVNG_SEEK_SET);
			f << "##value,";
		}
		else
		{
			if (chart && i<2) // use incremenet
				chart->getAxis(2).m_scaling[i]=float(value);
			if (value<0 || value>0)
			{
				char const *wh[]= {"low", "high", "incr"};
				f << wh[i] << "=" << value << ",";
			}
		}
	}
	val = int(libwps::readU8(m_input));
	if ((val&1)==0)
		f << "min=manual,";
	if ((val&2)==0)
		f << "max=manual,";
	if ((val&4)==0)
		f << "increment=manual,";
	if (val&8)
		f << "log,";
	if (val&0x10)
		f << "grid,";
	if (chart)
	{
		auto &axis=chart->getAxis(2);
		//axis.m_automaticScaling=((val&7)==0);
		if (val&0x8)
			axis.m_type= WKSChart::Axis::A_Logarithmic;
		else
			axis.m_type= WKSChart::Axis::A_Numeric;
	}
	val &= 0xE0;
	if (val) f << "#" << std::hex << val << std::dec << ",";
	f << "],";
	for (int i=0; i<2; ++i)
	{
		f << (i==0 ? "title" : "other") << "=[";
		val = int(libwps::readU8(m_input));
		if (val&0x80)
		{
			f << "bold,";
			val &=0x7f;
		}
		if (val) f << "fmt=" << val << ",";
		val = int(libwps::readU8(m_input));
		if (val) f << "sz=" << val << ",";
		val = int(libwps::readU8(m_input));
		if (val&1)
			f << "it,";
		if (val&2)
			f << "underline,";
		if (val&4)
			f << "strike,";
		/* does not use the standard color id, but
		   0:auto,1:black,2:blue,3:cyan,4:green,5:magenta
		   red, yellow, gray, white, dark blue, dark cyan,
		   dark green, dark magenta, dark red, light gray
		 */
		val>>=3;
		if (val) f << "col=" << val << ",";
		f << "],";
	}
	val = int(libwps::readU8(m_input));
	if (val) f << "g0=" << val << ",";
	for (int i = 0; i < 5; i++)
	{
		/* I find g2=5|7 */
		val = int(libwps::readU16(m_input));
		if (!val) continue;
		f << "g" << i << "=" << std::hex << val << std::dec << ",";
	}
	val = int(libwps::readU8(m_input)); // 7
	if (val) f << "g5=" << std::hex << val << std::dec << ",";
	long actPos = m_input->tell();
	librevenge::RVNGString name;
	if (!m_mainParser.readCString(name,40))
		f << "##name,";
	else if (!name.empty())
	{
		if (chart)
			chart->getAxis(2).m_title=name;
		f << "ySecondTitle=" << name.cstr() << ",";
	}
	m_input->seek(actPos+40, librevenge::RVNG_SEEK_SET);
	for (int i = 0; i < 5; i++)
	{
		// some cell zone ?? :
		int dim[2];
		for (auto &d : dim) d = int(libwps::read16(m_input));
		if (Vec2i(dim[0],dim[1]) != Vec2i(-1,0))
			f << "cell" << i << "=C" << Vec2i(dim[0],dim[1]) << ",";
	}
	val = int(libwps::read16(m_input));
	if (val) f << "h0=" << val << ","; // 0 or 2

	// look like TL?=1,1.25,BR?=9,6,pageLength?=11,8.5
	f << "dim?=[";
	for (int i = 0; i < 6; i++)
	{
		val = int(libwps::read16(m_input));
		f << val/1440. << ",";
	}
	f << "]";

	if (long(m_input->tell()) != endPos) ascii().addDelimiter(m_input->tell(), '#');
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	return true;
}

bool WKS4Chart::readChartSeries()
{
	libwps::DebugStream f;

	long pos = m_input->tell();
	auto type = long(libwps::read16(m_input));
	if (type != 0x5415)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartSeries: not a series' data\n"));
		return false;
	}
	auto sz = long(libwps::readU16(m_input));

	if (sz < 0x1e)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartSeries: chart definition too short\n"));
		ascii().addPos(pos);
		ascii().addNote("Entries(ChartSeries):###");
		return true;
	}

	std::shared_ptr<WKS4ChartInternal::Chart> chart;
	if (m_state->m_chartList.empty())
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartSeries: can not find the current chart\n"));
	}
	else
		chart = m_state->m_chartList.back();
	f << "Entries(ChartSeries):";
	for (int i = 0; i < 6; i++)
	{
		f << "S" << i << "=[";
		WKSChart::Serie *serie=chart ? chart->getSerie(i, false) : nullptr;
		const auto serieType = int(libwps::readU8(m_input)); // checkme, unsure
		if (chart)
			chart->m_seriesData[i].m_type = serieType;
		if (serieType&1)
			f << "bar,";
		else if (serie && serie->m_type==WKSChart::Serie::S_Bar)
			f << "line,";
		if (serieType&2)
			f << "Ysecond,";
		if (serieType & 0xFC)
			f << "#type=" << std::hex << (serieType&0xFC) << ",";
		int dim[2];
		for (auto &d : dim) d = int(libwps::read16(m_input));
		if (Vec2i(dim[0],dim[1]) != Vec2i(-1,0))
		{
			WKSChart::Position range(Vec2i(dim[0],dim[1]),m_mainParser.getSheetName(0));
			f << range << ",";
			if (serie)
				serie->m_legendRange=range;
		}
		f << "],";
	}

	if (sz != 0x1e) ascii().addDelimiter(m_input->tell(),'#');
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	return true;
}

bool WKS4Chart::readChartSeriesStyles()
{
	libwps::DebugStream f;

	long pos = m_input->tell();
	auto type = long(libwps::read16(m_input));
	if (type != 0x5416)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartSeriesStyles: not a series styles\n"));
		return false;
	}
	auto sz = long(libwps::readU16(m_input));

	if (sz%6)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartSeriesStyles: chart definition too short\n"));
		ascii().addPos(pos);
		ascii().addNote("Entries(ChartSerStyl):###");
		return true;
	}

	std::shared_ptr<WKS4ChartInternal::Chart> chart;
	if (m_state->m_chartList.empty() || m_state->m_chartList.back()->m_seriesStylesZoneFound)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartSeriesStyles: can not find the current chart\n"));
	}
	else
	{
		chart = m_state->m_chartList.back();
		chart->m_seriesStylesZoneFound = true;
	}
	auto N=int(sz/6);
	f << "Entries(ChartSerStyl):";
	for (int i = 0; i<N; ++i)
	{
		auto id=int(libwps::readU16(m_input));
		f << "S" << id << "=[";
		int format[3]= {id,0,0};
		for (int j=0; j<3; ++j)   // f1=0|1, f3=5|7
		{
			format[j]=int(libwps::readU8(m_input));
			if (!format[j]) continue;
			char const *wh[]= {"color", "pat[id]", "point[id]"};
			f << wh[j] << "=" << format[j] << ",";
		}
		auto flag=int(libwps::readU8(m_input));
		if ((flag&1)==0)
			f << "use[color],";
		else
			format[0]=-1;
		if ((flag&2)==0)
			f << "use[pat],";
		else
			format[1]=-1;
		if ((flag&4)==0)
			f << "use[point],";
		else
			format[2]=-1;
		if (flag&0xF8) f << "#fl=" << std::hex << (flag&0xf8) << std::dec << ",";
		f << "],";
		if (!chart || id<0 || id>=6) continue;
		auto &serieData = chart->m_seriesData[id];
		for (int j=0; j<3; ++j) serieData.m_ids[j]=format[j];
	}

	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	return true;
}

bool WKS4Chart::readChartSeriesColorMap()
{
	libwps::DebugStream f;

	long pos = m_input->tell();
	auto type = long(libwps::read16(m_input));
	if (type != 0x5431)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartSeriesColorMap: not a series styles\n"));
		return false;
	}
	auto sz = long(libwps::readU16(m_input));

	if (sz%8)
	{
		WPS_DEBUG_MSG(("WKS4Chart::readChartSeriesColorMap: chart definition too short\n"));
		ascii().addPos(pos);
		ascii().addNote("Entries(ChartSerColor):###");
		return true;
	}

	auto N=int(sz/8);
	f << "Entries(ChartSerColor):";
	for (int i = 0; i<N; ++i) // USEME. Note, black can correspond to auto color or black color
	{
		auto id=int(libwps::read16(m_input));
		unsigned char col[3];
		for (auto &c : col) c=static_cast<unsigned char>(libwps::read16(m_input)>>8);
		f << "S" << id << "=" << WPSColor(col[0],col[1],col[2]) << ",";
	}
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	return true;
}

////////////////////////////////////////////////////////////
// send data
////////////////////////////////////////////////////////////

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
