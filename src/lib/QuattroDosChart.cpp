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

#include "QuattroDos.h"

#include "QuattroDosChart.h"

namespace QuattroDosChartInternal
{

///////////////////////////////////////////////////////////////////
//! the chart of a Quattro Pro Dos

class Chart final : public WKSChart
{
public:
	//! constructor
	explicit Chart(QuattroDosChart &parser, RVNGInputStreamPtr const &input)
		: WKSChart()
		, m_spreadsheetId(-1)
		, m_parser(parser)
		, m_input(input)
	{
	}
	//! send the zone content (called when the zone is of text type)
	void sendContent(TextZone const &zone, WPSListenerPtr &listener) const final;

	//! the chart range
	Position m_ranges[2];
	//! the sheet id
	int m_spreadsheetId;
protected:
	//! the parser
	QuattroDosChart &m_parser;
	//! the input
	RVNGInputStreamPtr m_input;
};

void Chart::sendContent(TextZone const &zone, WPSListenerPtr &listener) const
{
	if (!listener.get())
	{
		WPS_DEBUG_MSG(("QuattroDosChartInternal::Chart::sendContent: no listener\n"));
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

//! the state of QuattroDosChart
struct State
{
	//! constructor
	State()
		: m_eof(-1)
		, m_version(-1)
		, m_chartType(-1)
		, m_idToChartMap()
	{
	}
	//! returns the pattern corresponding to a pattern id between 0 and 15
	static bool getPattern(int id, WPSGraphicStyle::Pattern &pattern);

	//! the last file position
	long m_eof;
	//! the file version
	int m_version;
	//! the chart type
	int m_chartType;
	//! map sheet id to chart
	std::multimap<int,std::shared_ptr<Chart> > m_idToChartMap;

};

bool State::getPattern(int id, WPSGraphicStyle::Pattern &pat)
{
	if (id<0 || id>15)
	{
		WPS_DEBUG_MSG(("QuattroDosChartInternal::State::getPattern(): unknown pattern id: %d\n",id));
		return false;
	}
	static const uint16_t  patterns[]=
	{
		0xffff, 0xffff, 0xffff, 0xffff, 0x0000, 0x0000, 0x0000, 0x0000, 0x00ff, 0xff00, 0x00ff, 0xff00, 0xeedd, 0xbb77, 0xeedd, 0xbb77, // 0-3
		0x3366, 0xcc99, 0x3366, 0xcc99, 0x8844, 0x2211, 0x8844, 0x2211, 0x99cc, 0x6633, 0x99cc, 0x6633, 0xff10, 0x1010, 0xff10, 0xff10, // 4-7
		0xbf7f, 0xfefc, 0x7bb7, 0xcfdf, 0xaa55, 0xaa55, 0xaa55, 0xaa55, 0x7fff, 0xffff, 0xf7ff, 0xffff, 0x77ff, 0xddff, 0x77ff, 0xddff, // 8-9
		0x990f, 0x050f, 0x99f0, 0x50f0, 0x0101, 0x01ff, 0x1010, 0x10ff, 0xbf7f, 0xfefc, 0x7bb7, 0xcfdf, 0xf77f, 0xbfdf, 0xeffe, 0xfdfb // 12-15
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
QuattroDosChart::QuattroDosChart(QuattroDosParser &parser)
	: m_input(parser.getInput())
	, m_listener()
	, m_mainParser(parser)
	, m_state(new QuattroDosChartInternal::State)
	, m_asciiFile(parser.ascii())
{
}

QuattroDosChart::~QuattroDosChart()
{
}

int QuattroDosChart::version() const
{
	if (m_state->m_version<0)
		m_state->m_version=m_mainParser.version();
	return m_state->m_version;
}

bool QuattroDosChart::checkFilePosition(long pos)
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

int QuattroDosChart::getNumSpreadsheets() const
{
	auto it = m_state->m_idToChartMap.end();
	if (it==m_state->m_idToChartMap.begin())
		return 0;
	--it;
	if (it->first>255)
	{
		WPS_DEBUG_MSG(("QuattroDosChart::getNumSpreadsheets: the number of spreadsheets seems to big: %d\n", it->first));
		return 256;
	}
	return it->first+1;
}

void QuattroDosChart::getChartPositionMap(int sheetId, std::map<Vec2i,Vec2i> &cellList) const
{
	cellList.clear();
	auto it = m_state->m_idToChartMap.lower_bound(sheetId);
	while (it!=m_state->m_idToChartMap.end() && it->first==sheetId)
	{
		if (it->second)
			cellList[it->second->m_ranges[0].m_pos]=it->second->m_ranges[1].m_pos;
		++it;
	}
}

bool QuattroDosChart::sendChart(int sheetId, Vec2i const &cell, Vec2f const &chartSize)
{
	if (m_listener.get() == nullptr)
	{
		WPS_DEBUG_MSG(("QuattroDosChart::sendChart: I can not find the listener\n"));
		return false;
	}
	auto it = m_state->m_idToChartMap.lower_bound(sheetId);
	while (it!=m_state->m_idToChartMap.end() && it->first==sheetId)
	{
		if (it->second && it->second->m_ranges[0].m_pos==cell)
		{
			Vec2f size(chartSize[0]>0 ? chartSize[0] : 100, chartSize[1]>0 ? chartSize[1] : 100);
			WPSPosition pos(Vec2f(0,0), size, librevenge::RVNG_POINT);
			pos.m_anchorTo = WPSPosition::Cell;
			auto endRange=it->second->m_ranges[1];
			endRange.m_pos += Vec2i(1,1);
			pos.m_anchorCellName = endRange.getCellName();
			// the chart is in fact a 1024*7?? windows scaled back on the chosen cells
			it->second->m_dimension=Vec2f(1024,700);
			m_listener->insertChart(pos, *it->second);
			return true;
		}
		++it;
	}
	WPS_DEBUG_MSG(("QuattroDosChart::sendChart: can not find chart %dx%d[%d]\n", cell[0], cell[1], sheetId));
	return false;
}

bool QuattroDosChart::sendText(WPSEntry const &entry)
{
	if (m_listener.get() == nullptr)
	{
		WPS_DEBUG_MSG(("QuattroDosChart::sendText: I can not find the listener\n"));
		return false;
	}
	if (!entry.valid())
		return true;
	m_input->seek(entry.begin(), librevenge::RVNG_SEEK_SET);
	std::string text;
	for (long i=0; i<entry.length(); ++i)
	{
		auto c = char(libwps::readU8(m_input));
		if (c == '\0') continue;
		text.push_back(c);
	}
	if (!text.empty())
		m_listener->insertUnicodeString(libwps_tools_win::Font::unicodeString(text, m_mainParser.getDefaultFontType()));
	return true;
}
////////////////////////////////////////////////////////////
// low level

////////////////////////////////////////////////////////////
// general
////////////////////////////////////////////////////////////
bool QuattroDosChart::readChartSetType()
{

	long pos = m_input->tell();
	auto type = long(libwps::readU16(m_input));
	if (type != 0xb8 && type != 0xca)
	{
		WPS_DEBUG_MSG(("QuattroDosChart::readChartSetType: not a chart definition\n"));
		return false;
	}
	libwps::DebugStream f;
	f << "Entries(" << (type==0xb8 ? "Chart3d" : "ChartBubble") << "):";
	auto sz = long(libwps::readU16(m_input));
	if (sz!=1)
	{
		WPS_DEBUG_MSG(("QuattroDosChart::readChartSetType: find unexpected size\n"));
		f << "###";
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		m_input->seek(sz, librevenge::RVNG_SEEK_CUR);
		return true;
	}
	m_state->m_chartType=int(libwps::readU8(m_input));
	f << "type=" << m_state->m_chartType << ",";
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

bool QuattroDosChart::readChartName()
{

	long pos = m_input->tell();
	auto type = long(libwps::readU16(m_input));
	if (type != 0xb9)
	{
		WPS_DEBUG_MSG(("QuattroDosChart::readChartName: not a chart definition\n"));
		return false;
	}
	libwps::DebugStream f;
	f << "Entries(ChartName):";
	auto sz = long(libwps::readU16(m_input));
	if (sz!=0x10)
	{
		WPS_DEBUG_MSG(("QuattroDosChart::readChartName: find unexpected size\n"));
		f << "###";
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		m_input->seek(sz, librevenge::RVNG_SEEK_CUR);
		return true;
	}
	librevenge::RVNGString name("");
	if (!m_mainParser.readPString(name,15))
		f << "##sSz,";
	else if (!name.empty())
		f << name.cstr() << ',';
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

bool QuattroDosChart::readChart()
{
	int vers=version();
	libwps::DebugStream f;

	long pos = m_input->tell();
	auto type = long(libwps::readU16(m_input));
	if (type != 0x2D && type != 0x2e)
	{
		WPS_DEBUG_MSG(("QuattroDosChart::readChart: not a chart definition\n"));
		return false;
	}
	auto sz = long(libwps::readU16(m_input));
	int normalSz = 0x2ca;
	if (type == 0x2e) normalSz += 0x10;
	if (vers>=2) normalSz+=2*26+4;
	if (sz < normalSz)
	{
		WPS_DEBUG_MSG(("QuattroDosChart::readChart: chart definition too short\n"));
		ascii().addPos(pos);
		ascii().addNote("Entries(ChartDef):###");
		return true;
	}

	f << "Entries(ChartDef):";
	std::shared_ptr<QuattroDosChartInternal::Chart> chart(new QuattroDosChartInternal::Chart(*this, m_input));
	if (type == 0x2e)
	{
		long actPos = m_input->tell();
		librevenge::RVNGString name("");
		if (!m_mainParser.readPString(name,15))
			f << "##sSz,";
		else if (!name.empty())
		{
			chart->m_name=name;
			f << "title=" << name.cstr() << ',';
		}
		m_input->seek(actPos+16, librevenge::RVNG_SEEK_SET);
	}
	bool hasLegend=false;
	for (int i = 0; i < 13; i++) // 0: X-serie, 1-6: data[series%d], 7-12:label serie%d
	{
		WKSChart::Position ranges[2];
		for (auto &range : ranges)
		{
			int dim[3];
			for (int d=0; d<(vers>=2 ? 3 : 2); ++d) dim[d] = int(libwps::read16(m_input));
			if (dim[0]==-1)
				continue;
			range = WKSChart::Position(Vec2i(dim[0],dim[1]), m_mainParser.getSheetName(vers<2 ? 0 : dim[2]));
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
				if (ranges[0]!=ranges[1])
				{
					WPS_DEBUG_MSG(("QuattroDosChart::readChart: unexpected label ranges for %d\n", i));
				}
			}
		}
	}
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	char const *axisNames[]= {"X", "Y", "Y2"};
	pos=m_input->tell();
	f.str("");
	f << "ChartDef-A:"; // GraphDef0
	auto chartType=int(libwps::readU8(m_input));
	if (m_state->m_chartType>=0)
	{
		chartType=m_state->m_chartType;
		m_state->m_chartType=-1;
	}
	WKSChart::Serie::Type serieType=WKSChart::Serie::S_Bar;
	switch (chartType)
	{
	case 0: // XY
		serieType=WKSChart::Serie::S_Scatter;
		break;
	case 1: // basic column
		break;
	case 2: // pie
	case 6: // column (like circle)
		serieType=WKSChart::Serie::S_Circle;
		break;
	case 3: // area stacked
		serieType=WKSChart::Serie::S_Area;
		chart->m_dataStacked=true;
		break;
	case 4: // line
		serieType=WKSChart::Serie::S_Line;
		break;
	case 5: // basic column
		chart->m_dataStacked=true;
		break;
	case 7: // min-max
		serieType=WKSChart::Serie::S_Stock;
		break;
	case 8: // basic bar
		chart->m_dataVertical=true;
		break;
	case 9: // bubble
		serieType=WKSChart::Serie::S_Bubble;
		break;
	case 10: // 3D-bar
	case 12: // 3D-bar tie together
		chart->m_is3D=true;
		break;
	case 11: // 3D-ribbon
		chart->m_is3D=true;
		serieType=WKSChart::Serie::S_Line;
		break;
	case 13:
		chart->m_is3D=true;
		serieType=WKSChart::Serie::S_Area;
		chart->m_dataStacked=true;
		break;
	default:
		WPS_DEBUG_MSG(("QuattroDosChart::readChart: unknown chart type\n"));
		f << "###";
	}
	if (chartType) f << "type=" << chartType << ",";

	chart->m_type=serieType;
	if (chartType==3 || chartType==5) chart->m_dataStacked=true;
	if (chartType==8) chart->m_dataVertical=true;
	auto pointType = serieType==WKSChart::Serie::S_Scatter ? WKSChart::Serie::P_Automatic : WKSChart::Serie::P_None;
	for (int i=0; i<6; ++i)
	{
		auto *serie=chart->getSerie(i, false);
		if (!serie) continue;
		serie->m_type=serieType;
		serie->m_pointType=pointType;
		serie->m_style.m_lineWidth=1;
	}
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
		// ok to find string before i < 6
		// checkme after
		long actPos = m_input->tell();
		int dataSz = i < 4 ? 40 : 20;
		librevenge::RVNGString name("");
		if (!m_mainParser.readPString(name,dataSz-1))
			f << "##sSz,";
		if (name.empty())
		{
			m_input->seek(actPos+dataSz, librevenge::RVNG_SEEK_SET);
			continue;
		}
		if (i<2)
		{
			WPSEntry entry;
			entry.setBegin(actPos+1);
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
			auto *serie=chart->getSerie(i-3, false);
			if (serie)
			{
				serie->m_legendText=name;
				hasLegend=true;
			}
			f << "serie" << i-3 <<  "=" << name.cstr() << ",";
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
	for (int i=0; i<2; ++i)
	{
		val=int(libwps::readU8(m_input));
		if (val==1 || val==255)
			f << "display[scale" << axisNames[i] << "]=no,";
		else if (val)
			f << "display[scale" << axisNames[i] << "]=##" << val << ",";
	}
	f << "pattern[series]=[";
	int patternSeriesId[6];
	for (auto &pId : patternSeriesId)
	{
		pId=int(libwps::readU16(m_input));
		if (pId==0)
			f << "_,";
		else
			f << pId << ",";
	}
	f << "],";
	f << "colors2=[";
	for (int i=0; i<3; ++i)
		f << int(libwps::readU16(m_input)) << ",";
	f << "],";
	f << "color[series]=[";
	WPSColor colorSeriesId[6];
	for (auto &color : colorSeriesId)
	{
		auto cId=int(libwps::readU16(m_input));
		if (m_mainParser.getColor(cId, color))
			f << color << ",";
		else
		{
			color=WPSColor(128,128,128);
			f << "###" << cId << ",";
		}
	}
	f << "],";
	f << "unkn=[";
	for (int i=0; i<4; ++i) // 3 colors + big number
		f << int(libwps::readU16(m_input)) << ",";
	f << "],";
	auto cId=int(libwps::readU8(m_input));
	WPSColor backColor(WPSColor::white());
	if (m_mainParser.getColor(cId, backColor))
	{
		chart->m_style.setSurfaceColor(backColor);
		f << "col[background]=" << backColor << ",";
	}
	else
		f << "col[background]=###" << val << ",";
	val=int(libwps::readU8(m_input));
	if (val) f << "f0=" << val << ",";
	m_input->seek(pos+46, librevenge::RVNG_SEEK_SET);
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	pos=m_input->tell();
	f.str("");
	f << "ChartDef-fonts:";
	for (int i = 0; i < 6; ++i)   // the series font?
	{
		WPSFont font;
		libwps_tools_win::Font::Type fFontType;
		if (!m_mainParser.readFont(font, fFontType))
		{
			f << "###";
			m_input->seek(pos+(i+1)*8, librevenge::RVNG_SEEK_SET);
		}
		else
		{
			auto *serie=chart->getSerie(i, false);
			if (serie)
				serie->m_font=font;
		}
		f << "[" << font << "],";
	}
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	pos=m_input->tell();
	f.str("");
	f << "ChartDef-C:";
	val=int(libwps::readU16(m_input));
	if (val) f << "f0=" << std::hex << val << std::dec << ",";
	for (int i=0; i<6; ++i)
	{
		f << "serie" << i << "=[";
		auto *serie = chart->getSerie(i, false);
		val=int(libwps::readU8(m_input));
		if (val==1 || val==255)
		{
			if (serie) serie->m_useSecondaryY=true;
			f << "secondary,";
		}
		else if (val)
			f << "coordinate=###" << val << ",";
		val=int(libwps::readU8(m_input));
		if (val==1)
		{
			if (serie && serie->m_type==WKSChart::Serie::S_Line)
				serie->m_type=WKSChart::Serie::S_Bar;
			f << "bar,";
		}
		else if (val==2)
		{
			if (serie && serie->m_type==WKSChart::Serie::S_Bar)
				serie->m_type=WKSChart::Serie::S_Line;
			f << "line,";
		}
		else if (val)
			f << "type=##" << val << ",";
		f << "],";
	}
	val=int(libwps::readU8(m_input));
	auto &axisYSecond=chart->getAxis(2);
	if (val==1)
	{
		f << "scaleY2[manual],";
		axisYSecond.m_automaticScaling=false;
	}
	else if (val)
		f << "scaleY2=##" << val << ",";
	for (int i=0; i<2; ++i)
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
				f << (i==0 ? "lowY2" : "highY2") << "=" << value << ",";
			axisYSecond.m_scaling[i]=float(value);
		}
	}
	val=int(libwps::readU8(m_input));
	if (val)
		f << "fmtY2=" << std::hex << val << std::dec << ",";
	val=int(libwps::readU8(m_input));
	if (val)
		f << "f1=" << val << ",";
	m_input->seek(pos+33, librevenge::RVNG_SEEK_SET);
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	pos=m_input->tell();
	f.str("");
	f << "ChartDef-name2:";
	librevenge::RVNGString name("");
	if (!m_mainParser.readPString(name,39))
		f << "##sSz,";
	else if (!name.empty()) f << name.cstr() << ",";
	m_input->seek(pos+40, librevenge::RVNG_SEEK_SET);
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	pos=m_input->tell();
	f.str("");
	f << "ChartDef-D:";
	val=int(libwps::readU8(m_input));
	if (val==0)
		f << "text[pos]=bottom,";
	else if (val==1)
		f << "text[pos]=right,";
	else if (val==2)
		f << "text[pos]=none,";
	else
		f << "text[pos]=##" << val << ",";
	for (int i=0; i<3; ++i) // USEME
	{
		val=int(libwps::readU8(m_input));
		if (val==6) continue;
		f << (i==0 ? "title" : i==1 ? "legend" : "graph") << "[ouline]=";
		if (val>=0 && val<=7)
		{
			char const *wh[]= {"box", "doubleLine", "thickLine", "shadow", "3d", "rndRect", "none","sculpted"};
			f << wh[val] << ",";
		}
		else
			f << "###" << val << ",";
	}
	for (int i=0; i<3; ++i)
	{
		auto &axis = chart->getAxis(i);
		f << "axis" << (i==0 ? "X" : i==1 ? "Y" : "Y2") << "=[";
		val=int(libwps::readU16(m_input));
		if (val) f << "f0=" << std::hex << val << std::dec << ",";
		val=int(libwps::readU8(m_input));
		if (val==0)
			axis.m_type= WKSChart::Axis::A_Numeric;
		else if (val==1 || val==255)
		{
			f << "log,"; // 0:normal
			axis.m_type= WKSChart::Axis::A_Logarithmic;
		}
		else f << "scale=###" << val << ",";
		long actPos=m_input->tell();
		double value;
		bool isNaN;
		if (!libwps::readDouble8(m_input, value, isNaN))
		{
			m_input->seek(actPos+8, librevenge::RVNG_SEEK_SET);
			f << "##value,";
		}
		else if (value<0 || value>0) // USEME
			f << "increment=" << value << ",";
		f << "],";
	}
	val=int(libwps::readU8(m_input));
	if (val && val<=7)
	{
		// USEME
		char const *wh[]= {"solid", "dotted", "center-line", "dashed", "heavy solid",
		                   "heavy dotted", "heavy centered", "heavy dashed"
		                  };
		f << "style[gridline]=" << wh[val] << ",";
	}
	else if (val)
		f << "##style[gridline]=" << val << ",";
	cId=int(libwps::readU8(m_input));
	WPSColor gridColor(WPSColor::black());
	if (m_mainParser.getColor(cId, gridColor))
	{
		chart->setGridColor(gridColor);
		f << "color[line/grid]=" << gridColor << ",";
	}
	else
		f << "##color[line/grid]=" << cId << ",";
	ascii().addDelimiter(m_input->tell(),'|');
	m_input->seek(pos+49, librevenge::RVNG_SEEK_SET);
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	// time to update the serie pattern
	if (hasLegend)
	{
		chart->getLegend().m_show=true;
		chart->getLegend().m_autoPosition=true;
		chart->getLegend().m_relativePosition=WPSBorder::RightBit;
	}
	for (int i=0; i<6; ++i)
	{
		auto *serie=chart->getSerie(i, false);
		if (!serie) continue;
		WPSGraphicStyle::Pattern pattern;
		if (!m_state->getPattern(patternSeriesId[i],pattern))
		{
			WPS_DEBUG_MSG(("QuattroDosChart::readChart: oops, can not find pattern %d\n", i));
			serie->setPrimaryColor(colorSeriesId[i]);
			serie->setSecondaryColor(gridColor);
			continue;
		}
		pattern.m_colors[0]=colorSeriesId[i];
		pattern.m_colors[1]=backColor;
		serie->setPrimaryPattern(pattern);
		serie->setSecondaryColor(gridColor);
	}

	pos=m_input->tell();
	f.str("");
	f << "ChartDef-E:";
	val=int(libwps::readU8(m_input));
	if (val) f << "tickX[alternate],";
	val=int(libwps::readU16(m_input));
	if (val) f << "f0=" << std::hex << val << std::dec << ",";
	val=int(libwps::readU8(m_input));
	if (val) f << "use[depth],";
	ascii().addDelimiter(m_input->tell(),'|');
	m_input->seek(24, librevenge::RVNG_SEEK_CUR);
	ascii().addDelimiter(m_input->tell(),'|');
	val=int(libwps::read16(m_input));
	if (val) f << "bar[width]=" << val << "%,";
	for (int i=0; i<2; ++i)   // 0: local graph position, 1: title/sub title frames
	{
		int dim[4];
		for (auto &d : dim) d = int(libwps::read16(m_input));
		if (dim[0]==-1)
			continue;
		WPSBox2i box(Vec2i(dim[0],dim[1]), Vec2i(dim[2],dim[3]));
		if (box==WPSBox2i()) continue;
		f << (i==0 ? "grid" : "title") << "[pos]=" << box << ",";
		if (i==0) // the size seems to be the windows size 1024*???
			chart->m_plotAreaPosition=
			    WPSBox2f(Vec2f(float(dim[0])/1024.f,float(dim[1])/700.f),Vec2f(float(dim[2])/1024.f,float(dim[3])/700.f));
	}
	f << "unk=[";
	for (int i=0; i<9; ++i)
	{
		val=int(libwps::readU8(m_input));
		if (val==255)
			f << "*,";
		else if (val)
			f << val << ",";
		else
			f << "_,";
	}
	f << "],";
	WPSColor color;
	cId=int(libwps::readU8(m_input));
	if (m_mainParser.getColor(cId, color))
	{
		chart->m_wallStyle.setSurfaceColor(color);
		f << "col[fill]=" << color << ",";
	}
	else
		f << "col[fill]=###" << color << ",";
	val=int(libwps::readU8(m_input));
	if (val) f << "f1=" << std::hex << val << std::dec << ",";
	for (auto &range : chart->m_ranges)
	{
		int dim[3];
		for (int d=0; d<(vers>=2 ? 3 : 2); ++d) dim[d] = int(libwps::read16(m_input));
		if (dim[0]==-1)
			continue;
		chart->m_spreadsheetId=vers<2 ? 0 : dim[2];
		range = WKSChart::Position(Vec2i(dim[0],dim[1]), m_mainParser.getSheetName(chart->m_spreadsheetId));
	}
	if (chart->m_ranges[0].valid(chart->m_ranges[1]))
	{
		f << "position=" << chart->m_ranges[0] << ":" << chart->m_ranges[1] << ",";
		if (chart->m_ranges[1].m_pos[0]>255 || chart->m_ranges[1].m_pos[1]>65535)
		{
			f << "###";
			WPS_DEBUG_MSG(("QuattroDosChart::readChart: oops, the maximum position seems bad\n"));
		}
		else if (chart->m_ranges[0].m_sheetName!=chart->m_ranges[1].m_sheetName)
		{
			f << "###";
			WPS_DEBUG_MSG(("QuattroDosChart::readChart: oops, the position is on different sheet\n"));
		}
		else
		{
			m_state->m_idToChartMap.insert(std::multimap<int,std::shared_ptr<QuattroDosChartInternal::Chart> >::value_type(chart->m_spreadsheetId, chart));
		}
	}
	if (sz != normalSz) ascii().addDelimiter(m_input->tell(),'#');
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	return true;
}

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
