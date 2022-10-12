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

#include <algorithm>
#include <cctype>
#include <cmath>
#include <sstream>
#include <limits>
#include <stack>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WPSCell.h"
#include "WKSContentListener.h"
#include "WPSEntry.h"
#include "WPSFont.h"
#include "WPSStream.h"
#include "WPSTable.h"

#include "Quattro.h"
#include "QuattroFormula.h"

#include "QuattroSpreadsheet.h"

namespace QuattroSpreadsheetInternal
{
//! a class used to store a style of a cell in QuattroSpreadsheet
struct Style final : public WPSCellFormat
{
	//! construtor
	explicit Style(libwps_tools_win::Font::Type type)
		: WPSCellFormat()
		, m_fontType(type)
		, m_fileFormat(0xFF)
		, m_alignAcrossColumn(false)
		, m_extra("")
	{
	}
	Style(Style const &)=default;
	//! destructor
	~Style() final;
	//! operator<<
	friend std::ostream &operator<<(std::ostream &o, Style const &style);
	//! operator==
	bool operator==(Style const &st) const;
	//! operator!=
	bool operator!=(Style const &st) const
	{
		return !(*this==st);
	}
	//! font encoding type
	libwps_tools_win::Font::Type m_fontType;
	//! the file format
	int m_fileFormat;
	//! flag to know if we must align across column
	bool m_alignAcrossColumn;
	/** extra data */
	std::string m_extra;
};

Style::~Style()
{
}

//! operator<<
std::ostream &operator<<(std::ostream &o, Style const &style)
{
	o << static_cast<WPSCellFormat const &>(style) << ",";
	if (style.m_fileFormat!=0xFF)
		o << "format=" << std::hex << style.m_fileFormat << std::dec << ",";
	if (style.m_extra.length())
		o << "extra=[" << style.m_extra << "],";

	return o;
}

bool Style::operator==(Style const &st) const
{
	if (m_fontType!=st.m_fontType || m_fileFormat!=st.m_fileFormat) return false;
	int diff = WPSCellFormat::compare(st);
	if (diff) return false;
	return m_fileFormat==st.m_fileFormat && m_alignAcrossColumn==st.m_alignAcrossColumn && m_extra==st.m_extra;
}

//! a cellule of a Quattro spreadsheet
class Cell final : public WPSCell
{
public:
	/// constructor
	explicit Cell(libwps_tools_win::Font::Type type)
		: m_fontType(type)
		, m_fileFormat(0xFF)
		, m_styleId(-1)
		, m_alignAcrossColumn(false)
		, m_content()
		, m_hasGraphic(false)
		, m_stream() { }

	//! operator<<
	friend std::ostream &operator<<(std::ostream &o, Cell const &cell);

	//! call when a cell must be send
	bool send(WPSListenerPtr &/*listener*/) final;

	//! call when the content of a cell must be send
	bool sendContent(WPSListenerPtr &/*listener*/) final
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheetInternal::Cell::sendContent: must not be called\n"));
		return false;
	}
	//! update the cell format using file format
	void updateFormat()
	{
		if (m_fileFormat==0xFF)
			return;
		switch ((m_fileFormat>>4)&7)
		{
		case 0:
		case 6: // checkme: date format set by system
			switch (m_fileFormat&0xF)
			{
			case 1: // +/- : kind of bool
				setFormat(F_BOOLEAN);
				break;
			case 2: // default
				break;
			case 3:
				setFormat(F_TEXT);
				break;
			case 4:
				setFormat(F_TEXT);
				m_font.m_attributes |= WPS_HIDDEN_BIT;
				break;
			case 5:
				setDTFormat(F_DATE, "%d %b %y");
				break;
			case 6:
				setDTFormat(F_DATE, "%d %b");
				break;
			case 7:
				setDTFormat(F_DATE, "%b-%d");
				break;
			case 8:
				setDTFormat(F_DATE, "%m/%d/%y");
				break;
			case 9:
				setDTFormat(F_DATE, "%m/%d");
				break;
			case 0xa:
				setDTFormat(F_TIME, "%I:%M:%S%p");
				break;
			case 0xb:
				setDTFormat(F_TIME, "%I:%M%p");
				break;
			case 0xc:
				setDTFormat(F_TIME, "%H:%M:%S");
				break;
			case 0xd:
				setDTFormat(F_TIME, "%H:%M");
				break;
			case 0xe:
				setDTFormat(F_TIME, "%y");
				break;
			case 0xf:
				setDTFormat(F_TIME, "%b");
				break;
			default:
				WPS_DEBUG_MSG(("QuattroSpreadsheetInternal::Cell::updateFormat: unknown format %x\n", unsigned(m_fileFormat)));
				break;
			}
			break;
		case 1: // fixed
			setFormat(F_NUMBER, 1);
			setDigits(m_fileFormat&0xF);
			break;
		case 2: // scientific
			setFormat(F_NUMBER, 2);
			setDigits(m_fileFormat&0xF);
			break;
		case 3: // currency
			setFormat(F_NUMBER, 4);
			setDigits(m_fileFormat&0xF);
			break;
		case 4: // percent
			setFormat(F_NUMBER, 3);
			setDigits(m_fileFormat&0xF);
			break;
		case 5: // decimal
			setFormat(F_NUMBER, 1);
			setDigits(m_fileFormat&0xF);
			break;
		case 7:   // fixme use UserFormat (m_fileFormat&0xF) in Quattro.cpp at least to decode date...
		{
			static bool first=true;
			if (first)
			{
				first=false;
				WPS_DEBUG_MSG(("QuattroSpreadsheetInternal::Cell::updateFormat: user defined format is not supported\n"));
			}
			break;
		}
		default:
			break;
		}
	}
	//! font encoding type
	libwps_tools_win::Font::Type m_fontType;
	//! the file format
	int m_fileFormat;
	//! the style id
	int m_styleId;
	//! flag to know if we must align across column
	bool m_alignAcrossColumn;
	//! the content
	WKSContentListener::CellContent m_content;
	//! a flag to know a cell has some graphic
	bool m_hasGraphic;
	//! the text stream(used to send text's zone)
	std::shared_ptr<WPSStream> m_stream;
};
bool Cell::send(WPSListenerPtr &/*listener*/)
{
	WPS_DEBUG_MSG(("QuattroSpreadsheetInternal::Cell::send: must not be called\n"));
	return false;
}

//! operator<<
std::ostream &operator<<(std::ostream &o, Cell const &cell)
{
	o << reinterpret_cast<WPSCell const &>(cell) << cell.m_content << ",";
	if (cell.m_fileFormat!=0xFF)
		o << "format=" << std::hex << cell.m_fileFormat << std::dec << ",";
	return o;
}

//! the spreadsheet of a Quattro Spreadsheet
class Spreadsheet
{
public:
	//! a constructor
	Spreadsheet(int id, libwps_tools_win::Font::Type fontType)
		: m_id(id)
		, m_numCols(0)
		, m_rowHeightMap()
		, m_heightDefault(13) // fixme: use zone d2
		, m_widthCols()
		, m_widthDefault(54) // fixme: use zone d4
		, m_positionToCellMap()
		, m_dummyCell(fontType)
	{
	}
	//! return a cell corresponding to a spreadsheet, create one if needed
	Cell &getCell(Vec2i const &pos, libwps_tools_win::Font::Type type)
	{
		if (m_positionToCellMap.find(pos)==m_positionToCellMap.end())
		{
			Cell cell(type);
			cell.setPosition(pos);
			if (pos[0]<0 || pos[0]>255)
			{
				WPS_DEBUG_MSG(("QuattroSpreadsheetInternal::Spreadsheet::getCell: find unexpected col=%d\n", pos[0]));
				return m_dummyCell;
			}
			m_positionToCellMap.insert(std::map<Vec2i, Cell>::value_type(pos,cell));
		}
		return m_positionToCellMap.find(pos)->second;
	}
	//! returns true if the spreedsheet is empty
	bool empty() const
	{
		return m_positionToCellMap.empty();
	}
	//! set the columns size
	void setColumnWidth(int col, int w=-1)
	{
		if (col < 0) return;
		if (col >= int(m_widthCols.size())) m_widthCols.resize(size_t(col)+1, -1);
		m_widthCols[size_t(col)] = w;
		if (col >= m_numCols) m_numCols=col+1;
	}

	//! return the columns format
	std::vector<WPSColumnFormat> getWidths() const
	{
		std::vector<WPSColumnFormat> widths;
		WPSColumnFormat defWidth(m_widthDefault), actWidth;
		defWidth.m_useOptimalWidth=true;
		int repeat=0;
		for (auto const &w : m_widthCols)
		{
			WPSColumnFormat newWidth;
			if (w < 0)
				newWidth=defWidth;
			else
				newWidth=WPSColumnFormat(float(w)/20.f);
			if (repeat && newWidth!=actWidth)
			{
				actWidth.m_numRepeat=repeat;
				widths.push_back(actWidth);
				repeat=0;
			}
			if (repeat==0)
				actWidth=newWidth;
			++repeat;
		}
		if (repeat)
		{
			actWidth.m_numRepeat=repeat;
			widths.push_back(actWidth);
		}
		return widths;
	}
	//! set the rows size in TWIP
	void setRowHeight(int row, int h)
	{
		auto rIt=m_rowHeightMap.lower_bound(Vec2i(-1,row));
		if (rIt!=m_rowHeightMap.end() && rIt->first[0]<=row && rIt->first[1]>=row)
		{
			WPS_DEBUG_MSG(("QuattroSpreadsheetInternal::Spreadsheet::setRowHeight: oops, row %d is already set\n", row));
			return;
		}
		if (h>=0)
			m_rowHeightMap[Vec2i(row,row)]=h;
	}
	//! set the rows size in TWIP
	void setRowHeights(int minRow, int maxRow, int h)
	{
		auto rIt=m_rowHeightMap.lower_bound(Vec2i(-1,minRow));
		while (rIt!=m_rowHeightMap.end())
		{
			auto const &cells=rIt->first;
			if (cells[0]>maxRow) break;
			if (cells[1]>=minRow)
			{
				WPS_DEBUG_MSG(("QuattroSpreadsheetInternal::Spreadsheet::setRowHeight: oops, some rows are already set in %dx%d\n", minRow, maxRow));
				return;
			}
			++rIt;
		}
		if (h>=0)
			m_rowHeightMap[Vec2i(minRow,maxRow)]=h;
	}
	//! returns the row size in point
	float getRowHeight(int row) const
	{
		auto rIt=m_rowHeightMap.lower_bound(Vec2i(-1,row));
		if (rIt!=m_rowHeightMap.end() && rIt->first[0]<=row && rIt->first[1]>=row)
			return float(rIt->second)/20.f;
		return m_heightDefault;
	}
	//! returns the height of a row in point and updated repeated row
	float getRowHeight(int row, int &numRepeated) const
	{
		auto rIt=m_rowHeightMap.lower_bound(Vec2i(-1,row));
		if (rIt!=m_rowHeightMap.end() && rIt->first[0]<=row && rIt->first[1]>=row)
		{
			numRepeated=rIt->first[1]-row+1;
			return float(rIt->second)/20.f;
		}
		numRepeated=10000;
		return m_heightDefault;
	}
	//! try to compress the list of row height
	void compressRowHeights()
	{
		auto oldMap=m_rowHeightMap;
		m_rowHeightMap.clear();
		int actHeight=-1;
		Vec2i actPos(0,-1);
		for (auto rIt : oldMap)
		{
			// first check for not filled row
			if (rIt.first[0]!=actPos[1]+1)
			{
				if (actHeight==int(m_heightDefault)*20)
					actPos[1]=rIt.first[0]-1;
				else
				{
					if (actPos[1]>=actPos[0])
						m_rowHeightMap[actPos]=actHeight;
					actHeight=int(m_heightDefault)*20;
					actPos=Vec2i(actPos[1]+1, rIt.first[0]-1);
				}
			}
			if (rIt.second!=actHeight)
			{
				if (actPos[1]>=actPos[0])
					m_rowHeightMap[actPos]=actHeight;
				actPos[0]=rIt.first[0];
				actHeight=rIt.second;
			}
			actPos[1]=rIt.first[1];
		}
		if (actPos[1]>=actPos[0])
			m_rowHeightMap[actPos]=actHeight;
	}
	//! returns the cell position
	Vec2f getPosition(Vec2i const &cell) const
	{
		float c=0;
		int numWidth=int(m_widthCols.size());
		for (int i=0; i<cell[0]; ++i)
		{
			if (i>=numWidth)
			{
				c+=float(i+1-numWidth)*m_widthDefault;
				break;
			}
			int w=m_widthCols[size_t(i)];
			if (w < 0)
				c+=m_widthDefault;
			else
				c+=float(w)/20.f;
		}
		int r=0, prevR=0;
		for (auto it : m_rowHeightMap)
		{
			int maxR=std::min(it.first[1],cell[1]-1);
			if (prevR<it.first[0])
			{
				r+=(maxR-prevR)*int(m_heightDefault)*20;
				prevR=maxR;
			}
			if (maxR<it.first[0])
				break;
			r+=(maxR+1-it.first[0])*it.second;
			prevR=maxR;
		}
		if (prevR<cell[1]) r+=(cell[1]-prevR)*int(m_heightDefault)*20;
		return Vec2f(c,float(r/20));
	}
	//! the spreadsheet id
	int m_id;
	/** the number of columns */
	int m_numCols;

	/** the map Vec2i(min row, max row) to size in TWIP */
	std::map<Vec2i,int> m_rowHeightMap;
	/** the default row size in point */
	float m_heightDefault;
	/** the column size in TWIP */
	std::vector<int> m_widthCols;
	/** the default width size in point */
	float m_widthDefault;
	/** a map cell to not empty cells */
	std::map<Vec2i, Cell> m_positionToCellMap;
	/** a dummy cell */
	mutable Cell m_dummyCell;
};

//! the state of QuattroSpreadsheet
struct State
{
	//! constructor
	explicit State(QuattroFormulaManager::CellReferenceFunction const &readCellReference)
		: m_version(-1)
		, m_maxDimension(0,0,0)
		, m_actSheet(-1)
		, m_stylesList()
		, m_formulaManager(readCellReference, 1)
		, m_spreadsheetMap()
		, m_idToSheetNameMap()
		, m_idToUserFormatMap()
	{
	}
	//! returns the ith real spreadsheet
	std::shared_ptr<Spreadsheet> getSheet(int id, libwps_tools_win::Font::Type fontType)
	{
		auto it=m_spreadsheetMap.find(id);
		if (it!=m_spreadsheetMap.end())
			return it->second;
		std::shared_ptr<Spreadsheet> sheet(new Spreadsheet(id, fontType));
		sheet->setColumnWidth(m_maxDimension[0]);
		if (id<0 || id>m_maxDimension[2])
		{
			WPS_DEBUG_MSG(("QuattroSpreadsheetInternal::State::getSheet: find unexpected id=%d\n", id));
			if (id<0 || id>255) // too small or too big, return dummy spreadsheet
				return sheet;
		}
		m_spreadsheetMap[id]=sheet;
		return sheet;
	}
	//! returns the ith spreadsheet
	librevenge::RVNGString getSheetName(int id) const
	{
		auto it = m_idToSheetNameMap.find(id);
		if (it!=m_idToSheetNameMap.end() && !it->second.empty())
			return it->second;
		librevenge::RVNGString name;
		name.sprintf("Sheet%d", id+1);
		return name;
	}
	//! the file version
	int m_version;
	//! the maximum col, row, sheet
	WPSVec3i m_maxDimension;
	//! the actual sheet
	int m_actSheet;
	//! the list of styles
	std::vector<Style> m_stylesList;
	//! the formula manager
	QuattroFormulaManager m_formulaManager;

	//! the map of spreadsheet
	std::map<int, std::shared_ptr<Spreadsheet> > m_spreadsheetMap;
	//! the map id to sheet's name
	std::map<int, librevenge::RVNGString> m_idToSheetNameMap;
	//! map id to user format string
	std::map<int, librevenge::RVNGString> m_idToUserFormatMap;
};

}

// constructor, destructor
QuattroSpreadsheet::QuattroSpreadsheet(QuattroParser &parser)
	: m_listener()
	, m_mainParser(parser)
	, m_state()
{
	m_state.reset(new QuattroSpreadsheetInternal::State(getReadCellReferenceFunction()));
}

QuattroSpreadsheet::~QuattroSpreadsheet()
{
}

void QuattroSpreadsheet::cleanState()
{
	m_state.reset(new QuattroSpreadsheetInternal::State(getReadCellReferenceFunction()));
}

void QuattroSpreadsheet::updateState()
{
}

int QuattroSpreadsheet::version() const
{
	if (m_state->m_version<0)
		m_state->m_version=m_mainParser.version();
	return m_state->m_version;
}

QuattroFormulaManager::CellReferenceFunction QuattroSpreadsheet::getReadCellReferenceFunction()
{
	return [this](std::shared_ptr<WPSStream> const &stream, long endPos,
	              QuattroFormulaInternal::CellReference &ref,
	              Vec2i const &pos, int sheetId)
	{
		return this->readCellReference(stream, endPos, ref, pos, sheetId);
	};
}

int QuattroSpreadsheet::getNumSpreadsheets() const
{
	if (m_state->m_spreadsheetMap.empty())
		return m_state->m_maxDimension[2]+1;
	auto it=m_state->m_spreadsheetMap.end();
	--it;
	return std::max(it->first,m_state->m_maxDimension[2])+1;
}

librevenge::RVNGString QuattroSpreadsheet::getSheetName(int id) const
{
	return m_state->getSheetName(id);
}

Vec2f QuattroSpreadsheet::getPosition(int sheetId, Vec2i const &cell) const
{
	auto it=m_state->m_spreadsheetMap.find(sheetId);
	if (it==m_state->m_spreadsheetMap.end() || !it->second)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::getPosition: can not find the sheet %d\n", sheetId));
		return Vec2f(float(cell[0]*50), float(cell[1]*13));
	}
	return it->second->getPosition(cell);
}

void QuattroSpreadsheet::addDLLIdName(int id, librevenge::RVNGString const &name, bool func1)
{
	m_state->m_formulaManager.addDLLIdName(id, name, func1);
}

void QuattroSpreadsheet::addUserFormat(int id, librevenge::RVNGString const &name)
{
	if (name.empty())
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::addUserFormat: called with empty name for id=%d\n", id));
		return;
	}
	if (m_state->m_idToUserFormatMap.find(id)!=m_state->m_idToUserFormatMap.end())
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::addUserFormat: called with dupplicated id=%d\n", id));
	}
	else
		m_state->m_idToUserFormatMap[id]=name;
}

////////////////////////////////////////////////////////////
// low level

////////////////////////////////////////////////////////////
//   parse sheet data
////////////////////////////////////////////////////////////
bool QuattroSpreadsheet::readCell(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if ((type < 0xc || type > 0x10) && (type!=0x33))
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readCell: not a cell property\n"));
		return false;
	}
	long sz = libwps::readU16(input);
	long endPos = pos+4+sz;

	if (sz < 5)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readCell: cell def is too short\n"));
		return false;
	}
	int cellPos[2];
	cellPos[0]=int(libwps::readU8(input));
	auto sheetId=int(libwps::readU8(input));
	cellPos[1]=int(libwps::read16(input));
	if (cellPos[1] < 0)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readCell: cell pos is bad\n"));
		return false;
	}
	if (sheetId)
		f << "sheet[id]=" << sheetId << ",";

	auto defFontType=m_mainParser.getDefaultFontType();
	auto sheet = m_state->getSheet(sheetId, defFontType);
	auto &cell=sheet->getCell(Vec2i(cellPos[0],cellPos[1]), defFontType);
	auto format=int(libwps::readU16(input));
	int id=format>>3;
	// format&7: reserved
	if (id<0 || id>int(m_state->m_stylesList.size()))
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readCell: can not find cell format\n"));
		f << "###Ce" << id << ",";
	}
	else if (id)
	{
		auto const &style=m_state->m_stylesList[size_t(id-1)];
		if (type!=0x33)
		{
			cell.m_styleId=id-1;
			cell.m_fileFormat=style.m_fileFormat;
			cell.m_fontType=style.m_fontType;
			static_cast<WPSCellFormat &>(cell)=style;
			cell.m_alignAcrossColumn=style.m_alignAcrossColumn;
		}
		f << "Ce" << id-1 << ",";
	}

	long dataPos = input->tell();
	auto dataSz = int(endPos-dataPos);

	bool ok = true;
	switch (type)
	{
	case 12:
	{
		if (dataSz == 0)
		{
			cell.m_content.m_contentType=WKSContentListener::CellContent::C_NONE;
			break;
		}
		ok = false;
		break;
	}
	case 13:
	{
		if (dataSz == 2)
		{
			cell.m_content.m_contentType=WKSContentListener::CellContent::C_NUMBER;
			cell.m_content.setValue(libwps::read16(input));
			break;
		}
		ok = false;
		break;
	}
	case 14:
	{
		double val;
		bool isNaN;
		if (dataSz == 8 && libwps::readDouble8(input, val, isNaN))
		{
			cell.m_content.m_contentType=WKSContentListener::CellContent::C_NUMBER;
			cell.m_content.setValue(val);
			break;
		}
		ok = false;
		break;
	}
	case 15:
		cell.m_content.m_contentType=WKSContentListener::CellContent::C_TEXT;
		WPS_FALLTHROUGH;
	case 0x33: // formula res
	{
		long begText=input->tell()+1;
		std::string s("");
		// align + c string
		auto align=char(libwps::readU8(input));
		if (align=='\'') cell.setHAlignment(WPSCellFormat::HALIGN_DEFAULT);
		else if (align=='^') cell.setHAlignment(WPSCellFormat::HALIGN_CENTER);
		else if (align=='\"') cell.setHAlignment(WPSCellFormat::HALIGN_RIGHT);
		else if (align=='\\') f << "repeat,"; // USEME
		else if (align==0x7c) f << "break,"; // FIXME remove "::" in text
		else if (align) f << "#align=" << int(align) << ",";

		librevenge::RVNGString text("");
		if (!m_mainParser.readCString(stream,text,dataSz-1))
			f << "##sSz,";
		else
		{
			if (endPos!=input->tell() && endPos!=input->tell()+1)
			{
				f << "#extra,";
				ascFile.addDelimiter(input->tell(), '|');
			}
			cell.m_stream=stream;
			cell.m_content.m_textEntry.setBegin(begText);
			cell.m_content.m_textEntry.setEnd(input->tell()-1);
			if (!text.empty())
				f << text.cstr() << ",";
		}
		break;
	}
	case 16:
	{
		double val;
		bool isNaN;
		if (dataSz >= 10 && libwps::readDouble8(input, val, isNaN))
		{
			cell.m_content.m_contentType=WKSContentListener::CellContent::C_FORMULA;
			cell.m_content.setValue(val);
			auto state=int(libwps::readU16(input));
			if (state)
			{
				f << "state[";
				if (state&0x8) f << "constant,";
				if (state==0x10) f << "volatile,";
				if (state==0x100) f << "inArray,";
				if (state==0x200) f << "useDLL,";
				if (state&0xfcf3) f << "#state=" << std::hex << (state&0xfcf3) << std::dec << ",";
				f << "],";
			}
			std::string error;
			if (!m_state->m_formulaManager.readFormula(stream, endPos, cell.position(), sheetId, cell.m_content.m_formula, error))
			{
				cell.m_content.m_contentType=WKSContentListener::CellContent::C_NUMBER;
				ascFile.addDelimiter(input->tell()-1, '#');
			}
			if (error.length()) f << error;
			break;
		}
		ok = false;
		break;
	}
	default:
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readCell: unknown type=%ld\n", type));
		ok = false;
		break;
	}
	if (!ok) ascFile.addDelimiter(dataPos, '#');

	input->seek(pos+sz, librevenge::RVNG_SEEK_SET);

	std::string extra=f.str();
	f.str("");
	f << cell << "," << extra;

	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	return true;
}

bool QuattroSpreadsheet::readCellStyle(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if (type != 0xce)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readCellStyle: not a style zone\n"));
		return false;
	}
	long sz = libwps::readU16(input);
	f << "[Ce" << m_state->m_stylesList.size() << "],";
	QuattroSpreadsheetInternal::Style style(m_mainParser.getDefaultFontType());
	if (sz<8)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readCellStyle: size seems bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		m_state->m_stylesList.push_back(style);
		return true;
	}
	style.m_fileFormat=int(libwps::readU8(input));
	if (style.m_fileFormat!=0xFF)
		f << "form=" << std::hex << style.m_fileFormat << std::dec << ",";
	auto flag=int(libwps::readU8(input));
	switch (flag&7)
	{
	case 1:
		style.setHAlignment(WPSCellFormat::HALIGN_LEFT);
		f << "left,";
		break;
	case 2:
		style.setHAlignment(WPSCellFormat::HALIGN_CENTER);
		f << "center,";
		break;
	case 3:
		style.setHAlignment(WPSCellFormat::HALIGN_RIGHT);
		f << "right,";
		break;
	case 4:
		style.setHAlignment(WPSCellFormat::HALIGN_FULL);
		f << "block,";
		break;
	case 6:
	{
		style.setHAlignment(WPSCellFormat::HALIGN_CENTER);
		style.m_alignAcrossColumn=true;
		f << "center[across],";
		break;
	}
	default:
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readCellStyle: find unexpected alignment\n"));
		f << "###align=" << (flag&7) << ",";
		break;
	case 0: // standart
		break;
	}
	if (sz>=12)
	{
		switch ((flag>>3)&3)
		{
		case 0: // default
			style.setVAlignment(WPSCellFormat::VALIGN_BOTTOM);
			break;
		case 1:
			style.setVAlignment(WPSCellFormat::VALIGN_CENTER);
			f << "vAlign=center,";
			break;
		case 2:
			style.setVAlignment(WPSCellFormat::VALIGN_TOP);
			f << "vAlign=top,";
			break;
		default:
			WPS_DEBUG_MSG(("QuattroSpreadsheet::readCellStyle: find unexpected alignment\n"));
			f << "###valign=3,";
			break;
		}
		if (flag&0x20)
		{
			style.setTextRotation(270);
			f << "top[down],";
		}
		if (flag&0x80)
		{
			style.setWrapping(WPSCellFormat::WRAP_WRAP);
			f << "wrap,";
		}
		flag &= 0x40;
	}
	else
	{
		switch ((flag>>6)&3) // checkme, maybe diferent in wb2, ie. I find difference with qprostyle.cpp
		{
		default:
		case 0: // standart
			break;
		case 1:
			f << "label[only],";
			break;
		case 2:
			f << "date[only],";
			break;
		case 3:
			f << "##input=3,";
			break;
		}
		flag&=0x38;
	}
	if (flag)
		f << "#fl=" << std::hex << flag << std::dec << ",";
	auto val=int(libwps::readU8(input));
	int color[3]= {val>>4, val&0xf, 0}; // col2: shade, col1: shade, textcolor
	val=int(libwps::readU8(input));
	color[2]=val>>4;
	int blend=(val&0x7);
	WPSColor colors[]= {WPSColor::white(), WPSColor::black(), WPSColor::black()};
	for (int i=0; i<3; ++i)
	{
		int const expected[]= {0,3,3};
		if (color[i]==expected[i]) continue;
		if (m_mainParser.getColor(color[i], colors[i]))
			f << "color" << i << "=" << colors[i] << ",";
		else
			f << "##color" << i << "=" << color[i] << ",";
	}
	if (blend==7)
		f << "###blend=7,";
	else
	{
		int const percent[]= {0,6,3,1,2,5,4};
		float fPercent=float(percent[blend])/6.f;
		if (blend)
			f << "blend=" << 100.f *fPercent << "%,";
		style.setBackgroundColor(WPSColor::barycenter(fPercent,colors[1],1.f-fPercent,colors[0]));
		// percent col2 + (1-percent) col1
	}
	if (val&8) f << "fl[8],";
	auto fId=int(libwps::readU8(input));
	WPSFont font;
	if (fId)
	{
		if (!m_mainParser.getFont(fId-1, font, style.m_fontType))
			f << "###";
		f << "F" << fId-1 << ",";
	}
	font.m_color=colors[2];
	style.setFont(font);
	auto bFlags=int(libwps::readU8(input));
	// 80:has[textColor], 40:has[protect], 20:has[borders], 10:has[fonts], 8:has[color], 4:[hasAlign], 2:[hasForm], 1:[hasProtect+0x40?]...
	val=int(libwps::readU8(input));
	val &= 0x41;
	if (val==0x41)
		f << "protect=no,";
	else if (val) f << "fl2=" << std::hex << val << std::dec << ",";
	val=int(libwps::readU8(input)); // USEME
	if (val) f << "style[id]=" << val << ",";
	WPSColor borderColors[4];
	for (auto &c : borderColors) c=WPSColor::black();
	if (sz>=12)   // sz=12, pre-wb2, wb2?
	{
		f << "borders[color]=[";
		for (int i=0; i<2; ++i)
		{
			val=int(libwps::readU8(input));
			for (int j=0; j<2; ++j)
			{
				int c=j==1 ? (val>>4) : (val&0xf);
				if (!m_mainParser.getColor(c, borderColors[2*i+j]))
					f << "##color=" << c << ",";
				else if (borderColors[2*i+j].isBlack())
					f << "_,";
				else
					f << borderColors[2*i+j] << ",";
			}
		}
		f << "],";
		val=int(libwps::readU16(input));
		switch (val&3)
		{
		default:
		case 0: // standart
			break;
		case 1:
			f << "label[only],";
			break;
		case 2:
			f << "date[only],";
			break;
		case 3:
			f << "##input=3,";
			break;
		}
		if (val&4) f << "use[lineColor],"; // or page color
		val&=0xfff8;
		if (val) f << "fl3=" << val << ",";
	}
	if (bFlags)
	{
		f << "borders=[";
		for (int i=0, depl=0; i<4; ++i, depl+=2)   // BRTL
		{
			int bType=(bFlags>>depl)&3;
			if (!bType) continue;
			char const *wh[]= {"L","T","R","B"};
			WPSBorder border;
			switch (bType)
			{
			case 1: // normal
				f << wh[i] << ",";
				break;
			case 2: // double
				border.m_type=WPSBorder::Double;
				f << wh[i] << "=double,";
				break;
			case 3: // width*2
				border.m_width=2;
				f << wh[i] << "=w2,";
				break;
			default: // impossible
				break;
			}
			border.m_color=borderColors[i];
			int const which[]= {WPSBorder::LeftBit, WPSBorder::TopBit, WPSBorder::RightBit, WPSBorder::BottomBit};
			style.setBorders(which[i], border);
		}
		f << "],";
	}
	m_state->m_stylesList.push_back(style);
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroSpreadsheet::readSheetSize(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if (type != 0x6)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readSheetSize: not a sheet zone\n"));
		return false;
	}
	long sz = libwps::readU16(input);
	if (sz < 8)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readSheetSize: block is too short\n"));
		return false;
	}
	bool ok=true;
	for (int i = 0; i < 2; i++)   // min can be invalid if the file is an extract file
	{
		f << (i==0 ? "min": "max") << "=[";
		int nCol = int(libwps::readU8(input))+1;
		f << "col=" << nCol << ",";
		auto nSheet = int(libwps::readU8(input));
		int nRow = libwps::read16(input);
		f << "row=" << nRow << ",";
		if (nSheet)
			f << "sheet=" << nSheet << ",";
		f << "],";
		if (i==0)
			continue;
		m_state->m_maxDimension=WPSVec3i(nCol,nRow,nSheet);
		if (nRow<0)
			ok=(nRow==-1 && nCol==1); // empty spreadsheet
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return ok;
}

bool QuattroSpreadsheet::readColumnRowDefaultSize(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if (type < 0xd2 || type > 0xd5)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readColumnRowDefaultSize: not a column size zone\n"));
		return false;
	}
	long sz = libwps::readU16(input);
	if (sz != 2)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readColumnRowDefaultSize: block is too short\n"));
		return false;
	}
	int val=int(libwps::readU16(input));
	if (val&0x8000)
	{
		f << "user,";
		val &= 0x7fff;
	}
	f << float(val)/20.f << ",";
	if (type==0xd2 || type==0xd4)
	{
		auto defFontType=m_mainParser.getDefaultFontType();
		auto sheet=m_state->getSheet(m_state->m_actSheet, defFontType);
		if (type==0xd2)
			sheet->m_heightDefault=float(val)/20.f;
		else
			sheet->m_widthDefault=float(val)/20.f;
	}

	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroSpreadsheet::readColumnSize(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if (type != 0xd8 && type != 0xd9)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readColumnSize: not a column size zone\n"));
		return false;
	}
	long sz = libwps::readU16(input);
	if (sz < 4)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readColumnSize: block is too short\n"));
		return false;
	}

	int col = libwps::read16(input);
	int width = libwps::readU16(input);

	auto defFontType=m_mainParser.getDefaultFontType();
	auto sheet=m_state->getSheet(m_state->m_actSheet, defFontType);
	bool ok = col >= 0 && col < sheet->m_numCols+10;
	f << "Col" << col << ":";
	if (width&0x8000)
	{
		f << "user,";
		width &= 0x7fff;
	}
	f << "width=" << float(width)/72.f << ",";
	if (ok && type==0xd8)
	{
		if (col >= sheet->m_numCols)
		{
			static bool first = true;
			if (first)
			{
				first = false;
				WPS_DEBUG_MSG(("QuattroSpreadsheet::readColumnSize: I must increase the number of columns\n"));
			}
			f << "#col[inc],";
		}
		sheet->setColumnWidth(col, width);
	}
	else if (col>256 && type==0xd8)
		f << "###,";
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	return true;
}

bool QuattroSpreadsheet::readRowSize(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if (type != 0xd6 && type != 0xd7)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readRowSize: not a row size zone\n"));
		return false;
	}
	long sz = libwps::readU16(input);
	if (sz != 4)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readRowSize: block is too short\n"));
		return false;
	}

	int row = libwps::read16(input);
	int height = libwps::readU16(input);

	f << "Row" << row << ",";
	if (height&0x8000)   // maybe set by hand?
	{
		f << "user,";
		height &= 0x7fff;
	}
	f << "h=" << float(height)/20.f << ",";
	if (type==0xd6)
	{
		if (row>=0 && m_state->m_actSheet>=0)
		{
			auto defFontType=m_mainParser.getDefaultFontType();
			m_state->getSheet(m_state->m_actSheet, defFontType)->setRowHeight(row, height);
		}
		else
		{
			WPS_DEBUG_MSG(("QuattroSpreadsheet::readRowSize: can not find the current sheet\n"));
			f << "###";
		}
	}

	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	return true;
}

bool QuattroSpreadsheet::readRowRangeSize(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if (type != 0x105 && type != 0x106)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readRowRangeSize: not a row size zone\n"));
		return false;
	}
	long sz = libwps::readU16(input);
	if (sz != 6)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readRowRangeSize: block is too short\n"));
		return false;
	}

	int minRow = libwps::read16(input);
	int maxRow = libwps::read16(input);
	int height = libwps::readU16(input);

	f << "Row" << minRow << "<->R" << maxRow << ",";
	if (height&0x8000)   // maybe set by hand?
	{
		f << "user,";
		height &= 0x7fff;
	}
	f << "h=" << float(height)/20.f << ",";
	if (type==0x105)
	{
		if (minRow>=0 && minRow<=maxRow && m_state->m_actSheet>=0)
		{
			auto defFontType=m_mainParser.getDefaultFontType();
			m_state->getSheet(m_state->m_actSheet, defFontType)->setRowHeights(minRow, maxRow, height);
		}
		else
		{
			WPS_DEBUG_MSG(("QuattroSpreadsheet::readRowSize: can not find the current sheet\n"));
			f << "###";
		}
	}

	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	return true;
}

////////////////////////////////////////////////////////////
// general
////////////////////////////////////////////////////////////
bool QuattroSpreadsheet::readBeginEndSheet(std::shared_ptr<WPSStream> const &stream, int &sheetId)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if (type != 0xca && type != 0xcb)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readBeginEndSheet: not a zoneB type\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	if (sz != 1)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readBeginEndSheet: size seems bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto sheet=int(libwps::readU8(input));
	f << "sheet=" << sheet << ",";
	if (type==0xca)
	{
		if (m_state->m_actSheet>=0)
		{
			WPS_DEBUG_MSG(("QuattroSpreadsheet::readBeginEndSheet: oops, does not find the previous end\n"));
			f << "###";
		}
		sheetId=m_state->m_actSheet=sheet;
	}
	else
	{
		if (m_state->m_actSheet!=sheet)
		{
			WPS_DEBUG_MSG(("QuattroSpreadsheet::readBeginEndSheet: oops, end sheet id does not match with begin sheet id\n"));
			f << "###";
		}
		sheetId=m_state->m_actSheet=-1;
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroSpreadsheet::readSheetName(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if (type != 0xcc)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readSheetName: not a zoneB type\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	if (sz < 1)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readSheetName: size seems bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	librevenge::RVNGString name;
	if (!m_mainParser.readCString(stream,name,sz) || name.empty())
		f << "###";
	else
	{
		f << name.cstr() << ",";
		if (m_state->m_idToSheetNameMap.find(m_state->m_actSheet)!=m_state->m_idToSheetNameMap.end())
		{
			WPS_DEBUG_MSG(("QuattroSpreadsheet::readSheetName: id dupplicated\n"));
			f << "###id";
		}
		else
			m_state->m_idToSheetNameMap[m_state->m_actSheet]=name;
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroSpreadsheet::readViewInfo(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if (type != 0x197 && type !=0x198)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readViewInfo: not a sheet zone\n"));
		return false;
	}
	long sz = libwps::readU16(input);
	long endPos=pos+4+sz;
	if (sz < 21)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readViewInfo: block is too short\n"));
		return false;
	}
	auto id=int(libwps::read8(input));
	f << "id=" << id << ",";

	int val = libwps::readU16(input); // 0|1|3|f
	f << "show=[";
	if (val&1) f << "rowHeading,";
	if (val&2) f << "colHeading,";
	if (val&4) f << "horiGrid,";
	if (val&8) f << "vertGrid,";
	// val&0x10: reserved
	val &=0xfff0;
	if (val)
		f << "f0=" << std::hex << val << std::dec << ",";
	f << "],";
	f << "range=";
	for (int i=0; i<2; ++i)
	{
		f << "C" << int(libwps::readU8(input));
		f << "S" << int(libwps::readU8(input));
		f << "R" << int(libwps::readU16(input));
		f << (i==0 ? "<->" : ",");
	}
	val = libwps::readU16(input);
	switch (val)
	{
	case 0: // default
		break;
	case 1:
		f << "title[hori],";
		break;
	case 2:
		f << "title[verti],";
		break;
	case 3:
		f << "title[both],";
		break;
	default:
		f << "##title=" << val << ",";
		break;
	}
	f << "cell[TL]=C" << int(libwps::readU8(input));
	f << "S" << int(libwps::readU8(input));
	f << "R" << int(libwps::readU16(input)) << ",";
	f << "num[row]=" << int(libwps::readU16(input)) << ",";
	f << "num[col]=" << int(libwps::readU16(input)) << ",";
	if (input->tell()!=endPos)
		ascFile.addDelimiter(input->tell(),'|');
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

////////////////////////////////////////////////////////////
// formula
////////////////////////////////////////////////////////////
bool QuattroSpreadsheet::readCell
(std::shared_ptr<WPSStream> const &stream, Vec2i actPos, WKSContentListener::FormulaInstruction &instr, int sheetId, librevenge::RVNGString const &fName)
{
	RVNGInputStreamPtr input = stream->m_input;
	instr=WKSContentListener::FormulaInstruction();
	instr.m_type=WKSContentListener::FormulaInstruction::F_Cell;
	instr.m_fileName=fName;
	bool ok = true;
	int pos[3]; // col, sheet, fl|row
	bool relative[3] = { false, false, false};
	for (int d=0; d<2; ++d) pos[d]=int(libwps::readU8(input));
	pos[2]=int(libwps::readU16(input));
	if (pos[2]&0x8000)
	{
		pos[1] = int8_t(pos[1])+sheetId;
		relative[1] = true;
	}
	if (pos[2]&0x4000)
	{
		pos[0] = int8_t(pos[0])+actPos[0];
		relative[0] = true;
	}
	if (pos[2]&0x2000)
	{
		pos[2] = actPos[1]+(int16_t((pos[2]&0x1fff)<<3)>>3);
		relative[2] = true;
	}
	else
		pos[2] &= 0x1fff;
	if (pos[0] < 0 || pos[0] > 255 || pos[2] < 0)
	{
		if (ok)
		{
			WPS_DEBUG_MSG(("QuattroSpreadsheet::readCell: can not read cell position\n"));
		}
		return false;
	}
	instr.m_position[0]=Vec2i(pos[0],pos[2]);
	instr.m_positionRelative[0]=Vec2b(relative[0],relative[2]);
	if (!fName.empty())   // external file, assume default name
	{
		librevenge::RVNGString name;
		name.sprintf("Sheet%d", pos[1]+1);
		instr.m_sheetName[0]=name;
	}
	else
		instr.m_sheetId[0]=pos[1];
	return ok;
}

bool QuattroSpreadsheet::readCellReference(std::shared_ptr<WPSStream> const &stream, long endPos,
                                           QuattroFormulaInternal::CellReference &ref,
                                           Vec2i const &cPos, int sheetId) const
{
	ref.m_cells.clear();
	RVNGInputStreamPtr input = stream->m_input;
	long pos = input->tell();
	if (pos+4>endPos) return false;
	auto type=int(libwps::readU16(input));
	int cellType=type>>12;
	if (cellType>4) return false;
	if (cellType==4)
	{
		// type==4: +6 unused bit then a 233 data field (checkme)
		WPS_DEBUG_MSG(("QuattroSpreadsheet::readCellReference: find a cell collection 4\n"));
		return false;
	}

	WKSContentListener::FormulaInstruction instr;
	if (cellType==3)
	{
		int dataSize=(type&0x3ff);
		if (pos+2+dataSize>endPos)
		{
			WPS_DEBUG_MSG(("QuattroSpreadsheet::readCellReference: can not read the cell collection data size\n"));
			return false;
		}
		if (type&0xc00) // check for deletion
		{
			input->seek(dataSize, librevenge::RVNG_SEEK_CUR);
			return true;
		}
		endPos=pos+2+dataSize;
		while (input->tell()<endPos)
		{
			QuattroFormulaInternal::CellReference cells;
			if (!readCellReference(stream, endPos, cells, cPos, sheetId))
			{
				WPS_DEBUG_MSG(("QuattroSpreadsheet::readCellReference: can not read a cell\n"));
				return false;
			}
			for (auto const &c : cells.m_cells) ref.addInstruction(c);
		}
		return true;
	}
	int const expectedSize[]= {4,8,2};
	if (pos+2+expectedSize[cellType]>endPos) return false;
	if (type&0xc00)
	{
		input->seek(expectedSize[cellType], librevenge::RVNG_SEEK_CUR);
		return true;
	}
	librevenge::RVNGString fileName;
	if ((type&0x3ff))
	{
		if (!m_mainParser.getExternalFileName((type&0x3ff), fileName))
			return false;
	}
	if (cellType==0 && pos+6<=endPos)
	{
		if (!readCell(stream, cPos, instr, sheetId, fileName))
			return false;
		ref.addInstruction(instr);
		return true;
	}
	else if (cellType==2)
	{
		auto fId=int(libwps::readU16(input));
		librevenge::RVNGString text;
		return m_mainParser.getField(fId, text, ref, fileName);
	}
	else if (cellType==1 && pos+10<=endPos)
	{
		WKSContentListener::FormulaInstruction cell2;
		if (!readCell(stream, cPos, instr, sheetId, fileName) ||
		        !readCell(stream, cPos, cell2, sheetId, fileName))
			return false;
		instr.m_type=WKSContentListener::FormulaInstruction::F_CellList;
		instr.m_position[1]=cell2.m_position[0];
		instr.m_positionRelative[1]=cell2.m_positionRelative[0];
		instr.m_sheetId[1]=cell2.m_sheetId[0];
		instr.m_sheetName[1]=cell2.m_sheetName[0];
		ref.addInstruction(instr);
		return true;
	}
	return false;
}

////////////////////////////////////////////////////////////
// send data
////////////////////////////////////////////////////////////
void QuattroSpreadsheet::sendSpreadsheet(int sId, std::vector<Vec2i> const &listGraphicCells)
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::sendSpreadsheet: I can not find the listener\n"));
		return;
	}
	auto defFontType=m_mainParser.getDefaultFontType();
	auto sheet = m_state->getSheet(sId, defFontType);
	for (auto c: listGraphicCells)
		sheet->getCell(c, defFontType).m_hasGraphic=true;
	m_listener->openSheet(sheet->getWidths(), m_state->getSheetName(sId));
	m_mainParser.sendPageGraphics(sId);
	sheet->compressRowHeights();
	auto it = sheet->m_positionToCellMap.begin();
	int prevRow = -1;
	while (it != sheet->m_positionToCellMap.end())
	{
		int row=it->first[1];
		auto const &cell=(it++)->second;
		if (row>prevRow+1)
		{
			while (row > prevRow+1)
			{
				if (prevRow != -1) m_listener->closeSheetRow();
				int numRepeat;
				float h=sheet->getRowHeight(prevRow+1, numRepeat);
				if (row<prevRow+1+numRepeat)
					numRepeat=row-1-prevRow;
				m_listener->openSheetRow(WPSRowFormat(h), numRepeat);
				prevRow+=numRepeat;
			}
		}
		if (row!=prevRow)
		{
			if (prevRow != -1) m_listener->closeSheetRow();
			m_listener->openSheetRow(WPSRowFormat(sheet->getRowHeight(++prevRow)));
		}
		if (cell.m_alignAcrossColumn)   // we must look for "merged" cell
		{
			auto firstCol=cell.position()[0], lastCol=firstCol+1;
			auto fIt=it;
			while (fIt!=sheet->m_positionToCellMap.end() && fIt->first==Vec2i(lastCol,row))
			{
				auto const &nextCell=fIt->second;
				if (nextCell.m_styleId!=cell.m_styleId)
					break;
				auto const &nextContent=nextCell.m_content;
				if ((nextContent.m_contentType== nextContent.C_NUMBER && !nextContent.isValueSet()) ||
				        nextContent.empty())
				{
					++fIt;
					++lastCol;
				}
				else
					break;
			}
			if (lastCol!=firstCol+1)
			{
				const_cast<QuattroSpreadsheetInternal::Cell &>(cell).setNumSpannedCells(Vec2i(lastCol-firstCol,1));
				it=fIt;
			}
		}
		sendCellContent(cell, sId);
	}
	if (prevRow!=-1) m_listener->closeSheetRow();
	m_listener->closeSheet();
}

namespace libwps
{
// basic function which probably does not exist on Windows, so rewrite it
static int strncasecmp(char const *s1, char const *s2, size_t n)
{
	if (n == 0)
		return 0;

	while (n-- != 0 && std::tolower(*s1) == std::tolower(*s2))
	{
		if (n == 0 || *s1 == '\0' || *s2 == '\0')
			break;
		s1++;
		s2++;
	}

	return std::tolower(*s1) - std::tolower(*s2);
}
}

void QuattroSpreadsheet::updateCellWithUserFormat(QuattroSpreadsheetInternal::Cell &cell, librevenge::RVNGString const &format)
{
	if (format.empty())
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::updateCellWithUserFormat: called with empty format\n"));
		return;
	}
	char const *ptr=format.cstr();
	auto c=char(std::toupper(*(ptr++)));
	// first N/n: numeric, T/t: date
	// all: *: fill with last data, 'string': strings, \x: x
	if (c=='N')
	{
		bool scientific=false;
		bool hasThousand=false;
		bool percent=false;
		int digits=-1;
		bool end=false;
		// numeric 0:always a digits, 9: potential digit, %: percent, ,:thousand, .:decimal, ;different format pos,equal,neg, [Ee][+-], other string
		while (*ptr)
		{
			c=char(std::toupper(*(ptr++)));
			bool ok=true;
			switch (c)
			{
			case '0':
			case '9':
				if (digits>=0 && !scientific) ++digits;
				break;
			case ',':
				if (digits<0 && !scientific)
					hasThousand=true;
				else
					ok=false;
				break;
			case 'E':
				if (digits<0)
					scientific=true;
				else
					ok=false;
				break;
			case '.':
				if (digits<0 && !scientific)
					digits=0;
				else
					ok=false;
				break;
			case '+':
			case '-':
				ok=scientific;
				break;
			case ';':
				end=true;
				break;
			case '%':
				percent=true;
				break;
			default:
				if (digits || scientific)
					end=true;
				else
					ok=false;
				break;
			}
			if (!ok)
			{
				WPS_DEBUG_MSG(("QuattroSpreadsheet::updateCellWithUserFormat: unsure how to format %s\n", format.cstr()));
				cell.setFormat(cell.F_NUMBER, 0);
				return;
			}
			if (end)
				break;
		}
		if (digits>0)
			cell.setDigits(digits);
		if (scientific)
			cell.setFormat(cell.F_NUMBER,4);
		else if (percent)
			cell.setFormat(cell.F_NUMBER,3);
		else
			cell.setFormat(cell.F_NUMBER, hasThousand ? 5 : 1);
		return;
	}
	if (c!='T')
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::updateCellWithUserFormat: unsure how to format %s\n", format.cstr()));
		return;
	}
	// d: day(1-31), dd: day(01-31), wday(sun), weekday(sunday),
	// m: month(1-12) or minute if preceded by h or hh, mm: month(01-12) or..
	// mo: month(1-12), mmo: month(01-12), mon(jan), month(january)
	// yy: year(00-99), yyyy: year(0001-9999), h: hour(0-23) except if ampm, hh: hour(00-23) except if ampm
	// mi: minute, mmi, s: second, ss
	// ampm:
	std::string dtFormat;
	bool hasHour=false;
	bool hasDate=false;
	bool inString=false;
	while (*ptr)
	{
		c=*(ptr++);
		if (inString)
		{
			if (c=='\'')
				inString=false;
			else if (c=='\\')
			{
				if (*ptr)
					dtFormat+=*(ptr++);
			}
			else
				dtFormat+=c;
			continue;
		}
		c=char(std::toupper(c));
		switch (c)
		{
		case 'A':
			if (libwps::strncasecmp(ptr, "mpm",3)==0)
			{
				dtFormat+="%p";
				ptr+=3;
				hasHour=true;
			}
			else
				dtFormat+=c;
			break;
		case 'D':
			if (*ptr=='d' || *ptr=='D')
				++ptr;
			dtFormat+="%d";
			hasDate=true;
			break;
		case 'H':
			if (*ptr=='h' || *ptr=='H')
				++ptr;
			dtFormat+="%H";
			hasHour=true;
			break;
		case 'M':
			if (*ptr=='m' || *ptr=='M')
				++ptr;
			if (*ptr=='o' || *ptr=='O')
			{
				if (libwps::strncasecmp(ptr, "onth",4)==0)
				{
					dtFormat+="%B";
					ptr+=4;
				}
				else if (libwps::strncasecmp(ptr, "on",2)==0)
				{
					dtFormat+="%b";
					ptr+=2;
				}
				else
				{
					dtFormat+="%m";
					++ptr;
				}
				hasDate=true;
			}
			else if (*ptr=='i' || *ptr=='I')
			{
				hasHour=true;
				dtFormat+="%M";
				++ptr;
			}
			else if (hasHour)
				dtFormat+="%M";
			else
				dtFormat+="%m";
			break;
		case 'S':
			if (*ptr=='s' || *ptr=='S')
				++ptr;
			dtFormat+="%S";
			hasHour=true;
			break;
		case 'W':
			if (libwps::strncasecmp(ptr, "day",3)==0)
			{
				dtFormat+="%a";
				ptr+=3;
				hasDate=true;
			}
			else if (libwps::strncasecmp(ptr, "eekday",6)==0)
			{
				dtFormat+="%A";
				ptr+=6;
				hasDate=true;
			}
			else
				dtFormat+=c;
			break;
		case 'Y':
			if (libwps::strncasecmp(ptr, "yyy",3)==0)
			{
				dtFormat+="%Y";
				ptr+=3;
				hasDate=true;
			}
			else if (libwps::strncasecmp(ptr, "y",1)==0)
			{
				dtFormat+="%y";
				ptr+=1;
				hasDate=true;
			}
			else
				dtFormat+=c;
			break;
		case '\'':
			inString=true;
			break;
		case '\\':
			if (*ptr)
				dtFormat+=*(ptr++);
			break;
		default:
			dtFormat+=c;
			break;
		}
	}
	cell.setDTFormat((hasDate||!hasHour) ? cell.F_DATE : cell.F_TIME, dtFormat);
}

void QuattroSpreadsheet::sendCellContent(QuattroSpreadsheetInternal::Cell const &cell, int sheetId)
{
	if (m_listener.get() == nullptr)
	{
		WPS_DEBUG_MSG(("QuattroSpreadsheet::sendCellContent: I can not find the listener\n"));
		return;
	}

	libwps_tools_win::Font::Type fontType = cell.m_fontType;
	m_listener->setFont(cell.getFont());

	QuattroSpreadsheetInternal::Cell finalCell(cell);
	auto &content=finalCell.m_content;
	for (auto &f : content.m_formula)
	{
		if (f.m_type==WKSContentListener::FormulaInstruction::F_Cell ||
		        f.m_type==WKSContentListener::FormulaInstruction::F_CellList)
		{
			int dim=f.m_type==WKSContentListener::FormulaInstruction::F_Cell ? 1 : 2;
			for (int i=0; i<dim; ++i)
			{
				if (f.m_sheetId[i]>=0 && f.m_sheetName[i].empty() && (f.m_sheetId[i]!=sheetId || !f.m_fileName.empty()))
					f.m_sheetName[i]=getSheetName(f.m_sheetId[i]);
			}
			continue;
		}
		if (f.m_type!=WKSContentListener::FormulaInstruction::F_Text)
			continue;
		std::string &text=f.m_content;
		librevenge::RVNGString finalString=libwps_tools_win::Font::unicodeString(text, fontType);
		if (finalString.empty())
			text.clear();
		else
			text=finalString.cstr();
	}
	if ((finalCell.m_fileFormat>>4)==7)
	{
		auto it=m_state->m_idToUserFormatMap.find(finalCell.m_fileFormat&0xf);
		if (it==m_state->m_idToUserFormatMap.end() || it->second.empty())
		{
			WPS_DEBUG_MSG(("QuattroSpreadsheet::sendCellContent: can not find an user format\n"));
		}
		else
			updateCellWithUserFormat(finalCell, it->second);
	}
	else
		finalCell.updateFormat();
	m_listener->openSheetCell(finalCell, content);
	if (cell.m_hasGraphic)
		m_mainParser.sendGraphics(sheetId, cell.position());
	if (cell.m_content.m_textEntry.valid())
	{
		if (!cell.m_stream || !cell.m_stream->m_input)
		{
			WPS_DEBUG_MSG(("QuattroSpreadsheet::sendCellContent: oops can not find the text's stream\n"));
		}
		else
		{
			auto input=cell.m_stream->m_input;
			input->seek(cell.m_content.m_textEntry.begin(), librevenge::RVNG_SEEK_SET);
			bool prevEOL=false;
			std::string text;
			while (input->tell()<=cell.m_content.m_textEntry.end())
			{
				bool last=input->isEnd() || input->tell()>=cell.m_content.m_textEntry.end();
				auto c=last ? '\0' : char(libwps::readU8(input));
				if ((c==0 || c==0xa || c==0xd) && !text.empty())
				{
					m_listener->insertUnicodeString(libwps_tools_win::Font::unicodeString(text, fontType));
					text.clear();
				}
				if (last) break;
				if (c==0xd)
				{
					m_listener->insertEOL();
					prevEOL=true;
				}
				else if (c==0xa)
				{
					if (!prevEOL)
					{
						WPS_DEBUG_MSG(("QuattroSpreadsheet::sendCellContent: find 0xa without 0xd\n"));
					}
					prevEOL=false;
				}
				else
				{
					if (c)
						text.push_back(c);
					prevEOL=false;
				}
			}
		}
	}
	m_listener->closeSheetCell();
}

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
