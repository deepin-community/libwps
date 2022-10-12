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
#include <stack>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WPSCell.h"
#include "WKSContentListener.h"
#include "WPSEntry.h"
#include "WPSFont.h"
#include "WPSTable.h"

#include "QuattroDos.h"

#include "QuattroDosSpreadsheet.h"

namespace QuattroDosSpreadsheetInternal
{

///////////////////////////////////////////////////////////////////
//! a class used to store a style of a cell in QuattroDosSpreadsheet
struct Style final : public WPSCellFormat
{
	//! construtor
	explicit Style(libwps_tools_win::Font::Type type)
		: WPSCellFormat()
		, m_fontType(type)
		, m_fileFormat(0xFF)
		, m_extra("")
	{
	}
	Style(Style const &)=default;
	Style &operator=(Style const &)=default;
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
	return m_extra==st.m_extra;
}

///////////////////////////////////////////////////////////////////
//! the style manager
class StyleManager
{
public:
	StyleManager()
		: m_idStyleMap() {}
	//! add a new style and returns its id
	void add(int id, Style const &st)
	{
		if (m_idStyleMap.find(id)!=m_idStyleMap.end())
		{
			WPS_DEBUG_MSG(("QuattroDosParserInternal::StyleManager::add style %d already exists\n", id));
			return;
		}
		m_idStyleMap.insert(std::map<int, Style>::value_type(id,st));
	}
	//! returns the style with id
	bool get(int id, Style &style) const
	{
		if (m_idStyleMap.find(id)==m_idStyleMap.end())
		{
			WPS_DEBUG_MSG(("QuattroDosParserInternal::StyleManager::get can not find style %d\n", id));
			return false;
		}
		style=m_idStyleMap.find(id)->second;
		return true;
	}
	//! print a style
	void print(int id, std::ostream &o) const
	{
		if (m_idStyleMap.find(id)!=m_idStyleMap.end())
			o << ", style=" << m_idStyleMap.find(id)->second;
		else
		{
			WPS_DEBUG_MSG(("QuattroDosParserInternal::StyleManager::print: can not find a style\n"));
			o << ", ###style=" << id;
		}
	}

protected:
	//! the styles
	std::map<int, Style> m_idStyleMap;
};

//! a cellule of a Quattro spreadsheet
class Cell final : public WPSCell
{
public:
	/// constructor
	explicit Cell(libwps_tools_win::Font::Type type)
		: m_fontType(type)
		, m_fileFormat(0xFF)
		, m_content() { }

	//! operator<<
	friend std::ostream &operator<<(std::ostream &o, Cell const &cell);

	//! call when a cell must be send
	bool send(WPSListenerPtr &/*listener*/) final;

	//! call when the content of a cell must be send
	bool sendContent(WPSListenerPtr &/*listener*/) final
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheetInternal::Cell::sendContent: must not be called\n"));
		return false;
	}

	//! update the cell format using file format
	void updateFormat()
	{
		switch ((m_fileFormat>>4)&7)
		{
		case 0: // fixed
			setFormat(F_NUMBER, 1);
			setDigits(m_fileFormat&0xF);
			break;
		case 1: // scientific
			setFormat(F_NUMBER, 2);
			setDigits(m_fileFormat&0xF);
			break;
		case 2: // currency
			setFormat(F_NUMBER, 4);
			setDigits(m_fileFormat&0xF);
			break;
		case 3: // percent
			setFormat(F_NUMBER, 3);
			setDigits(m_fileFormat&0xF);
			break;
		case 4: // decimal
			setFormat(F_NUMBER, 1);
			setDigits(m_fileFormat&0xF);
			break;
		case 5:
			switch (m_fileFormat&0xF)
			{
			case 4: // a date but no sure which format is good
				setDTFormat(F_DATE, "%m/%d/%y");
				break;
			case 5:
				setDTFormat(F_DATE, "%m/%d");
				break;
			default:
				WPS_DEBUG_MSG(("QuattroDosSpreadsheetInternal::Cell::updateFormat: unknown format %x\n", unsigned(m_fileFormat)));
				break;
			}
			break;
		case 6:
			WPS_DEBUG_MSG(("QuattroDosSpreadsheetInternal::Cell::updateFormat: unknown format %x\n", unsigned(m_fileFormat)));
			break;
		case 7:
			switch (m_fileFormat&0xF)
			{
			case 0: // +/- : kind of bool
				setFormat(F_BOOLEAN);
				break;
			case 1:
				setFormat(F_NUMBER, 0);
				break;
			case 2:
				setDTFormat(F_DATE, "%d %B %y");
				break;
			case 3:
				setDTFormat(F_DATE, "%d %B");
				break;
			case 4:
				setDTFormat(F_DATE, "%B %y");
				break;
			case 5:
				setFormat(F_TEXT);
				break;
			case 6:
				setFormat(F_TEXT);
				m_font.m_attributes |= WPS_HIDDEN_BIT;
				break;
			case 7:
				setDTFormat(F_TIME, "%I:%M:%S%p");
				break;
			case 8:
				setDTFormat(F_TIME, "%I:%M%p");
				break;
			case 9:
				setDTFormat(F_DATE, "%m/%d/%y");
				break;
			case 0xa:
				setDTFormat(F_DATE, "%m/%d");
				break;
			case 0xb:
				setDTFormat(F_TIME, "%H:%M:%S");
				break;
			case 0xc:
				setDTFormat(F_TIME, "%H:%M");
				break;
			case 0xd:
				setFormat(F_TEXT);
				break;
			case 0xf: // automatic
				break;
			default:
				break;
			}
			break;
		default:
			break;
		}
	}
	//! font encoding type
	libwps_tools_win::Font::Type m_fontType;
	//! the file format
	int m_fileFormat;
	//! the content
	WKSContentListener::CellContent m_content;
};

bool Cell::send(WPSListenerPtr &/*listener*/)
{
	WPS_DEBUG_MSG(("QuattroDosSpreadsheetInternal::Cell::send: must not be called\n"));
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

///////////////////////////////////////////////////////////////////
//! the spreadsheet of a QuattroDos Spreadsheet
class Spreadsheet
{
public:
	//! the spreadsheet type
	enum Type { T_Spreadsheet, T_Filter, T_Report };

	//! a constructor
	Spreadsheet(Type type=T_Spreadsheet, int id=0)
		: m_type(type)
		, m_id(id)
		, m_numCols(0)
		, m_widthCols()
		, m_rowHeightMap()
		, m_heightDefault(16)
		, m_widthDefault(76)
		, m_positionToCellMap()
		, m_lastCellPos()
		, m_rowPageBreaksList() {}
	//! return a cell corresponding to a spreadsheet, create one if needed
	Cell &getCell(Vec2i const &pos, libwps_tools_win::Font::Type type)
	{
		if (m_positionToCellMap.find(pos)==m_positionToCellMap.end())
		{
			Cell cell(type);
			cell.setPosition(pos);
			m_positionToCellMap.insert(std::map<Vec2i, Cell>::value_type(pos,cell));
		}
		m_lastCellPos=pos;
		return m_positionToCellMap.find(pos)->second;
	}
	//! returns the last cell
	Cell *getLastCell()
	{
		if (m_positionToCellMap.find(m_lastCellPos)==m_positionToCellMap.end())
			return nullptr;
		return &m_positionToCellMap.find(m_lastCellPos)->second;
	}
	void insertChartPositions(std::map<Vec2i,Vec2i> const &chartPosMap, libwps_tools_win::Font::Type type)
	{
		int maxCol=-1;
		for (auto it : chartPosMap)
		{
			if (it.second[0]>maxCol) maxCol=it.second[0];
			if (m_positionToCellMap.find(it.first)!=m_positionToCellMap.end())
				continue;
			Cell cell(type);
			cell.setPosition(it.first);
			m_positionToCellMap.insert(std::map<Vec2i, Cell>::value_type(it.first,cell));
		}
		if (maxCol>=int(m_widthCols.size()))
			m_widthCols.resize(size_t(maxCol+1),-1);
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
		WPSColumnFormat defWidth(static_cast<float>(m_widthDefault)), actWidth;
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
	//! returns the row size in point
	float getRowHeight(int row) const
	{
		auto rIt=m_rowHeightMap.lower_bound(Vec2i(-1,row));
		if (rIt!=m_rowHeightMap.end() && rIt->first[0]<=row && rIt->first[1]>=row)
			return float(rIt->second);
		return float(m_heightDefault);
	}
	//! returns the height of a row in point and updated repeated row
	float getRowHeight(int row, int &numRepeated) const
	{
		auto rIt=m_rowHeightMap.lower_bound(Vec2i(-1,row));
		if (rIt!=m_rowHeightMap.end() && rIt->first[0]<=row && rIt->first[1]>=row)
		{
			numRepeated=rIt->first[1]-row+1;
			return float(rIt->second);
		}
		numRepeated=10000;
		return float(m_heightDefault);
	}
	//! set the rows size
	void setRowHeight(int row, int h)
	{
		if (h>=0)
			m_rowHeightMap[Vec2i(row,row)]=h;
	}
	//! return the size of the zone containning two cells
	Vec2f getZoneSize(Vec2i const &cell0, Vec2i const &cell1) const
	{
		int w=0;
		auto numCol=int(m_widthCols.size());
		for (int i=cell0[0]; i<=cell1[0]; ++i)
		{
			if (i<0 || i>=numCol)
				w+=m_widthDefault;
			else
				w+=m_widthCols[size_t(i)];
		}
		int h=0;
		auto rIt=m_rowHeightMap.lower_bound(Vec2i(-1,cell0[1]));
		for (int i=cell0[1]; i<=cell1[1]; ++i)
		{
			if (rIt==m_rowHeightMap.end())
				h+=m_heightDefault;
			else if (rIt->first[0]<=i && rIt->first[1]>=i)
				h+=rIt->second;
			else
			{
				if (rIt->first[1]<i) ++rIt;
				if (rIt==m_rowHeightMap.end() || rIt->first[0]>i)
					h+=m_heightDefault;
				else
					h+=rIt->second;
			}
		}
		return Vec2f(float(w),float(h));
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
				if (actHeight==m_heightDefault)
					actPos[1]=rIt.first[0]-1;
				else
				{
					if (actPos[1]>=actPos[0])
						m_rowHeightMap[actPos]=actHeight;
					actHeight=m_heightDefault;
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
	//! returns true if the spreedsheet is empty
	bool empty() const
	{
		return m_positionToCellMap.empty();
	}
	//! the spreadsheet type
	Type m_type;
	//! the spreadsheet id
	int m_id;
	/** the number of columns */
	int m_numCols;

	/** the column size in TWIP (?) */
	std::vector<int> m_widthCols;
	/** the map Vec2i(min row, max row) to size in points */
	std::map<Vec2i,int> m_rowHeightMap;
	/** the default row size in point */
	int m_heightDefault;
	/** the default width size in point */
	int m_widthDefault;
	/** a map cell to not empty cells */
	std::map<Vec2i, Cell> m_positionToCellMap;
	/** the last cell position */
	Vec2i m_lastCellPos;
	/** the list of row page break */
	std::vector<int> m_rowPageBreaksList;

};

//! the state of QuattroDosSpreadsheet
struct State
{
	//! constructor
	State()
		:  m_eof(-1)
		, m_version(-1)
		, m_styleManager()
		, m_spreadsheetList()
		, m_spreadsheetStack()
	{
		pushNewSheet(std::shared_ptr<Spreadsheet>(new Spreadsheet(Spreadsheet::T_Spreadsheet, 0)));
	}
	//! returns the maximal spreadsheet
	int getMaximalSheet(Spreadsheet::Type type=Spreadsheet::T_Spreadsheet) const
	{
		int max=-1;
		for (auto sheet : m_spreadsheetList)
		{
			if (!sheet || sheet->m_type != type || sheet->m_id<=max || sheet->empty()) continue;
			max=sheet->m_id;
		}
		return max;
	}
	//! returns the ith real spreadsheet
	std::shared_ptr<Spreadsheet> getSheet(Spreadsheet::Type type, int id)
	{
		for (auto sheet : m_spreadsheetList)
		{
			if (!sheet || sheet->m_type!=type || sheet->m_id!=id)
				continue;
			return sheet;
		}
		return std::shared_ptr<Spreadsheet>();
	}
	//! returns the ith spreadsheet
	static librevenge::RVNGString getSheetName(int id)
	{
		librevenge::RVNGString name;
		name.sprintf("Sheet%d", id+1);
		return name;
	}
	//! returns the actual sheet
	Spreadsheet &getActualSheet()
	{
		return *m_spreadsheetStack.top();
	}
	//! create a new sheet and stack id
	void pushNewSheet(std::shared_ptr<Spreadsheet> sheet)
	{
		if (!sheet)
		{
			WPS_DEBUG_MSG(("QuattroDosSpreadsheetInternal::State::pushSheet: can not find the sheet\n"));
			return;
		}
		m_spreadsheetStack.push(sheet);
		m_spreadsheetList.push_back(sheet);
	}
	//! try to pop the actual sheet
	bool popSheet()
	{
		if (m_spreadsheetStack.size()<=1)
		{
			WPS_DEBUG_MSG(("QuattroDosSpreadsheetInternal::State::popSheet: can pop the main sheet\n"));
			return false;
		}
		m_spreadsheetStack.pop();
		return true;
	}
	//! the last file position
	long m_eof;
	//! the file version
	int m_version;
	//! the style manager
	StyleManager m_styleManager;

	//! the list of spreadsheet ( first: main spreadsheet, other report spreadsheet )
	std::vector<std::shared_ptr<Spreadsheet> > m_spreadsheetList;
	//! the stack of spreadsheet id
	std::stack<std::shared_ptr<Spreadsheet> > m_spreadsheetStack;
};

}

// constructor, destructor
QuattroDosSpreadsheet::QuattroDosSpreadsheet(QuattroDosParser &parser)
	: m_input(parser.getInput())
	, m_listener()
	, m_mainParser(parser)
	, m_state(new QuattroDosSpreadsheetInternal::State)
	, m_asciiFile(parser.ascii())
{
}

QuattroDosSpreadsheet::~QuattroDosSpreadsheet()
{
}

int QuattroDosSpreadsheet::version() const
{
	if (m_state->m_version<0)
		m_state->m_version=m_mainParser.version();
	return m_state->m_version;
}

bool QuattroDosSpreadsheet::checkFilePosition(long pos)
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

int QuattroDosSpreadsheet::getNumSpreadsheets() const
{
	return m_state->getMaximalSheet(QuattroDosSpreadsheetInternal::Spreadsheet::T_Spreadsheet)+1;
}

librevenge::RVNGString QuattroDosSpreadsheet::getSheetName(int id) const
{
	return m_state->getSheetName(id);
}

////////////////////////////////////////////////////////////
// low level

////////////////////////////////////////////////////////////
//   parse sheet data
////////////////////////////////////////////////////////////
bool QuattroDosSpreadsheet::readSheetSize()
{
	libwps::DebugStream f;
	long pos = m_input->tell();
	long type = libwps::read16(m_input);
	if (type != 0x6)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readSheetSize: not a sheet zone\n"));
		return false;
	}
	long sz = libwps::readU16(m_input);
	int const vers=version();
	if (sz < (vers>=2 ? 12 : 8))
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readSheetSize: block is too short\n"));
		return false;
	}
	f << "Entries(SheetSize):";
	for (int i = 0; i < (vers>=2 ? 3 : 2); i++)   // always 2 zero ?
	{
		int val = libwps::read16(m_input);
		if (!val) continue;
		f << "f" << i << "=" << std::hex << val << std::dec << ",";
	}
	int nCol = libwps::read16(m_input)+1;
	f << "nCols=" << nCol << ",";
	int nRow =  libwps::read16(m_input);
	f << "nRow=" << nRow << ",";
	int nSheet =  vers<=1 ? 0 : libwps::read16(m_input);
	if (nSheet>0)
		f << "nSheet=" << nSheet << ",";
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	// empty spreadsheet
	if (nRow==-1 && nCol==0) return true;
	if (nRow < 0 || nCol <= 0) return false;
	m_state->getActualSheet().setColumnWidth(nCol-1);
	return true;

}

bool QuattroDosSpreadsheet::readRowSize()
{
	libwps::DebugStream f;
	long pos = m_input->tell();
	long type = libwps::read16(m_input);
	if (type != 0xe0)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readRowSize: not a row size zone\n"));
		return false;
	}
	long sz = libwps::readU16(m_input);
	if (sz < 3)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readRowSize: block is too short\n"));
		return false;
	}

	int row = libwps::read16(m_input);
	int height = libwps::readU8(m_input);

	bool ok = row >= 0;
	f << "Entries(Row):Row" << row << "";
	if (!ok) f << "###";
	f << ":height=" << height << ",";

	if (ok)
		m_state->getActualSheet().setRowHeight(row, height); // point to TWIP
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	return ok;
}

bool QuattroDosSpreadsheet::readColumnSize()
{
	libwps::DebugStream f;
	long pos = m_input->tell();
	long type = libwps::read16(m_input);
	if (type != 0x8 && type != 0xe2)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readColumnSize: not a column size zone\n"));
		return false;
	}
	long sz = libwps::readU16(m_input);
	if (sz < 3)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readColumnSize: block is too short\n"));
		return false;
	}

	int col = libwps::read16(m_input);
	int width = libwps::readU8(m_input);

	bool ok = col >= 0 && col < m_state->getActualSheet().m_numCols+10;
	f << "Entries(Column):Col" << col << "";
	if (!ok) f << "###";
	f << ":width=" << width << ",";

	if (ok)
	{
		if (col >= m_state->getActualSheet().m_numCols)
		{
			static bool first = true;
			if (first)
			{
				first = false;
				WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readColumnSize: I must increase the number of columns\n"));
			}
			f << "#col[inc],";
		}
		// checkme: unit in character(?) -> TWIP
		m_state->getActualSheet().setColumnWidth(col, width*105);
	}
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	return ok;
}

bool QuattroDosSpreadsheet::readHiddenColumns()
{
	libwps::DebugStream f;
	long pos = m_input->tell();
	long type = libwps::read16(m_input);
	if (type != 0x64)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readHiddenColumns: not a column hidden zone\n"));
		return false;
	}
	long sz = libwps::readU16(m_input);
	if (sz != 32)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readHiddenColumns: block size seems bad\n"));
		return false;
	}

	f << "Entries(HiddenCol):";
	for (int i=0; i<32; ++i)
	{
		auto val=int(libwps::readU8(m_input));
		if (!val) continue;
		static bool first=true;
		if (first)
		{
			WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readHiddenColumns: find some hidden col, ignored\n"));
			first=false;
		}
		// checkme
		for (int j=0, depl=1; j<8; ++j, depl<<=1)
		{
			if (val&depl)
				f << "col" << j+8*i << "[hidden],";
		}
	}
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	return true;
}

bool QuattroDosSpreadsheet::readCellStyle()
{
	libwps::DebugStream f;
	long pos = m_input->tell();
	long type = libwps::read16(m_input);
	int const vers=version();
	if (type != 0xd8)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readCellStyle: not a style zone\n"));
		return false;
	}
	long sz = libwps::readU16(m_input);
	f << "Entries(CellStyle):";
	if ((vers==1&&(sz%12)!=0) || (vers>1 && sz!=0x16))
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readCellStyle: size seems bad\n"));
		f << "###";
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		return true;
	}
	if (vers>1)
	{
		auto pId=int(libwps::readU16(m_input));
		QuattroDosSpreadsheetInternal::Style style(m_mainParser.getDefaultFontType());
		auto id=int(libwps::readU16(m_input));
		f.str("");
		if (!m_state->m_styleManager.get(pId, style))
			f << "###";
		if (pId!=0xFF0F)
		{
			f << "Ce" << (pId>>8);
			if (pId&0xFF) f << "-" << (pId&0xFF);
			f << "[parent],";
		}
		int val;
		for (int i=0; i<4; ++i)   // small int, unknown
		{
			val=int(libwps::readU8(m_input));
			if (val) f << "f" << i << "=" << val << ",";
		}
		WPSFont font;
		auto flags = int(libwps::readU16(m_input));
		uint32_t attributes = 0;
		if (flags & 1) attributes |= WPS_BOLD_BIT;
		if (flags & 2) attributes |= WPS_ITALICS_BIT;
		if (flags & 8) attributes |= WPS_UNDERLINE_BIT;

		font.m_attributes=attributes;
		if (flags & 0xFFF4)
			f << "fl=" << std::hex << (flags & 0xFFF4) << std::dec << ",";
		auto fId=int(libwps::readU16(m_input));
		f << "fId=" << fId << ",";
		auto fSize = int(libwps::readU16(m_input));
		if (fSize >= 1 && fSize <= 50)
			font.m_size=double(fSize);
		else
			f << "###fSize=" << fSize << ",";
		auto color = int(libwps::readU16(m_input));
		if (color && !m_mainParser.getColor(color, font.m_color))
		{
			WPS_DEBUG_MSG(("QuattroDosParser::readCellStyle: unknown color\n"));
			f << "##color=" << color << ",";
		}
		style.setFont(font);

		val=int(libwps::readU8(m_input));
		if (val)
		{
			for (int i=0, depl=0; i<4; ++i, depl+=2)
			{
				int bd=(val>>depl)&3;
				if (!bd) continue;
				WPSBorder border;
				switch (bd)
				{
				case 1: // normal
					break;
				case 2:
					border.m_type = WPSBorder::Double;
					break;
				case 3:
					border.m_width=2;
					break;
				default:
					break;
				}
				int const wh[]= {WPSBorder::TopBit, WPSBorder::LeftBit, WPSBorder::BottomBit, WPSBorder::RightBit};
				style.setBorders(wh[i], border);
			}
		}
		val=int(libwps::readU8(m_input)); // 8|10|18
		if (val) f << "g0=" << val << ",";
		val=int(libwps::readU8(m_input));
		switch (val)
		{
		case 0:
			break;
		case 1:
			style.setBackgroundColor(WPSColor(0x80,0x80,0x80));
			break;
		case 2:
			style.setBackgroundColor(WPSColor::black());
			break;
		default:
			if ((val&3)==3)
			{
				WPSColor col;
				if (!m_mainParser.getColor(val>>2, col))
				{
					WPS_DEBUG_MSG(("QuattroDosParser::readCellStyle: unknown background color\n"));
					f << "##color=" << val << ",";
				}
				else
					style.setBackgroundColor(col);
			}
			else
				f << "##background=" << val << ",";
			break;
		}
		style.m_fileFormat=int(libwps::readU8(m_input));
		for (int i=0; i<2; ++i)   // g1=0|f8
		{
			val=int(libwps::readU8(m_input));
			if (val) f << "g" << i+1 << "=" << val << ",";
		}

		style.m_extra=f.str();
		m_state->m_styleManager.add(id, style);
		f.str("");
		f << "Entries(CellStyle):Ce" << (id>>8);
		if (id&0xFF)
			f << "-" << (id&0xFF) << "#";
		f << "," << style << ",";

		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		return true;
	}
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	auto N=int(sz/12);
	for (int i=0; i<N; ++i)
	{
		pos=m_input->tell();
		QuattroDosSpreadsheetInternal::Style style(m_mainParser.getDefaultFontType());
		f.str("");
		auto id=int(libwps::readU16(m_input));

		WPSFont font;
		auto flags = int(libwps::readU16(m_input));
		uint32_t attributes = 0;
		if (flags & 1) attributes |= WPS_BOLD_BIT;
		if (flags & 2) attributes |= WPS_ITALICS_BIT;
		if (flags & 8) attributes |= WPS_UNDERLINE_BIT;

		font.m_attributes=attributes;
		if (flags & 0xFFF4)
			f << "fl=" << std::hex << (flags & 0xFFF4) << std::dec << ",";
		auto fId=int(libwps::readU16(m_input));
		f << "fId=" << fId << ",";
		auto fSize = int(libwps::readU16(m_input));
		if (fSize >= 1 && fSize <= 50)
			font.m_size=double(fSize);
		else
			f << "###fSize=" << fSize << ",";
		auto color = int(libwps::readU16(m_input));
		if (color && !m_mainParser.getColor(color, font.m_color))
		{
			WPS_DEBUG_MSG(("QuattroDosParser::readCellStyle: unknown color\n"));
			f << "##color=" << color << ",";
		}
		style.setFont(font);
		auto val=int(libwps::readU16(m_input));
		f << "f0=" << std::hex << val << std::dec << ",";

		style.m_extra=f.str();
		m_state->m_styleManager.add(id, style);
		f.str("");
		f << "CellStyle-" << i << ":Ce" << id << "," << style << ",";

		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		m_input->seek(pos+12, librevenge::RVNG_SEEK_SET);
	}
	return true;
}

bool QuattroDosSpreadsheet::readCellProperty()
{
	libwps::DebugStream f;
	long pos = m_input->tell();
	long type = libwps::read16(m_input);
	auto defFontType=m_mainParser.getDefaultFontType();

	if (type != 0x9d)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readCellProperty: not a cell property zone zone\n"));
		return false;
	}
	f << "Entries(CellProperty):";
	long sz = libwps::readU16(m_input);
	if (sz!=7)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readCellProperty: the size seems bad\n"));
		f << "###sz";
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		return true;
	}
	auto format=int(libwps::readU8(m_input));
	auto col=int(libwps::read16(m_input));
	auto row=int(libwps::read16(m_input));
	QuattroDosSpreadsheetInternal::Cell emptyCell(defFontType);
	auto *cell=&emptyCell;
	if (col<0 || row<0)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readCellProperty: the position seems bad\n"));
		f << "###";
	}
	else
		cell = &m_state->getActualSheet().getCell(Vec2i(col,row),defFontType);
	f << "C" << Vec2i(col,row) << ",";
	auto flag=int(libwps::readU8(m_input));
	auto id=int(libwps::readU8(m_input));
	if (id&0x80)
	{
		f << "Ce" << (id&0x7f) << ",";
		QuattroDosSpreadsheetInternal::Style style(cell->m_fontType);
		if (!m_state->m_styleManager.get((id&0x7f), style))
		{
			WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readCellProperty: can not find a style\n"));
			f << "###";
		}
		else
		{
			if (style.m_fileFormat==0xFF)
				cell->m_fileFormat=style.m_fileFormat;
			cell->m_fontType=style.m_fontType;
			cell->setFont(style.getFont());
			cell->setBackgroundColor(style.backgroundColor());
			if (style.hasBorders())
				cell->setBorders(style.borders());
		}
	}
	else if (id)
	{
		if (id&0x7c)
			f << "Fo" << (id>>2) << ",";
		WPSFont font;
		if (!m_mainParser.getFont(id>>2, font, cell->m_fontType))
		{
			WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readCellProperty: can not find a font\n"));
			f << "###";
		}
		else
			cell->setFont(font);
		if (id&0x3)
			f << "f0=" << (id&3) << ",";
	}
	if (format!=0xFF)
	{
		cell->m_fileFormat=format;
		f << "form=" << std::hex << format << std::dec << ",";
	}
	switch (flag>>6)
	{
	case 1:
		cell->setHAlignment(WPSCellFormat::HALIGN_LEFT);
		f << "left,";
		break;
	case 2:
		cell->setHAlignment(WPSCellFormat::HALIGN_RIGHT);
		f << "right,";
		break;
	case 3:
		cell->setHAlignment(WPSCellFormat::HALIGN_CENTER);
		f << "center,";
		break;
	default:
	case 0:
		break;
	}
	for (int i=0; i<2; ++i)
	{
		int bd=(flag>>(2*i))&0x3;
		if (!bd) continue;
		f << "bord" << (i==0 ? "T" : "L");
		WPSBorder border;
		switch (bd)
		{
		case 2:
			border.m_type = WPSBorder::Double;
			f << "[double]";
			break;
		case 3:
			border.m_width=2;
			f << "[w=2]";
			break;
		default:
		case 1: // normal
			break;
		}
		f << ",";
		cell->setBorders(i==0 ? WPSBorder::TopBit : WPSBorder::LeftBit, border);
	}
	switch ((flag>>4)&3)
	{
	case 1:
		cell->setBackgroundColor(WPSColor(0x80,0x80,0x80));
		f << "back[grey],";
		break;
	case 2:
		cell->setBackgroundColor(WPSColor::black());
		f << "back[black],";
		break;
	case 3:
		f << "#back[color]=3";
		break;
	case 0:
	default:
		break;
	}
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

bool QuattroDosSpreadsheet::readUserStyle()
{
	libwps::DebugStream f;
	long pos = m_input->tell();
	long type = libwps::read16(m_input);
	int const vers=version();
	if (type != 0xc9)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readUserStyle: not a style zone\n"));
		return false;
	}
	long sz = libwps::readU16(m_input);
	f << "Entries(UserStyle):";
	if ((vers==1&&sz!=0x2a) || (vers>1 && sz!=0x24))
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readUserStyle: size seems bad\n"));
		f << "###";
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		return true;
	}
	if (vers>1)
	{
		f.str("");
		librevenge::RVNGString name("");
		if (!m_mainParser.readPString(name,15))
			f << "##sSz,";
		else if (!name.empty())
			f << name.cstr() << ",";
		m_input->seek(pos+20, librevenge::RVNG_SEEK_SET);

		QuattroDosSpreadsheetInternal::Style style(m_mainParser.getDefaultFontType());
		auto id=int(libwps::readU16(m_input));
		int val;
		for (int i=0; i<4; ++i)   // small int, unknown
		{
			val=int(libwps::readU8(m_input));
			if (val) f << "f" << i << "=" << val << ",";
		}

		WPSFont font;
		auto flags = int(libwps::readU16(m_input));
		uint32_t attributes = 0;
		if (flags & 1) attributes |= WPS_BOLD_BIT;
		if (flags & 2) attributes |= WPS_ITALICS_BIT;
		if (flags & 8) attributes |= WPS_UNDERLINE_BIT;

		font.m_attributes=attributes;
		if (flags & 0xFFF4)
			f << "fl=" << std::hex << (flags & 0xFFF4) << std::dec << ",";
		auto fId=int(libwps::readU16(m_input));
		f << "fId=" << fId << ",";
		auto fSize = int(libwps::readU16(m_input));
		if (fSize >= 1 && fSize <= 50)
			font.m_size=double(fSize);
		else
			f << "###fSize=" << fSize << ",";
		auto color = int(libwps::readU16(m_input));
		if (color && !m_mainParser.getColor(color, font.m_color))
		{
			WPS_DEBUG_MSG(("QuattroDosParser::readCellStyle: unknown color\n"));
			f << "##color=" << color << ",";
		}
		style.setFont(font);
		val=int(libwps::readU8(m_input));
		if (val)
		{
			for (int i=0, depl=0; i<4; ++i, depl+=2)
			{
				int bd=(val>>depl)&3;
				if (!bd) continue;
				WPSBorder border;
				switch (bd)
				{
				case 1: // normal
					break;
				case 2:
					border.m_type = WPSBorder::Double;
					break;
				case 3:
					border.m_width=2;
					break;
				default:
					break;
				}
				int const wh[]= {WPSBorder::TopBit, WPSBorder::LeftBit, WPSBorder::BottomBit, WPSBorder::RightBit};
				style.setBorders(wh[i], border);
			}
		}
		val=int(libwps::readU8(m_input)); // 8|10|18
		if (val) f << "g0=" << val << ",";
		val=int(libwps::readU8(m_input));
		switch (val)
		{
		case 0:
			break;
		case 1:
			style.setBackgroundColor(WPSColor(0x80,0x80,0x80));
			break;
		case 2:
			style.setBackgroundColor(WPSColor::black());
			break;
		default:
			f << "##background=" << val << ",";
			break;
		}
		style.m_fileFormat=int(libwps::readU8(m_input));
		for (int i=0; i<2; ++i)   // g1=0|f8
		{
			val=int(libwps::readU8(m_input));
			if (val) f << "g" << i+1 << "=" << val << ",";
		}

		style.m_extra=f.str();
		m_state->m_styleManager.add(id, style);
		f.str("");
		f << "Entries(UserStyle):";
		if (id!=0xFF0F)
		{
			f << "Ce" << (id>>8) << "-" << (id&0xFF);
			if ((id&0xFF)!=8)
				f  << "#";
		}
		f << "," << style << ",";

		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		return true;
	}

	QuattroDosSpreadsheetInternal::Style style(m_mainParser.getDefaultFontType());
	auto id=int(libwps::readU16(m_input));
	f.str("");
	auto flags = int(libwps::readU16(m_input));

	WPSFont font;
	uint32_t attributes = 0;
	if (flags & 1) attributes |= WPS_BOLD_BIT;
	if (flags & 2) attributes |= WPS_ITALICS_BIT;
	if (flags & 8) attributes |= WPS_UNDERLINE_BIT;

	font.m_attributes=attributes;
	if (flags & 0xFFF4)
		f << "font[fl]=" << std::hex << (flags & 0xFFF4) << std::dec << ",";
	auto fId=int(libwps::readU16(m_input));
	f << "fId=" << fId << ",";
	auto fSize = int(libwps::readU16(m_input));
	if (fSize >= 1 && fSize <= 50)
		font.m_size=double(fSize);
	else
		f << "###fSize=" << fSize << ",";
	auto color = int(libwps::readU16(m_input));
	if (color && !m_mainParser.getColor(color, font.m_color))
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readUserStyle: unknown color\n"));
		f << "##font[color]=" << color << ",";
	}
	style.setFont(font);
	auto val=int(libwps::read16(m_input));
	if (val!=-1) f << "f1=" << val << ",";
	librevenge::RVNGString name("");
	if (!m_mainParser.readPString(name,15))
		f << "##sSz,";
	else if (!name.empty())
		f << name.cstr() << ",";
	m_input->seek(pos+32, librevenge::RVNG_SEEK_SET);
	val=int(libwps::readU16(m_input)); // probably list of modified flags
	if (val)
		f << "wh=" << std::hex << val << std::dec << ",";
	for (int i=0; i<4; ++i)
	{
		val=int(libwps::readU8(m_input));
		if (!val) continue;
		WPSBorder border;
		switch (val)
		{
		case 1: // normal
			break;
		case 2:
			border.m_type = WPSBorder::Double;
			break;
		case 3:
			border.m_width=2;
			break;
		default:
			f << "#border" << i << "=" << val << ",";
			break;
		}
		int const wh[]= {WPSBorder::TopBit, WPSBorder::LeftBit, WPSBorder::BottomBit, WPSBorder::RightBit};
		style.setBorders(wh[i], border);
	}
	val=int(libwps::readU8(m_input));
	switch (val)
	{
	case 0:
		break;
	case 1:
		style.setBackgroundColor(WPSColor(0x80,0x80,0x80));
		break;
	case 2:
		style.setBackgroundColor(WPSColor::black());
		break;
	default:
		f << "#back[color]=" << val << ",";
		break;
	}
	val=int(libwps::readU8(m_input));
	switch (val)
	{
	case 0:
		break;
	case 1:
		style.setHAlignment(WPSCellFormat::HALIGN_LEFT);
		break;
	case 2:
		style.setHAlignment(WPSCellFormat::HALIGN_RIGHT);
		break;
	case 3:
		style.setHAlignment(WPSCellFormat::HALIGN_CENTER);
		break;
	default:
		f << "#align=" << val << ",";
		break;
	}
	val=int(libwps::readU8(m_input));
	switch (val)
	{
	case 0:
		break;
	case 1:
		f << "input=labels[only],";
		break;
	case 2:
		f << "input=dates[only],";
		break;
	default:
		f << "#input=" << val << ",";
		break;
	}
	style.m_fileFormat=int(libwps::readU8(m_input));
	for (int i=0; i<3; ++i)
	{
		val=int(libwps::read8(m_input));
		if (val) f << "f" << i+2 << "=" << val << ",";
	}

	style.m_extra=f.str();
	m_state->m_styleManager.add(id, style);
	f.str("");
	f << "Entries(UserStyle):Ce" << id << "," << style << ",";
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

////////////////////////////////////////////////////////////
// general
////////////////////////////////////////////////////////////

bool QuattroDosSpreadsheet::readCell()
{
	libwps::DebugStream f;

	long pos = m_input->tell();
	long type = libwps::read16(m_input);
	if (type < 0xc || type > 0x10)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readCell: not a cell property\n"));
		return false;
	}
	long sz = libwps::readU16(m_input);
	long endPos = pos+4+sz;

	if (sz < 5)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readCell: cell def is too short\n"));
		return false;
	}
	int const vers=version();
	auto defFontType=m_mainParser.getDefaultFontType();
	bool dosFile = vers<=1;
	int format=0xFF;
	if (dosFile)
		format = int(libwps::readU8(m_input));
	int cellPos[2];
	cellPos[0]=int(libwps::readU8(m_input));
	auto sheetId=int(libwps::readU8(m_input));
	cellPos[1]=int(libwps::read16(m_input));
	if (cellPos[1] < 0)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readCell: cell pos is bad\n"));
		return false;
	}
	if (sheetId)
	{
		if (vers==1)
		{
			WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readCell: find unexpected sheet id\n"));
			f << "###";
		}
		f << "sheet[id]=" << sheetId << ",";
	}

	auto &cell=m_state->getActualSheet().getCell(Vec2i(cellPos[0],cellPos[1]), defFontType);
	cell.m_fileFormat=format;
	if (!dosFile)
	{
		auto id=int(libwps::readU16(m_input));
		if (id!=0xFF0F)
		{
			f << "Ce" << (id>>8);
			if (id&0xFF) f << "-" << (id&0xFF);
			f << ",";
		}
		QuattroDosSpreadsheetInternal::Style style(defFontType);
		if (!m_state->m_styleManager.get(id, style))
			f << "###";
		else
		{
			cell.m_fileFormat=style.m_fileFormat;
			cell.m_fontType=style.m_fontType;
			cell.setFont(style.getFont());
			cell.setBackgroundColor(style.backgroundColor());
			if (style.hasBorders())
				cell.setBorders(style.borders());
		}
	}

	long dataPos = m_input->tell();
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
			cell.m_content.setValue(libwps::read16(m_input));
			break;
		}
		ok = false;
		break;
	}
	case 14:
	{
		double val;
		bool isNaN;
		if (dataSz == 8 && libwps::readDouble8(m_input, val, isNaN))
		{
			cell.m_content.m_contentType=WKSContentListener::CellContent::C_NUMBER;
			cell.m_content.setValue(val);
			break;
		}
		ok = false;
		break;
	}
	case 15:
	{
		cell.m_content.m_contentType=WKSContentListener::CellContent::C_TEXT;
		long begText=m_input->tell()+2, endText=begText+dataSz-2;
		std::string s("");
		// pascal string
		auto align=char(libwps::readU8(m_input));
		if (align=='\'') cell.setHAlignment(WPSCellFormat::HALIGN_DEFAULT);
		else if (align=='\\') cell.setHAlignment(WPSCellFormat::HALIGN_LEFT);
		else if (align=='^') cell.setHAlignment(WPSCellFormat::HALIGN_CENTER);
		else if (align=='\"') cell.setHAlignment(WPSCellFormat::HALIGN_RIGHT);
		else f << "#align=" << int(align) << ",";

		librevenge::RVNGString text("");
		if (!m_mainParser.readPString(text,dataSz-2))
			f << "##sSz,";
		else
		{
			if (endText!=m_input->tell())
			{
				endText=m_input->tell();
				f << "#extra,";
				ascii().addDelimiter(m_input->tell(), '|');
			}
			if (!text.empty())
				f << text.cstr() << ",";
		}
		cell.m_content.m_textEntry.setBegin(begText);
		cell.m_content.m_textEntry.setEnd(endText);
		break;
	}
	case 16:
	{
		double val;
		bool isNaN;
		if (dataSz >= 8 && libwps::readDouble8(m_input, val, isNaN))
		{
			cell.m_content.m_contentType=WKSContentListener::CellContent::C_FORMULA;
			cell.m_content.setValue(val);
			std::string error;
			if (!readFormula(endPos, cell.position(), sheetId, cell.m_content.m_formula, error))
			{
				cell.m_content.m_contentType=WKSContentListener::CellContent::C_NUMBER;
				ascii().addDelimiter(m_input->tell()-1, '#');
			}
			if (error.length()) f << error;
			break;
		}
		ok = false;
		break;
	}
	default:
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readCell: unknown type=%ld\n", type));
		ok = false;
		break;
	}
	if (!ok) ascii().addDelimiter(dataPos, '#');

	if (dosFile)
	{
		switch (cell.m_content.m_contentType)
		{
		case WKSContentListener::CellContent::C_NONE:
			break;
		case WKSContentListener::CellContent::C_TEXT:
			cell.setFormat(WPSCell::F_TEXT);
			break;
		case WKSContentListener::CellContent::C_NUMBER:
		case WKSContentListener::CellContent::C_FORMULA:
		case WKSContentListener::CellContent::C_UNKNOWN:
		default:
			cell.setFormat(WPSCell::F_NUMBER);
			break;
		}

	}
	m_input->seek(pos+sz, librevenge::RVNG_SEEK_SET);

	std::string extra=f.str();
	f.str("");
	f << "Entries(CellContent):" << cell << "," << extra;

	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	return true;
}

bool QuattroDosSpreadsheet::readCellFormulaResult()
{
	libwps::DebugStream f;

	long pos = m_input->tell();
	long type = libwps::read16(m_input);
	if (type != 0x33)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readCellFormulaResult: not a cell property\n"));
		return false;
	}
	long sz = libwps::readU16(m_input);
	if (sz<6)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readCellFormulaResult: the zone seems to short\n"));
		return false;
	}
	long endPos = pos+4+sz;

	bool dosFile = version() <= 1;
	// skip format for dosFile
	m_input->seek(dosFile ? pos+5 : pos+4, librevenge::RVNG_SEEK_SET);
	f << "CellContent[res]:";
	int dim[2];
	for (int &i : dim) i=int(libwps::readU16(m_input));
	f << "C" << dim[0] << "x" << dim[1] << ",";
	// skip format for windows file
	if (!dosFile) m_input->seek(2, librevenge::RVNG_SEEK_CUR);
	librevenge::RVNGString text("");
	if (!m_mainParser.readPString(text,endPos-m_input->tell()-1))
		f << "##sSz,";
	else if (!text.empty())
		f << text.cstr() << ",";
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

////////////////////////////////////////////////////////////
// Data
////////////////////////////////////////////////////////////
bool QuattroDosSpreadsheet::readSpreadsheetName()
{
	libwps::DebugStream f;

	long pos = m_input->tell();
	auto type = long(libwps::readU16(m_input));
	if (type != 0xde)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readSpreadsheetName: not a spreadsheet header\n"));
		return false;
	}
	auto sz = int(libwps::readU16(m_input));
	f << "Entries(SheetName):";
	librevenge::RVNGString name("");
	if (!m_mainParser.readPString(name,sz-1))
		f << "##sSz,";
	else if (!name.empty())
		f << name.cstr() << ",";
	/* TODO: store the name and use it to define the spreadsheet name

	   note: this will imply that we reconstruct the formula to set
	   the final name before sending them to the listener.
	 */
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

bool QuattroDosSpreadsheet::readSpreadsheetOpen()
{
	libwps::DebugStream f;

	long pos = m_input->tell();
	auto type = long(libwps::readU16(m_input));
	if (type != 0xdc)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readSpreadsheetOpen: not a spreadsheet header\n"));
		return false;
	}
	auto sz = long(libwps::readU16(m_input));
	f << "Entries(Spreadsheet)[beg]:";
	if (sz!=2)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readSpreadsheetOpen: the field size seems odd\n"));
		f << "###";
	}
	auto id=int(libwps::readU16(m_input));
	if (id<0||id>255)
	{
		f << "###";
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readSpreadsheetOpen: the id seems odd\n"));
	}
	else if (id==0 && m_state->m_spreadsheetStack.size()!=1)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readSpreadsheetOpen: find 0 but the stack is not empty\n"));
		f << "###stack";
		m_state->m_spreadsheetStack.push(m_state->m_spreadsheetList[0]);
	}
	else if (id)
		m_state->pushNewSheet(std::shared_ptr<QuattroDosSpreadsheetInternal::Spreadsheet>
		                      (new QuattroDosSpreadsheetInternal::Spreadsheet
		                       (QuattroDosSpreadsheetInternal::Spreadsheet::T_Spreadsheet, id)));
	f << id << ",";
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

bool QuattroDosSpreadsheet::readSpreadsheetClose()
{
	libwps::DebugStream f;

	long pos = m_input->tell();
	auto type = long(libwps::readU16(m_input));
	if (type != 0xdd)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readSpreadsheetClose: not a spreadsheet header\n"));
		return false;
	}
	auto sz = long(libwps::readU16(m_input));
	f << "Entries(Spreadsheet)[end]:";
	auto const &sheet=m_state->getActualSheet();
	if (sheet.m_type!=QuattroDosSpreadsheetInternal::Spreadsheet::T_Spreadsheet)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readSpreadsheetClose: can not find spreadsheet spreadsheet\n"));
		f << "###[noOpen],";
	}
	else if (m_state->m_spreadsheetStack.size() > 1)
		m_state->popSheet();
	if (sz!=0)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readSpreadsheetClose: the field size seems odd\n"));
		f << "###";
	}
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

bool QuattroDosSpreadsheet::readCell
(Vec2i actPos, WKSContentListener::FormulaInstruction &instr, bool hasSheetId, int sheetId)
{
	instr=WKSContentListener::FormulaInstruction();
	instr.m_type=WKSContentListener::FormulaInstruction::F_Cell;
	bool ok = true;
	int pos[3]= {0,0,0};
	bool absolute[3] = { true, true, true};
	for (int dim = 0; dim < (hasSheetId ? 3 : 2); dim++)
	{
		auto val = int(libwps::readU16(m_input));
		if (dim==0 && (val&0xF00) && (val&0xF000)!=0xF000)
		{
			// checkme: probably (val>>8)&0x1f
			instr.m_fileName=m_mainParser.getFileName(int((val>>8)&0xF));
			val &= 0xF0FF;
		}
		if ((val & 0xF000) == 0); // absolute value ?
		else if ((val & 0xc000) == 0x8000)   // relative ?
		{
			if (version()==1 || dim==0)
			{
				val &= 0xFF;
				if ((val & 0x80) && val+actPos[dim] >= 0x100)
					// sometimes this value is odd, so do not generate errors here
					val = val - 0x100;
			}
			else
			{
				val &= 0x3FFF;
				if (val>0x1000) val = val - 0x2000;
			}
			if (dim==2)
				val += sheetId;
			else
				val += actPos[dim];
			absolute[dim] = false;
		}
		else if (val==0xFFFF)
		{
			static bool first=true;
			if (first)   // in general associated with a nan value, so maybe be normal
			{
				WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readCell: find some ffff cell\n"));
				first=false;
			}
			ok = false;
		}
		else
		{
			WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readCell: can not read cell %x\n", unsigned(val)));
			ok = false;
		}
		pos[dim] = val;
	}

	if (pos[0] < 0 || pos[1] < 0)
	{
		if (ok)
		{
			WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readCell: can not read cell position\n"));
		}
		return false;
	}
	if (pos[0]>255) // FIXME:  can happens in some list cell, so...
		pos[0]&=0xFF;
	instr.m_position[0]=Vec2i(pos[0],pos[1]);
	instr.m_positionRelative[0]=Vec2b(!absolute[0],!absolute[1]);
	if ((hasSheetId && pos[2]!=sheetId) || !instr.m_fileName.empty())
		instr.m_sheetName[0]=m_state->getSheetName(pos[2]);
	return ok;
}

namespace QuattroDosSpreadsheetInternal
{
struct Functions
{
	char const *m_name;
	int m_arity;
};

static Functions const s_listFunctions[] =
{
	{ "", 0} /*SPEC: number*/, {"", 0}/*SPEC: cell*/, {"", 0}/*SPEC: cells*/, {"=", 1} /*SPEC: end of formula*/,
	{ "(", 1} /* SPEC: () */, {"", 0}/*SPEC: number*/, { "", -2} /*SPEC: text*/, {"", -2}/*unused*/,
	{ "-", 1}, {"+", 2}, {"-", 2}, {"*", 2},
	{ "/", 2}, { "^", 2}, {"=", 2}, {"<>", 2},

	{ "<=", 2},{ ">=", 2},{ "<", 2},{ ">", 2},
	{ "And", 2},{ "Or", 2}, { "Not", 1}, { "+", 1},
	{ "&", 2}, { "", -2} /*unused*/,{ "", -2} /*unused*/,{ "", -2} /*unused*/,
	{ "", -2} /*unused*/,{ "", -2} /*unused*/,{ "", -2} /*unused*/,{ "NA", 0} /*checkme*/,

	{ "NA", 0} /* Error*/,{ "Abs", 1},{ "Int", 1},{ "Sqrt", 1},
	{ "Log10", 1},{ "Ln", 1},{ "Pi", 0},{ "Sin", 1},
	{ "Cos", 1},{ "Tan", 1},{ "Atan2", 2},{ "Atan", 1},
	{ "Asin", 1},{ "Acos", 1},{ "Exp", 1},{ "Mod", 2},

	{ "Choose", -1},{ "IsNa", 1},{ "IsError", 1},{ "False", 0},
	{ "True", 0},{ "Rand", 0},{ "Date", 3},{ "Now", 0},
	{ "PMT", 3} /*BAD*/,{ "PV", 3} /*BAD*/,{ "FV", 3} /*BAD*/,{ "IF", 3},
	{ "Day", 1},{ "Month", 1},{ "Year", 1},{ "Round", 2},

	{ "Time", 3},{ "Hour", 1},{ "Minute", 1},{ "Second", 1},
	{ "IsNumber", 1},{ "IsText", 1},{ "Len", 1},{ "Value", 1},
	{ "Fixed", 2}, { "Mid", 3}, { "Char", 1},{ "Ascii", 1},
	{ "Find", 3},{ "DateValue", 1} /*checkme*/,{ "TimeValue", 1} /*checkme*/,{ "CellPointer", 1} /*checkme*/,

	{ "Sum", -1},{ "Average", -1},{ "COUNT", -1},{ "Min", -1},
	{ "Max", -1},{ "VLookUp", 3},{ "NPV", 2}, { "Var", -1},
	{ "StDev", -1},{ "IRR", 2} /*BAD*/, { "HLookup", 3},{ "DSum", 3},
	{ "DAvg", 3},{ "DCount", 3},{ "DMin", 3},{ "DMax", 3},

	{ "DVar", 3},{ "DStd", 3},{ "Index", 3}, { "Columns", 1},
	{ "Rows", 1},{ "Rept", 2},{ "Upper", 1},{ "Lower", 1},
	{ "Left", 2},{ "Right", 2},{ "Replace", 4}, { "Proper", 1},
	{ "Cell", 1} /*checkme*/,{ "Trim", 1},{ "Clean", 1} /*UNKN*/,{ "T", 1},

	{ "IsNonText", 1},{ "Exact", 2},{ "", -2} /*UNKN*/,{ "Indirect", 1},
	{ "RRI", 3}, { "TERM", 3}, { "CTERM", 3}, { "SLN", 3},
	{ "SYD", 4},{ "DDB", 4}, { "StDevP", -1}, { "VarP", -1},
	{ "DBStdDevP", -1}, { "DBVarP", -1}, { "PV", 3}, { "PMT", 5},

	{ "FV", 3}, { "Nper", 5}, { "Rate", 5}, { "Ipmt", 4},
	{ "Ppmt", 6}, { "SumProduct", 2}, { "", -2} /*UNKN*/,{ "", -2} /*UNKN*/,
	{ "", -2} /*UNKN*/,{ "", -2} /*UNKN*/,{ "", -2} /*UNKN*/,{ "", -2} /*UNKN*/,
	{ "", -2} /*UNKN*/,{ "And", -1},{ "Or", -1},{ "Not", 1},

};
}

bool QuattroDosSpreadsheet::readFormula(long endPos, Vec2i const &position, int sheetId,
                                        std::vector<WKSContentListener::FormulaInstruction> &formula, std::string &error)
{
	int const vers=version();
	formula.resize(0);
	error = "";
	long pos = m_input->tell();
	if (endPos - pos < (vers==1 ? 6 : 4)) return false;
	auto sz = int(libwps::readU16(m_input));
	if (endPos-pos-2 != sz) return false;

	std::vector<WKSContentListener::FormulaInstruction> listCellsPos;
	int fieldPos[2]= {0,sz};
	for (int &i : fieldPos)
	{
		i= int(libwps::readU16(m_input));
		if (vers>1) break;
	}
	if (fieldPos[0]<0||fieldPos[0]>fieldPos[1]||fieldPos[1]>sz)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readFormula: can not find the field header\n"));
		error="###fieldPos";
		return false;
	}
	int NSingle=0;
	if (fieldPos[0]!=sz)
	{
		m_input->seek(pos+2+fieldPos[0], librevenge::RVNG_SEEK_SET);
		ascii().addDelimiter(pos+2+fieldPos[0],'|');
		if (vers==1)
		{
			int N=int(sz-fieldPos[0])/4;
			NSingle=int(fieldPos[1]-fieldPos[0])/4;
			for (int i=0; i<N; ++i)
			{
				WKSContentListener::FormulaInstruction cell;
				if (!readCell(position, cell))
				{
					ascii().addDelimiter(pos+2+fieldPos[0]+i*4,'@');
					WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readFormula: can not read some cell\n"));
					error="###cell,";
					break;
				}
				listCellsPos.push_back(cell);
			}
		}
		else
		{
			while (!m_input->isEnd())
			{
				long actPos=m_input->tell();
				if (actPos+8>endPos) break;
				auto type=int(libwps::readU16(m_input));
				WKSContentListener::FormulaInstruction cell, cell2;
				if (type==0)
				{
					if (!readCell(position, cell, true, sheetId))
					{
						m_input->seek(actPos, librevenge::RVNG_SEEK_SET);
						break;
					}
					listCellsPos.push_back(cell);
					continue;
				}
				if (type!=0x1000 || actPos+14>endPos || !readCell(position, cell, true, sheetId) ||
				        !readCell(position, cell2, true, sheetId))
				{
					m_input->seek(actPos, librevenge::RVNG_SEEK_SET);
					break;
				}
				cell.m_type=WKSContentListener::FormulaInstruction::F_CellList;
				cell.m_position[1]=cell2.m_position[0];
				cell.m_positionRelative[1]=cell2.m_positionRelative[0];
				cell.m_sheetName[1]=cell2.m_sheetName[0];
				listCellsPos.push_back(cell);
			}
			if (m_input->tell() !=endPos)
			{
				ascii().addDelimiter(m_input->tell(),'@');
				WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readFormula: can not read some cell\n"));
				error="###cell,";
			}
		}
		m_input->seek(pos+4+(vers==1 ? 2 : 0), librevenge::RVNG_SEEK_SET);
		endPos=pos+2+fieldPos[0];
	}
	std::stringstream f;
	std::vector<std::vector<WKSContentListener::FormulaInstruction> > stack;
	bool ok = true;
	size_t actCellId=0, actDualCellId=size_t(NSingle);
	while (long(m_input->tell()) != endPos)
	{
		double val;
		bool isNaN;
		pos = m_input->tell();
		if (pos > endPos) return false;
		auto wh = int(libwps::readU8(m_input));
		int arity = 0;
		WKSContentListener::FormulaInstruction instr;
		switch (wh)
		{
		case 0x0:
			if (endPos-pos<9 || !libwps::readDouble8(m_input, val, isNaN))
			{
				f.str("");
				f << "###number";
				error=f.str();
				ok = false;
				break;
			}
			instr.m_type=WKSContentListener::FormulaInstruction::F_Double;
			instr.m_doubleValue=val;
			break;
		case 0x1:
			if (actCellId>=listCellsPos.size())
			{
				f.str("");
				f << "###unknCell" << actCellId;
				error=f.str();
				ok = false;
				break;
			}
			instr=listCellsPos[actCellId++];
			break;
		case 0x2:
			if (vers>=2)
			{
				if (actCellId>=listCellsPos.size())
				{
					f.str("");
					f << "###unknListCell" << actCellId;
					error=f.str();
					ok = false;
					break;
				}
				instr=listCellsPos[actCellId++];
				break;
			}
			if (actDualCellId+1>=listCellsPos.size())
			{
				f.str("");
				f << "###unknListCell" << actDualCellId;
				error=f.str();
				ok = false;
				break;
			}
			instr=listCellsPos[actDualCellId++];
			instr.m_type=WKSContentListener::FormulaInstruction::F_CellList;
			instr.m_position[1]=listCellsPos[actDualCellId].m_position[0];
			instr.m_positionRelative[1]=listCellsPos[actDualCellId].m_positionRelative[0];
			++actDualCellId;
			break;
		case 0x5:
			instr.m_type=WKSContentListener::FormulaInstruction::F_Long;
			instr.m_longValue=long(libwps::read16(m_input));
			break;
		case 0x6:
			instr.m_type=WKSContentListener::FormulaInstruction::F_Text;
			while (!m_input->isEnd())
			{
				if (m_input->tell() >= endPos)
				{
					ok=false;
					break;
				}
				auto c = char(libwps::readU8(m_input));
				if (c==0) break;
				instr.m_content += c;
			}
			break;
		default:
			if (wh >= 0x90 || QuattroDosSpreadsheetInternal::s_listFunctions[wh].m_arity == -2)
			{
				f.str("");
				f << "##Funct" << std::hex << wh;
				error=f.str();
				ok = false;
				break;
			}
			instr.m_type=WKSContentListener::FormulaInstruction::F_Function;
			instr.m_content=QuattroDosSpreadsheetInternal::s_listFunctions[wh].m_name;
			ok=!instr.m_content.empty();
			arity = QuattroDosSpreadsheetInternal::s_listFunctions[wh].m_arity;
			if (arity == -1) arity = int(libwps::read8(m_input));
			break;
		}

		if (!ok) break;
		std::vector<WKSContentListener::FormulaInstruction> child;
		if (instr.m_type!=WKSContentListener::FormulaInstruction::F_Function)
		{
			child.push_back(instr);
			stack.push_back(child);
			continue;
		}
		size_t numElt = stack.size();
		if (int(numElt) < arity)
		{
			f.str("");
			f << instr.m_content << "[##" << arity << "]";
			error=f.str();
			ok = false;
			break;
		}
		//
		// first treat the special cases
		//
		if (arity==3 && instr.m_type==WKSContentListener::FormulaInstruction::F_Function && instr.m_content=="TERM")
		{
			// @TERM(pmt,pint,fv) -> NPER(pint,-pmt,pv=0,fv)
			auto pmt=stack[size_t(int(numElt)-3)];
			auto pint=stack[size_t(int(numElt)-2)];
			auto fv=stack[size_t(int(numElt)-1)];

			stack.resize(size_t(++numElt));
			// pint
			stack[size_t(int(numElt)-4)]=pint;
			//-pmt
			auto &node=stack[size_t(int(numElt)-3)];
			instr.m_type=WKSContentListener::FormulaInstruction::F_Operator;
			instr.m_content="-";
			node.resize(0);
			node.push_back(instr);
			instr.m_content="(";
			node.push_back(instr);
			node.insert(node.end(), pmt.begin(), pmt.end());
			instr.m_content=")";
			node.push_back(instr);
			//pv=zero
			instr.m_type=WKSContentListener::FormulaInstruction::F_Long;
			instr.m_longValue=0;
			stack[size_t(int(numElt)-2)].resize(0);
			stack[size_t(int(numElt)-2)].push_back(instr);
			//fv
			stack[size_t(int(numElt)-1)]=fv;
			arity=4;
			instr.m_type=WKSContentListener::FormulaInstruction::F_Function;
			instr.m_content="NPER";
		}
		else if (arity==3 && instr.m_type==WKSContentListener::FormulaInstruction::F_Function && instr.m_content=="CTERM")
		{
			// @CTERM(pint,fv,pv) -> NPER(pint,pmt=0,-pv,fv)
			auto pint=stack[size_t(int(numElt)-3)];
			auto fv=stack[size_t(int(numElt)-2)];
			auto pv=stack[size_t(int(numElt)-1)];
			stack.resize(size_t(++numElt));
			// pint
			stack[size_t(int(numElt)-4)]=pint;
			// pmt=0
			instr.m_type=WKSContentListener::FormulaInstruction::F_Long;
			instr.m_longValue=0;
			stack[size_t(int(numElt)-3)].resize(0);
			stack[size_t(int(numElt)-3)].push_back(instr);
			// -pv
			auto &node=stack[size_t(int(numElt)-2)];
			instr.m_type=WKSContentListener::FormulaInstruction::F_Operator;
			instr.m_content="-";
			node.resize(0);
			node.push_back(instr);
			instr.m_content="(";
			node.push_back(instr);
			node.insert(node.end(), pv.begin(), pv.end());
			instr.m_content=")";
			node.push_back(instr);

			//fv
			stack[size_t(int(numElt)-1)]=fv;
			arity=4;
			instr.m_type=WKSContentListener::FormulaInstruction::F_Function;
			instr.m_content="NPER";
		}

		if ((instr.m_content[0] >= 'A' && instr.m_content[0] <= 'Z') || instr.m_content[0] == '(')
		{
			if (instr.m_content[0] != '(')
				child.push_back(instr);

			instr.m_type=WKSContentListener::FormulaInstruction::F_Operator;
			instr.m_content="(";
			child.push_back(instr);
			for (int i = 0; i < arity; i++)
			{
				if (i)
				{
					instr.m_content=";";
					child.push_back(instr);
				}
				auto const &node=stack[size_t(int(numElt)-arity+i)];
				child.insert(child.end(), node.begin(), node.end());
			}
			instr.m_content=")";
			child.push_back(instr);

			stack.resize(size_t(int(numElt)-arity+1));
			stack[size_t(int(numElt)-arity)] = child;
			continue;
		}
		if (arity==1)
		{
			instr.m_type=WKSContentListener::FormulaInstruction::F_Operator;
			stack[numElt-1].insert(stack[numElt-1].begin(), instr);
			if (wh==3)
				break;
			continue;
		}
		if (arity==2)
		{
			instr.m_type=WKSContentListener::FormulaInstruction::F_Operator;
			stack[numElt-2].push_back(instr);
			stack[numElt-2].insert(stack[numElt-2].end(), stack[numElt-1].begin(), stack[numElt-1].end());
			stack.resize(numElt-1);
			continue;
		}
		ok=false;
		error = "### unexpected arity";
		break;
	}

	if (!ok) ;
	else if (stack.size()==1 && stack[0].size()>1 && stack[0][0].m_content=="=")
	{
		formula.insert(formula.begin(),stack[0].begin()+1,stack[0].end());
		if (m_input->tell()!=endPos)
		{
			WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readFormula: find some extra data\n"));
			error="##extra data";
			ascii().addDelimiter(m_input->tell(),'#');
		}
		return true;
	}
	else
		error = "###stack problem";

	static bool first = true;
	if (first)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::readFormula: I can not read some formula\n"));
		first = false;
	}

	f.str("");
	for (auto const &i : stack)
	{
		for (auto const &j : i)
			f << j << ",";
	}
	f << error << "###";
	error = f.str();
	return false;
}

////////////////////////////////////////////////////////////
// send data
////////////////////////////////////////////////////////////
void QuattroDosSpreadsheet::sendSpreadsheet(int sId, std::map<Vec2i,Vec2i> const &chartPosMap)
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::sendSpreadsheet: I can not find the listener\n"));
		return;
	}
	auto sheet = m_state->getSheet(QuattroDosSpreadsheetInternal::Spreadsheet::T_Spreadsheet, sId);
	if (!sheet)
	{
		if (sId==0)
		{
			WPS_DEBUG_MSG(("QuattroDosSpreadsheet::sendSpreadsheet: oops can not find the actual sheet\n"));
		}
		sheet.reset(new QuattroDosSpreadsheetInternal::Spreadsheet);
	}
	sheet->insertChartPositions(chartPosMap, m_mainParser.getDefaultFontType());
	m_listener->openSheet(sheet->getWidths(), m_state->getSheetName(sId));
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
		auto cIt=chartPosMap.find(cell.position());
		if (cIt==chartPosMap.end())
			sendCellContent(cell);
		else
			sendCellContent(cell, sId, sheet->getZoneSize(cIt->first, cIt->second));
	}
	if (prevRow!=-1) m_listener->closeSheetRow();
	m_listener->closeSheet();
}

void QuattroDosSpreadsheet::sendCellContent(QuattroDosSpreadsheetInternal::Cell const &cell, int cellChartSheetId, Vec2f const &chartSize)
{
	if (m_listener.get() == nullptr)
	{
		WPS_DEBUG_MSG(("QuattroDosSpreadsheet::sendCellContent: I can not find the listener\n"));
		return;
	}

	libwps_tools_win::Font::Type fontType = cell.m_fontType;
	m_listener->setFont(cell.getFont());

	QuattroDosSpreadsheetInternal::Cell finalCell(cell);
	auto &content=finalCell.m_content;
	for (auto &f : content.m_formula)
	{
		if (f.m_type!=WKSContentListener::FormulaInstruction::F_Text)
			continue;
		std::string &text=f.m_content;
		auto finalText=libwps_tools_win::Font::unicodeString(text, fontType);
		if (finalText.empty())
			text.clear();
		else
			text=finalText.cstr();
	}
	finalCell.updateFormat();
	m_listener->openSheetCell(finalCell, content);

	if (cell.m_content.m_textEntry.valid())
	{
		m_input->seek(cell.m_content.m_textEntry.begin(), librevenge::RVNG_SEEK_SET);
		bool prevEOL=false;
		std::string text;
		while (m_input->tell()<=cell.m_content.m_textEntry.end())
		{
			bool last=m_input->isEnd() || m_input->tell()>=cell.m_content.m_textEntry.end();
			auto c=last ? '\0' : char(libwps::readU8(m_input));
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
					WPS_DEBUG_MSG(("QuattroDosSpreadsheet::sendCellContent: find 0xa without 0xd\n"));
				}
				prevEOL=false;
			}
			else
			{
				text.push_back(c);
				prevEOL=false;
			}
		}
	}
	if (cellChartSheetId>=0)
		m_mainParser.sendChart(cellChartSheetId, cell.position(), chartSize);
	m_listener->closeSheetCell();
}

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
