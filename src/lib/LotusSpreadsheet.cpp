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
#include <limits>
#include <map>
#include <regex>
#include <set>
#include <stack>
#include <sstream>
#include <utility>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WKSContentListener.h"
#include "WKSSubDocument.h"
#include "WPSEntry.h"
#include "WPSFont.h"
#include "WPSCell.h"
#include "WPSStream.h"
#include "WPSTable.h"

#include "Lotus.h"
#include "LotusStyleManager.h"

#include "LotusSpreadsheet.h"

static const int MAX_COLUMNS = 255;

namespace LotusSpreadsheetInternal
{

///////////////////////////////////////////////////////////////////
//! a class used to store a style of a cell in LotusSpreadsheet
struct Style final : public WPSCellFormat
{
	//! construtor
	explicit Style(libwps_tools_win::Font::Type type)
		: WPSCellFormat()
		, m_fontType(type)
		, m_extra("")
	{
		m_font.m_size=10;
	}
	Style(Style const &)=default;
	Style &operator=(Style const &)=default;
	//! destructor
	~Style() final;
	//! operator<<
	friend std::ostream &operator<<(std::ostream &o, Style const &style)
	{
		o << static_cast<WPSCellFormat const &>(style) << ",";
		if (!style.m_extra.empty()) o << style.m_extra;
		return o;
	}
	//! operator==
	bool operator==(Style const &st) const
	{
		if (m_fontType!=st.m_fontType || WPSCellFormat::compare(st)!=0) return false;
		return true;
	}
	//! operator!=
	bool operator!=(Style const &st) const
	{
		return !(*this==st);
	}
	//! font encoding type
	libwps_tools_win::Font::Type m_fontType;
	/** extra data */
	std::string m_extra;
};

Style::~Style()
{
}
//! the extra style
struct ExtraStyle
{
	//! constructor
	ExtraStyle()
		: m_color(WPSColor::black())
		, m_backColor(WPSColor::white())
		, m_format(0)
		, m_flag(0)
		, m_borders(0)
	{
	}
	//! returns true if the style is empty
	bool empty() const
	{
		// find also f[8-c]ffffffXX, which seems to have a different meaning
		if ((m_format&0xf0)==0xf0) return true;
		return m_color.isBlack() && m_backColor.isWhite() && (m_format&0x38)==0 && m_borders==0;
	}
	//! update the cell style
	void update(Style &style) const
	{
		WPSFont font=style.getFont();
		if (m_format&0x38)
		{
			if (m_format&0x8) font.m_attributes |= WPS_BOLD_BIT;
			if (m_format&0x10) font.m_attributes |= WPS_ITALICS_BIT;
			if (m_format&0x20) font.m_attributes |= WPS_UNDERLINE_BIT;
		}
		font.m_color=m_color;
		style.setFont(font);
		style.setBackgroundColor(m_backColor);
		if (m_borders)
		{
			for (int i=0,decal=0; i<4; ++i, decal+=2)
			{
				int type=(m_borders>>decal)&3;
				if (type==0) continue;
				static int const wh[]= {WPSBorder::LeftBit,WPSBorder::RightBit,WPSBorder::TopBit,WPSBorder::BottomBit};
				WPSBorder border;
				if (type==2) border.m_width=2;
				else if (type==3) border.m_type=WPSBorder::Double;
				style.setBorders(wh[i],border);
			}
		}
	}
	//! the font color
	WPSColor m_color;
	//! the backgroun color
	WPSColor m_backColor;
	//! the format
	int m_format;
	//! the second flag: graph
	int m_flag;
	//! the border
	int m_borders;
};

//! a class used to store the styles of a row in LotusSpreadsheet
struct RowStyles
{
	//! constructor
	RowStyles()
		: m_colsToStyleMap()
	{
	}
	//! a map Vec2i(minCol,maxCol) to style
	std::map<Vec2i, Style> m_colsToStyleMap;
};

//! a class used to store the extra style of a row in LotusSpreadsheet
struct ExtraRowStyles
{
	//! constructor
	ExtraRowStyles()
		: m_colsToStyleMap()
	{
	}

	//! returns true if all style are empty
	bool empty() const
	{
		for (auto it : m_colsToStyleMap)
		{
			if (!it.second.empty()) return false;
		}
		return true;
	}
	//! a map Vec2i(minCol,maxCol) to style
	std::map<Vec2i, ExtraStyle> m_colsToStyleMap;
};

//! a list of position of a Lotus spreadsheet
struct CellsList
{
	//! constructor
	CellsList()
		: m_positions()
	{
		for (auto &id : m_ids) id=0;
	}
	//! operator<<
	friend std::ostream &operator<<(std::ostream &o, CellsList const &pos)
	{
		o << pos.m_positions;
		for (int i=0; i<2; ++i)
		{
			if (pos.m_ids[i]) o << "[sheet" << i << "=" << pos.m_ids[i] << "]";
		}
		o << ",";
		return o;
	}
	//! the sheets id
	int m_ids[2];
	//! the first and last position
	WPSBox2i m_positions;
};

//! a cellule of a Lotus spreadsheet
class Cell final : public WPSCell
{
public:
	/// constructor
	Cell()
		: m_input()
		, m_styleId(-1)
		, m_hAlignement(WPSCellFormat::HALIGN_DEFAULT)
		, m_content()
		, m_comment() { }
	/// constructor
	explicit Cell(RVNGInputStreamPtr const &input)
		: m_input(input)
		, m_styleId(-1)
		, m_hAlignement(WPSCellFormat::HALIGN_DEFAULT)
		, m_content()
		, m_comment() { }
	//! operator<<
	friend std::ostream &operator<<(std::ostream &o, Cell const &cell);

	//! call when a cell must be send
	bool send(WPSListenerPtr &/*listener*/) final
	{
		WPS_DEBUG_MSG(("LotusSpreadsheetInternal::Cell::send: must not be called\n"));
		return false;
	}

	//! call when the content of a cell must be send
	bool sendContent(WPSListenerPtr &/*listener*/) final;

	//! the input
	RVNGInputStreamPtr m_input;
	//! the style
	int m_styleId;
	//! the horizontal align (in dos file)
	WPSCellFormat::HorizontalAlignment m_hAlignement;
	//! the content
	WKSContentListener::CellContent m_content;
	//! the comment entry
	WPSEntry m_comment;
};
bool Cell::sendContent(WPSListenerPtr &/*listener*/)
{
	WPS_DEBUG_MSG(("LotusSpreadsheetInternal::Cell::sendContent: must not be called\n"));
	return false;
}

//! operator<<
std::ostream &operator<<(std::ostream &o, Cell const &cell)
{
	o << reinterpret_cast<WPSCell const &>(cell) << cell.m_content << ",";
	if (cell.m_styleId>=0) o << "style=" << cell.m_styleId << ",";
	switch (cell.m_hAlignement)
	{
	case WPSCellFormat::HALIGN_LEFT:
		o << "left,";
		break;
	case WPSCellFormat::HALIGN_CENTER:
		o << "centered,";
		break;
	case WPSCellFormat::HALIGN_RIGHT:
		o << "right,";
		break;
	case WPSCellFormat::HALIGN_FULL:
		o << "full,";
		break;
	case WPSCellFormat::HALIGN_DEFAULT:
	default:
		break; // default
	}
	return o;
}

///////////////////////////////////////////////////////////////////
//! the spreadsheet of a LotusSpreadsheet
class Spreadsheet
{
public:
	//! a constructor
	Spreadsheet()
		: m_name("")
		, m_numCols(0)
		, m_numRows(0)
		, m_boundsColsMap()
		, m_widthCols()
		, m_rowHeightMap()
		, m_heightDefault(16)
		, m_rowPageBreaksList()
		, m_positionToCellMap()
		, m_rowToStyleIdMap()
		, m_rowToExtraStyleMap() {}
	//! return a cell corresponding to a spreadsheet, create one if needed
	Cell &getCell(RVNGInputStreamPtr input, Vec2i const &pos)
	{
		if (m_positionToCellMap.find(pos)==m_positionToCellMap.end())
		{
			Cell cell(input);
			cell.setPosition(pos);
			m_positionToCellMap[pos]=cell;
		}
		return m_positionToCellMap.find(pos)->second;
	}
	//! set the columns size
	void setColumnWidth(int col, WPSColumnFormat const &format)
	{
		if (col >= int(m_widthCols.size()))
		{
			// sanity check
			if (col>MAX_COLUMNS || (!m_boundsColsMap.empty() && col >= int(m_widthCols.size())+10 &&
			                        m_boundsColsMap.find(col)==m_boundsColsMap.end()))
			{
				WPS_DEBUG_MSG(("LotusSpreadsheetInternal::Spreadsheet::setColumnWidth: the column %d seems bad\n", col));
				return;
			}
			WPSColumnFormat defCol;
			defCol.m_useOptimalWidth=true;
			m_widthCols.resize(size_t(col)+1, defCol);
		}
		m_widthCols[size_t(col)] = format;
		if (col >= m_numCols) m_numCols=col+1;
	}
	//! returns the row size in point
	WPSRowFormat getRowHeight(int row) const
	{
		auto rIt=m_rowHeightMap.lower_bound(Vec2i(-1,row));
		if (rIt!=m_rowHeightMap.end() && rIt->first[0]<=row && rIt->first[1]>=row)
			return rIt->second;
		WPSRowFormat format(m_heightDefault);
		format.m_isMinimalHeight=true;
		return format;
	}
	//! set the rows size
	void setRowHeight(int row, WPSRowFormat const &format)
	{
		m_rowHeightMap[Vec2i(row,row)]=format;
	}
	//! returns the position corresponding to a cell
	Vec2f getPosition(Vec2i const &cell) const
	{
		// first compute the height
		int lastRow=0;
		float h=0;
		auto rIt=m_rowHeightMap.begin();
		while (rIt!=m_rowHeightMap.end() && rIt->first[1]<cell[1])
		{
			if (rIt->first[0]>lastRow)
			{
				h+=float(rIt->first[0]-lastRow)*m_heightDefault;
				lastRow=rIt->first[0];
			}
			float rHeight=rIt->second.m_height>=0 ? rIt->second.m_height : m_heightDefault;
			h+=float(rIt->first[1]+1-lastRow)*rHeight;
			lastRow=rIt->first[1]+1;
			++rIt;
		}
		if (lastRow<cell[1])
		{
			if (rIt!=m_rowHeightMap.end() && rIt->first[0]<cell[1] && rIt->second.m_height>=0)
				h+=float(cell[1]-lastRow)*rIt->second.m_height;
			else
				h+=float(cell[1]-lastRow)*m_heightDefault;

		}
		// now compute the width
		size_t const numCols = m_widthCols.size();
		float w=0;
		for (size_t i = 0; i < numCols && i<size_t(cell[0]); i++)
			w+=m_widthCols[i].m_width>=0 ? m_widthCols[i].m_width : 72;
		if (numCols<size_t(cell[0]))
			w+=72*float(size_t(cell[0])-numCols);
		return Vec2f(w, h);
	}
	//! try to compress the list of row height
	void compressRowHeights()
	{
		auto oldMap=m_rowHeightMap;
		m_rowHeightMap.clear();
		auto rIt=oldMap.begin();
		WPSRowFormat actHeight;
		WPSRowFormat defHeight(m_heightDefault);
		defHeight.m_isMinimalHeight=true;
		Vec2i actPos(0,-1);
		while (rIt!=oldMap.end())
		{
			// first check for not filled row
			if (rIt->first[0]!=actPos[1]+1)
			{
				if (actHeight==defHeight)
					actPos[1]=rIt->first[0]-1;
				else
				{
					if (actPos[1]>=actPos[0])
						m_rowHeightMap[actPos]=actHeight;
					actHeight=defHeight;
					actPos=Vec2i(actPos[1]+1, rIt->first[0]-1);
				}
			}
			if (rIt->second!=actHeight)
			{
				if (actPos[1]>=actPos[0])
					m_rowHeightMap[actPos]=actHeight;
				actPos[0]=rIt->first[0];
				actHeight=rIt->second;
			}
			actPos[1]=rIt->first[1];
			++rIt;
		}
		if (actPos[1]>=actPos[0])
			m_rowHeightMap[actPos]=actHeight;
	}
	//! convert the m_widthColsInChar in a vector of of point size
	std::vector<WPSColumnFormat> getWidths() const
	{
		std::vector<WPSColumnFormat> widths;
		WPSColumnFormat actWidth;
		int repeat=0;
		for (auto const &newWidth : m_widthCols)
		{
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
	//! returns the row style id corresponding to a sheetId (or -1)
	int getRowStyleId(int row) const
	{
		auto it=m_rowToStyleIdMap.lower_bound(Vec2i(-1, row));
		if (it!=m_rowToStyleIdMap.end() && it->first[0]<=row && row<=it->first[1])
			return int(it->second);
		return -1;
	}

	//! returns true if the spreedsheet is empty
	bool empty() const
	{
		return m_positionToCellMap.empty() && m_rowToStyleIdMap.empty() && m_name.empty();
	}
	/** the sheet name */
	librevenge::RVNGString m_name;
	/** the number of columns */
	int m_numCols;
	/** the number of rows */
	int m_numRows;
	/** a map used to stored the min/max row of each columns */
	std::map<int, Vec2i> m_boundsColsMap;
	/** the column size */
	std::vector<WPSColumnFormat> m_widthCols;
	/** the map Vec2i(min row, max row) to size in points */
	std::map<Vec2i,WPSRowFormat> m_rowHeightMap;
	/** the default row size in point */
	float m_heightDefault;
	/** the list of row page break */
	std::vector<int> m_rowPageBreaksList;
	/** a map cell to not empty cells */
	std::map<Vec2i, Cell> m_positionToCellMap;
	//! map Vec2i(min row, max row) to state row style id
	std::map<Vec2i,size_t> m_rowToStyleIdMap;
	//! map row to extra style
	std::map<int,ExtraRowStyles> m_rowToExtraStyleMap;
};

// ------------------------------------------------------------
// Lotus 123
// ------------------------------------------------------------
//! the format style for lotus 123
struct Format123Style final : public WPSCellFormat
{
	//! constructor
	Format123Style()
		: m_alignAcrossColumn(false)
	{
	}
	Format123Style(Format123Style const &)=default;
	Format123Style &operator=(Format123Style const &)=default;
	//! destructor
	~Format123Style() final;
	//! update the cell style
	void update(Style &style) const
	{
		style.setDTFormat(getFormat(), getDTFormat());
		style.setFormat(getFormat(), getSubFormat());
		style.setDigits(digits());
	}
	//! operator==
	bool operator==(Format123Style const &f) const
	{
		return m_alignAcrossColumn==f.m_alignAcrossColumn && compare(f)==0;
	}
	//! flag to know if we must align across column
	bool m_alignAcrossColumn;
};

Format123Style::~Format123Style()
{
}

//! the extra style for lotus 123
struct Extra123Style
{
	//! constructor
	Extra123Style()
	{
		for (auto &border : m_borders)
			border.m_style=WPSBorder::None;
	}
	//! returns true if the style is empty
	bool empty() const
	{
		for (const auto &border : m_borders)
		{
			if (!border.isEmpty()) return false;
		}
		return true;
	}
	//! operator==
	bool operator==(Extra123Style const &f) const
	{
		for (int i=0; i<2; ++i)
		{
			if (m_borders[i]!=f.m_borders[i])
				return false;
		}
		return true;
	}
	//! update the cell style
	void update(Style &style) const
	{
		for (int i=0; i<2; ++i)
		{
			if (m_borders[i].isEmpty()) continue;
			style.setBorders(i==0 ? WPSBorder::TopBit : WPSBorder::LeftBit,m_borders[i]);
		}
	}
	//! the top/left border
	WPSBorder m_borders[2];
};

//! a class used to store the styles of a table in LotusSpreadsheet in a lotus 123 files
struct Table123Styles
{
	//! constructor
	Table123Styles()
		: m_defaultCellId(-1)
		, m_rowsToColsToCellIdMap()
		, m_rowsToColsToExtraStyleMap()
		, m_rowsToColsToFormatStyleMap()
	{
	}
	//! add a style to a list of cell
	void addCellStyle(Vec2i const &cols, Vec2i const &rows, int cellId)
	{
		if (m_rowsToColsToCellIdMap.find(rows)==m_rowsToColsToCellIdMap.end())
			m_rowsToColsToCellIdMap[rows]= std::map<Vec2i,int>();
		auto &map=m_rowsToColsToCellIdMap.find(rows)->second;
		if (map.find(cols)!=map.end())
		{
			WPS_DEBUG_MSG(("LotusSpreadsheetInternal::Table123Styles::addCellStyle: find dupplicated cell\n"));
		}
		auto it=map.lower_bound(cols);
		if (!map.empty() && it!=map.begin()) --it;
		if (it!=map.end() && it->first[1]+1==cols[0] && it->second==cellId)
		{
			Vec2i newCols=it->first;
			map.erase(newCols);
			newCols[1]=cols[1];
			map[newCols]=cellId;
		}
		else
			map[cols]=cellId;
	}
	//! add a extra style to a list of cell
	void addCellStyle(Vec2i const &cols, Vec2i const &rows, Extra123Style const &extra)
	{
		if (m_rowsToColsToExtraStyleMap.find(rows)==m_rowsToColsToExtraStyleMap.end())
			m_rowsToColsToExtraStyleMap[rows]= std::map<Vec2i,Extra123Style>();
		auto &map=m_rowsToColsToExtraStyleMap.find(rows)->second;
		// checkme: sometimes, we can retrieve the same cells again
		auto it=map.lower_bound(cols);
		if (!map.empty() && it!=map.begin()) --it;
		if (it!=map.end() && it->first[1]+1==cols[0] && it->second==extra)
		{
			Vec2i newCols=it->first;
			map.erase(newCols);
			newCols[1]=cols[1];
			map[newCols]=extra;
		}
		else
			map[cols]=extra;
	}
	//! add a extra style to a list of cell
	void addCellStyle(Vec2i const &cols, Vec2i const &rows, Format123Style const &format)
	{
		if (m_rowsToColsToFormatStyleMap.find(rows)==m_rowsToColsToFormatStyleMap.end())
			m_rowsToColsToFormatStyleMap[rows]= std::map<Vec2i,Format123Style>();
		auto &map=m_rowsToColsToFormatStyleMap.find(rows)->second;
		if (map.find(cols)!=map.end())
		{
			WPS_DEBUG_MSG(("LotusSpreadsheetInternal::Table123Styles::addCellStyle: find dupplicated cell\n"));
		}
		auto it=map.lower_bound(cols);
		if (!map.empty() && it!=map.begin()) --it;
		if (it!=map.end() && it->first[1]+1==cols[0] && it->second==format)
		{
			Vec2i newCols=it->first;
			map.erase(newCols);
			newCols[1]=cols[1];
			map[newCols]=format;
		}
		else
			map[cols]=format;
	}
	//! the default cell style
	int m_defaultCellId;
	//! map rows to cols to cell id
	std::map<Vec2i, std::map<Vec2i,int> > m_rowsToColsToCellIdMap;
	//! map rows to cols to extra style
	std::map<Vec2i, std::map<Vec2i,Extra123Style> > m_rowsToColsToExtraStyleMap;
	//! map rows to cols to format style
	std::map<Vec2i, std::map<Vec2i,Format123Style> > m_rowsToColsToFormatStyleMap;
};

//! the state of LotusSpreadsheet
struct State
{
	//! constructor
	State()
		:  m_version(-1)
		, m_spreadsheetList()
		, m_nameToCellsMap()
		, m_rowStylesList()
		, m_rowSheetIdToStyleIdMap()
		, m_rowSheetIdToChildRowIdMap()
		, m_sheetIdToTableStyleMap()
		, m_sheetCurrentId(-1)
	{
		m_spreadsheetList.resize(1);
	}
	//! returns the number of spreadsheet
	int getNumSheet() const
	{
		return int(m_spreadsheetList.size());
	}
	//! returns the ith spreadsheet
	Spreadsheet &getSheet(int id)
	{
		if (id<0||id>=int(m_spreadsheetList.size()))
		{
			WPS_DEBUG_MSG(("LotusSpreadsheetInternal::State::getSheet: can not find spreadsheet %d\n", id));
			static Spreadsheet empty;
			return empty;
		}
		return m_spreadsheetList[size_t(id)];
	}
	//! returns a table style for a sheet(if it exists)
	Table123Styles const *getTableStyle(int id) const
	{
		Vec2i pos(-1,id);
		auto cIt=m_sheetIdToTableStyleMap.lower_bound(pos);
		if (cIt==m_sheetIdToTableStyleMap.end() || cIt->first[0]>id || cIt->first[1]<id)
			return nullptr;
		return &cIt->second;
	}
	//! returns a table style for a sheet zone, create it if needed
	Table123Styles *getTablesStyle(Vec2i pos)
	{
		auto cIt=m_sheetIdToTableStyleMap.lower_bound(pos);
		if (cIt==m_sheetIdToTableStyleMap.end() || cIt->first[0]>pos[1] || cIt->first[1]<pos[0])
		{
			m_sheetIdToTableStyleMap[pos]=Table123Styles();
			return &m_sheetIdToTableStyleMap.find(pos)->second;
		}
		if (cIt->first==pos)
			return &cIt->second;
		Vec2i actPos=cIt->first;
		if (actPos[0]>pos[0] || actPos[1]<pos[1])
		{
			WPS_DEBUG_MSG(("LotusSpreadsheetInternal::State::getTablesStyle: problem when creating spreadsheet %d,%d\n", pos[0], pos[1]));
			return nullptr;
		}
		auto const &table=cIt->second;
		if (actPos[0]<pos[0])
		{
			Vec2i newPos(actPos[0],pos[0]);
			m_sheetIdToTableStyleMap[newPos]=table;
		}
		m_sheetIdToTableStyleMap[pos]=table;
		if (actPos[1]>pos[1])
		{
			Vec2i newPos(pos[1],actPos[1]);
			m_sheetIdToTableStyleMap[newPos]=table;
		}
		m_sheetIdToTableStyleMap.erase(actPos);
		return &m_sheetIdToTableStyleMap.find(pos)->second;
	}
	//! returns the ith spreadsheet name
	librevenge::RVNGString getSheetName(int id) const
	{
		if (id>=0 && id<int(m_spreadsheetList.size()) && !m_spreadsheetList[size_t(id)].m_name.empty())
			return m_spreadsheetList[size_t(id)].m_name;
		librevenge::RVNGString name;
		name.sprintf("Sheet%d", id+1);
		return name;
	}
	//! the file version
	int m_version;
	//! the list of spreadsheet ( first: main spreadsheet, other report spreadsheet )
	std::vector<Spreadsheet> m_spreadsheetList;
	//! map name to position
	std::map<std::string, CellsList> m_nameToCellsMap;
	//! the list of row styles
	std::vector<RowStyles> m_rowStylesList;
	//! map Vec2i(row, sheetId) to row style id
	std::map<Vec2i,size_t> m_rowSheetIdToStyleIdMap;
	//! map Vec2i(row, sheetId) to child style
	std::multimap<Vec2i,Vec2i> m_rowSheetIdToChildRowIdMap;
	//! map Vec2i(sheetMin, sheetMax) to table style
	std::map<Vec2i,Table123Styles> m_sheetIdToTableStyleMap;
	//! the sheet id
	int m_sheetCurrentId;
};

//! Internal: the subdocument of a LotusSpreadsheet
class SubDocument final : public WKSSubDocument
{
public:
	//! constructor for a text entry
	SubDocument(RVNGInputStreamPtr const &input, LotusSpreadsheet &sheetParser, WPSEntry const &entry)
		: WKSSubDocument(input, nullptr)
		, m_sheetParser(&sheetParser)
		, m_entry(entry) {}
	//! destructor
	~SubDocument() final {}

	//! operator==
	bool operator==(std::shared_ptr<WPSSubDocument> const &doc) const final
	{
		if (!doc || !WKSSubDocument::operator==(doc))
			return false;
		auto const *sDoc = dynamic_cast<SubDocument const *>(doc.get());
		if (!sDoc) return false;
		if (m_sheetParser != sDoc->m_sheetParser) return false;
		if (m_entry != sDoc->m_entry) return false;
		return true;
	}

	//! the parser function
	void parse(std::shared_ptr<WKSContentListener> &listener, libwps::SubDocumentType subDocumentType) final;
	//! the spreadsheet parse
	LotusSpreadsheet *m_sheetParser;
	//! a text zone entry
	WPSEntry m_entry;
private:
	SubDocument(SubDocument const &) = delete;
	SubDocument &operator=(SubDocument const &) = delete;
};

void SubDocument::parse(std::shared_ptr<WKSContentListener> &listener, libwps::SubDocumentType)
{
	if (!m_sheetParser)
	{
		listener->insertCharacter(' ');
		WPS_DEBUG_MSG(("LotusSpreadsheetInternal::SubDocument::parse: bad parser\n"));
		return;
	}
	m_sheetParser->sendTextNote(m_input, m_entry);
}
}

// constructor, destructor
LotusSpreadsheet::LotusSpreadsheet(LotusParser &parser)
	: m_listener()
	, m_mainParser(parser)
	, m_styleManager(parser.m_styleManager)
	, m_state(new LotusSpreadsheetInternal::State)
{
}

LotusSpreadsheet::~LotusSpreadsheet()
{
}

void LotusSpreadsheet::cleanState()
{
	m_state.reset(new LotusSpreadsheetInternal::State);
}

bool LotusSpreadsheet::getLeftTopPosition(Vec2i const &cell, int sheetId, Vec2f &pos)
{
	// set to default
	pos=Vec2f(float(cell[0]>=0 ? cell[0]*72 : 0), float(cell[1]>=0 ? cell[1]*16 : 0));
	if (sheetId<0||sheetId>=m_state->getNumSheet() || cell[0]<0 || cell[1]<0)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::getLeftTopPosition: the sheet %d seems bad\n", sheetId));
		return true;
	}
	auto const &sheet = m_state->getSheet(sheetId);
	pos=sheet.getPosition(cell);
	return true;
}

librevenge::RVNGString LotusSpreadsheet::getSheetName(int id) const
{
	return m_state->getSheetName(id);
}

void LotusSpreadsheet::updateState()
{
	// update the state correspondance between row and row's styles
	if (!m_state->m_rowSheetIdToChildRowIdMap.empty())
	{
		std::set<Vec2i> seens;
		std::stack<Vec2i> toDo;
		for (auto it : m_state->m_rowSheetIdToStyleIdMap)
			toDo.push(it.first);
		while (!toDo.empty())
		{
			Vec2i pos=toDo.top();
			toDo.pop();
			if (seens.find(pos)!=seens.end())
			{
				WPS_DEBUG_MSG(("LotusSpreadsheet::updateState: dupplicated position, something is bad\n"));
				continue;
			}
			seens.insert(pos);
			auto cIt=m_state->m_rowSheetIdToChildRowIdMap.lower_bound(pos);
			if (cIt==m_state->m_rowSheetIdToChildRowIdMap.end() || cIt->first!=pos)
				continue;
			if (m_state->m_rowSheetIdToStyleIdMap.find(pos)==m_state->m_rowSheetIdToStyleIdMap.end())
			{
				WPS_DEBUG_MSG(("LotusSpreadsheet::updateState: something is bad\n"));
				continue;
			}
			size_t finalPos=m_state->m_rowSheetIdToStyleIdMap.find(pos)->second;
			while (cIt!=m_state->m_rowSheetIdToChildRowIdMap.end() && cIt->first==pos)
			{
				Vec2i const &cPos=cIt++->second;
				m_state->m_rowSheetIdToStyleIdMap[cPos]=finalPos;
				toDo.push(cPos);
			}
		}
	}

	// time to update each sheet rows style map
	for (auto it=m_state->m_rowSheetIdToStyleIdMap.begin(); it!=m_state->m_rowSheetIdToStyleIdMap.end();)
	{
		int sheetId=it->first[1];
		LotusSpreadsheetInternal::Spreadsheet *sheet=nullptr;
		if (sheetId>=0 && sheetId<int(m_state->m_spreadsheetList.size()))
			sheet=&m_state->m_spreadsheetList[size_t(sheetId)];
		else
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::updateState: can not find sheet %d\n", sheetId));
		}
		int lastStyleId=-1;
		Vec2i rows(0,-1);
		while (it!=m_state->m_rowSheetIdToStyleIdMap.end() && it->first[1]==sheetId)
		{
			if (lastStyleId!=int(it->second) || it->first[0]!=rows[1]+1)
			{
				if (lastStyleId>=0 && sheet)
					sheet->m_rowToStyleIdMap[rows]=size_t(lastStyleId);
				lastStyleId=int(it->second);
				rows=Vec2i(it->first[0], it->first[0]);
			}
			else
				++rows[1];
			++it;
		}
		if (lastStyleId>=0 && sheet)
			sheet->m_rowToStyleIdMap[rows]=size_t(lastStyleId);
	}
}

void LotusSpreadsheet::setLastSpreadsheetId(int id)
{
	if (id<0)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::setLastSpreadsheetId: the id:%d seems bad\n", id));
		return;
	}
	m_state->m_spreadsheetList.resize(size_t(id+1));
}

int LotusSpreadsheet::version() const
{
	if (m_state->m_version<0)
		m_state->m_version=m_mainParser.version();
	return m_state->m_version;
}

bool LotusSpreadsheet::hasSomeSpreadsheetData() const
{
	for (auto const &sheet : m_state->m_spreadsheetList)
	{
		if (!sheet.empty())
			return true;
	}
	return false;
}

////////////////////////////////////////////////////////////
// low level

////////////////////////////////////////////////////////////
//   parse sheet data
////////////////////////////////////////////////////////////
bool LotusSpreadsheet::readColumnDefinition(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::read16(input));
	if (type != 0x1f)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readColumnDefinition: not a column definition\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	f << "Entries(ColDef):";
	if (sz<8 || (sz%4))
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readColumnDefinition: the zone is too short\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto sheetId=int(libwps::readU8(input));
	f << "sheet[id]=" << sheetId << ",";
	auto col=int(libwps::readU8(input));
	f << "col=" << col << ",";
	auto N=int(libwps::readU8(input));
	if (N!=1) f << "N=" << N << ",";
	auto val=int(libwps::readU8(input)); // between 0 and 94
	if (val) f << "f0=" << val << ",";
	if (sz!=4+4*N)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readColumnDefinition: the number of columns seems bad\n"));
		f << "###N,";
		if (sz==8)
			N=1;
		else
		{
			ascFile.addPos(pos);
			ascFile.addNote(f.str().c_str());
			return true;
		}
	}
	Vec2i bound;
	for (int n=0; n<N; ++n)
	{
		int rowPos[2];
		for (int &i : rowPos) i=int(libwps::readU16(input));
		if (n==0)
			bound=Vec2i(rowPos[0], rowPos[1]);
		else
		{
			if (rowPos[0]<bound[0])
				bound[0]=rowPos[0];
			if (rowPos[1]>bound[1])
				bound[1]=rowPos[1];
		}
		f << "row" << n << "[bound]=" << Vec2i(rowPos[0], rowPos[1]) << ",";
	}
	if (sheetId<0||sheetId>=m_state->getNumSheet())
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readColumnDefinition the zone id seems bad\n"));
		f << "##id";
	}
	else
	{
		auto &sheet=m_state->getSheet(sheetId);
		if (sheet.m_boundsColsMap.find(col)!=sheet.m_boundsColsMap.end())
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readColumnDefinition the zone col seems bad\n"));
			f << "##col";
		}
		else
			sheet.m_boundsColsMap[col]=bound;
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusSpreadsheet::readColumnSizes(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::read16(input));
	if (type != 0x7)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readColumnSizes: not a column size name\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	f << "Entries(ColSize):";
	if (sz < 4 || (sz%2))
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readColumnSizes: the zone is too odd\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto sheetId=int(libwps::readU8(input));
	f << "id[sheet]=" << sheetId << ",";
	LotusSpreadsheetInternal::Spreadsheet empty, *sheet=nullptr;
	if (sheetId<0||sheetId>=int(m_state->m_spreadsheetList.size()))
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readColumnSizes: can find spreadsheet %d\n", sheetId));
		sheet=&empty;
		f << "###";
	}
	else
		sheet=&m_state->m_spreadsheetList[size_t(sheetId)];
	auto val=int(libwps::readU8(input)); // always 0?
	if (val) f << "f0=" << val << ",";
	f << "f1=" << std::hex << libwps::readU16(input) << std::dec << ","; // big number
	auto N=int((sz-4)/2);
	f << "widths=[";
	for (int i=0; i<N; ++i)
	{
		auto col=int(libwps::readU8(input));
		auto width=int(libwps::readU8(input)); // width in char, default 12...
		sheet->setColumnWidth(col, WPSColumnFormat(float(7*width)));
		f << width << "C:col" << col << ",";
	}
	f << "],";
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	return true;
}

bool LotusSpreadsheet::readRowFormats(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::read16(input));
	if (type != 0x13)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readRowFormats: not a row definition\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos = pos+4+sz;
	f << "Entries(RowFormat):";
	if (sz<8)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readRowFormats: the zone is too short\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto sheetId=int(libwps::readU8(input));
	auto rowType=int(libwps::readU8(input));
	auto row=int(libwps::readU16(input));
	int val;
	f << "sheet[id]=" << sheetId << ",";
	if (row) f << "row=" << row << ",";
	switch (rowType)
	{
	case 0:
	{
		f << "def,";
		size_t rowStyleId=m_state->m_rowStylesList.size();
		m_state->m_rowStylesList.resize(rowStyleId+1);

		int actCell=0;
		auto &stylesList=m_state->m_rowStylesList.back();
		f << "[";
		while (input->tell()<endPos)
		{
			int numCell;
			LotusSpreadsheetInternal::Style style(m_mainParser.getDefaultFontType());
			if (!readRowFormat(stream, style, numCell, endPos))
			{
				f << "###";
				WPS_DEBUG_MSG(("LotusSpreadsheet::readRowFormats: find extra data\n"));
				break;
			}
			if (numCell>=1)
				stylesList.m_colsToStyleMap.insert
				(std::map<Vec2i,LotusSpreadsheetInternal::Style>::value_type
				 (Vec2i(actCell,actCell+numCell-1),style));
			f << "[" << style << "]";
			if (numCell>1)
				f << "x" << numCell;
			f << ",";
			actCell+=numCell;
		}
		f << "],";
		m_state->m_rowSheetIdToStyleIdMap[Vec2i(row,sheetId)]=rowStyleId;
		if (actCell>256)
		{
			f << "###";
			WPS_DEBUG_MSG(("LotusSpreadsheet::readRowFormats: find too much cells\n"));
		}
		break;
	}
	case 1: // the last row definition, maybe the actual row style ?
		f << "last,";
		if (sz<12)
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readRowFormats: the size seems bad\n"));
			f << "###sz,";
			break;
		}
		for (int i=0; i<8; ++i)  // f0=0|32|41|71|7e|fe,f1=0|1|40|50|c0, f2=0|4|5|b|41, f3=0|40|54|5c, f4=27, other 0
		{
			val=int(libwps::readU8(input));
			if (val)
				f << "f" << i << "=" << std::hex << val << std::dec << ",";
		}
		break;
	case 2:
	{
		f << "dup,";
		if (sz!=8)
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readRowFormats: the size seems bad\n"));
			f << "###sz,";
			break;
		}
		auto sheetId2=int(libwps::readU8(input));
		if (sheetId2!=sheetId)
			f << "#sheetId2=" << sheetId2 << ",";
		val=int(libwps::readU8(input)); // always 0?
		if (val)
			f << "f0=" << val << ",";
		val=int (libwps::readU16(input));
		if (val>=row)
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readRowFormats: the original row seems bad\n"));
			f << "#";
		}
		m_state->m_rowSheetIdToChildRowIdMap.insert
		(std::multimap<Vec2i,Vec2i>::value_type(Vec2i(val,sheetId2),Vec2i(row,sheetId)));
		f << "orig[row]=" << val << ",";
		break;
	}
	default:
		WPS_DEBUG_MSG(("LotusSpreadsheet::readRowFormats: find unknown row type\n"));
		f << "###type=" << rowType << ",";
		break;
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

bool LotusSpreadsheet::readRowFormat(std::shared_ptr<WPSStream> stream, LotusSpreadsheetInternal::Style &style, int &numCell, long endPos)
{
	if (!stream) return false;
	numCell=1;

	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugStream f;
	long actPos=input->tell();
	if (endPos-actPos<4)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readRowFormats: the zone seems too short\n"));
		return false;
	}

	int value[3];
	for (int i=0; i<3; ++i)
		value[i]=(i==1) ? int(libwps::readU16(input)) : int(libwps::readU8(input));
	WPSFont font;
	if (value[2]&0x80)
	{
		if (actPos+5>endPos)
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readRowFormats: the zone seems too short\n"));
			input->seek(actPos, librevenge::RVNG_SEEK_SET);
			return false;
		}
		value[2]&=0x7F;
		numCell=1+int(libwps::readU8(input));
	}
	if ((value[0]&0x80)==0)
		f << "protected?,";
	switch ((value[0]>>4)&7)
	{
	case 0: // fixed
		f << "fixed,";
		style.setFormat(WPSCellFormat::F_NUMBER, 1);
		style.setDigits(value[0]&0xF);
		break;
	case 1: // scientific
		style.setFormat(WPSCellFormat::F_NUMBER, 2);
		style.setDigits(value[0]&0xF);
		break;
	case 2: // currency
		style.setFormat(WPSCellFormat::F_NUMBER, 4);
		style.setDigits(value[0]&0xF);
		break;
	case 3: // percent
		style.setFormat(WPSCellFormat::F_NUMBER, 3);
		style.setDigits(value[0]&0xF);
		break;
	case 4: // decimal
		style.setFormat(WPSCellFormat::F_NUMBER, 1);
		style.setDigits(value[0]&0xF);
		break;
	case 7:
		switch (value[0]&0xF)
		{
		case 0: // +/- : kind of bool
			style.setFormat(WPSCellFormat::F_BOOLEAN);
			f << "+/-,";
			break;
		case 1:
			style.setFormat(WPSCellFormat::F_NUMBER, 0);
			break;
		case 2:
			style.setDTFormat(WPSCellFormat::F_DATE, "%d %B %y");
			break;
		case 3:
			style.setDTFormat(WPSCellFormat::F_DATE, "%d %B");
			break;
		case 4:
			style.setDTFormat(WPSCellFormat::F_DATE, "%B %y");
			break;
		case 5:
			style.setFormat(WPSCellFormat::F_TEXT);
			break;
		case 6:
			style.setFormat(WPSCellFormat::F_TEXT);
			font.m_attributes |= WPS_HIDDEN_BIT;
			break;
		case 7:
			style.setDTFormat(WPSCellFormat::F_TIME, "%I:%M:%S%p");
			break;
		case 8:
			style.setDTFormat(WPSCellFormat::F_TIME, "%I:%M%p");
			break;
		case 9:
			style.setDTFormat(WPSCellFormat::F_DATE, "%m/%d/%y");
			break;
		case 0xa:
			style.setDTFormat(WPSCellFormat::F_DATE, "%m/%d");
			break;
		case 0xb:
			style.setDTFormat(WPSCellFormat::F_TIME, "%H:%M:%S");
			break;
		case 0xc:
			style.setDTFormat(WPSCellFormat::F_TIME, "%H:%M");
			break;
		case 0xd:
			style.setFormat(WPSCellFormat::F_TEXT);
			f << "label,";
			break;
		case 0xf: // automatic
			break;
		default:
			WPS_DEBUG_MSG(("LotusSpreadsheet::readRowFormat: find unknown 7e format\n"));
			f << "Fo=##7e,";
			break;
		}
		break;
	default:
		WPS_DEBUG_MSG(("LotusSpreadsheet::readRowFormat: find unknown %x format\n", static_cast<unsigned int>(value[0]&0x7F)));
		f << "##Fo=" << std::hex << (value[0]&0x7F) << std::dec << ",";
		break;
	}

	switch (value[2]&3)
	{
	case 1:
		style.setHAlignment(WPSCellFormat::HALIGN_LEFT);
		break;
	case 2:
		style.setHAlignment(WPSCellFormat::HALIGN_RIGHT);
		break;
	case 3:
		style.setHAlignment(WPSCellFormat::HALIGN_CENTER);
		break;
	default: // general
		break;
	}

	if (value[1]&1)
		f << "red[neg],";
	if (value[1]&2)
		f << "add[parenthesis],";
	/*
	  Now can either find some font definitions or a type id. I am not
	  sure how to distinguish these two cases ; this code does not
	  seem robust and may fail on some files...
	 */
	int wh=(value[2]>>2);
	int const vers=version();
	if (vers==1 && (wh&0x10))
	{
		int fId=(value[1]>>6)&0x3F;
		if (fId==0)
			;
		else if ((wh&0xf)==5)
		{
			/* unsure about this code, seems to work for Lotus123 mac,
			   but I never find a cell style in Lotus123 pc and this
			   part can be called*/
			if (!m_styleManager->updateCellStyle(fId, style, font, style.m_fontType))
				f << "#";
			f << "Ce" << fId << ",";
			wh &= 0xE0;
		}
		else if ((wh&0xf)==0)
		{
			if (!m_styleManager->updateFontStyle(fId, font, style.m_fontType))
				f << "#";
			f << "FS" << fId << ",";
			wh &= 0xE0;
		}
		else
			f << "#fId=" << fId << ",";
		value[1] &= 0xF03C;
	}
	else if (wh&0x10)
	{
		int fId=(value[1]>>6);
		if (fId==0)
			;
		else if ((wh&0xf)==0)
		{
			if (!m_styleManager->updateCellStyle(fId, style, font, style.m_fontType))
				f << "#";
			f << "Ce" << fId << ",";
			wh &= 0xE0;
		}
		else
			f << "#fId=" << fId << ",";
		value[1] &= 0x3C;
	}
	else
	{
		if (value[1]&0x40)
			font.m_attributes |= WPS_BOLD_BIT;
		if (value[1]&0x80)
			font.m_attributes |= WPS_ITALICS_BIT;
		if (value[1]>>11)
			font.m_size=(value[1]>>11);
		// values[1]&0x20 is often set in this case...
		value[1] &= 0x033c;
	}
	if (value[1])
		f << "f1=" << std::hex << value[1] << std::dec << ",";
	if (wh)
		f << "wh=" << std::hex << wh << std::dec << ",";
	if (font.m_size<=0)
		font.m_size=10; // if the size is not defined, set it to default
	style.setFont(font);
	style.m_extra=f.str();
	return true;
}

bool LotusSpreadsheet::readCellsFormat801(std::shared_ptr<WPSStream> stream, WPSVec3i const &minC, WPSVec3i const &maxC, int subZoneId)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::read16(input));
	if (type != 0x801)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readCellsFormat801: not a cells formats definition\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos = pos+4+sz;
	if (sz<0 || !stream->checkFilePosition(endPos))
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	f << "Entries(Zone8):";
	if (minC==maxC)
		f << "setStyle[" << minC << "],";
	else
		f << "setStyle[" << minC << "<->" << maxC << "],";
	LotusSpreadsheetInternal::Table123Styles *tableStyle=nullptr;
	int const vers=version();
	if (minC[0]<=maxC[0] && minC[0]>=0)
		tableStyle=m_state->getTablesStyle(Vec2i(minC[0],maxC[0]));

	int val;
	switch (sz)
	{
	case 2: // in general the second one
		val=int(libwps::readU16(input));
		if ((val>>8)==0x50)
		{
			if (tableStyle) tableStyle->addCellStyle(Vec2i(minC[1],maxC[1]),Vec2i(minC[2],maxC[2]),(val&0xFF));
			f << "Ce" << (val&0xFF);
		}
		else
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readCellsFormat801: find unexpected format type\n"));
			f << "##" << std::hex << val << std::dec << ",";
		}
		break;
	case 4:   // in general the first one
	{
		LotusSpreadsheetInternal::Format123Style format;
		int values[4];
		for (int &value : values) value=int(libwps::readU8(input));
		if ((values[0]&0x80)==0)
			f << "protected[no],";
		else
			values[0] &= 0x7f;
		if ((values[3]&0x80) && (values[0]>>4)==0)
		{
			if (values[2]>=0x7a)
				format.setDTFormat(WPSCellFormat::F_TIME, "%I:%M:%S%p");
			else if (values[2]>=0x63)
				format.setDTFormat(WPSCellFormat::F_DATE, "%m/%d/%y");
			else
			{
				format.setFormat(WPSCellFormat::F_NUMBER, 4);
				format.setDigits(values[0]&0xF);
			}
			values[3]&=0x7f;
		}
		else if (values[0]!=0x7f)
		{
			switch ((values[0]>>4))
			{
			case 0: // fixed
				f << "fixed,";
				format.setFormat(WPSCellFormat::F_NUMBER, 1);
				format.setDigits(values[0]&0xF);
				break;
			case 1: // scientific
				format.setFormat(WPSCellFormat::F_NUMBER, 2);
				format.setDigits(values[0]&0xF);
				break;
			case 2: // currency
				format.setFormat(WPSCellFormat::F_NUMBER, 4);
				format.setDigits(values[0]&0xF);
				break;
			case 3: // percent
				format.setFormat(WPSCellFormat::F_NUMBER, 3);
				format.setDigits(values[0]&0xF);
				break;
			case 4: // decimal
				format.setFormat(WPSCellFormat::F_NUMBER, 1);
				format.setDigits(values[0]&0xF);
				break;
			case 7: // checkme
				switch (values[0]&0xF)
				{
				case 0: // +/- : kind of bool
					format.setFormat(WPSCellFormat::F_BOOLEAN);
					f << "+/-,";
					break;
				case 1:
					format.setFormat(WPSCellFormat::F_NUMBER, 0);
					break;
				case 2:
					format.setDTFormat(WPSCellFormat::F_DATE, "%d %B %y");
					break;
				case 3:
					format.setDTFormat(WPSCellFormat::F_DATE, "%d %B");
					break;
				case 4:
					format.setDTFormat(WPSCellFormat::F_DATE, "%B %y");
					break;
				case 5:
					format.setFormat(WPSCellFormat::F_TEXT);
					break;
				case 6:
					format.setFormat(WPSCellFormat::F_TEXT);
					// font.m_attributes |= WPS_HIDDEN_BIT;
					break;
				case 7:
					format.setDTFormat(WPSCellFormat::F_TIME, "%I:%M:%S%p");
					break;
				case 8:
					format.setDTFormat(WPSCellFormat::F_TIME, "%I:%M%p");
					break;
				case 9:
					format.setDTFormat(WPSCellFormat::F_DATE, "%m/%d/%y");
					break;
				case 0xa:
					format.setDTFormat(WPSCellFormat::F_DATE, "%m/%d");
					break;
				case 0xb:
					format.setDTFormat(WPSCellFormat::F_TIME, "%H:%M:%S");
					break;
				case 0xc:
					format.setDTFormat(WPSCellFormat::F_TIME, "%H:%M");
					break;
				case 0xd:
					format.setFormat(WPSCellFormat::F_TEXT);
					f << "label,";
					break;
				default:
					WPS_DEBUG_MSG(("LotusSpreadsheet::readCellsFormat801: find unknown 7e format\n"));
					f << "Fo=##7e,";
					break;
				}
				break;
			default:
				WPS_DEBUG_MSG(("LotusSpreadsheet::readCellsFormat801: find unknown %x format\n", static_cast<unsigned int>(values[0])));
				f << "##Fo=" << std::hex << values[0] << std::dec << ",";
				break;
			}
		}
		f << format << ",";

		if (values[1]&1) f << "neg[value,red],";
		if (values[1]&2) f << "add[parenthesis],";
		if (values[1]&0x10)
		{
			format.m_alignAcrossColumn=true;
			f << "align[across,column],";
		}
		if (values[1]&0x20) f << "hidden,";
		values[1] &=0xCC;
		for (int i=1; i<4; ++i)
		{
			if (values[i]) f << "f" << i << "=" << std::hex << values[i] << std::dec << ",";
		}
		if (tableStyle) tableStyle->addCellStyle(Vec2i(minC[1],maxC[1]),Vec2i(minC[2],maxC[2]),format);
		break;
	}
	case 8:   // in general the third one, this define border, ...
	{
		LotusSpreadsheetInternal::Extra123Style style;
		for (int i=0; i<2; ++i)
		{
			auto col=int(libwps::readU8(input));
			val=int(libwps::readU8(input));
			if ((val&0xF)==0xF) continue;
			f << (i==0 ?  "bordT" : "bordL") << "=[";
			WPSBorder border;
			switch (val&0xF)
			{
			case 0:
				border.m_style=WPSBorder::None;
				break;
			case 1: // single
				break;
			case 2:
				border.m_type=WPSBorder::Double;
				break;
			case 3:
				border.m_width=2;
				break;
			case 4: // dot[1x1]
				border.m_style=WPSBorder::Dot;
				break;
			case 5:
				border.m_style=WPSBorder::Dash;
				f << "dash[1x3],";
				break;
			case 6:
				border.m_style=WPSBorder::Dash;
				f << "dash[3x1],";
				break;
			case 7:
				border.m_style=WPSBorder::Dash;
				f << "dash[1x1,3x1],";
				break;
			case 8:
				border.m_style=WPSBorder::Dash;
				f << "dash[1x1,1x1,2x1],";
				break;
			default:
				f << "##type=" << (val&0xF) << ",";
				break;
			}
			if (!m_styleManager->getColor256(col,border.m_color))
				f << "##colId=" << col << ",";
			f << border;
			f << "],";
			style.m_borders[i]=border;
		}
		f << "unk0=[";
		for (int i=0; i<4; ++i)
		{
			val=int(libwps::readU8(input));
			if (val)
				f <<std::hex << val << std::dec << ",";
			else
				f << "_,";
		}
		f << "],";
		if (tableStyle) tableStyle->addCellStyle(Vec2i(minC[1],maxC[1]),Vec2i(minC[2],maxC[2]),style);
		break;
	}
	case 12:   // column data, in general the fourst one
	{
		if (subZoneId==0) f << "col,";
		else if (subZoneId==1) f << "row,";
		else if (subZoneId!=-1)
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readCellsFormat801: the zone8 id seems bad\n"));
			f << "###zone15=" << subZoneId << ",";
		}
		int width=-1;
		bool isDefault=false;
		bool isWidthDef=true;
		f << "colUnkn=[";
		for (int i=0; i<7; ++i)   // f1=[01][04][01], f2: big number, f3=0|1, f4=0|2d|3c|240
		{
			if (i==2 && vers>=5) continue;
			val=(vers<5 && (i==2||i==4)) ? int(libwps::readU8(input)) : int(libwps::readU16(input));
			if (i==1)
			{
				if ((val&1)==0)
				{
					isDefault=true;
					f << "no[w],";
				}
				if (val&2) f << "hidden,";
				if (val&0x20) f << "page[break],";
				if (val&0x40)
				{
					isWidthDef=false;
					f << "w[def],";
				}
				if (val&0x100)
					f << "fl100,";
				val &= 0xFE9C;
				if (val)
					f << "##fl=" << std::hex << val << std::dec << ",";
			}
			else if (i==3)
			{
				width=val;
				if (!isDefault)
					f << "w=" << width << ",";
			}
			else if (val)
				f <<std::hex << val << std::dec << ",";
			else
				f << "_,";
		}
		f << "],";
		if (isDefault || minC[1]<0 || width<0) break;
		if (subZoneId<0 || subZoneId>1) break;
		for (int i=minC[0]; i<=maxC[0]; ++i)
		{
			auto &sheet=m_state->getSheet(i);
			for (int c=minC[1]; c<=std::min(maxC[1], MAX_COLUMNS); ++c)
			{
				if (subZoneId==0)
				{
					WPSColumnFormat format(!isWidthDef ? 72 : vers>=5 ? float(width)/16.f : float(width));
					sheet.setColumnWidth(c, format);
				}
				else
				{
					WPSRowFormat format(vers>=5 ? float(width)/16.f : float(width));
					format.m_useOptimalHeight=!isWidthDef;
					sheet.setRowHeight(c, format);
				}
			}
		}
		break;
	}
	case 30: // table data, rare, the last one
		for (int i=0; i<2; ++i)   // f0=71|7e
		{
			val=int(libwps::readU16(input));
			if (val)
				f << "f" << i << "=" << val << ",";
		}
		f << "],";
		val=int(libwps::readU16(input));
		if ((val>>8)==0x50)
		{
			if (tableStyle) tableStyle->m_defaultCellId=(val&0xFF);
			f << "Ce" << (val&0xFF) << ",";
		}
		else
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readCellsFormat801: find unexpected format type\n"));
			f << "##" << std::hex << val << std::dec << ",";
		}
		f << "tableUnk=[";
		for (int i=0; i<12; ++i)
		{
			// f0=168|438|3000|3600, f2=f0|14a|a008|c00,f4=[02]0[04c]0,f6=0|3
			// f8=[37]5d4,f10=0|4,f11=1003(L3?)
			val=int(libwps::readU16(input));
			if (val)
				f <<std::hex << val << std::dec << ",";
			else
				f << "_,";
		}
		f << "],";
		break;
	default:
		break;
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	if (input->tell()!=endPos)
	{
		ascFile.addDelimiter(input->tell(),'|');
		input->seek(endPos, librevenge::RVNG_SEEK_SET);
	}
	return true;
}

bool LotusSpreadsheet::readRowSizes(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	long sz=endPos-pos;
	f << "Entries(RowSize):";
	if (sz<10 || (sz%8)!=2)
	{
		WPS_DEBUG_MSG(("LotusParser::readRowSizes: the zone size seems bad\n"));
		f << "###";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}

	auto sheetId=int(libwps::readU8(input));
	f << "id[sheet]=" << sheetId << ",";
	LotusSpreadsheetInternal::Spreadsheet empty, *sheet=nullptr;
	if (sheetId<0||sheetId>=int(m_state->m_spreadsheetList.size()))
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readRowSizes: can find spreadsheet %d\n", sheetId));
		sheet=&empty;
		f << "###";
	}
	else
		sheet=&m_state->m_spreadsheetList[size_t(sheetId)];
	auto val=int(libwps::readU8(input));  // always 0
	if (val) f << "f0=" << val << ",";
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	auto N=int(sz/8);
	for (int i=0; i<N; ++i)
	{
		pos=input->tell();
		f.str("");
		f << "RowSize-" << i << ":";
		auto row=int(libwps::readU16(input));
		f << "row=" << row << ",";
		val=int(libwps::readU16(input));
		if (val!=0xFFFF)
		{
			f << "dim=" << float(val+31)/32.f << ",";
			sheet->setRowHeight(row, WPSRowFormat(float(val+31)/32.f));
		}
		for (int j=0; j<2; ++j)
		{
			val=int(libwps::read16(input));
			if (val!=j-1)
				f << "f" << j << "=" << val <<",";
		}
		input->seek(pos+8, librevenge::RVNG_SEEK_SET);
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
	}

	return true;
}

bool LotusSpreadsheet::readSheetName(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::read16(input));
	if (type != 0x23)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readSheetName: not a sheet name\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	f << "Entries(SheetName):";
	if (sz < 5)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readSheetName: sheet name is too short\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto val=int(libwps::read16(input)); // always 14000
	if (val!=14000)
		f << "f0=" << std::hex << val << std::dec << ",";
	auto sheetId=int(libwps::readU8(input));
	f << "id[sheet]=" << sheetId << ",";
	val=int(libwps::readU8(input)); // always 0
	if (val)
		f << "f1=" << val << ",";
	std::string name;
	for (long i = 0; i < sz-4; i++)
	{
		char c = char(libwps::readU8(input));
		if (c == '\0') break;
		name.push_back(c);
	}
	f << name << ",";
	if (input->tell()!=pos+4+sz && input->tell()+1!=pos+4+sz)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readSheetName: the zone seems too short\n"));
		f << "##";
		ascFile.addDelimiter(input->tell(), '|');
	}
	if (sheetId<0||sheetId>=m_state->getNumSheet())
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readSheetName: the zone id seems bad\n"));
		f << "##id";
	}
	else if (!name.empty())
		m_state->getSheet(sheetId).m_name=libwps_tools_win::Font::unicodeString(name, m_mainParser.getDefaultFontType());
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusSpreadsheet::readSheetHeader(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::read16(input));
	if (type!=0xc3)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readSheetHeader: not a sheet header\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	f << "Entries(FMTSheetBegin):";
	if (sz != 0x22)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readSheetHeader: the zone size seems bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto id=int(libwps::read16(input));
	if (id<0)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readSheetHeader: the id seems bad\n"));
		f << "###";
		m_state->m_sheetCurrentId=-1;
	}
	else
		m_state->m_sheetCurrentId=id;
	f << "id=" << id << ",";
	for (int i=0; i<16; ++i)   // always 0
	{
		auto val=int(libwps::read16(input));
		if (val) f << "f" << i << "=" << val << ",";
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusSpreadsheet::readExtraRowFormats(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::read16(input));
	if (type!=0xc5)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readExtraRowFormats: not a sheet header\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	f << "Entries(FMTRowForm):";
	if (sz < 9 || (sz%5)!=4)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readExtraRowFormats: the zone size seems bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto row=int(libwps::readU16(input));
	f << "row=" << row << ",";

	auto &sheet=m_state->getSheet(m_state->m_sheetCurrentId);
	auto val=int(libwps::readU8(input));
	sheet.setRowHeight(row, WPSRowFormat(float(val)));
	if (val!=14) f << "height=" << val << ",";
	val=int(libwps::readU8(input)); // 10|80
	if (val) f << "f0=" << std::hex << val << std::dec << ",";
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	LotusSpreadsheetInternal::ExtraRowStyles badRow;
	auto *rowStyles=&badRow;
	if (sheet.m_rowToExtraStyleMap.find(row)==sheet.m_rowToExtraStyleMap.end())
	{
		sheet.m_rowToExtraStyleMap[row]=LotusSpreadsheetInternal::ExtraRowStyles();
		rowStyles=&sheet.m_rowToExtraStyleMap.find(row)->second;
	}
	else
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readExtraRowFormats: row %d is already defined\n", row));
	}
	auto N=int(sz/5);
	int begPos=0;
	for (int i=0; i<N; ++i)
	{
		pos=input->tell();
		f.str("");
		f << "FMTRowForm-" << i << ":";
		LotusSpreadsheetInternal::ExtraStyle style;
		val=style.m_format=int(libwps::readU8(input));
		// find also f[8-c]ffffffXX, which seems to have a different meaning
		if ((val>>4)==0xf) f << "#";
		if (val&0x7) f << "font[id]=" << (val&0x7) << ","; // useMe
		if (val&0x8) f << "bold,";
		if (val&0x10) f << "italic,";
		if (val&0x20) f << "underline,";
		val &= 0xC0;
		if (val) f << "fl=" << std::hex << val << std::dec << ",";
		style.m_flag=val=int(libwps::readU8(input));
		if (val&0x20) f << "special,";
		if (!m_styleManager->getColor8(val&7, style.m_color))
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readExtraRowFormats: can not read a color\n"));
			f << "##colId=" << (val&7) << ",";
		}
		else if (!style.m_color.isBlack()) f << "col=" << style.m_color << ",";
		val &= 0xD8;
		if (val) f << "fl1=" << std::hex << val << std::dec << ",";
		val=int(libwps::readU8(input));
		if (val&7)
		{
			if ((val&7)==7) // checkMe
				style.m_backColor=WPSColor::black();
			else if (!m_styleManager->getColor8(val&7, style.m_backColor))
			{
				WPS_DEBUG_MSG(("LotusSpreadsheet::readExtraRowFormats: can not read a color\n"));
				f << "##colId=" << (val&7) << ",";
			}
			else
				f << "col[back]=" << style.m_backColor << ",";
		}
		if (val&0x10) f << "shadow2";
		if (val&0x20) f << "shadow,";
		val &= 0xD8;
		if (val) f << "f0=" << std::hex << val << std::dec << ",";
		style.m_borders=val=int(libwps::readU8(input));
		if (val) f << "border=" << std::hex << val << std::dec << ",";
		val=int(libwps::readU8(input));
		if (val) f << "dup=" << val << ",";
		rowStyles->m_colsToStyleMap[Vec2i(begPos, begPos+val)]=style;
		begPos+=1+val;
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		input->seek(pos+5, librevenge::RVNG_SEEK_SET);
	}
	if (begPos!=256)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readExtraRowFormats: the number of columns for row %d seems bad\n", row));
	}
	return true;
}

////////////////////////////////////////////////////////////
// Cell
////////////////////////////////////////////////////////////
bool LotusSpreadsheet::readCellName(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::read16(input));
	if (type!=9)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readCellName: not a cell name cell\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	f << "Entries(CellName):";
	if (sz < 0x1a)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readCellName: the zone is too short\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto val=int(libwps::read16(input)); // 0 or 1
	if (val)
		f << "f0=" << val << ",";
	std::string name("");
	for (int i=0; i<16; ++i)
	{
		auto c=char(libwps::readU8(input));
		if (!c) break;
		name += c;
	}
	f << name << ",";
	input->seek(pos+4+18, librevenge::RVNG_SEEK_SET);
	LotusSpreadsheetInternal::CellsList cells;
	for (int i=0; i<2; ++i)
	{
		auto row=int(libwps::readU16(input));
		auto sheetId=int(libwps::readU8(input));
		auto col=int(libwps::readU8(input));
		if (i==0)
			cells.m_positions.setMin(Vec2i(col,row));
		else
			cells.m_positions.setMax(Vec2i(col,row));
		cells.m_ids[i]=sheetId;
	}
	f << cells << ",";
	if (m_state->m_nameToCellsMap.find(name)!=m_state->m_nameToCellsMap.end())
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readCellName: cell with name %s already exists\n", name.c_str()));
		f << "##name=" << name << ",";
	}
	else
		m_state->m_nameToCellsMap[name]=cells;
	std::string note("");
	auto remain=int(endPos-input->tell());
	for (int i=0; i<remain; ++i)
	{
		auto c=char(libwps::readU8(input));
		if (!c) break;
		note+=c;
	}
	if (!note.empty())
		f << "note=" << note << ",";
	if (input->tell()!=endPos)
		ascFile.addDelimiter(input->tell(),'|');

	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusSpreadsheet::readCell(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::read16(input));
	std::string what("");
	if (type == 0x16)
		what="TextCell";
	else if (type == 0x17)
		what="Doub10Cell";
	else if (type == 0x18)
		what="DoubU16Cell";
	else if (type == 0x19)
		what="Doub10FormCell";
	else if (type == 0x1a) // checkme
		what="TextFormCell";
	else if (type == 0x25)
		what="DoubU32Cell";
	else if (type == 0x26)
		what="CommentCell";
	else if (type == 0x27)
		what="Doub8Cell";
	else if (type == 0x28)
		what="Doub8FormCell";
	else
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readCell: not a cell's cell\n"));
		return false;
	}

	auto sz = long(libwps::readU16(input));
	if (sz < 5)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readCell: the zone is too short\n"));
		f << "Entries(" << what << "):###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	long endPos=pos+4+sz;
	auto row=int(libwps::readU16(input));
	auto sheetId=int(libwps::readU8(input));
	auto col=int(libwps::readU8(input));
	if (sheetId) f << "sheet[id]=" << sheetId << ",";

	LotusSpreadsheetInternal::Spreadsheet empty, *sheet=nullptr;
	if (sheetId<0||sheetId>=int(m_state->m_spreadsheetList.size()))
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readCell: can find spreadsheet %d\n", sheetId));
		sheet=&empty;
		f << "###";
	}
	else
		sheet=&m_state->m_spreadsheetList[size_t(sheetId)];

	auto &cell=sheet->getCell(stream->m_input, Vec2i(col, row));
	switch (type)
	{
	case 0x16:
	case 0x1a:
	case 0x26:   // comment
	{
		std::string text("");
		long begText=input->tell();
		for (long i=4; i<sz; ++i)
		{
			auto c=char(libwps::readU8(input));
			if (!c) break;
			if (i==4)
			{
				bool done=true;
				if (c=='\'') cell.m_hAlignement=WPSCellFormat::HALIGN_DEFAULT; // sometimes followed by 0x7C
				else if (c=='\\') cell.m_hAlignement=WPSCellFormat::HALIGN_LEFT;
				else if (c=='^') cell.m_hAlignement=WPSCellFormat::HALIGN_CENTER;
				else if (c=='\"') cell.m_hAlignement=WPSCellFormat::HALIGN_RIGHT;
				else done=false;
				if (done)
				{
					++begText;
					continue;
				}
			}
			text += c;
		}
		f << "\"" << getDebugStringForText(text) << "\",";
		WPSEntry entry;
		entry.setBegin(begText);
		entry.setEnd(endPos);
		if (type==0x16)
		{
			cell.m_content.m_contentType=WKSContentListener::CellContent::C_TEXT;
			cell.m_content.m_textEntry=entry;
		}
		else if (type==0x1a)
		{
			if (cell.m_content.m_contentType!=WKSContentListener::CellContent::C_FORMULA)
				cell.m_content.m_contentType=WKSContentListener::CellContent::C_TEXT;
			cell.m_content.m_textEntry=entry;
		}
		else if (type==0x26)
			cell.m_comment=entry;
		else
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readCell: find unexpected type %x\n", unsigned(type)));
			f << "###type";
		}

		if (input->tell()!=endPos && input->tell()+1!=endPos)
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readCell: the string zone seems too short\n"));
			f << "###";
		}
		break;
	}
	case 0x17:
	{
		if (sz!=14)
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readCell: the double10 zone seems too short\n"));
			f << "###";
		}
		double res;
		bool isNaN;
		if (!libwps::readDouble10(input, res, isNaN))
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readCell: can read a double10 zone\n"));
			f << "###";
			break;
		}
		if (cell.m_content.m_contentType!=WKSContentListener::CellContent::C_FORMULA)
			cell.m_content.m_contentType=WKSContentListener::CellContent::C_NUMBER;
		cell.m_content.setValue(res);
		break;
	}
	case 0x18:
	{
		if (sz!=6)
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readCell: the uint16 zone seems too short\n"));
			f << "###";
		}
		double res;
		bool isNaN;
		if (!libwps::readDouble2Inv(input, res, isNaN))
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readCell: can read a uint16 zone\n"));
			f << "###";
			break;
		}
		if (cell.m_content.m_contentType!=WKSContentListener::CellContent::C_FORMULA)
			cell.m_content.m_contentType=WKSContentListener::CellContent::C_NUMBER;
		cell.m_content.setValue(res);
		break;
	}
	case 0x19:
	{
		if (sz<=14)
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readCell: the double10+formula zone seems too short\n"));
			f << "###";
		}
		double res;
		bool isNaN;
		if (!libwps::readDouble10(input, res, isNaN))
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readCell: can read double10+formula for zone\n"));
			f << "###";
			break;
		}
		cell.m_content.m_contentType=WKSContentListener::CellContent::C_FORMULA;
		cell.m_content.setValue(res);
		ascFile.addDelimiter(input->tell(),'|');
		std::string error;
		if (!readFormula(*stream, endPos, sheetId, false, cell.m_content.m_formula, error))
		{
			cell.m_content.m_contentType=WKSContentListener::CellContent::C_NUMBER;
			ascFile.addDelimiter(input->tell()-1, '#');
			if (error.length()) f << error;
			break;
		}
		if (error.length()) f << error;
		if (input->tell()+1>=endPos)
			break;
		static bool first=true;
		if (first)
		{
			first=false;
			WPS_DEBUG_MSG(("LotusSpreadsheet::readCell: find err message for double10+formula\n"));
		}
		// find in one file "Formula failed to convert"
		error="";
		auto remain=int(endPos-input->tell());
		for (int i=0; i<remain; ++i) error += char(libwps::readU8(input));
		f << "#err[msg]=" << error << ",";
		break;
	}
	case 0x25:
	{
		if (sz!=8)
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readCell: the uint32 zone seems too short\n"));
			f << "###";
		}
		double res;
		bool isNaN;
		if (!libwps::readDouble4Inv(input, res, isNaN))
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readCell: can read a uint32 zone\n"));
			f << "###";
			break;
		}
		if (cell.m_content.m_contentType!=WKSContentListener::CellContent::C_FORMULA)
			cell.m_content.m_contentType=WKSContentListener::CellContent::C_NUMBER;
		cell.m_content.setValue(res);
		break;
	}
	case 0x27:
	{
		if (sz!=12)
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readCell: the double8 zone seems too short\n"));
			f << "###";
		}
		double res;
		bool isNaN;
		if (!libwps::readDouble8(input, res, isNaN))
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readCell: can read a double8 zone\n"));
			f << "###";
			break;
		}
		if (cell.m_content.m_contentType!=WKSContentListener::CellContent::C_FORMULA)
			cell.m_content.m_contentType=WKSContentListener::CellContent::C_NUMBER;
		cell.m_content.setValue(res);
		break;
	}
	case 0x28:
	{
		if (sz<=12)
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readCell: the double8 formula zone seems too short\n"));
			f << "###";
		}
		double res;
		bool isNaN;
		if (!libwps::readDouble8(input, res, isNaN))
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::readCell: can read a double8 formula zone\n"));
			f << "###";
			break;
		}
		cell.m_content.m_contentType=WKSContentListener::CellContent::C_FORMULA;
		cell.m_content.setValue(res);
		ascFile.addDelimiter(input->tell(),'|');
		std::string error;
		if (!readFormula(*stream, endPos, sheetId, true, cell.m_content.m_formula, error))
		{
			cell.m_content.m_contentType=WKSContentListener::CellContent::C_NUMBER;
			ascFile.addDelimiter(input->tell()-1, '#');
		}
		else if (input->tell()+1<endPos)
		{
			// often end with another bytes 03, probably for alignement
			WPS_DEBUG_MSG(("LotusSpreadsheet::readCell: find extra data for double8 formula zone\n"));
			f << "###extra";
		}
		if (error.length()) f << error;
		break;
	}
	default:
		WPS_DEBUG_MSG(("LotusSpreadsheet::readCell: oops find unimplemented type\n"));
		break;
	}
	std::string extra=f.str();
	f.str("");
	f << "Entries(" << what << "):" << cell << "," << extra;
	if (input->tell()!=endPos)
		ascFile.addDelimiter(input->tell(),'|');
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	return true;
}

////////////////////////////////////////////////////////////
// filter
////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////
// report
////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////
// send data
////////////////////////////////////////////////////////////
void LotusSpreadsheet::sendSpreadsheet(int sheetId)
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::sendSpreadsheet: I can not find the listener\n"));
		return;
	}
	if (sheetId<0||sheetId>=m_state->getNumSheet())
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::sendSpreadsheet: the sheet %d seems bad\n", sheetId));
		return;
	}
	auto &sheet = m_state->getSheet(sheetId);
	m_listener->openSheet(sheet.getWidths(), getSheetName(sheetId));
	m_mainParser.sendGraphics(sheetId);
	sheet.compressRowHeights();
	/* create a set to know which row needed to be send, each value of
	   the set corresponding to a position where the rows change
	   excepted the last position */
	std::set<int> newRowSet;
	newRowSet.insert(0);
	std::map<Vec2i, LotusSpreadsheetInternal::Cell>::const_iterator cIt;
	int prevRow=-1;
	for (cIt=sheet.m_positionToCellMap.begin(); cIt!=sheet.m_positionToCellMap.end(); ++cIt)
	{
		if (prevRow==cIt->first[1])
			continue;
		prevRow=cIt->first[1];
		newRowSet.insert(prevRow);
		newRowSet.insert(prevRow+1);
	}
	size_t numRowStyle=m_state->m_rowStylesList.size();
	for (auto rIt : sheet.m_rowToStyleIdMap)
	{
		Vec2i const &rows=rIt.first;
		size_t listId=rIt.second;
		if (listId>=numRowStyle)
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::sendSpreadsheet: can not find list %d\n", int(listId)));
			continue;
		}
		newRowSet.insert(rows[0]);
		newRowSet.insert(rows[1]+1);
	}
	for (auto rIt : sheet.m_rowHeightMap)
	{
		Vec2i const &rows=rIt.first;
		newRowSet.insert(rows[0]);
		newRowSet.insert(rows[1]+1);
	}
	for (auto rIt : sheet.m_rowToExtraStyleMap)
	{
		if (rIt.second.empty()) continue;
		newRowSet.insert(rIt.first);
		newRowSet.insert(rIt.first+1);
	}

	LotusSpreadsheetInternal::Table123Styles const *table123Styles=m_state->getTableStyle(sheetId);
	if (table123Styles)
	{
		for (auto rIt : table123Styles->m_rowsToColsToCellIdMap)
		{
			Vec2i const &rows=rIt.first;
			newRowSet.insert(rows[0]);
			newRowSet.insert(rows[1]+1);
		}
		for (auto rIt : table123Styles->m_rowsToColsToExtraStyleMap)
		{
			Vec2i const &rows=rIt.first;
			newRowSet.insert(rows[0]);
			newRowSet.insert(rows[1]+1);
		}
		for (auto rIt : table123Styles->m_rowsToColsToFormatStyleMap)
		{
			Vec2i const &rows = rIt.first;
			newRowSet.insert(rows[0]);
			newRowSet.insert(rows[1]+1);
		}
	}
	for (auto sIt=newRowSet.begin(); sIt!=newRowSet.end();)
	{
		int row=*(sIt++);
		if (row<0)
		{
			WPS_DEBUG_MSG(("LotusSpreadsheet::sendSpreadsheet: find a negative row %d\n", row));
			continue;
		}
		if (sIt==newRowSet.end())
			break;
		m_listener->openSheetRow(sheet.getRowHeight(row), *sIt-row);
		sendRowContent(sheet, row, table123Styles);
		m_listener->closeSheetRow();
	}
	m_listener->closeSheet();
}

void LotusSpreadsheet::sendRowContent(LotusSpreadsheetInternal::Spreadsheet const &sheet, int row, LotusSpreadsheetInternal::Table123Styles const *table123Styles)
{
	if (m_listener.get() == nullptr)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::sendRowContent: I can not find the listener\n"));
		return;
	}

	// create a set of column we need to generate
	std::set<int> newColSet;
	newColSet.insert(0);

	auto cIt=sheet.m_positionToCellMap.lower_bound(Vec2i(-1, row));
	bool checkCell=cIt!=sheet.m_positionToCellMap.end() && cIt->first[1]==row;
	if (checkCell)
	{
		for (; cIt!=sheet.m_positionToCellMap.end() && cIt->first[1]==row ; ++cIt)
		{
			newColSet.insert(cIt->first[0]);
			newColSet.insert(cIt->first[0]+1);
		}
		cIt=sheet.m_positionToCellMap.lower_bound(Vec2i(-1, row));
	}

	LotusSpreadsheetInternal::RowStyles const *styles=nullptr;
	int styleId=sheet.getRowStyleId(row);
	if (styleId>=int(m_state->m_rowStylesList.size()))
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::sendRowContent: I can not row style %d\n", styleId));
	}
	else if (styleId>=0)
		styles=&m_state->m_rowStylesList[size_t(styleId)];
	std::map<Vec2i, LotusSpreadsheetInternal::Style>::const_iterator sIt;
	if (styles)
	{
		for (sIt=styles->m_colsToStyleMap.begin(); sIt!=styles->m_colsToStyleMap.end(); ++sIt)
		{
			newColSet.insert(sIt->first[0]);
			newColSet.insert(sIt->first[1]+1);
		}
		sIt=styles->m_colsToStyleMap.begin();
		if (sIt==styles->m_colsToStyleMap.end())
			styles=nullptr;
	}

	LotusSpreadsheetInternal::ExtraRowStyles const *extraStyles=nullptr;
	if (sheet.m_rowToExtraStyleMap.find(row)!=sheet.m_rowToExtraStyleMap.end())
		extraStyles=&sheet.m_rowToExtraStyleMap.find(row)->second;
	std::map<Vec2i, LotusSpreadsheetInternal::ExtraStyle>::const_iterator eIt;
	if (extraStyles)
	{
		for (eIt=extraStyles->m_colsToStyleMap.begin(); eIt!=extraStyles->m_colsToStyleMap.end(); ++eIt)
		{
			if (eIt->second.empty()) continue;
			newColSet.insert(eIt->first[0]);
			newColSet.insert(eIt->first[1]+1);
		}
		eIt=extraStyles->m_colsToStyleMap.begin();
		if (eIt==extraStyles->m_colsToStyleMap.end())
			extraStyles=nullptr;
	}

	auto const defFontType=m_mainParser.getDefaultFontType();
	LotusSpreadsheetInternal::Style defaultStyle(defFontType);

	std::map<Vec2i, LotusSpreadsheetInternal::Style> colToCellIdMap;
	std::map<Vec2i, LotusSpreadsheetInternal::Style>::const_iterator c123It;
	std::map<Vec2i, LotusSpreadsheetInternal::Extra123Style> colToExtraStyleMap;
	std::map<Vec2i, LotusSpreadsheetInternal::Extra123Style>::const_iterator e123It;
	std::map<Vec2i, LotusSpreadsheetInternal::Format123Style> colToFormatStyleMap;
	std::map<Vec2i, LotusSpreadsheetInternal::Format123Style>::const_iterator f123It;

	std::map<int,int> potentialMergeMap;
	if (table123Styles)
	{
		for (auto rIt : table123Styles->m_rowsToColsToCellIdMap)
		{
			if (rIt.first[0]>row || rIt.first[1]<row) continue;
			for (auto colIt : rIt.second)
			{
				LotusSpreadsheetInternal::Style style(defFontType);
				WPSFont font;
				if (!m_styleManager->updateCellStyle(colIt.second, style, font, style.m_fontType))
					continue;
				style.setFont(font);
				colToCellIdMap.insert(std::map<Vec2i, LotusSpreadsheetInternal::Style>::value_type(colIt.first,style));
				newColSet.insert(colIt.first[0]);
				newColSet.insert(colIt.first[1]+1);
			}
		}
		for (auto rIt : table123Styles->m_rowsToColsToExtraStyleMap)
		{
			if (rIt.first[0]>row || rIt.first[1]<row) continue;
			for (auto colIt : rIt.second)
			{
				colToExtraStyleMap.insert(std::map<Vec2i, LotusSpreadsheetInternal::Extra123Style>::value_type(colIt.first,colIt.second));
				newColSet.insert(colIt.first[0]);
				newColSet.insert(colIt.first[1]+1);
			}
		}
		for (auto rIt : table123Styles->m_rowsToColsToFormatStyleMap)
		{
			if (rIt.first[0]>row || rIt.first[1]<row) continue;
			for (auto colIt : rIt.second)
			{
				colToFormatStyleMap.insert(std::map<Vec2i, LotusSpreadsheetInternal::Format123Style>::value_type(colIt.first,colIt.second));
				newColSet.insert(colIt.first[0]);
				newColSet.insert(colIt.first[1]+1);
				if (colIt.first[0]!=colIt.first[1] && colIt.second.m_alignAcrossColumn)
					potentialMergeMap[colIt.first[0]]=colIt.first[1]+1;
			}
		}
		c123It=colToCellIdMap.begin();
		e123It=colToExtraStyleMap.begin();
		f123It=colToFormatStyleMap.begin();
		if (table123Styles->m_defaultCellId>=0)
		{
			WPSFont font;
			if (m_styleManager->updateCellStyle(table123Styles->m_defaultCellId, defaultStyle, font, defaultStyle.m_fontType))
				defaultStyle.setFont(font);
		}
		if (colToCellIdMap.empty() && colToExtraStyleMap.empty() && colToFormatStyleMap.empty())
			table123Styles=nullptr;
	}

	LotusSpreadsheetInternal::Cell emptyCell;
	for (auto colIt=newColSet.begin(); colIt!=newColSet.end();)
	{
		int const col=*colIt;
		++colIt;
		if (colIt==newColSet.end())
			break;
		int const endCol=*colIt;
		if (styles)
		{
			while (sIt->first[1] < col)
			{
				++sIt;
				if (sIt==styles->m_colsToStyleMap.end())
				{
					styles=nullptr;
					break;
				}
			}
		}
		if (extraStyles)
		{
			while (eIt->first[1] < col)
			{
				++eIt;
				if (eIt==extraStyles->m_colsToStyleMap.end())
				{
					extraStyles=nullptr;
					break;
				}
			}
		}
		auto style=styles ? sIt->second : defaultStyle;
		bool hasStyle=styles!=nullptr;
		if (table123Styles)
		{
			while (c123It!=colToCellIdMap.end() && c123It->first[1] < col)
				++c123It;
			if (c123It!=colToCellIdMap.end() && c123It->first[0] <= col && c123It->first[1] >= col)
			{
				style=c123It->second;
				hasStyle=true;
			}

			while (e123It!=colToExtraStyleMap.end() && e123It->first[1] < col)
				++e123It;
			if (e123It!=colToExtraStyleMap.end() && e123It->first[0] <= col && e123It->first[1] >= col)
			{
				e123It->second.update(style);
				hasStyle=true;
			}
			while (f123It!=colToFormatStyleMap.end() && f123It->first[1] < col)
				++f123It;
			if (f123It!=colToFormatStyleMap.end() && f123It->first[0] <= col && f123It->first[1] >= col)
			{
				f123It->second.update(style);
				hasStyle=true;
			}
			if (c123It==colToCellIdMap.end() && e123It==colToExtraStyleMap.end() && f123It==colToFormatStyleMap.end())
				table123Styles=nullptr;
		}

		if (extraStyles && !eIt->second.empty())
		{
			eIt->second.update(style);
			hasStyle=true;
		}
		bool hasCell=false;
		if (checkCell)
		{
			while (cIt->first[1]==row && cIt->first[0]<col)
			{
				++cIt;
				if (cIt==sheet.m_positionToCellMap.end())
				{
					checkCell=false;
					break;
				}
			}
			if (checkCell && cIt->first[1]!=row) checkCell=false;
			hasCell=checkCell && cIt->first[0]==col;
		}
		if (!hasCell && !hasStyle) continue;
		if (!hasCell) emptyCell.setPosition(Vec2i(col, row));
		bool canMerge=false;
		if (hasCell && potentialMergeMap.find(col)!=potentialMergeMap.end())
		{
			int newEndCol=potentialMergeMap.find(col)->second;
			auto newCIt=cIt;
			if (newCIt!=sheet.m_positionToCellMap.end()) ++newCIt;
			canMerge=(newCIt==sheet.m_positionToCellMap.end() || newCIt->first[1]!=row || newCIt->first[0]>=newEndCol);
		}
		if (canMerge)
		{
			int newEndCol=potentialMergeMap.find(col)->second;
			while (colIt!=newColSet.end() && *colIt<newEndCol)
			{
				++colIt;
				continue;
			}
			auto cell=cIt->second;
			cell.setNumSpannedCells(Vec2i(newEndCol-col,1));
			sendCellContent(cell, style, 1);
		}
		else
			sendCellContent(hasCell ? cIt->second : emptyCell, style, endCol-col);
	}
}

void LotusSpreadsheet::sendCellContent(LotusSpreadsheetInternal::Cell const &cell,
                                       LotusSpreadsheetInternal::Style const &style,
                                       int numRepeated)
{
	if (m_listener.get() == nullptr)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::sendCellContent: I can not find the listener\n"));
		return;
	}

	LotusSpreadsheetInternal::Style cellStyle(style);
	if (cell.m_hAlignement!=WPSCellFormat::HALIGN_DEFAULT)
		cellStyle.setHAlignment(cell.m_hAlignement);

	auto fontType = cellStyle.m_fontType;

	m_listener->setFont(cellStyle.getFont());

	LotusSpreadsheetInternal::Cell finalCell(cell);
	finalCell.WPSCellFormat::operator=(cellStyle);
	WKSContentListener::CellContent content(cell.m_content);
	for (auto &f : content.m_formula)
	{
		if (f.m_type!=WKSContentListener::FormulaInstruction::F_Text)
			continue;
		std::string &text=f.m_content;
		librevenge::RVNGString finalString=libwps_tools_win::Font::unicodeString(text, fontType);
		text=finalString.cstr();
	}
	m_listener->openSheetCell(finalCell, content, numRepeated);

	if (cell.m_input && cell.m_content.m_textEntry.valid())
	{
		RVNGInputStreamPtr input=cell.m_input;
		input->seek(cell.m_content.m_textEntry.begin(), librevenge::RVNG_SEEK_SET);
		sendText(input, cell.m_content.m_textEntry.end(), cellStyle);
	}
	if (cell.m_comment.valid())
	{
		WPSSubDocumentPtr subdoc(new LotusSpreadsheetInternal::SubDocument
		                         (cell.m_input, *this, cell.m_comment));
		m_listener->insertComment(subdoc);
	}
	m_listener->closeSheetCell();
}

////////////////////////////////////////////////////////////
// formula
////////////////////////////////////////////////////////////
bool LotusSpreadsheet::parseVariable(std::string const &variable, WKSContentListener::FormulaInstruction &instr)
{
	std::regex exp("<<([^>]+)>>([^:]+):([A-Z]+)([0-9]+)");
	std::regex exp2("<<([^>]+)>>([^:]+):([A-Z]+)([0-9]+)\\.\\.([^:]+):([A-Z]+)([0-9]+)");
	std::smatch base_match;
	size_t dim=0;
	if (std::regex_match(variable,base_match,exp) && base_match.size() == 5)
		dim=1;
	else if (std::regex_match(variable,base_match,exp2) && base_match.size() == 8)
	{
		if (base_match[2].str()!=base_match[5].str())
			instr.m_sheetName[1]=base_match[5].str().c_str();
		dim=2;
	}
	else // can also be variable, <<File>>variable, db.field ?
		return false;
	instr.m_fileName=base_match[1].str().c_str();
	instr.m_sheetName[0]=base_match[2].str().c_str();
	instr.m_type=dim==1 ? WKSContentListener::FormulaInstruction::F_Cell : WKSContentListener::FormulaInstruction::F_CellList;
	for (size_t d=0; d<dim; ++d)
	{
		int col=0;
		std::string sCol=base_match[3+3*d].str();
		for (auto c:sCol)
		{
			if (col > (std::numeric_limits<int>::max() - int(c-'A')) / 26)
			{
				WPS_DEBUG_MSG(("LotusSpreadsheet::parseVariable: oops the column seems bad\n"));
				return false;
			}
			col=26*col+int(c-'A');
		}
		int row=0;
		std::string sRow=base_match[4+3*d].str();
		for (auto c:sRow)
		{
			if (row > (std::numeric_limits<int>::max() - int(c-'0')) / 10)
			{
				WPS_DEBUG_MSG(("LotusSpreadsheet::parseVariable: oops the row seems bad\n"));
				return false;
			}
			row=10*row+int(c-'0');
		}
		instr.m_position[d]=Vec2i(col,row-1);
		instr.m_positionRelative[d]=Vec2b(true,true);
	}
	return true;
}

bool LotusSpreadsheet::readCell(WPSStream &stream, int sId, bool isList, WKSContentListener::FormulaInstruction &instr)
{
	RVNGInputStreamPtr &input=stream.m_input;
	instr=WKSContentListener::FormulaInstruction();
	instr.m_type=isList ? WKSContentListener::FormulaInstruction::F_CellList :
	             WKSContentListener::FormulaInstruction::F_Cell;
	auto flags=int(libwps::readU8(input));
	for (int i=0; i<2; ++i)
	{
		auto row=int(libwps::readU16(input));
		auto sheetId=int(libwps::readU8(input));
		auto col=int(libwps::readU8(input));
		instr.m_position[i]=Vec2i(col, row);
		int wh=(i==0) ? (flags&0xF) : (flags>>4);
		instr.m_positionRelative[i]=Vec2b((wh&1)!=0, (wh&2)!=0);
		if (sheetId!=sId)
			instr.m_sheetName[i]=getSheetName(sheetId);
		if (!isList) break;
	}
	return true;
}

namespace LotusSpreadsheetInternal
{
struct Functions
{
	char const *m_name;
	int m_arity;
};

static Functions const s_listFunctions[] =
{
	{ "", 0} /*SPEC: number*/, {"", 0}/*SPEC: cell*/, {"", 0}/*SPEC: cells*/, {"=", 1} /*SPEC: end of formula*/,
	{ "(", 1} /* SPEC: () */, {"", 0}/*SPEC: number*/, { "", 0} /*SPEC: text*/, {"", 0}/*name reference*/,
	{ "", 0}/* SPEC: abs name ref*/, {"", 0}/* SPEC: err range ref*/, { "", 0}/* SPEC: err cell ref*/, {"", 0}/* SPEC: err constant*/,
	{ "", -2} /* unused*/, { "", -2} /*unused*/, {"-", 1}, {"+", 2},

	{ "-", 2},{ "*", 2},{ "/", 2},{ "^", 2},
	{ "=", 2},{ "<>", 2},{ "<=", 2},{ ">=", 2},
	{ "<", 2},{ ">", 2},{ "And", 2},{ "Or", 2},
	{ "Not", 1}, { "+", 1}, { "&", 2}, { "NA", 0} /* not applicable*/,

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
	{ "Text", 2}/* or fixed*/, { "Mid", 3}, { "Char", 1},{ "Ascii", 1},
	{ "Find", 3},{ "DateValue", 1} /*checkme*/,{ "TimeValue", 1} /*checkme*/,{ "CellPointer", 1} /*checkme*/,

	{ "Sum", -1},{ "Average", -1},{ "COUNT", -1},{ "Min", -1},
	{ "Max", -1},{ "VLookUp", 3},{ "NPV", 2}, { "Var", -1},
	{ "StDev", -1},{ "IRR", 2} /*BAD*/, { "HLookup", 3},{ "DSum", 3},
	{ "DAvg", 3},{ "DCount", 3},{ "DMin", 3},{ "DMax", 3},

	{ "DVar", 3},{ "DStd", 3},{ "Index", 3}, { "Columns", 1},
	{ "Rows", 1},{ "Rept", 2},{ "Upper", 1},{ "Lower", 1},
	{ "Left", 2},{ "Right", 2},{ "Replace", 4}, { "Proper", 1},
	{ "Cell", 2} /*checkme*/,{ "Trim", 1},{ "Clean", 1} /*UNKN*/,{ "T", 1},

	{ "IsNonText", 1},{ "Exact", 2},{ "", -2} /*App not implemented*/,{ "", 3} /*UNKN*/,
	{ "Rate", 3} /*BAD*/,{ "TERM", 3}, { "CTERM", 3}, { "SLN", 3},
	{ "SYD", 4},{ "DDB", 4},{ "SplFunc", -1} /*SplFunc*/,{ "Sheets", 1},
	{ "Info", 1},{ "SumProduct", -1},{ "IsRange", 1},{ "DGet", -1},

	{ "DQuery", -1},{ "Coord", 4}, { "", -2} /*reserved*/, { "Today", 0},
	{ "Vdb", -1},{ "Dvars", -1},{ "Dstds", -1},{ "Vars", -1},
	{ "Stds", -1},{ "D360", 2},{ "", -2} /*reserved*/,{ "IsApp", 0},
	{ "IsAaf", -1},{ "Weekday", 1},{ "DateDiff", 3},{ "Rank", -1},

	{ "NumberString", 2},{ "DateString", 1}, { "Decimal", 1}, { "Hex", 1},
	{ "Db", 4},{ "PMTI", 4},{ "SPI", 4},{ "Fullp", 1},
	{ "Halfp", 1},{ "PureAVG", -1},{ "PureCount", -1},{ "PureMax", -1},
	{ "PureMin", -1},{ "PureSTD", -1},{ "PureVar", -1},{ "PureSTDS", -1},

	{ "PureVars", -1},{ "PMT2", 3}, { "PV2", 3}, { "FV2", 3},
	{ "TERM2", 3},{ "", -2} /*UNKN*/,{ "D360", 2},{ "", -2} /*UNKN*/,
	{ "", -2} /*UNKN*/,{ "", -2} /*UNKN*/,{ "", -2} /*UNKN*/, { "", -2} /*UNKN*/,
	{ "", -2} /*UNKN*/,{ "", -2} /*UNKN*/,{ "", -2} /*UNKN*/,{ "", -2} /*UNKN*/,
};
}

bool LotusSpreadsheet::readFormula(WPSStream &stream, long endPos, int sheetId, bool newFormula,
                                   std::vector<WKSContentListener::FormulaInstruction> &formula, std::string &error)
{
	int const vers=version();
	RVNGInputStreamPtr &input=stream.m_input;
	formula.resize(0);
	error = "";
	long pos = input->tell();
	if (endPos - pos < 1) return false;

	std::stringstream f;
	std::vector<std::vector<WKSContentListener::FormulaInstruction> > stack;
	bool ok = true;
	while (long(input->tell()) != endPos)
	{
		double val = 0.0;
		bool isNaN;
		pos = input->tell();
		if (pos > endPos) return false;
		auto wh = int(libwps::readU8(input));
		int arity = 0;
		WKSContentListener::FormulaInstruction instr;
		switch (wh)
		{
		case 0x0:
			if ((!newFormula && (endPos-pos<11 || !libwps::readDouble10(input, val, isNaN))) ||
			        (newFormula && (endPos-pos<9 || !libwps::readDouble8(input, val, isNaN))))
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
		{
			if (endPos-pos<6 || !readCell(stream, sheetId, false, instr))
			{
				f.str("");
				f << "###cell short";
				error=f.str();
				ok = false;
				break;
			}
			break;
		}
		case 0x2:
		{
			if (endPos-pos<10 || !readCell(stream, sheetId, true, instr))
			{
				f.str("");
				f << "###list cell short";
				error=f.str();
				ok = false;
				break;
			}
			break;
		}
		case 0x5:
			instr.m_type=WKSContentListener::FormulaInstruction::F_Double;
			if ((!newFormula && (endPos-pos<3 || !libwps::readDouble2Inv(input, val, isNaN))) ||
			        (newFormula && (endPos-pos<5 || !libwps::readDouble4Inv(input, val, isNaN))))
			{
				WPS_DEBUG_MSG(("LotusSpreadsheet::readFormula: can read a uint16/32 zone\n"));
				f << "###uint16/32";
				error=f.str();
				break;
			}
			instr.m_doubleValue=val;
			break;
		case 0x6:
			instr.m_type=WKSContentListener::FormulaInstruction::F_Text;
			while (!input->isEnd())
			{
				if (input->tell() >= endPos)
				{
					ok=false;
					break;
				}
				auto c = char(libwps::readU8(input));
				if (c==0) break;
				instr.m_content += c;
			}
			break;
		case 0x7:
		case 0x8:
		{
			std::string variable("");
			while (!input->isEnd())
			{
				if (input->tell() >= endPos)
				{
					ok=false;
					break;
				}
				auto c = char(libwps::readU8(input));
				if (c==0) break;
				variable += c;
			}
			if (!ok)
				break;
			if (m_state->m_nameToCellsMap.find(variable)==m_state->m_nameToCellsMap.end())
			{
				if (parseVariable(variable, instr)) break;
				// can also be a database field, ...
				WPS_DEBUG_MSG(("LotusSpreadsheet::readFormula: can not find variable %s\n", variable.c_str()));
				f << "##variable=" << variable << ",";
				error=f.str();
				instr.m_type=WKSContentListener::FormulaInstruction::F_Text;
				instr.m_content = variable;
				break;
			}
			auto const &cells=m_state->m_nameToCellsMap.find(variable)->second;
			instr.m_position[0]=cells.m_positions[0];
			instr.m_position[1]=cells.m_positions[1];
			instr.m_positionRelative[0]=instr.m_positionRelative[1]=Vec2b(wh==7,wh==7);
			for (int i=0; i<2; ++i)
			{
				if (cells.m_ids[i] != sheetId)
					instr.m_sheetName[i]=getSheetName(cells.m_ids[i]);
			}
			instr.m_type=cells.m_positions[0]==cells.m_positions[1] ?
			             WKSContentListener::FormulaInstruction::F_Cell :
			             WKSContentListener::FormulaInstruction::F_CellList;
			break;
		}
		default:
			if (wh >= 0xb0 || LotusSpreadsheetInternal::s_listFunctions[wh].m_arity == -2)
			{
				f.str("");
				f << "##Funct" << std::hex << wh;
				error=f.str();
				ok = false;
				break;
			}
			instr.m_type=WKSContentListener::FormulaInstruction::F_Function;
			instr.m_content=LotusSpreadsheetInternal::s_listFunctions[wh].m_name;
			ok=!instr.m_content.empty();
			arity = LotusSpreadsheetInternal::s_listFunctions[wh].m_arity;
			if (arity == -1)
				arity = int(libwps::read8(input));
			if (wh==0x7a)   // special Spell function
			{
				auto sSz=int(libwps::readU16(input));
				if (input->tell()+sSz>endPos || (vers>=3 && sSz<2))
				{
					WPS_DEBUG_MSG(("LotusSpreadsheet::readFormula: can not find spell function length\n"));
					f << "###spell[length]=" << sSz << ",";
					error = f.str();
					ok=false;
				}
				if (vers>=3)   // 1801
				{
					f << "f0=" << std::dec << libwps::readU16(input) << std::dec << ",";
					sSz-=2;
				}
				WKSContentListener::FormulaInstruction lastArg;
				lastArg.m_type=WKSContentListener::FormulaInstruction::F_Text;
				for (int i=0; i<sSz; ++i)
				{
					auto c = char(libwps::readU8(input));
					if (c==0) break;
					lastArg.m_content += c;
				}
				std::vector<WKSContentListener::FormulaInstruction> child;
				child.push_back(lastArg);
				stack.push_back(child);
				++arity;
				break;
			}
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
			if (wh==3) break;
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
		return true;
	}
	else
		error = "###stack problem";

	static bool first = true;
	if (first)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readFormula: I can not read some formula\n"));
		first = false;
	}

	f.str("");
	for (size_t i = 0; i < stack.size(); ++i)
	{
		if (i) f << "##";
		for (auto const &j : stack[i])
			f << j << ",";
	}
	f << error;
	error = f.str();
	return false;
}

// ------------------------------------------------------------
// zone 1b
// ------------------------------------------------------------
bool LotusSpreadsheet::readSheetName1B(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	long sz=endPos-pos;
	f << "Entries(SheetName):";
	if (sz<3)
	{
		WPS_DEBUG_MSG(("LotusParser::readSheetName1B: the zone size seems bad\n"));
		f << "###";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto sheetId=int(libwps::readU16(input));
	f << "id=" << sheetId << ",";
	std::string name;
	for (long i=2; i<sz; ++i)
	{
		char c = char(libwps::readU8(input));
		if (c == '\0') break;
		name.push_back(c);
	}
	f << name << ",";
	if (sheetId<0||sheetId>=m_state->getNumSheet())
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::readSheetName: the zone id seems bad\n"));
		f << "##id";
	}
	else if (!name.empty())
		m_state->getSheet(sheetId).m_name=libwps_tools_win::Font::unicodeString(name, m_mainParser.getDefaultFontType());
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusSpreadsheet::readNote(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	long sz=endPos-pos;
	f << "Entries(Note):";
	if (sz<4)
	{
		WPS_DEBUG_MSG(("LotusParser::readNote: the zone size seems bad\n"));
		f << "###";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	static bool first=true;
	if (first)
	{
		first=false;
		WPS_DEBUG_MSG(("LotusParser::readNote: this spreadsheet contains some notes, but there is no code to retrieve them\n"));
	}
	f << "id=" << int(libwps::readU8(input)) << ",";
	for (int i=0; i<2; ++i)   // f0=1, f1=2|4
	{
		auto val=int(libwps::readU8(input));
		if (val!=i+1) f << "f" << i << "=" << val << ",";
	}
	std::string text;
	for (long i=0; i<sz-3; ++i) text+=char(libwps::readU8(input));
	f << getDebugStringForText(text) << ",";
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	return true;
}

//////////////////////////////////////////////////////////////////////
// formatted text
//////////////////////////////////////////////////////////////////////
void LotusSpreadsheet::sendText(RVNGInputStreamPtr &input, long endPos, LotusSpreadsheetInternal::Style const &style) const
{
	if (!input || !m_listener)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::sendText: can not find the listener\n"));
		return;
	}
	libwps_tools_win::Font::Type fontType = style.m_fontType;
	WPSFont font=style.getFont();
	m_listener->setFont(font);
	bool prevEOL=false;
	std::string text;
	while (!input->isEnd())
	{
		long pos=input->tell();
		auto c=pos>=endPos ? '\0' : char(libwps::readU8(input));
		if ((c==0 || c==1 || c==0xa || c==0xd) && !text.empty())
		{
			m_listener->insertUnicodeString(libwps_tools_win::Font::unicodeString(text, fontType));
			text.clear();
		}
		if (pos>=endPos) break;
		switch (c)
		{
		case 0x1:
			if (pos+1>=endPos)
			{
				WPS_DEBUG_MSG(("LotusSpreadsheet::sendText: can not read the escape value\n"));
				break;
			}
			c=char(libwps::readU8(input));
			switch (c)
			{
			case 0x1e:
				if (pos+2>=endPos)
				{
					WPS_DEBUG_MSG(("LotusSpreadsheet::sendText: can not read the escape value\n"));
					break;
				}
				c=char(libwps::readU8(input));
				switch (c)
				{
				case 'b': // bold
					font.m_attributes |= WPS_BOLD_BIT;
					m_listener->setFont(font);
					break;
				case 'i': // italic
					font.m_attributes |= WPS_ITALICS_BIT;
					m_listener->setFont(font);
					break;
				default:
					if (c>='0' && c<='7')
					{
						if (pos+3>=endPos)
						{
							WPS_DEBUG_MSG(("LotusSpreadsheet::sendText: can not read the escape value\n"));
							break;
						}
						auto c2=char(libwps::readU8(input));
						if (c2=='c' && m_styleManager->getColor8(c-'0', font.m_color))
							m_listener->setFont(font);
						else if (c2!='F')   // F for font?
						{
							WPS_DEBUG_MSG(("LotusSpreadsheet::sendText: unknown int sequence\n"));
							break;
						}
						break;
					}
					WPS_DEBUG_MSG(("LotusSpreadsheet::sendText: unknown sequence\n"));
					break;
				}
				break;
			case 0x1f: // end
				font=style.getFont();
				m_listener->setFont(font);
				break;
			case ';': // unknown, ie. the text can begin with 27013b in some mac file
				break;
			default:
				WPS_DEBUG_MSG(("LotusSpreadsheet::sendText: unknown debut sequence\n"));
				break;
			}
			break;
		case 0xd:
			m_listener->insertEOL();
			prevEOL=true;
			break;
		case 0xa:
			if (!prevEOL)
			{
				WPS_DEBUG_MSG(("LotusSpreadsheet::sendText: find 0xa without 0xd\n"));
			}
			prevEOL=false;
			break;
		default:
			if (c)
				text.push_back(c);
			break;
		}
	}
}

void LotusSpreadsheet::sendTextNote(RVNGInputStreamPtr &input, WPSEntry const &entry) const
{
	if (!input || !m_listener)
	{
		WPS_DEBUG_MSG(("LotusSpreadsheet::sendTextNote: can not find the listener\n"));
		return;
	}
	bool prevEOL=false;
	auto fontType=m_mainParser.getDefaultFontType();
	input->seek(entry.begin(), librevenge::RVNG_SEEK_SET);
	long endPos=entry.end();
	std::string text;
	while (!input->isEnd())
	{
		long pos=input->tell();
		auto c=pos>=endPos ? '\0' : char(libwps::readU8(input));
		if ((c==0 || c==0xa || c==0xd) && !text.empty())
		{
			m_listener->insertUnicodeString(libwps_tools_win::Font::unicodeString(text, fontType));
			text.clear();
		}
		if (pos>=endPos) break;
		switch (c)
		{
		case 0xd:
			m_listener->insertEOL();
			prevEOL=true;
			break;
		case 0xa:
			if (!prevEOL)
			{
				WPS_DEBUG_MSG(("LotusSpreadsheet::sendTextNote: find 0xa without 0xd\n"));
			}
			prevEOL=false;
			break;
		default:
			if (c)
				text.push_back(c);
			break;
		}
	}
}

std::string LotusSpreadsheet::getDebugStringForText(std::string const &text)
{
	std::string res;
	size_t len=text.length();
	for (size_t i=0; i<len; ++i)
	{
		char c=text[i];
		switch (c)
		{
		case 0x1:
			if (i+1>=len)
			{
				WPS_DEBUG_MSG(("LotusSpreadsheet::getDebugStringForText: can not read the escape value\n"));
				res+= "[##escape]";
				break;
			}
			c=text[++i];
			switch (c)
			{
			case 0x1e:
				if (i+1>=len)
				{
					WPS_DEBUG_MSG(("LotusSpreadsheet::getDebugStringForText: can not read the escape value\n"));
					res+= "[##escape1]";
					break;
				}
				c=text[++i];
				switch (c)
				{
				case 'b': // bold
				case 'i': // italic
					res+= std::string("[") + c + "]";
					break;
				default:
					if (c>='0' && c<='8')
					{
						if (i+1>=len)
						{
							WPS_DEBUG_MSG(("LotusSpreadsheet::getDebugStringForText: can not read the escape value\n"));
							res+= "[##escape1]";
							break;
						}
						char c2=text[++i];
						if (c2!='c' && c2!='F')   // color and F for font?
						{
							WPS_DEBUG_MSG(("LotusSpreadsheet::getDebugStringForText: unknown int sequence\n"));
							res+= std::string("[##") + c + c2 + "]";
							break;
						}
						res += std::string("[") + c + c2 + "]";
						break;
					}
					WPS_DEBUG_MSG(("LotusSpreadsheet::getDebugStringForText: unknown sequence\n"));
					res+= std::string("[##") + c + "]";
					break;
				}
				break;
			case 0x1f: // end
				res+= "[^]";
				break;
			// case ';': ?
			default:
				WPS_DEBUG_MSG(("LotusSpreadsheet::getDebugStringForText: unknown debut sequence\n"));
				res+= std::string("[##") +c+ "]";
				break;
			}
			break;
		case 0xd:
			res+="\\n";
			break;
		default:
			res+=c;
			break;
		}
	}
	return res;
}

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
