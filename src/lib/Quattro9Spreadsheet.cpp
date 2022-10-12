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
#include <limits>
#include <set>
#include <sstream>
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

#include "QuattroFormula.h"
#include "Quattro9.h"

#include "Quattro9Spreadsheet.h"

namespace Quattro9SpreadsheetInternal
{
//! a class used to store a style of a cell in Quattro9Spreadsheet
struct Style final : public WPSCellFormat
{
	//! construtor
	explicit Style()
		: WPSCellFormat()
		, m_fileFormat(-1)
		, m_alignAcrossColumn(false)
		, m_extra("")
	{
	}
	Style(Style const &) = default;
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
	if (m_fileFormat!=st.m_fileFormat) return false;
	int diff = WPSCellFormat::compare(st);
	if (diff) return false;
	return m_fileFormat==st.m_fileFormat && m_alignAcrossColumn==st.m_alignAcrossColumn && m_extra==st.m_extra;
}

//! a cellule of a Quattro spreadsheet
class Cell final : public WPSCell
{
public:
	/// constructor
	explicit Cell()
	{ }

	//! update the cell format using file format
	void updateFormat(int fileFormat)
	{
		if (fileFormat<0)
			return;
		switch (fileFormat>>5)
		{
		case 0:
			switch (fileFormat&0x1f)
			{
			case 1: // checkme: +/- : kind of bool
				setFormat(F_BOOLEAN);
				break;
			case 2: // generic
				break;
			case 3:
				setFormat(F_TEXT);
				break;
			case 4:   // hidden
			{
				auto font=getFont();
				font.m_attributes |= WPS_HIDDEN_BIT;
				setFont(font);
				break;
			}
			case 5:
				setDTFormat(F_DATE, "%d-%b-%y");
				break;
			case 6:
				setDTFormat(F_DATE, "%d %b");
				break;
			case 7:
				setDTFormat(F_DATE, "%b-%y");
				break;
			case 8:
				setDTFormat(F_DATE, "%A %d %B %Y");
				break;
			case 9:
				setDTFormat(F_DATE, "%m/%d/%Y");
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
				setDTFormat(F_DATE, "%d-%b-%Y");
				break;
			case 0xf:
				setDTFormat(F_DATE, "%b-%Y");
				break;
			default:
				WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::Cell::send: unknown format %x\n", fileFormat));
				break;
			}
			break;
		case 1: // basic, neg red
		case 2: // neg (), neg red ()
		case 3: // basic+thousand, neg red
		case 4:
			setFormat(F_NUMBER, 1);
			setDigits(fileFormat&0xF);
			break;
		case 5: // currency (see previous)
		case 6:
		case 7:
		case 8:
			setFormat(F_NUMBER, 4);
			setDigits(fileFormat&0xF);
			break;
		case 9: // percent, scientific
			setFormat(F_NUMBER, (fileFormat&0x10) ? 2 : 3);
			setDigits(fileFormat&0xF);
			break;
		case 0xa: // fraction (note fileFormat&0x1f seems to define the format of the denominator)
			setFormat(F_NUMBER, 7);
			break;
		case 0xb:
		{
			// id=(fileFormat>>4)-0x16
			static bool first=true;
			if (first)
			{
				first=false;
				WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::Cell::updateFormat: user defined format is not supported\n"));
			}
			break;
		}
		default:
			WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::Cell::send: unknown format %x\n", fileFormat));
			break;
		}
	}
	//! call when a cell must be send
	bool send(WPSListenerPtr &/*listener*/) final;

	//! call when the content of a cell must be send
	bool sendContent(WPSListenerPtr &/*listener*/) final
	{
		WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::Cell::sendContent: must not be called\n"));
		return false;
	}
};
bool Cell::send(WPSListenerPtr &/*listener*/)
{
	WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::Cell::send: must not be called\n"));
	return false;
}


//! a class used to store the cell(s) content of a Quattro Spreadsheet
struct CellData
{
	//! constructor
	CellData()
		: m_type(0)
		, m_rows()
		, m_span(1,1)
		, m_style(-1)
		, m_intList()
		, m_doubleList()
		, m_flagList()
	{
	}
	//! returns true if the cell contains no data (and is not a merged cell)
	bool empty() const
	{
		return (m_type&0x1f)==1 && m_span==Vec2i(1,1);
	}
	//! returns the double value corresponding to a row
	double getDouble(int row) const
	{
		if (m_doubleList.empty())
		{
			WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::CellData::getDouble: no int value\n"));
			return 0;
		}
		if (row<m_rows[0]||row>m_rows[1])
		{
			WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::CellData::getDouble: called with bad row=%d\n", row));
			return 0;
		}
		switch ((m_type>>5)&3)
		{
		case 0:
			return m_doubleList[0];
		case 2:
			if (row-m_rows[0]>=int(m_doubleList.size()))
			{
				WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::CellData::getDouble: oops, can not find row=%d\n", row));
				return 0;
			}
			return m_doubleList[size_t(row-m_rows[0])];
		case 3:
			if (m_doubleList.size()!=2)
			{
				WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::CellData::getDouble: oops, unexpected data size\n"));
				return 0;
			}
			return m_doubleList[0]+(row-m_rows[0])*m_doubleList[1];
		default:
			break;
		}
		WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::CellData::getDouble: oops, unexpected type\n"));
		return 0;
	}
	//! returns the flag value corresponding to a row
	int getFlag(int row) const
	{
		if (m_flagList.empty())
		{
			WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::CellData::getFlag: no flag value\n"));
			return 0;
		}
		if (row<m_rows[0]||row>m_rows[1])
		{
			WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::CellData::getFlag: called with bad row=%d\n", row));
			return 0;
		}
		switch ((m_type>>5)&3)
		{
		case 0:
			return m_flagList[0];
		case 2:
			if (row-m_rows[0]>=int(m_flagList.size()))
			{
				WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::CellData::getFlag: oops, can not find row=%d\n", row));
				return 0;
			}
			return m_flagList[size_t(row-m_rows[0])];
		case 3:
			if (m_flagList.size()!=2)
			{
				WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::CellData::getFlag: oops, unexpected data size\n"));
				return 0;
			}
			return m_flagList[0]+(row-m_rows[0])*m_flagList[1];
		default:
			break;
		}
		WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::CellData::getFlag: oops, unexpected type\n"));
		return 0;
	}
	//! returns the int value corresponding to a row
	int getInt(int row) const
	{
		if (m_intList.empty())
		{
			WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::CellData::getInt: no int value\n"));
			return 0;
		}
		if (row<m_rows[0]||row>m_rows[1])
		{
			WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::CellData::getInt: called with bad row=%d\n", row));
			return 0;
		}
		switch ((m_type>>5)&3)
		{
		case 0:
			return m_intList[0];
		case 2:
			if (row-m_rows[0]>=int(m_intList.size()))
			{
				WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::CellData::getInt: oops, can not find row=%d\n", row));
				return 0;
			}
			return m_intList[size_t(row-m_rows[0])];
		case 3:
			if (m_intList.size()!=2)
			{
				WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::CellData::getInt: oops, unexpected data size\n"));
				return 0;
			}
			return m_intList[0]+(row-m_rows[0])*m_intList[1];
		default:
			break;
		}
		WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::CellData::getInt: oops, unexpected type\n"));
		return 0;
	}
	//! the cell type
	int m_type;
	//! the min/max row
	Vec2i m_rows;
	//! the column row/span
	Vec2i m_span;
	//! the style id
	int m_style;
	//! the list of int values
	std::vector<int> m_intList;
	//! the list of double values
	std::vector<double> m_doubleList;
	//! a list of flag (for formula)
	std::vector<int> m_flagList;
};

//! Internal: a list of cell and result of Quattro9SpreadsheetInternal
struct Column
{
	//! constructor
	Column()
		: m_rowsToCellMap()
		, m_rowToCellResMap()
	{
	}
	//! add a cell/list of cells data
	void add(Vec2i limits, CellData const &cell)
	{
		auto rIt=m_rowsToCellMap.lower_bound(Vec2i(-1,limits[0]));
		while (rIt!=m_rowsToCellMap.end())
		{
			auto const &cells=rIt->first;
			if (cells[0]>limits[1]) break;
			if (cells[1]>=limits[0])
			{
				WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::column::addCells: oops, some rows are already set in %dx%d\n", limits[0], limits[1]));
				return;
			}
			++rIt;
		}
		m_rowsToCellMap[limits]=cell;
	}
	//! add a cell result
	void add(int row, std::shared_ptr<WPSStream> const &stream, Quattro9ParserInternal::TextEntry const &entry)
	{
		if (m_rowToCellResMap.find(row)!=m_rowToCellResMap.end())
		{
			WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::column::addCells: oops, a result exists for row=%d\n", row));
			return;
		}
		m_rowToCellResMap[row]=std::make_pair(stream,entry);
	}
	//! update a list of row (needed to know what row must be sent)
	void updateListOfRows(std::set<int> &rows) const
	{
		for (auto const &it : m_rowsToCellMap)
		{
			rows.insert(it.first[0]);
			if (it.second.m_span[1]>1)
			{
				rows.insert(it.first[0]+1);
				rows.insert(it.first[0]+it.second.m_span[1]);
				continue;
			}
			rows.insert(it.first[1]+1);
			if ((it.second.m_type&0x1f)==1) continue;
			for (int r=it.first[0]+1; r<=it.first[1]; ++r) rows.insert(r);
		}
	}
	//! returns the cell corresponding to a row (if it exist)
	CellData const *getCell(int row) const
	{
		auto rIt=m_rowsToCellMap.lower_bound(Vec2i(-1,row));
		if (rIt!=m_rowsToCellMap.end() && rIt->first[1]<row) ++rIt;
		if (rIt!=m_rowsToCellMap.end() && rIt->first[0]<=row && rIt->first[1]>=row)
			return &rIt->second;
		return nullptr;
	}
	//! a map rows to cell data
	std::map<Vec2i, CellData> m_rowsToCellMap;
	//! a map row to cell result
	std::map<int, std::pair<std::shared_ptr<WPSStream>, Quattro9ParserInternal::TextEntry> > m_rowToCellResMap;
};

//! the spreadsheet of a Quattro Spreadsheet
class Spreadsheet
{
public:
	//! a constructor
	explicit Spreadsheet(int id)
		: m_id(id)
		, m_defaultSizes(1080,260)
		, m_mergedCellList()
		, m_colToColumnMap()
		, m_invalidColumn()
	{
	}
	//! update the spreadsheet: check for merged cell, ...
	void update()
	{
		for (auto const &box : m_mergedCellList)
		{
			int const minRow=box[0][1];
			int const maxRow=box[1][1];
			for (int col=box[0][0]; col<=box[1][0]; ++col)
			{
				auto cIt=m_colToColumnMap.find(col);
				bool const firstCol=col==box[0][0];
				if (cIt==m_colToColumnMap.end())
				{
					if (firstCol)   // we must create it
					{
						auto &column=getColumn(col);
						CellData cell;
						cell.m_type=1; // no data
						cell.m_rows=Vec2i(minRow, minRow);
						cell.m_span=Vec2i(box.size())+Vec2i(1,1);
						column.add(cell.m_rows, cell);
					}
					continue;
				}
				std::vector<CellData> listCellToAdd;
				auto &column=getColumn(col);
				auto rIt=column.m_rowsToCellMap.lower_bound(Vec2i(-1,minRow));
				if (rIt!=column.m_rowsToCellMap.end() && rIt->first[1]<minRow) ++rIt;
				if (rIt!=column.m_rowsToCellMap.end() && rIt->first[0]<minRow && rIt->first[1]>=minRow)
				{
					// we need to split the first data
					auto rows=rIt->first;
					auto cell=rIt->second;
					++rIt;
					column.m_rowsToCellMap.erase(rows);
					cell.m_rows[1]=minRow-1;
					listCellToAdd.push_back(cell);
					if (firstCol)
					{
						cell.m_rows=Vec2i(minRow, minRow);
						cell.m_span=Vec2i(box.size())+Vec2i(1,1);
						listCellToAdd.push_back(cell);
					}
				}
				else if (firstCol)
				{
					if (rIt!=column.m_rowsToCellMap.end() && rIt->first[0]==minRow)
					{
						// the first cell exist, update span
						auto cell=rIt->second;
						cell.m_rows=Vec2i(minRow, minRow);
						cell.m_span=Vec2i(box.size())+Vec2i(1,1);
						listCellToAdd.push_back(cell);
					}
					else
					{
						// create a new cell
						CellData cell;
						cell.m_type=1; // no data
						cell.m_rows=Vec2i(minRow, minRow);
						cell.m_span=Vec2i(box.size())+Vec2i(1,1);
						column.add(cell.m_rows, cell);
					}
				}
				while (rIt!=column.m_rowsToCellMap.end() && rIt->first[1]<=maxRow)
				{
					// remove cells between minRow and maxRow
					auto rows=rIt->first;
					++rIt;
					column.m_rowsToCellMap.erase(rows);
				}
				if (rIt!=column.m_rowsToCellMap.end() && rIt->first[0]<=maxRow)
				{
					// we need to split the last data
					auto rows=rIt->first;
					auto cell=rIt->second;
					++rIt;
					column.m_rowsToCellMap.erase(rows);
					cell.m_rows[0]=maxRow+1;
					listCellToAdd.push_back(cell);
				}
				for (auto const &cell: listCellToAdd) column.add(cell.m_rows, cell);
			}
		}
	}
	//! returns the row size in point
	WPSRowFormat getRowHeight(int row) const
	{
		auto rIt=m_sizesMap[1].lower_bound(Vec2i(-1,row));
		if (rIt!=m_sizesMap[1].end() && rIt->first[0]<=row && rIt->first[1]>=row)
			return WPSRowFormat(float(rIt->second)/20.f);
		return WPSRowFormat(float(m_defaultSizes[1])/20.f);
	}
	//! return the columns format
	std::vector<WPSColumnFormat> getWidths() const
	{
		std::vector<WPSColumnFormat> widths;
		// g++-mp-7 does not like float(m_defaultSizes[0])/20.f), but this->m_default is ok...
		WPSColumnFormat defWidth(float(this->m_defaultSizes[0])/20.f), actWidth;
		defWidth.m_useOptimalWidth=true;
		int prevRow=-1;
		for (auto const &it : m_sizesMap[0])
		{
			if (it.first[0]<prevRow+1)
			{
				WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::Spreadsheet::getWidths: oops, some limits are bad\n"));
				continue;
			}
			if (it.first[0]>prevRow+1)
			{
				defWidth.m_numRepeat=it.first[0]-(prevRow+1);
				widths.push_back(defWidth);
			}
			WPSColumnFormat width(float(it.second)/20.f);
			width.m_numRepeat=(it.first[1]+1-it.first[0]);
			widths.push_back(width);
			prevRow=it.first[1];
		}
		if (widths.empty())
		{
			defWidth.m_numRepeat=256;
			widths.push_back(defWidth);
		}
		return widths;
	}
	//! set the col/row size in TWIP
	void setColRowSize(int which, int pos, int w)
	{
		if (which!=0 && which!=1)
		{
			WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::Spreadsheet::setColRowSize: oops, which=%d is bad\n", which));
			return;
		}
		auto rIt=m_sizesMap[which].lower_bound(Vec2i(-1,pos));
		if (rIt!=m_sizesMap[which].end() && rIt->first[0]<=pos && rIt->first[1]>=pos)
		{
			WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::Spreadsheet::setColRowSize: oops, pos %d is already set\n", pos));
			return;
		}
		if (w>=0)
			m_sizesMap[which][Vec2i(pos,pos)]=w;
	}
	//! set the col/row size in TWIP
	void setColRowSizes(int which, int minPos, int maxPos, int w)
	{
		if (which!=0 && which!=1)
		{
			WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::Spreadsheet::setColRowSizes: oops, which=%d is bad\n", which));
			return;
		}
		auto rIt=m_sizesMap[which].lower_bound(Vec2i(-1,minPos));
		while (rIt!=m_sizesMap[which].end())
		{
			auto const &cells=rIt->first;
			if (cells[0]>maxPos) break;
			if (cells[1]>=minPos)
			{
				WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::Spreadsheet::setColRowSizes: oops, some rows are already set in %dx%d\n", minPos, maxPos));
				return;
			}
			++rIt;
		}
		if (w>=0)
			m_sizesMap[which][Vec2i(minPos,maxPos)]=w;
	}
	//! returns the cell position
	Vec2f getPosition(Vec2i const &cell) const
	{
		Vec2f res(0,0);
		for (int which=0; which<2; ++which)
		{
			int prevRow=-1;
			int width=0;
			for (auto const &it : m_sizesMap[0])
			{
				if (it.first[0]<prevRow+1) continue;
				if (it.first[0]>prevRow+1)
				{
					if (it.first[0]>cell[which])
					{
						width+=(cell[which]-(prevRow+1))*m_defaultSizes[which];
						prevRow=cell[which];
						break;
					}
					width+=(it.first[0]-(prevRow+1))*m_defaultSizes[which];
				}
				if (it.first[1]>cell[which])
				{
					width+=(cell[which]-it.first[0])*it.second;
					prevRow=cell[which];
					break;
				}
				width+=(it.first[1]+1-it.first[0])*it.second;
				prevRow=it.first[1];
			}
			if (cell[which]>prevRow+1)
				width+=(cell[which]-(prevRow+1))*m_defaultSizes[which];
			res[which]=float(width)/20.f;
		}
		return res;
	}
	//! returns a ref to column data (or to invalidColumn if called with bad column)
	Column &getColumn(int col)
	{
		auto it=m_colToColumnMap.find(col);
		if (it!=m_colToColumnMap.end()) return it->second;
		if (col<0 || col>m_defaultSizes[0])
		{
			WPS_DEBUG_MSG(("Quattro9SpreadsheetInternal::Spreadsheet::getColumn: called with invalid col=%d\n", col));
			return m_invalidColumn;
		}
		m_colToColumnMap[col]=Column();
		return m_colToColumnMap.find(col)->second;
	}
	//! returns the list of rows which need to be opened, ...
	std::set<int> getListSendRow() const
	{
		std::set<int> rows;
		rows.insert(0);
		for (auto const &it : m_colToColumnMap) it.second.updateListOfRows(rows);
		for (auto const &it :  m_sizesMap[1])
		{
			rows.insert(it.first[0]);
			rows.insert(it.first[1]+1);
		}
		for (auto const &box : m_mergedCellList)
		{
			rows.insert(box[0][1]);
			rows.insert(box[1][1]+1);
		}
		return rows;
	}
	//! returns the list of columns in which a row cell exists
	std::vector<Vec2i> getListCellsInARow(int row) const
	{
		std::vector<Vec2i> cols;
		CellData const *prevCell=nullptr;
		int prevCol=-1, numRepeated=0;
		for (auto const &it : m_colToColumnMap)
		{
			auto newCell=it.second.getCell(row);
			if (it.first==prevCol+numRepeated && prevCell && newCell &&
			        prevCell->empty() && newCell->empty() && prevCell->m_style==newCell->m_style)
			{
				numRepeated++;
				continue;
			}
			if (prevCell) cols.push_back(Vec2i(prevCol,prevCol+numRepeated-1));
			prevCol=it.first;
			prevCell=newCell;
			numRepeated=1;
		}
		if (prevCell) cols.push_back(Vec2i(prevCol,prevCol+numRepeated-1));
		return cols;
	}
	//! returns the cell corresponding to a position (if it exist)
	CellData const *getCell(Vec2i const &pos) const
	{
		auto rIt=m_colToColumnMap.find(pos[0]);
		if (rIt==m_colToColumnMap.end()) return nullptr;
		return rIt->second.getCell(pos[1]);
	}
	//! the spreadsheet id
	int m_id;

	/** the default col/row size in TWIP */
	Vec2i m_defaultSizes;
	/** the map Vec2i(min, max) to col/row size in TWIP */
	std::map<Vec2i,int> m_sizesMap[2];
	/** the merge cells */
	std::vector<WPSBox2i> m_mergedCellList;
	/** a map col to column data */
	std::map<int, Column> m_colToColumnMap;
private:
	/** a extra column used to return an invalid column */
	Column m_invalidColumn;
};

//! the state of Quattro9Spreadsheet
struct State
{
	//! constructor
	explicit State(QuattroFormulaManager::CellReferenceFunction const &readCellReference)
		: m_version(-1)
		, m_documentStrings()
		, m_documentFormulas()
		, m_formulaManager(readCellReference, 2)
		, m_actualSpreadsheet()
		, m_actualColumn(-1)
		, m_stylesList()
		, m_spreadsheetMap()
		, m_idToSheetNameMap()
		, m_idToUserFormatMap()
	{
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
	//! the document strings
	std::pair<std::shared_ptr<WPSStream>, std::vector<Quattro9ParserInternal::TextEntry> > m_documentStrings;
	//! the document formulas
	std::pair<std::shared_ptr<WPSStream>, std::vector<WPSEntry> > m_documentFormulas;
	//! the formula manager
	QuattroFormulaManager m_formulaManager;
	//! the actual sheet
	std::shared_ptr<Spreadsheet> m_actualSpreadsheet;
	//! the actual column
	int m_actualColumn;
	//! the list of styles
	std::vector<Style> m_stylesList;
	//! the map of spreadsheet
	std::map<int, std::shared_ptr<Spreadsheet> > m_spreadsheetMap;
	//! the map id to sheet's name
	std::map<int, librevenge::RVNGString> m_idToSheetNameMap;
	//! map id to user format string
	std::map<int, librevenge::RVNGString> m_idToUserFormatMap;
};

}

// constructor, destructor
Quattro9Spreadsheet::Quattro9Spreadsheet(Quattro9Parser &parser)
	: m_listener()
	, m_mainParser(parser)
	, m_state()
{
	m_state.reset(new Quattro9SpreadsheetInternal::State(getReadCellReferenceFunction()));
}

Quattro9Spreadsheet::~Quattro9Spreadsheet()
{
}

void Quattro9Spreadsheet::cleanState()
{
	m_state.reset(new Quattro9SpreadsheetInternal::State(getReadCellReferenceFunction()));
}

void Quattro9Spreadsheet::updateState()
{
}

int Quattro9Spreadsheet::version() const
{
	if (m_state->m_version<0)
		m_state->m_version=m_mainParser.version();
	return m_state->m_version;
}

QuattroFormulaManager::CellReferenceFunction Quattro9Spreadsheet::getReadCellReferenceFunction()
{
	return [this](std::shared_ptr<WPSStream> const &stream, long endPos,
	              QuattroFormulaInternal::CellReference &ref,
	              Vec2i const &pos, int sheetId)
	{
		return this->readCellReference(stream, endPos, ref, pos, sheetId);
	};
}

int Quattro9Spreadsheet::getNumSpreadsheets() const
{
	if (m_state->m_spreadsheetMap.empty())
		return 0;
	auto it=m_state->m_spreadsheetMap.end();
	--it;
	return it->first+1;
}

Vec2f Quattro9Spreadsheet::getPosition(int sheetId, Vec2i const &cell) const
{
	auto it=m_state->m_spreadsheetMap.find(sheetId);
	if (it==m_state->m_spreadsheetMap.end() || !it->second)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::getPosition: can not find the sheet %d\n", sheetId));
		return Vec2f(float(cell[0]*50), float(cell[1]*13));
	}
	return it->second->getPosition(cell);
}

librevenge::RVNGString Quattro9Spreadsheet::getSheetName(int id) const
{
	return m_state->getSheetName(id);
}

void Quattro9Spreadsheet::addDLLIdName(int id, librevenge::RVNGString const &name, bool func1)
{
	m_state->m_formulaManager.addDLLIdName(id, name, func1);
}

void Quattro9Spreadsheet::addUserFormat(int id, librevenge::RVNGString const &name)
{
	if (name.empty())
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::addUserFormat: called with empty name for id=%d\n", id));
		return;
	}
	if (m_state->m_idToUserFormatMap.find(id)!=m_state->m_idToUserFormatMap.end())
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::addUserFormat: called with dupplicated id=%d\n", id));
	}
	else
		m_state->m_idToUserFormatMap[id]=name;
}

void Quattro9Spreadsheet::addDocumentStrings(std::shared_ptr<WPSStream> const &stream,
                                             std::vector<Quattro9ParserInternal::TextEntry> const &entries)
{
	if (!m_state->m_documentStrings.second.empty())
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::addDocumentStrings: the entries list is not empty\n"));
	}
	else
	{
		m_state->m_documentStrings.first=stream;
		m_state->m_documentStrings.second=entries;
	}
}

////////////////////////////////////////////////////////////
// low level

////////////////////////////////////////////////////////////
//   parse sheet data: file zones
////////////////////////////////////////////////////////////
bool Quattro9Spreadsheet::readCellStyles(std::shared_ptr<WPSStream> const &stream)
{
	int const vers=version();
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::readU16(input));
	bool const bigBlock=type&0x8000;
	if ((type&0x7fff) !=0xa)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellStyles: not a cell property\n"));
		return false;
	}
	long sz = bigBlock ? long(libwps::readU32(input)) : long(libwps::readU16(input));
	long N= long(libwps::readU32(input));
	f << "N=" << N << ",";
	// qpw9[v=2001]: sz=30, qpwX2[v=2013] and qpwX9[v=2020]: sz=36, we can assume that sz=36 for X3..X8
	int dataSz= vers<2012 ? 30 : 36;
	if (N>=0 && sz>=4 && (sz-4)/dataSz!=N)
	{
		if (vers==2001 || (vers>=2013 && vers<=2020)) // we are "sure" of the data size, try to modify N
		{
			// this situation can happen when QP9 writes data in consecutive blocks(when needed), in this case, N is bad
			if (sz%dataSz==4)
			{
				N=(sz-4)/dataSz;
				WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellStyles: reset N to %ld\n", N));
			}
		}
		else // the data size is unknown, try to modify the data size
		{
			dataSz=int((sz-4)/N);
			if (dataSz>=30)
			{
				WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellStyles: reset data size to %d\n", dataSz));
			}
		}
	}
	if (N<0 || sz<4 || dataSz<30 || (sz-4)/dataSz!=N)
	{
		f << "###";
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellStyles: not a cell property\n"));
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return false;
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	long actSize=long(m_state->m_stylesList.size());
	m_state->m_stylesList.resize(size_t(actSize+N));
	for (long i=actSize; i<actSize+N; ++i)
	{
		pos=input->tell();
		f.str("");
		f << "Cell[style-Ce" << i << "]:";
		auto &style=m_state->m_stylesList[size_t(i)];
		auto fId=int(libwps::readU16(input));
		WPSFont font;
		if (fId)
		{
			if (!m_mainParser.getFont(fId-1, font))
				f << "###";
			f << "F" << fId-1 << ",";
		}
		style.m_fileFormat=int(libwps::readU16(input));
		f << "form=" << std::hex << style.m_fileFormat << std::dec << ",";
		int flag=int(libwps::readU16(input));
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
		case 6: // useme
		{
			style.setHAlignment(WPSCellFormat::HALIGN_CENTER);
			style.m_alignAcrossColumn=true;
			f << "center[across],";
			break;
		}
		case 7: // useme
			f << "ident,";
			break;
		default:
			WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellStyles: find unexpected alignment\n"));
			f << "###align=" << (flag&7) << ",";
			break;
		case 0: // standart
			break;
		}
		switch ((flag>>4)&0x3)
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
			WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellStyle: find unexpected alignment\n"));
			f << "###valign=3,";
			break;
		}
		if (flag&0x80)
		{
			style.setTextRotation(270);
			f << "top[down],";
		}
		// flag&0x100 // maybe related to rotation
		if (flag&0x400)
		{
			style.setWrapping(WPSCellFormat::WRAP_WRAP);
			f << "wrap,";
		}
		flag &= 0xfb08;
		if (flag) f << "fl=" << std::hex << flag << std::dec << ",";
		int val=int(libwps::readU16(input));
		if (val&1) f << "not[protected],";
		val&=0xfffe;
		if (val) f << "f1=" << std::hex << val << std::dec << ",";
		val=int(libwps::readU16(input)); // useme
		if (val) f << "style[id]=" << val << ",";
		val=int(libwps::readU16(input));
		WPSColor color;
		if (!m_mainParser.getColor(val, color))
			f << "##color[font]=" << val << ",";
		else if (!color.isBlack())
		{
			font.m_color=color;
			f << "color[font]=" << color << ",";
		}
		style.setFont(font);
		val=int(libwps::readU8(input));
		if (val)
		{
			f << "rot=" << val << ",";
			style.setTextRotation(val);
		}
		val=int(libwps::readU16(input));
		if (val!=0x50)   // 80=US, 82=Pt
		{
			f << "lang=" << val << ",";
		}
		int borderTypes[5];
		for (auto &t : borderTypes) t=int(libwps::readU8(input));
		int borderColors[5];
		for (auto &c : borderColors) c=int(libwps::readU8(input));
		WPSColor surfColors[2]= {WPSColor::white(), WPSColor::black()};
		for (int s=0; s<2; s++)
		{
			int c=int(libwps::readU8(input));
			if (!m_mainParser.getColor(c, color))
			{
				f << "###col" << s << "=" << c << ",";
				continue;
			}
			surfColors[s]=color;
			if ((s==0 && !color.isWhite()) || (s==1 && !color.isBlack()))
				f << "col" << s << "=" << color << ",";
		}
		int patId=int(libwps::readU8(input));

		auto flags=libwps::readU8(input);
		// unsure, these flags seem set when a file is modified but
		// not when a file is imported => probably safe to ignore

		//if (flags&1) f << "set[align],";
		//if (flags&2) f << "set[format],";
		//if (flags&4) f << "set[align2],";
		//if (flags&8) f << "set[style],";
		//if (flags&0x10) f << "set[font],";
		//if (flags&0x20) f << "set[border],";
		//if (flags&0x40) f << "set[protection],";
		if (flags&0x80) f << "fl2[80],";
		val=int(libwps::readU8(input)); // 0,1,ff
		if (val) f << "g0=" << val << ",";

		// time to write the border and the background
		for (int b=0; b<5; ++b)
		{
			if (!borderTypes[b]) continue;
			WPSBorder border;
			char const *wh[]= {"bordL","bordT","bordR","bordB","bordall"/* or last, checkme*/};
			switch (borderTypes[b])
			{
			case 1: // normal
				f << wh[b] << ",";
				break;
			case 2:
				border.m_type=WPSBorder::Double;
				f << wh[b] << "=double,";
				break;
			case 3:
				border.m_width=2;
				f << wh[b] << "=w2,";
				break;
			case 4:
				border.m_style=WPSBorder::Dot;
				f << wh[b] << "=dot,";
				break;
			case 5:
				border.m_style=WPSBorder::LargeDot;
				f << wh[b] << "=dot,large,";
				break;
			case 6: // 3-1-1-1-1-1
				border.m_style=WPSBorder::Dash;
				f << wh[b] << "=dash311";
				break;
			case 7: // 3-1-1-1
				border.m_style=WPSBorder::Dash;
				f << wh[b] << "=dash31";
				break;
			case 8:
				border.m_style=WPSBorder::Dash;
				f << wh[b] << "=dash";
				break;
			case 9: // 2-2
				border.m_width=2;
				border.m_style=WPSBorder::LargeDot;
				f << wh[b] << "=dot,w2,";
				break;
			case 10: // 3-1-1-1-1-1
				border.m_width=2;
				border.m_style=WPSBorder::Dash;
				f << wh[b] << "=dash311,w2";
				break;
			case 11: // 3-1-1-1
				border.m_width=2;
				border.m_style=WPSBorder::Dash;
				f << wh[b] << "=dash31,w2";
				break;
			case 12: // 3-1
				border.m_width=2;
				border.m_style=WPSBorder::Dash;
				f << wh[b] << "=dash,w2";
				break;
			default:
				WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellStyle: unknown border type\n"));
				break;
			}
			if (!m_mainParser.getColor(borderColors[b], border.m_color))
				f << ",##col=" << borderColors[b];
			else if (!border.m_color.isBlack())
				f << border.m_color;
			f << ",";
			if (b==4) continue;
			int const which[]= {WPSBorder::LeftBit, WPSBorder::TopBit, WPSBorder::RightBit, WPSBorder::BottomBit};
			style.setBorders(which[b], border);
		}
		if (patId==1)
			style.setBackgroundColor(surfColors[0]);
		else if (patId)
		{
			WPSGraphicStyle::Pattern pattern;
			if (!m_mainParser.getPattern(patId, pattern))
				f << "###";
			else
			{
				pattern.m_colors[0]=surfColors[1];
				pattern.m_colors[1]=surfColors[0];
				if (pattern.getAverageColor(color))
					style.setBackgroundColor(color);
			}
			f << "pat=" << patId << ",";
		}
		else
			f << "pat=none,";
		if (input->tell()!=pos+dataSz)
			ascFile.addDelimiter(input->tell(),'|');
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		input->seek(pos+dataSz, librevenge::RVNG_SEEK_SET);
	}
	return true;
}

////////////////////////////////////////////////////////////
//   parse sheet data: document zones
////////////////////////////////////////////////////////////
bool Quattro9Spreadsheet::readDocumentFormulas(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type=int(libwps::readU16(input));
	if ((type&0x7fff) !=0x408)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readDocumentFormulas: not a spreadsheet zone\n"));
		return false;
	}
	long sz = (type&0x8000) ? long(libwps::readU32(input)) : long(libwps::readU16(input));
	int const headerSize = (type&0x8000) ? 6 : 4;
	long endPos=pos+headerSize+sz;
	long N=long(libwps::readU16(input));
	if (sz<12 || (sz-headerSize-8)/4<N || !stream->checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readDocumentFormulas: the size seems bad\n"));
		return false;
	}
	f << "N=" << N << ",";
	int val=int(libwps::readU16(input));
	if (val) f << "f0=" << val << ",";
	f << "f1=" << libwps::readU32(input) << ",";
	f << "f2=" << libwps::readU32(input) << ",";
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	m_state->m_documentFormulas.first = stream;
	auto &entriesList=m_state->m_documentFormulas.second;
	if (!entriesList.empty())
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readDocumentFormulas: oops, we have already some formula\n"));
		entriesList.clear();
	}
	entriesList.reserve(size_t(N));
	for (long i=0; i<N; ++i)
	{
		pos=input->tell();
		f.str("");
		f << "Document[formula-Fo" << i+1 << "]:";
		int dSz=int(libwps::readU16(input));
		if (pos+4+dSz>endPos)
		{
			input->seek(pos, librevenge::RVNG_SEEK_SET);
			break;
		}
		WPSEntry entry;
		entry.setBegin(pos);
		entry.setLength(4+dSz);
		entriesList.push_back(entry);
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		input->seek(pos+4+dSz, librevenge::RVNG_SEEK_SET);
	}
	if (input->tell()!=endPos)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readDocumentFormulas: can not read some formulas\n"));
		ascFile.addPos(input->tell());
		ascFile.addNote("Document[formula]:###extra");
	}
	return true;
}

////////////////////////////////////////////////////////////
//   parse sheet data: spreadsheet zones
////////////////////////////////////////////////////////////
bool Quattro9Spreadsheet::readBeginSheet(std::shared_ptr<WPSStream> const &stream, int &sheetId)
{
	if (m_state->m_actualSpreadsheet)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readBeginSheet: the last spreadsheet is not closed\n"));
	}
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if (type !=0x601)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readBeginSheet: not a spreadsheet zone\n"));
		return false;
	}
	long sz = libwps::readU16(input);
	long endPos=pos+4+sz;
	if (sz<22 || !stream->checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readBeginSheet: the size seems bad\n"));
		return false;
	}
	sheetId=int(libwps::readU16(input));
	f << "id=" << sheetId << ",";
	auto &sheet=m_state->m_actualSpreadsheet;
	sheet.reset(new Quattro9SpreadsheetInternal::Spreadsheet(sheetId));
	int dim[2];
	for (auto &d : dim) d=int(libwps::readU16(input));
	f << "cols[window?]=" << Vec2i(dim[0],dim[1]) << ","; // or 255-0 empty?
	int val=int(libwps::read16(input)); // 0,1,-1
	if (val) f << "f0=" << val << ",";
	for (auto &d : dim) d=int(libwps::readU16(input));
	f << "rows[window?]=" << Vec2i(dim[0],dim[1]) << ",";
	for (int i=0; i<3; ++i)
	{
		val=int(libwps::read16(input));
		int const expected[]= {0,-1,0};
		if (val!=expected[i]) f << "f" << i+1 << "=" << val << ",";
	}
	Quattro9ParserInternal::TextEntry entry;
	if (m_mainParser.readPString(stream, endPos, entry))
	{
		f << entry.getDebugString(stream) << ",";
		auto name=entry.getString(stream);
		if (!name.empty())
		{
			if (m_state->m_idToSheetNameMap.find(sheetId)==m_state->m_idToSheetNameMap.end())
				m_state->m_idToSheetNameMap[sheetId]=name;
		}
	}
	else
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readBeginSheet: can not read the spreadsheet name\n"));
		f << "###";
	}
	if (sheetId>=1024)
	{
		// checkme: what is the maximum of the number of sheet...
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readBeginSheet: id=%d seems to big\n", sheetId));
	}
	else if (m_state->m_spreadsheetMap.find(sheetId)!=m_state->m_spreadsheetMap.end())
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readBeginSheet: id=%d sheet already exists\n", sheetId));
		f << "###id,";
	}
	else
		m_state->m_spreadsheetMap[sheetId]=sheet;
	if (input->tell()!=endPos)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readBeginSheet: find extra data\n"));
		f << "###";
		ascFile.addDelimiter(input->tell(),'|');
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool Quattro9Spreadsheet::readEndSheet(std::shared_ptr<WPSStream> const &stream)
{
	if (!m_state->m_actualSpreadsheet)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readEndSheet: no spreadsheet are opened\n"));
	}
	else if (m_state->m_actualColumn>=0)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readEndSheet: the last column is not closed\n"));
		m_state->m_actualColumn=-1;
	}
	m_state->m_actualSpreadsheet.reset();
	long filePos[2]; // f0=pointer to a previous zone 601|602
	Quattro9Parser::readFilePositions(stream, filePos);
	return true;
}

bool Quattro9Spreadsheet::readColRowDefault(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if (type !=0x631 && type != 0x632)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readColRowDefault: not a dimension zone\n"));
		return false;
	}
	long sz = libwps::readU16(input);
	long endPos=pos+4+sz;
	if (sz!=2 || !stream->checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readColRowDefault: unexpected size\n"));
		return false;
	}
	int val=int(libwps::readU16(input)); // height/width in dpi?
	if (val&0x8000)
		f << "size=" << (val&0x7FFF) << ","; // exact
	else
		f << "size=" << val << "*,"; // automatic update
	if (!m_state->m_actualSpreadsheet)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readColRowDimension: can not find the spreadsheet\n"));
	}
	else
		m_state->m_actualSpreadsheet->m_defaultSizes[type==0x631 ? 1 : 0]=(val&0x7fff);
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool Quattro9Spreadsheet::readColRowDimension(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if (type !=0x633 && type != 0x634)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readColRowDimension: not a dimension zone\n"));
		return false;
	}
	long sz = libwps::readU16(input);
	long endPos=pos+4+sz;
	if (sz!=6 || !stream->checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readColRowDimension: unexpected size\n"));
		return false;
	}
	int posi=int(libwps::readU32(input));
	f << "id=" << posi << ",";
	int val=int(libwps::readU16(input)); // height/width in dpi?
	if (val&0x8000)
		f << "size=" << (val&0x7FFF) << ","; // exact
	else
		f << "size=" << val << "*,"; // automatic update
	if (!m_state->m_actualSpreadsheet)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readColRowDimension: can not find the spreadsheet\n"));
	}
	else
		m_state->m_actualSpreadsheet->setColRowSize(type==0x633 ? 1 : 0, posi, (val&0x7fff));
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool Quattro9Spreadsheet::readColRowDimensions(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if (type !=0x635 && type != 0x636)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readColRowDimensions: not a dimension zone\n"));
		return false;
	}
	long sz = libwps::readU16(input);
	long endPos=pos+4+sz;
	if (sz!=10 || !stream->checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readColRowDimension: unexpected size\n"));
		return false;
	}
	int dim[2];
	for (auto &d : dim) d=int(libwps::readU32(input));
	f << "limits=" << Vec2i(dim[0],dim[1]) << ",";
	int val=int(libwps::readU16(input)); // height/width in dpi?
	if (val&0x8000)
		f << "size=" << (val&0x7FFF) << ","; // exact
	else
		f << "size=" << val << "*,"; // automatic update
	if (!m_state->m_actualSpreadsheet)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readColRowDimension: can not find the spreadsheet\n"));
	}
	else if (dim[0]<=dim[1])
		m_state->m_actualSpreadsheet->setColRowSizes(type==0x635 ? 1 : 0, dim[0], dim[1], (val&0x7fff));
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool Quattro9Spreadsheet::readMergedCells(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if (type !=0x61d)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readMergedCells: not a merged cells zone\n"));
		return false;
	}
	long sz = libwps::readU16(input);
	if (sz!=16)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readMergedCells: unexpected size\n"));
		return false;
	}
	int dim[4];
	for (auto &d : dim) d=int(libwps::readU32(input));
	WPSBox2i box(Vec2i(dim[0],dim[2]),Vec2i(dim[1],dim[3]));
	if (dim[0]<0 || dim[0]>dim[1] || dim[2]<0 || dim[2]>dim[3])
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readMergedCells: the selection seems bad\n"));
		f << "###";
	}
	else if (m_state->m_actualSpreadsheet)
		m_state->m_actualSpreadsheet->m_mergedCellList.push_back(box);
	else
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readMergedCells: can not find the main cell\n"));
	}
	f << box << ",";
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool Quattro9Spreadsheet::readPageBreak(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if (type !=0x617)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readPageBreak: not a pagebreak zone\n"));
		return false;
	}
	long sz = libwps::readU16(input);
	if (sz<2 || (sz%2)!=0)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readPageBreak: unexpected size\n"));
		return false;
	}
	int val=int(libwps::read16(input)); // 800=col?, a00=row?
	if (val) f << "fl=" << std::hex << val << std::dec << ",";
	int N=int((sz-2)/2);
	f << "break=[";
	for (int i=0; i<N; ++i) f << libwps::readU16(input) << ",";
	f << "],";

	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

////////////////////////////////////////////////////////////
//   parse sheet data: column zones
////////////////////////////////////////////////////////////
bool Quattro9Spreadsheet::readBeginColumn(std::shared_ptr<WPSStream> const &stream)
{
	if (!m_state->m_actualSpreadsheet)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readBeginColumn: called outside a spreadsheet\n"));
	}
	else if (m_state->m_actualColumn>=0)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readBeginColumn: the last column is not closed\n"));
	}
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if (type !=0xa01)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readBeginColumn: not a col[begin] zone\n"));
		return false;
	}
	long sz = libwps::readU16(input);
	long endPos=pos+4+sz;
	if (sz!=10 || !stream->checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readBeginColumn: unexpected size\n"));
		return false;
	}
	int col=int(libwps::readU16(input));
	f << "col=" << col << ",";
	if (m_state->m_actualSpreadsheet) m_state->m_actualColumn=col;
	int dim[2];
	for (auto &d : dim) d=int(libwps::readU32(input));
	f << "limits=" << Vec2i(dim[0],dim[1]) << ",";
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool Quattro9Spreadsheet::readEndColumn(std::shared_ptr<WPSStream> const &stream)
{
	if (m_state->m_actualColumn<0)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readEndColumn: the last column is not opened\n"));
	}
	m_state->m_actualColumn=-1;
	long filePos[2]; // f0=pointer to a previous zone a01|a02
	Quattro9Parser::readFilePositions(stream, filePos);
	return true;
}

////////////////////////////////////////////////////////////
//   parse sheet data: cell zones
////////////////////////////////////////////////////////////
bool Quattro9Spreadsheet::readCellList(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if (type !=0xc01)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellList: not a cell zone\n"));
		return false;
	}
	long sz = libwps::readU16(input);
	long endPos=pos+4+sz;
	if (sz<8 || !stream->checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellList: the size seems bad\n"));
		return false;
	}
	int row=int(libwps::readU32(input));
	if (row) f << "first[row]=" << row << ",";
	int nCells=int(libwps::readU32(input));
	if (nCells) f << "num[cells]=" << nCells << ",";
	int lastRow=row+nCells;
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	int const col= m_state->m_actualSpreadsheet ? m_state->m_actualColumn : -1;
	if (col < 0)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellList: call outside a col,begin zone\n"));
	}
	Quattro9SpreadsheetInternal::Column invalidColumn;
	Quattro9SpreadsheetInternal::Column &column = col>=0 ? m_state->m_actualSpreadsheet->getColumn(col) : invalidColumn;
	while (input->tell() < endPos)
	{
		pos=input->tell();
		f.str("");
		f << "Spreadsheet[cell]:";
		Quattro9SpreadsheetInternal::CellData cell;
		int cType=cell.m_type=int(libwps::readU8(input));
		bool ok=true;
		if (cType&0x80)
		{
			ok=(input->tell()+2<=endPos);
			if (ok)
			{
				cell.m_style=int(libwps::readU16(input));
				f << "Ce" << cell.m_style-1 << ",";
			}
			cType &= 0x7f;
		}
		int numData=1, N=1;
		if (ok)
		{
			if ((cType&0x60)==0x40)   // list
			{
				ok=(input->tell()+2<=endPos);
				if (ok)
				{
					N=numData=int(libwps::readU16(input));
					f << "N=" << numData << ",";
				}
				cType &= 0x1f;
			}
			else if ((cType&0x60)==0x60)   // first value followed by increment value
			{
				ok=(input->tell()+2<=endPos);
				if (ok)
				{
					N=int(libwps::readU16(input));
					f << "serie, N=" << N << ",";
					numData=2;
				}
				cType &= 0x1f;
			}
			else if ((cType&0x60)==0x20)
			{
				/* CHECKME what is it ?
				   I tried to generate a list with 2|(n>2) similar values
				   but it generates a list :-~
				*/
				WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellList: argh, list[cType]=0x20, some cells will be lost\n"));
				ok=false;
			}
		}
		if (col>=0)
		{
			if (N>1)
				f << "C" << col << "R" << row << "-" << row+N-1 << ",";
			else
				f << "C" << col << "R" << row << ",";
		}
		cell.m_rows=Vec2i(row, row+N-1);
		row += N;
		f << "type=" << std::hex << cType << std::dec << ",";
		if (ok)
		{
			long actPos=input->tell();
			ok=false;
			switch (cType)
			{
			case 1: // no data?
				ok=true;
				break;
			case 0x2: // unsigned int value
				if (actPos+numData*2>endPos) break;
				f << "values=[";
				for (int i=0; i<numData; ++i)
				{
					cell.m_intList.push_back(int(libwps::readU16(input)));
					f << cell.m_intList.back() << ",";
				}
				f << "],";
				ok=true;
				break;
			case 0x3: // signed int
				if (actPos+numData*2>endPos) break;
				f << "values=[";
				for (int i=0; i<numData; ++i)
				{
					cell.m_intList.push_back(int(libwps::read16(input)));
					f << cell.m_intList.back() << ",";
				}
				f << "],";
				ok=true;
				break;
			case 4: // double4 value
				if (actPos+4*numData>endPos) break;
				f << "values=[";
				for (int i=0; i<numData; ++i)
				{
					double value;
					bool isNaN;
					if (libwps::readDouble4(input, value, isNaN))
						f << value << ",";
					else
					{
						value=0;
						f << "###,";
						input->seek(actPos+(i+1)*8, librevenge::RVNG_SEEK_SET);
					}
					cell.m_doubleList.push_back(value);
				}
				f << "],";
				ok=true;
				break;
			case 5:   // double
			{
				if (actPos+numData*8>endPos) break;
				f << "values=[";
				for (int i=0; i<numData; ++i)
				{
					double value;
					bool isNaN;
					if (libwps::readDouble8(input, value, isNaN))
						f << value << ",";
					else
					{
						value=0;
						f << "###,";
						input->seek(actPos+(i+1)*8, librevenge::RVNG_SEEK_SET);
					}
					cell.m_doubleList.push_back(value);
				}
				f << "],";
				ok=true;
				break;
			}
			case 7: // string index
				if (actPos+4*numData>endPos) break;
				f << "values=[";
				for (int i=0; i<numData; ++i)   // the ith string
				{
					cell.m_intList.push_back(int(libwps::readU32(input)));
					f << "Str" << cell.m_intList.back() << ",";
				}
				f << "],";
				ok=true;
				break;
			case 8: // formula
				if (actPos+14*numData>endPos) break;
				f << "values=[";
				for (int i=0; i<numData; ++i)
				{
					double value;
					bool isNaN;
					if (libwps::readDouble8(input, value, isNaN))
						f << value << ",";
					else
					{
						f << "###,";
						input->seek(actPos+i*14+8, librevenge::RVNG_SEEK_SET);
						value=0;
					}
					cell.m_doubleList.push_back(value);
					cell.m_flagList.push_back(int(libwps::readU16(input))); // high byte=0-3, lower byte=0|8|10
					if (cell.m_flagList.back()) f << "fl=" << std::hex << cell.m_flagList.back() << std::dec << ",";
					cell.m_intList.push_back(int(libwps::readU32(input)));
					f << "Fo" << cell.m_intList.back() << ",";
				}
				f << "],";
				ok=true;
				break;
			default:
				WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellList: argh, find unknown type %d, some cells will be lost\n", int(cType)));
				break;
			}
		}
		if (!ok)
		{
			input->seek(pos, librevenge::RVNG_SEEK_SET);
			break;
		}
		if ((cell.m_type&0x9f)!=1) // no need to add cell with no style and no content
			column.add(cell.m_rows, cell);
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
	}
	if (input->tell() < endPos)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellList: find extra data\n"));
		ascFile.addPos(input->tell());
		ascFile.addNote("Spreadsheet[cell]:###");
	}
	else if (lastRow!=row)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellList: read an unexpected number of cells %d != %d\n", row, lastRow));
	}
	input->seek(endPos, librevenge::RVNG_SEEK_SET);
	return true;
}

bool Quattro9Spreadsheet::readCellResult(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if (type !=0xc02)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellResult: not a result zone\n"));
		return false;
	}
	long sz = libwps::readU16(input);
	long endPos=pos+4+sz;
	if (sz<10 || !stream->checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellResult: the size seems bad\n"));
		return false;
	}

	int col=int(libwps::readU16(input));
	int row=int(libwps::readU32(input));
	if (col!=m_state->m_actualColumn)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellResult: unexpected called\n"));
		f << "###";
	}
	f << "C" << col << "R" << row << ",";
	Quattro9ParserInternal::TextEntry entry;
	if (m_mainParser.readPString(stream, endPos, entry))
	{
		f << entry.getDebugString(stream) << ",";
		if (m_state->m_actualSpreadsheet && col>=0)
			m_state->m_actualSpreadsheet->getColumn(col).add(row,stream,entry);
	}
	else
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellResult: can not read a string\n"));
		f << "###";
	}

	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}
////////////////////////////////////////////////////////////
// formula
////////////////////////////////////////////////////////////
bool Quattro9Spreadsheet::readCell
(std::shared_ptr<WPSStream> const &stream, Vec2i actPos, WKSContentListener::FormulaInstruction &instr, int sheetId, librevenge::RVNGString const &fName)
{
	RVNGInputStreamPtr input = stream->m_input;
	instr=WKSContentListener::FormulaInstruction();
	instr.m_type=WKSContentListener::FormulaInstruction::F_Cell;
	instr.m_fileName=fName;
	int pos[3]; // col, sheet, fl|row
	bool relative[3] = { false, false, false};
	for (auto &p : pos) p=int(libwps::readU16(input));
	int fl=int(libwps::readU16(input));
	if (fl&0x8000)
	{
		pos[1] = int(uint16_t(pos[1]+sheetId)); // unsure what is the maximum
		relative[1] = true;
	}
	if (fl&0x4000)
	{
		pos[0] = int(uint16_t(pos[0]+actPos[0]));
		relative[0] = true;
	}
	if (fl&0x2000)
	{
		pos[2] = int(uint16_t(pos[2]+actPos[1]));
		relative[2] = true;
	}
	//fl&0x1fff=0 | 1fff depending on the sign of pos[2] probably junk
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
	return true;
}

bool Quattro9Spreadsheet::readCellReference(std::shared_ptr<WPSStream> const &stream, long endPos,
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
		// type==4:
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellReference: find a cell collection 4\n"));
		return false;
	}
	WKSContentListener::FormulaInstruction instr;
	if (cellType==3)
	{
		// checkme
		int dataSize=(type&0x3ff);
		if (pos+2+dataSize>endPos)
		{
			WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellReference: can not read the cell collection data size\n"));
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
				WPS_DEBUG_MSG(("Quattro9Spreadsheet::readCellReference: can not read a cell\n"));
				return false;
			}
			for (auto const &c : cells.m_cells) ref.addInstruction(c);
		}
		return true;
	}
	int const expectedSize[]= {8,16,2};
	if (pos+2+expectedSize[cellType]>endPos) return false;
	// checkme: what means type&0x800?
	if (type&0x400)
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
	if (cellType==0 && pos+10<=endPos)
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
	else if (cellType==1 && pos+18<=endPos)
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
void Quattro9Spreadsheet::sendSpreadsheet(int sId)
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::sendSpreadsheet: I can not find the listener\n"));
		return;
	}
	std::shared_ptr<Quattro9SpreadsheetInternal::Spreadsheet> sheet;
	auto sheetIt=m_state->m_spreadsheetMap.find(sId);
	if (sheetIt==m_state->m_spreadsheetMap.end())
		sheet.reset(new Quattro9SpreadsheetInternal::Spreadsheet(sId));
	else
		sheet=sheetIt->second;
	sheet->update();
	m_listener->openSheet(sheet->getWidths(), m_state->getSheetName(sId));
	m_mainParser.sendPageGraphics(sId);
	std::set<int> listRows=sheet->getListSendRow();
	for (auto rIt=listRows.begin(); rIt!=listRows.end(); ++rIt)
	{
		int row=*rIt;
		auto nextRIt=rIt;
		int numRow=(++nextRIt==listRows.end()) ? 1 : *nextRIt-row;
		m_listener->openSheetRow(sheet->getRowHeight(row), numRow);
		std::vector<Vec2i> cols=sheet->getListCellsInARow(row);
		for (auto const &c : cols)
		{
			Vec2i pos(c[0],row);
			sendCellContent(sheet->getCell(pos), pos, sId, 1+c[1]-c[0]);
		}
		m_listener->closeSheetRow();
	}
	m_listener->closeSheet();
}

void Quattro9Spreadsheet::sendCellContent(Quattro9SpreadsheetInternal::CellData const *cell, Vec2i pos, int sheetId, int numRepeated)
{
	if (!cell) return;
	if (m_listener.get() == nullptr)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::sendCellContent: I can not find the listener\n"));
		return;
	}
	Quattro9SpreadsheetInternal::Cell finalCell;
	WKSContentListener::CellContent content;
	finalCell.setPosition(pos);
	finalCell.setNumSpannedCells(cell->m_span);
	WPSFont font;
	libwps_tools_win::Font::Type fontType=m_mainParser.getDefaultFontType();
	if (cell->m_style>0 && cell->m_style<=int(m_state->m_stylesList.size()))
	{
		auto const &style=m_state->m_stylesList[size_t(cell->m_style-1)];
		static_cast<WPSCellFormat &>(finalCell)=style;
		if (style.m_fileFormat>0) finalCell.updateFormat(style.m_fileFormat);
		font=style.getFont();
		if (!font.m_name.empty())
		{
			fontType=libwps_tools_win::Font::getFontType(font.m_name);
			if (fontType==libwps_tools_win::Font::UNKNOWN)
				fontType=m_mainParser.getDefaultFontType();
			else
				finalCell.setFont(font);
		}
	}
	else if (cell->m_style>0)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::sendCellContent: unknown style %d\n", cell->m_style-1));
	}
	m_listener->setFont(font);
	int stringId=-1;
	switch ((cell->m_type&0x1f))
	{
	case 1: // none
		break;
	case 2:
	case 3:
		content.m_contentType=WKSContentListener::CellContent::C_NUMBER;
		content.setValue(cell->getInt(pos[1]));
		break;
	case 4:
	case 5:
		content.m_contentType=WKSContentListener::CellContent::C_NUMBER;
		content.setValue(cell->getDouble(pos[1]));
		break;
	case 7:
		content.m_contentType=WKSContentListener::CellContent::C_TEXT;
		stringId=cell->getInt(pos[1]);
		break;
	case 8:   // changeme
	{
		content.m_contentType=WKSContentListener::CellContent::C_NUMBER;
		content.setValue(cell->getDouble(pos[1]));
		int fId=cell->getInt(pos[1]);
		if (fId>0 && fId<=int(m_state->m_documentFormulas.second.size()) &&
		        m_state->m_documentFormulas.first)
		{
			auto const &entry=m_state->m_documentFormulas.second[size_t(fId-1)];
			auto &stream=m_state->m_documentFormulas.first;
			auto &input=stream->m_input;
			libwps::DebugFile &ascFile=stream->m_ascii;
			libwps::DebugStream f;
			long actPos=input->tell();
			std::string error;
			input->seek(entry.begin(), librevenge::RVNG_SEEK_SET);
			bool formulaOk=m_state->m_formulaManager.readFormula(stream, entry.end(), pos, sheetId, content.m_formula, error);
			for (auto const &fo : content.m_formula)
				f << fo;
			f << "," << error;
			if (formulaOk)
			{
				content.m_contentType=WKSContentListener::CellContent::C_FORMULA;
				for (auto &fo : content.m_formula)
				{
					if (fo.m_type==WKSContentListener::FormulaInstruction::F_Cell ||
					        fo.m_type==WKSContentListener::FormulaInstruction::F_CellList)
					{
						int dim=fo.m_type==WKSContentListener::FormulaInstruction::F_Cell ? 1 : 2;
						for (int i=0; i<dim; ++i)
						{
							if (fo.m_sheetId[i]>=0 && fo.m_sheetName[i].empty() && (fo.m_sheetId[i]!=sheetId || !fo.m_fileName.empty()))
								fo.m_sheetName[i]=getSheetName(fo.m_sheetId[i]);
						}
						continue;
					}
					if (fo.m_type!=WKSContentListener::FormulaInstruction::F_Text)
						continue;
					std::string &text=fo.m_content;
					librevenge::RVNGString finalString=libwps_tools_win::Font::unicodeString(text, fontType);
					if (finalString.empty())
						text.clear();
					else
						text=finalString.cstr();
				}
			}
			else
				content.m_formula.clear();
			ascFile.addPos(entry.begin());
			ascFile.addNote(f.str().c_str());
			input->seek(actPos, librevenge::RVNG_SEEK_SET);
		}
		break;
	}
	default:
		break;
	}
	m_listener->openSheetCell(finalCell, content, numRepeated);
	if (stringId>0 && stringId<=int(m_state->m_documentStrings.second.size()) &&
	        m_state->m_documentStrings.first)
		m_state->m_documentStrings.second[size_t(stringId-1)].send(m_state->m_documentStrings.first, font, fontType, m_listener);
	else if (stringId>0)
	{
		WPS_DEBUG_MSG(("Quattro9Spreadsheet::sendCellContent: can not find the string %d\n", stringId));
	}
	m_listener->closeSheetCell();
}
////////////////////////////////////////////////////////////
// cell reference
////////////////////////////////////////////////////////////

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
