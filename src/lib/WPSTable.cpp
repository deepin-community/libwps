/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: t; c-basic-offset: 4 -*- */
/* libwps
 * Version: MPL 2.0 / LGPLv2.1+
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *
 * Major Contributor(s):
 * Copyright (C) 2009, 2011 Alonso Laurent (alonso@loria.fr)
 * Copyright (C) 2006, 2007 Andrew Ziem
 * Copyright (C) 2004-2006 Fridrich Strba (fridrich.strba@bluewin.ch)
 * Copyright (C) 2004 Marc Maurer (uwog@uwog.net)
 * Copyright (C) 2003-2005 William Lachance (william.lachance@sympatico.ca)
 *
 * For minor contributions see the git repository.
 *
 * Alternatively, the contents of this file may be used under the terms
 * of the GNU Lesser General Public License Version 2.1 or later
 * (LGPLv2.1+), in which case the provisions of the LGPLv2.1+ are
 * applicable instead of those above.
 *
 * For further information visit http://libwps.sourceforge.net
 */

/*
 * Structure to store and construct a table from an unstructured list
 * of cell
 *
 */

#include <iostream>
#include <set>

#include "WPSContentListener.h"
#include "WPSCell.h"

#include "WPSTable.h"

////////////////////////////////////////////////////////////
// WPSColumnFormat
std::ostream &operator<<(std::ostream &o, WPSColumnFormat const &column)
{
	if (column.m_width>=0)
	{
		if (column.m_isPercentWidth)
			o<<"w=" << column.m_width << "%,";
		else
			o<<"w=" << column.m_width << ",";
	}
	if (column.m_useOptimalWidth) o << "optimal[h],";
	if (column.m_isHeader) o << "table[header],";
	if (column.m_numRepeat>1) o << "repeat=" << column.m_numRepeat << ",";
	return o;
}

void WPSColumnFormat::addTo(librevenge::RVNGPropertyList &propList) const
{
	if (m_width>=0)
		propList.insert("style:column-width", double(m_width), m_isPercentWidth ? librevenge::RVNG_PERCENT : librevenge::RVNG_POINT);
	if (m_useOptimalWidth)
		propList.insert("style:use-optimal-column-width", true);
	if (m_isHeader)
		propList.insert("librevenge:is-header-column", true); // checkme
	if (m_numRepeat>1)
		propList.insert("table:number-columns-repeated", m_numRepeat);
}

////////////////////////////////////////////////////////////
// WPSRowFormat
std::ostream &operator<<(std::ostream &o, WPSRowFormat const &row)
{
	if (row.m_height>=0)
	{
		if (row.m_isMinimalHeight)
			o<<"h[min]=" << row.m_height << ",";
		else
			o<<"h=" << row.m_height << ",";
	}
	if (row.m_useOptimalHeight) o << "optimal[h],";
	if (row.m_isHeader) o << "table[header],";
	return o;
}

void WPSRowFormat::addTo(librevenge::RVNGPropertyList &propList) const
{
	if (m_height>=0)
	{
		if (m_isMinimalHeight)
			propList.insert("style:min-row-height", double(m_height), librevenge::RVNG_POINT);
		else
			propList.insert("style:row-height", double(m_height), librevenge::RVNG_POINT);
	}
	if (m_useOptimalHeight)
		propList.insert("style:use-optimal-row-height", true);
	propList.insert("librevenge:is-header-row", m_isHeader);
}

////////////////////////////////////////////////////////////
// destructor, ...
WPSTable::~WPSTable()
{
}

void WPSTable::add(WPSCellPtr &cell)
{
	m_cellsList.push_back(cell);
}

WPSCellPtr WPSTable::getCell(int id)
{
	if (id < 0 || id >= int(m_cellsList.size()))
	{
		WPS_DEBUG_MSG(("WPSTable::get: cell %d does not exists\n",id));
		return std::shared_ptr<WPSCell>();
	}
	return m_cellsList[size_t(id)];
}

////////////////////////////////////////////////////////////
// build the table structure
bool WPSTable::buildStructures()
{
	if (m_colsSize.size())
		return true;

	size_t nCells = m_cellsList.size();
	std::vector<float> listPositions[2];
	for (int dim = 0; dim < 2; dim++)
	{
		WPSCell::Compare compareFunction(dim);
		std::set<WPSCell::Compare::Point,
		    WPSCell::Compare> set(compareFunction);
		for (size_t c = 0; c < nCells; c++)
		{
			set.insert(WPSCell::Compare::Point(0, m_cellsList[c].get()));
			set.insert(WPSCell::Compare::Point(1, m_cellsList[c].get()));
		}

		std::vector<float> positions;
		float maxPosiblePos=0;
		int actCell = -1;
		for (auto it : set)
		{
			float pos = it.getPos(dim);
			if (actCell < 0 || pos > maxPosiblePos)
			{
				actCell++;
				positions.push_back(pos);
				maxPosiblePos = pos+2.0f; // 2 pixel ok
			}
			if (it.m_which == 0 && it.getPos(dim)-2.0f < maxPosiblePos)
				maxPosiblePos = pos;
		}
		listPositions[dim] = positions;
	}
	std::vector<int> numYSet(listPositions[1].size(), 0);
	std::vector<int> numYUnset(listPositions[1].size(), 0);
	for (auto &cell : m_cellsList)
	{
		int cellPos[2], spanCell[2];
		for (int dim = 0; dim < 2; dim++)
		{
			float pt[2] = { cell->box().min()[dim],
			                cell->box().max()[dim]
			              };
			std::vector<float> &pos = listPositions[dim];
			size_t numPos = pos.size();
			size_t i = 0;
			while (i+1 < numPos && pos[i+1] < pt[0])
				i++;
			while (i+1 < numPos && (pos[i]+pos[i+1])/2 < pt[0])
				i++;
			if (i+1 > numPos)
			{
				WPS_DEBUG_MSG(("WPSTable::buildStructures: impossible to find cell position !!!\n"));
				return false;
			}
			cellPos[dim] = int(i);
			while (i+1 < numPos && pos[i+1] < pt[1])
				i++;
			if (i+1 < numPos && (pos[i]+pos[i+1])/2 < pt[1])
				i++;
			spanCell[dim] = int(i)-cellPos[dim];
			if (spanCell[dim]==0 &&
			        (cell->box().size()[dim]<0||cell->box().size()[dim]>0))
			{
				WPS_DEBUG_MSG(("WPSTable::buildStructures: impossible to find span number !!!\n"));
				return false;
			}
			if (spanCell[dim] > 1 &&
			        pos[size_t(cellPos[dim])]+2.0f > pos[size_t(cellPos[dim]+1)])
			{
				spanCell[dim]--;
				cellPos[dim]++;
			}
			if (spanCell[dim] > 1 &&
			        pos[size_t(cellPos[dim])]+2.0f > pos[size_t(cellPos[dim]+1)])
			{
				spanCell[dim]--;
				cellPos[dim]++;
			}
		}
		cell->m_position = Vec2i(cellPos[0], cellPos[1]);
		cell->m_numberCellSpanned = Vec2i(spanCell[0], spanCell[1]);
		if (spanCell[1] > 0)
		{
			for (int x = cellPos[0]; x < cellPos[0]+spanCell[0]; x++)
			{
				if (cell->isVerticalSet())
					numYSet[size_t(cellPos[1]+spanCell[1]-1)]++;
				else
					numYUnset[size_t(cellPos[1]+spanCell[1]-1)]++;
			}
		}
	}
	// finally update the row/col size
	for (int dim = 0; dim < 2; dim++)
	{
		std::vector<float> const &pos = listPositions[dim];
		size_t numPos = pos.size();
		if (!numPos) continue;
		std::vector<float> &res = (dim==0) ? m_colsSize : m_rowsSize;
		res.resize(numPos-1);
		for (size_t i = 0; i < numPos-1; i++)
		{
			if (dim==0 || numYUnset[i]==0)
				res[i] = pos[i+1]-pos[i];
			else if (numYSet[i])
				res[i] = -(pos[i+1]-pos[i]);
			else
				res[i] = 0;
		}
	}

	return true;
}

////////////////////////////////////////////////////////////
// try to send the table
bool WPSTable::sendTable(WPSContentListenerPtr listener)
{
	if (!buildStructures())
		return false;
	if (!listener)
		return true;

	size_t numCols = m_colsSize.size();
	size_t numRows = m_rowsSize.size();
	if (!numCols || !numRows)
		return false;
	std::vector<int> cellsId(numCols*numRows, -1);
	size_t nCells = m_cellsList.size();
	for (size_t c = 0; c < nCells; c++)
	{
		if (!m_cellsList[c]) continue;
		Vec2i const &pos=m_cellsList[c]->m_position;
		Vec2i const &span=m_cellsList[c]->m_numberCellSpanned;

		for (int x = pos[0]; x < pos[0]+span[0]; x++)
		{
			if (x >= int(numCols))
			{
				WPS_DEBUG_MSG(("WPSTable::sendTable: x is too big !!!\n"));
				return false;
			}
			for (int y = pos[1]; y < pos[1]+span[1]; y++)
			{
				if (y >= int(numRows))
				{
					WPS_DEBUG_MSG(("WPSTable::sendTable: y is too big !!!\n"));
					return false;
				}
				auto tablePos = size_t(y*int(numCols)+x);
				if (cellsId[tablePos] != -1)
				{
					WPS_DEBUG_MSG(("WPSTable::sendTable: cells is used!!!\n"));
					return false;
				}
				if (x == pos[0] && y == pos[1])
					cellsId[tablePos] = int(c);
				else
					cellsId[tablePos] = -2;
			}
		}
	}

	listener->openTable(m_colsSize, librevenge::RVNG_POINT);
	WPSListenerPtr listen=listener;
	for (size_t r = 0; r < numRows; r++)
	{
		listener->openTableRow(m_rowsSize[r], librevenge::RVNG_POINT);
		for (size_t c = 0; c < numCols; c++)
		{
			size_t tablePos = r*numCols+c;
			int id = cellsId[tablePos];
			if (id == -1)
				listener->addEmptyTableCell(Vec2i(int(c), int(r)));
			if (id < 0)
				continue;

			m_cellsList[size_t(id)]->send(listen);
		}
		listener->closeTableRow();
	}
	listener->closeTable();

	return true;
}


////////////////////////////////////////////////////////////
// try to send the table
bool WPSTable::sendAsText(WPSContentListenerPtr listener)
{
	if (!listener) return true;

	WPSListenerPtr listen=listener;
	for (auto &cell : m_cellsList)
	{
		if (!cell) continue;
		cell->sendContent(listen);
		listener->insertEOL();
	}
	return true;
}
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
