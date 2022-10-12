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

#ifndef WPS_TABLE
#  define WPS_TABLE

#include <iostream>
#include <vector>

#include "libwps_internal.h"

/*
 * Structure to store the column properties
 *
 * \note use only to define sheet properties, to be changed
 */
struct WPSColumnFormat
{
public:
	//! constructor
	explicit WPSColumnFormat(float width=-1)
		: m_width(width)
		, m_isPercentWidth(false)
		, m_useOptimalWidth(false)
		, m_isHeader(false)
		, m_numRepeat(1)
	{
	}
	//! add to the propList
	void addTo(librevenge::RVNGPropertyList &propList) const;
	/** a comparison  function
		\note this comparison function does ignore m_numRepeat
	 */
	int compare(WPSColumnFormat const &col) const
	{
		if (m_width<col.m_width) return 1;
		if (m_width>col.m_width) return -1;
		if (m_isPercentWidth!=col.m_isPercentWidth) return m_isPercentWidth ? 1 : -1;
		if (m_useOptimalWidth!=col.m_useOptimalWidth) return m_useOptimalWidth ? 1 : -1;
		if (m_isHeader!=col.m_isHeader) return m_isHeader ? 1 : -1;
		return 0;
	}
	//! operator==
	bool operator==(WPSColumnFormat const &col) const
	{
		return compare(col)==0;
	}
	//! operator!=
	bool operator!=(WPSColumnFormat const &col) const
	{
		return compare(col)!=0;
	}
	//! operator<
	bool operator<(WPSColumnFormat const &col) const
	{
		return compare(col)<0;
	}
	//! operator<<
	friend std::ostream &operator<<(std::ostream &o, WPSColumnFormat const &col);

	//! the column width, if known
	float m_width;
	//! a flag to know if the width is in percent (or in point)
	bool m_isPercentWidth;
	//! a flag to know if we need to see use-optimal column width
	bool m_useOptimalWidth;
	//! a flag to know if the column is a header column
	bool m_isHeader;
	//! the number times a column is repeated
	int m_numRepeat;
};

/*
 * Structure to store the row properties
 *
 * \note use only to define sheet properties, to be changed
 */
struct WPSRowFormat
{
public:
	//! constructor
	explicit WPSRowFormat(float height=-1)
		: m_height(height)
		, m_isMinimalHeight(false)
		, m_useOptimalHeight(false)
		, m_isHeader(false)
	{
	}
	//! add to the propList
	void addTo(librevenge::RVNGPropertyList &propList) const;
	//! a comparison  function
	int compare(WPSRowFormat const &row) const
	{
		if (m_height<row.m_height) return 1;
		if (m_height>row.m_height) return -1;
		if (m_isMinimalHeight!=row.m_isMinimalHeight) return m_isMinimalHeight ? 1 : -1;
		if (m_useOptimalHeight!=row.m_useOptimalHeight) return m_useOptimalHeight ? 1 : -1;
		if (m_isHeader!=row.m_isHeader) return m_isHeader ? 1 : -1;
		return 0;
	}
	//! operator==
	bool operator==(WPSRowFormat const &row) const
	{
		return compare(row)==0;
	}
	//! operator!=
	bool operator!=(WPSRowFormat const &row) const
	{
		return compare(row)!=0;
	}
	//! operator<
	bool operator<(WPSRowFormat const &row) const
	{
		return compare(row)<0;
	}
	//! operator<<
	friend std::ostream &operator<<(std::ostream &o, WPSRowFormat const &row);

	//! the row height, if known
	float m_height;
	//! a flag to know if the height is only a minimum
	bool m_isMinimalHeight;
	//! a flag to know if we need to see use-optimal row height
	bool m_useOptimalHeight;
	//! a flag to know if the row is a header row
	bool m_isHeader;
};

/*
 * Structure to store and construct a table from an unstructured list
 * of cell
 *
 */
class WPSTable
{
public:
	//! the constructor
	WPSTable()
		: m_cellsList()
		, m_rowsSize()
		, m_colsSize() {}
	WPSTable &operator=(WPSTable const &)=default;
	//! the destructor
	virtual ~WPSTable();

	//! add a new cells
	void add(WPSCellPtr &cell);

	//! returns the number of cell
	int numCells() const
	{
		return int(m_cellsList.size());
	}
	//! returns the i^th cell
	WPSCellPtr getCell(int id);

	/** try to send the table

	Note: either send the table ( and returns true ) or do nothing.
	 */
	bool sendTable(WPSContentListenerPtr listener);

	/** try to send the table as basic text */
	bool sendAsText(WPSContentListenerPtr listener);

protected:
	//! create the correspondance list, ...
	bool buildStructures();

	/** the list of cells */
	std::vector<WPSCellPtr> m_cellsList;
	/** the final row and col size (in point) */
	std::vector<float> m_rowsSize, m_colsSize;
};

#endif
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
