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
 * Copyright (C) 2003-2005 William Lachance (william.lachance@sympatico.ca)
 * Copyright (C) 2003 Marc Maurer (uwog@uwog.net)
 *
 * For minor contributions see the git repository.
 *
 * Alternatively, the contents of this file may be used under the terms
 * of the GNU Lesser General Public License Version 2.1 or later
 * (LGPLv2.1+), in which case the provisions of the LGPLv2.1+ are
 * applicable instead of those above.
 */

#ifndef QUATTRO_SPREADSHEET_H
#define QUATTRO_SPREADSHEET_H

#include <ostream>
#include <map>
#include <vector>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"

#include "WPSDebug.h"
#include "WKSContentListener.h"

#include "QuattroFormula.h"

namespace QuattroSpreadsheetInternal
{
class Cell;
class SpreadSheet;
struct State;
}

class QuattroParser;

/**
 * This class parses Quattro Pro spreadsheet file
 *
 */
class QuattroSpreadsheet
{
public:
	friend class QuattroParser;

	//! constructor
	explicit QuattroSpreadsheet(QuattroParser &parser);
	//! destructor
	~QuattroSpreadsheet();
	//! sets the listener
	void setListener(WKSContentListenerPtr &listen)
	{
		m_listener = listen;
	}
	//! clean internal state
	void cleanState();
	//! update the state (need to be called before sending data)
	void updateState();

protected:
	//! return the file version
	int version() const;
	//! returns the function to read a cell's reference
	QuattroFormulaManager::CellReferenceFunction getReadCellReferenceFunction();

	//! returns the number of spreadsheet
	int getNumSpreadsheets() const;
	//! returns the name of the id's spreadsheet
	librevenge::RVNGString getSheetName(int id) const;
	//! send the sheetId'th spreadsheet
	void sendSpreadsheet(int sheetId, std::vector<Vec2i> const &listGraphicCells);
	//! returns the beginning position of a cell
	Vec2f getPosition(int sheetId, Vec2i const &cell) const;
	//! send the cell data
	void sendCellContent(QuattroSpreadsheetInternal::Cell const &cell, int sheetId);

	//! add a dll's correspondance between an id and a name
	void addDLLIdName(int id, librevenge::RVNGString const &name, bool func1);
	//! add a user format's correspondance between an id and a name
	void addUserFormat(int id, librevenge::RVNGString const &name);

	//
	// low level
	//
	//////////////////////// spreadsheet //////////////////////////////

	//! reads a cell content data: zone 0xc-0x10 or 33
	bool readCell(std::shared_ptr<WPSStream> const &stream);
	//! reads sheet size: zone 06
	bool readSheetSize(std::shared_ptr<WPSStream> const &stream);
	//! reads the sheet column/row default size: zone d2-d5
	bool readColumnRowDefaultSize(std::shared_ptr<WPSStream> const &stream);
	//! reads the column size: zone d8,d9
	bool readColumnSize(std::shared_ptr<WPSStream> const &stream);
	//! reads the row size: zone d6,d7
	bool readRowSize(std::shared_ptr<WPSStream> const &stream);
	//! reads the row size: zone 105,106
	bool readRowRangeSize(std::shared_ptr<WPSStream> const &stream);

	//! read the begin/end of a sheet zone: zone ca and cb
	bool readBeginEndSheet(std::shared_ptr<WPSStream> const &stream, int &sheetId);
	//! reads sheet name: zone cc
	bool readSheetName(std::shared_ptr<WPSStream> const &stream);
	//! reads a cell attribute: zone 0xce
	bool readCellStyle(std::shared_ptr<WPSStream> const &stream);

	//! reads a view info: zone 197/198
	static bool readViewInfo(std::shared_ptr<WPSStream> const &stream);

	/* reads a cell */
	static bool readCell(std::shared_ptr<WPSStream> const &stream, Vec2i actPos, WKSContentListener::FormulaInstruction &instr, int sheetId, librevenge::RVNGString const &fName);
	//! try to update the cell's format using the user format
	static void updateCellWithUserFormat(QuattroSpreadsheetInternal::Cell &cell, librevenge::RVNGString const &format);
	//! try to read a cell reference
	bool readCellReference(std::shared_ptr<WPSStream> const &stream, long endPos,
	                       QuattroFormulaInternal::CellReference &ref,
	                       Vec2i const &pos=Vec2i(0,0), int sheetId=0) const;

private:
	QuattroSpreadsheet(QuattroSpreadsheet const &orig) = delete;
	QuattroSpreadsheet &operator=(QuattroSpreadsheet const &orig) = delete;
	std::shared_ptr<WKSContentListener> m_listener; /** the listener (if set)*/
	//! the main parser
	QuattroParser &m_mainParser;
	//! the internal state
	std::shared_ptr<QuattroSpreadsheetInternal::State> m_state;
};

#endif /* WPS4_H */
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
