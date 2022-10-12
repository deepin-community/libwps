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

#ifndef QUATTRO9_SPREADSHEET_H
#define QUATTRO9_SPREADSHEET_H

#include <ostream>
#include <map>
#include <vector>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"

#include "WPSDebug.h"
#include "WKSContentListener.h"

#include "QuattroFormula.h"

namespace Quattro9SpreadsheetInternal
{
struct CellData;
class SpreadSheet;
struct State;
}

namespace Quattro9ParserInternal
{
struct TextEntry;
}

class Quattro9Parser;

/**
 * This class parses Quattro9 Pro spreadsheet file
 *
 */
class Quattro9Spreadsheet
{
public:
	friend class Quattro9Parser;

	//! constructor
	explicit Quattro9Spreadsheet(Quattro9Parser &parser);
	//! destructor
	~Quattro9Spreadsheet();
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
	void sendSpreadsheet(int sheetId);
	//! returns the beginning position of a cell
	Vec2f getPosition(int sheetId, Vec2i const &cell) const;
	//! add a dll's correspondance between an id and a name
	void addDLLIdName(int id, librevenge::RVNGString const &name, bool func1);
	//! add a user format's correspondance between an id and a name
	void addUserFormat(int id, librevenge::RVNGString const &name);
	//! set the document strings list
	void addDocumentStrings(std::shared_ptr<WPSStream> const &stream,
	                        std::vector<Quattro9ParserInternal::TextEntry> const &entries);

	//
	// low level
	//

	//! send the cell data
	void sendCellContent(Quattro9SpreadsheetInternal::CellData const *cell, Vec2i pos, int sheetId, int numRepeated);

	//! reads a cell attribute: zone a
	bool readCellStyles(std::shared_ptr<WPSStream> const &stream);
	//! reads the document formulas: zone 408
	bool readDocumentFormulas(std::shared_ptr<WPSStream> const &stream);
	//! read the begin sheet zone: zone 601
	bool readBeginSheet(std::shared_ptr<WPSStream> const &stream, int &sheetId);
	//! read the end sheet zone: zone 602
	bool readEndSheet(std::shared_ptr<WPSStream> const &stream);
	//! read the page break zone: zone 617
	static bool readPageBreak(std::shared_ptr<WPSStream> const &stream);
	//! read the merged cell: zone 61d
	bool readMergedCells(std::shared_ptr<WPSStream> const &stream);
	//! read a col/row default dimension: zone 631, 632
	bool readColRowDefault(std::shared_ptr<WPSStream> const &stream);
	//! read a col/row dimension: zone 633, 634
	bool readColRowDimension(std::shared_ptr<WPSStream> const &stream);
	//! read a col/row dimensions: zone 635, 636
	bool readColRowDimensions(std::shared_ptr<WPSStream> const &stream);
	//! read the begin column zone: zone a01
	bool readBeginColumn(std::shared_ptr<WPSStream> const &stream);
	//! read the end column zone: zone a02
	bool readEndColumn(std::shared_ptr<WPSStream> const &stream);
	//! reads a cell list zone: zone c01
	bool readCellList(std::shared_ptr<WPSStream> const &stream);
	//! reads a cell result zone: zone c02
	bool readCellResult(std::shared_ptr<WPSStream> const &stream);

	/* reads a cell */
	static bool readCell(std::shared_ptr<WPSStream> const &stream, Vec2i actPos, WKSContentListener::FormulaInstruction &instr, int sheetId, librevenge::RVNGString const &fName);
	//! try to read a cell reference
	bool readCellReference(std::shared_ptr<WPSStream> const &stream, long endPos,
	                       QuattroFormulaInternal::CellReference &ref,
	                       Vec2i const &pos=Vec2i(0,0), int sheetId=0) const;
	//////////////////////// spreadsheet //////////////////////////////

private:
	Quattro9Spreadsheet(Quattro9Spreadsheet const &orig) = delete;
	Quattro9Spreadsheet &operator=(Quattro9Spreadsheet const &orig) = delete;
	std::shared_ptr<WKSContentListener> m_listener; /** the listener (if set)*/
	//! the main parser
	Quattro9Parser &m_mainParser;
	//! the internal state
	std::shared_ptr<Quattro9SpreadsheetInternal::State> m_state;
};

#endif /* WPS4_H */
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
