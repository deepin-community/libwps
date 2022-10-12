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

#ifndef MULTIPLAN_H
#define MULTIPLAN_H

#include <vector>

#include <librevenge-stream/librevenge-stream.h>
#include "libwps/libwps.h"

#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WKSContentListener.h"

#include "WKSParser.h"

namespace libwps
{

namespace MultiplanParserInternal
{
struct State;
}

/**
 * This class parses Microsoft Multiplan DOS spreadsheet v1
 *
 */
class MultiplanParser final : public WKSParser
{
public:
	//! constructor
	MultiplanParser(RVNGInputStreamPtr &input, WPSHeaderPtr &header,
	                libwps_tools_win::Font::Type encoding=libwps_tools_win::Font::UNKNOWN,
	                char const *password=nullptr);
	//! destructor
	~MultiplanParser() override;
	//! called by WPSDocument to parse the file
	void parse(librevenge::RVNGSpreadsheetInterface *documentInterface) final;
	//! checks if the document header is correct (or not)
	bool checkHeader(WPSHeader *header, bool strict=false);

protected:
	//! return true if the pos is in the file, update the file size if need
	bool checkFilePosition(long pos);
	//! return the file version
	int version() const;
	/** returns the default font type, ie. the encoding given by the constructor if given
		or the encoding deduiced from the version.
	 */
	libwps_tools_win::Font::Type getDefaultFontType() const;

	/** creates the main listener */
	static std::shared_ptr<WKSContentListener> createListener(librevenge::RVNGSpreadsheetInterface *interface);
	//! try to send the main spreadsheet
	bool sendSpreadsheet();
	//! try to send a cell
	bool sendCell(Vec2i const &cellPos, int pos);

	//
	// low level
	//

	/** finds the different zones (spreadsheet, chart, print, ...) */
	bool readZones();
	//! read the columns width
	bool readColumnsWidth();
	//! read the spreadsheet zone list
	bool readZonesList();
	//! read the spreadsheet zone list v2
	bool readZonesListV2();
	//! read the cell data position
	bool readCellDataPosition(WPSEntry const &entry);
	//! read a link
	bool readLink(int pos, WKSContentListener::FormulaInstruction &instruction);
	//! read a link filename
	bool readFilename(int pos, librevenge::RVNGString &filename);
	//! read a shared data
	bool readSharedData(int pos, int cellType, Vec2i const &cellPos, WKSContentListener::CellContent &content);
	//! reads a name and returns the cell's instruction
	bool readName(int pos, WKSContentListener::FormulaInstruction &instruction);
	//! try to read the function names: v2
	bool readFunctionNamesList();

	//////////////////////// encryption ////////////////////////////////////

	//! check if the password corresponds to a ket
	bool checkPassword(char const *password) const;
	//! try to decode a stream, if successful, replace the stream'input by the new one
	RVNGInputStreamPtr decodeStream(RVNGInputStreamPtr input);
	//! try to guess a password supposing that the Zone0 content is default
	bool retrievePasswordKeys();

	//////////////////////// unknown zone //////////////////////////////

	//! read an unknown zone
	bool readZoneB();


protected:
	//! try to read a double value
	bool readDouble(double &value);
	//! try to read a formula
	bool readFormula(Vec2i const &cellPos, std::vector<WKSContentListener::FormulaInstruction> &formula, long endPos, std::string &extra);
	//! try to read a formula V2
	bool readFormulaV2(Vec2i const &cellPos, std::vector<WKSContentListener::FormulaInstruction> &formula, long endPos, std::string &extra);

	std::shared_ptr<WKSContentListener> m_listener; /** the listener (if set)*/
	//! the internal state
	std::shared_ptr<MultiplanParserInternal::State> m_state;
};

}

#endif /* MULTIPLAN_H */
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
