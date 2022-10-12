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

#ifndef QUATTRO_DOS_H
#define QUATTRO_DOS_H

#include <vector>

#include <librevenge-stream/librevenge-stream.h>
#include "libwps/libwps.h"

#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WKSParser.h"

namespace QuattroDosParserInternal
{
class SubDocument;
struct State;
}

class QuattroDosChart;
class QuattroDosSpreadsheet;

/**
 * This class parses Quattro Pro spreadsheet: .wq1 and wq2
 *
 */
class QuattroDosParser final : public WKSParser
{
	friend class QuattroDosParserInternal::SubDocument;
	friend class QuattroDosChart;
	friend class QuattroDosSpreadsheet;
public:
	//! constructor
	QuattroDosParser(RVNGInputStreamPtr &input, WPSHeaderPtr &header,
	                 libwps_tools_win::Font::Type encoding=libwps_tools_win::Font::UNKNOWN);
	//! destructor
	~QuattroDosParser() final;
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
	//! returns the name of the fId file
	librevenge::RVNGString getFileName(int fId) const;

	//
	// interface with QuattroDosSpreadsheet
	//

	//! returns the color corresponding to an id
	bool getColor(int id, WPSColor &color) const;
	//! returns the font corresponding to an id
	bool getFont(int id, WPSFont &font, libwps_tools_win::Font::Type &type) const;
	//! returns the name of the id's spreadsheet
	librevenge::RVNGString getSheetName(int id) const;

	//
	// interface with QuattroDosChart
	//

	//! try to send the chart corresponding to sheetId and a position
	bool sendChart(int sheetId, Vec2i const &cell, Vec2f const &chartSize);

	/** creates the main listener */
	std::shared_ptr<WKSContentListener> createListener(librevenge::RVNGSpreadsheetInterface *interface);
	//! send the header/footer
	void sendHeaderFooter(bool header);

	//
	// low level
	//

	/** finds the different zones (spreadsheet, chart, print, ...) */
	bool readZones();
	//! reads a zone
	bool readZone();

	//////////////////////// generic ////////////////////////////////////

	//! reads the user fonts
	bool readUserFonts();
	//! try to read a font
	bool readFont(WPSFont &font, libwps_tools_win::Font::Type &type);
	//! reads the header/footer
	bool readHeaderFooter(bool header);

	//! read a list of field name + ...
	bool readFieldName();
	//! read a file name
	bool readFileName();

	//! try to read a basic pascal string, knowing the maximum size(excluding string size)
	bool readPString(librevenge::RVNGString &string, long maxSize);
	//////////////////////// unknown zone //////////////////////////////

	//! reads windows record 0:7|0:9
	bool readWindowRecord();
	/** reads some unknown spreadsheet zones 0:18|0:19|0:20|0:27|0:2a

	 \note this zones seems to consist of a list of flags potentially
	 followed by other data*/
	bool readUnknown1();

	std::shared_ptr<WKSContentListener> m_listener; /** the listener (if set)*/
	//! the internal state
	std::shared_ptr<QuattroDosParserInternal::State> m_state;
	//! the spreadsheet manager
	std::shared_ptr<QuattroDosSpreadsheet> m_spreadsheetParser;
	//! the chart manager
	std::shared_ptr<QuattroDosChart> m_chartParser;
};

#endif /* WPS4_H */
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
