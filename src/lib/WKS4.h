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

#ifndef WKS4_H
#define WKS4_H

#include <vector>

#include <librevenge-stream/librevenge-stream.h>
#include "libwps/libwps.h"

#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WKSParser.h"

namespace WKS4ParserInternal
{
class SubDocument;
struct State;
}

class WKS4Chart;
class WKS4Spreadsheet;

/**
 * This class parses Microsoft Works spreadsheet or a database file
 *
 */
class WKS4Parser final : public WKSParser
{
	friend class WKS4ParserInternal::SubDocument;
	friend class WKS4Chart;
	friend class WKS4Spreadsheet;
public:
	//! constructor
	WKS4Parser(RVNGInputStreamPtr &input, WPSHeaderPtr &header,
	           libwps_tools_win::Font::Type encoding=libwps_tools_win::Font::UNKNOWN,
	           char const *password=nullptr);
	//! destructor
	~WKS4Parser() override;
	//! called by WPSDocument to parse the file
	void parse(librevenge::RVNGSpreadsheetInterface *documentInterface) final;
	//! checks if the document header is correct (or not)
	bool checkHeader(WPSHeader *header, bool strict=false);

	//! try to read a basic C string, knowing the maximum size
	bool readCString(librevenge::RVNGString &string, long maxSize);

protected:
	//! return true if the pos is in the file, update the file size if need
	bool checkFilePosition(long pos);
	//! return the file version
	int version() const;
	//! return the file creator
	libwps::WPSCreator creator() const;
	/** returns the default font type, ie. the encoding given by the constructor if given
		or the encoding deduiced from the version.
	 */
	libwps_tools_win::Font::Type getDefaultFontType() const;
	//! returns the creator
	libwps::WPSCreator getCreator() const;

	//
	// interface with WKS4Spreadsheet
	//

	//! returns the color corresponding to an id
	bool getColor(int id, WPSColor &color) const;
	//! returns the font corresponding to an id
	bool getFont(int id, WPSFont &font, libwps_tools_win::Font::Type &type) const;
	//! returns the name of the id's spreadsheet
	librevenge::RVNGString getSheetName(int id) const;

	/** creates the main listener */
	std::shared_ptr<WKSContentListener> createListener(librevenge::RVNGSpreadsheetInterface *interface);
	//! send the header/footer
	void sendHeaderFooter(bool header);

	//
	// low level
	//

	/// check for the existence of a format stream, if it exists, parse it
	bool parseFormatStream();
	/// reset the main input
	void resetMainInput(RVNGInputStreamPtr newInput);

	/** finds the different zones (spreadsheet, chart, print, ...) */
	bool readZones();
	//! reads a zone
	bool readZone();
	//! reads a Quattro Pro zone
	bool readZoneQuattro();

	//////////////////////// generic ////////////////////////////////////

	//! reads a mswork font
	bool readFont();

	//! reads a printer data ?
	bool readPrnt();

	//! reads another printer data. Seem simillar to ZZPrnt
	bool readPrn2();

	//! reads the header/footer
	bool readHeaderFooter(bool header);

	//! read a list of field name + ...
	bool readFieldName();

	//////////////////////// decode a lotus stream //////////////////////////////

	//! try to decode a stream, if successful, replace the stream'input by the new one
	static RVNGInputStreamPtr decodeStream(RVNGInputStreamPtr input, long endPos, std::vector<uint8_t> const &key);

	//////////////////////// unknown zone //////////////////////////////

	//! reads windows record 0:7|0:9
	bool readWindowRecord();
	/** reads some unknown spreadsheet zones 0:18|0:19|0:20|0:27|0:2a

	 \note this zones seems to consist of a list of flags potentially
	 followed by other data*/
	bool readUnknown1();

	std::shared_ptr<WKSContentListener> m_listener; /** the listener (if set)*/
	//! the internal state
	std::shared_ptr<WKS4ParserInternal::State> m_state;
	//! the chart manager
	std::shared_ptr<WKS4Chart> m_chartParser;
	//! the spreadsheet manager
	std::shared_ptr<WKS4Spreadsheet> m_spreadsheetParser;
};

#endif /* WPS4_H */
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
