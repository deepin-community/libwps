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

#ifndef QUATTRO_9_H
#define QUATTRO_9_H

#include <map>
#include <vector>

#include <librevenge-stream/librevenge-stream.h>
#include "libwps/libwps.h"

#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WKSParser.h"
#include "WPSGraphicStyle.h"

namespace Quattro9ParserInternal
{
class SubDocument;
struct State;

struct TextEntry
{
	//! constructor
	TextEntry()
		: m_entry()
		, m_fontsList()
		, m_posFontIdMap()
		, m_flag(0)
		, m_extra()
	{
	}
	//! returns true if the string is empty
	bool empty() const
	{
		return !m_entry.valid();
	}
	//! returns the string
	librevenge::RVNGString getString(std::shared_ptr<WPSStream> const &stream, libwps_tools_win::Font::Type type=libwps_tools_win::Font::WIN3_WEUROPE) const;
	//! sends the text to the main listener
	void send(std::shared_ptr<WPSStream> const &stream, WPSFont const &font, libwps_tools_win::Font::Type type, WKSContentListenerPtr &listener);
	//! internal: returns a debug string
	std::string getDebugString(std::shared_ptr<WPSStream> const &stream) const;
	//! the text entry
	WPSEntry m_entry;
	//! the list of fonts
	std::vector<WPSFont> m_fontsList;
	//! the position to font map(complex text)
	std::map<int, int> m_posFontIdMap;
	//! the flag
	int m_flag;
	//! extra data
	std::string m_extra;
};
}

class Quattro9Graph;
class Quattro9Spreadsheet;

namespace QuattroFormulaInternal
{
class CellReference;
}

/**
 * This class parses Quattro Pro WP spreadsheet: .qpw
 *
 */
class Quattro9Parser final : public WKSParser
{
	friend class Quattro9ParserInternal::SubDocument;
	friend class Quattro9Graph;
	friend class Quattro9Spreadsheet;

public:
	//! constructor
	Quattro9Parser(RVNGInputStreamPtr &input, WPSHeaderPtr &header,
	               libwps_tools_win::Font::Type encoding=libwps_tools_win::Font::UNKNOWN,
	               char const *password=nullptr);
	//! destructor
	~Quattro9Parser() final;
	//! called by WPSDocument to parse the file
	void parse(librevenge::RVNGSpreadsheetInterface *documentInterface) final;
	//! checks if the document header is correct (or not)
	bool checkHeader(WPSHeader *header, bool strict=false);

protected:
	//! return the file version
	int version() const;
	/** returns the default font type, ie. the encoding given by the constructor if given
		or the encoding deduced from the version.
	 */
	libwps_tools_win::Font::Type getDefaultFontType() const;
	//! returns the name of the fId file
	bool getExternalFileName(int fId, librevenge::RVNGString &fName) const;
	//! returns the text and cell of a field instruction
	bool getField(int fId, librevenge::RVNGString &text,
	              QuattroFormulaInternal::CellReference &ref,
	              librevenge::RVNGString const &fileName) const;

	//
	// interface with Quattro9Spreadsheet
	//

	//! returns the font corresponding to an id
	bool getFont(int id, WPSFont &font) const;
	//! returns the beginning position of a cell
	Vec2f getCellPosition(int sheetId, Vec2i const &cell) const;

	//
	// interface with Quattro9Graph
	//

	//! returns the color corresponding to an id
	bool getColor(int id, WPSColor &color) const;
	//! returns the pattern corresponding to a pattern id between 0 and 24
	bool getPattern(int id, WPSGraphicStyle::Pattern &pattern) const;
	//! send the page graphic corresponding to a sheet
	bool sendPageGraphics(int sheetId) const;

	//
	//
	//

	/** creates the main listener */
	std::shared_ptr<WKSContentListener> createListener(librevenge::RVNGSpreadsheetInterface *interface);

	//! send the header/footer
	void sendHeaderFooter(bool header);

	//
	// low level
	//

	//! checks if the document header is correct (or not)
	bool checkHeader(std::shared_ptr<WPSStream> stream, bool strict);
	/** finds the different zones in the main OLE stream (spreadsheet, chart, print, ...) */
	bool readZones();
	/** try to read a zone */
	bool readZone(std::shared_ptr<WPSStream> &stream);

	//////////////////////// generic ////////////////////////////////////

	/** try to read a string: length[2bytes], unknow[bytes] followed by the string */
	bool readPString(std::shared_ptr<WPSStream> const &stream, long endPos, Quattro9ParserInternal::TextEntry &entry);
	//! read a font name zone: zone 6
	bool readFontNames(std::shared_ptr<WPSStream> const &stream);
	//! read a font style zone: zone 7
	bool readFontStyles(std::shared_ptr<WPSStream> const &stream);
	//! read a zone which contains two files position (prev and next): zone 2,402,602,a02
	static bool readFilePositions(std::shared_ptr<WPSStream> const &stream, long (&filePos)[2]);
	//! read a zone 406 which contains a list of fields
	bool readDocumentFields(std::shared_ptr<WPSStream> const &stream);
	//! read a zone 407 which contains a list of stream
	bool readDocumentStrings(std::shared_ptr<WPSStream> const &stream);

	//! read a text entry style zone
	bool readTextStyles(std::shared_ptr<WPSStream> const &stream, long endPos, Quattro9ParserInternal::TextEntry &entry);
	//! read a font style in a text entry zone
	bool readTextFontStyles(std::shared_ptr<WPSStream> const &stream, int dataSz, WPSFont &font);

	//! add the document strings to the spreadsheetParser
	void addDocumentStrings();

	//////////////////////// unknown zone ////////////////////////////////////


	//////////////////////// Ole specific zone ////////////////////////////////////

	/** try to parse the OLE stream(except the main stream) */
	bool parseOLEStream(RVNGInputStreamPtr input, std::string const &avoid="");
	/** try to read the link info sub stream */
	static bool readOleLinkInfo(std::shared_ptr<WPSStream> stream);
	/** try to read the BOlePart sub stream: a zone which contains 5 long */
	static bool readOleBOlePart(std::shared_ptr<WPSStream> stream);

	//////////////////////// decode a quattro stream //////////////////////////////

	//! try to decode a stream, if successful, replace the stream'input by the new one
	static RVNGInputStreamPtr decodeStream(RVNGInputStreamPtr input, std::vector<uint8_t> const &key);

	std::shared_ptr<WKSContentListener> m_listener; /** the listener (if set)*/
	//! the internal state
	std::shared_ptr<Quattro9ParserInternal::State> m_state;
	//! the graph manager
	std::shared_ptr<Quattro9Graph> m_graphParser;
	//! the spreadsheet manager
	std::shared_ptr<Quattro9Spreadsheet> m_spreadsheetParser;
};

#endif /* WPS4_H */
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
