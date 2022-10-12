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

#ifndef LOTUS_H
#define LOTUS_H

#include <vector>

#include <librevenge-stream/librevenge-stream.h>
#include "libwps_internal.h"
#include "libwps_tools_win.h"
#include "WPSDebug.h"

#include "WKSParser.h"

namespace LotusParserInternal
{
class SubDocument;
struct State;
}

class LotusGraph;
class LotusChart;
class LotusSpreadsheet;
class LotusStyleManager;
class WPSGraphicStyle;
class WPSOLE1Parser;

/* .wk3: a spreadsheet is composed in two files
       + a wk3 file which contains the spreadsheet data
       + a fm3 file which contains the different formatings

   .wk4: the file contains three parts:
       + the wk3 previous file
	   + the fm3 file
	   + an unknown part, which may code the file structure,

	   Normally the wk3 and the fm3 are a sequence of small zones,
	   but picture seems to be appeared at random position inside the
	   fm3 part (and even inside some structure fm3 structures...)

	   search for .ole and OLE1

	.123: the file contains at least two parts:
	   + the 123 storing the spreadsheet's data and format
	   + the last part containing the file's structure
	   + some optional part containing chart, picture, ...
 */

/**
 * This class parses a wk3,wk4,123 Lotus spreadsheet
 *
 */
class LotusParser final : public WKSParser
{
	friend class LotusParserInternal::SubDocument;
	friend class LotusChart;
	friend class LotusGraph;
	friend class LotusSpreadsheet;
	friend class LotusStyleManager;
public:
	//! constructor
	LotusParser(RVNGInputStreamPtr &input, WPSHeaderPtr &header,
	            libwps_tools_win::Font::Type encoding=libwps_tools_win::Font::UNKNOWN,
	            char const *password=nullptr);
	//! destructor
	~LotusParser() final;
	//! called by WPSDocument to parse the file
	void parse(librevenge::RVNGSpreadsheetInterface *documentInterface) final;
	//! checks if the document header is correct (or not)
	bool checkHeader(WPSHeader *header, bool strict=false);

	//! basic struct used to store link
	struct Link
	{
		//! constructor
		Link() : m_name(), m_linkName()
		{
		}
		//! the basic name(used to retrieve a data)
		std::string m_name;
		//! the cell positions
		WPSVec3i m_cells[2];
		//! the link name
		librevenge::RVNGString m_linkName;
	};
protected:
	//! return the file version
	int version() const;

	//
	// interface
	//

	//! returns the font corresponding to an id
	bool getFont(int id, WPSFont &font, libwps_tools_win::Font::Type &type) const;
	/** returns the default font type, ie. the encoding given by the constructor if given
		or the encoding deduced from the version.
	 */
	libwps_tools_win::Font::Type getDefaultFontType() const;
	//! returns a list of links corresponding to an id
	std::vector<Link> getLinksList(int lId) const;

	//
	// interface with LotusChart
	//

	//! try to send a chart
	bool sendChart(int cId, WPSPosition const &pos, WPSGraphicStyle const &style);

	//
	// interface with LotusGraph
	//

	//! return true if the sheet sheetId has some graphic
	bool hasGraphics(int sheetId) const;
	//! send the graphics corresponding to a sheetId
	void sendGraphics(int sheetId);

	//
	// interface with LotusSpreadsheet
	//

	//! returns the left top position of a cell
	bool getLeftTopPosition(Vec2i const &cell, int spreadsheet, Vec2f &pos) const;
	//! returns the name of the id's spreadsheet
	librevenge::RVNGString getSheetName(int id) const;

	//
	// interface with WPSOLE1Parser
	//

	/** try to retrieve the content of a graphic, knowing it local id */
	bool updateEmbeddedObject(int id, WPSEmbeddedObject &object) const;

	/** try to parse the different zones */
	bool createZones();
	/** creates the main listener */
	bool createListener(librevenge::RVNGSpreadsheetInterface *interface);

	//
	// low level
	//

	/// check for the existence of a format stream, if it exists, parse it
	bool parseFormatStream();

	//! checks if the document header is correct (or not)
	bool checkHeader(std::shared_ptr<WPSStream> stream, bool mainStream, bool strict);
	/** finds the different zones (spreadsheet, chart, print, ...) */
	bool readZones(std::shared_ptr<WPSStream> stream);
	/** parse the different zones 1B */
	bool readDataZone(std::shared_ptr<WPSStream> stream);
	//! reads a zone
	bool readZone(std::shared_ptr<WPSStream> &stream);
	//! reads a zone of type 1: 123 files
	bool readZone1(std::shared_ptr<WPSStream> stream);
	//! reads a zone of type 2: 123 files
	bool readSheetZone(std::shared_ptr<WPSStream> stream);
	//! reads a zone of type 4: 123 files
	static bool readZone4(std::shared_ptr<WPSStream> stream);
	//! reads a zone of type 5: 123 files
	static bool readChartZone(std::shared_ptr<WPSStream> stream);
	//! reads a zone of type 6: 123 files
	static bool readRefZone(std::shared_ptr<WPSStream> stream);
	//! reads a zone of type 7: 123 files
	static bool readZone7(std::shared_ptr<WPSStream> stream);
	//! reads a zone of type 8: 123 files
	bool readZone8(std::shared_ptr<WPSStream> stream);
	//! reads a zone of type a: 123 files
	static bool readVersionZone(std::shared_ptr<WPSStream> stream);
	//! parse a wk123 zone
	static bool readZoneV3(std::shared_ptr<WPSStream> stream);
	//////////////////////// generic ////////////////////////////////////

	//! reads a mac font name
	bool readMacFontName(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a format style name: b6
	static bool readFMTStyleName(std::shared_ptr<WPSStream> stream);
	//! reads a link
	bool readLinkZone(std::shared_ptr<WPSStream> stream);
	//! reads a mac document info zone: zone 1b, then 2af8
	bool readDocumentInfoMac(std::shared_ptr<WPSStream> stream, long endPos);

	//////////////////////// decode a stream //////////////////////////////

	//! try to decode a stream, if successful, replace the stream'input by the new one
	static RVNGInputStreamPtr decodeStream(RVNGInputStreamPtr input, long endPos, std::vector<uint8_t> const &key);
	//! try to guess a password knowing its file keys. Returns the keys if it founds a valid password
	static std::vector<uint8_t> retrievePasswordKeys(std::vector<uint8_t> const &fileKeys);

	std::shared_ptr<WKSContentListener> m_listener; /** the listener (if set)*/
	//! the internal state
	std::shared_ptr<LotusParserInternal::State> m_state;
	//! the style manager
	std::shared_ptr<LotusStyleManager> m_styleManager;
	//! the chart manager
	std::shared_ptr<LotusChart> m_chartParser;
	//! the graph manager
	std::shared_ptr<LotusGraph> m_graphParser;
	//! the spreadsheet manager
	std::shared_ptr<LotusSpreadsheet> m_spreadsheetParser;
	//! the ole1 parser
	std::shared_ptr<WPSOLE1Parser> m_ole1Parser;
};

#endif /* LOTUS_H */
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
