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

#ifndef QUATTRO_H
#define QUATTRO_H

#include <vector>

#include <librevenge-stream/librevenge-stream.h>
#include "libwps/libwps.h"

#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WKSParser.h"

namespace QuattroParserInternal
{
class SubDocument;
struct State;
}

//class QuattroChart;
class QuattroGraph;
class QuattroSpreadsheet;

namespace QuattroFormulaInternal
{
class CellReference;
}
/**
 * This class parses Quattro Pro spreadsheet: .wb1, ..., .wb3
 *
 */
class QuattroParser final : public WKSParser
{
	friend class QuattroParserInternal::SubDocument;
	//friend class QuattroChart;
	friend class QuattroGraph;
	friend class QuattroSpreadsheet;
public:
	//! constructor
	QuattroParser(RVNGInputStreamPtr &input, WPSHeaderPtr &header,
	              libwps_tools_win::Font::Type encoding=libwps_tools_win::Font::UNKNOWN,
	              char const *password=nullptr);
	//! destructor
	~QuattroParser() final;
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
	// interface with QuattroSpreadsheet
	//

	//! returns the color corresponding to an id
	bool getColor(int id, WPSColor &color) const;
	//! returns the font corresponding to an id
	bool getFont(int id, WPSFont &font, libwps_tools_win::Font::Type &type) const;
	//! returns the beginning position of a cell
	Vec2f getCellPosition(int sheetId, Vec2i const &cell) const;

	//
	// interface with QuattroGraph
	//

	//! send the page graphic corresponding to a sheet
	bool sendPageGraphics(int sheetId) const;
	//! send the graphic corresponding to a cell
	bool sendGraphics(int sheetId, Vec2i const &cell) const;

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
	/** finds the different zones (spreadsheet, chart, print, ...) */
	bool readZones();
	/** finds the different OLE zones: wb2 */
	bool readOLEZones(std::shared_ptr<WPSStream> &stream);
	/** try to read a zone */
	bool readZone(std::shared_ptr<WPSStream> &stream);

	//////////////////////// generic ////////////////////////////////////

	//! read a list of field name + ...: zone b
	bool readFieldName(std::shared_ptr<WPSStream> stream);
	/** read the cell's position: zone 96*/
	static bool readCellPosition(std::shared_ptr<WPSStream> stream);
	//! read a external filename/name: zone 97,98
	bool readExternalData(std::shared_ptr<WPSStream> stream);
	//! read a font: zone cf, fc and 110
	bool readFontDef(std::shared_ptr<WPSStream> stream);
	//! read a color lits: zone e8
	bool readColorList(std::shared_ptr<WPSStream> stream);
	//! read a style name: zone d0
	bool readStyleName(std::shared_ptr<WPSStream> stream);
	//! reads the header/footer: zone 25,26
	bool readHeaderFooter(std::shared_ptr<WPSStream> stream, bool header);
	//! read the pane attribute: d1
	bool readPaneAttribute(std::shared_ptr<WPSStream> stream);
	//! read the first optimizer zone: 103
	bool readOptimizer(std::shared_ptr<WPSStream> stream);
	//! read the table query command zone: 12f
	bool readQueryCommand(std::shared_ptr<WPSStream> stream);
	//! read the serie extension zone: 2dc
	static bool readSerieExtension(std::shared_ptr<WPSStream> stream);
	//! try to read a basic C string, knowing the maximum size
	bool readCString(std::shared_ptr<WPSStream> stream, librevenge::RVNGString &string, long maxSize);

	//////////////////////// unknown zone ////////////////////////////////////

	/** reads some cell reference list (potential followed by other data)*/
	bool readBlockList(std::shared_ptr<WPSStream> stream);
	/** reads a big zone(chart?) which contains sub zones: 341 */
	bool readZone341(std::shared_ptr<WPSStream> stream);

	//////////////////////// Ole specific zone ////////////////////////////////////

	/** try to parse the OLE stream(except the main stream) */
	bool parseOLEStream(RVNGInputStreamPtr input, std::string const &avoid="");
	/** try to read the link info sub stream */
	bool readOleLinkInfo(std::shared_ptr<WPSStream> stream, librevenge::RVNGString &link);
	/** try to read the BOlePart sub stream: a zone which contains 5 long */
	static bool readOleBOlePart(std::shared_ptr<WPSStream> stream);

	//////////////////////// decode a quattro stream //////////////////////////////

	//! try to decode a stream, if successful, replace the stream'input by the new one
	RVNGInputStreamPtr decodeStream(RVNGInputStreamPtr input, std::vector<uint8_t> const &key) const;

	std::shared_ptr<WKSContentListener> m_listener; /** the listener (if set)*/
	//! the internal state
	std::shared_ptr<QuattroParserInternal::State> m_state;
	// the chart manager
	// std::shared_ptr<QuattroChart> m_chartParser;
	//! the graph manager
	std::shared_ptr<QuattroGraph> m_graphParser;
	//! the spreadsheet manager
	std::shared_ptr<QuattroSpreadsheet> m_spreadsheetParser;
};

#endif /* WPS4_H */
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
