/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: t; c-basic-offset: 4 -*- */
/* libwps
 * Version: MPL 2.0 / LGPLv2.1+
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *
 * Major Contributor(s):
 * Copyright (C) 2020 Laurent Alonso (alonso.laurent@gmail.com)
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

#ifndef XY_WRITE_H
#define XY_WRITE_H

#include <memory>
#include <vector>

#include <librevenge-stream/librevenge-stream.h>
#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WPSParser.h"

namespace XYWriteParserInternal
{
class SubDocument;

struct Cell;
struct Format;
struct State;
}

/**
 * This class parses XYWrite Dos and Win4 files
 *
 */
class XYWriteParser : public WPSParser
{
	friend struct XYWriteParserInternal::Cell;
	friend class XYWriteParserInternal::SubDocument;

public:
	//! constructor
	XYWriteParser(RVNGInputStreamPtr &input, WPSHeaderPtr &header,
	              libwps_tools_win::Font::Type encoding=libwps_tools_win::Font::UNKNOWN);
	//! destructor
	~XYWriteParser() override;
	//! called by WPSDocument to parse the file
	void parse(librevenge::RVNGTextInterface *documentInterface) override;
	//! checks if the document header is correct (or not)
	bool checkHeader(WPSHeader *header, bool strict=false);

private:
	XYWriteParser(const XYWriteParser &) = delete;
	XYWriteParser &operator=(const XYWriteParser &) = delete;

	/** creates the main listener */
	std::shared_ptr<WPSContentListener> createListener(librevenge::RVNGTextInterface *interface);

protected:
	//! check if the file position is correct or not
	bool checkFilePosition(long pos) const;
	/** tries to find the end of main zone, the meta data zones (only Win4) */
	bool findAllZones();
	/** tries to parse the main text zone */
	bool parseTextZone(WPSEntry const &entry, std::string const &styleName="");
	/** tries to parse a frame */
	bool parseFrameZone(XYWriteParserInternal::Format const &frameFormat);
	/** tries to parse a picture */
	bool parsePictureZone(XYWriteParserInternal::Format const &pictureFormat);
	/** tries to parse the meta data zone */
	bool parseMetaData(WPSEntry const &entry);
	/** tries to parse a format: first character 0xae is read */
	bool parseFormat(XYWriteParserInternal::Format &format);
	/** tries to parse again a format to create a list of child: SS, FA, FM1, ... */
	bool createFormatChildren(XYWriteParserInternal::Format &format, size_t fPos=2);
	//! try to update the listener data(font,paragraph style, ...)
	bool update(XYWriteParserInternal::Format const &format, libwps_tools_win::Font::Type &fontType) const;
	/** tries to create a table */
	bool createTable(XYWriteParserInternal::Format const &format, long endPos);

	// State

	/** the listener (if set)*/
	std::shared_ptr<WPSContentListener> m_listener;
	/** the main state*/
	std::shared_ptr<XYWriteParserInternal::State> m_state;
};

#endif /* MS_WRITE_H */
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
