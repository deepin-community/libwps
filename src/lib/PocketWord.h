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

#ifndef POCKET_WORD_H
#define POCKET_WORD_H

#include <memory>
#include <vector>

#include <librevenge-stream/librevenge-stream.h>
#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WPSParser.h"

namespace PocketWordParserInternal
{
class SubDocument;

struct State;
}

/**
 * This class parses some Pocket Word
 *
 */
class PocketWordParser : public WPSParser
{
	friend class PocketWordParserInternal::SubDocument;

public:
	//! constructor
	PocketWordParser(RVNGInputStreamPtr &input, WPSHeaderPtr &header,
	                 libwps_tools_win::Font::Type encoding=libwps_tools_win::Font::UNKNOWN);
	//! destructor
	~PocketWordParser() override;
	//! called by WPSDocument to parse the file
	void parse(librevenge::RVNGTextInterface *documentInterface) override;
	//! checks if the document header is correct (or not)
	bool checkHeader(WPSHeader *header, bool strict=false);

private:
	PocketWordParser(const PocketWordParser &) = delete;
	PocketWordParser &operator=(const PocketWordParser &) = delete;

	/** creates the main listener */
	std::shared_ptr<WPSContentListener> createListener(librevenge::RVNGTextInterface *interface);

protected:
	//! try to read the different zones
	bool createZones();
	//! try to read the font names
	bool readFontNames(WPSEntry const &entry);
	//! try to read the page dimension
	bool readPageDims(WPSEntry const &entry);
	//! try to read a paragraph list
	bool readParagraphList(WPSEntry const &entry, std::vector<int> &paraId);
	//! try to read a paragraph dimensions' zone which follows the paragraph zone
	bool readParagraphDims(WPSEntry const &entry);
	//! try to read a paragraph unknown zone: tabs, link?
	bool readParagraphUnkn(WPSEntry const &entry);
	//! try to read a sound definition (maybe a picture)
	bool readSound(WPSEntry const &entry, WPSEmbeddedObject &object);
	//! try to read a sound data (maybe a picture)
	bool readSoundData(WPSEntry const &entry, long pictSize, WPSEmbeddedObject &object);
	//! try to read the unknown zone 8: one by file, maybe prefs
	bool readUnkn8(WPSEntry const &entry);
	//! try to read the unknown zone 20 and 21: one by file, find always no data, maybe style
	bool readUnkn2021(WPSEntry const &entry, int type);
	//! check if the file position is correct or not
	bool checkFilePosition(long pos) const;
	//! try to parse the unparsed zones
	void checkUnparsed();

	//! try to send all the data
	void sendData();
	//! try to read and send a paragraph
	bool sendParagraph(size_t paraId);

	// State

	/** the listener (if set)*/
	std::shared_ptr<WPSContentListener> m_listener;
	/** the main state*/
	std::shared_ptr<PocketWordParserInternal::State> m_state;
};

#endif /* POCKET_WORD_H */
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
