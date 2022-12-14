/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: t; c-basic-offset: 4 -*- */
/* libwps
 * Version: MPL 2.0 / LGPLv2.1+
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *
 * Major Contributor(s):
 * Copyright (C) 2015 Sean Young <sean@mess.org>
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

#ifndef DOSWORD_H
#define DOSWORD_H

#include <librevenge-stream/librevenge-stream.h>
#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "MSWrite.h"
#include "WPSParser.h"
#include "WPSEntry.h"
#include "WPSPageSpan.h"

/**
 * This class parses Microsoft Word for DOS
 *
 */
class DosWordParser final : public MSWriteParser
{
public:
	DosWordParser(RVNGInputStreamPtr &input, WPSHeaderPtr &header,
	              libwps_tools_win::Font::Type encoding=libwps_tools_win::Font::UNKNOWN);

	~DosWordParser() final;

	//! checks if the document header is correct (or not)
	bool checkHeader(WPSHeader *header, bool strict=false);

private:
	DosWordParser(const DosWordParser &) = delete;
	DosWordParser &operator=(const DosWordParser &) = delete;
	libwps_tools_win::Font::Type getFileEncoding(libwps_tools_win::Font::Type hint) final;

	static WPSColor color(int clr);

	void readSECT(uint32_t fcSep, uint32_t fcLim);
	void readSED() final;
	void readFFNTB() final;
	void readCHP(uint32_t fcFirst, uint32_t fcLim, unsigned cch) final;
	void readPAP(uint32_t fcFirst, uint32_t fcLim, unsigned cch) final;
	void readSUMD() final;
	void readFNTB() final;

	void insertSpecial(uint8_t val, uint32_t fc, MSWriteParserInternal::Paragraph::Location location) final;
	void insertControl(uint8_t val, uint32_t fc) final;
};

#endif /* DOSWORD_H */
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
