/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: t; c-basic-offset: 4 -*- */
/* libwps
 * Version: MPL 2.0 / LGPLv2.1+
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *
 * Major Contributor(s):
 * Copyright (C) 2002 William Lachance (william.lachance@sympatico.ca)
 * Copyright (C) 2002-2004 Marc Maurer (uwog@uwog.net)
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

#include <string.h>

#include "libwps_internal.h"

#include "WPSHeader.h"

using namespace libwps;

WPSHeader::WPSHeader(RVNGInputStreamPtr &input, RVNGInputStreamPtr &fileInput, int majorVersion, WPSKind kind, WPSCreator creator)
	: m_input(input)
	, m_fileInput(fileInput)
	, m_majorVersion(majorVersion)
	, m_kind(kind)
	, m_creator(creator)
	, m_isEncrypted(false)
	, m_needEncodingFlag(false)
{
}

WPSHeader::~WPSHeader()
{
}


/**
 * So far, we have identified three categories of Works documents.
 *
 * Works documents versions 3 and later use a MS OLE container, so we detect
 * their type by checking for OLE stream names.  Works version 2 is like
 * Works 3 without OLE, so those two types use the same parser.
 *
 */
WPSHeader *WPSHeader::constructHeader(RVNGInputStreamPtr &input)
{
	if (!input->isStructured())
	{
		input->seek(0, librevenge::RVNG_SEEK_SET);
		uint8_t val[6];
		for (unsigned char &i : val) i = libwps::readU8(input);

		if (val[0] < 6 && val[1] == 0xFE)
		{
			WPS_DEBUG_MSG(("WPSHeader::constructHeader: Microsoft Works v2 format detected\n"));
			return new WPSHeader(input, input, 2);
		}
		// works1 dos file begin by 2054
		if ((val[0] == 0xFF || val[0] == 0x20) && val[1]==0x54)
		{
			WPS_DEBUG_MSG(("WPSHeader::constructHeader: Microsoft Works wks database\n"));
			return new WPSHeader(input, input, 1, WPS_DATABASE);
		}
		if (val[0] == 0xFF && val[1] == 0 && val[2]==2)
		{
			WPS_DEBUG_MSG(("WPSHeader::constructHeader: Microsoft Works wks detected\n"));
			return new WPSHeader(input, input, 3, WPS_SPREADSHEET);
		}
		if (val[0] == 00 && val[1] == 0 && val[2]==2)
		{
			if (val[3]==0 && (val[4]==0x20 || val[4]==0x21) && val[5]==0x51)
			{
				WPS_DEBUG_MSG(("WPSHeader::constructHeader: Quattro Pro wq1 or wq2 detected\n"));
				return new WPSHeader(input, input, 2, WPS_SPREADSHEET, WPS_QUATTRO_PRO);
			}
			if (val[3]==0 && (val[4]==1 || val[4]==2) && val[5]==0x10)
			{
				WPS_DEBUG_MSG(("WPSHeader::constructHeader: Quattro Pro wb1 or wb2 detected\n"));
				return new WPSHeader(input, input, 1000, WPS_SPREADSHEET, WPS_QUATTRO_PRO);
			}
			WPS_DEBUG_MSG(("WPSHeader::constructHeader: potential Lotus|Microsft Works|Quattro Pro spreadsheet detected\n"));
			return new WPSHeader(input, input, 2, WPS_SPREADSHEET);
		}
		if (val[0] == 00 && val[1] == 0x0 && val[2]==0x1a)
		{
			WPS_DEBUG_MSG(("WPSHeader::constructHeader: Lotus spreadsheet detected\n"));
			return new WPSHeader(input, input, 101, WPS_SPREADSHEET, WPS_LOTUS);
		}
		if ((val[0] == 0x31 || val[0] == 0x32) && val[1] == 0xbe && val[2] == 0 && val[3] == 0 && val[4] == 0 && val[5] == 0xab)
		{
			// This value is always 0 for Word for DOS
			input->seek(96, librevenge::RVNG_SEEK_SET);
			if (libwps::readU16(input))
			{
				WPS_DEBUG_MSG(("WPSHeader::constructHeader: Microsoft Write detected\n"));
				return new WPSHeader(input, input, 3, WPS_TEXT, WPS_MSWRITE);
			}
			else
			{
				WPS_DEBUG_MSG(("WPSHeader::constructHeader: Microsoft Word for DOS detected\n"));
				return new WPSHeader(input, input, 0, WPS_TEXT, WPS_DOSWORD);
			}
		}
		if (val[0]==0x7b && val[1]==0x5c && val[2]==0x70 && val[3]==0x77 &&
		        val[4]==0x69 && val[5]==0x15)
		{
			WPS_DEBUG_MSG(("WPSHeader::constructHeader: PocketWord document detected\n"));
			return new WPSHeader(input, input, 1, WPS_TEXT, WPS_POCKETWORD);
		}
		if (val[0] == 0x08 && val[1] == 0xe7)
		{
			WPS_DEBUG_MSG(("WPSHeader::constructHeader: Multiplan spreadsheet v1 detected\n"));
			return new WPSHeader(input, input, 1, WPS_SPREADSHEET, WPS_MULTIPLAN);
		}
		if (val[0] == 0x0c && (val[1] == 0xec || val[1] == 0xed))
		{
			int vers=int(val[1]-0xeb);
			WPS_DEBUG_MSG(("WPSHeader::constructHeader: Multiplan spreadsheet v%d detected\n", vers));
			return new WPSHeader(input, input, vers, WPS_SPREADSHEET, WPS_MULTIPLAN);
		}
		// now look at the end of file
		input->seek(-1, librevenge::RVNG_SEEK_END);
		val[0]=libwps::readU8(input);
		if (val[0]==0x1a) // Dos XYWrite ends with 0x1a, Win4 XYWrite ends with fe fc fe 01 00
		{
			WPS_DEBUG_MSG(("WPSHeader::constructHeader: potential XYWrite document detected\n"));
			return new WPSHeader(input, input, 0, WPS_TEXT, WPS_XYWRITE);
		}
		if (val[0]==0)
		{
			input->seek(-5, librevenge::RVNG_SEEK_END);
			if (libwps::readU32(input)==0x1fefcfe)
			{
				WPS_DEBUG_MSG(("WPSHeader::constructHeader: potential XYWrite document detected\n"));
				return new WPSHeader(input, input, 1, WPS_TEXT, WPS_XYWRITE);
			}
		}
		return nullptr;
	}

	RVNGInputStreamPtr document_mn0(input->getSubStreamByName("MN0"));
	if (document_mn0)
	{
		// can be a mac or a pc document
		// each must contains a MM Ole which begins by 0x444e: Mac or 0x4e44: PC
		RVNGInputStreamPtr document_mm(input->getSubStreamByName("MM"));
		if (document_mm && libwps::readU16(document_mm) == 0x4e44)
		{
			WPS_DEBUG_MSG(("WPSHeader::constructHeader: Microsoft Works Mac v4 format detected\n"));
			return nullptr;
		}
		// now, look if this is a database document
		uint16_t fileMagic=libwps::readU16(document_mn0);
		if (fileMagic == 0x54FF)
		{
			WPS_DEBUG_MSG(("WPSHeader::constructHeader: Microsoft Works Database format detected\n"));
			return new WPSHeader(document_mn0, input, 4, WPS_DATABASE);
		}
		WPS_DEBUG_MSG(("WPSHeader::constructHeader: Microsoft Works v4 format detected\n"));
		return new WPSHeader(document_mn0, input, 4);
	}

	RVNGInputStreamPtr document_contents(input->getSubStreamByName("CONTENTS"));
	if (document_contents)
	{
		/* check the Works 2000/7/8 format magic */
		document_contents->seek(0, librevenge::RVNG_SEEK_SET);

		char fileMagic[8];
		int i = 0;
		for (; i<7 && !document_contents->isEnd(); i++)
			fileMagic[i] = char(libwps::readU8(document_contents.get()));
		fileMagic[i] = '\0';

		/* Works 7/8 */
		if (0 == strcmp(fileMagic, "CHNKWKS"))
		{
			WPS_DEBUG_MSG(("WPSHeader::constructHeader: Microsoft Works v8 (maybe 7) format detected\n"));
			return new WPSHeader(document_contents, input, 8);
		}

		/* Works 2000 */
		if (0 == strcmp(fileMagic, "CHNKINK"))
		{
			return new WPSHeader(document_contents, input, 5);
		}
	}
	if (input->existsSubStream("PerfectOffice_MAIN"))
	{
		RVNGInputStreamPtr stream(input->getSubStreamByName("PerfectOffice_MAIN"));
		if (stream && stream->seek(0, librevenge::RVNG_SEEK_SET) == 0 && libwps::readU16(stream)==0 &&
		        libwps::readU8(stream)==2 && libwps::readU8(stream)==0 &&
		        libwps::readU8(stream)==7 && libwps::readU8(stream)==0x10)
		{
			WPS_DEBUG_MSG(("WPSHeader::constructHeader: find a Quatto Pro wb3 spreadsheet\n"));
			return new WPSHeader(stream, input, 1003, WPS_SPREADSHEET, WPS_QUATTRO_PRO);
		}
	}
	if (input->existsSubStream("NativeContent_MAIN"))
	{
		RVNGInputStreamPtr stream(input->getSubStreamByName("NativeContent_MAIN"));
		// check that the first field has type=1, size=0xe, data=QPW9...
		if (stream && stream->seek(0, librevenge::RVNG_SEEK_SET) == 0 && libwps::readU16(stream)==1 &&
		        libwps::readU16(stream)==0xe && libwps::readU32(stream)==0x39575051)
		{
			WPS_DEBUG_MSG(("WPSHeader::constructHeader: find a Quatto Pro qpw spreadsheet\n"));
			return new WPSHeader(stream, input, 2000, WPS_SPREADSHEET, WPS_QUATTRO_PRO);
		}
	}

	/* check for a lotus 123 zip file containing a WK3 and FM3
	   or a old lotus file containing WK1 and FMT
	 */
	if (input->existsSubStream("WK1") && input->existsSubStream("FMT"))
	{
		RVNGInputStreamPtr stream(input->getSubStreamByName("WK1"));
		if (stream && stream->seek(0, librevenge::RVNG_SEEK_SET) == 0 && libwps::readU16(stream)==0 &&
		        libwps::readU8(stream)==2 && libwps::readU8(stream)==0)
		{
			WPS_DEBUG_MSG(("WPSHeader::constructHeader: find a zip Lotus spreadsheet\n"));
			return new WPSHeader(stream, input, 2, WPS_SPREADSHEET, WPS_LOTUS);
		}
	}
	if (input->existsSubStream("WK3") && input->existsSubStream("FM3"))
	{
		RVNGInputStreamPtr stream(input->getSubStreamByName("WK3"));
		if (stream && stream->seek(0, librevenge::RVNG_SEEK_SET) == 0 && libwps::readU16(stream)==0 &&
		        libwps::readU8(stream)==0x1a && libwps::readU8(stream)==0)
		{
			WPS_DEBUG_MSG(("WPSHeader::constructHeader: find a zip Lotus spreadsheet\n"));
			return new WPSHeader(stream, input, 101, WPS_SPREADSHEET, WPS_LOTUS);
		}
	}
	return nullptr;
}
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
