/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: t; c-basic-offset: 4 -*- */
/* libwps
 * Version: MPL 2.0 / LGPLv2.1+
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *
 * Major Contributor(s):
 * Copyright (C) 2009, 2011 Alonso Laurent (alonso@loria.fr)
 * Copyright (C) 2006, 2007 Andrew Ziem
 * Copyright (C) 2004-2006 Fridrich Strba (fridrich.strba@bluewin.ch)
 * Copyright (C) 2004 Marc Maurer (uwog@uwog.net)
 * Copyright (C) 2003-2005 William Lachance (william.lachance@sympatico.ca)
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

/*
 *  freely inspired from istorage :
 * ------------------------------------------------------------
 *      Generic OLE Zones furnished with a copy of the file header
 *
 * Compound Storage (32 bit version)
 * Storage implementation
 *
 * This file contains the compound file implementation
 * of the storage interface.
 *
 * Copyright 1999 Francis Beaudet
 * Copyright 1999 Sylvain St-Germain
 * Copyright 1999 Thuy Nguyen
 * Copyright 2005 Mike McCormack
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * ------------------------------------------------------------
 */

#ifndef WPS_OLE_PARSER_H
#define WPS_OLE_PARSER_H

#include <functional>
#include <string>
#include <map>
#include <vector>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WPSDebug.h"
#include "WPSPosition.h"

namespace libwps
{
class Storage;
}

namespace WPSOLEParserInternal
{
class CompObj;
struct State;
}

/** \brief a class used to parse some basic oles
    Tries to read the different ole parts and stores their contents in form of picture.
 */
class WPSOLEParser
{
public:
	/** constructor
	    \param mainName name of the main ole, we must avoid to parse */
	WPSOLEParser(const std::string &mainName,
	             libwps_tools_win::Font::Type fontType,
	             std::function<int(std::string const &)> const &dirToIdFunc=getIdFromDirectory);

	/** destructor */
	~WPSOLEParser();

	/** tries to parse basic OLE (excepted mainName)
	    \return false if fileInput is not an Ole file */
	bool parse(RVNGInputStreamPtr fileInput);
	/** update the meta data, using information find in SummaryInformation */
	void updateMetaData(librevenge::RVNGPropertyList &metaData) const;

	//! returns the list of unknown ole
	std::vector<std::string> const &getNotParse() const;
	//! returns the list of data positions which have been read
	std::map<int,WPSEmbeddedObject> const &getObjectsMap() const;

protected:
	//! the summary information and the doc summary information
	bool readSummaryInformation(RVNGInputStreamPtr input, std::string const &oleName,
	                            librevenge::RVNGPropertyList &pList,
	                            libwps::DebugFile &ascii) const;
	//! try to read a summary property
	bool readSummaryProperty(RVNGInputStreamPtr input, long endPos, int type,
	                         libwps::DebugFile &ascii, libwps::DebugStream &f) const;
	//! try to read a summary property: type 1e
	bool readSummaryPropertyString(RVNGInputStreamPtr input, long endPos, int type, librevenge::RVNGString &string,
	                               libwps::DebugStream &f) const;
	//! try to read a summary property: type 2,3,9,12,
	static bool readSummaryPropertyLong(RVNGInputStreamPtr input, long endPos, int type, long &value,
	                                    libwps::DebugStream &f);

	//!  the "Ole" small structure : unknown contain
	static bool readOle(RVNGInputStreamPtr &ip, std::string const &oleName,
	                    libwps::DebugFile &ascii);
	//!  the "MM" small structure : seems to contain the file versions
	static bool readMM(RVNGInputStreamPtr &input, std::string const &oleName,
	                   libwps::DebugFile &ascii);
	//!  the "ObjInfo" small structure : seems to contain 3 ints=0,3,4
	static bool readObjInfo(RVNGInputStreamPtr &input, std::string const &oleName,
	                        libwps::DebugFile &ascii);
	//!  the "CompObj" contains : UserType,ClipName,ProgIdName
	bool readCompObj(RVNGInputStreamPtr &ip, std::string const &oleName,
	                 libwps::DebugFile &ascii);

	/** the OlePres001 seems to contain standart picture file and size */
	static bool isOlePres(RVNGInputStreamPtr &ip, std::string const &oleName);
	/** extracts the picture of OlePres001 if it is possible */
	static bool readOlePres(RVNGInputStreamPtr &ip, WPSEmbeddedObject &obj, libwps::DebugFile &ascii);

	//! theOle10Native : basic Windows' picture, with no size
	static bool isOle10Native(RVNGInputStreamPtr &ip, std::string const &oleName);
	/** extracts the picture if it is possible */
	static bool readOle10Native(RVNGInputStreamPtr &ip, WPSEmbeddedObject &obj, libwps::DebugFile &ascii);

	/** \brief the Contents : in general a picture : a PNG, an JPEG, a basic metafile,
	 * I find also a Word art picture, which are not sucefully read
	 */
	static bool readContents(RVNGInputStreamPtr &input, std::string const &oleName,
	                         WPSEmbeddedObject &obj, libwps::DebugFile &ascii);

	/** the CONTENTS : seems to store a header size, the header
	 * and then a object in EMF (with the same header)...
	 * \note I only find such lib in 2 files, so the parsing may be incomplete
	 *  and many such Ole rejected
	 */
	static bool readCONTENTS(RVNGInputStreamPtr &input, std::string const &oleName,
	                         WPSEmbeddedObject &obj, libwps::DebugFile &ascii);

	//!  the "MN0" small structure : can contains a WKS file...
	static bool readMN0AndCheckWKS(RVNGInputStreamPtr &input, std::string const &oleName,
	                               WPSEmbeddedObject &obj, libwps::DebugFile &ascii);

	/** the default function which uses the last integer of dirName to return the final id,
		ie. it converts "MatOST/MatadorObject1", "Object 1" in 1*/
	static int getIdFromDirectory(std::string const &dirName);

	//! if filled, does not parse content with this name
	std::string m_avoidOLE;
	//! the main state
	std::shared_ptr<WPSOLEParserInternal::State> m_state;
};

#endif
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
