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
 */

#ifndef WPS_OLE_OBJECT_H
#define WPS_OLE_OBJECT_H

#include <string>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"
#include "libwps_tools_win.h"

/** \brief a class used to parse/retrieve an OLE object
 */
class WPSOLEObject
{
public:
	/** constructor knowing the file stream */
	WPSOLEObject() {}

	/** destructor */
	~WPSOLEObject() {}

	//! try to read a metafile data
	static bool readMetafile(std::shared_ptr<WPSStream> stream, WPSEmbeddedObject &object,
	                         long endPos=-1, bool strict=false);
	//! try to read a OLE: 0x0105 ...
	static bool readOLE(std::shared_ptr<WPSStream> stream, WPSEmbeddedObject &object, long endPos=-1);
	/** try to read a wmf file: 0x0[12]00 0xXX00(with XX>=9)

		\see http://www.fileformat.info/format/wmf/egff.htm */
	static bool readWMF(std::shared_ptr<WPSStream> stream, WPSEmbeddedObject &object,long endPos=-1);

protected:
	//! try to read a stream
	static bool readString(std::shared_ptr<WPSStream> stream, std::string &name, long endPos);
	//! check if a wmf header
	static bool checkIsWMF(std::shared_ptr<WPSStream> stream, long endPos);

	//! try to read a embedded OLE: 0x0105 0000 0200 ...
	static bool readEmbeddedOLE(std::shared_ptr<WPSStream> stream, WPSEmbeddedObject &object, long endPos);
	//! try to read a static OLE: 0x0105 0000 0[35]00 ...
	static bool readStaticOLE(std::shared_ptr<WPSStream> stream, WPSEmbeddedObject &object, long endPos);
};

#endif
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
