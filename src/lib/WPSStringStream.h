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

#ifndef WPS_STRING_STREAM_H
#define WPS_STRING_STREAM_H

#include <memory>

#include <librevenge-stream/librevenge-stream.h>

class WPSStringStreamPrivate;

/** internal class used to create a RVNGInputStream from a unsigned char's pointer

    \note this class (highly inspired from librevenge) does not
    implement the isStructured's protocol, ie. it only returns false.
 */
class WPSStringStream final: public librevenge::RVNGInputStream
{
public:
	//! constructor
	WPSStringStream(const unsigned char *data, const unsigned int dataSize);
	//! destructor
	~WPSStringStream() final;

	//! append some data at the end of the string
	void append(const unsigned char *data, const unsigned int dataSize);
	/**! reads numbytes data.

	 * \return a pointer to the read elements
	 */
	const unsigned char *read(unsigned long numBytes, unsigned long &numBytesRead) final;
	//! returns actual offset position
	long tell() final;
	/*! \brief seeks to a offset position, from actual, beginning or ending position
	 * \return 0 if ok
	 */
	int seek(long offset, librevenge::RVNG_SEEK_TYPE seekType) final;
	//! returns true if we are at the end of the section/file
	bool isEnd() final;

	/** returns true if the stream is ole

	 \sa returns always false*/
	bool isStructured() final;
	/** returns the number of sub streams.

	 \sa returns always 0*/
	unsigned subStreamCount() final;
	/** returns the ith sub streams name

	 \sa returns always 0*/
	const char *subStreamName(unsigned) final;
	/** returns true if a substream with name exists

	 \sa returns always false*/
	bool existsSubStream(const char *name) final;
	/** return a new stream for a ole zone

	 \sa returns always 0 */
	librevenge::RVNGInputStream *getSubStreamByName(const char *name) final;
	/** return a new stream for a ole zone

	 \sa returns always 0 */
	librevenge::RVNGInputStream *getSubStreamById(unsigned) final;

private:
	/// the string stream data
	std::unique_ptr<WPSStringStreamPrivate> m_data;
	WPSStringStream(const WPSStringStream &) = delete; // copy is not allowed
	WPSStringStream &operator=(const WPSStringStream &) = delete; // assignment is not allowed
};

#endif

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
