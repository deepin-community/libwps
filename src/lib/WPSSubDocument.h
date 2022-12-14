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
 *
 * For further information visit http://libwps.sourceforge.net
 */

#ifndef WPSSUBDOCUMENT_H
#define WPSSUBDOCUMENT_H

#include "libwps_internal.h"

class WPSContentListener;
class WPSParser;

/** virtual class to define a sub document */
class WPSSubDocument
{
public:
	/// constructor
	WPSSubDocument(RVNGInputStreamPtr const &input, int id=0);
	/// destructor
	virtual ~WPSSubDocument();

	/// returns the input
	RVNGInputStreamPtr &getInput()
	{
		return m_input;
	}
	/// get the identificator
	int id() const
	{
		return m_id;
	}
	/// set the identificator
	void setId(int i)
	{
		m_id = i;
	}

	/// an operator =
	virtual bool operator==(std::shared_ptr<WPSSubDocument> const &doc) const;
	bool operator!=(std::shared_ptr<WPSSubDocument> const &doc) const
	{
		return !operator==(doc);
	}

protected:
	RVNGInputStreamPtr m_input;
	int m_id;
private:
	WPSSubDocument(const WPSSubDocument &) = delete;
	WPSSubDocument &operator=(const WPSSubDocument &) = delete;

};
#endif
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */

