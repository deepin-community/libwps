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

#include <iomanip>
#include <iostream>

#include <librevenge/librevenge.h>

#include "libwps_internal.h"
#include "WPSContentListener.h"
#include "WPSEntry.h"
#include "WPSFont.h"
#include "WPSOLEObject.h"
#include "WPSOLEParser.h"
#include "WPSParagraph.h"
#include "WPSPosition.h"
#include "WPSStream.h"

#include "WPS4.h"

#include "WPS4Graph.h"

/** Internal: the structures of a WPS4Graph */
namespace WPS4GraphInternal
{
//! Internal: the state of a WPS4Graph
struct State
{
	State()
		: m_version(-1)
		, m_numPages(0)
		, m_idToObjectMap()
	{}
	//! the version
	int m_version;
	//! the number page
	int m_numPages;

	//! the map id to objects
	std::map<int,WPSEmbeddedObject> m_idToObjectMap;
};
}

// constructor/destructor
WPS4Graph::WPS4Graph(WPS4Parser &parser)
	: m_listener()
	, m_mainParser(parser)
	, m_state(new WPS4GraphInternal::State)
	, m_asciiFile(parser.ascii())
{
}

WPS4Graph::~WPS4Graph()
{
}

// small functions: version/numpages/update position
int WPS4Graph::version() const
{
	if (m_state->m_version <= 0)
		m_state->m_version = m_mainParser.version();
	return m_state->m_version;
}

int WPS4Graph::numPages() const
{
	return m_state->m_idToObjectMap.empty() ? 0 : 1;
}

void WPS4Graph::computePositions() const
{
}

// update the positions and send data to the listener
void WPS4Graph::storeObjects(std::map<int,WPSEmbeddedObject> const &objectsMap)
{
	for (auto it : objectsMap)
	{
		if (m_state->m_idToObjectMap.find(it.first)!=m_state->m_idToObjectMap.end())
			continue;
		m_state->m_idToObjectMap[it.first]=it.second;
	}
}

// send object
void WPS4Graph::sendObject(WPSPosition const &position, int id)
{
	if (m_listener.get() == nullptr)
	{
		WPS_DEBUG_MSG(("WPS4Graph::sendObject: listener is not set\n"));
		return;
	}

	auto it=m_state->m_idToObjectMap.find(id);
	if (it==m_state->m_idToObjectMap.end())
	{
		WPS_DEBUG_MSG(("WPS4Graph::sendObject: can not find %d object\n", id));
		return;
	}
	auto const &obj=it->second;
	obj.m_sent=true;
	WPSPosition posi(position);
	posi.setNaturalSize(obj.m_size);
	m_listener->insertObject(posi, obj);
}

void WPS4Graph::sendObjects(int page)
{
	if (page != -1) return;
	if (m_listener.get() == nullptr)
	{
		WPS_DEBUG_MSG(("WPS4Graph::sendObjects: listener is not set\n"));
		return;
	}

#ifdef DEBUG
	bool firstSend = false;
#endif
	for (auto it : m_state->m_idToObjectMap)
	{
		if (it.second.m_sent) continue;
#ifdef DEBUG
		if (!firstSend)
		{
			firstSend = true;
			WPS_DEBUG_MSG(("WPS4Graph::sendObjects: find some extra pictures\n"));
			m_listener->setFont(WPSFont::getDefault());
			m_listener->setParagraph(WPSParagraph());
			m_listener->insertEOL();
			librevenge::RVNGString message = "--------- The original document has some extra pictures: -------- ";
			m_listener->insertUnicodeString(message);
			m_listener->insertEOL();
		}
#endif
		// as we do not have the size of the data, we insert small picture
		auto const &obj=it.second;
		obj.m_sent=true;
		WPSPosition pos(Vec2f(),obj.m_size==Vec2f() ? Vec2f(1.,1.) : obj.m_size);
		pos.setRelativePosition(WPSPosition::CharBaseLine);
		pos.m_wrapping = WPSPosition::WDynamic;
		m_listener->insertObject(pos, obj);
	}
}

////////////////////////////////////////////////////////////
//  low level
////////////////////////////////////////////////////////////
int WPS4Graph::readObject(RVNGInputStreamPtr input, WPSEntry const &entry)
{
	if (!entry.valid() || entry.length() <= 4)
	{
		WPS_DEBUG_MSG(("WPS4Graph::readObject: invalid object\n"));
		return -1;
	}
	long pos=entry.begin();
	long endPos = entry.end();
	input->seek(pos, librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f;
	f << "ZZEOBJ" << entry.id() << "(Contents):";
	int type = libwps::readU16(input);
	int oleId=-1;
	if (type == 0x4f4d && entry.length()>=8)   // OM
	{
		oleId = libwps::read16(input);
		f << "Ole" << oleId << ",";
		int numData=entry.length()>=10 ? 3 : 2;
		for (int i=0; i<numData; ++i)
		{
			auto val = int(libwps::readU16(input));
			if (val) f << "f" << i << "=" << std::hex << val << std::dec << ",";
		}
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
	}
	else
	{
		WPSEmbeddedObject object;
		auto stream=std::make_shared<WPSStream>(input, ascii());
		bool ok=false;
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		if (type==0x501)
			ok=WPSOLEObject::readOLE(stream, object, endPos);
		else // type==8 seems be followed by a standart metafile
			ok=WPSOLEObject::readMetafile(stream, object, endPos,type!=8);
		if (ok)
		{
			if (m_state->m_idToObjectMap.empty())
				oleId=0;
			else
			{
				auto it=m_state->m_idToObjectMap.end();
				oleId=(--it)->first+1;
			}
			m_state->m_idToObjectMap[oleId]=object;
		}
	}
	if (input->tell() != endPos)
	{
		if (type!=0x4f4d || m_state->m_idToObjectMap.find(oleId)==m_state->m_idToObjectMap.end())
		{
			WPS_DEBUG_MSG(("WPS4Graph::readObject: find extra data\n"));
			ascii().addPos(input->tell());
			ascii().addNote("ZZEOBJ(Contents):##extra");
			ascii().addPos(endPos);
			ascii().addNote("_");
		}
		else
		{
			ascii().addPos(input->tell());
			ascii().addNote("_");
			ascii().addPos(endPos);
			ascii().addNote("_");
		}
	}
	return oleId;
}
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
