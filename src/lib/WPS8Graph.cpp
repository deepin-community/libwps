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

#include <string.h>

#include <iomanip>
#include <iostream>

#include <librevenge/librevenge.h>

#include "WPSContentListener.h"
#include "WPSEntry.h"
#include "WPSFont.h"
#include "WPSOLEObject.h"
#include "WPSOLEParser.h"
#include "WPSParagraph.h"
#include "WPSPosition.h"
#include "WPSStream.h"

#include "WPS8.h"

#include "WPS8Graph.h"


/** Internal: the structures of a WPS8Graph */
namespace WPS8GraphInternal
{
/** Internal: a complex border picture
 *
 * It consists in series of at most 8 pictures which forms the border */
struct Border
{
	//! constructor
	Border()
		: m_name("")
		, m_pictList()
		, m_parsed(false)
	{
		for (int &i : m_borderSize) i = -1;
		for (int &i : m_borderId) i = -1;
	}

	//! Internal: the border name
	std::string m_name;
	/** three value to stored the border's sizes
	 *
	 * Checkme: corner size, following by horizontal, vertical size in points?
	 */
	int m_borderSize[3];
	//! the picture uses to draw TopLeft corner, Top border, TopRight corner, Right border...
	int m_borderId[8];
	//! the border's pictures: 1 to 8 pictures..
	std::vector<WPSEmbeddedObject> m_pictList;
	//! flag to know if the border was sent to the listener
	bool m_parsed;
};

//! Internal: the state of a WPS8Graph
struct State
{
	State()
		: m_version(-1)
		, m_numPages(0)
		, m_borderMap()
		, m_ibgfMap()
		, m_pictMap()
		, m_oleMap() {}
	//! the version
	int m_version;
	//! the number page
	int m_numPages;
	//! a map id -> border
	std::map<int, Border> m_borderMap;
	//! a map id -> ibgf entry (background picture entry)
	std::map<int, WPSEntry> m_ibgfMap;
	//! a map id -> pictData
	std::map<int, WPSEmbeddedObject> m_pictMap;
	//! a map id -> OleData
	std::map<int, WPSEmbeddedObject> m_oleMap;
};
}

////////////////////////////////////////////////////////////
// constructor/destructor
////////////////////////////////////////////////////////////
WPS8Graph::WPS8Graph(WPS8Parser &parser)
	: m_listener()
	, m_mainParser(parser)
	, m_state()
	, m_asciiFile(parser.ascii())
{
	m_state.reset(new WPS8GraphInternal::State);
}

WPS8Graph::~WPS8Graph()
{ }

////////////////////////////////////////////////////////////
// update the positions and send data to the listener
////////////////////////////////////////////////////////////
int WPS8Graph::version() const
{
	if (m_state->m_version <= 0)
		m_state->m_version = m_mainParser.version();
	return m_state->m_version;
}

int WPS8Graph::numPages() const
{
	return m_state->m_numPages;
}

////////////////////////////////////////////////////////////
// update the positions and send data to the listener
////////////////////////////////////////////////////////////
void WPS8Graph::computePositions() const
{
	m_state->m_numPages = (m_state->m_pictMap.size() || m_state->m_oleMap.size()) ? 1 : 0;
}

void WPS8Graph::storeObjects(std::map<int,WPSEmbeddedObject> const &objectsMap)
{
	m_state->m_oleMap=objectsMap;
}

////////////////////////////////////////////////////////////
// find all structures which correspond to the picture
////////////////////////////////////////////////////////////
bool WPS8Graph::readStructures(RVNGInputStreamPtr const &input)
{
	auto &nameTable = m_mainParser.getNameEntryMap();

	// contains a text and 8 borders cells?
	auto pos = nameTable.lower_bound("BDR ");
	while (nameTable.end() != pos)
	{
		WPSEntry const &entry = pos++->second;
		if (!entry.hasName("BDR ")) break;
		if (!entry.hasType("WBDR")) continue;
		readBDR(input, entry);
	}

	// read IBGF zone : picture type (image background f...? )
	pos = nameTable.lower_bound("IBGF");
	while (pos != nameTable.end())
	{
		WPSEntry const &entry = pos++->second;
		if (!entry.hasName("IBGF")) break;
		if (!entry.hasType("IBGF")) continue;

		readIBGF(input, entry);
	}

	pos = nameTable.lower_bound("PICT");
	while (pos != nameTable.end())
	{
		WPSEntry const &entry = pos++->second;
		if (!entry.hasName("PICT")) break;

		readPICT(input, entry);
	}

	return true;
}

////////////////////////////////////////////////////////////
// different function to send data to a listener
////////////////////////////////////////////////////////////
bool WPS8Graph::sendObject(WPSPosition const &posi, int id, bool ole)
{
	if (m_listener.get() == nullptr)
	{
		WPS_DEBUG_MSG(("WPS8Graph::sendObject: listener is not set\n"));
		return false;
	}
	auto &map = ole ? m_state->m_oleMap : m_state->m_pictMap;

	auto pos = map.find(id);
	if (pos == map.end())
	{
		WPS_DEBUG_MSG(("WPS8Graph::sendObject: can not find %dth object\n", id));
		return false;
	}
	auto &pict = pos->second;
	pict.m_sent = true;

	Vec2f size(posi.size()), naturalSize;
	if (size.x() <= 0 || size.y()<=0)
	{
		size=pict.m_size;
		if (size.x() <= 0 || size.y()<=0)
		{
			WPS_DEBUG_MSG(("WPS8Graph::sendObject: can not find object size\n"));
			size=Vec2f(0.5,0.5);
		}
	}
	else if (pict.m_size.x() > 0 && pict.m_size.y() > 0)
		naturalSize = pict.m_size;
	WPSPosition finalPos(posi);
	finalPos.setSize(size);
	finalPos.setNaturalSize(naturalSize);
	m_listener->insertObject(finalPos, pict);
	return true;
}

bool WPS8Graph::sendIBGF(WPSPosition const &posi, int ibgfId)
{
	auto pos = m_state->m_ibgfMap.find(ibgfId);
	if (pos == m_state->m_ibgfMap.end())
	{
		WPS_DEBUG_MSG(("WPS8Graph::sendIBGF: can not find background\n"));
		return false;
	}

	WPSEntry const &ent = pos->second;
	if (!ent.hasName("PICT"))
	{
		WPS_DEBUG_MSG(("WPS8Graph::sendIBGF: unknown ibgf type\n"));
		return false;
	}
	return sendObject(posi, ent.id(), false);
}

void WPS8Graph::sendObjects(int page, int)
{
	if (page != -1) return;
	if (m_listener.get() == nullptr)
	{
		WPS_DEBUG_MSG(("WPS8Graph::sendObjects: listener is not set\n"));
		return;
	}

#ifdef DEBUG
	bool firstSend = false;
#endif
	for (int st = 0; st < 2; st++)
	{
		auto &map = (st == 0) ? m_state->m_pictMap : m_state->m_oleMap;
		auto pos = map.begin();

		while (pos != map.end())
		{
			auto &pict = pos++->second;
			if (pict.m_sent) continue;

#ifdef DEBUG
			if (!firstSend)
			{
				firstSend = true;
				WPS_DEBUG_MSG(("WPS8Graph::sendObjects: find some extra pictures\n"));
				m_listener->setFont(WPSFont::getDefault());
				m_listener->setParagraph(WPSParagraph());
				m_listener->insertEOL();
				librevenge::RVNGString message = "--------- The original document has some extra pictures: -------- ";
				m_listener->insertUnicodeString(message);
				m_listener->insertEOL();
			}
#endif

			pict.m_sent = true;
			Vec2f size=pict.m_size;
			// if we do not have the size of the data, we insert small picture
			if (size.x() <= 0 || size.y() <= 0) size.set(1.0, 1.0);
			WPSPosition position(Vec2f(),size);
			position.setNaturalSize(pict.m_size);
			position.setRelativePosition(WPSPosition::CharBaseLine);
			position.m_wrapping = WPSPosition::WDynamic;
			m_listener->insertObject(position, pict);
		}
	}

#ifdef DEBUG
	// check the border
	for (auto pos : m_state->m_borderMap)
	{
		int id = pos.first;
		bool parsed = pos.second.m_parsed;
		if (parsed) continue;
		if (!firstSend)
		{
			firstSend = true;
			m_listener->setFont(WPSFont::getDefault());
			m_listener->setParagraph(WPSParagraph());
			m_listener->insertEOL();
			librevenge::RVNGString message;
			message = "--------- The original document used some complex border: -------- ";
			m_listener->insertUnicodeString(message);
			m_listener->insertEOL();
		}
		sendBorder(id);
	}
#endif
}

void WPS8Graph::sendBorder(int borderId)
{
	if (!m_listener || m_state->m_borderMap.find(borderId) == m_state->m_borderMap.end()) return;

	auto &border = m_state->m_borderMap[borderId];
	if (border.m_parsed) return;

	border.m_parsed = true;

	m_listener->setFont(WPSFont::getDefault());
	m_listener->setParagraph(WPSParagraph());
	librevenge::RVNGString message("-------");
	message.append(border.m_name.c_str());
	message.append("---------");
	m_listener->insertUnicodeString(message);

	WPSPosition pos(Vec2f(),Vec2f(0.5,0.5));
	pos.setRelativePosition(WPSPosition::CharBaseLine);
	pos.m_wrapping = WPSPosition::WDynamic;
	for (int i = 0; i < 8; i++)
	{
		if (i == 0 || i == 3 || i == 5) m_listener->insertEOL();
		static int const wh[8] = {0, 1, 2, 7, 3, 6, 5, 4 };
		auto id = size_t(border.m_borderId[wh[i]]);
		if (border.m_pictList[id].m_size.x() > 0 &&
		        border.m_pictList[id].m_size.y() > 0)
			pos.setSize(border.m_pictList[id].m_size);
		m_listener->insertObject(pos, border.m_pictList[id]);
		if (i == 3)
		{
			message = librevenge::RVNGString("-----");
			m_listener->insertUnicodeString(message);
		}
	}
}

////////////////////////////////////////////////////////////
//
//  low level
//
////////////////////////////////////////////////////////////

// Read a PICT/MEF4 entry :  read uncompressed picture of sx*sy of rgb
bool WPS8Graph::readPICT(RVNGInputStreamPtr const &input, WPSEntry const &entry)
{
	long page_offset = entry.begin();
	long length = entry.length();
	long endPos = entry.end();

	WPSEmbeddedObject pict;

	// too small, we return 0
	if (length < 24) return false;

	if (!entry.hasType("MEF4"))
	{
		WPS_DEBUG_MSG(("WPS8Graph::readPICT: warning: PICT name=%s, type=%s\n",
		               entry.name().c_str(), entry.type().c_str()));
		return false;
	}

	input->seek(page_offset, librevenge::RVNG_SEEK_SET);
	std::string name;
	for (int i = 0; i < 4; i++) name += char(libwps::readU8(input));
	if (strncmp("MEF4", name.c_str(), 4))
	{
		WPS_DEBUG_MSG(("WPS8Graph::readPICT: warning: PICT unknown header=%s\n",
		               name.c_str()));
		return false;
	}

	libwps::DebugStream f;
	f << "Header:";
	// unkn0, unkn1 : always 0 ?
	for (int i = 0; i < 2; i++)
	{
		auto val = long(libwps::readU32(input)); // is one the number of meta file ?
		if (val) f << "unknA" << i << "=" << std::hex << val << std::dec << ",";
	}
	pict.m_size.setX(float(libwps::readU32(input))/914400.f);
	pict.m_size.setY(float(libwps::readU32(input))/914400.f);
	f << "pSz(inches)=" << pict.m_size << ",";

	ascii().addPos(page_offset);
	ascii().addNote(f.str().c_str());
	entry.setParsed();

	auto stream=std::make_shared<WPSStream>(input, ascii());
	bool ok=WPSOLEObject::readWMF(stream, pict, endPos); // we may also want to check if this is a emf file
	if (ok)
	{
		if (m_state->m_pictMap.find(entry.id()) != m_state->m_pictMap.end())
			WPS_DEBUG_MSG(("WPS8Graph::readPICT: Pict entry %d already exists\n",  entry.id()));
		else
			m_state->m_pictMap[entry.id()] = pict;
	}
	else
		input->seek(page_offset+24, librevenge::RVNG_SEEK_SET);

	if (input->tell() != endPos)
	{
		ascii().addPos(input->tell());
		ascii().addNote("PICT###");
	}


	return ok;
}


// reads a IBGF zone
// Warning: only seems very simple IBGF, complex may differ
bool WPS8Graph::readIBGF(RVNGInputStreamPtr const &input, WPSEntry const &entry)
{
	libwps::DebugStream f;
	if (!entry.hasType(entry.name()))
	{
		WPS_DEBUG_MSG(("WPS8Graph::readIBGF: warning: IBGF name=%s, type=%s\n",
		               entry.name().c_str(), entry.type().c_str()));
		return false;
	}

	long page_offset = entry.begin();
	long length = entry.length();

	if (length != 26)
	{
		WPS_DEBUG_MSG(("WPS8Graph::readIBGF: IBGF length=0x%lx\n", static_cast<unsigned long>(length)));
		return false;
	}

	entry.setParsed();
	input->seek(page_offset, librevenge::RVNG_SEEK_SET);
	std::string name;
	for (int i = 0; i < 4; i++)
	{
		auto val = char(libwps::readU8(input));
		if (val >= '0' && val <= 'z')
		{
			name+=val;
			continue;
		}
		WPS_DEBUG_MSG(("WPS8Graph::readIBGF: invalid name %s\n", name.c_str()));
		return false;
	}

	auto id = int(libwps::read16(input));

	WPSEntry res;
	res.setName(name);
	res.setId(id);

	// name = PICT ?
	f << "indexEntry='" << name << "':" << id;

	for (int i = 0; i < 10; i++)
	{
		auto val = libwps::read16(input);
		if (val) f << ", f"<<i <<"=" << val;
	}

	ascii().addPos(page_offset);
	ascii().addNote(f.str().c_str());

	if (m_state->m_ibgfMap.find(entry.id()) != m_state->m_ibgfMap.end())
		WPS_DEBUG_MSG(("WPS8Graph::readIBGF: warning: IBGF entry %d already exists\n",
		               entry.id()));
	else
		m_state->m_ibgfMap[entry.id()] = res;

	return true;
}

// BDR/WBDR Code : read a BDR Code
bool WPS8Graph::readBDR(RVNGInputStreamPtr const &input, WPSEntry const &entry)
{
	long page_offset = entry.begin();
	long length = entry.length();
	long endPos = entry.end();

	// too small, we return 0
	if (length < 20)
	{
		WPS_DEBUG_MSG(("WPS8Graph::readBDR: length=%ld is too short\n", length));
		return false;
	}

	if (!entry.hasType("WBDR"))
	{
		WPS_DEBUG_MSG(("WPS8Graph::readBDR: warning: BDR name=%s, type=%s\n",
		               entry.name().c_str(), entry.type().c_str()));
		return false;
	}

	entry.setParsed();
	input->seek(page_offset, librevenge::RVNG_SEEK_SET);

	WPS8GraphInternal::Border border;
	border.m_name = entry.extra();
	libwps::DebugStream f;
	if (!border.m_name.empty()) f << "Header:borderName='"<<border.m_name << "',";

	// 1, unk? = ~ 1/20*following size ?
	for (int i = 0; i < 2; i++) f << libwps::read16(input) << ",";
	// 3*size (corner size, following by x/y steps) in points?
	f << "sizes=(";
	for (int &i : border.m_borderSize)
	{
		i = libwps::read16(input);
		f << i << ",";
	}
	f << "),";

	bool ok = true;
	// ids of picture which corresponds to the frame borders:
	// TL, T, TR, R, BR, B, BL, L ?
	f << "id=(";
	for (int &i : border.m_borderId)
	{
		auto id = int(libwps::read8(input));
		if (id < 0 || id> 8)
		{
			ok = false;
			break;
		}
		i = id;
		f << id << ",";
	}
	f << "),";
	auto N = int(libwps::read8(input));
	f << "Nbdr="<<N << ",";
	auto unkn = int(libwps::read8(input));
	if (unkn) f << "###unkn=" << unkn << ",";

	if (!ok || N < 0 || N > 8 || 20+N*4 > length)
	{
		WPS_DEBUG_MSG(("WPS8Graph::readBDR: can not read the pictures\n"));
		f << "###";
		ascii().addPos(page_offset);
		ascii().addNote(f.str().c_str());
		return false;
	}

	long debPos = page_offset+4*N+20;
	f << "ptr(" << std::hex;

	std::vector<long> listPtrs;
	listPtrs.push_back(debPos);

	for (int i = 0; i < N; i++)
	{
		f << debPos << ",";
		auto sz = long(libwps::readU32(input));
		debPos += sz;
		if (sz < 0 || debPos > endPos)
		{
			ok = false;
			break;
		}
		listPtrs.push_back(debPos);
	}
	f << debPos << ")," << std::dec;

	ascii().addPos(page_offset);
	ascii().addNote(f.str().c_str());
	if (!ok) return false;

	if (listPtrs[size_t(N)] != endPos)
	{
		ascii().addPos(listPtrs[size_t(N)]);
		ascii().addNote("###BDR");
	}

	auto stream=std::make_shared<WPSStream>(input, ascii());
	for (size_t bd = 0; bd < size_t(N); bd++)
	{
		long debP = listPtrs[bd];
		long endP = listPtrs[bd+1];

		input->seek(debP, librevenge::RVNG_SEEK_SET);
		f.str("");
		f << "BDR(" << bd << "):";

		if (debP+12 > endP || libwps::readU32(input) != 0x4154454d)   // META
		{
			WPS_DEBUG_MSG(("WPS8Graph::readBDR: unknown type can not read the picture %d\n", int(bd)));
			f << "###";
			ascii().addPos(debP);
			ascii().addNote(f.str().c_str());
			ok = false;
			continue;
		}

		WPSEmbeddedObject pict;
		int dim[4];
		for (int &i : dim)
		{
			i = int(libwps::read16(input));
			f << i << ",";
		}
		// final picture size
		pict.m_size.set(float(dim[2]-dim[0])/1440.f, float(dim[3]-dim[1])/1440.f);
		bool correct=WPSOLEObject::readWMF(stream, pict, endP);
		if (!correct) f << "###";
		ascii().addPos(debP);
		ascii().addNote(f.str().c_str());

		if (!correct)
		{
			ok = false;
			continue;
		}

		if (ok) border.m_pictList.push_back(pict);
	}

	if (!ok) return false;
	m_state->m_borderMap[entry.id()] = border;
	return true;
}

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
