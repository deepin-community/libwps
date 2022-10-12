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
 * Copyright (C) 2004 Marc Maurer (uwog@uwog.net)
 * Copyright (C) 2004-2006 Fridrich Strba (fridrich.strba@bluewin.ch)
 *
 * For minor contributions see the git repository.
 *
 * Alternatively, the contents of this file may be used under the terms
 * of the GNU Lesser General Public License Version 2.1 or later
 * (LGPLv2.1+), in which case the provisions of the LGPLv2.1+ are
 * applicable instead of those above.
 */

#include <stdlib.h>
#include <string.h>

#include <map>
#include <sstream>
#include <utility>
#include <vector>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"

#include "WPSContentListener.h"
#include "WPSDebug.h"
#include "WPSEntry.h"
#include "WPSFont.h"
#include "WPSHeader.h"
#include "WPSPageSpan.h"
#include "WPSParagraph.h"
#include "WPSPosition.h"
#include "WPSTextSubDocument.h"

#include "PocketWord.h"

namespace PocketWordParserInternal
{

//! the state of PocketWord
struct State
{
	//! constructor
	explicit State(libwps_tools_win::Font::Type encoding)
		: m_eof(-1)
		, m_version(6)
		, m_encoding(encoding)
		, m_badFile(false)
		, m_typeEntryList()
		, m_idToEntryMap()
		, m_typeToEntryMap()
		, m_pageSpan()
		, m_idToFontNameMap()
	{
		m_pageSpan.setMarginLeft(0.1);
		m_pageSpan.setMarginRight(0.1);
		m_pageSpan.setMarginTop(0.1);
		m_pageSpan.setMarginBottom(0.1);
	}
	//! returns an entry corresponding to a file identificator
	WPSEntry getEntry(int id, int &type) const
	{
		auto it=m_idToEntryMap.find(id);
		if (it==m_idToEntryMap.end() || it->second >= m_typeEntryList.size())
		{
			WPS_DEBUG_MSG(("PocketWordParserInternal::State::getEntry can not find entry for id=%d\n", id));
			type=-1;
			return WPSEntry();
		}
		auto const &typeEntrie=m_typeEntryList[it->second];
		type=typeEntrie.first;
		typeEntrie.second.setParsed(true);
		return typeEntrie.second;
	}
	//! try to retrieve a color
	bool getColor(int id, WPSColor &col) const
	{
		static WPSColor colors[16]=
		{
			WPSColor(0,0,0), WPSColor(128,128,128), WPSColor(192,192,192), WPSColor(255,255,255),
			WPSColor(255,0,0), WPSColor(0,255,0), WPSColor(0,0,255), WPSColor(0,255,255),
			WPSColor(255,0,255), WPSColor(0,255,0,255), WPSColor(128,0,0), WPSColor(0,128,0),
			WPSColor(0,0,128), WPSColor(0,128,128),	WPSColor(128,0,128), WPSColor(0,128,0,128)
		};
		if (id<0 || id>=16)
			return false;
		col=colors[id];
		return true;
	}
	//! the last file position
	long m_eof;
	//! the file version
	int m_version;
	//! the file encoding
	libwps_tools_win::Font::Type m_encoding;
	//! some file does not respect the unique indice conditions, ...
	bool m_badFile;

	//! the list of type, entry
	std::vector<std::pair<int, WPSEntry> > m_typeEntryList;
	//! the file id to (type, entry) index
	std::map<int, size_t > m_idToEntryMap;
	//! a type to a list of (type,entry) indices
	std::map<int, std::vector<size_t> > m_typeToEntryMap;

	//! the page span
	WPSPageSpan m_pageSpan;
	//! the correspondance between index and font name
	std::map<int, librevenge::RVNGString> m_idToFontNameMap;
};

}

// constructor, destructor
PocketWordParser::PocketWordParser(RVNGInputStreamPtr &input, WPSHeaderPtr &header, libwps_tools_win::Font::Type encoding)
	: WPSParser(input, header)
	, m_listener()
	, m_state()
{
	m_state.reset(new PocketWordParserInternal::State(encoding));
}

PocketWordParser::~PocketWordParser()
{
}

bool PocketWordParser::checkFilePosition(long pos) const
{
	if (m_state->m_eof < 0)
	{
		RVNGInputStreamPtr input = const_cast<PocketWordParser *>(this)->getInput();
		long actPos = input->tell();
		input->seek(0, librevenge::RVNG_SEEK_END);
		m_state->m_eof=input->tell();
		input->seek(actPos, librevenge::RVNG_SEEK_SET);
	}
	return pos>=0 && pos <= m_state->m_eof;
}

// listener
std::shared_ptr<WPSContentListener> PocketWordParser::createListener(librevenge::RVNGTextInterface *interface)
{
	// look for page dimensions
	auto it=m_state->m_typeToEntryMap.find(7);
	if (it==m_state->m_typeToEntryMap.end() || it->second.empty() ||
	        it->second[0]>=m_state->m_typeEntryList.size())
	{
		WPS_DEBUG_MSG(("PocketWordParser::createListener: can not find the page dimensions\n"));
	}
	else
	{
		if (it->second.size()>1)
		{
			WPS_DEBUG_MSG(("PocketWordParser::createListener: using multiple page dimensions is unimplemented\n"));
		}
		WPSEntry const &dEntry=m_state->m_typeEntryList[it->second[0]].second;
		if (dEntry.valid())
			readPageDims(dEntry);
	}
	std::vector<WPSPageSpan> pageList;
	WPSPageSpan ps(m_state->m_pageSpan);
	pageList.push_back(ps);
	auto listener=std::make_shared<WPSContentListener>(pageList, interface);
	return listener;
}

////////////////////////////////////////////////////////////
// main funtions to parse a document
////////////////////////////////////////////////////////////

// main function to parse the document
void PocketWordParser::parse(librevenge::RVNGTextInterface *documentInterface)
{
	RVNGInputStreamPtr input=getInput();
	if (!input)
	{
		WPS_DEBUG_MSG(("PocketWordParser::parse: does not find main input\n"));
		throw (libwps::ParseException());
	}
	if (!checkHeader(nullptr, true))
		throw (libwps::ParseException());

	ascii().setStream(input);
	ascii().open("main-1");
	try
	{
		checkHeader(nullptr);
		if (!createZones())
			throw (libwps::ParseException());
		m_listener=createListener(documentInterface);
		if (!m_listener)
		{
			WPS_DEBUG_MSG(("PocketWordParser::parse: can not create the listener\n"));
			throw (libwps::ParseException());
		}
		m_listener->startDocument();
		sendData();
#ifdef DEBUG
		checkUnparsed();
#endif
		m_listener->endDocument();
	}
	catch (...)
	{
		WPS_DEBUG_MSG(("PocketWordParser::parse: exception catched when parsing the main document\n"));
		throw (libwps::ParseException());
	}
	m_listener.reset();
	ascii().reset();
}

bool PocketWordParser::createZones()
{
	auto input=getInput();
	if (!input)
		throw (libwps::ParseException());
	int lastId=-1;
	while (checkFilePosition(input->tell()+6))
	{
		long pos=input->tell();
		int type=libwps::readU16(input);
		int id=int(libwps::readU16(input));
		long len=long(libwps::readU16(input));
		if (type!=85)
			len*=4;
		else
		{
			// checkme: I always find id=0, so id may store the hi of len
			len=long(len + (id<<16));
			id=65536+lastId;
		}
		if (len<0 || !checkFilePosition(pos+6+len))
		{
			input->seek(pos, librevenge::RVNG_SEEK_SET);
			break;
		}
		WPSEntry entry;
		entry.setBegin(pos+6);
		entry.setLength(len);
		entry.setId(id);

		size_t newId=m_state->m_typeEntryList.size();
		m_state->m_typeEntryList.push_back(std::make_pair(type,entry));

		// normally, one id by zone, but some file the same id for each zone which shared the same type :-~
		//           for instance, PocketWord files created by OpenOffice/LibreOffice :-~
		if (m_state->m_idToEntryMap.find(id)==m_state->m_idToEntryMap.end())
			m_state->m_idToEntryMap[id]=newId;
		else if (!m_state->m_badFile)
		{
			WPS_DEBUG_MSG(("PocketWordParser::createZones: this file contains zone with similar indices!!!\n"));
			m_state->m_badFile=true;
		}

		if (m_state->m_typeToEntryMap.find(type)==m_state->m_typeToEntryMap.end())
			m_state->m_typeToEntryMap[type]=std::vector<size_t>();
		m_state->m_typeToEntryMap.find(type)->second.push_back(newId);
		lastId=id;
		input->seek(entry.end(), librevenge::RVNG_SEEK_SET);
	}
	if (!input->isEnd())
	{
		ascii().addPos(input->tell());
		ascii().addNote("Entries(Bad):###");
	}
	return true;
}

void PocketWordParser::checkUnparsed()
{
	auto input=getInput();
	for (size_t i=0; i<m_state->m_typeEntryList.size(); ++i)
	{
		auto const &type=m_state->m_typeEntryList[i].first;
		auto const &entry=m_state->m_typeEntryList[i].second;
		if (entry.isParsed())
			continue;
		bool done=false;
		std::string name;
		switch (type)
		{
		case 0:
			done=readFontNames(entry);
			break;
		case 7:
			done=readPageDims(entry);
			break;
		case 8:
			done=readUnkn8(entry);
			break;
		case 20:
		case 21:
			done=readUnkn2021(entry, type);
			break;
		case 64:
		{
			std::vector<int> paraId;
			done=readParagraphList(entry, paraId);
			break;
		}
		case 65:
			done=m_listener && sendParagraph(i);
			break;
		case 66:
			done=readParagraphDims(entry);
			break;
		case 67:
			done=readParagraphUnkn(entry);
			break;
		case 84:   // find only sound, but this may be some picture
		{
			WPSEmbeddedObject object;
			done=readSound(entry, object);
			break;
		}
		case 130:
			name="End";
			break;
		default:
			break;
		}
		if (!done)
		{
			libwps::DebugStream f;
			if (name.empty())
				f << "Entries(Zone" << type << "A):";
			else
				f << "Entries(" << name << "):";
			f << "id=" << entry.id() << ",";
			ascii().addPos(entry.begin());
			ascii().addNote(f.str().c_str());
		}
		if (input->tell()!=entry.end())
			ascii().addDelimiter(input->tell(), '|');
	}
}

bool PocketWordParser::readFontNames(WPSEntry const &entry)
{
	entry.setParsed(true);

	auto input=getInput();
	long pos=entry.begin();
	input->seek(pos, librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f;
	f << "Entries(FontNames):id=" << entry.id() << ",";
	if ((entry.length()%80) != 16)
	{
		WPS_DEBUG_MSG(("PocketWordParser::readFontNames: the length seems bad\n"));
		f << "###";
		ascii().addPos(pos-6);
		ascii().addNote(f.str().c_str());
		return true;
	}
	int val;
	for (int i=0; i<2; ++i)   // f3=1|2, f5=5|6
	{
		val=int(libwps::readU16(input));
		int const expected[]= {0,0xa};
		if (val==expected[i]) continue;
		f << "f" << i << "=" << val << ",";
	}
	int numFonts=int(libwps::readU16(input));
	if (numFonts!=1)
		f << "num[fonts]=" <<numFonts << ",";
	for (int i=0; i<5; ++i)   // g0=..*100?, g1=..x100?
	{
		val=int(libwps::readU16(input));
		if (val==0)
			continue;
		f << "g" << i << "=" << val << ",";
	}
	if (numFonts<=0 || 80*numFonts+16>entry.length())
	{
		WPS_DEBUG_MSG(("PocketWordParser::readFontNames: the number of fonts seems bad\n"));
		f << "###N,";
		numFonts=int(entry.length()/80);
	}
	ascii().addPos(pos-6);
	ascii().addNote(f.str().c_str());
	for (int i=0; i< numFonts; ++i)
	{
		pos=input->tell();
		f.str("");
		f << "FontNames-" << i << ":";

		int fId=int(libwps::readU16(input));
		f << "fId=" << fId << ",";
		val=int(libwps::readU16(input));
		if (val!=1)
			f << "f0=" << val << ",";
		val=int(libwps::readU16(input)); // 0|1|2|8
		if (val)
			f << "f1=" << val << ",";
		val=int(libwps::readU16(input));
		if (val) // 0|1|202|2cc
			f << "fl=" << std::hex << val << std::dec << ",";
		for (int j=0; j<4; ++j)   // f2=0-3
		{
			val=int(libwps::readU16(input));
			if (val)
				f << "f" << j+2 << "=" << val << ",";
		}
		librevenge::RVNGString name;
		for (int j=0; j<32; ++j)
		{
			val=int(libwps::readU16(input));
			if (!val)
				break;
			libwps::appendUnicode(uint32_t(val), name);
		}
		f << name.cstr();
		if (m_state->m_idToFontNameMap.find(fId)!=m_state->m_idToFontNameMap.end())
		{
			WPS_DEBUG_MSG(("PocketWordParser::readFontNames: a font with id=%d already exists\n", fId));
			f << "###fId,";
		}
		else
			m_state->m_idToFontNameMap[fId]=name;
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		input->seek(pos+80, librevenge::RVNG_SEEK_SET);
	}
	return true;
}

bool PocketWordParser::readPageDims(WPSEntry const &entry)
{
	auto input=getInput();
	if (!input)
		throw (libwps::ParseException());
	entry.setParsed(true);
	long pos=entry.begin();
	input->seek(pos, librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f;
	f << "Entries(PageDims):id=" << entry.id() << ",";
	if (entry.length()!=84)
	{
		WPS_DEBUG_MSG(("PocketWordParser::readPageDims: the length seems bad\n"));
		f << "###";
		ascii().addPos(pos-6);
		ascii().addNote(f.str().c_str());
		return true;
	}
	int val=int(libwps::readU16(input)); // 10|42|50
	if (val)
		f << "fl0=" << std::hex << val << std::dec << ",";
	val=int(libwps::readU16(input));
	if (val!=1)
		f << "f0=" << val << ",";
	double dims[2];
	for (auto &d : dims)
	{
		d=double(libwps::readU16(input));
		d+=double(libwps::readU16(input))/65536;
	}
	f << "dim=" << dims[0]/20 << "x" << dims[1]/20 << ",";
	if (dims[0]>1440)
		m_state->m_pageSpan.setFormWidth(dims[0]/1440.);
	if (dims[1]>1440)
		m_state->m_pageSpan.setFormLength(dims[1]/1440.);
	val=int(libwps::readU16(input)); // f|226
	if (val!=0xf)
		f << "fl1=" << std::hex << val << std::dec << ",";
	val=int(libwps::readU16(input));
	if (val)
		f << "f1=" << val << ",";
	int margins[4]; // assume LTRB
	f << "margins=[";
	for (auto &m : margins)
	{
		m=int(libwps::readU16(input));
		f << double(m)/20. << ",";
		input->seek(2, librevenge::RVNG_SEEK_CUR);
	}
	f << "],";
	if (double(margins[0]+margins[2])<dims[0]/2)
	{
		m_state->m_pageSpan.setMarginLeft(double(margins[0])/1440);
		m_state->m_pageSpan.setMarginRight(double(margins[2])/1440);
	}
	else
	{
		WPS_DEBUG_MSG(("PocketWordParser::readPageDims: the LR margins seem bad\n"));
		f << "###";
	}
	if (double(margins[1]+margins[3])<dims[1]/2)
	{
		m_state->m_pageSpan.setMarginTop(double(margins[1])/1440);
		m_state->m_pageSpan.setMarginBottom(double(margins[3])/1440);
	}
	else
	{
		WPS_DEBUG_MSG(("PocketWordParser::readPageDims: the TB margins seem bad\n"));
		f << "###";
	}
	for (int i=0; i<3; ++i)   // 0
	{
		val=int(libwps::readU16(input));
		if (val)
			f << "f" << i+2 << "=" << val << ",";
	}
	for (int d=0; d<2; ++d)
	{
		f << "unkn" << d << "=[";
		for (int i=0; i<8; ++i)
		{
			val=int(libwps::readU16(input));
			int const expected = i==4 ? 0xa : i==7 ? 4 : 0;
			if (val!=expected)
				f << "f" << i << "=" << val << ",";
		}
		f << "],";
	}
	for (int i=0; i<7; ++i)   // g3=0|1
	{
		val=int(libwps::readU16(input));
		if (val)
			f << "g" << i << "=" << val << ",";
	}
	ascii().addPos(pos-6);
	ascii().addNote(f.str().c_str());
	return true;
}

bool PocketWordParser::sendParagraph(size_t paraId)
{
	auto input=getInput();
	if (!input)
		throw (libwps::ParseException());
	if (paraId>=m_state->m_typeEntryList.size() || m_state->m_typeEntryList[paraId].first!=65)
	{
		WPS_DEBUG_MSG(("PocketWordParser::sendParagraph: can not find paragraph %d\n", int(paraId)));
		return true;
	}
	WPSEntry const &entry=m_state->m_typeEntryList[paraId].second;
	entry.setParsed(true);
	long pos=entry.begin();
	input->seek(pos, librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f;
	f << "Entries(Paragraph):id=" << entry.id() << ",";
	if (entry.length()<22)
	{
		WPS_DEBUG_MSG(("PocketWordParser::sendParagraph: the length seems bad\n"));
		f << "###";
		ascii().addPos(pos-6);
		ascii().addNote(f.str().c_str());
		return true;
	}
	int textLen=int(libwps::readU16(input));
	f << "text[len]=" << textLen << ",";
	int textFormLen=int(libwps::readU16(input));
	f << "text[form,len]=" << textFormLen << ",";
	if (22+textFormLen>entry.length())
	{
		WPS_DEBUG_MSG(("PocketWordParser::sendParagraph: the form length seems bad\n"));
		f << "###";
		ascii().addPos(pos-6);
		ascii().addNote(f.str().c_str());
		return true;
	}
	int numLines=int(libwps::readU16(input));
	if (numLines!=1)
		f << "num[line]=" << numLines << ",";
	int val=int(libwps::read16(input));
	if (val==0)
		f << "*,";
	else if (val!=-1)
		f << "f0=" << val << ",";
	f << "id[paraDims]=" << libwps::readU16(input) << ","; // ? id
	WPSParagraph para;
	for (int i=0; i<4; ++i)
	{
		val=i==1 ? int(libwps::read16(input)) : int(libwps::readU16(input));
		if (!val) continue;
		char const *wh[]= {"f1", "indent[spec]", "left[marg]", "right[marg]"};
		if (i>=1 && i<=3)
			para.m_margins[i-1]=double(val)/72/20; // checkme: what is the unit, something between pt and dpi
		f << wh[i] << "=" << val << ",";
	}
	for (int i=0; i<2; ++i)
	{
		val=int(libwps::readU8(input));
		if (!val) continue;
		char const *wh[]= {"bullet", "align"};
		f << wh[i] << "=" << std::hex << val << std::dec << ",";
		if (i==0 && val==0xff)
		{
			para.m_listLevel.m_type=libwps::BULLET;
			libwps::appendUnicode(0x2022, para.m_listLevel.m_bullet);
			para.m_listLevelIndex=1;
			para.m_listLevel.m_labelIndent=para.m_margins[1];
			para.m_margins[1]=0;
		}
		else
		{
			switch (val)
			{
			case 0:
				break;
			case 1:
				para.m_justify=libwps::JustificationRight;
				break;
			case 2:
				para.m_justify=libwps::JustificationCenter;
				break;
			default:
				WPS_DEBUG_MSG(("PocketWordParser::sendParagraph: find unknown justification=%d\n", val));
				f << "###";
				break;
			}
		}
	}

	for (int i=0; i<2; ++i)
	{
		val=int(libwps::readU16(input));
		if (!val) continue;
		f << "f" << i+2 << "=" << val << ",";
	}
	m_listener->setParagraph(para);
	ascii().addPos(pos-6);
	ascii().addNote(f.str().c_str());

	pos=input->tell();
	long endPos=entry.end();
	f.str("");
	f << "Paragraph-Text:";
	// 83 04=? 9f05=? 9480=? 9980=' a305=t romanian fe 523 163 21b
	static std::map<uint8_t,int> const normalSpecialMap=
	{
		{0xc1,1}, {0xc2,2}, {0xc3,2}, {0xc4,1}, {0xc5,2},
		{0xe5,2}, {0xe6,2}, {0xe7,2}, {0xe8,2},
		{0xe9,1}, {0xea,1}, {0xeb,1}, {0xec,1},
		{0xef,3}
	};
	// the only tag created by the OpenOffice/LibreOffice convertor
	static std::map<uint8_t,int> const badSpecialMap=
	{
		{0xc4,1},
		{0xe5,2}, {0xe6,2}, {0xe7,2}, {0xe8,2},
		{0xe9,1}, {0xea,1}, {0xeb,1}, {0xec,1},
	};
	std::map<uint8_t,int> const &specialMap=m_state->m_badFile ? badSpecialMap : normalSpecialMap;
	libwps_tools_win::Font::Type fontType=m_state->m_encoding==libwps_tools_win::Font::UNKNOWN ?
	                                      libwps_tools_win::Font::WIN3_WEUROPE : m_state->m_encoding;
	m_listener->setFont(WPSFont());
	while (input->tell()<endPos)
	{
		auto c=libwps::readU8(input);
		auto it=specialMap.find(c);
		if (it!=specialMap.end() && input->tell()+it->second<=endPos)
		{
			bool done=true;
			long actPos=input->tell();
			if (it->second==4)
				val=int(libwps::readU32(input));
			else if (it->second==3)
				val=int(libwps::readU16(input))+int(libwps::readU8(input)<<16);
			else if (it->second==2)
				val=int(libwps::readU16(input));
			else
				val=int(libwps::readU8(input));
			if (c==0xc4 && val==0) // end of parsing
				break;
			switch (c)
			{
			case 0xc1:
				c=(unsigned char)(val);
				done=false;
				break;
			case 0xc2:
				f << "[ParaUnkn" << val << "]";
				break;
			// 0xc3: unsure, find one time an id to a paragraph and an id to a unexisting zone
			case 0xc4:
				if (val==4)
				{
					f << "\t";
					m_listener->insertTab();
				}
				else if (m_state->m_badFile && val>0x1f)
				{
					done=false;
					input->seek(actPos, librevenge::RVNG_SEEK_SET);
				}
				else
					f << "[##" << std::hex << int(c) << std::dec << "=" << val << "]";
				break;
			case 0xc5:
			{
				int cType;
				WPSEntry cEntry=m_state->getEntry(val, cType);
				f << "[Obj" << val << "]";
				if (!cEntry.valid())
				{
					f << "###";
					break;
				}
				if (cType!=84)
				{
					WPS_DEBUG_MSG(("PocketWordParser::sendParagraph: object %d does not correspond to a 84 zone\n", val));
					f << "###";
					break;
				}
				WPSEmbeddedObject object;
				actPos=input->tell();
				if (readSound(cEntry, object) && !object.isEmpty())
				{
					WPSPosition objPos(Vec2f(0,0), Vec2f(72,72), librevenge::RVNG_POINT);
					objPos.setRelativePosition(WPSPosition::Char);
					m_listener->insertObject(objPos, object);
				}
				input->seek(actPos, librevenge::RVNG_SEEK_SET);
				break;
			}
			case 0xe5:
			case 0xe6:
			case 0xe7:
			case 0xe8:
			case 0xe9:
			case 0xea:
			case 0xeb:
			case 0xec:
			{
				WPSFont font=m_listener->getFont();
				if (c==0xe5)
				{
					f << "[FN" << val << "]";
					auto fIt=m_state->m_idToFontNameMap.find(val);
					if (fIt!=m_state->m_idToFontNameMap.end())
						font.m_name=fIt->second;
					else if (val==0) // checkme: maybe system
						font.m_name="courier";
					else
					{
						WPS_DEBUG_MSG(("PocketWordParser::sendParagraph: can not find font %d\n", val));
						f << "###";
					}
				}
				else if (c==0xe6)
				{
					f << "[FS=" << val << "]";
					font.m_size=double(val ? val : 12); // checkme: what is default size
				}
				else if (c==0xe7)
				{
					if (m_state->getColor(val, font.m_color))
						f << "[FC=" << font.m_color << "]";
					else
					{
						WPS_DEBUG_MSG(("PocketWordParser::sendParagraph: unknown color %d\n", val));
						f << "[FC=" << val << "]###";
					}
				}
				else if (c==0xe8)
				{
					f << "[Fw=" << val << "]";
					if (val==4 || val==1) // normal, fine
						font.m_attributes&=~(unsigned int)(WPS_BOLD_BIT);
					else if (val==7 || val==8) // bold, thick
						font.m_attributes|=WPS_BOLD_BIT;
					else
					{
						WPS_DEBUG_MSG(("PocketWordParser::sendParagraph: unknown font weigth %d\n", val));
						f << "###";
					}
				}
				else if (c==0xe9)
				{
					f << "[FIt=" << val << "]";
					if (val==0)
						font.m_attributes&=~(unsigned int)(WPS_ITALICS_BIT);
					else if (val==1)
						font.m_attributes|=WPS_ITALICS_BIT;
					else
					{
						WPS_DEBUG_MSG(("PocketWordParser::sendParagraph: unknown italic flag %d\n", val));
						f << "###";
					}
				}
				else if (c==0xea)
				{
					f << "[FUnd=" << val << "]";
					if (val==0)
						font.m_attributes&=~(unsigned int)(WPS_UNDERLINE_BIT);
					else if (val==1)
						font.m_attributes|=WPS_UNDERLINE_BIT;
					else
					{
						WPS_DEBUG_MSG(("PocketWordParser::sendParagraph: unknown underline flag %d\n", val));
						f << "###";
					}
				}
				else if (c==0xeb)
				{
					f << "[FStr=" << val << "]";
					if (val==0)
						font.m_attributes&=~(unsigned int)(WPS_STRIKEOUT_BIT);
					else if (val==1)
						font.m_attributes|=WPS_STRIKEOUT_BIT;
					else
					{
						WPS_DEBUG_MSG(("PocketWordParser::sendParagraph: unknown strike flag %d\n", val));
						f << "###";
					}
				}
				else
				{
					f << "[FHil=" << val << "]";
					if (val==0)
						font.m_attributes&=~(unsigned int)(WPS_REVERSEVIDEO_BIT);
					else if (val==1)
						font.m_attributes|=WPS_REVERSEVIDEO_BIT;
					else
					{
						WPS_DEBUG_MSG(("PocketWordParser::sendParagraph: unknown hilite flag %d\n", val));
						f << "###";
					}
				}
				m_listener->setFont(font);
				break;
			}
			default:
				f << "[C###" << std::hex << int(c) << std::dec << "=" << val << "]";
				break;
			}
			if (done)
			{
				if (val>100 && it->second<3) f << "##";
				continue;
			}

		}
		if (c<0x1f)
			f << "[###" << std::hex << int(c) << std::dec << "]";
		else
		{
			f << char(c);
			m_listener->insertUnicode(uint32_t(libwps_tools_win::Font::unicode((unsigned char)(c), fontType)));
		}
	}
	m_listener->insertEOL();
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	return true;
}

bool PocketWordParser::readParagraphDims(WPSEntry const &entry)
{
	auto input=getInput();
	if (!input)
		throw (libwps::ParseException());
	long pos=entry.begin();
	input->seek(pos, librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f;
	f << "Entries(ParaDims):id=" << entry.id() << ",";

	if ((entry.length()%2)!=0 || (entry.length()%10)>2)
	{
		WPS_DEBUG_MSG(("PocketWordParser::readParagraphDims: the form length seems bad\n"));
		f << "###";
		ascii().addPos(pos-6);
		ascii().addNote(f.str().c_str());
		return true;
	}

	ascii().addPos(pos-6);
	ascii().addNote(f.str().c_str());

	int N=int(entry.length()/10);
	for (int i=0; i<N; ++i)
	{
		pos=input->tell();
		f.str("");
		f << "ParaDims-L" << i << ":";
		f << "num[char]=" << libwps::readU16(input) << ",";
		int val=int(libwps::readU32(input));
		if (val)
			f << "fill?=" << val << ",";
		f << "w=" << libwps::readU16(input) << ","; // in point/2 ?
		f << "h=" << int(libwps::readU8(input)) << ",";
		// 2-6 some enum corresponding to h? 2:h~=10, 3:h~=12, 4:h~=14
		f << "fl=" << int(libwps::readU8(input)) << ",";
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		input->seek(pos+10, librevenge::RVNG_SEEK_SET);
	}
	return true;
}

bool PocketWordParser::readParagraphUnkn(WPSEntry const &entry)
{
	// maybe some tabs
	auto input=getInput();
	if (!input)
		throw (libwps::ParseException());
	long pos=entry.begin();
	input->seek(pos, librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f;
	f << "Entries(ParaUnkn):id=" << entry.id() << ",";

	if (entry.length()<11)   // 20 or 24
	{
		WPS_DEBUG_MSG(("PocketWordParser::readParagraphUnkn: the form length seems bad\n"));
		f << "###";
		ascii().addPos(pos-6);
		ascii().addNote(f.str().c_str());
		return true;
	}
	int val=int(libwps::readU16(input)); // f0 with sz=20|f2 with sz=24
	f << "fl=" << std::hex << val << std::dec << ",";
	val=int(libwps::readU16(input)); // 40
	if (val!=0x40)
		f << "f0=" << val << ",";
	f << "id2=" << libwps::readU16(input) << ","; // 1-3
	val=int(libwps::readU16(input)); // 1
	if (val!=1)
		f << "f1=" << val << ",";
	int sz2=libwps::readU16(input);
	if ((sz2%3)==0 && input->tell()+sz2<=entry.end())
	{
		// really unsure
		int N=sz2/3;
		for (int i=0; i<N; ++i)
		{
			f << "unkn" << i << "=[";
			for (int j=0; j<3; ++j)
			{
				val=int(libwps::readU8(input));
				if (val)
					f << val << ",";
				else
					f << "_,";
			}
			f << "],";
		}
	}
	else
	{
		WPS_DEBUG_MSG(("PocketWordParser::readParagraphUnkn: something seems bads\n"));
		f << "##sz[data]=" << sz2 << ","; // 6,9,c
	}
	ascii().addPos(pos-6);
	ascii().addNote(f.str().c_str());
	return true;
}

bool PocketWordParser::readParagraphList(WPSEntry const &entry, std::vector<int> &paraId)
{
	auto input=getInput();
	if (!input)
		throw (libwps::ParseException());
	entry.setParsed(true);
	long pos=entry.begin();
	input->seek(pos, librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f;
	f << "Entries(ParaList):id=" << entry.id() << ",";

	if ((entry.length()%8)!=0 || entry.length()<24)
	{
		WPS_DEBUG_MSG(("PocketWordParser::readParagraphList: the form length seems bad\n"));
		f << "###";
		ascii().addPos(pos-6);
		ascii().addNote(f.str().c_str());
		return true;
	}
	f << "vals=[";
	for (int i=0; i<6; ++i)   // some big number often [X,Y,Z,W,0,0] with X~Y~W and Z much bigger
	{
		f << libwps::readU32(input) << ",";
	}
	f << "],";
	ascii().addPos(pos-6);
	ascii().addNote(f.str().c_str());

	int N=int(entry.length()/8)-3;
	for (int i=0; i<N; ++i)
	{
		pos=input->tell();
		f.str("");
		f << "ParaList-L" << i << ":";
		int val=int(libwps::readU16(input));
		if (val!=1)
			f << "num[lines]=" << val << ",";
		f << "num[char]=" << int(libwps::readU16(input)) << ",";
		paraId.push_back(int(libwps::readU16(input)));
		f << "id=" << paraId.back() << ",";
		val=int(libwps::readU16(input)); // always 0?
		if (val)
			f << "f0=" << val << ",";
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		input->seek(pos+8, librevenge::RVNG_SEEK_SET);
	}
	return true;
}

bool PocketWordParser::readSound(WPSEntry const &entry, WPSEmbeddedObject &object)
{
	// unsure, may also be a picture
	auto input=getInput();
	if (!input)
		throw (libwps::ParseException());
	entry.setParsed(true);
	long pos=entry.begin();
	input->seek(pos, librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f;
	f << "Entries(Sound):id=" << entry.id() << ",";
	// unsure find many times entry.length()==430, but one time entry.length()==1274...
	// I suspect that testing if entry.length()<0x32e must be ok
	if (entry.length()<0x430)
	{
		WPS_DEBUG_MSG(("PocketWordParser::readSound: the length seems bad\n"));
		f << "###";
		ascii().addPos(pos-6);
		ascii().addNote(f.str().c_str());
		return true;
	}
	for (int i=0; i<4; ++i)
	{
		int val=int(libwps::readU16(input));
		int const expected[]= {0,1,0x49,0};
		if (val!=expected[i])
			f << "f" << i << "=" << val << ",";
	}
	long pictSize=libwps::readU32(input);
	f << "sz=" << pictSize << ",";
	int val=int(libwps::readU16(input));
	if (val) // 0 or big number
		f << "unkn=" << std::hex << val << std::dec << ",";
	val=int(libwps::readU16(input));
	if (val) // 5-7 or 4450
		f << "f4=" << std::hex << val << std::dec << ",";
	for (int i=0; i<2; ++i)   // 0
	{
		val=int(libwps::readU16(input));
		if (val)
			f << "f" << i+5 << "=" << val << ",";
	}
	f << "checksum?=" << std::hex << libwps::readU32(input) << std::dec << ",";
	for (int i=0; i<50; ++i)   // 0
	{
		val=int(libwps::readU16(input));
		if (val)
			f << "g" << i << "=" << val << ",";
	}
	ascii().addPos(pos-6);
	ascii().addNote(f.str().c_str());

	for (int st=0; st<2; ++st)
	{
		pos=input->tell();
		f.str("");
		f << "Sound-" << st << ":";
		for (int i=0; i<107; ++i)   // 0
		{
			val=int(libwps::readU16(input));
			if (val)
				f << "f" << i << "=" << val << ",";
		}
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
	}

	pos=input->tell();
	f.str("");
	f << "Sound-name:";
	librevenge::RVNGString name; // always VoiceNote.wav
	for (int j=0; j<128; ++j)
	{
		val=int(libwps::readU16(input));
		if (!val)
			break;
		libwps::appendUnicode(uint32_t(val), name);
	}
	f << name.cstr();
	input->seek(pos+256, librevenge::RVNG_SEEK_SET);
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	pos=input->tell();
	f.str("");
	f << "Sound-2:";
	for (int i=0; i<132; ++i)   // 0
	{
		val=int(libwps::readU16(input));
		if (val)
			f << "f" << i << "=" << val << ",";
	}
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	// try to read the wav file
	int cType;
	WPSEntry cEntry=m_state->getEntry(int(entry.id()+65536), cType);
	if (!cEntry.valid() || cType!=85)
	{
		WPS_DEBUG_MSG(("PocketWordParser::readSound: can not find data corresponding to %d\n", entry.id()));
		return true;
	}
	readSoundData(cEntry, pictSize, object);
	input->seek(entry.end(), librevenge::RVNG_SEEK_SET);
	return true;
}

bool PocketWordParser::readSoundData(WPSEntry const &entry, long pictSize, WPSEmbeddedObject &object)
{
	// unsure, may also be a picture
	auto input=getInput();
	if (!input)
		throw (libwps::ParseException());
	entry.setParsed();
	long pos=entry.begin();
	input->seek(pos, librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f;
	f << "Entries(SndData):id=" << entry.id() << ",";
	if (entry.length()<20 || entry.length()<pictSize ||	entry.length()>pictSize+20)
	{
		WPS_DEBUG_MSG(("PocketWordParser::readSoundData: the length seems bad\n"));
		f << "###";
		ascii().addPos(pos-6);
		ascii().addNote(f.str().c_str());
		return true;
	}
	ascii().addPos(pos-6);
	ascii().addNote(f.str().c_str());

	librevenge::RVNGBinaryData data;
	if (!libwps::readData(input,(unsigned long)pictSize,data))
	{
		WPS_DEBUG_MSG(("PocketWordParser::readSoundData: can not read the sound\n"));
		f << "###";
		ascii().addPos(pos);
		ascii().addNote("SndData:###");
		return true;
	}
	object=WPSEmbeddedObject(data); // checkme: probably better to set media-type="application/vnd.sun.star.media"

	static int pictId=0;
	std::stringstream s;
	s << "Snd" << ++pictId << ".wav";
	libwps::Debug::dumpFile(data, s.str().c_str());
	ascii().skipZone(pos, entry.end()-1);
	return true;
}

bool PocketWordParser::readUnkn8(WPSEntry const &entry)
{
	// checkme can be pref
	auto input=getInput();
	if (!input)
		throw (libwps::ParseException());
	long pos=entry.begin();
	input->seek(pos, librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f;
	f << "Entries(UnknC):id=" << entry.id() << ",";
	if (entry.length()!=64)
	{
		WPS_DEBUG_MSG(("PocketWordParser::readUnkn8: the length seems bad\n"));
		f << "###";
		ascii().addPos(pos-6);
		ascii().addNote(f.str().c_str());
		return true;
	}
	int val=int(libwps::readU16(input));
	if (val!=1)
		f << "f0=" << val << ",";
	for (int i=0; i<9; ++i)   // f2=0|2, f6=0|10e|170|b, f8=0|1f62|237d
	{
		val=int(libwps::readU16(input));
		if (val)
			f << "f" << i+1 << "=" << std::hex << val << std::dec << ",";
	}
	val=int(libwps::readU16(input)); // 8|20
	if (val!=20)
		f << "f10=" << val << ",";
	for (int i=0; i<17; ++i)   // g3~g11,g5=g13 some dimension/positions in points?
	{
		val=int(libwps::readU16(input));
		if (val)
			f << "g" << i << "=" << std::hex << val << std::dec << ",";
	}
	// then 0000000000000000 or 09040000e4040100 or 09040904e4040101 some item id?
	ascii().addPos(pos-6);
	ascii().addNote(f.str().c_str());
	return true;
}
bool PocketWordParser::readUnkn2021(WPSEntry const &entry, int type)
{
	// checkme follow the font name, so this can be the styles
	auto input=getInput();
	long pos=entry.begin();
	input->seek(pos, librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f;
	f << "Entries(Unkn" << (type==20 ? "A": "B") << "):id=" << entry.id() << ",";
	if (entry.length() < 4)
	{
		WPS_DEBUG_MSG(("PocketWordParser::readUnkn2021: the length seems bad\n"));
		f << "###";
		ascii().addPos(pos-6);
		ascii().addNote(f.str().c_str());
		return true;
	}
	for (int i=0; i<2; ++i)   // a number of data + ???
	{
		int val=int(libwps::readU16(input));
		if (!val)
			continue;
		f << "f" << i << "=" << val << ",";
	}
	if (entry.length() > 4)
	{
		WPS_DEBUG_MSG(("PocketWordParser::readUnkn2021: find unexpected data in zone %d\n", type));
		f << "###";
	}
	ascii().addPos(pos-6);
	ascii().addNote(f.str().c_str());
	return true;
}

////////////////////////////////////////////////////////////
// low level
////////////////////////////////////////////////////////////


// basic function to check if the header is ok
bool PocketWordParser::checkHeader(WPSHeader *header, bool strict)
{
	RVNGInputStreamPtr input = getInput();
	if (!input || !checkFilePosition(0x74)) // size of the header and a font
	{
		WPS_DEBUG_MSG(("PocketWordParser::checkHeader: file is too short\n"));
		return false;
	}

	input->seek(0x0, librevenge::RVNG_SEEK_SET);
	if (libwps::readU32(input)!=0x77705c7b || libwps::readU32(input)!=0x1569 || libwps::readU16(input)!=0x101)
		return false;
	libwps::DebugStream f;
	f << "FileHeader:";
	int vers=int(libwps::readU16(input));
	if (vers<6 || vers>7)
	{
		WPS_DEBUG_MSG(("PocketWordParser::checkHeader: unknown version %d\n", vers));
		return false;
	}
	f << "v" << vers << ",";
	int val=int(libwps::readU16(input));
	if (val)
		f << "f0=" << val << ",";
	if (strict)
	{
		long pos=input->tell();
		// check the first length
		input->seek(4, librevenge::RVNG_SEEK_CUR); // type, id
		val=int(libwps::readU16(input));
		if (!checkFilePosition(input->tell()+4*val))
		{
			WPS_DEBUG_MSG(("PocketWordParser::checkHeader: can read the font name length\n"));
			return false;
		}
		input->seek(pos, librevenge::RVNG_SEEK_SET);
	}
	ascii().addPos(0);
	ascii().addNote(f.str().c_str());

	m_state->m_version=vers;
	if (header)
	{
		header->setMajorVersion(vers);
		header->setNeedEncoding(true);
	}
	return true;
}

////////////////////////////////////////////////////////////
// send the different data
////////////////////////////////////////////////////////////
void PocketWordParser::sendData()
{
	auto input=getInput();
	if (!input || !m_listener)
		throw (libwps::ParseException());

	// first, retrieves the font names
	auto it=m_state->m_typeToEntryMap.find(0);
	if (it==m_state->m_typeToEntryMap.end())
	{
		WPS_DEBUG_MSG(("PocketWordParser::sendData: can not find the fonts names\n"));
	}
	else
	{
		for (auto const &id : it->second)
		{
			if (id >= m_state->m_typeEntryList.size() || !m_state->m_typeEntryList[id].second.valid())
			{
				WPS_DEBUG_MSG(("PocketWordParser::sendData: oops pb when reading the fonts names\n"));
				continue;
			}
			readFontNames(m_state->m_typeEntryList[id].second);
		}
	}

	// now, retrieve the list of zone, we need to send
	if (!m_state->m_badFile)
	{
		auto lIt=m_state->m_typeToEntryMap.find(64);
		if (lIt==m_state->m_typeToEntryMap.end())
		{
			WPS_DEBUG_MSG(("PocketWordParser::sendData: can not find any paragraph list\n"));
		}
		else
		{
			for (auto const &id : lIt->second)
			{
				if (id >= m_state->m_typeEntryList.size() || !m_state->m_typeEntryList[id].second.valid())
				{
					WPS_DEBUG_MSG(("PocketWordParser::sendData: oops pb when reading some paragraph list\n"));
					continue;
				}
				std::vector<int> paraId;
				readParagraphList(m_state->m_typeEntryList[id].second, paraId);
				for (auto &pId : paraId)
				{
					if (pId==0)
						continue;
					auto pIt=m_state->m_idToEntryMap.find(pId);
					if (pIt==m_state->m_idToEntryMap.end())
					{
						WPS_DEBUG_MSG(("PocketWordParser::sendData: can not find paragraph %d\n", pId));
						continue;
					}
					sendParagraph(pIt->second);
				}
			}
		}
		return;
	}
	auto pIt=m_state->m_typeToEntryMap.find(65);
	if (pIt==m_state->m_typeToEntryMap.end())
	{
		WPS_DEBUG_MSG(("PocketWordParser::sendData: can not find any paragraph\n"));
		return;
	}
	for (auto const &id : pIt->second)
		sendParagraph(id);
}

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
