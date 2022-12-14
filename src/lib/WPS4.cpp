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

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WPSContentListener.h"
#include "WPSEntry.h"
#include "WPSHeader.h"
#include "WPSOLEParser.h"
#include "WPSPageSpan.h"
#include "WPSTextSubDocument.h"

#include "WPS4Graph.h"
#include "WPS4Text.h"

#include "WPS4.h"

namespace WPS4ParserInternal
{
//! Internal: the subdocument of a WPS4Parser
class SubDocument final : public WPSTextSubDocument
{
public:
	//! type of an entry stored in textId
	enum Type { Unknown, MN };
	//! constructor for a text entry
	SubDocument(RVNGInputStreamPtr const &input, WPS4Parser &pars, WPSEntry const &entry)
		: WPSTextSubDocument(input, &pars), m_entry(entry) {}
	//! destructor
	~SubDocument() final {}

	//! operator==
	bool operator==(std::shared_ptr<WPSSubDocument> const &doc) const final
	{
		if (!doc || !WPSTextSubDocument::operator==(doc))
			return false;
		auto const *sDoc = dynamic_cast<SubDocument const *>(doc.get());
		if (!sDoc) return false;
		return m_entry == sDoc->m_entry;
	}

	//! the parser function
	void parse(std::shared_ptr<WPSContentListener> &listener, libwps::SubDocumentType subDocumentType) final;
	//! the entry
	WPSEntry m_entry;
};

void SubDocument::parse(std::shared_ptr<WPSContentListener> &listener, libwps::SubDocumentType subDocumentType)
{
	if (!listener.get())
	{
		WPS_DEBUG_MSG(("WPS4ParserInternal::SubDocument::parse: no listener\n"));
		return;
	}
	if (!dynamic_cast<WPSContentListener *>(listener.get()))
	{
		WPS_DEBUG_MSG(("WPS4ParserInternal::SubDocument::parse: bad listener\n"));
		return;
	}

	auto &listen =  reinterpret_cast<WPSContentListenerPtr &>(listener);
	if (!m_parser)
	{
		listen->insertCharacter(' ');
		WPS_DEBUG_MSG(("WPS4ParserInternal::SubDocument::parse: bad parser\n"));
		return;
	}

	if (m_entry.isParsed() && subDocumentType != libwps::DOC_HEADER_FOOTER)
	{
		listen->insertCharacter(' ');
		WPS_DEBUG_MSG(("WPS4ParserInternal::SubDocument::parse: this zone is already parsed\n"));
		return;
	}
	m_entry.setParsed(true);
	if (m_entry.type() != "TEXT")
	{
		listen->insertCharacter(' ');
		WPS_DEBUG_MSG(("WPS4ParserInternal::SubDocument::parse: send not Text entry is not implemented\n"));
		return;
	}

	if (!m_entry.valid())
	{
		if (subDocumentType != libwps::DOC_COMMENT_ANNOTATION)
		{
			WPS_DEBUG_MSG(("SubDocument::parse: empty document found...\n"));
		}
		listen->insertCharacter(' ');
		return;
	}

	auto *mnParser = dynamic_cast<WPS4Parser *>(m_parser);
	if (!mnParser)
	{
		WPS_DEBUG_MSG(("SubDocument::parse: bad parser...\n"));
		listen->insertCharacter(' ');
		return;
	}
	mnParser->send(m_entry, subDocumentType);
}

//! the state of WPS4
struct State
{
	explicit State(libwps_tools_win::Font::Type fontType)
		: m_isDosFile(false)
		, m_eof(-1)
		, m_fontType(fontType)
		, m_OEMFontType(libwps_tools_win::Font::UNKNOWN)
		, m_pageSpan()
		, m_noFirstPageHeader(false)
		, m_noFirstPageFooter(false)
		, m_numColumns(1)
		, m_actPage(0)
		, m_numPages(0)
	{
	}
	//! flag to know if the file is or not a dos file
	bool m_isDosFile;
	//! the last file position
	long m_eof;
	//! the user font type
	libwps_tools_win::Font::Type m_fontType;
	//! the OEM code page
	libwps_tools_win::Font::Type m_OEMFontType;
	//! the actual document size
	WPSPageSpan m_pageSpan;
	bool m_noFirstPageHeader /* true if the first page has no header */;
	bool m_noFirstPageFooter /* true if the first page has no footer */;
	int m_numColumns /** the number of columns */;
	int m_actPage /** the actual page*/, m_numPages /* the number of pages */;
};
}

// constructor, destructor
WPS4Parser::WPS4Parser(RVNGInputStreamPtr &input, WPSHeaderPtr &header,
                       libwps_tools_win::Font::Type encoding)
	: WPSParser(input, header)
	, m_listener()
	, m_graphParser()
	, m_textParser()
	, m_state()
{
	m_state.reset(new WPS4ParserInternal::State(encoding));
	m_graphParser.reset(new WPS4Graph(*this));
	m_textParser.reset(new WPS4Text(*this, input));
}

WPS4Parser::~WPS4Parser()
{
}

// small funtion ( dimension, color, fileSize, ...)
float WPS4Parser::pageHeight() const
{
	return float(m_state->m_pageSpan.getFormLength()-m_state->m_pageSpan.getMarginTop()-m_state->m_pageSpan.getMarginBottom());
}

float WPS4Parser::pageWidth() const
{
	return float(m_state->m_pageSpan.getFormWidth()-m_state->m_pageSpan.getMarginLeft()-m_state->m_pageSpan.getMarginRight());
}

int WPS4Parser::numColumns() const
{
	return m_state->m_numColumns;
}

libwps_tools_win::Font::Type WPS4Parser::getDefaultFontType() const
{
	if (m_state->m_fontType != libwps_tools_win::Font::UNKNOWN)
		return m_state->m_fontType;
	return version()<=2 ? libwps_tools_win::Font::DOS_850 : libwps_tools_win::Font::WIN3_WEUROPE;
}

libwps_tools_win::Font::Type WPS4Parser::getOEMFontType() const
{
	return m_state->m_OEMFontType;
}

bool WPS4Parser::getColor(int id, WPSColor &color) const
{
	if (m_state->m_isDosFile)
	{
		static const uint32_t colorDosMap[]=
		{
			0x0, 0xFF0000, 0x00FF00, 0x0000FF,
			0x00FFFF, 0xFF00FF, 0xFFFF00
		};
		if (id < 0 || id >= 7)
		{
			WPS_DEBUG_MSG(("WPS4Parser::getColor(): unknown Dos color id: %d\n",id));
			return false;
		}
		color=WPSColor(colorDosMap[id]);
		return true;
	}
	static const uint32_t colorMap[]=
	{
		// 0x00RRGGBB
		0, // auto
		0,
		0x0000FF, 0x00FFFF,
		0x00FF00, 0xFF00FF,	0xFF0000, 0xFFFF00,
		0x808080, 0xFFFFFF,	0x000080, 0x008080,
		0x008000, 0x800080,	0x800000, 0xC0C0C0
	};
	if (id < 0 || id >= 16)
	{
		WPS_DEBUG_MSG(("WPS4Parser::getColor(): unknown color id: %d\n",id));
		return false;
	}
	color = WPSColor(colorMap[id]);
	return true;
}

void WPS4Parser::setSizeFile(long sz)
{
	if (sz > m_state->m_eof)
		m_state->m_eof = sz;
}

bool WPS4Parser::checkFilePosition(long pos)
{
	if (m_state->m_eof < 0)
	{
		RVNGInputStreamPtr input = getInput();
		long actPos = input->tell();
		input->seek(0, librevenge::RVNG_SEEK_END);
		m_state->m_eof=input->tell();
		input->seek(actPos, librevenge::RVNG_SEEK_SET);
	}
	return pos <= m_state->m_eof;
}

// listener, new page
void WPS4Parser::setListener(std::shared_ptr<WPSContentListener> const &listener)
{
	m_listener = listener;
	m_graphParser->setListener(m_listener);
	m_textParser->setListener(m_listener);
}

std::shared_ptr<WPSContentListener> WPS4Parser::createListener(librevenge::RVNGTextInterface *interface)
{
	std::vector<WPSPageSpan> pageList;
	WPSPageSpan page1(m_state->m_pageSpan), ps(m_state->m_pageSpan);

	WPSEntry ent = m_textParser->getHeaderEntry();
	if (ent.valid())
	{
		WPSSubDocumentPtr subdoc(new WPS4ParserInternal::SubDocument
		                         (getInput(), *this, ent));
		ps.setHeaderFooter(WPSPageSpan::HEADER, WPSPageSpan::ALL, subdoc);
		if (!m_state->m_noFirstPageHeader)
			page1.setHeaderFooter(WPSPageSpan::HEADER, WPSPageSpan::ALL, subdoc);
	}
	ent = m_textParser->getFooterEntry();
	if (ent.valid())
	{
		WPSSubDocumentPtr subdoc(new WPS4ParserInternal::SubDocument
		                         (getInput(), *this, ent));
		ps.setHeaderFooter(WPSPageSpan::FOOTER, WPSPageSpan::ALL, subdoc);
		if (!m_state->m_noFirstPageFooter)
			page1.setHeaderFooter(WPSPageSpan::FOOTER, WPSPageSpan::ALL, subdoc);
	}

	int numPages = m_textParser->numPages();
	int graphPages = m_graphParser->numPages();
	if (graphPages>=numPages) numPages = graphPages;
#ifdef DEBUG
	// create all the pages + an empty page, if we have some remaining data...
	numPages++;
#endif
	pageList.push_back(page1);
	for (int i = 1; i < numPages; i++) pageList.push_back(ps);
	m_state->m_numPages=numPages;
	return std::shared_ptr<WPSContentListener>
	       (new WPSContentListener(pageList, interface));
}

void WPS4Parser::newPage(int number)
{
	if (number <= m_state->m_actPage || number > m_state->m_numPages)
		return;

	while (m_state->m_actPage < number)
	{
		m_state->m_actPage++;

		if (m_listener.get() == nullptr || m_state->m_actPage == 1)
			continue;
		m_listener->insertBreak(WPS_PAGE_BREAK);
		m_graphParser->sendObjects(m_state->m_actPage);
	}
}

////////////////////////////////////////////////////////////
// interface with the graph/text parser
////////////////////////////////////////////////////////////
int WPS4Parser::readObject(RVNGInputStreamPtr input, WPSEntry const &entry)
{
	long actPos = input->tell();
	int id=m_graphParser->readObject(input, entry);
	input->seek(actPos, librevenge::RVNG_SEEK_SET);
	return id;
}

void WPS4Parser::sendObject(WPSPosition const &position, int id)
{
	return m_graphParser->sendObject(position, id);
}

void WPS4Parser::send(WPSEntry const &entry, libwps::SubDocumentType)
{
	if (!entry.hasType("TEXT"))
	{
		WPS_DEBUG_MSG(("WPS4Parser::send: unknown entry type '%s'\n", entry.type().c_str()));
		if (m_listener.get()) m_listener->insertCharacter(' ');
		return;
	}
	RVNGInputStreamPtr input = getInput();
	long actPos = input->tell();
	m_textParser->readText(entry);
	input->seek(actPos, librevenge::RVNG_SEEK_SET);
}

void WPS4Parser::createDocument(WPSEntry const &entry, libwps::SubDocumentType type)
{
	if (m_listener.get() == nullptr) return;
	WPSSubDocumentPtr subdoc(new WPS4ParserInternal::SubDocument
	                         (getInput(), *this, entry));
	if (type == libwps::DOC_COMMENT_ANNOTATION)
		m_listener->insertComment(subdoc);
	else
	{
		WPS_DEBUG_MSG(("WPS4Parser:createDocument error: unknown type: \"%d\"\n", int(type)));
	}
}

void WPS4Parser::createNote(WPSEntry const &entry, librevenge::RVNGString const &label)
{
	if (m_listener.get() == nullptr) return;
	WPSSubDocumentPtr subdoc(new WPS4ParserInternal::SubDocument
	                         (getInput(), *this, entry));
	m_listener->insertLabelNote(WPSContentListener::FOOTNOTE, label, subdoc);
}

void WPS4Parser::createTextBox(WPSEntry const &entry, WPSPosition const &pos, librevenge::RVNGPropertyList &extras)
{
	if (m_listener.get() == nullptr) return;
	WPSSubDocumentPtr subdoc(new WPS4ParserInternal::SubDocument
	                         (getInput(), *this, entry));
	m_listener->insertTextBox(pos, subdoc, extras);
}

////////////////////////////////////////////////////////////
// main funtions to parse a document, its OLE structures and its main input
////////////////////////////////////////////////////////////

// main function to parse the document
void WPS4Parser::parse(librevenge::RVNGTextInterface *documentInterface)
{
	RVNGInputStreamPtr input=getInput();
	if (!input)
	{
		WPS_DEBUG_MSG(("WPS4Parser::parse: does not find main ole\n"));
		throw (libwps::ParseException());
	}
	if (!checkHeader(nullptr, true))
		throw (libwps::ParseException());
	try
	{
		createOLEStructures();
	}
	catch (...)
	{
		WPS_DEBUG_MSG(("WPS4Parser::parse: exception catched when parsing secondary OLEs\n"));
	}

	ascii().setStream(input);
	ascii().open("MN0");
	try
	{
		createStructures();
	}
	catch (...)
	{
		WPS_DEBUG_MSG(("WPS4Parser::parse: exception catched when parsing MN0\n"));
		throw (libwps::ParseException());
	}

	setListener(createListener(documentInterface));
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("WPS4Parser::parse: can not create the listener\n"));
		throw (libwps::ParseException());
	}
	m_listener->startDocument();
	m_textParser->sendObjects(-1);
	WPSEntry ent = m_textParser->getMainTextEntry();
	if (ent.valid())
		m_textParser->readText(ent);
	else
	{
		WPS_DEBUG_MSG(("WPS4Parser::parse: can not find main text entry\n"));
		throw (libwps::ParseException());
	}
#ifdef DEBUG
	m_textParser->flushExtra();
	m_graphParser->sendObjects(-1);
#endif
	m_listener->endDocument();
	m_listener.reset();

	ascii().reset();
}

// find and create all the zones ( normal/ole )
bool WPS4Parser::createStructures()
{
	if (!findZones()) throw libwps::ParseException();
	if (!m_textParser->readStructures()) throw libwps::ParseException();
	m_graphParser->computePositions();

#ifdef DEBUG
	WPS4Parser::NameMultiMap::iterator pos;
	pos = getNameEntryMap().find("PRNT");
	if (pos != getNameEntryMap().end()) readPrnt(pos->second);
	pos = getNameEntryMap().find("DocWInfo");
	if (pos != getNameEntryMap().end()) readDocWindowsInfo(pos->second);
#endif

	return true;
}

bool WPS4Parser::createOLEStructures()
{
	RVNGInputStreamPtr input=getFileInput();
	if (!input) return false;

	if (!input->isStructured()) return true;
	WPSOLEParser oleParser("MN0", getDefaultFontType());
	if (!oleParser.parse(input)) return false;

	m_graphParser->storeObjects(oleParser.getObjectsMap());
#ifdef DEBUG
	// there can remain some embedded Works subdocument ( WKS, ... )
	// with name MN0 and some unknown picture ole
	auto unparsed = oleParser.getNotParse();

	for (auto const &name : unparsed)
	{
		if (name == "MN0")
			continue;
		WPS_DEBUG_MSG(("WPS4Parser::createOLEStructures:: Find unparsed ole: %s\n", name.c_str()));

#ifdef DEBUG_WITH_FILES
		RVNGInputStreamPtr ole(input->getSubStreamByName(name.c_str()));
		if (!ole.get())
		{
			WPS_DEBUG_MSG(("WPS4Parser::createOLEStructures: error: can find OLE part: \"%s\"\n", name.c_str()));
			continue;
		}

		librevenge::RVNGBinaryData data;
		if (libwps::readDataToEnd(ole,data))
			libwps::Debug::dumpFile(data, name.c_str());
#endif
	}
#endif
	return true;
}

////////////////////////////////////////////////////////////
// low level
////////////////////////////////////////////////////////////

// ------------------------------------------------------------
// function to read/check the file header or to check that an entry is ok
// ------------------------------------------------------------

// parse a basic entry: ie offset + size and check if it is valid
bool WPS4Parser::parseEntry(std::string const &name)
{
	RVNGInputStreamPtr input = getInput();
	long actPos = input->tell();
	WPSEntry zone;
	zone.setBegin(long(libwps::readU32(input)));
	zone.setLength(libwps::readU16(input));
	zone.setType(name);

	bool ok = zone.valid() && checkFilePosition(zone.end());

	if (ok)
	{
		getNameEntryMap().insert(NameMultiMap::value_type(zone.type(), zone));

		ascii().addPos(zone.begin());
		std::string nm = "ZZ";
		nm+= name;
		ascii().addNote(nm.c_str());
		ascii().addPos(zone.end());
		ascii().addNote("_");
	}

	libwps::DebugStream f;
	if (ok)
		f << "Entries(ZZ"<< name << ")="
		  << std::hex << zone.begin() << "(" << zone.length() << ")";
	else
		f << "___";
	ascii().addPos(actPos);
	ascii().addNote(f.str().c_str());

	return ok;
}

// basic function to check if the header is ok
bool WPS4Parser::checkHeader(WPSHeader *header, bool /*strict*/)
{
	RVNGInputStreamPtr input = getInput();
	if (!input || !checkFilePosition(0x100))
	{
		WPS_DEBUG_MSG(("WPS4Parser::checkHeader: file is too short\n"));
		return false;
	}

	/* let's do the strict minimum, we does not want to break old code */
	input->seek(0x0, librevenge::RVNG_SEEK_SET);
	uint8_t firstOffset = libwps::readU8(input);
	uint8_t secondOffset = libwps::readU8(input);
	if (secondOffset != 0xFE || firstOffset > 7) return false;
	if (header)
	{
		if (firstOffset==1)
			header->setNeedEncoding(true);
		header->setMajorVersion(firstOffset<4 ? 2 : firstOffset<6 ? 3 : 4);
	}
	return true;
}

// read the document structure ...
bool WPS4Parser::findZones()
{
	RVNGInputStreamPtr input = getInput();

	if (!checkFilePosition(0x100))
	{
		WPS_DEBUG_MSG(("WPS4Parser::findZones: error: incomplete header\n"));
		throw libwps::ParseException();
	}

	input->seek(0x0, librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f, f2;
	f << "Entries(ZZHeader):";
	int vers = libwps::read8(input);
	long val = libwps::read8(input); // always 0xfe?
	int apCreator = libwps::readU16(input);
	int worksVersion = 0;
	switch (vers)
	{
	case 1:
		m_state->m_isDosFile = true;
		switch (apCreator)
		{
		case 0xda1:
			apCreator = 2;
			WPS_FALLTHROUGH;
		case 0:
		case 1:
			worksVersion = 1;
			f << "vers=dos" << 1+apCreator << ",";
			apCreator = 0;
			break;
		default:
			//  checkme
			worksVersion= (apCreator<10000) ? 1 : 2;
			f << "vers=dos3/win2,";
			break;
		}
		break;
	case 4:
		worksVersion=3;
		if (apCreator==0x4e27)
		{
			f << "vers=Win3.0,";
			apCreator = 0;
		}
		else
			f << "vers=Win3,";
		break;
	case 6:
		worksVersion=4;
		if (apCreator==0x5375)
		{
			f << "vers=Win4.0,";
			apCreator = 0;
		}
		else
			f << "vers=Win4,";
		break;
	default:
		WPS_DEBUG_MSG(("WPS4Parser::findZones: can not read the version\n"));
		f << "vers=unkn:" << vers << ",";
		break;
	}
	if (worksVersion)
		setVersion(worksVersion);
	if (val != -2) f << "##unk=" << val << ",";
	if (apCreator && apCreator != 0x4755)
		f << "##apCreator=" << std::hex << apCreator << std::dec << ",";

	f << "unkn1=("; // in general : same number appear two time
	for (int i = 0; i < 2; i++)
		f << std::hex << libwps::readU32(input) << ",";
	f << "),dim?=(";
	for (int i = 0; i < 2; i++)
		f << libwps::readU16(input)/1440. << ",";
	f << "),";
	auto oem=int(libwps::readU16(input));
	if ((oem>>4)&0x7ff)
	{
		auto type=libwps_tools_win::Font::getTypeForOEM((oem>>4)&0x7ff);
		if (type != libwps_tools_win::Font::UNKNOWN)
		{
			m_state->m_OEMFontType = type;
			f << "codePage=" << libwps_tools_win::Font::getTypeName(type).cstr() << ",";
		}
		else
			f << "#codePage=" << std::hex << ((oem>>4)&0x7ff) << std::dec << ",";
	}
	if (oem & 0x800f) f << "codePage[extra]=" << std::hex << (oem & 0x800f) << std::dec << ",";
	val=libwps::read16(input);
	if (val) f << "unkn2=" << val << ",";
	f << "dim2=" << libwps::read16(input)/1440. << ",";
	val = libwps::read32(input);
	if (val) f << std::dec << "unkn3=" << val <<",";
	ascii().addPos(0);
	ascii().addNote(f.str().c_str());

	f.str("");

	// 0x1e -> 0x64
	if (!m_textParser->readEntries()) return false;
	// 0x64 -> 0x80
	long actPos = 0x64;
	input->seek(actPos, librevenge::RVNG_SEEK_SET);
	readDocDim();

	if (version() <= 1)
	{
		// CHECKME:
		ascii().addPos(0x80);
		ascii().addNote("ZZHeader-I");
		ascii().addPos(0xd0);
		ascii().addNote("ZZHeader-filename");
		return true;
	}

	actPos = 0x80;
	input->seek(actPos, librevenge::RVNG_SEEK_SET);
	parseEntry("EOBJ");

	actPos = 0x86;
	input->seek(actPos, librevenge::RVNG_SEEK_SET);
	f.str("");
	f << std::hex;
	// {-1,-1}|{0,0}, 0,0, 0x[08][03][235=2col 6=3col 8=4col], 0|425|720
	for (int i = 0; i < 4; i++)
	{
		val = libwps::readU16(input);
		if (!val)
		{
			f << "_,";
			continue;
		}
		f << val << ",";
	}
	f << std::dec;
	val = libwps::readU16(input);
	if (val & 0xFF00) f << "#unkn=" << (val >> 8) << ",";
	val &= 0xFF;
	auto numCols = int(val/2);
	if (numCols >= 1 && numCols <= 13)
	{
		if (version() >= 3)
			m_state->m_numColumns = numCols;
		else
		{
			// can this happen ?
			WPS_DEBUG_MSG(("WPS4Parser::findZones: find some column in dos file: ignored\n"));
		}
		if (numCols != 1)
		{
			f << "numCols=" << numCols;
			if (val & 1) f << "[borderLim]";
		}
		f << ",";
	}
	else if (val)
		f << "##cols,";
	int colSep = libwps::readU16(input);
	if (colSep)
		f << "colSep=" << colSep/1440. << ",";
	ascii().addPos(actPos);
	if (f.str().length())
	{
		f2.str("");
		f2 << "ZZHeader-I(unkn):"<< f.str();
		ascii().addNote(f2.str().c_str());
	}
	else ascii().addNote("___");

	parseEntry("End1");

	actPos = 0x98;
	input->seek(actPos, librevenge::RVNG_SEEK_SET);
	f.str("");
	f << "ZZHeader-II:";
	bool empty = true;
	auto begP = long(libwps::readU32(input));
	if (begP)
	{
		if (begP <= 0 || !checkFilePosition(begP)) f << "###";
		else
		{
			ascii().addPos(begP);
			ascii().addNote("EOText");
		}
		f << std::hex << "EOText=" << begP << ",";
		empty = false;
	}

	int num = libwps::readU16(input);
	long sz = libwps::readU16(input);
	begP = long(libwps::readU32(input));
	if (begP)
	{
		if (begP <= 0 || !checkFilePosition(begP)) f << "###";
		else if (checkFilePosition(begP+num *sz))
		{
			WPSEntry zone;
			zone.setType("PRNT");
			for (int i = 0; i < num; i++)
			{
				zone.setBegin(begP + i*sz);
				zone.setLength(sz);
				zone.setId(i);
				getNameEntryMap().insert(NameMultiMap::value_type(zone.type(), zone));

				ascii().addPos(zone.begin());
				std::string nm = "ZZPRNT(";
				nm+= char('0'+i);
				nm+=')';
				ascii().addNote(nm.c_str());
				ascii().addPos(zone.end());
				ascii().addNote("_");
			}
		}
		else
		{
			ascii().addPos(begP);
			ascii().addNote("ZZPRNT");
		}

		f << std::hex << "Entries(ZZPRNT)=" << begP << "(" << sz << ")";
		if (num != 1) f << "x" << num;
		empty = false;
	}

	ascii().addPos(actPos);
	if (empty) ascii().addNote("___");
	else ascii().addNote(f.str().c_str());

	parseEntry("DTTM");
	parseEntry("DocWInfo");

	actPos = 0xb0;
	input->seek(actPos, librevenge::RVNG_SEEK_SET);
	f.str("");
	f << "ZZHeader-III:";
	empty = true;

	// ok, try to see if we can find other things...
	int i = 0;
	while (input->tell()+4 <= 0xd0)
	{
		val = long(libwps::readU32(input));
		if (val)
		{
			f << "f" << std::dec << i << "=" << std::hex << val << ",";
			empty = false;
		}
		i++;
	}

	ascii().addPos(actPos);
	if (empty) ascii().addNote("___");
	else ascii().addNote(f.str().c_str());

	// find always a list of 0 here
	ascii().addPos(0xd0);
	ascii().addNote("ZZHeader-filename");

	return true;
}

// ------------------------------------------------------------
// functions to read the document's dimension or printer structures
// ------------------------------------------------------------

// Read the page format from the file.
bool WPS4Parser::readDocDim()
{
	WPSPageSpan page = WPSPageSpan();
	RVNGInputStreamPtr &input = getInput();
	input->seek(0x64, librevenge::RVNG_SEEK_SET);
	long actPos = input->tell();

	libwps::DebugStream f;
	f << "Entries(DocDim):";
	double margin[4], size[2];
	for (double &i : margin) i = libwps::readU16(input)/1440.;
	for (double &i : size) i = libwps::readU16(input)/1440.;

	if (margin[0]+margin[1] > size[0] || margin[2]+margin[3] > size[1])
	{
		WPS_DEBUG_MSG(("WPS4Parser::readDocDim: error: the margins are too large for the page size\n"));
		return false;
	}
	page.setMarginTop(margin[0]);
	page.setMarginBottom(margin[1]);
	page.setMarginLeft(margin[2]);
	// decrease a little right margin if possible
	double rMargin = margin[3] > 0.4 ? margin[3]-0.2 : 0.5*margin[3];
	page.setMarginRight(rMargin);
	page.setFormLength(size[0]);
	page.setFormWidth(size[1]);

	int unkns[8];
	for (int &unkn : unkns) unkn = libwps::readU16(input);
	int page_orientation = unkns[5];
	if (page_orientation == 0)
		page.setFormOrientation(WPSPageSpan::PORTRAIT);
	else if (page_orientation == 1)
		page.setFormOrientation(WPSPageSpan::LANDSCAPE);
	else
	{
		WPS_DEBUG_MSG(("WPS4Parser::readDocDim: error: bad page orientation code\n"));
	}
	m_state->m_pageSpan = page;

	f << "margin=("<<  margin[2] << "x" << margin[0]
	  << ", " << margin[3] << "x" << margin[1] << "), ";
	f << "size=" << size[1] << "x" << size[0] << ",";
	if (page_orientation) f << "orien=" << page_orientation << ",";
	if (unkns[0]!=1) f << "firstPage=" << unkns[0] << ",";
	if (unkns[1]==1)
	{
		m_state->m_noFirstPageHeader=true;
		f << "noPage1Header,";
	}
	else if (unkns[1]) f << "#noPage1Header=" << unkns[1] << ",";
	if (unkns[2]==1)
	{
		m_state->m_noFirstPageHeader=true;
		f << "noPage1Footer,";
	}
	else if (unkns[2]) f << "#noPage1Footer=" << unkns[2] << ",";
	f << "headerH=" << unkns[3]/1440. << ",";
	f << "footerH=" << unkns[4]/1440. << ",";
	if (unkns[6] != 100) f << "zoom=" << unkns[6] << "%,";
	if (unkns[7]) f << "#unkn=" << unkns[7] << ","; // alway 0 ?
	ascii().addPos(actPos);
	ascii().addNote(f.str().c_str());
	return true;
}

// PRNT : the printer definition
bool WPS4Parser::readPrnt(WPSEntry const &entry)
{
	if (!entry.valid()) return false;

	RVNGInputStreamPtr &input = getInput();
	input->seek(entry.begin(), librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f;

	long length = entry.length();
	if (length < 0x174)
	{
		WPS_DEBUG_MSG(("WPS4Parser::readPrnt: length::=%ld is to short\n", length));
		return false;
	}
	f << std::hex;

	for (int st = 0; st < 2; st++)
	{
		float dim[8];
		for (int i = 0; i < 8; i++)
		{
			if (i == 4 || i == 5)
				dim[i] = float(libwps::readU32(input)/1440.);
			else
				dim[i] = float(libwps::read32(input)/1440.);
		}
		f << "dim"<< st << "=" << dim[5] << "x" << dim[4] << ",";
		f << "margin"<< st << "=[" << dim[0] << "x" << dim[2] << ","
		  << dim[3] << "x" << dim[1] << "],";
		f << "head/foot??"<< st << "=" << dim[6] << "x" << dim[7] << ",";
	}

	f << std::dec;
	long val;

	/* I find f0=1, f1=0|1, f2=0|1, f4=100, f5=0|15, f23=0|372
	   Note: f23=372 and strLen!=0 only when f1=1 and f2=1 ( probably related)
	 */
	for (int i = 0; i < 24; i++)
	{
		val = libwps::read32(input);
		if (val) f << "f" << i << "=" << val << ",";
	}
	val = libwps::read32(input);
	if (val) f << "strLen?=" << val << ",";

	/* I only find 0 here */
	for (int i = 0; i < 52; i++)
	{
		val = libwps::read32(input);
		if (val) f << "g" << i << "=" << val << ",";
	}

	ascii().addPos(entry.begin());
	ascii().addNote(f.str().c_str());

	length -= 372;

	if (length)
	{
		/* In my files, I find one time a strLen string: a header/footer? */
		ascii().addPos(input->tell());
		f.str("");
		f << "ZZPRNT(II):";
		for (long i = 0; i < length; i++)
			f << char(libwps::readU8(input));
		ascii().addNote(f.str().c_str());
	}
	return true;
}

// Link : in Works4 seems to contain data corresponding to the page border, ...
bool WPS4Parser::readDocWindowsInfo(WPSEntry const &entry)
{
	if (!entry.valid()) return false;

	RVNGInputStreamPtr &input = getInput();
	long length = entry.length();
	if (length < 0x154)
	{
		WPS_DEBUG_MSG(("WPS4Parser::readDocWindowsInfo: length::=%ld is to short\n", length));
		return false;
	}

	input->seek(entry.begin(), librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f;

	std::string str("");
	int pos = 0, debSPos = -1;
	// Find in one file str[0]="C:\Databases\Elem 02-03.wdb",str[120]="Query1"
	while (pos++ < 0x132)   // try to find some strings
	{
		auto c = char(libwps::read8(input));
		if (c == '\0')
		{
			if (debSPos>=0)
				f << "str[" << std::hex << debSPos << "]=\"" << str << "\",";
			debSPos = -1;
			str = std::string("");
			continue;
		}

		if (debSPos < 0) debSPos = pos-1;
		str += c;
	}
	if (debSPos>=0)
		f << "str[" << debSPos << "]=\"" << str << "\",";
	ascii().addPos(entry.begin());
	ascii().addNote(f.str().c_str());
	f.str("");


	input->seek(entry.begin()+0x132, librevenge::RVNG_SEEK_SET);
	f << "ZZDocWInfo(II):" << std::hex;
	// f0=f1=-1 in one file, f0=f1=0 in another file
	// f0=e6|1b0|2d0 ( but 2d0 in 2/3 of the files), 100<f1<438
	for (int i = 0; i < 2; i++)
	{
		long val = libwps::read32(input);
		if (val)
			f << "f" << i << "=" << val << ",";
	}

	int dim[2]; // TB, LR
	for (int &i : dim) i = libwps::read16(input);
	if (dim[0] || dim[1])
		f << "pageBorderDist=" << dim[1]/1440. << "x" << dim[0]/1440. << ",";
	long val = libwps::readU8(input);
	WPSColor color;
	if (val && getColor(int(val), color))
		f << "pageBorderColor=" << color << ",";
	else if (val)
		f << "#pageBorderColor=" << std::hex << val << std::dec << ",";
	val = libwps::readU8(input);
	if (val)
		f << "pageBorderStyle=" << val << ",";
	val = long(libwps::readU32(input));
	if (val&2)
		f << "pageBorderShaded,";
	if (val&1)
		f << "firstPageBorder,";
	val &= 0xFFFFFFFCL;
	if (val)
		f << "#pageBorder?=" << std::hex << val << std::dec << ",";

	// in one file: nothing ( case when f0=f1=-1)
	//
	//        h0-h1: seems to correspond to a text zone
	for (int i = 0; i < 4; i++)
	{
		long val_ = libwps::read32(input);
		if (i == 0)
		{
			f << "textpos?=[" << 0x100+val_<< ",";
			continue;
		}
		if (i == 1)
		{
			f << 0x100+val_ << "],";
			continue;
		}
		if (val_)
			f << "h" << i << "=" << val_ << ",";
	}

	ascii().addPos(entry.begin()+0x132);
	ascii().addNote(f.str().c_str());

	if (input->tell() != entry.end())
	{
		// can this happens ?
		ascii().addPos(input->tell());
		ascii().addNote("ZZDocWInfo(III)");
	}
	return true;
}

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
