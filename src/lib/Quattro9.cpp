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

#include <algorithm>
#include <cmath>
#include <sstream>
#include <stack>
#include <utility>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WKSContentListener.h"
#include "WKSSubDocument.h"

#include "WPSCell.h"
#include "WPSEntry.h"
#include "WPSFont.h"
#include "WPSHeader.h"
#include "WPSOLEParser.h"
#include "WPSOLEStream.h"
#include "WPSPageSpan.h"
#include "WPSStream.h"
#include "WPSStringStream.h"

#include "Quattro9Graph.h"
#include "Quattro9Spreadsheet.h"

#include "Quattro9.h"

using namespace libwps;

//! Internal: namespace to define internal class of Quattro9Parser
namespace Quattro9ParserInternal
{
//! Internal: the subdocument of a WPS4Parser
class SubDocument final : public WKSSubDocument
{
public:
	//! constructor for a text entry
	SubDocument(RVNGInputStreamPtr const &input, Quattro9Parser &pars, bool header) :
		WKSSubDocument(input, &pars), m_header(header) {}
	//! destructor
	~SubDocument() final {}

	//! operator==
	bool operator==(std::shared_ptr<WPSSubDocument> const &doc) const final
	{
		if (!doc || !WKSSubDocument::operator==(doc))
			return false;
		auto const *sDoc = dynamic_cast<SubDocument const *>(doc.get());
		if (!sDoc) return false;
		return m_header == sDoc->m_header;
	}

	//! the parser function
	void parse(std::shared_ptr<WKSContentListener> &listener, libwps::SubDocumentType subDocumentType) final;
	//! a flag to known if we need to send the header or the footer
	bool m_header;
};

void SubDocument::parse(std::shared_ptr<WKSContentListener> &listener, libwps::SubDocumentType)
{
	if (!listener.get())
	{
		WPS_DEBUG_MSG(("Quattro9ParserInternal::SubDocument::parse: no listener\n"));
		return;
	}
	if (!dynamic_cast<WKSContentListener *>(listener.get()))
	{
		WPS_DEBUG_MSG(("Quattro9ParserInternal::SubDocument::parse: bad listener\n"));
		return;
	}

	Quattro9Parser *pser = m_parser ? dynamic_cast<Quattro9Parser *>(m_parser) : nullptr;
	if (!pser)
	{
		listener->insertCharacter(' ');
		WPS_DEBUG_MSG(("Quattro9ParserInternal::SubDocument::parse: bad parser\n"));
		return;
	}
	pser->sendHeaderFooter(m_header);
}

//! a zone name
struct ZoneName
{
	//! constructor
	explicit ZoneName(char const *name, char const *extra=nullptr)
		: m_name(name)
		, m_extra(extra==nullptr ? "" : extra)
	{
	}
	//! the zone name
	std::string m_name;
	//! the extra data
	std::string m_extra;
};

//! the state of Quattro9Parser
struct State
{
	//! constructor
	State(libwps_tools_win::Font::Type fontType, char const *password)
		: m_fontType(fontType)
		, m_version(-1)
		, m_password(password)
		, m_metaData()
		, m_fontNamesList()
		, m_fontsList()
		, m_idToExternalFileMap()
		, m_idToExternalNameMap()
		, m_idToFieldMap()
		, m_actualSheet(-1)
		, m_actualStrings()
		, m_isEncrypted(false)
		, m_isDecoded(false)
		, m_idToZoneNameMap()
	{
	}

	//! return the default font style
	libwps_tools_win::Font::Type getDefaultFontType() const
	{
		if (m_fontType != libwps_tools_win::Font::UNKNOWN)
			return m_fontType;
		return libwps_tools_win::Font::WIN3_WEUROPE;
	}
	//! returns a default font with file's version */
	static WPSFont getDefaultFont()
	{
		WPSFont res;
		res.m_name="Times New Roman";
		res.m_size=12;
		return res;
	}
	//! init the zone name map
	void initZoneNameMap();
	//! the user font type
	libwps_tools_win::Font::Type m_fontType;
	//! the file version
	int m_version;
	//! the password (if known)
	char const *m_password;
	//! the meta data
	librevenge::RVNGPropertyList m_metaData;
	//! the font name list
	std::vector<librevenge::RVNGString> m_fontNamesList;
	//! the font list
	std::vector<WPSFont> m_fontsList;
	//! map id to filename
	std::map<int, librevenge::RVNGString> m_idToExternalFileMap;
	//! map id to external name
	std::map<int, librevenge::RVNGString> m_idToExternalNameMap;
	//! map id to field
	std::map<int, std::pair<librevenge::RVNGString,QuattroFormulaInternal::CellReference> >m_idToFieldMap;
	//! the actual sheet id
	int m_actualSheet;
	//! the actual string list
	std::pair<std::shared_ptr<WPSStream>, std::vector<TextEntry> > m_actualStrings;
	//! true if the file is encrypted
	bool m_isEncrypted;
	//! true if the main stream has been decoded
	bool m_isDecoded;
	//! map zone id to zone name
	std::map<int, ZoneName> m_idToZoneNameMap;
private:
	State(State const &)=delete;
	State &operator=(State const &)=delete;
};

void State::initZoneNameMap()
{
	if (!m_idToZoneNameMap.empty())
		return;
	m_idToZoneNameMap=std::map<int,ZoneName>
	{
		{ 0x1, ZoneName("File", "header")},
		{ 0x2, ZoneName("File", "pointer")},
		{ 0x3, ZoneName("File", "setting")},
		{ 0x4, ZoneName("File", "password")},
		{ 0x5, ZoneName("File", "user")},
		{ 0x6, ZoneName("Font", "name")},
		{ 0x7, ZoneName("Font", "style")},
		{ 0x8, ZoneName("UserFormat")},
		{ 0x9, ZoneName("Style", "name")},
		{ 0xa, ZoneName("Cell", "style")},
		{ 0xb, ZoneName("DLLIdFunct", "lib") },
		{ 0xc, ZoneName("DLLIdFunct", "func") },
		{ 0x101, ZoneName("Group") }, // group of spreadsheet, column, cell, ...
		{ 0x401, ZoneName("Document", "begin") },
		{ 0x402, ZoneName("Document", "end") },
		{ 0x403, ZoneName("Document", "external,link") },
		{ 0x404, ZoneName("Document", "external,link,end") },
		{ 0x405, ZoneName("Document", "external,link,name") },
		{ 0x406, ZoneName("Document", "fields")},
		{ 0x407, ZoneName("Document", "strings")},
		{ 0x408, ZoneName("Document", "formula")},
		{ 0x411, ZoneName("Document", "sep")}, // the document main data seems to be define before
		{ 0x430, ZoneName("Document", "num,sheet") },
		{ 0x488, ZoneName("Selection")}, // checkme
		{ 0x601, ZoneName("Spreadsheet", "begin") }, // or maybe a shape
		{ 0x602, ZoneName("Spreadsheet", "end") },
		{ 0x613, ZoneName("Spreadsheet", "info") },
		{ 0x617, ZoneName("Spreadsheet", "page,break") },
		{ 0x61d, ZoneName("Spreadsheet", "join,cells") },
		{ 0x631, ZoneName("Spreadsheet", "row,def") },
		{ 0x632, ZoneName("Spreadsheet", "col,def") },
		{ 0x633, ZoneName("Spreadsheet", "row,size") },
		{ 0x634, ZoneName("Spreadsheet", "col,size") },
		{ 0x635, ZoneName("Spreadsheet", "rows,size") },
		{ 0x636, ZoneName("Spreadsheet", "cols,size") },
		{ 0x637, ZoneName("Spreadsheet", "row,dSize") }, // related to a change of size determined by the cell content
		{ 0x638, ZoneName("Spreadsheet", "col,dSize") }, // checkme
		{ 0xa01, ZoneName("Spreadsheet", "col,begin") },
		{ 0xa02, ZoneName("Spreadsheet", "col,end") },
		{ 0xa03, ZoneName("Spreadsheet", "col,sep") }, // always follow a01, a separator?
		{ 0xc01, ZoneName("Spreadsheet", "cell,list") },
		{ 0xc02, ZoneName("Spreadsheet", "cell,res") },
		{ 0x1401, ZoneName("Graph", "begin") },
		{ 0x1402, ZoneName("Graph", "end") },
		{ 0x2001, ZoneName("Graph", "zone,begin") },
		{ 0x2002, ZoneName("Graph", "zone,end") },
		{ 0x2051, ZoneName("Graph", "header") },
		{ 0x2052, ZoneName("Graph", "name") },
		{ 0x2073, ZoneName("Button", "name") },

		{ 0x2131, ZoneName("Frame", "fill") },
		{ 0x2141, ZoneName("Frame", "pattern") },
		{ 0x2151, ZoneName("Frame", "unknown") },
		{ 0x2161, ZoneName("Frame", "name") },
		{ 0x2171, ZoneName("Frame", "header") },
		{ 0x2184, ZoneName("Frame", "name2") },

		{ 0x21d1, ZoneName("OLE", "name") },

		{ 0x2221, ZoneName("Shape", "basic") },
		{ 0x2371, ZoneName("Textbox", "style") },
		{ 0x2372, ZoneName("Textbox", "string") },
		{ 0x2374, ZoneName("Textbox", "zone4") },
		{ 0x2375, ZoneName("Textbox", "zone5") },

		{ 0x23d1, ZoneName("Shape", "unknown") }, // picture or group?

		{ 0x2f30, ZoneName("Pict", "image") },

		{ 0x2ff1, ZoneName("Pict", "zone1") },
		{ 0x2ff2, ZoneName("Pict", "dir") },
		{ 0x2ff3, ZoneName("Pict", "zone2") },
		{ 0x2ff4, ZoneName("Pict", "fill[bitmap]") },
	};
}

////////////////////////////////////////////////////////////
//   Text entry
////////////////////////////////////////////////////////////
librevenge::RVNGString TextEntry::getString(std::shared_ptr<WPSStream> const &stream, libwps_tools_win::Font::Type type) const
{
	if (empty()) return "";
	if (!stream || !stream->m_input)
	{
		WPS_DEBUG_MSG(("Quattro9ParserInternal::TextEntry::getString: can not read a stream\n"));
		return "";
	}
	auto input=stream->m_input;
	long actPos=input->tell();
	input->seek(m_entry.begin(), librevenge::RVNG_SEEK_SET);
	std::string string;
	for (long i=0; i<m_entry.length(); ++i)
	{
		auto c = char(libwps::readU8(input));
		if (c == '\0') break;
		if (c == 0xd)
		{
			static bool first=true;
			if (first)
			{
				WPS_DEBUG_MSG(("Quattro9ParserInternal::TextEntry::getString: find some EOL in string, use send\n"));
				first=false;
			}
			string.push_back(' ');
			continue;
		}
		string.push_back(c);
	}
	input->seek(actPos, librevenge::RVNG_SEEK_SET);
	return libwps_tools_win::Font::unicodeString(string, type);
}

std::string TextEntry::getDebugString(std::shared_ptr<WPSStream> const &stream) const
{
	std::stringstream s;
	s << getString(stream).cstr();
	if (m_flag) s << "[fl=" << std::hex << m_flag << std::dec <<  "]";
	if (!m_extra.empty()) s << "[" << m_extra << "]";
	return s.str();
}

void TextEntry::send(std::shared_ptr<WPSStream> const &stream, WPSFont const &font, libwps_tools_win::Font::Type type, WKSContentListenerPtr &listener)
{
	if (!listener)
	{
		WPS_DEBUG_MSG(("Quattro9ParserInternal::TextEntry::send: called without listener\n"));
		return;
	}
	if (empty()) return;
	if (!stream || !stream->m_input)
	{
		WPS_DEBUG_MSG(("Quattro9ParserInternal::TextEntry::send: can not read a stream\n"));
		return;
	}
	auto input=stream->m_input;
	long actPos=input->tell();
	input->seek(m_entry.begin(), librevenge::RVNG_SEEK_SET);
	bool prevEOL=false;
	int numFonts=int(m_fontsList.size());
	auto fontType=type;
	std::string text;
	for (long i=0; i<=m_entry.length(); ++i)
	{
		auto c=i==m_entry.length() ? '\0' : char(libwps::readU8(input));
		auto fIt=m_posFontIdMap.find(int(i));
		if ((fIt!=m_posFontIdMap.end() || c==0 || c==0x9 || c==0xa || c==0xd) && !text.empty())
		{
			listener->insertUnicodeString(libwps_tools_win::Font::unicodeString(text, fontType));
			text.clear();
		}
		if (i==m_entry.length()) break;
		if (fIt!=m_posFontIdMap.end())
		{
			auto id=fIt->second;
			if (id==0)
			{
				fontType=type;
				listener->setFont(font);
			}
			else if (id>0 && id<=numFonts)
			{
				auto newFont=m_fontsList[size_t(id-1)];
				if (!newFont.m_name.empty())
				{
					auto newType=libwps_tools_win::Font::getFontType(newFont.m_name);
					if (newType!=libwps_tools_win::Font::UNKNOWN)
						fontType=newType;
				}
				listener->setFont(newFont);
			}
			else
			{
				WPS_DEBUG_MSG(("Quattro9ParserInternal::TextEntry::send: unknown font %d\n", id));
			}
		}
		switch (c)
		{
		case 0:
			break;
		case 0x9:
			listener->insertTab();
			break;
		case 0xa:
			if (!prevEOL)
			{
				WPS_DEBUG_MSG(("Quattro9ParserInternal::TextEntry::send: find 0xa without 0xd\n"));
			}
			break;
		case 0xd:
			listener->insertEOL();
			break;
		default:
			text.push_back(c);
			break;
		}
		prevEOL=c==0xd;
	}
	input->seek(actPos, librevenge::RVNG_SEEK_SET);
}

}

// constructor, destructor
Quattro9Parser::Quattro9Parser(RVNGInputStreamPtr &input, WPSHeaderPtr &header, libwps_tools_win::Font::Type encoding, char const *password)
	: WKSParser(input, header)
	, m_listener()
	, m_state(new Quattro9ParserInternal::State(encoding, password))
	, m_graphParser(new Quattro9Graph(*this))
	, m_spreadsheetParser(new Quattro9Spreadsheet(*this))
{
}

Quattro9Parser::~Quattro9Parser()
{
}

int Quattro9Parser::version() const
{
	return m_state->m_version;
}

libwps_tools_win::Font::Type Quattro9Parser::getDefaultFontType() const
{
	return m_state->getDefaultFontType();
}

bool Quattro9Parser::getExternalFileName(int fId, librevenge::RVNGString &fName) const
{
	auto it = m_state->m_idToExternalFileMap.find(fId);
	if (it!=m_state->m_idToExternalFileMap.end())
	{
		fName=it->second;
		return true;
	}
	if (fId==2)
	{
		// checkme current file?
		return true;
	}
	WPS_DEBUG_MSG(("Quattro9Parser::getExternalFileName: can not find %d name\n", fId));
	return false;
}

bool Quattro9Parser::getField(int fId, librevenge::RVNGString &text,
                              QuattroFormulaInternal::CellReference &ref,
                              librevenge::RVNGString const &fileName) const
{
	ref.m_cells.clear();
	if (fId&0x4000)
	{
		auto it=m_state->m_idToExternalNameMap.find(fId & 0xbfff);
		if (it!=m_state->m_idToExternalNameMap.end() && !it->second.empty())
		{
			text=it->second;
			WKSContentListener::FormulaInstruction instr;
			instr.m_type=instr.F_Text;
			if (!fileName.empty())
			{
				instr.m_content="[";
				instr.m_content+=fileName.cstr();
				instr.m_content+="]";
			}
			instr.m_content+=it->second.cstr();
			ref.addInstruction(instr);
			return true;
		}
		WPS_DEBUG_MSG(("Quattro9Parser::getField: can not find %d name\n", fId&0xbfff));
		return false;
	}
	auto it = m_state->m_idToFieldMap.find(fId);
	if (it!=m_state->m_idToFieldMap.end())
	{
		text=it->second.first;
		ref=it->second.second;
		if (!fileName.empty())   // unsure
		{
			for (auto &r : ref.m_cells)
			{
				if (r.m_type==r.F_Cell || r.m_type==r.F_CellList)
					r.m_fileName=fileName;
			}
		}
		return true;
	}
	WPS_DEBUG_MSG(("Quattro9Parser::getField: can not find %d field\n", fId));
	return false;
}
//////////////////////////////////////////////////////////////////////
// interface with Quattro9Graph
//////////////////////////////////////////////////////////////////////
bool Quattro9Parser::getColor(int id, WPSColor &color) const
{
	return m_graphParser->getColor(id, color);
}

bool Quattro9Parser::getPattern(int id, WPSGraphicStyle::Pattern &pattern) const
{
	return m_graphParser->getPattern(id, pattern);
}

bool Quattro9Parser::sendPageGraphics(int sheetId) const
{
	return m_graphParser->sendPageGraphics(sheetId);
}

//////////////////////////////////////////////////////////////////////
// interface with Quattro9Spreadsheet
//////////////////////////////////////////////////////////////////////
bool Quattro9Parser::getFont(int id, WPSFont &font) const
{
	if (id < 0 || id>=int(m_state->m_fontsList.size()))
	{
		WPS_DEBUG_MSG(("Quattro9Parser::getFont: can not find font %d\n", id));
		return false;
	}
	font=m_state->m_fontsList[size_t(id)];
	return true;
}

Vec2f Quattro9Parser::getCellPosition(int sheetId, Vec2i const &cell) const
{
	return m_spreadsheetParser->getPosition(sheetId, cell);
}

void Quattro9Parser::addDocumentStrings()
{
	if (!m_state->m_actualStrings.first || m_state->m_actualStrings.second.empty())
		return;
	m_spreadsheetParser->addDocumentStrings(m_state->m_actualStrings.first, m_state->m_actualStrings.second);
	m_state->m_actualStrings.first.reset();
	m_state->m_actualStrings.second.clear();
}

// main function to parse the document
void Quattro9Parser::parse(librevenge::RVNGSpreadsheetInterface *documentInterface)
{
	RVNGInputStreamPtr input=getInput();
	if (!input)
	{
		WPS_DEBUG_MSG(("Quattro9Parser::parse: does not find main ole\n"));
		throw (libwps::ParseException());
	}

	if (!checkHeader(nullptr)) throw(libwps::ParseException());

	bool ok=false;
	try
	{
		ascii().setStream(input);
		ascii().open("MN0");
		if (checkHeader(nullptr) && readZones())
			m_listener=createListener(documentInterface);
		if (m_listener)
		{
			m_spreadsheetParser->setListener(m_listener);
			m_graphParser->setListener(m_listener);
			m_graphParser->updateState();
			m_spreadsheetParser->updateState();

			m_listener->startDocument();
			int numSheet=m_spreadsheetParser->getNumSpreadsheets();
			if (numSheet==0) ++numSheet;
			for (int i=0; i<numSheet; ++i)
				m_spreadsheetParser->sendSpreadsheet(i);
			m_listener->endDocument();
			m_listener.reset();
			ok = true;
		}
	}
	catch (...)
	{
		WPS_DEBUG_MSG(("Quattro9Parser::parse: exception catched when parsing MN0\n"));
		throw (libwps::ParseException());
	}

	ascii().reset();
	if (!ok)
		throw(libwps::ParseException());
}

std::shared_ptr<WKSContentListener> Quattro9Parser::createListener(librevenge::RVNGSpreadsheetInterface *interface)
{
	std::vector<WPSPageSpan> pageList;
	WPSPageSpan ps;
	int numSheet=m_spreadsheetParser->getNumSpreadsheets();
	if (numSheet<=0) numSheet=1;
	ps.setPageSpan(numSheet);
	pageList.push_back(ps);
	auto listener=std::make_shared<WKSContentListener>(pageList, interface);
	listener->setMetaData(m_state->m_metaData);
	return listener;
}

////////////////////////////////////////////////////////////
// low level
////////////////////////////////////////////////////////////
// read the header
////////////////////////////////////////////////////////////
bool Quattro9Parser::checkHeader(WPSHeader *header, bool strict)
{
	m_state.reset(new Quattro9ParserInternal::State(m_state->m_fontType, m_state->m_password));
	auto input=getInput();
	auto mainStream=std::make_shared<WPSStream>(input, ascii());
	if (!mainStream || !checkHeader(mainStream, strict))
		return false;
	if (header)
	{
		header->setMajorVersion(m_state->m_version);
		header->setCreator(libwps::WPS_QUATTRO_PRO);
		header->setKind(libwps::WPS_SPREADSHEET);
		header->setIsEncrypted(m_state->m_isEncrypted);
		header->setNeedEncoding(true);
	}
	return true;
}

bool Quattro9Parser::checkHeader(std::shared_ptr<WPSStream> stream, bool strict)
{
	if (!stream || !stream->checkFilePosition(14))
	{
		WPS_DEBUG_MSG(("Quattro9Parser::checkHeader: file is too short\n"));
		return false;
	}
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	m_state->m_version=2000;
	input->seek(0,librevenge::RVNG_SEEK_SET);
	// basic check: check if the first zone has type=1, length=14 and begins with QPW9
	if (libwps::readU16(input)!=1 || libwps::readU16(input)!=0xe || libwps::readU32(input)!=0x39575051)
		return false;
	if (strict)
	{
		input->seek(0,librevenge::RVNG_SEEK_SET);
		for (int i=0; i < 6; ++i)
		{
			if (!readZone(stream)) return false;
			if (m_state->m_isEncrypted) break;
		}
	}
	ascFile.addPos(0);
	ascFile.addNote(f.str().c_str());

	return true;
}

bool Quattro9Parser::readZones()
{
	m_graphParser->cleanState();
	m_spreadsheetParser->cleanState();
	m_state->initZoneNameMap();

	auto stream=std::make_shared<WPSStream>(getInput(), ascii());
	RVNGInputStreamPtr &input = stream->m_input;
	input->seek(0, librevenge::RVNG_SEEK_SET);
	while (stream->checkFilePosition(input->tell()+4))
	{
		if (!readZone(stream))
			break;
		if (m_state->m_isEncrypted && !m_state->m_isDecoded)
			throw(libwps::PasswordException());
	}
	if (!input->isEnd())
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readZones: find extra data\n"));
		libwps::DebugFile &ascFile=stream->m_ascii;
		ascFile.addPos(input->tell());
		ascFile.addNote("Entries(Unknown):###");
	}
	parseOLEStream(getFileInput(), "NativeContent_MAIN");
	return m_spreadsheetParser->getNumSpreadsheets();
}

bool Quattro9Parser::parseOLEStream(RVNGInputStreamPtr input, std::string const &avoid)
{
	if (!input || !input->isStructured())
	{
		WPS_DEBUG_MSG(("Quattro9Parser::parseOLEStream: oops, can not find the input stream\n"));
		return false;
	}
	std::map<std::string,size_t> dirToIdMap;
	WPSOLEParser oleParser(avoid, libwps_tools_win::Font::WIN3_WEUROPE,
	                       [&dirToIdMap](std::string const &dir)
	{
		if (dirToIdMap.find(dir)==dirToIdMap.end())
			dirToIdMap[dir]=dirToIdMap.size();
		return int(dirToIdMap.find(dir)->second);
	});
	oleParser.parse(input);
	oleParser.updateMetaData(m_state->m_metaData);
	auto objectMap=oleParser.getObjectsMap();
	std::map<librevenge::RVNGString,WPSEmbeddedObject> nameToObjectsMap;
	for (auto it : dirToIdMap)
	{
		if (it.first.empty()) continue;
		auto linkName=it.first;
		auto pos = linkName.find_last_of('/');
		if (pos != std::string::npos)
			linkName=linkName.substr(pos+1);
		if (!linkName.empty())
			nameToObjectsMap[linkName.c_str()]=objectMap.find(int(it.second))->second;
		for (int wh=0; wh<2; ++wh)
		{
			std::string name=it.first+"/"+(wh==0 ? "LinkInfo" : "BOlePart");
			RVNGInputStreamPtr cOle(input->getSubStreamByName(name.c_str()));
			if (!cOle)
			{
				WPS_DEBUG_MSG(("Quattro9Parser::parseOLEStream: oops, can not find link info for dir %s\n", name.c_str()));
				continue;
			}
			libwps::DebugFile asciiFile(cOle);
			asciiFile.open(libwps::Debug::flattenFileName(name));
			if (wh==1)
				readOleBOlePart(std::make_shared<WPSStream>(cOle,asciiFile));
			else
				readOleLinkInfo(std::make_shared<WPSStream>(cOle,asciiFile));
		}
	}
	if (!nameToObjectsMap.empty())
		m_graphParser->storeObjects(nameToObjectsMap);
	return true;
}

bool Quattro9Parser::readZone(std::shared_ptr<WPSStream> &stream)
{
	if (!stream)
		return false;
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto id = int(libwps::readU16(input));
	bool const bigBlock=id&0x8000;
	long sz = bigBlock ? long(libwps::readU32(input)) : long(libwps::readU16(input));
	int const headerSize=bigBlock ? 6 : 4;
	id &= 0x7fff;
	if (sz<0 || !stream->checkFilePosition(pos+headerSize+sz))
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readZone: size is bad\n"));
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}

	auto zIt=m_state->m_idToZoneNameMap.find(id);
	if (zIt==m_state->m_idToZoneNameMap.end())
		f << "Entries(Zone" << std::hex << id << std::dec << "A):";
	else if (zIt->second.m_extra.empty())
		f << "Entries(" << zIt->second.m_name << "):";
	else
		f << "Entries(" << zIt->second.m_name << ")[" << zIt->second.m_extra << "]:";
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	f.str("");
	bool ok = true, isParsed = false, needWriteInAscii = false;
	int val;
	input->seek(pos, librevenge::RVNG_SEEK_SET);
	switch (id)
	{
	case 0x1:
	{
		if (sz!=14) break;
		input->seek(pos+headerSize, librevenge::RVNG_SEEK_SET);
		std::string type; // QWP9
		for (int i=0; i<4; ++i) type+=char(libwps::readU8(input));
		f << type << ",";
		val=int(libwps::readU16(input)); // QPW 9: 0, QPW X9: 19
		m_state->m_version=2000+val+1;
		if (val) f << "vers=" << val << ",";
		for (int i=0; i<4; ++i)   // f1=ee|201, f4=1-5
		{
			val=int(libwps::read16(input));
			int const expected[]= {0, 1, 0, 0};
			if (val!=expected[i]) f << "f" << i << "=" << val << ",";
		}
		isParsed=needWriteInAscii=true;
		break;
	}
	// no data
	case 0x400:
	case 0x481:
	case 0x4a1:
	case 0x4a3:
		if (sz!=0) break;
		isParsed=needWriteInAscii=true;
		break;

	// bool
	case 0x48c: // 1
	case 0x48d: // 0
	case 0x4a4: // 0|3|4
	case 0x4a8: // 0
	case 0x4aa: // 0
	case 0x4ab: // 0
	case 0x4ac: // 0
	case 0x4ad: // 0
	case 0x4b2: // 1
	case 0x4b6: // 1
	case 0x4b9: // 0
		if (sz!=1) break;
		input->seek(pos+headerSize, librevenge::RVNG_SEEK_SET);
		val=int(libwps::read8(input));
		if (val)
			f << "f0=" << val << ",";
		isParsed=needWriteInAscii=true;
		break;
	// int
	case 0x430: // num spreadsheet?
	case 0x48a: // 0|3
	case 0x485: // 64
	case 0x487: // 0
	case 0x4b0: // 1
	case 0x4b4: // 64
	case 0x4b5: // 1|9 (2 by files)
	case 0xa03: // 0 (after a01 before c01 a03)
		if (sz!=2) break;
		input->seek(pos+headerSize, librevenge::RVNG_SEEK_SET);
		val=int(libwps::read16(input));
		if (val)
			f << "f0=" << val << ",";
		isParsed=needWriteInAscii=true;
		break;
	// 2 int
	case 0x482: // 0
	case 0x4ae: // 0
	case 0x4af: // 0
	case 0x2374: // 0
	case 0x2375: // 0
		if (sz!=4) break;
		input->seek(pos+headerSize, librevenge::RVNG_SEEK_SET);
		for (int i=0; i<2; ++i)
		{
			val=int(libwps::read16(input));
			if (val)
				f << "f" << i << "=" << val << ",";
		}
		isParsed=needWriteInAscii=true;
		break;
	// 3 int
	case 0x4a6: // 0
	case 0x4a7: // 0
		if (sz!=6) break;
		input->seek(pos+headerSize, librevenge::RVNG_SEEK_SET);
		for (int i=0; i<3; ++i)
		{
			val=int(libwps::read16(input));
			if (val)
				f << "f" << i << "=" << val << ",";
		}
		isParsed=needWriteInAscii=true;
		break;

	// zone with pointer

	// 2 file positions
	case 0x2: // f0=0, f1=pointer to a zone 402
	case 0x402:   // f0= pointer to a previous zone 401, f1=0
	{
		long fPos[2];
		isParsed=readFilePositions(stream,fPos);
		break;
	}
	case 0x3:
		if (sz<2) break;
		input->seek(pos+headerSize, librevenge::RVNG_SEEK_SET);
		val=int(libwps::readU16(input));
		if (val==0x20) break; // normal
		if (val&3) f << "protection=" << (val&3) << ",";
		if (val&0xffdc)
			f << "fl=" << std::hex << (val&0xffdc) << std::dec << ",";
		break;
	case 0x4:
		m_state->m_isEncrypted=true;
		if (sz==20)
		{
			m_state->m_isEncrypted=true;
			input->seek(pos+headerSize, librevenge::RVNG_SEEK_SET);
			uint16_t fileKey(libwps::readU16(input));
			f << "pass=" << std::hex << fileKey << std::dec << ",";
			f << "len=" << int(libwps::readU16(input)) << ",";
			isParsed = needWriteInAscii = true;
			std::vector<uint8_t> keys;
			keys.resize(16);
			for (auto &k : keys) k=uint8_t(libwps::readU8(input));
			// to check users password:
			//   call libwps::encodeLotusPassword(m_state->m_password, key, lotusKeys, someDefValues);
			//   and check if  int16_t(key<<8|key>>8)==fileKey
			if (!m_state->m_isDecoded)
			{
				auto newInput=decodeStream(input, keys);
				if (newInput)
				{
					// let's replace the current input by the decoded input
					m_state->m_isDecoded=true;
					stream->m_input=newInput;
					stream->m_ascii.setStream(newInput);
				}
			}
		}
		if (!m_state->m_isDecoded)
		{
			WPS_DEBUG_MSG(("Quattro9Parser::parse: can not decode the file\n"));
		}
		break;
	case 0x5:
	{
		if (sz<4) break;
		input->seek(pos+headerSize, librevenge::RVNG_SEEK_SET);
		for (int i=0; i<2; ++i)
		{
			Quattro9ParserInternal::TextEntry entry;
			if (readPString(stream, pos+headerSize+sz, entry))
				f << entry.getDebugString(stream) << ",";
			else
			{
				WPS_DEBUG_MSG(("Quattro9Parser::readZone[user]: can not read a string\n"));
				f << "###";
				break;
			}
		}
		isParsed=needWriteInAscii=true;
		break;
	}
	case 0x6:
		isParsed=readFontNames(stream);
		break;
	case 0x7:
		isParsed=readFontStyles(stream);
		break;
	case 0x8:
	{
		if (sz<6) break;
		input->seek(pos+headerSize, librevenge::RVNG_SEEK_SET);
		int fId=int(libwps::readU16(input));
		f << "id=" << fId << ",";
		val=int(libwps::readU16(input)); // 0
		if (val) f << "f0=" << val << ",";
		std::string format;
		for (long i=0; i<sz-4; ++i)
		{
			auto c = char(libwps::readU8(input));
			if (c == '\0') break;
			format.push_back(c);
		}
		if (!format.empty())
		{
			m_spreadsheetParser->addUserFormat(fId, libwps_tools_win::Font::unicodeString(format, libwps_tools_win::Font::WIN3_WEUROPE));
			f << format << ",";
		}
		if (input->tell()!=pos+headerSize+sz) ascFile.addDelimiter(input->tell(),'|');
		isParsed=needWriteInAscii=true;
		break;
	}
	case 0x9:
	{
		if (sz<2) break;
		input->seek(pos+headerSize, librevenge::RVNG_SEEK_SET);
		int N=int(libwps::readU16(input));
		for (int i=0; i<N; ++i)
		{
			int fId=int(libwps::readU16(input));
			f << "Styl" << fId << "=[";
			val=int(libwps::readU16(input));
			if (val!=fId) f << "id2=" << val << ",";
			Quattro9ParserInternal::TextEntry entry;
			if (readPString(stream, pos+headerSize+sz, entry))
				f << entry.getDebugString(stream) << ",";
			else
			{
				WPS_DEBUG_MSG(("Quattro9Parser::readZone[style]: can not read a name\n"));
				f << "###";
				break;
			}
			f << "],";
		}
		isParsed=needWriteInAscii=true;
		break;
	}
	case 0xa:
		ok=isParsed=m_spreadsheetParser->readCellStyles(stream);
		break;
	case 0xb:
	case 0xc:
	{
		if (sz<6) break;
		input->seek(pos+headerSize, librevenge::RVNG_SEEK_SET);
		val=int(libwps::readU16(input));
		f << "id=" << val << ",";
		Quattro9ParserInternal::TextEntry entry;
		if (!readPString(stream, pos+headerSize+sz, entry))
			f << "###";
		else
		{
			m_spreadsheetParser->addDLLIdName(val, entry.getString(stream), id==0xb);
			f << entry.getDebugString(stream) << ",";
		}
		isParsed=needWriteInAscii=true;
		break;
	}
	case 0x101:   // list of 601|a01|c01 zone...
	{
		if (sz<6 || (sz%4)!=2) break;
		input->seek(pos+headerSize, librevenge::RVNG_SEEK_SET);
		f << "type=" << std::hex << int(libwps::readU16(input)) << std::dec << ",";
		int dim[2];
		for (auto &d : dim) d=int(libwps::read16(input));
		f << "limits=" << Vec2i(dim[0], dim[1]) << ",";
		int N=int((sz-6)/4);
		f << "pos=[";
		for (int i=0; i<N; ++i)
			f << std::hex << libwps::readU32(input) << std::dec << ",";
		f << "],";
		isParsed=needWriteInAscii=true;
		break;
	}
	case 0x403:
	{
		if (sz<16) break;
		input->seek(pos+headerSize, librevenge::RVNG_SEEK_SET);
		int lId=int(libwps::readU16(input));
		f << "id=" << lId << ",";
		for (int i=0; i<2; ++i)   // 0
		{
			val=int(libwps::read16(input));
			if (val) f << "f" << i << "=" << val << ",";
		}
		Quattro9ParserInternal::TextEntry entry;
		if (!readPString(stream, pos+headerSize+sz, entry))
			f << "###";
		else
		{
			f << entry.getDebugString(stream) << ",";
			auto &map = m_state->m_idToExternalFileMap;
			if (map.find(lId)!=map.end())
			{
				WPS_DEBUG_MSG(("Quattro9Parser::readZone[link,file]: a link with id=%d already exists\n", lId));
				f << "###dup,";
			}
			else
				map[lId]=entry.getString(stream);
		}
		isParsed=needWriteInAscii=true;
		break;
	}
	case 0x404:
		if (sz!=0) break;
		m_state->m_actualStrings.first.reset();
		m_state->m_actualStrings.second.clear(); // clean potential external strings
		isParsed=needWriteInAscii=true;
		break;
	case 0x405:
	{
		if (sz<20) break;
		input->seek(pos+headerSize, librevenge::RVNG_SEEK_SET);
		val=int(libwps::read16(input)); // 0
		if (val) f << "f0=" << val << ",";
		int lId=int(libwps::readU16(input));
		f << "id=" << lId << ",";
		for (int i=0; i<2; ++i)   // 0, lId
		{
			val=int(libwps::read16(input));
			if (val) f << "f" << i+2 << "=" << val << ",";
		}
		Quattro9ParserInternal::TextEntry entry;
		if (!readPString(stream, pos+headerSize+sz, entry))
			f << "###";
		else
		{
			f << entry.getDebugString(stream) << ",";
			auto &map = m_state->m_idToExternalNameMap;
			if (map.find(lId)!=map.end())
			{
				WPS_DEBUG_MSG(("Quattro9Parser::readZone[link,name]: a link with id=%d already exists\n", lId));
				f << "###dup,";
			}
			else
				map[lId]=entry.getString(stream);
		}
		isParsed=needWriteInAscii=true;
		break;
	}
	case 0x406:
		isParsed=readDocumentFields(stream);
		break;
	case 0x407:
		// used to store the strings which appears in cell, but also some strings which appear in link, ...
		ok=isParsed=readDocumentStrings(stream);
		break;
	case 0x408:
		addDocumentStrings();
		ok=isParsed=m_spreadsheetParser->readDocumentFormulas(stream);
		break;
	case 0x411: // always with 0
		if (sz<1) break;
		val=int(libwps::readU8(input));
		if (val) f << "f0=" << val << ",";
		addDocumentStrings(); // last time to save the main document strings
		isParsed=needWriteInAscii=true;
		break;
	case 0x601:
		ok=isParsed=m_spreadsheetParser->readBeginSheet(stream, m_state->m_actualSheet);
		break;
	case 0x602:
		ok=isParsed=m_spreadsheetParser->readEndSheet(stream);
		m_state->m_actualSheet=-1;
		break;
	case 0x613:
		if (sz!=24) break;
		input->seek(pos+headerSize, librevenge::RVNG_SEEK_SET);
		for (int i=0; i<6; ++i)   // 0
		{
			val=int(libwps::read16(input));
			if (val) f << "f" << i << "=" << val << ",";
		}
		for (int i=0; i<4; ++i)
		{
			val=int(libwps::readU16(input));
			int const expected[]= {0x759c,0x8800,0xe43c, 0x7e37};
			if (val!=expected[i]) f << "f" << i+6 << "=" << std::hex << val << std::dec << ",";
		}
		for (int i=0; i<4; ++i)
		{
			val=int(libwps::read8(input));
			int const expected[]= {4,3,5,4};
			if (val!=expected[i]) f << "f" << i+10 << "=" << val << ",";
		}
		isParsed=needWriteInAscii=true;
		break;
	case 0x617:
		isParsed=m_spreadsheetParser->readPageBreak(stream);
		break;
	case 0x61d:
		isParsed=m_spreadsheetParser->readMergedCells(stream);
		break;
	case 0x631:
	case 0x632:
		isParsed=m_spreadsheetParser->readColRowDefault(stream);
		break;
	case 0x633:
	case 0x634:
		isParsed=m_spreadsheetParser->readColRowDimension(stream);
		break;
	case 0x635:
	case 0x636:
		isParsed=m_spreadsheetParser->readColRowDimensions(stream);
		break;
	case 0x637:
	case 0x638: // checkme
		if (sz!=6) break;
		input->seek(pos+headerSize, librevenge::RVNG_SEEK_SET);
		f << "id=" << libwps::readU32(input) << ",";
		val=int(libwps::readU8(input)); // height/width in dpi?
		f << "size?=" << val << ",";
		val=int(libwps::readU8(input)); // 0, da, 5b
		if (val) f << "fl=" << std::hex << val << std::dec << ",";
		isParsed=needWriteInAscii=true;
		break;
	case 0xa01:
		ok=isParsed=m_spreadsheetParser->readBeginColumn(stream);
		break;
	case 0xa02:
		ok=isParsed=m_spreadsheetParser->readEndColumn(stream);
		break;
	case 0xc01:
		ok = isParsed = m_spreadsheetParser->readCellList(stream);
		break;
	case 0xc02:
		isParsed = m_spreadsheetParser->readCellResult(stream);
		break;

	case 0x1401:
	case 0x1402:
		isParsed = m_graphParser->readBeginEnd(stream, m_state->m_actualSheet);
		break;

	case 0x2001:
	case 0x2002:
		isParsed = m_graphParser->readBeginEndZone(stream);
		break;

	case 0x2051:
		isParsed=m_graphParser->readGraphHeader(stream);
		break;

	case 0x2131:
		isParsed=m_graphParser->readFrameStyle(stream);
		break;
	case 0x2141:
		isParsed=m_graphParser->readFramePattern(stream);
		break;
	case 0x2171:
		isParsed=m_graphParser->readFrameHeader(stream);
		break;
	case 0x21d1:
		isParsed=m_graphParser->readOLEName(stream);
		break;
	case 0x2221:
	case 0x23d1:
		isParsed=m_graphParser->readShape(stream);
		break;
	case 0x2371:
		isParsed=m_graphParser->readTextboxStyle(stream);
		break;
	case 0x2372:
		isParsed=m_graphParser->readTextboxText(stream);
		break;

	// case 0x2371: fSz1,fSz2, 0,0, name1,name2, ....
	case 0x2052:
	case 0x2073:
	case 0x2161:
	case 0x2184:
	{
		if (sz<2) break;
		input->seek(pos+headerSize, librevenge::RVNG_SEEK_SET);
		Quattro9ParserInternal::TextEntry entry;
		if (readPString(stream, pos+headerSize+sz, entry))
			f << entry.getDebugString(stream) << ",";
		else
		{
			WPS_DEBUG_MSG(("Quattro9Parser::readZone[name]: can not read a string\n"));
			f << "###";
			break;
		}
		isParsed=needWriteInAscii=true;
		break;
	}

	default:
		break;
	}
	if (!ok)
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	input->seek(pos+headerSize+sz, librevenge::RVNG_SEEK_SET);
	if (isParsed)
	{
		if (needWriteInAscii)
		{
			ascFile.addPos(pos);
			ascFile.addNote(f.str().c_str());
		}
		return true;
	}
	if (sz)
		ascFile.addDelimiter(pos+headerSize,'|');
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

////////////////////////////////////////////////////////////
//   generic
////////////////////////////////////////////////////////////
bool Quattro9Parser::readPString(std::shared_ptr<WPSStream> const &stream, long endPos, Quattro9ParserInternal::TextEntry &entry)
{
	RVNGInputStreamPtr input = stream->m_input;
	long pos = input->tell();
	entry=Quattro9ParserInternal::TextEntry();
	if (pos+4>endPos || !stream->checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readPString: string's size seems bad\n"));
		return false;
	}
	int dSz=int(libwps::readU16(input));
	if (pos+4+dSz>endPos)
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readPString: string's size seems bad\n"));
		return false;
	}
	entry.m_flag=int(libwps::readU8(input));

	entry.m_entry.setBegin(input->tell());
	entry.m_entry.setLength(dSz);
	input->seek(pos+4+dSz, librevenge::RVNG_SEEK_SET);
	// entry.m_flag&0x20 frequent no data
	if (entry.m_flag&0x2)
	{
		if (!readTextStyles(stream, endPos, entry))
			return false;
	}
	if (entry.m_flag&0xdd)
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readPString: find some unknown flag, some data may be lost\n"));
		entry.m_extra="";
	}
	return true;
}

bool Quattro9Parser::readFontNames(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type != 6)
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readFontNames: not a font zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	int N=int(libwps::readU16(input));
	long endPos=pos+4+sz;
	for (int i=0; i<N; ++i)
	{
		pos=input->tell();
		f.str("");
		f << "Font-FN" << i << ":";
		Quattro9ParserInternal::TextEntry entry;
		librevenge::RVNGString font;
		if (readPString(stream, endPos, entry))
		{
			font=entry.getString(stream);
			f << entry.getDebugString(stream) << ",";
		}
		else
		{
			WPS_DEBUG_MSG(("Quattro9Parser::readFontNames: can not read a string\n"));
			f << "###";
			ascFile.addPos(pos);
			ascFile.addNote(f.str().c_str());
			return true;
		}
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		m_state->m_fontNamesList.push_back(font);
	}
	return true;
}

bool Quattro9Parser::readFontStyles(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type != 7)
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readFontStyles: not a font zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	int N=int(libwps::readU16(input));
	f << "N=" << N << ",";
	if (2+16*N!=sz)
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readFontStyles: the number of data seems bad\n"));
		f << "###";
		return false;
	}
	for (int i=0; i<N; ++i)
	{
		long pos=input->tell();
		f.str("");
		f << "Font-F" << i << ":";
		WPSFont font;
		int fId=int(libwps::readU16(input));
		if (fId<int(m_state->m_fontNamesList.size()))
			font.m_name=m_state->m_fontNamesList[size_t(fId)];
		else
			f << "###FN" << fId << ",";
		int val=int(libwps::readU16(input)); // system font id?
		f << "id" << val << ",";
		val=int(libwps::readU16(input)); // 0
		if (val) f << "f0=" << val << ",";
		auto fSize = int(libwps::readU16(input));
		if (fSize >= 1 && fSize <= 50) // check maximumn
			font.m_size=double(fSize);
		else
			f << "###fSize=" << fSize << ",";
		val=int(libwps::readU16(input)); // 0
		if (val) f << "f1=" << val << ",";

		uint32_t attributes = 0;
		val=int(libwps::readU16(input));
		if (val&1) attributes|=WPS_UNDERLINE_BIT;
		if (val&0x10) attributes|=WPS_UNDERLINE_BIT;
		if (val&0x20) attributes|=WPS_DOUBLE_UNDERLINE_BIT;
		val &= 0xffce;
		if (val) f << "f2=" << val << ",";
		val=int(libwps::readU16(input));
		if (val&1) attributes|=WPS_ITALICS_BIT;
		if (val&0x10) attributes|=WPS_STRIKEOUT_BIT;
		val &= 0xffee;
		if (val) f << "f3=" << val << ",";
		val=int(libwps::readU16(input));
		if (val==700) attributes|=WPS_BOLD_BIT;
		else if (val!=400) // 400=normal
			f << "f4=" << val << ",";
		font.m_attributes=attributes;
		m_state->m_fontsList.push_back(font);
		f << font;
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		input->seek(pos+16, librevenge::RVNG_SEEK_SET);
	}
	return true;
}

bool Quattro9Parser::readTextStyles(std::shared_ptr<WPSStream> const &stream, long endPos, Quattro9ParserInternal::TextEntry &entry)
{
	RVNGInputStreamPtr input = stream->m_input;
	long pos=input->tell();
	if (pos+2>endPos)
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readTextStyles: the zone is too short\n"));
		return false;
	}
	int dSz=int(libwps::readU16(input));
	if (dSz<6 || pos+dSz>endPos)
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readTextStyles: the zone size is bad\n"));
		return false;
	}
	endPos=pos+dSz;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	f << "Entries(TextStyle):";
	int nPos=int(libwps::readU16(input));
	f << "n[pos]=" << nPos << ",";
	int nFont=int(libwps::readU16(input));
	f << "n[font]=" << nFont << ",";
	if (dSz<6+4*nPos+42*nFont)
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readTextStyles: the number of position seems bad\n"));
		f << "###";
		nPos=nFont=0;
	}
	int actC=0;
	for (int i=0; i<nPos; ++i)
	{
		int nC=int(libwps::readU16(input));
		int id=int(libwps::readU16(input));
		entry.m_posFontIdMap[actC]=id;
		f << actC << ":Ft" << id << ",";
		actC+=nC;
	}
	entry.m_fontsList.resize(size_t(nFont));
	int dataSz=nFont>0 ? (dSz-6-4*nPos)/nFont : 42;
	for (auto &font : entry.m_fontsList)
	{
		if (readTextFontStyles(stream, dataSz, font))
			continue;
		break;
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	input->seek(endPos, librevenge::RVNG_SEEK_SET);
	return true;
}

bool Quattro9Parser::readTextFontStyles(std::shared_ptr<WPSStream> const &stream, int dataSz, WPSFont &font)
{
	font=WPSFont();
	RVNGInputStreamPtr input = stream->m_input;
	long pos=input->tell();
	if (dataSz<42 || !stream->checkFilePosition(pos+dataSz))
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readTextFontStyles: the zone is too short\n"));
		return false;
	}
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	f << "TextStyle[font]";
	auto fSize = int(libwps::readU16(input));
	if (fSize >= 1 && fSize <= 50) // check maximumn
		font.m_size=double(fSize);
	else
		f << "###fSize=" << fSize << ",";
	uint32_t attributes = 0;
	auto flags = int(libwps::readU16(input));
	if (flags & 1) attributes |= WPS_BOLD_BIT;
	if (flags & 2) attributes |= WPS_ITALICS_BIT;
	if (flags & 4) attributes |= WPS_UNDERLINE_BIT;
	if (flags & 8) attributes |= WPS_SUBSCRIPT_BIT; // reserved
	if (flags & 0x10) attributes |= WPS_SUPERSCRIPT_BIT; // reserved
	if (flags & 0x20) attributes |= WPS_STRIKEOUT_BIT;
	if (flags & 0x40) attributes |= WPS_DOUBLE_UNDERLINE_BIT; // reserved
	if (flags & 0x80) attributes |= WPS_OUTLINE_BIT; // reserved
	if (flags & 0x100) attributes |= WPS_SHADOW_BIT; // reserved
	font.m_attributes=attributes;
	if (flags&0xfe00)
		f << "##fl=" << std::hex << (flags&0xfe00) << std::dec << ",";

	auto const fontType = getDefaultFontType();
	std::string name;
	for (long i=0; i<32; ++i)
	{
		auto c = char(libwps::readU8(input));
		if (c == '\0') break;
		name.push_back(c);
	}
	if (!name.empty())
		font.m_name=libwps_tools_win::Font::unicodeString(name, fontType);
	input->seek(pos+36, librevenge::RVNG_SEEK_SET);
	unsigned char col[4];
	for (auto &c: col) c=static_cast<unsigned char>(libwps::readU8(input));
	font.m_color=WPSColor(col[0],col[1],col[2]);
	f << font;
	if (dataSz==42)
	{
		int val=int(libwps::readU16(input));
		if (val!=flags) f << "fl2=" << std::hex << val << std::dec << ",";
	}
	else
	{
		ascFile.addDelimiter(input->tell(),'|');
		input->seek(pos+dataSz, librevenge::RVNG_SEEK_SET);
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool Quattro9Parser::readDocumentFields(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if (type !=0x406)
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readDocumentFields: not a spreadsheet zone\n"));
		return false;
	}
	long sz = libwps::readU16(input);
	long endPos=pos+4+sz;
	long N=long(libwps::readU16(input));
	f << "N=" << N << ",";
	if (sz<2 || 2+N*28>sz || !stream->checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readDocumentFields: the size seems bad\n"));
		return false;
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	auto fontType=getDefaultFontType();
	for (long i=0; i<N; ++i)
	{
		pos=input->tell();
		if (pos+28>endPos) break;
		f.str("");
		f << "Document[fields]:Field" << i+1 << ",";
		int fSz=int(libwps::readU16(input));
		long endFieldPos=pos+fSz;
		if (fSz<28 || endFieldPos>endPos)
		{
			input->seek(pos, librevenge::RVNG_SEEK_SET);
			break;
		}
		bool hasRef=false;
		for (int j=0; j<3; ++j)
		{
			int val=int(libwps::readU16(input));
			if (!val) continue;
			if (j==1)
			{
				if (val&0x40)
				{
					f << "hasRef,";
					hasRef=true;
				}
				val &= 0xffbf;
			}
			if (val)
				f << "f" << j << "=" << val << ",";
		}
		Quattro9ParserInternal::TextEntry entry;
		bool ok=true;
		if (!readPString(stream, endPos, entry) || input->tell()+16>endFieldPos)
		{
			WPS_DEBUG_MSG(("Quattro9Parser::readDocumentFields: can not read the field name\n"));
			f << "###";
			ok=false;
		}
		else
			f << entry.getDebugString(stream) << ",";
		if (ok && hasRef)
		{
			QuattroFormulaInternal::CellReference instr;
			if (!m_spreadsheetParser->readCellReference(stream, endFieldPos, instr))
			{
				WPS_DEBUG_MSG(("Quattro9Parser::readDocumentFields: can not read some reference\n"));
				f << "###";
				ok=false;
			}
			else
			{
				f << instr;
				m_state->m_idToFieldMap[int(i+1)]=
				    std::pair<librevenge::RVNGString,QuattroFormulaInternal::CellReference>(entry.getString(stream, fontType), instr);
			}
		}
		if (ok && input->tell()+16<=endFieldPos)
		{
			for (int j=0; j<8; ++j)   // 0
			{
				int val=int(libwps::readU16(input));
				if (!val) continue;
				f << "g" << j << "=" << val << ",";
			}
		}
		if (input->tell()!=endFieldPos)
			ascFile.addDelimiter(input->tell(),'|');
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		input->seek(endFieldPos, librevenge::RVNG_SEEK_SET);
	}
	if (input->tell()!=endPos)
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readDocumentFields: find extra data\n"));
		ascFile.addPos(input->tell());
		ascFile.addNote("Document[fields]:###extra");
	}
	return true;
}

bool Quattro9Parser::readDocumentStrings(std::shared_ptr<WPSStream> const &stream)
{
	m_state->m_actualStrings.first=stream;
	auto &entries=m_state->m_actualStrings.second;
	entries.clear();
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::readU16(input));
	bool const bigBlock=type&0x8000;
	int const headerSize = bigBlock ? 6 : 4;
	if ((type&0x7fff) !=0x407)
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readDocumentStrings: not a spreadsheet zone\n"));
		return false;
	}
	long sz = bigBlock ? long(libwps::readU32(input)) : long(libwps::readU16(input));
	long endPos=pos+headerSize+sz;
	long N=long(libwps::readU32(input));
	f << "N=" << N << ",";
	if (sz<12 || N<0 || (sz-headerSize-8)/4<N || !stream->checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readDocumentStrings: the size seems bad\n"));
		return false;
	}
	f << "f0=" << libwps::readU32(input) << ",";
	f << "f1=" << libwps::readU32(input) << ",";
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	entries.reserve(size_t(N));
	for (long i=0; i<N; ++i)
	{
		pos=input->tell();
		librevenge::RVNGString text;
		f.str("");
		f << "Document[strings]:Str" << i+1 << ",";
		Quattro9ParserInternal::TextEntry entry;
		if (!readPString(stream, endPos, entry))
		{
			input->seek(pos, librevenge::RVNG_SEEK_SET);
			break;
		}
		entries.push_back(entry);
		f << entry.getDebugString(stream) << ",";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
	}
	if (input->tell()!=endPos)
	{
		if (int(entries.size())==N)
		{
			WPS_DEBUG_MSG(("Quattro9Parser::readDocumentStrings: find extra data\n"));
		}
		ascFile.addPos(input->tell());
		ascFile.addNote("Document[strings]:###extra");
	}
	return true;
}

bool Quattro9Parser::readFilePositions(std::shared_ptr<WPSStream> const &stream, long (&filePos)[2])
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	if (!stream->checkFilePosition(pos+12))
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readFilePositions: the zone is too short\n"));
		return false;
	}
	auto type = long(libwps::readU16(input)&0x7fff);
	if ((type&0xf) !=0x2)
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readFilePositions: unexpected type\n"));
		return false;
	}
	long sz = libwps::readU16(input);
	long endPos=pos+4+sz;
	if (sz<8 || !stream->checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readFilePositions: the size seems bad\n"));
		return false;
	}
	for (int i=0; i<2; ++i)
	{
		auto &lPos=filePos[i];
		lPos=long(libwps::readU32(input));
		if (lPos==0) continue;
		if (lPos<0 || !stream->checkFilePosition(lPos))
		{
			filePos[i]=0;
			WPS_DEBUG_MSG(("Quattro9Parser::readFilePositions: a position seems bad\n"));
			f << "###";
		}
		f << "pos" << i << "=" << std::hex << lPos << std::dec << ",";
	}
	if (sz!=8)
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readFilePositions: find extra data\n"));
		f << "###";
		ascFile.addDelimiter(input->tell(),'|');
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}
// ----------------------------------------------------------------------
// Header/Footer
// ----------------------------------------------------------------------
void Quattro9Parser::sendHeaderFooter(bool /*header*/)
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("Quattro9Parser::sendHeaderFooter: can not find the listener\n"));
		return;
	}
	WPS_DEBUG_MSG(("Quattro9Parser::sendHeaderFooter: not implemented\n"));
	m_listener->setFont(m_state->getDefaultFont());
}

////////////////////////////////////////////////////////////
//   ole stream
////////////////////////////////////////////////////////////
bool Quattro9Parser::readOleLinkInfo(std::shared_ptr<WPSStream> stream)
{
	if (!stream || !stream->checkFilePosition(4))
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readLinkInfo: unexpected zone\n"));
		return false;
	}
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	f << "Entries(LinkInfo):";
	int val=libwps::readU8(input);
	if (val!=0x53)
		f << "f0=" << std::hex << val << std::dec << ",";
	val=libwps::readU16(input); // 2 or 3
	if (val) f << "f1=" << val << ",";
	std::string name;
	while (!input->isEnd())
	{
		char c=char(libwps::readU8(input));
		if (!c) break;
		name+=c;
	}
	if (!name.empty())
		f << "name=" << name << ",";
	ascFile.addPos(0);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool Quattro9Parser::readOleBOlePart(std::shared_ptr<WPSStream> stream)
{
	if (!stream || !stream->checkFilePosition(20))
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readOleBOlePart: unexpected zone\n"));
		return false;
	}
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	f << "Entries(BOlePart):";
	for (int i=0; i<5; ++i)   // f0=1, f1=f2=small int(often 1), f3=f4=small int(often 1)
	{
		auto val=int(libwps::read32(input));
		if (val!=1) f << "f" << i << "=" << val << ",";
	}
	ascFile.addPos(0);
	ascFile.addNote(f.str().c_str());
	return true;
}

////////////////////////////////////////////////////////////
//   decode
////////////////////////////////////////////////////////////
RVNGInputStreamPtr Quattro9Parser::decodeStream(RVNGInputStreamPtr input, std::vector<uint8_t> const &key)
{
	//int const vers=version();
	if (!input || key.size()!=16)
	{
		WPS_DEBUG_MSG(("Quattro9Parser::decodeStream: the arguments seems bad\n"));
		return RVNGInputStreamPtr();
	}
	long actPos=input->tell();
	input->seek(0,librevenge::RVNG_SEEK_SET);
	librevenge::RVNGBinaryData data;
	if (!libwps::readDataToEnd(input, data) || !data.getDataBuffer())
	{
		WPS_DEBUG_MSG(("Quattro9Parser::decodeStream: can not read the original input\n"));
		return RVNGInputStreamPtr();
	}
	auto *buf=const_cast<unsigned char *>(data.getDataBuffer());
	auto endPos=long(data.size());
	input->seek(actPos,librevenge::RVNG_SEEK_SET);
	std::stack<long> stack;
	stack.push(endPos);
	while (!input->isEnd() && !stack.empty())
	{
		long pos=input->tell();
		if (pos+4>stack.top()) break;
		auto id=libwps::readU16(input);
		bool const bigBlock=id&0x8000;
		id &= 0x7fff;
		auto sSz=bigBlock ? long(libwps::readU32(input)) : long(libwps::readU16(input));
		int const headerSize=bigBlock ? 6 : 4;
		if (sSz<0 || pos+headerSize+sSz>stack.top())
		{
			input->seek(pos,librevenge::RVNG_SEEK_SET);
			break;
		}
		uint32_t d7=uint32_t(input->tell())&0xf;
		for (int i=0; i<sSz; ++i)
		{
			auto c=uint8_t(libwps::readU8(input));
			c=(c^key[(d7++)&0xf]);
			buf[pos+headerSize+i]=uint8_t((c>>5)|(c<<3));
		}
		// main zone ends with zone 2
		if (id==2)
		{
			input->seek(stack.top(),librevenge::RVNG_SEEK_SET);
			stack.pop();
		}
	}
	if (input->tell()!=endPos)
	{
		WPS_DEBUG_MSG(("Quattro9Parser::decodeStream: can not decode the end of the file, data may be bad %lx %lx\n", static_cast<unsigned long>(input->tell()), static_cast<unsigned long>(endPos)));
	}
	RVNGInputStreamPtr res(new WPSStringStream(data.getDataBuffer(), static_cast<unsigned int>(endPos)));
	res->seek(actPos, librevenge::RVNG_SEEK_SET);
	return res;
}

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
