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
#include <iterator>
#include <map>
#include <sstream>
#include <stack>
#include <utility>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WKSContentListener.h"

#include "WPSCell.h"
#include "WPSEntry.h"
#include "WPSFont.h"
#include "WPSHeader.h"
#include "WPSOLE1Parser.h"
#include "WPSPageSpan.h"
#include "WPSStream.h"
#include "WPSStringStream.h"
#include "WPSTable.h"

#include "LotusChart.h"
#include "LotusGraph.h"
#include "LotusSpreadsheet.h"
#include "LotusStyleManager.h"

#include "Lotus.h"

using namespace libwps;

//! Internal: namespace to define internal class of LotusParser
namespace LotusParserInternal
{
//! the font of a LotusParser
struct Font final : public WPSFont
{
	//! constructor
	explicit Font(libwps_tools_win::Font::Type type)
		: WPSFont()
		, m_type(type)
	{
	}
	Font(Font const &)=default;
	~Font() final;
	//! font encoding type
	libwps_tools_win::Font::Type m_type;
};
Font::~Font()
{
}
//! the state of LotusParser
struct State
{
	//! constructor
	State(libwps_tools_win::Font::Type fontType, char const *password)
		: m_fontType(fontType)
		, m_version(-1)
		, m_isMacFile(false)
		, m_inMainContentBlock(false)
		, m_fontsMap()
		, m_pageSpan()
		, m_maxSheet(0)
		, m_actualZoneId(0)
		, m_actualZoneParentId(0)
		, m_sheetZoneIdList()
		, m_dataZoneIdToSheetZoneIdMap()
		, m_linkIdToLinkMap()
		, m_actualLevels()
		, m_zone1Stack()
		, m_sheetSubZoneOpened(0x20, false)
		, m_actPage(0)
		, m_numPages(0)
		, m_metaData()
		, m_password(password)
		, m_isEncrypted(false)
		, m_isDecoded(false)
	{
	}
	//! return the default font style
	libwps_tools_win::Font::Type getDefaultFontType() const
	{
		if (m_fontType != libwps_tools_win::Font::UNKNOWN)
			return m_fontType;
		return libwps_tools_win::Font::WIN3_WEUROPE;
	}

	//! returns a default font (Courier12) with file's version to define the default encoding */
	WPSFont getDefaultFont() const
	{
		WPSFont res;
		if (m_version <= 2)
			res.m_name="Courier";
		else
			res.m_name="Times New Roman";
		res.m_size=12;
		return res;
	}
	//! returns the min and max cell
	void getLevels(WPSVec3i &minC, WPSVec3i &maxC) const
	{
		size_t numLevels=m_actualLevels.size();
		for (size_t i=0; i<3; ++i)
		{
			if (i+1<numLevels)
			{
				minC[int(i)]=m_actualLevels[i+1][0];
				maxC[int(i)]=m_actualLevels[i+1][1]-1;
			}
			else
				minC[int(i)]=maxC[int(i)]=-1;
		}
	}
	//! returns a map dataZoneId to sheet final id
	std::map<int,int> getDataZoneIdToSheetIdMap() const
	{
		std::map<int,int> zoneIdToSheetMap;
		for (size_t i=0; i<m_sheetZoneIdList.size(); ++i)
			zoneIdToSheetMap[m_sheetZoneIdList[i]]=int(i);
		std::map<int,int> res;
		for (auto it : m_dataZoneIdToSheetZoneIdMap)
		{
			if (zoneIdToSheetMap.find(it.second)==zoneIdToSheetMap.end())
			{
				WPS_DEBUG_MSG(("LotusParserInternal::State::getDataZoneIdToSheetIdMap: can not find the sheet corresponding to %d\n", it.second));
				continue;
			}
			res[zoneIdToSheetMap.find(it.second)->second]=it.first;
		}
		return res;
	}
	//! returns a name corresponding to the actual level(for debugging)
	std::string getLevelsDebugName() const
	{
		std::stringstream s;
		for (size_t i=0; i<m_actualLevels.size(); ++i)
		{
			auto const &level=m_actualLevels[i];
			if (i==0 && level==Vec2i(0,0)) continue;
			if (i<4)
			{
				char const *wh[]= {"Z", "T", "C", "R"};
				s << wh[i];
			}
			else
				s << "[F" << i << "]";
			if (level[0]==level[1])
				s << "_";
			else if (level[0]==level[1]-1)
				s << level[0];
			else
				s << level[0] << "x" << level[1]-1;
		}
		return s.str();
	}
	//! returns a name corresponding to the zone1 stack(for debugging)
	std::string getZone1StackDebugName() const
	{
		if (m_zone1Stack.empty())
			return "";
		std::stringstream s;
		s << "ids=[";
		for (auto const &id : m_zone1Stack)
			s << std::hex << id << std::dec << ",";
		s << "],";
		return s.str();
	}
	//! the user font type
	libwps_tools_win::Font::Type m_fontType;
	//! the file version
	int m_version;
	//! flag to know if this is a mac file
	bool m_isMacFile;
	//! a flag used to know if we are in the main block or no
	bool m_inMainContentBlock;
	//! the fonts map
	std::map<int, Font> m_fontsMap;
	//! the actual document size
	WPSPageSpan m_pageSpan;
	//! the last sheet number
	int m_maxSheet;
	//! the actual zone id
	int m_actualZoneId;
	//! the actual zone parent id
	int m_actualZoneParentId;
	//! the list of sheet main zone id
	std::vector<int> m_sheetZoneIdList;
	//! a map to retrieve the sheet zone id from the data sheet zone id
	std::map<int,int> m_dataZoneIdToSheetZoneIdMap;
	//! a multimap link id to link zone
	std::multimap<int, LotusParser::Link> m_linkIdToLinkMap;
	//! the actual zone: (0,0), table list, col list, row list
	std::vector<Vec2i> m_actualLevels;
	//! a unknown Zone1 stack of increasing(?) numbers
	std::vector<unsigned long> m_zone1Stack;
	//! some sheet sub zones (SheetZone)
	std::vector<bool> m_sheetSubZoneOpened;
	int m_actPage /** the actual page*/, m_numPages /* the number of pages */;
	//! the metadata
	librevenge::RVNGPropertyList m_metaData;

	//! the password (if known)
	char const *m_password;
	//! true if the file is encrypted
	bool m_isEncrypted;
	//! true if the main stream has been decoded
	bool m_isDecoded;
private:
	State(State const &) = delete;
	State operator=(State const &) = delete;
};

}

// constructor, destructor
LotusParser::LotusParser(RVNGInputStreamPtr &input, WPSHeaderPtr &header,
                         libwps_tools_win::Font::Type encoding,
                         char const *password)
	: WKSParser(input, header)
	, m_listener()
	, m_state(new LotusParserInternal::State(encoding, password))
	, m_styleManager()
	, m_chartParser()
	, m_graphParser()
	, m_spreadsheetParser()
	, m_ole1Parser()
{
	m_styleManager.reset(new LotusStyleManager(*this));
	m_chartParser.reset(new LotusChart(*this));
	m_graphParser.reset(new LotusGraph(*this));
	m_spreadsheetParser.reset(new LotusSpreadsheet(*this));
}

LotusParser::~LotusParser()
{
}

int LotusParser::version() const
{
	return m_state->m_version;
}

//////////////////////////////////////////////////////////////////////
// interface
//////////////////////////////////////////////////////////////////////
libwps_tools_win::Font::Type LotusParser::getDefaultFontType() const
{
	return m_state->getDefaultFontType();
}

bool LotusParser::getFont(int id, WPSFont &font, libwps_tools_win::Font::Type &type) const
{
	if (m_state->m_fontsMap.find(id)==m_state->m_fontsMap.end())
	{
		WPS_DEBUG_MSG(("LotusParser::getFont: can not find font %d\n", id));
		return false;
	}
	auto const &ft=m_state->m_fontsMap.find(id)->second;
	font=ft;
	type=ft.m_type;
	return true;
}

std::vector<LotusParser::Link> LotusParser::getLinksList(int lId) const
{
	std::vector<LotusParser::Link> res;
	auto range = m_state->m_linkIdToLinkMap.equal_range(lId);
	std::transform(range.first, range.second, std::back_inserter(res), [](std::pair<int,Link> const &element)
	{
		return element.second;
	});
	return res;

}

bool LotusParser::hasGraphics(int sheetId) const
{
	return m_graphParser->hasGraphics(sheetId);
}

void LotusParser::sendGraphics(int sheetId)
{
	m_graphParser->sendGraphics(sheetId);
}

bool LotusParser::getLeftTopPosition(Vec2i const &cell, int spreadsheet, Vec2f &pos) const
{
	return m_spreadsheetParser->getLeftTopPosition(cell, spreadsheet, pos);
}

librevenge::RVNGString LotusParser::getSheetName(int id) const
{
	return m_spreadsheetParser->getSheetName(id);
}

bool LotusParser::sendChart(int cId, WPSPosition const &pos, WPSGraphicStyle const &style)
{
	return m_chartParser->sendChart(cId, pos, style);
}

bool LotusParser::updateEmbeddedObject(int id, WPSEmbeddedObject &object) const
{
	if (!m_ole1Parser)
	{
		WPS_DEBUG_MSG(("LotusParser::updateEmbeddedObject: can not find the ole1 parser\n"));
		return false;
	}
	return m_ole1Parser->updateEmbeddedObject(id, object);
}

//////////////////////////////////////////////////////////////////////
// parsing
//////////////////////////////////////////////////////////////////////

// main function to parse the document
void LotusParser::parse(librevenge::RVNGSpreadsheetInterface *documentInterface)
{
	RVNGInputStreamPtr input=getInput();
	if (!input)
	{
		WPS_DEBUG_MSG(("LotusParser::parse: does not find main ole\n"));
		throw (libwps::ParseException());
	}

	if (!checkHeader(nullptr)) throw(libwps::ParseException());

	bool ok=false;
	try
	{
		ascii().setStream(input);
		ascii().open("MN0");
		if (checkHeader(nullptr) && createZones())
			createListener(documentInterface);
		if (m_listener)
		{
			m_styleManager->updateState();
			m_chartParser->updateState();
			m_spreadsheetParser->updateState();
			m_graphParser->updateState(m_state->getDataZoneIdToSheetIdMap(),
			                           m_chartParser->getNameToChartIdMap());

			m_chartParser->setListener(m_listener);
			m_graphParser->setListener(m_listener);
			m_spreadsheetParser->setListener(m_listener);

			m_listener->startDocument();
			for (int i=0; i<=m_state->m_maxSheet; ++i)
				m_spreadsheetParser->sendSpreadsheet(i);
			if (version()<=1 && !m_state->m_isMacFile && m_chartParser->getNumCharts())
			{
				std::vector<WPSColumnFormat> widths;
				WPSColumnFormat width(72.);
				width.m_numRepeat=20;
				widths.push_back(width);
				m_listener->openSheet(widths, "Charts");
				m_chartParser->sendCharts();
				m_listener->closeSheet();
			}
			m_listener->endDocument();
			m_listener.reset();
			ok = true;
		}
	}
	catch (libwps::PasswordException())
	{
		ascii().reset();
		WPS_DEBUG_MSG(("LotusParser::parse: password exception catched when parsing MN0\n"));
		throw (libwps::PasswordException());
	}
	catch (...)
	{
		WPS_DEBUG_MSG(("LotusParser::parse: exception catched when parsing MN0\n"));
		throw (libwps::ParseException());
	}

	ascii().reset();
	if (!ok)
		throw(libwps::ParseException());
}

bool LotusParser::createListener(librevenge::RVNGSpreadsheetInterface *interface)
{
	std::vector<WPSPageSpan> pageList;
	WPSPageSpan ps(m_state->m_pageSpan);
	int numPages=m_state->m_maxSheet+1;
	if (numPages<=0) numPages=1;
	for (int i=0; i<numPages; ++i) pageList.push_back(ps);
	m_listener.reset(new WKSContentListener(pageList, interface));
	m_listener->setMetaData(m_state->m_metaData);
	return true;
}

////////////////////////////////////////////////////////////
// low level
////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////
// read the header
////////////////////////////////////////////////////////////
bool LotusParser::checkHeader(WPSHeader *header, bool strict)
{
	m_state.reset(new LotusParserInternal::State(m_state->m_fontType, m_state->m_password));
	std::shared_ptr<WPSStream> mainStream(new WPSStream(getInput(), ascii()));
	if (!checkHeader(mainStream, true, strict))
		return false;
	if (header)
	{
		header->setMajorVersion(uint8_t(100+m_state->m_version));
		header->setCreator(libwps::WPS_LOTUS);
		header->setKind(libwps::WPS_SPREADSHEET);
		header->setNeedEncoding(true);
		header->setIsEncrypted(m_state->m_isEncrypted);
	}
	return true;
}

bool LotusParser::checkHeader(std::shared_ptr<WPSStream> stream, bool mainStream, bool strict)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	if (!stream->checkFilePosition(12))
	{
		WPS_DEBUG_MSG(("LotusParser::checkHeader: file is too short\n"));
		return false;
	}

	input->seek(0,librevenge::RVNG_SEEK_SET);
	auto firstOffset = int(libwps::readU8(input));
	auto type = int(libwps::read8(input));
	auto val=int(libwps::read16(input));
	f << "FileHeader:";
	if (firstOffset == 0 && type == 0 && val==0x1a)
	{
		m_state->m_version=1;
		f << "DOS,";
	}
	else
	{
		WPS_DEBUG_MSG(("LotusParser::checkHeader: find unexpected first data\n"));
		return false;
	}
	val=int(libwps::readU16(input));
	if (!mainStream)
	{
		if (val!=0x8007)
		{
			WPS_DEBUG_MSG(("LotusParser::checkHeader: find unknown lotus file format\n"));
			return false;
		}
		f << "lotus123[FMT],";
	}
	else if (val>=0x1000 && val<=0x1005)
	{
		WPS_DEBUG_MSG(("LotusParser::checkHeader: find lotus123 file\n"));
		m_state->m_version=(val-0x1000)+1;
		f << "lotus123[" << m_state->m_version << "],";
	}
#ifdef DEBUG
	else if (val==0x8007)
	{
		WPS_DEBUG_MSG(("LotusParser::checkHeader: find lotus file format, sorry parsing this file is only implemented for debugging, not output will be created\n"));
		f << "lotus123[FMT],";
	}
#endif
	else
	{
		WPS_DEBUG_MSG(("LotusParser::checkHeader: unknown lotus 123 header\n"));
		return false;
	}

	input->seek(0, librevenge::RVNG_SEEK_SET);
	if (strict)
	{
		for (int i=0; i < 4; ++i)
		{
			if (!readZone(stream)) return false;
			if (m_state->m_isEncrypted) break;
		}
	}
	ascFile.addPos(0);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusParser::createZones()
{
	RVNGInputStreamPtr input=getInput();
	if (!input)
	{
		WPS_DEBUG_MSG(("LotusParser::createZones: can not find the main input\n"));
		return false;
	}
	m_styleManager->cleanState();
	m_chartParser->cleanState();
	m_graphParser->cleanState();
	m_spreadsheetParser->cleanState();

	int const vers=version();

	std::shared_ptr<WPSStream> mainStream(new WPSStream(input, ascii()));
	if (vers>=3)
	{
		m_ole1Parser.reset(new WPSOLE1Parser(mainStream));
		m_ole1Parser->createZones();
		std::shared_ptr<WPSStream> wkStream=m_ole1Parser->getStreamForName(vers==3 ? "WK3" : "123");
		if (wkStream)
		{
			if (!readZones(wkStream)) return false;
			m_ole1Parser->updateMetaData(m_state->m_metaData, getDefaultFontType());
			if (vers==3)
			{
				std::shared_ptr<WPSStream> fmStream=m_ole1Parser->getStreamForName("FM3");
				if (fmStream) readZones(fmStream);
			}
			return true;
		}
	}
	input->seek(0, librevenge::RVNG_SEEK_SET);
	if (!readZones(mainStream)) return false;
	if (vers<=2) parseFormatStream();
	return true;
}

bool LotusParser::parseFormatStream()
{
	RVNGInputStreamPtr file=getFileInput();
	if (!file || !file->isStructured()) return false;

	RVNGInputStreamPtr formatInput(file->getSubStreamByName("FM3"));
	if (!formatInput)
	{
		WPS_DEBUG_MSG(("LotusParser::parseFormatStream: can not find the format stream\n"));
		return false;
	}

	std::shared_ptr<WPSStream> formatStream(new WPSStream(formatInput));
	formatInput->seek(0, librevenge::RVNG_SEEK_SET);
	formatStream->m_ascii.open("FM3");
	formatStream->m_ascii.setStream(formatInput);
	if (!checkHeader(formatStream, false, false))
	{
		WPS_DEBUG_MSG(("LotusParser::parseFormatStream: can not read format stream\n"));
		return false;
	}
	return readZones(formatStream);
}

bool LotusParser::readZones(std::shared_ptr<WPSStream> stream)
{
	if (!stream)
	{
		WPS_DEBUG_MSG(("LotusParser::readZones: can not find the stream\n"));
		return false;
	}
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;

	bool mainDataRead=false;
	// data, format and ?
	for (int wh=0; wh<2; ++wh)
	{
		if (input->isEnd())
			break;

		while (readZone(stream))
		{
			if (m_state->m_isEncrypted && !m_state->m_isDecoded)
				throw(libwps::PasswordException());
		}

		//
		// look for ending
		//
		long pos = input->tell();
		if (!stream->checkFilePosition(pos+4))
			break;
		auto type = int(libwps::readU16(input)); // 1
		auto length = int(libwps::readU16(input));
		if (type==1 && length==0)
		{
			ascFile.addPos(pos);
			ascFile.addNote("Entries(EOF)");
			if (!mainDataRead)
				mainDataRead=m_state->m_inMainContentBlock;
			// end of block, look for other blocks
			continue;
		}
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		break;
	}

	while (!input->isEnd())
	{
		long pos=input->tell();
		if (pos>=stream->m_eof) break;
		auto id = int(libwps::readU8(input));
		auto type = int(libwps::readU8(input));
		auto sz = long(libwps::readU16(input));
		if ((type>0x2a) || sz<0 || !stream->checkFilePosition(pos+4+sz))
		{
			input->seek(pos, librevenge::RVNG_SEEK_SET);
			break;
		}
		libwps::DebugStream f;
		f << "Entries(UnknZon" << std::hex << id << "):";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		input->seek(pos+4+sz, librevenge::RVNG_SEEK_SET);
	}

	if (!input->isEnd() && input->tell()<stream->m_eof)
	{
		ascFile.addPos(input->tell());
		ascFile.addNote("Entries(Unknown)");
	}

	return mainDataRead || m_spreadsheetParser->hasSomeSpreadsheetData();
}

bool LotusParser::readZone(std::shared_ptr<WPSStream> &stream)
{
	if (!stream)
		return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto id = int(libwps::readU8(input));
	auto type = int(libwps::readU8(input));
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	if ((type>0x2a) || sz<0 || !stream->checkFilePosition(endPos))
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	f << "Entries(Lotus";
	if (type) f << std::hex << type << std::dec << "A";
	f << std::hex << id << std::dec << "E):";
	bool ok = true, isParsed = false, needWriteInAscii = false;
	int val;
	input->seek(pos, librevenge::RVNG_SEEK_SET);
	int const vers=version();
	switch (type)
	{
	case 0:
		switch (id)
		{
		case 0:
		{
			if (sz!=26)
			{
				ok=false;
				break;
			}
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			f.str("");
			f << "Entries(BOF):";
			val=int(libwps::readU16(input));
			m_state->m_inMainContentBlock=false;
			if (val==0x8007)
				f << "FMT,";
			else if (val>=0x1000 && val <= 0x1005)
			{
				m_state->m_inMainContentBlock=true;
				f << "version=" << (val-0x1000) << ",";
			}
			else
				f << "#version=" << std::hex << val << std::dec << ",";
			for (int i=0; i<4; ++i)   // f0=4, f3 a small number
			{
				val=int(libwps::read16(input));
				if (val)
					f << "f" << i << "=" << val << ",";
			}
			val=int(libwps::readU8(input));
			if (m_state->m_inMainContentBlock)
			{
				m_spreadsheetParser->setLastSpreadsheetId(val);
				m_state->m_maxSheet=val;
			}
			if (val && m_state->m_inMainContentBlock)
				f << "max[sheet]=" << val << ",";
			else if (val)
				f << "max[fmt]=" << val << ",";

			for (int i=0; i<7; ++i)   // g0/g1=0..fd, g2=0|4, g3=0|5|7|1e|20|30, g4=0|8c|3d, g5=1|10, g6=2|a
			{
				val=int(libwps::readU8(input));
				if (val)
					f << "g" << i << "=" << std::hex << val << std::dec << ",";
			}
			isParsed=needWriteInAscii=true;
			break;
		}
		case 0x1: // EOF
			ok = false;
			break;
		case 0x2:
			m_state->m_isEncrypted=true;
			if (sz==16)
			{
				input->seek(pos+4, librevenge::RVNG_SEEK_SET);
				std::vector<uint8_t> fileKeys;
				for (int i=0; i<16; ++i)
				{
					unsigned char c(libwps::readU8(input));
					fileKeys.push_back(uint8_t(c));
				}
				isParsed=needWriteInAscii=true;
				if (!m_state->m_isDecoded)
				{
					static uint8_t const defValues[]=
					{
						0xb9,0x5f, 0xd7,0x31, 0xdb,0x75, 9,0x72,
						0x5d,0x85, 0x32,0x11, 0x5,0x11, 0x58,0
					};
					uint16_t key;
					std::vector<uint8_t> keys;
					if (m_state->m_password && libwps::encodeLotusPassword(m_state->m_password, key, keys, defValues))
					{
						bool passwordOk=fileKeys.size()==keys.size();
						if (passwordOk)
						{
							/* check that the password is ok, normally,
							   all keys must be equal excepted:
							   - fileKey7=key7^(key>>8),
							   - and fileKey13=key13^key.

							   This also means that knowing fileKeys,
							   it is possible to retrieve the password
							   if it is short enough.
							*/
							int numSame=0;
							for (size_t c=0; c < fileKeys.size(); ++c)
							{
								if (keys[c]==fileKeys[c])
									++numSame;
							}
							passwordOk=numSame>=14;
							if (!passwordOk)
							{
								WPS_DEBUG_MSG(("LotusParser::parse: the password seems bad\n"));
							}
						}
						if (!passwordOk)
						{
							keys=retrievePasswordKeys(fileKeys);
							passwordOk=keys.size()==16;
						}
						RVNGInputStreamPtr newInput;
						if (passwordOk) newInput=decodeStream(input, stream->m_eof, keys);
						if (newInput)
						{
							// let's replace the current input by the decoded input
							m_state->m_isDecoded=true;
							stream->m_input=newInput;
							stream->m_ascii.setStream(newInput);
						}
					}
				}
			}
			else
			{
				WPS_DEBUG_MSG(("LotusParser::parse: find unexpected password field\n"));
				throw (libwps::PasswordException());
			}
			f.str("");
			f << "Entries(Password):";
			break;
		case 0x3:
			if (sz!=6)
			{
				ok=false;
				break;
			}
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			for (int i=0; i<3; ++i)   // f0=1, f2=1|32
			{
				val=int(libwps::read16(input));
				if (val)
					f << "f" << i << "=" << val << ",";
			}
			isParsed=needWriteInAscii=true;
			break;
		case 0x4:
			if (sz!=28)
			{
				ok=false;
				break;
			}
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			for (int i=0; i<2; ++i)   // f0=1|3
			{
				val=int(libwps::read8(input));
				if (val!=1)
					f << "f" << i << "=" << val << ",";
			}
			for (int i=0; i<2; ++i)   // f2=1-3, f1=0|1
			{
				val=int(libwps::read16(input));
				if (val)
					f << "f" << i+1 << "=" << val << ",";
			}
			isParsed=needWriteInAscii=true;
			break;
		case 0x5:
		{
			f.str("");
			f << "Entries(SheetUnknA):";
			if (sz!=16)
			{
				ok=false;
				break;
			}
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			val=int(libwps::readU8(input));
			if (val)
				f << "sheet[id]=" << val << ",";
			val=int(libwps::read8(input)); // always 0?
			if (val)
				f << "f0=" << val << ",";

			isParsed=needWriteInAscii=true;
			break;
		}
		case 0x6: // one by sheet
			f.str("");
			f << "Entries(SheetUnknB):";
			if (sz!=5)
			{
				ok=false;
				break;
			}
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			val=int(libwps::readU8(input));
			if (val)
				f << "sheet[id]=" << val << ",";
			for (int i=0; i<4; ++i)   // f0=0, f2=0|1, f3=7-9
			{
				val=int(libwps::read8(input)); // always 0?
				if (val)
					f << "f" << i << "=" << val << ",";
			}
			isParsed=needWriteInAscii=true;
			break;
		case 0x7:
			ok=isParsed=m_spreadsheetParser->readColumnSizes(stream);
			break;
		case 0x9:
			ok=isParsed=m_spreadsheetParser->readCellName(stream);
			break;
		case 0xa:
			ok=isParsed=readLinkZone(stream);
			break;
		case 0xb: // 0,1,-1
		case 0x1e: // always with 0
		case 0x21:
			if (sz!=1)
			{
				ok=false;
				break;
			}
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			val=int(libwps::read8(input));
			if (val==1)
				f << "true,";
			else if (val)
				f << "val=" << val << ",";
			break;
		case 0xc: // find 0 or 4 int with value 0|1|ff
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			for (long i=0; i<sz; ++i)
			{
				val=int(libwps::read8(input));
				if (val==1) f << "f" << i << ",";
				else if (val) f << "f" << i << "=" << val << ",";
			}
			isParsed=needWriteInAscii=true;
			break;
		case 0xe:
			if (sz<30)
			{
				ok=false;
				break;
			}
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			for (int i=0; i<30; ++i)   // f7=0|f, f8=0|60, f9=0|54, f17=80, f18=0|ff, f19=3f|40, f26=0|f8, f27=80|ff, f28=b|c,f29=40
			{
				val=int(libwps::read8(input));
				if (val) f << "f" << i << "=" << val << ",";
			}
			if (sz>=32)
			{
				val=int(libwps::read16(input)); // always 1?
				if (val!=1) f << "f30=" << val << ",";
			}
			isParsed=needWriteInAscii=true;
			break;
		case 0xf:
			if (sz<0x56)
			{
				ok=false;
				break;
			}
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			val=int(libwps::read8(input)); // 1|2
			if (val!=1) f << "f0=" << val << ",";
			for (int i=0; i<3; ++i)
			{
				long actPos=input->tell();
				std::string name("");
				for (int j=0; j<16; ++j)
				{
					auto c=char(libwps::readU8(input));
					if (!c) break;
					name += c;
				}
				if (!name.empty())
					f << "str" << i << "=" << name << ",";
				input->seek(actPos+16, librevenge::RVNG_SEEK_SET);
			}
			for (int i=0; i<17; ++i)   // f2=f11=1,f15=0|1, f16=0|2, f17=0|1|2
			{
				val=int(libwps::read8(input));
				if (val) f << "f" << i+1 << "=" << val << ",";
			}
			for (int i=0; i<10; ++i)   // g0=0|1,g1=Ø|1, g2=4|a, g3=4c|50|80, g4=g5=0|2, g6=42, g7=41|4c, g8=3c|42|59
			{
				val=int(libwps::read16(input));
				if (val) f << "g" << i << "=" << val << ",";
			}
			isParsed=needWriteInAscii=true;
			break;
		case 0x10: // CHECKME
		{
			if (sz<3)
			{
				ok=false;
				break;
			}
			f.str("");
			f << "Entries(Macro):";
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			for (int i=0; i<2; ++i)
			{
				val=int(libwps::readU8(input));
				if (val) f << "f" << i << "=" << val << ",";
			}
			std::string data("");
			for (long i=2; i<sz; ++i)
			{
				auto c=char(libwps::readU8(input));
				if (!c) break;
				data += c;
			}
			if (!data.empty())
				f << "data=" << data << ",";
			if (input->tell()!=endPos && input->tell()+1!=endPos)
			{
				WPS_DEBUG_MSG(("LotusParser::readZone: the string zone %d seems too short\n", id));
				f << "###";
			}
			isParsed=needWriteInAscii=true;
			break;
		}
		case 0x11:
			ok=isParsed=m_chartParser->readChart(stream);
			break;
		case 0x12:
			ok=isParsed=m_chartParser->readChartName(stream);
			break;
		case 0x13:
			isParsed=m_spreadsheetParser->readRowFormats(stream);
			break;
		case 0x15:
		case 0x1d:
			if (sz!=4)
			{
				WPS_DEBUG_MSG(("LotusParser::readZone: size of zone%d seems bad\n", id));
				f << "###";
				break;
			}
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			val=int(libwps::read16(input)); // small number 6-c maybe a style
			if (val) f << "f0=" << val << ",";
			for (int i=0; i<2; ++i)   // zone15: f1=3, f2=2-5, zone 1d: always 0
			{
				val=int(libwps::readU8(input));
				if (val) f << "f" << i+1 << "=" << val << ",";
			}
			isParsed=needWriteInAscii=true;
			break;
		case 0x16: // the cell text
		case 0x17: // double10 cell
		case 0x18: // uint16 double cell
		case 0x19: // double10+formula
		case 0x1a: // text formula result cell
		case 0x25: // uint32 double cell
		case 0x26: // comment cell
		case 0x27: // double8 cell
		case 0x28: // double8+formula
			ok=isParsed=m_spreadsheetParser->readCell(stream);
			break;
		case 0x1b:
			isParsed=readDataZone(stream);
			break;
		case 0x1c: // always 00002d000000
			if (sz!=6)
			{
				WPS_DEBUG_MSG(("LotusParser::readZone: size of zone%d seems bad\n", id));
				f << "###";
				break;
			}
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			for (int i=0; i<6; ++i)   // some int
			{
				val=int(libwps::readU8(input));
				if (val) f << "f" << i << "=" << std::hex << val << std::dec << ",";
			}
			isParsed=needWriteInAscii=true;
			break;
		case 0x1f:
			isParsed=ok=m_spreadsheetParser->readColumnDefinition(stream);
			break;
		case 0x23:
			isParsed=ok=m_spreadsheetParser->readSheetName(stream);
			break;
		// case 13: big structure

		//
		// format:
		//
		case 0x93: // 4
		case 0x96: // 0 or FF
		case 0x97: // 5F
		case 0x98: // 0|2|3
		case 0x99: // 0|4 or FF
		case 0x9c: // 0
		case 0xa3: // 0 or FF
		case 0xce: // 1
		case 0xcf: // 1
		case 0xd0: // 1
			if (m_state->m_inMainContentBlock) break;
			f.str("");
			f << "Entries(FMTByte" << std::hex << id << std::dec << "Z):";
			if (sz!=1)
			{
				f << "###";
				break;
			}
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			val=int(libwps::readU8(input));
			if (val==0xFF) f << "true,";
			else if (val) f << "val=" << val << ",";
			isParsed=needWriteInAscii=true;
			break;
		case 0x87: // always with 0000
		case 0x88: // always with 0000
		case 0x8e: // with 57|64
		case 0x9a: // with 800
		case 0x9b: // with 720
		case 0xcd: // with 57|64
			if (m_state->m_inMainContentBlock) break;
			f.str("");
			f << "Entries(FMTInt" << std::hex << id << std::dec << "Z):";
			if (sz!=2)
			{
				f << "###";
				break;
			}
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			val=int(libwps::readU16(input));
			if (val) f << "val=" << val << ",";
			isParsed=needWriteInAscii=true;
			break;
		case 0x86:
		case 0x89:
		case 0xba: // header?
		case 0xbb: // footer?
		{
			if (m_state->m_inMainContentBlock) break;
			f.str("");
			if (id==0x86)
				f << "Entries(FMTPrinter):";
			else if (id==0x89)
				f << "Entries(FMTPrinter):shortName,";
			else if (id==0xba)
				f << "Entries(FMTHeader):";
			else
				f << "Entries(FMTFooter):";

			if (sz<1)
			{
				f << "###";
				break;
			}
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			std::string text;
			for (long i=0; i<sz; ++i) text+=char(libwps::readU8(input));
			f << text << ",";
			isParsed=needWriteInAscii=true;
			break;
		}
		case 0xae:
			if (m_state->m_inMainContentBlock) break;
			isParsed=m_styleManager->readFMTFontName(stream);
			break;
		case 0xaf:
		case 0xb1:
			if (m_state->m_inMainContentBlock) break;
			isParsed=m_styleManager->readFMTFontSize(stream);
			break;
		case 0xb0:
			if (m_state->m_inMainContentBlock) break;
			isParsed=m_styleManager->readFMTFontId(stream);
			break;
		case 0xb6:
			if (m_state->m_inMainContentBlock) break;
			isParsed=readFMTStyleName(stream);
			break;
		case 0xb8: // always 0101
			if (m_state->m_inMainContentBlock) break;
			f.str("");
			f << "Entries(FMTInts" << std::hex << id << std::dec << "Z):";
			if (sz!=2)
			{
				f << "###";
				break;
			}
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			for (int i=0; i<2; ++i)
			{
				val=int(libwps::readU8(input));
				if (val!=1) f << "f" << i << "=" << val << ",";
			}
			isParsed=needWriteInAscii=true;
			break;
		case 0xc3:
			if (m_state->m_inMainContentBlock) break;
			isParsed=ok=m_spreadsheetParser->readSheetHeader(stream);
			break;
		case 0xc4: // with 0-8, 5c-15c
		case 0xcb: // unsure, seems appeared together a group with 1,1
			if (m_state->m_inMainContentBlock) break;
			f.str("");
			if (id==0xcb)
				f << "Entries(FMTGrpData):";
			else
				f << "Entries(FMTInt2" << std::hex << id << std::dec << "Z):";
			if (sz!=4)
			{
				f << "###";
				break;
			}
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			for (int i=0; i<2; ++i)
			{
				val=int(libwps::readU16(input));
				if (val) f << "f" << i << "=" << val << ",";
			}
			isParsed=needWriteInAscii=true;
			break;
		case 0xc5:
			if (m_state->m_inMainContentBlock) break;
			isParsed=ok=m_spreadsheetParser->readExtraRowFormats(stream);
			break;
		case 0xc9:
			if (m_state->m_inMainContentBlock) break;
			isParsed=ok=m_graphParser->readZoneBeginC9(stream);
			break;
		case 0xca: // a graphic
			if (m_state->m_inMainContentBlock) break;
			isParsed=ok=m_graphParser->readGraphic(stream);
			break;
		case 0xcc: // frame of a graphic
			if (m_state->m_inMainContentBlock) break;
			isParsed=ok=m_graphParser->readFrame(stream);
			break;
		case 0xd1: // the textbox data
			if (m_state->m_inMainContentBlock) break;
			isParsed=ok=m_graphParser->readTextBoxDataD1(stream);
			break;
		case 0xb7:
			if (m_state->m_inMainContentBlock) break;
			isParsed=ok=m_graphParser->readFMTPictName(stream);
			break;
		case 0xbf: // variable size, can also contain a name, ...
		case 0xc0: // size 1a: dim + a list of id?
		case 0xc2: // size 22: a list of id?
			if (m_state->m_inMainContentBlock) break;
			f.str("");
			f << "Entries(FMTPict" << std::hex << id << std::dec << "):";
			break;
		default:
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			if (!m_state->m_inMainContentBlock && id>=0x80)
			{
				f.str("");
				f << "Entries(FMT" << std::hex << id << std::dec << "E):";
			}
			break;
		}
		break;
	case 1:
		if (vers<=2)
		{
			ok=false;
			break;
		}
		ok = isParsed=readZone1(stream);
		break;
	case 2:
		if (vers<=2)
		{
			ok=false;
			break;
		}
		ok = isParsed=readSheetZone(stream);
		break;
	case 3:
		if (vers<=2)
		{
			ok=false;
			break;
		}
		ok = isParsed=m_graphParser->readGraphZone(stream, m_state->m_actualZoneParentId); // sheetZone.Data0
		break;
	case 4:
		if (vers<=2)
		{
			ok=false;
			break;
		}
		ok = isParsed=readZone4(stream);
		break;
	case 5:
		if (vers<=2)
		{
			ok=false;
			break;
		}
		ok = isParsed=readChartZone(stream);
		break;
	case 6:
		if (vers<=2)
		{
			ok=false;
			break;
		}
		ok = isParsed=readRefZone(stream);
		break;
	case 7:
		if (vers<=2)
		{
			ok=false;
			break;
		}
		ok = isParsed=readZone7(stream);
		break;
	case 8:
		if (vers<=2)
		{
			ok=false;
			break;
		}
		ok = isParsed=readZone8(stream);
		break;
	case 0xa:
		if (vers<=2)
		{
			ok=false;
			break;
		}
		ok = isParsed=readVersionZone(stream);
		break;
	default:
		// checkme: maybe <5 is ok
		if (vers<=2)
		{
			ok=false;
			break;
		}
		ok = isParsed=readZoneV3(stream);
		break;
	}
	if (!ok)
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	if (sz && input->tell()!=pos && input->tell()!=endPos)
		ascFile.addDelimiter(input->tell(),'|');
	input->seek(endPos, librevenge::RVNG_SEEK_SET);
	if (!isParsed || needWriteInAscii)
	{
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
	}
	return true;
}

bool LotusParser::readDataZone(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input));
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	if (type!=0x1b || sz<2)
	{
		WPS_DEBUG_MSG(("LotusParser::readDataZone: the zone seems odd\n"));
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	type = int(libwps::readU16(input));
	f << "Entries(Data" << std::hex << type << std::dec << "E):";
	bool isParsed=false, needWriteInAscii=false;
	sz-=2;
	int val;
	switch (type)
	{
	//
	// mac windows
	//
	case 0x7d2:
	{
		f.str("");
		f << "Entries(WindowsMacDef):";
		if (sz<26)
		{
			WPS_DEBUG_MSG(("LotusParser::readDataZone: the windows definition seems bad\n"));
			f << "###";
			break;
		}
		val=int(libwps::readU8(input));
		if (val) f << "id=" << val << ",";
		val=int(libwps::read8(input)); // find 0|2
		if (val) f << "f0=" << val << ",";
		int dim[4];
		for (int i=0; i<4; ++i)
		{
			dim[i]=int(libwps::read16(input));
			val=int(libwps::read16(input));
			if (!val) continue;
			if (i)
				f << "num[split]=" << val << ",";
			else
				f << "dim" << i << "[h]=" << val << ",";
		}
		f << "dim=" << WPSBox2i(Vec2i(dim[0],dim[1]),Vec2i(dim[2],dim[3])) << ",";
		for (int i=0; i<8; ++i)   // small value or 100
		{
			val=int(libwps::read8(input));
			if (val) f << "f" << i+1 << "=" << val << ",";
		}
		isParsed=needWriteInAscii=true;
		auto remain=int(sz-26);
		if (remain<=1) break;
		std::string name("");
		for (int i=0; i<remain; ++i)
			name+=char(libwps::readU8(input));
		f << name << ",";
		break;
	}
	case 0x7d3:
	{
		f.str("");
		f << "Entries(WindowsMacSplit):";
		if (sz<24)
		{
			WPS_DEBUG_MSG(("LotusParser::readDataZone: the windows split seems bad\n"));
			f << "###";
			break;
		}
		val=int(libwps::readU8(input));
		if (val) f << "id=" << val << ",";
		val=int(libwps::readU8(input));
		if (val) f << "split[id]=" << val << ",";
		for (int i=0; i<3; ++i)   // 0 or 1
		{
			val=int(libwps::read8(input));
			if (val) f << "f" << i+1 << "=" << val << ",";
		}
		int dim[4];
		for (int i=0; i<4; ++i)
		{
			val=int(libwps::read16(input));
			dim[i]=int(libwps::read16(input));
			if (val) f << "dim" << i <<"[h]=" << val << ",";
		}
		f << "dim=" << WPSBox2i(Vec2i(dim[0],dim[1]),Vec2i(dim[2],dim[3])) << ",";
		for (int i=0; i<3; ++i)
		{
			static int const expected[]= {0,-1,25};
			val=int(libwps::read8(input));
			if (val!=expected[i]) f << "g" << i << "=" << val << ",";
		}
		isParsed=needWriteInAscii=true;
		break;
	}
	case 0x7d4:
	{
		f.str("");
		f << "Entries(WindowsMacUnkn0)";
		if (sz<5)
		{
			WPS_DEBUG_MSG(("LotusParser::readDataZone: the windows unkn0 seems bad\n"));
			f << "###";
			break;
		}
		for (int i=0; i<4; ++i)   // always 2,1,1,2 ?
		{
			val=int(libwps::read8(input));
			if (val) f << "f" << i << "=" << val << ",";
		}
		isParsed=needWriteInAscii=true;
		auto remain=int(sz-4);
		if (remain<=1) break;
		std::string name("");
		for (int i=0; i<remain; ++i) // always LMBCS 1.2?
			name+=char(libwps::readU8(input));
		f << name << ",";
		break;
	}
	case 0x7d5: // frequently followed by Lotus13 block and SheetRow, ...
		f.str("");
		f << "Entries(SheetBegin):";
		if (sz!=11)
		{
			WPS_DEBUG_MSG(("LotusParser::readDataZone: the sheet begin zone seems bad\n"));
			f << "###";
			break;
		}
		// time to update the style manager state
		m_styleManager->updateState();

		val=int(libwps::readU8(input));
		if (val) f << "sheet[id]=" << val << ",";
		// then always 0a3fff00ffff508451ff ?
		isParsed=needWriteInAscii=true;
		break;
	case 0x7d7:
		isParsed=m_spreadsheetParser->readRowSizes(stream, endPos);
		break;
	case 0x7d8:
	case 0x7d9:
	{
		f.str("");
		int dataSz=type==0x7d8 ? 1 : 2;
		if (type==0x7d8)
			f << "Entries(ColMacBreak):";
		else
			f << "Entries(RowMacBreak):";

		if (sz<4 || (sz%dataSz))
		{
			WPS_DEBUG_MSG(("LotusParser::readDataZone: the page mac break seems bad\n"));
			f << "###";
			break;
		}
		val=int(libwps::readU8(input));
		if (val) f << "sheet[id]=" << val << ",";
		val=int(libwps::readU8(input)); // always 0
		if (val) f << "f0=" << val << ",";
		f << "break=[";
		auto N=int((sz-2)/dataSz);
		for (int i=0; i<N; ++i)
		{
			if (dataSz==1)
				f << int(libwps::readU8(input)) << ",";
			else
				f << libwps::readU16(input) << ",";
		}
		f << "],";
		isParsed=needWriteInAscii=true;
		break;
	}

	//
	// selection
	//
	case 0xbb8:
		f.str("");
		f << "Entries(MacSelect):";

		if (sz!=18)
		{
			WPS_DEBUG_MSG(("LotusParser::readDataZone: the mac selection seems bad\n"));
			f << "###";
			break;
		}
		for (int i=0; i<3; ++i)   // f0=0, f1=f2=1
		{
			val=int(libwps::read16(input));
			if (val) f << "f" << i << "=" << val << ",";
		}
		for (int i=0; i<3; ++i)
		{
			auto row=int(libwps::readU16(input));
			auto sheet=int(libwps::readU8(input));
			auto col=int(libwps::readU8(input));
			f << "C" << col << "-" << row;
			if (sheet) f << "[" << sheet << "],";
			else f << ",";
		}
		isParsed=needWriteInAscii=true;
		break;

	//
	// style
	//
	case 0xfa0: // wk3
		isParsed=m_styleManager->readFontStyleA0(stream, endPos);
		break;
	case 0xfa1: // wk6-wk9, with size 26
		f.str("");
		f << "Entries(FontStyle):";
		break;
	case 0xfaa: // 10Style
	case 0xfab:
		isParsed=m_styleManager->readLineStyle(stream, endPos, type==0xfaa ? 0 : 1);
		break;
	case 0xfb4: // 20 Style
		isParsed=m_styleManager->readColorStyle(stream, endPos);
		break;
	case 0xfbe: // 30Style
		isParsed=m_styleManager->readFormatStyle(stream, endPos);
		break;
	case 0xfc8: //
		isParsed=m_styleManager->readGraphicStyle(stream, endPos);
		break;
	case 0xfc9: // 40Style: lotus 123
		isParsed=m_styleManager->readGraphicStyleC9(stream, endPos);
		break;
	case 0xfd2: // 50Style
		isParsed=m_styleManager->readCellStyleD2(stream, endPos);
		break;
	case 0xfdc:
		isParsed=readMacFontName(stream, endPos);
		break;
	case 0xfe6: // wk5
		if (version()==3)
			isParsed=m_styleManager->readCellStyleE6(stream, endPos);
		else if (version()>3)
			isParsed=m_styleManager->readStyleE6(stream, endPos);
		break;
	case 0xff0: // wk5
		isParsed=m_styleManager->readFontStyleF0(stream, endPos);
		break;

	// 0xfe6: X X CeId : 60Style

	//
	// graphic
	//

	case 0x2328: // begin(wk3mac)
		isParsed=m_graphParser->readZoneBegin(stream, endPos);
		break;
	case 0x2332: // line(wk3mac)
	case 0x2346: // rect, rectoval, rect(wk3mac)
	case 0x2350: // arc(wk3mac)
	case 0x2352: // rect shadow(wk3mac)
	case 0x23f0: // frame(wk3mac)
		isParsed=m_graphParser->readZoneData(stream, endPos, type);
		break;
	case 0x23fa: // textbox data(wk3mac)
		isParsed=m_graphParser->readTextBoxData(stream, endPos);
		break;

	case 0x2710:   // chart data(wk3mac)
	{
		int chartId=-1;
		isParsed=m_chartParser->readMacHeader(stream, endPos, chartId);
		if (isParsed && chartId>=0)
			m_graphParser->setChartId(chartId);
		break;
	}
	case 0x2774:
		isParsed=m_chartParser->readMacPlacement(stream, endPos);
		break;
	case 0x277e:
		isParsed=m_chartParser->readMacLegend(stream, endPos);
		break;
	case 0x2788:
		isParsed=m_chartParser->readMacPlotArea(stream, endPos);
		break;
	case 0x27d8:
		isParsed=m_chartParser->readMacAxis(stream, endPos);
		break;
	case 0x27e2:
		isParsed=m_chartParser->readMacSerie(stream, endPos);
		break;
	// case 0x27ec: always 030000
	case 0x2846: // only 3d
		isParsed=m_chartParser->readMacFloor(stream, endPos);
		break;
	case 0x2904: // only if manual
		isParsed=m_chartParser->readMacPosition(stream, endPos);
		break;

	//
	// chart
	//

	case 0x2a30: // plot area style?
		isParsed=m_chartParser->readPlotArea(stream, endPos);
		break;
	case 0x2a31: // serie style
		isParsed=m_chartParser->readSerie(stream, endPos);
		break;
	case 0x2a32: // serie name
		isParsed=m_chartParser->readSerieName(stream, endPos);
		break;
	case 0x2a33: // serie width
		isParsed=m_chartParser->readSerieWidth(stream, endPos);
		break;
	case 0x2a34: // serie of font, before 2a35
		isParsed=m_chartParser->readFontsStyle(stream, endPos);
		break;
	case 0x2a35: // list of frames styles
		isParsed=m_chartParser->readFramesStyle(stream, endPos);
		break;
	//
	// mac pict
	//
	case 0x240e:
		isParsed=m_graphParser->readPictureDefinition(stream, endPos);
		break;
	case 0x2410:
		isParsed=m_graphParser->readPictureData(stream, endPos);
		break;

	//
	// mac printer
	//
	case 0x2af8:
		isParsed=readDocumentInfoMac(stream, endPos);
		break;
	case 0x2afa:
		f.str("");
		f << "Entries(PrinterMacUnkn1):";
		if (sz!=3)
		{
			WPS_DEBUG_MSG(("LotusParser::readDataZone: the printer unkn1 seems bad\n"));
			f << "###";
			break;
		}
		for (int i=0; i<3; ++i)
		{
			val=int(libwps::readU8(input));
			static int const expected[]= {0x1f, 0xe0, 0/*or 1*/};
			if (val!=expected[i])
				f << "f" << i << "=" << val << ",";
		}
		isParsed=needWriteInAscii=true;
		break;
	case 0x2afb:
	{
		f.str("");
		f << "Entries(PrinterMacName):";
		if (sz<3)
		{
			WPS_DEBUG_MSG(("LotusParser::readDataZone: the printername seems bad\n"));
			f << "###";
			break;
		}
		val=int(libwps::read16(input));
		if (val!=20) f << "f0=" << val << ",";
		std::string name("");
		for (long i=4; i<sz; ++i)
		{
			auto c=char(libwps::readU8(input));
			if (!c) break;
			name+=c;
		}
		f << name << ",";
		isParsed=needWriteInAscii=true;
		break;
	}
	case 0x2afc:
		f.str("");
		f << "Entries(PrintMacInfo):";
		if (sz<120)
		{
			WPS_DEBUG_MSG(("LotusParser::readDataZone: the printinfo seems bad\n"));
			f << "###";
			break;
		}
		isParsed=needWriteInAscii=true;
		break;

	case 0x32e7:
		isParsed=m_styleManager->readMenuStyleE7(stream, endPos);
		break;

	// 32e7: related to style ?
	case 0x36b0:
		isParsed=m_spreadsheetParser->readSheetName1B(stream, endPos);
		break;

	//
	// 4268, 4269
	//

	case 0x4a38:
		f.str("");
		f << "Entries(LinkUnkA):";
		break;
	case 0x4a39:
		f.str("");
		f << "Entries(LinkUnkB):";
		break;
	case 0x6590:
		isParsed=m_spreadsheetParser->readNote(stream, endPos);
		break;

	default:
		break;
	}
	if (!isParsed || needWriteInAscii)
	{
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
	}
	if (input->tell()!=endPos)
		ascFile.addDelimiter(input->tell(),'|');
	input->seek(endPos, librevenge::RVNG_SEEK_SET);
	return true;
}

bool LotusParser::readZoneV3(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input));
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	if (sz<0 || !stream->checkFilePosition(endPos))
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	f << "Entries(Data" << std::hex << type << std::dec << "N):";
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	if (input->tell()!=endPos && input->tell()!=pos)
		ascFile.addDelimiter(input->tell(),'|');
	input->seek(endPos, librevenge::RVNG_SEEK_SET);
	return true;
}

bool LotusParser::readZone1(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto id = int(libwps::readU8(input));
	if (libwps::readU8(input) != 1)
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	if (sz<0 || !stream->checkFilePosition(endPos))
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	f << "Entries(Zone1):";
	int val;
	bool isParsed=false;
	switch (id)
	{
	case 0: // list of ids all between 1-N
	case 3: // always follow an id zone, with some value between 0 and N
	case 0xb: // always follow a level[close] zone, with some value between 0 and N, meaning unsure...
		f << (id==0 ? "id=" : id==3 ? "parent[id]," : "parent2[id],");
		if (sz!=4)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone1: the size seems bad for zone %d\n", id));
			f << "###";
			break;
		}
		val=int(libwps::readU32(input));
		if (val)
		{
			if (id==3)
				m_state->m_actualZoneParentId=val;
			f << "Z" << val << ",";
		}
		if (id==0) m_state->m_actualZoneId=val;
		break;
	case 4:
		f << "stack1[open],";
		if (sz!=4)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone1: the size seems bad for zone 4\n"));
			f << "###";
			break;
		}
		// 01[XX][YY]00 where Y is a small even number
		m_state->m_zone1Stack.push_back(libwps::readU32(input));
		f << m_state->getZone1StackDebugName();
		break;
	case 5:
		f << "stack1[close],";
		if (sz!=4)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone1: the size seems bad for zone 5\n"));
			f << "###";
			break;
		}
		else
		{
			unsigned long lVal=libwps::readU32(input);
			if (m_state->m_zone1Stack.empty() || m_state->m_zone1Stack.back()!=lVal)
			{
				WPS_DEBUG_MSG(("LotusParser::readZone1: the value seems bad for zone 5\n"));
				f << "###val=" << std::hex << lVal << ",";
			}
			if (!m_state->m_zone1Stack.empty()) m_state->m_zone1Stack.pop_back();
			f << m_state->getZone1StackDebugName();
		}
		break;
	// level 1=table, 2=col, 3=row
	case 0x6: // no data
		f << "level[open],";
		m_state->m_actualLevels.push_back(Vec2i(0,0));
		f << "[" << m_state->getLevelsDebugName() << "],";
		break;
	case 0x7: // no data
		f << "level[close]";
		if (m_state->m_actualLevels.empty())
		{
			WPS_DEBUG_MSG(("LotusParser::readZone1: the level seems bad\n"));
			f << "###";
			break;
		}
		else
			m_state->m_actualLevels.pop_back();
		f << "[" << m_state->getLevelsDebugName() << "],";
		break;
	case 0x9: // appear one time at the beginning of the file, pos~8
		f << "dimension,";
		if (sz!=20)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone1: the size seems bad for zone 9\n"));
			f << "###";
			break;
		}
		int dim[4];
		for (int &i : dim) i=int(libwps::read32(input));
		f << "dim=" << WPSBox2i(Vec2i(dim[0],dim[1]), Vec2i(dim[2],dim[3])) << ",";
		for (int i=0; i<2; ++i)   // always 1,0
		{
			val=int(libwps::readU16(input));
			if (val!=1-i) f << "f" << i << "=" << val << ",";
		}
		break;
	case 0xa:   // appear one time at the beginning of the file, pos=3
	{
		f << "typea,";
		if (sz<24 || (sz%8)!=0)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone1: the size seems bad for zone a\n"));
			f << "###";
			break;
		}
		for (int i=0; i<11; ++i)   // f0=cd92|cebd|1cc8, f1=e|f
		{
			val=int(libwps::readU16(input));
			if (val) f << "f" << i << "=" << std::hex << val << std::dec << ",";
		}
		auto N=int(libwps::readU16(input));
		f << "N=" << N << ",";
		if (24+N*8!=sz)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone1: the N value seems bad for zone a\n"));
			f << "###";
			break;
		}
		for (int i=0; i<N; ++i)
		{
			f << "unk" << i << "=[";
			for (int j=0; j<4; ++j)   // f0=12|18|24|800, f1=15e|2913, f2=4|37|41
			{
				val=int(libwps::readU16(input));
				if (val) f << "f" << j << "=" << std::hex << val << std::dec << ",";
			}
			f << "],";
		}
		break;
	}
	case 0xc: // one by file, after a level[close], before a level[open]
		f << "typec,";
		if (sz!=4)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone1: the size seems bad for zone c\n"));
			f << "###";
			break;
		}
		for (int i=0; i<2; ++i)   // always 0
		{
			val=int(libwps::readU16(input));
			if (val) f << "f" << i << "=" << val << ",";
		}
		break;
	case 0xd: // after a line: the coordinate, after a texbox the text
		isParsed=m_graphParser->readGraphDataZone(stream,endPos);
		break;

	case 0xe: // no data, one by file
		f << "styles[def]=begin,";
		break;
	case 0xf: // no data, one by file
		f << "styles[def]=end,";
		break;
	default:
		f << "type=" << std::hex << id << std::dec << ",";
		break;
	}
	if (!isParsed)
	{
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
	}
	if (input->tell()!=endPos && input->tell()!=pos)
		ascFile.addDelimiter(input->tell(),'|');
	input->seek(endPos, librevenge::RVNG_SEEK_SET);
	return true;
}

bool LotusParser::readSheetZone(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto id = int(libwps::readU8(input));
	if (libwps::readU8(input) != 2)
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	if (sz<0 || !stream->checkFilePosition(endPos))
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	f << "Entries(SheetZone):";
	int val;
	switch (id)
	{
	case 0: // appear at the end of the file: pos~end-3
		f << "rList,";
		if (sz!=10)
		{
			WPS_DEBUG_MSG(("LotusParser::readSheetZone: the size seems bad for zone 200\n"));
			f << "###";
			break;
		}
		m_state->m_actualZoneParentId=0;
		f << "sheet[root]=Z" << int(libwps::readU32(input)) << ",";
		for (int i=0; i<3; ++i)   // f1=57
		{
			val=int(libwps::readU16(input));
			if (val) f << "f" << i << "=" << std::hex << val << std::dec << ",";
		}
		break;
	case 1: // root of all sheet node
		f << "root,";
		if (sz!=78)
		{
			WPS_DEBUG_MSG(("LotusParser::readSheetZone: the size seems bad for zone 201\n"));
			f << "###";
			break;
		}
		for (int i=0; i<10; ++i)   // f0=5|7, f6:big number, f8:big number
		{
			val=int(libwps::readU16(input));
			if (val) f << "f" << i << "=" << std::hex << val << std::dec << ",";
		}
		for (int i=0; i<24; ++i)   // 0
		{
			val=int(libwps::readU16(input));
			if (val) f << "g" << i << "=" << val << ",";
		}
		for (int i=0; i<5; ++i)
		{
			val=int(libwps::readU16(input));
			int const expected[]= {0x4001,0x2003,0x100,0x64,0};
			if (val!=expected[i]) f << "h" << i << "=" << std::hex << val << std::dec << ",";
		}
		break;
	case 2:    // appear after zone 0: pos~end-2
	{
		f << "list,";
		if (sz<16 || (sz%4)!=0)
		{
			WPS_DEBUG_MSG(("LotusParser::readSheetZone: the size seems bad for zone 202\n"));
			f << "###";
			break;
		}
		auto N=int(libwps::readU16(input));
		f << "N=" << N << ",";
		if (16+4*N!=sz)
		{
			WPS_DEBUG_MSG(("LotusParser::readSheetZone: the N value seems bad for zone 202\n"));
			f << "###";
			break;
		}
		if (!m_state->m_sheetZoneIdList.empty())
		{
			WPS_DEBUG_MSG(("LotusParser::readSheetZone: oops the sheet list is not empty\n"));
		}
		f << "zones=[";
		for (int i=0; i<N; ++i)
		{
			m_state->m_sheetZoneIdList.push_back(int(libwps::readU32(input)));
			f << "Z" << m_state->m_sheetZoneIdList.back() << ",";
		}
		f << "],";
		for (int i=0; i<7; ++i)   // f3=1-8
		{
			val=int(libwps::readU16(input));
			if (val) f << "f" << i << "=" << val << ",";
		}
		break;
	}
	case 4:
	{
		f << "name,";
		if (sz<14)
		{
			WPS_DEBUG_MSG(("LotusParser::readSheetZone: the size seems bad for zone 204\n"));
			f << "###";
			break;
		}
		for (int i=0; i<4; ++i)   // f2=0|1
		{
			val=int(libwps::readU16(input));
			int const expected[]= {0x200,0x11a,0,0};
			if (val!=expected[i]) f << "f" << i << "=" << std::hex << val << std::dec << ",";
		}
		auto N=int(libwps::readU16(input));
		f << "N=" << N << ",";
		if (14+N!=sz)
		{
			WPS_DEBUG_MSG(("LotusParser::readSheetZone: the N value seems bad for zone 204\n"));
			f << "###";
			break;
		}
		std::string text;
		for (int i=0; i<N; ++i) text+=char(libwps::readU8(input));
		f << text << ",";
		for (int i=0; i<2; ++i)   // g0=5f|76
		{
			val=int(libwps::readU16(input));
			if (val) f << "g" << i << "=" << val << ",";
		}
		break;
	}
	case 5: // defined a child of a sheetNames
		f << "Data0,";
		if (sz!=4)
		{
			WPS_DEBUG_MSG(("LotusParser::readSheetZone: the size seems bad for zone 205\n"));
			f << "###";
			break;
		}
		val=int(libwps::readU32(input));
		if (!val) break;
		if (m_state->m_dataZoneIdToSheetZoneIdMap.find(val) !=
		        m_state->m_dataZoneIdToSheetZoneIdMap.end())
		{
			WPS_DEBUG_MSG(("LotusParser::readSheetZone: the zone %d has already a parent\n", m_state->m_actualZoneId));
		}
		else
			m_state->m_dataZoneIdToSheetZoneIdMap[val]=m_state->m_actualZoneId;
		f << "Z" << val << ",";
		break;

	// two by files: first one a sheet sub zone, second close id
	case 0x82:
	case 0x83:
	case 0x84:
	case 0x93:
	case 0x94:
	case 0x95: // column definition
	case 0x96: // row definition
	{
		auto subZId=size_t(id&0x1F);
		f << "sheetC" << std::hex << subZId << std::dec << "[" << (m_state->m_sheetSubZoneOpened[subZId] ? "close" : "open") << "],";
		m_state->m_sheetSubZoneOpened[subZId]=!m_state->m_sheetSubZoneOpened[subZId];
		break;
	}
	case 0x80: // one by sheet, between 288 and 287
		f << "sheetB0,";
		if (sz!=8)
		{
			WPS_DEBUG_MSG(("LotusParser::readSheetZone: the size seems bad for zone 280\n"));
			f << "###";
			break;
		}
		for (int i=0; i<4; ++i)   // 0|1
		{
			val=int(libwps::readU16(input));
			if (val) f << "f" << i << "=" << val << ",";
		}
		break;
	case 0x81: // one by sheet, in sheetData0
		f << "sheetB1,";
		if (sz!=44)
		{
			WPS_DEBUG_MSG(("LotusParser::readSheetZone: the size seems bad for zone 281\n"));
			f << "###";
			break;
		}
		val=int(libwps::readU16(input));
		if (val!=1) f << "f0=" << val << ",";
		f << "unkn=[";
		for (int i=0; i<4; ++i)   // maybe some some id
		{
			val=int(libwps::readU16(input));
			if (val)
				f << val << ",";
			else
				f << "_,";
		}
		f << "],";
		for (int i=0; i<17; ++i)   // g4=ff, g6=1fff|ffff, g8=1, g9=0|small number
		{
			val=int(libwps::read16(input));
			if (val) f << "g" << i << "=" << std::hex << val << std::dec << ",";
		}
		break;
	case 0x85: // in general, I found them in sheetData0+1 and sheetData+2 (with the same value)
		f << "data1,";
		if (sz!=4)
		{
			WPS_DEBUG_MSG(("LotusParser::readSheetZone: the size seems bad for zone 285\n"));
			f << "###";
			break;
		}
		f << "id=Z" << int(libwps::readU32(input)) << ",";
		break;
	case 0x86: // one by file in general after sheetC16
		f << "sheetB6,";
		if (sz!=8 && sz!=10)
		{
			WPS_DEBUG_MSG(("LotusParser::readSheetZone: the size seems bad for zone 286\n"));
			f << "###";
			break;
		}
		for (long i=0; i<sz/2; ++i)   // 100,0,100,2000 in 123:v4.0 , 100,0,100,0,100:v4.5
		{
			val=int(libwps::readU16(input));
			if (val) f << "f" << i << "=" << std::hex << val << std::dec << ",";
		}
		break;
	case 0x87: // one by sheet, after 280
		f << "sheetB7,";
		if (sz!=6)
		{
			WPS_DEBUG_MSG(("LotusParser::readSheetZone: the size seems bad for zone 287\n"));
			f << "###";
			break;
		}
		for (int i=0; i<3; ++i)   // 1,1,0
		{
			val=int(libwps::readU16(input));
			if (val) f << "f" << i << "=" << val << ",";
		}
		break;
	case 0x88: // one by sheet, between 281 and 280
		f << "sheetB8,";
		if (sz!=4)
		{
			WPS_DEBUG_MSG(("LotusParser::readSheetZone: the size seems bad for zone 288\n"));
			f << "###";
			break;
		}
		val=int(libwps::readU32(input)); // always 1, maybe and id
		if (val!=1) f << "f0=" << val << ",";
		break;
	case 0x92: // in sheetC2. Checkme, zone with variable size
		f << "sheetB12,";
		if (sz<28)
		{
			WPS_DEBUG_MSG(("LotusParser::readSheetZone: the size seems bad for zone 288\n"));
			f << "###";
			break;
		}
		for (int i=0; i<14; ++i)   // f0=3600|438,f2=f0|c000|e040, f4=500[256]
		{
			val=int(libwps::readU16(input));
			int const expected[]= {0,0,0,0,0, 0x35d4,0,0x1003,0x2000,0, 0x60,0,0x60,0};
			if (val!=expected[i]) f << "f" << i << "=" << std::hex << val << std::dec << ",";
		}
		// after some 0 or 3ff0
		break;
	case 0x99: // in sheetC15, sheetC16
	case 0x9a: // in sheetC4, sheetC13, sheetC14
		f << "sheetB" << std::hex << (id-0x90) << std::dec << ",";
		if (sz!=10)
		{
			WPS_DEBUG_MSG(("LotusParser::readSheetZone: the size seems bad for zone %d\n", id));
			f << "###";
			break;
		}
		for (int i=0; i<5; ++i)   // always 0
		{
			val=int(libwps::readU16(input));
			if (val) f << "f" << i << "=" << val << ",";
		}
		break;
	default:
		f << "type=" << std::hex << id << std::dec << ",";
		break;
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	if (input->tell()!=endPos && input->tell()!=pos)
		ascFile.addDelimiter(input->tell(),'|');
	input->seek(endPos, librevenge::RVNG_SEEK_SET);
	return true;
}

bool LotusParser::readZone4(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto id = int(libwps::readU8(input));
	if (libwps::readU8(input)!=4)
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	if (sz<0 || !stream->checkFilePosition(endPos))
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	f << "Entries(Zone4):";
	int val;
	switch (id)
	{
	case 0:   // one by sheet, page definition ?
	{
		f << "sheet[page,def],";
		if (sz<90)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone4: the size seems bad for zone 0\n"));
			f << "###";
			break;
		}
		f << "dims=["; // margins + page dim?
		for (int i=0; i<6; ++i)
		{
			val=int(libwps::read32(input));
			if (val)
				f << val << ",";
			else
				f << "_,";
		}
		f << "],";
		f << "unkn=[";
		for (int i=0; i<9; ++i)   // list of 0|2|5
		{
			val=int(libwps::read16(input));
			if (val)
				f << val << ",";
			else
				f << "_,";
		}
		f << "],";
		for (int i=0; i<3; ++i)   // some zone
		{
			val=int(libwps::read32(input));
			if (val)
				f << "zone" << i << "=Z" << val << ",";
		}
		for (int i=0; i<12; ++i)   // f2=0|8, f4=119|131|137, f5=0|7, f11=0|9
		{
			val=int(libwps::read16(input));
			int const expected[]= {0,0,0,0,0x131, 0,1,1,0x270f,1, 0x64,1};
			if (val!=expected[i])
				f << "f" << i << "=" << val << ",";
		}
		f << "fl=[";
		for (int i=0; i<10; ++i)   // list of 0|1|3, often 1
		{
			val=int(libwps::readU8(input));
			if (val)
				f << val << ",";
			else
				f << "_,";
		}
		f << "],";
		if (sz<92) break;
		// unsure from here, often the main style name?
		std::string name;
		while (input->tell()<endPos)
		{
			auto c=char(libwps::readU8(input));
			if (!c) break;
			name+=c;
		}
		if (!name.empty()) f << name << ",";
		break;
	}
	case 1: // after zoneA0, in general 6 item f0=0-5, tab? or item style?
		f << "zoneA1,";
		if (sz!=7)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone4: the size seems bad for zone 1\n"));
			f << "###";
			break;
		}
		f << "id=" << int(libwps::readU8(input)) << ",";
		for (int i=0; i<3; ++i)   // 0
		{
			val=int(libwps::readU16(input));
			if (val) f << "f" << i << "=" << val << ",";
		}
		break;
	case 3:   // in sheet definition, after the style
	{
		f << "footerDef,";
		if (sz<31)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone4: the size seems bad for zone 3\n"));
			f << "###";
			break;
		}
		for (int i=0; i<13; ++i)   // f4==5,f6=4,f7=76,f8=2,f9=2,f10=66,f11=1
		{
			val=int(libwps::readU16(input));
			if (val) f << "f" << i << "=" << val << ",";
		}
		val=int(libwps::readU8(input)); // 0
		if (val) f << "f13=" << val << ",";
		for (int s=0; s<2; ++s)
		{
			auto sSz=int(libwps::readU16(input));
			if (input->tell()+sSz+(s==0 ? 2 : 0)>endPos)
			{
				WPS_DEBUG_MSG(("LotusParser::readZone4: the size seems bad for zone 3\n"));
				f << "###";
				break;
			}
			std::string name;
			for (int i=0; i<sSz; ++i)
			{
				auto c=char(libwps::readU8(input));
				if (c) name+=c;
				else if (i+1!=sSz)
				{
					WPS_DEBUG_MSG(("LotusParser::readZone4: find odd char in zone 3\n"));
					f << "###";
				}
			}
			if (!name.empty()) f << "string" << s << "=" << name << ",";
		}
		break;
	}
	case 0x80: // rare, present in sheet definition
		f << "chartSheet,";
		if (sz!=4)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone4: the size seems bad for zone 80\n"));
			f << "###";
			break;
		}
		f << "id=Z" << int(libwps::readU32(input)) << ",";
		break;
	case 0x81: // after 480. checkme chart series
		f << "chartSeries,";
		if (sz!=12)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone4: the size seems bad for zone 81\n"));
			f << "###";
			break;
		}
		f << "unkn=[";
		for (int i=0; i<3; ++i)
		{
			val=int(libwps::readU32(input));
			if (val)
				f << val << ",";
			else
				f << "_,";
		}
		f << "],";
		break;
	default:
		f << "type=" << std::hex << id << std::dec << ",";
		break;
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	if (input->tell()!=endPos && input->tell()!=pos)
		ascFile.addDelimiter(input->tell(),'|');
	input->seek(endPos, librevenge::RVNG_SEEK_SET);
	return true;
}

bool LotusParser::readChartZone(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto id = int(libwps::readU8(input));
	if (libwps::readU8(input)!=5)
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	if (sz<0 || !stream->checkFilePosition(endPos))
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	f << "Entries(ChartZone):";
	int val;
	switch (id)
	{
	case 0:
	{
		f << "name,";
		if (sz<6)
		{
			WPS_DEBUG_MSG(("LotusParser::readChartZone: the size seems bad for zone 0\n"));
			f << "###";
			break;
		}
		for (int i=0; i<2; ++i)   // always 0
		{
			val=int(libwps::readU16(input));
			if (val)
				f << "f" << i << "=" << val << ",";
		}
		auto sSz=int(libwps::readU16(input));
		if (6+sSz>sz)
		{
			WPS_DEBUG_MSG(("LotusParser::readChartZone: the size seems bad for zone 0\n"));
			f << "###";
			break;
		}
		std::string name;
		for (int i=0; i<sSz; ++i)
		{
			auto c=char(libwps::readU8(input));
			if (c) name+=c;
			else if (i+1!=sSz)
			{
				WPS_DEBUG_MSG(("LotusParser::readChartZone: find odd char in zone 0\n"));
				f << "###";
			}
		}
		if (!name.empty()) f << name << ",";
		break;
	}
	case 2:
		f << "series,";
		if (sz!=12)
		{
			WPS_DEBUG_MSG(("LotusParser::readChartZone: the size seems bad for zone 2\n"));
			f << "###";
			break;
		}
		f << "unkn=[";
		for (int i=0; i<3; ++i)
		{
			val=int(libwps::readU32(input));
			if (val)
				f << val << ",";
			else
				f << "_,";
		}
		f << "],";
		break;
	case 3: // last zone
		f << "end,";
		if (sz!=0)
		{
			WPS_DEBUG_MSG(("LotusParser::readChartZone: the size seems bad for zone 3\n"));
			f << "###";
			break;
		}
		break;
	default:
		f << "type=" << std::hex << id << std::dec << ",";
		break;
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	if (input->tell()!=endPos && input->tell()!=pos)
		ascFile.addDelimiter(input->tell(),'|');
	input->seek(endPos, librevenge::RVNG_SEEK_SET);
	return true;
}

bool LotusParser::readRefZone(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto id = int(libwps::readU8(input));
	if (libwps::readU8(input)!=6)
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	if (sz<0 || !stream->checkFilePosition(endPos))
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	f << "Entries(RefZone):";
	int val;
	switch (id)
	{
	case 0x40: // after 642
		f << "cells,";
		if (sz!=12)
		{
			WPS_DEBUG_MSG(("LotusParser::readRefZone: the size seems bad for zone 640\n"));
			f << "###";
			break;
		}
		for (int i=0; i<6; ++i) // C?R?S? <-> C?R?S?: checkme maybe this stores also the range
		{
			f << int(libwps::readU16(input));
			if (i==2) f << "<->";
			else if (i==5) f << ",";
			else f << ":";
		}
		break;
	case 0x42: // after 407
		f << "begin,";
		if (sz!=4)
		{
			WPS_DEBUG_MSG(("LotusParser::readRefZone: the size seems bad for zone 642\n"));
			f << "###";
			break;
		}
		val=int(libwps::readU32(input)); // always 3
		if (val!=3) f << "f0=" << val << ",";
		break;
	case 0x43:   // find something similar to A:E7 for a cell or to B:H3..B:H80 for a cell list
	{
		f << "name,";
		if (sz<=0)
		{
			WPS_DEBUG_MSG(("LotusParser::readRefZone: the size seems bad for zone 643\n"));
			f << "###";
			break;
		}
		std::string name;
		for (long i=0; i<sz; ++i)
		{
			auto c=char(libwps::readU8(input));
			if (c) name+=c;
			else if (i+1!=sz)
			{
				WPS_DEBUG_MSG(("LotusParser::readRefZone: find odd char in zone 643\n"));
				f << "###";
			}
		}
		if (!name.empty()) f << name << ",";
		break;
	}
	default:
		f << "type=" << std::hex << id << std::dec << ",";
		break;
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	if (input->tell()!=endPos && input->tell()!=pos)
		ascFile.addDelimiter(input->tell(),'|');
	input->seek(endPos, librevenge::RVNG_SEEK_SET);
	return true;
}

bool LotusParser::readZone7(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto id = int(libwps::readU8(input));
	if (libwps::readU8(input)!=7)
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	if (sz<0 || !stream->checkFilePosition(endPos))
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	f << "Entries(Zone7)[" << std::hex << id << std::dec << "]:";

	// normally, 780, ..., 701, 702, ..., 703, ..., 704, ...
	// in 704: the cell style
	int val;
	switch (id)
	{
	case 1:
		// empty zone (or container of 702)
		if (sz!=28)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone7: the size seems bad for zone 1\n"));
			f << "###";
			break;
		}
		for (int i=0; i<4; ++i)
		{
			val=int(libwps::readU16(input));
			if (val) f << "f" << i << "=" << std::hex << val << std::dec << ",";
		}
		f << "mat=[";
		for (int i=0; i<4; ++i)
		{
			double res;
			bool isNan;
			long actPos=pos;
			if (libwps::readDouble4(input, res, isNan))
				f << res << ",";
			else
			{
				f << "###";
				input->seek(actPos+2, librevenge::RVNG_SEEK_SET);
			}
		}
		f << "],";
		for (int i=0; i<2; ++i)
		{
			val=int(libwps::readU16(input));
			if (val) f << "g" << i << "=" << std::hex << val << std::dec << ",";
		}
		break;
	case 2:
		// precedes the LotusbE
		if (sz!=8)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone7: the size seems bad for zone 2\n"));
			f << "###";
			break;
		}
		for (int i=0; i<4; ++i)   // always 0
		{
			val=int(libwps::readU16(input));
			if (val) f << "f" << i << "=" << val << ",";
		}
		break;
	case 3:
		// precedes LotuscE, Lotus1cE, the col size, the link, the sheet name, the cells content
		f << "content,";
		if (sz!=6)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone7: the size seems bad for zone 3\n"));
			f << "###";
			break;
		}
		for (int i=0; i<3; ++i)   // list of 0|1
		{
			val=int(libwps::readU16(input));
			if (val) f << "f" << i << "=" << val << ",";
		}
		break;
	case 4:
		// precedes the cell styles
		f << "styles,";
		if (sz!=0)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone7: the size seems bad for zone 4\n"));
			f << "###";
			break;
		}
		break;
	case 0x80: // first zone, precedes Data105N,Data104N,Data100N,Lotus3E
		f << "first,";
		if (sz!=12)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone7: the size seems bad for zone 80\n"));
			f << "###";
			break;
		}
		for (int i=0; i<6; ++i)   // f0=6f|ef
		{
			val=int(libwps::readU16(input));
			int const expected[]= {0xef, 0, 7, 0, 0x5f, 0x57};
			if (val!=expected[i]) f << "f" << i << "=" << std::hex << val << std::dec << ",";
		}
		break;
	default:
		break;
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	if (input->tell()!=endPos && input->tell()!=pos)
		ascFile.addDelimiter(input->tell(),'|');
	input->seek(endPos, librevenge::RVNG_SEEK_SET);
	return true;
}

bool LotusParser::readZone8(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto id = int(libwps::readU8(input));
	if (id==1)
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		WPSVec3i minC, maxC;
		m_state->getLevels(minC, maxC);
		return m_spreadsheetParser->readCellsFormat801
		       (stream, minC, maxC, m_state->m_sheetSubZoneOpened[0x15] ? 0 :
		        m_state->m_sheetSubZoneOpened[0x16] ? 1 : -1);
	}
	if (libwps::readU8(input)!=8)
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	if (sz<0 || !stream->checkFilePosition(endPos))
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	int const vers=version();
	f << "Entries(Zone8):";
	int val;
	switch (id)
	{
	case 0: // v4: sz2, v4.5: sz4
	{
		f << "level[select],";
		int const expectedSz=vers<=4 ? 2 : 4; // checkme: unsure about vers==5
		if (sz!=expectedSz)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone8: the level size seems bad\n"));
			f << "###";
			break;
		}
		if (m_state->m_actualLevels.empty())
		{
			WPS_DEBUG_MSG(("LotusParser::readZone8: the level seems bad\n"));
			f << "###";
			break;
		}
		long count=int(sz>=4 ? libwps::readU32(input) : libwps::readU16(input));
		Vec2i &zone=m_state->m_actualLevels.back();
		if (int(zone[1]+count)<0)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone8: arg the delta bad\n"));
			f << "###delta=" << count << ",";
			count=0;
		}
		zone[0] = zone[1];
		zone[1] += int(count);
		f << "pos=[" << m_state->getLevelsDebugName() << "],";
		break;
	}
	// 1 already done
	case 2: // very often 802 and 803 are close to each other (in the sheet's zone)
	case 3:
		if (id==2)
			f << "column[def],";
		else
			f << "zoneA" << id << ",";
		if (sz!=2)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone8: the size seems bad for id=%d\n", id));
			f << "###";
			break;
		}
		val=int(libwps::readU16(input)); // 1|2
		if (val!=1) f << "f0=" << val  << ",";
		break;
	case 4:
	{
		f << "zoneA4,";
		if (m_state->m_sheetSubZoneOpened[0x15]) f << "cols,";
		else if (m_state->m_sheetSubZoneOpened[0x16]) f << "rows,";
		if (sz<4)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone8: the size seems bad for 804\n"));
			f << "###";
			break;
		}
		val=int(libwps::readU16(input));
		if (val!=3) f << "f0=" << val << ",";
		auto N=int(libwps::readU16(input));
		f << "N=" << N << ",";
		int const expectedSz=vers<=4 ? 2 : 4; // checkme: unsure about vers==5
		if (sz!=4+N*expectedSz)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone8: the N value seems bad for 804\n"));
			f << "###";
			break;
		}
		f << "unk=[";
		for (int i=0; i<N; ++i)
		{
			f << int(libwps::readU8(input));
			f << ":" << int(libwps::readU8(input));
			if (expectedSz==4)
			{
				f << "<->" << int(libwps::readU8(input));
				f << ":" << int(libwps::readU8(input));
			}
			f << ",";
		}
		f << "],";
		break;
	}
	case 0x83: // often the last 80X's zone
		f << "zoneB3,";
		if (sz!=5)
		{
			WPS_DEBUG_MSG(("LotusParser::readZone8: the size seems bad for 883\n"));
			f << "###";
			break;
		}
		for (int i=0; i<5; ++i)   // always 0
		{
			val=int(libwps::readU8(input));
			if (val) f << "f" << i << "=" << val  << ",";
		}
		break;
	default:
		f << "type=" << std::hex << id << std::dec << ",";
		break;
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	if (input->tell()!=endPos && input->tell()!=pos)
		ascFile.addDelimiter(input->tell(),'|');
	input->seek(endPos, librevenge::RVNG_SEEK_SET);
	return true;
}

bool LotusParser::readVersionZone(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto id = int(libwps::readU8(input));
	if (libwps::readU8(input)!=0xa)
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	if (sz<0 || !stream->checkFilePosition(endPos))
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	f << "Entries(VersionZone):";
	// TODO
	switch (id)
	{
	default:
		f << "type=" << std::hex << id << std::dec << ",";
		break;
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	if (input->tell()!=endPos && input->tell()!=pos)
		ascFile.addDelimiter(input->tell(),'|');
	input->seek(endPos, librevenge::RVNG_SEEK_SET);
	return true;
}
////////////////////////////////////////////////////////////
//   generic
////////////////////////////////////////////////////////////
bool LotusParser::readMacFontName(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	const int vers=version();
	long pos = input->tell();
	long sz=endPos-pos;
	f << "Entries(MacFontName):";
	if ((vers<=1 && sz<7) || (vers>1 && sz!=42))
	{
		WPS_DEBUG_MSG(("LotusParser::readMacFontName: the zone size seems bad\n"));
		f << "###";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	if (vers<=1)
	{
		// seems only to exist in a lotus mac file, so revert the default encoding to MacRoman if undef
		if (m_state->m_fontType==libwps_tools_win::Font::UNKNOWN)
			m_state->m_fontType=libwps_tools_win::Font::MAC_ROMAN;
		m_state->m_isMacFile=true;
		auto id=int(libwps::readU16(input));
		f << "FN" << id << ",";
		auto val=int(libwps::readU16(input)); // always 0?
		if (val)
			f << "f0=" << val << ",";
		val=int(libwps::read16(input)); // find -1, 30 (Geneva), 60 (Helvetica)
		if (val)
			f << "f1=" << val << ",";
		librevenge::RVNGString name("");
		bool nameOk=true;
		for (long i=0; i<sz-6; ++i)
		{
			auto c=char(libwps::readU8(input));
			if (!c) break;
			if (nameOk && !(c==' ' || (c>='0'&&c<='9') || (c>='a'&&c<='z') || (c>='A'&&c<='Z')))
			{
				nameOk=false;
				WPS_DEBUG_MSG(("LotusParser::readMacFontName: find odd character in name\n"));
				f << "#";
			}
			name.append(c);
		}
		f << name.cstr() << ",";
		if (m_state->m_fontsMap.find(id)!=m_state->m_fontsMap.end())
		{
			WPS_DEBUG_MSG(("LotusParser::readMacFontName: a font with id=%d already exists\n", id));
			f << "###id,";
		}
		else if (nameOk && !name.empty())
		{
			auto encoding=name!="Symbol" ? libwps_tools_win::Font::MAC_ROMAN : libwps_tools_win::Font::MAC_SYMBOL;
			LotusParserInternal::Font font(encoding);
			font.m_name=name;
			m_state->m_fontsMap.insert(std::map<int, LotusParserInternal::Font>::value_type(id,font));
		}
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}

	int id=0;
	for (int i=0; i<4; ++i)
	{
		auto val=int(libwps::readU8(input)); // 0|1
		if (i==1)
		{
			id=val;
			f << "FN" << id << ",";
		}
		else if (val)
			f << "fl" << i << "=" << val << ",";
	}
	for (int i=0; i<2; ++i)   // f1=0|1288
	{
		auto val=int(libwps::read16(input));
		if (val)
			f << "f" << i << "=" << val << ",";
	}
	librevenge::RVNGString name("");
	bool nameOk=true;
	for (int i=0; i<8; ++i)
	{
		auto c=char(libwps::read8(input));
		if (!c) break;
		if (nameOk && !(c==' ' || (c>='0'&&c<='9') || (c>='a'&&c<='z') || (c>='A'&&c<='Z')))
		{
			nameOk=false;
			WPS_DEBUG_MSG(("LotusParser::readMacFontName: find odd character in name\n"));
			f << "#";
		}
		name.append(c);
	}
	f << name.cstr() << ",";
	if (m_state->m_fontsMap.find(id)!=m_state->m_fontsMap.end())
	{
		WPS_DEBUG_MSG(("LotusParser::readMacFontName: a font with id=%d already exists\n", id));
		f << "###id,";
	}
	else if (nameOk && !name.empty())
	{
		LotusParserInternal::Font font(getDefaultFontType());
		font.m_name=name;
		m_state->m_fontsMap.insert(std::map<int, LotusParserInternal::Font>::value_type(id,font));
	}
	input->seek(pos+16, librevenge::RVNG_SEEK_SET);
	if (input->tell()!=endPos)
	{
		ascFile.addDelimiter(input->tell(),'|');
		input->seek(endPos, librevenge::RVNG_SEEK_SET);
	}
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusParser::readFMTStyleName(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = int(libwps::read16(input));
	if (type!=0xb6)
	{
		WPS_DEBUG_MSG(("LotusParser::readFMTStyleName: not a font name definition\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	if (sz<8)
	{
		WPS_DEBUG_MSG(("LotusParser::readFMTStyleName: the zone size seems bad\n"));
		ascFile.addPos(pos);
		ascFile.addNote("Entries(FMTStyleName):###");
		return true;
	}
	f << "Entries(FMTStyleName):";
	f << "id=" << libwps::readU16(input) << ",";
	std::string name;
	for (int i=0; i<6; ++i)
	{
		auto c=char(libwps::readU8(input));
		if (c==0) break;
		name+= c;
	}
	f << "title=" << name << ",";
	input->seek(pos+12, librevenge::RVNG_SEEK_SET);
	name.clear();
	for (long i=0; i<sz-8; ++i) name+= char(libwps::readU8(input));
	f << name << ",";
	if (input->tell()!=endPos)
	{
		WPS_DEBUG_MSG(("LotusParser::readFMTStyleName: find extra data\n"));
		f << "###extra";
		input->seek(endPos, librevenge::RVNG_SEEK_SET);
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusParser::readLinkZone(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = int(libwps::read16(input));
	if (type!=0xa)
	{
		WPS_DEBUG_MSG(("LotusParser::readLinkZone: not a link definition\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	f << "Entries(Link):";
	if (sz < 19)
	{
		WPS_DEBUG_MSG(("LotusParser::readLinkZone: the zone is too short\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	type=int(libwps::read16(input));
	if (type==0) // fixme: find if this is a note, so that we can retrieve it
		f << "chart/note/...,";
	else if (type==1)
		f << "file,";
	else
	{
		WPS_DEBUG_MSG(("LotusParser::readLinkZone: find unknown type\n"));
		f << "##type=" << type << ",";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << "ID=" << int(libwps::readU8(input)) << ","; // 0,19,42,53,ff
	auto id = int(libwps::readU8(input));
	f << "id=" << id << ",";

	Link link;
	// C0: current selection
	// ----- chart ----:
	// G[23-28] color series 0->5
	// G[2a-2f] hatch series 0->5
	// G[39-3e]: data series 0, 1, ...
	// G[3f]: chart axis 0
	// G[40-45]: legend serie 0->5
	// G[47][22,27,2c,31,36,3b,40,45,4a,4f,54,59,5e]: data serie 6-18 (+1 label)
	// G[48][23,28,2d,32]: serie 19-22 (+1 label)

	// G[4c-4e]: unit axis x,y,ysecond
	// G[4f-51]: label axis x,y,ysecond
	// G[52-55]: title, subtile, note1, note2
	// ----- unknown -----:
	// P3: can contains often a basic name or cells zone
	// Q[0-2]: contains often <<XXX>>YYY: link to another sheetname?
	for (int i=0; i<14; ++i)
	{
		auto c=char(libwps::readU8(input));
		if (!c) break;
		link.m_name += c;
	}
	f << "\"" << link.m_name << "\",";
	input->seek(pos+4+18, librevenge::RVNG_SEEK_SET);
	bool ok=true;
	switch (type)
	{
	case 0:
		if (sz<26)
		{
			WPS_DEBUG_MSG(("LotusParser::readLinkZone: the chart zone seems too short\n"));
			f << "###";
			break;
		}
		for (int i=0; i<2; ++i)
		{
			auto row=int(libwps::readU16(input));
			auto table=int(libwps::readU8(input));
			auto col=int(libwps::readU8(input));
			link.m_cells[i]=WPSVec3i(col,row,table);
			f << "C" << col << "-" << row;
			if (table) f << "[" << table << "]";
			if (i==0)
				f << "<->";
			else
				f << ",";
		}
		break;
	case 1:
	{
		if (sz>18)
			link.m_linkName=libwps_tools_win::Font::unicodeString(input.get(), static_cast<unsigned long>(sz-18), getDefaultFontType());
		f << "link=" << link.m_linkName.cstr() << ",";
		break;
	}
	default:
		ok=false;
		WPS_DEBUG_MSG(("LotusParser::readLinkZone: find unknown type\n"));
		f << "###";
		break;
	}
	if (ok)	m_state->m_linkIdToLinkMap.insert(std::multimap<int,Link>::value_type(id, link));
	if (input->tell()!=pos+4+sz && input->tell()+1!=pos+4+sz)
	{
		WPS_DEBUG_MSG(("LotusParser::readLinkZone: the zone seems too short\n"));
		f << "##";
		ascFile.addDelimiter(input->tell(), '|');
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

// ----------------------------------------------------------------------
// Header/Footer/PageDim
// ----------------------------------------------------------------------
bool LotusParser::readDocumentInfoMac(std::shared_ptr<WPSStream> stream, long endPos)
{
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	f << "Entries(DocMacInfo):";
	if (endPos-pos!=51)
	{
		WPS_DEBUG_MSG(("LotusParser::readDocumentInfoMac: the zone size seems bad\n"));
		f << "###";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	int dim[7];
	for (int i=0; i<7; ++i)
	{
		auto val=int(libwps::read8(input));
		if (i==0)
			f << "dim[unkn]=";
		else if (i==1)
			f << "margins=[";
		else if (i==5)
			f << "pagesize=[";
		dim[i]=int(libwps::read16(input));
		f << dim[i];
		if (val) f << "[" << val << "]";
		val=int(libwps::read8(input)); // always 0
		if (val) f << "[" << val << "]";
		f << ",";
		if (i==4 || i==6) f << "],";
	}
	// check order
	if (dim[5]>dim[1]+dim[3] && dim[6]>dim[2]+dim[4])
	{
		m_state->m_pageSpan.setFormWidth(dim[5]);
		m_state->m_pageSpan.setFormLength(dim[6]);
		m_state->m_pageSpan.setMarginLeft(dim[1]);
		m_state->m_pageSpan.setMarginTop(dim[2]);
		m_state->m_pageSpan.setMarginRight(dim[3]);
		m_state->m_pageSpan.setMarginBottom(dim[4]);
	}
	else
		f << "###";
	f << "unkn=[";
	for (int i=0; i<5; ++i)   // 1,1,1,100|inf,1
	{
		auto val=int(libwps::read16(input));
		if (val==9999)
			f << "inf,";
		else if (val)
			f << val << ",";
		else
			f << "_,";
	}
	f << "],";
	for (int i=0; i<13; ++i)   // always 0?
	{
		auto val=int(libwps::read8(input));
		if (val)
			f << "g" << i << "=" << val << ",";
	}
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	return true;
}

////////////////////////////////////////////////////////////
//   decode
////////////////////////////////////////////////////////////
RVNGInputStreamPtr LotusParser::decodeStream(RVNGInputStreamPtr input, long endPos, std::vector<uint8_t> const &key)
{
	if (!input || key.size()!=16)
	{
		WPS_DEBUG_MSG(("LotusParser::decodeStream: the arguments seems bad\n"));
		return RVNGInputStreamPtr();
	}
	long actPos=input->tell();
	input->seek(0,librevenge::RVNG_SEEK_SET);
	librevenge::RVNGBinaryData data;
	if (!libwps::readData(input, static_cast<unsigned long>(endPos), data) || long(data.size())!=endPos || !data.getDataBuffer())
	{
		WPS_DEBUG_MSG(("LotusParser::decodeStream: can not read the original input\n"));
		return RVNGInputStreamPtr();
	}
	auto *buf=const_cast<unsigned char *>(data.getDataBuffer());
	input->seek(actPos,librevenge::RVNG_SEEK_SET);
	uint8_t d7=0;
	bool transform=true;
	while (!input->isEnd())
	{
		long pos=input->tell();
		if (pos+4>endPos) break;
		auto type=int(libwps::readU16(input));
		auto sSz=int(libwps::readU16(input));
		if (pos+4+sSz>endPos)
		{
			input->seek(pos,librevenge::RVNG_SEEK_SET);
			break;
		}
		//   Special case :
		// 123 files:
		//   - the style zone (between 0x10e and 0x10f) is not transformed
		//   - the stack1[open|close] field are not transformed
		if (type==0x10e)
			transform=false;
		else if (type==0x10f)
			transform=true;
		if (type==0x104 || type==0x105 || !transform)
		{
			input->seek(pos+4+sSz,librevenge::RVNG_SEEK_SET);
			continue;
		}
		auto d4=uint8_t(sSz);
		uint8_t d5=key[13];
		for (int i=0; i<sSz; ++i)
		{
			auto c=uint8_t(libwps::readU8(input));
			buf[pos+4+i]=(c^key[d7&0xf]);
			d7=uint8_t(c+d4);
			d4=uint8_t(d4+d5++);
		}
	}
	if (input->tell()!=endPos)
	{
		WPS_DEBUG_MSG(("LotusParser::decodeStream: can not decode the end of the file, data may be bad %lx %lx\n", static_cast<unsigned long>(input->tell()), static_cast<unsigned long>(endPos)));
	}
	RVNGInputStreamPtr res(new WPSStringStream(data.getDataBuffer(), static_cast<unsigned int>(endPos)));
	res->seek(actPos, librevenge::RVNG_SEEK_SET);
	return res;
}

std::vector<uint8_t> LotusParser::retrievePasswordKeys(std::vector<uint8_t> const &fileKeys)
{
	/* let try to detect short password (|password|<=14) by using the
	   fact that fileKeys differ from the keys in two positions.

	   If the password length is less or equal to 12:
	   Using fileKeys[12] and fileKeys[14], we can "retrieve"
	   the password length. Then knowing this length, fileKeys[14]
	   and fileKeys[15] give us the key. Finally, we can retrieve the
	   password and check if it gives us again fileKeys.

	   We can also test password with length 13 or 14 similarly.

	   Note: if |password|>14, we can detect it by testing 256*256 posibilities, but :-~
	 */
	std::vector<uint8_t> res;
	if (fileKeys.size()!=16)
	{
		WPS_DEBUG_MSG(("LotusParser::retrievePasswordKeys: the file keys seems bad\n"));
		return res;
	}
	static uint8_t const defValues[]=
	{
		0xb9,0x5f, 0xd7,0x31, 0xdb,0x75, 9,0x72,
		0x5d,0x85, 0x32,0x11, 0x5,0x11, 0x58,0
	};
	std::map<uint8_t,size_t> diffToPosMap;
	for (size_t i=0; i<14; ++i)
		diffToPosMap[defValues[i+2]^defValues[i]]=i;
	uint8_t diff12=fileKeys[12]^fileKeys[14];
	std::vector<size_t> posToTest;
	if (diffToPosMap.find(diff12)!=diffToPosMap.end() && diffToPosMap.find(diff12)->second+2<14)
	{
		posToTest.push_back(diffToPosMap.find(diff12)->second+2);
		// defValues[0]^defValues[2]=defValues[1]^defValues[3]=0x6e => we must add by hand this position
		if (diff12==0x6e)
			posToTest.push_back(2);
	}
	// check also password with length 13 or 14
	posToTest.push_back(0);
	posToTest.push_back(1);
	for (size_t actPos : posToTest)
	{
		auto key=uint16_t(((fileKeys[14]^defValues[actPos])<<8)|(fileKeys[15]^defValues[actPos+1]));
		res=fileKeys;
		res[7]=uint8_t(res[7]^key);
		res[13]=uint8_t(res[13]^(key>>8));
		// now build the password
		std::string password;
		for (size_t i=0; i<size_t(16-actPos-2); ++i)
			password+=char(res[i]^(key>>((i%2)==0 ? 8 : 0)));
		// check if the password is correct
		uint16_t resKey;
		std::vector<uint8_t> resKeys;
		if (libwps::encodeLotusPassword(password.c_str(), resKey, resKeys, defValues) && key==resKey && res==resKeys)
		{
			WPS_DEBUG_MSG(("LotusParser::retrievePasswordKeys: Find password %s\n", password.c_str()));
			return res;
		}
	}
	return std::vector<uint8_t>();
}

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
