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

#include <cmath>
#include <set>
#include <sstream>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WKSContentListener.h"
#include "WKSSubDocument.h"

#include "WPSCell.h"
#include "WPSEntry.h"
#include "WPSFont.h"
#include "WPSHeader.h"
#include "WPSPageSpan.h"
#include "WPSStringStream.h"
#include "WPSTable.h"

#include "Multiplan.h"

namespace libwps
{

//! Internal: namespace to define internal class of MultiplanParser
namespace MultiplanParserInternal
{
//! the font of a MultiplanParser
struct Font final : public WPSFont
{
	//! constructor
	explicit Font(libwps_tools_win::Font::Type type) : WPSFont(), m_type(type)
	{
	}
	//! destructor
	~Font() final;
	//! font encoding type
	libwps_tools_win::Font::Type m_type;
};

Font::~Font()
{
}

//! a cellule of a Lotus spreadsheet
class Cell final : public WPSCell
{
public:
	/// constructor
	Cell() : WPSCell() { }
	//! call when a cell must be send
	bool send(WPSListenerPtr &/*listener*/) final;

	//! call when the content of a cell must be send
	bool sendContent(WPSListenerPtr &/*listener*/) final
	{
		WPS_DEBUG_MSG(("MultiplanInternal::Cell::sendContent: must not be called\n"));
		return false;
	}
};
bool Cell::send(WPSListenerPtr &/*listener*/)
{
	WPS_DEBUG_MSG(("MultiplanInternal::Cell::send: must not be called\n"));
	return false;
}
//! a data cell zone
struct Zone
{
	//! the different enum type
	enum Type { Zone0=0, Link, FileName, SharedData, Name,
	            CellData, CellPosition, Undef
	          };
	//! constructor
	Zone()
		: m_entry()
		, m_positionsSet()
	{
	}
	//! returns true if the zone entry is valid
	bool isValid() const
	{
		return m_entry.valid();
	}
	//! returns the name corresponding to a type
	static std::string getName(int type)
	{
		char const *wh[]= { "Zone0", "Link", "FileName", "SharedData", "Names",
		                    "CellData", "CellPos", "UndefZone"
		                  };
		if (type<0 || type>=int(WPS_N_ELEMENTS(wh)))
		{
			WPS_DEBUG_MSG(("MultiplanInternal::Zone::getName: called with unexpected type=%d\n", type));
			return "UndefZone";
		}
		return wh[type];
	}
	//! the zone data
	WPSEntry m_entry;
	//! list of position in the zone
	std::set<int> m_positionsSet;
};
//! the state of MultiplanParser
struct State
{
	//! constructor
	State(libwps_tools_win::Font::Type fontType, char const *password)
		: m_eof(-1)
		, m_version(-1)
		, m_fontType(fontType)
		, m_maximumCell()
		, m_columnWidths()
		, m_zonesList(1, Zone())
		, m_cellPositionsMap()
		, m_posToLinkMap()
		, m_posToNameMap()
		, m_posToSharedDataSeen()

		, m_password(password)
		, m_hash(0)
		, m_checksum(0)
	{
		for (auto &k : m_keys) k=0;
	}
	//! return the default font style
	libwps_tools_win::Font::Type getDefaultFontType() const
	{
		if (m_fontType != libwps_tools_win::Font::UNKNOWN)
			return m_fontType;
		return libwps_tools_win::Font::CP_437;
	}

	//! returns a default font (Courier12) with file's version to define the default encoding */
	static WPSFont getDefaultFont()
	{
		WPSFont res;
		res.m_name="Courier";
		res.m_size=12;
		return res;
	}
	//! returns the column width in point
	std::vector<WPSColumnFormat> getColumnsWidth() const;

	//! the last file position
	long m_eof;
	//! the file version
	int m_version;
	//! the user font type
	libwps_tools_win::Font::Type m_fontType;
	//! the maximumCell
	Vec2i m_maximumCell;
	//! the columns width in char, 0 means default
	std::vector<int> m_columnWidths;
	//! the list of cell's data zone
	std::vector<Zone> m_zonesList;
	//! the positions of each cell: a vector for each row
	std::map<int,std::vector<int> > m_cellPositionsMap;
	//! the different main spreadsheet zones
	WPSEntry m_entries[5];
	//! the list of link instruction
	std::map<int, WKSContentListener::FormulaInstruction> m_posToLinkMap;
	//! the map name's pos to name's cell instruction
	std::map<int, WKSContentListener::FormulaInstruction> m_posToNameMap;
	//! a set a shared data already seen
	std::set<int> m_posToSharedDataSeen;

	//! the password (if known)
	char const *m_password;
	//! the file hash
	int m_hash;
	//! the file checksum
	int m_checksum;
	//! the list of decoding keys
	uint8_t m_keys[16];
private:
	State(State const &) = delete;
	State &operator=(State const &) = delete;
};

std::vector<WPSColumnFormat> State::getColumnsWidth() const
{
	std::vector<WPSColumnFormat> res;
	WPSColumnFormat const defFormat(64);
	for (auto p : m_columnWidths)
	{
		if (p<0 || p>=255)
			res.push_back(defFormat);
		else
			res.push_back(WPSColumnFormat(float(8*p)));
	}
	if (res.size()<64) res.resize(64, defFormat);
	return res;
}
}

// constructor, destructor
MultiplanParser::MultiplanParser(RVNGInputStreamPtr &input, WPSHeaderPtr &header,
                                 libwps_tools_win::Font::Type encoding, char const *password)
	: WKSParser(input, header)
	, m_listener()
	, m_state(new MultiplanParserInternal::State(encoding, password))
{
}

MultiplanParser::~MultiplanParser()
{
}

int MultiplanParser::version() const
{
	return m_state->m_version;
}

bool MultiplanParser::checkFilePosition(long pos)
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

libwps_tools_win::Font::Type MultiplanParser::getDefaultFontType() const
{
	return m_state->getDefaultFontType();
}

////////////////////////////////////////////////////////////
// main function to parse the document
////////////////////////////////////////////////////////////
void MultiplanParser::parse(librevenge::RVNGSpreadsheetInterface *documentInterface)
{
	RVNGInputStreamPtr input=getInput();
	if (!input)
	{
		WPS_DEBUG_MSG(("MultiplanParser::parse: does not find input!!!\n"));
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
			m_listener->startDocument();
			sendSpreadsheet();
			m_listener->endDocument();
			ok = true;
		}
	}
	catch (libwps::PasswordException())
	{
		ascii().reset();
		WPS_DEBUG_MSG(("MultiplanParser::parse: password exception catched when parsing MN0\n"));
		throw (libwps::PasswordException());
	}
	catch (...)
	{
		WPS_DEBUG_MSG(("MultiplanParser::parse: exception catched when parsing MN0\n"));
		throw (libwps::ParseException());
	}

	m_listener.reset();
	ascii().reset();
	if (!ok)
		throw(libwps::ParseException());
}

std::shared_ptr<WKSContentListener> MultiplanParser::createListener(librevenge::RVNGSpreadsheetInterface *interface)
{
	std::vector<WPSPageSpan> pageList;
	WPSPageSpan ps;
	pageList.push_back(ps);
	return std::shared_ptr<WKSContentListener>(new WKSContentListener(pageList, interface));
}

////////////////////////////////////////////////////////////
// low level
////////////////////////////////////////////////////////////
// read the header
////////////////////////////////////////////////////////////
bool MultiplanParser::checkHeader(WPSHeader *header, bool strict)
{
	libwps::DebugStream f;

	RVNGInputStreamPtr input = getInput();
	if (!checkFilePosition(0x29a))
	{
		WPS_DEBUG_MSG(("MultiplanParser::checkHeader: file is too short\n"));
		return false;
	}

	input->seek(0,librevenge::RVNG_SEEK_SET);
	int fileSign=int(libwps::readU16(input));
	auto &vers=m_state->m_version;
	if (fileSign==0xe708)
		vers=1;
	else if (fileSign==0xec0c)
		vers=2;
	else if (fileSign==0xed0c)
		vers=3;
	else
		return false;
	f << "FileHeader:vers=" << vers << ",";
	if (vers==3)
	{
		input->seek(22, librevenge::RVNG_SEEK_SET);
		m_state->m_hash=int(libwps::readU16(input));
		m_state->m_checksum=int(libwps::readU16(input));
		if (m_state->m_hash || m_state->m_checksum)
		{
			WPS_DEBUG_MSG(("MultiplanParser::checkHeader: the file is protected with a password\n"));
		}
	}
	long const endHeader=vers== 1 ? 0xfa : 0x112;
	if (strict)
	{
		// read the spreadsheet dimension
		input->seek(endHeader, librevenge::RVNG_SEEK_SET);
		int dim[2];
		for (auto &d : dim) d=int(libwps::readU16(input));
		if (dim[0]==0 || dim[0]>=(vers==1 ? 256 : 4096) || dim[1]==0 || dim[1]>=256)
		{
			WPS_DEBUG_MSG(("MultiplanParser::checkHeader: the spreadsheet dimension is bad\n"));
			return false;
		}
		if (vers==1)
		{
			// read the last zone list position and check that it corresponds to a valid position
			input->seek(0x28a, librevenge::RVNG_SEEK_SET);
			int lastPos=0;
			for (int i=0; i<8; ++i)
			{
				int newPos=int(libwps::readU16(input));
				if (i==4) newPos+=lastPos; // length
				if (i==5)
				{
					lastPos=newPos;
					continue;
				}
				if (newPos<lastPos)
				{
					WPS_DEBUG_MSG(("MultiplanParser::checkHeader: find bad position\n"));
					return false;
				}
				if (i==1 && newPos-lastPos!=2*dim[0]*dim[1])
				{
					WPS_DEBUG_MSG(("MultiplanParser::checkHeader: the first zone size seems bad\n"));
					return false;
				}
				lastPos=newPos;
			}
			if (lastPos<4 || !checkFilePosition(0x29a+lastPos))
			{
				WPS_DEBUG_MSG(("MultiplanParser::checkHeader: can not find last spreadsheet position\n"));
				return false;
			}
		}
		else
		{
			// check if we can find the spreadsheet's data zone
			if (!checkFilePosition(0x3c7))
			{
				WPS_DEBUG_MSG(("MultiplanParser::checkHeader: can not find the data main position\n"));
				return false;
			}
			input->seek(0x3c5, librevenge::RVNG_SEEK_SET);
			long dataPos=long(libwps::readU16(input));
			if (dataPos<0x3c7 || !checkFilePosition(dataPos+20))
			{
				WPS_DEBUG_MSG(("MultiplanParser::checkHeader: the data main position seems bad\n"));
				return false;
			}
			input->seek(dataPos+2, librevenge::RVNG_SEEK_SET);
			int actType=6;
			// check if we can read the main spreadsheet zone
			while (!input->isEnd())
			{
				long pos=input->tell();
				bool ok=checkFilePosition(pos+6);
				int val=ok ? int(libwps::readU16(input)) : 0;
				if (val==6 && actType==12)   // row zone, ok if we have find some cell's data zones
				{
					input->seek(pos, librevenge::RVNG_SEEK_SET);
					break;
				}
				if (val<=12 && (val>actType || val==12))   // we can have main cell data's zone
				{
					actType=val;
					input->seek(2, librevenge::RVNG_SEEK_CUR);
					int dSz=int(libwps::readU16(input));
					if (dSz>=6 && checkFilePosition(pos+dSz))
					{
						input->seek(pos+dSz, librevenge::RVNG_SEEK_SET);
						continue;
					}
				}
				WPS_DEBUG_MSG(("MultiplanParser::checkHeader: can not read some zone\n"));
				return false;
			}
		}
	}
	ascii().addPos(0);
	ascii().addNote(f.str().c_str());
	input->seek(vers==1 ? 0x2 : 0x1a, librevenge::RVNG_SEEK_SET);
	for (int i=0; i<8; ++i)
	{
		long pos=input->tell();
		f.str("");
		f << "Entries(LinkFiles)[" << i << "]:";
		std::string file;
		for (int c=0; c<0x1f; ++c)
		{
			char ch=char(libwps::readU8(input));
			if (!ch) break;
			file += ch;
		}
		f << file;
		ascii().addPos(pos);
		if (file.empty())
			ascii().addNote("_");
		else
			ascii().addNote(f.str().c_str());
		input->seek(pos+0x1f, librevenge::RVNG_SEEK_SET);
	}
	if (header)
	{
		header->setMajorVersion(uint8_t(vers));
		header->setCreator(libwps::WPS_MULTIPLAN);
		header->setKind(libwps::WPS_SPREADSHEET);
		header->setNeedEncoding(true);
		header->setIsEncrypted(m_state->m_hash || m_state->m_checksum);
	}
	return true;
}

bool MultiplanParser::readZones()
{
	int const vers=version();
	if (!readZoneB()) return false;
	auto input = getInput();
	long pos=input->tell();
	libwps::DebugStream f;
	long const zoneCSz=vers==1 ? 22 : 28;
	if (!checkFilePosition(pos+zoneCSz*8))
	{
		WPS_DEBUG_MSG(("MultiplanParser::readZones: can not read zone C\n"));
		ascii().addPos(pos);
		ascii().addNote("Entries(Unknown):###extra");
		return false;
	}
	ascii().addPos(pos);
	ascii().addNote("Entries(ZoneC):");
	for (int i=0; i<8; ++i)
	{
		pos=input->tell();
		f.str("");
		f << "ZoneC" << i << ":";
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		input->seek(pos+zoneCSz, librevenge::RVNG_SEEK_SET);
	}
	if (!readColumnsWidth())
		return false;
	pos=input->tell();
	if (!checkFilePosition(pos+29))
	{
		WPS_DEBUG_MSG(("MultiplanParser::readZones: can not read zone D\n"));
		ascii().addPos(pos);
		ascii().addNote("Entries(Unknown):###extra");
		return false;
	}
	f.str("");
	f << "Entries(ZoneD):";
	if (m_state->m_version==1)
	{
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		input->seek(pos+29, librevenge::RVNG_SEEK_SET);
	}
	else
	{
		input->seek(pos+27, librevenge::RVNG_SEEK_SET);
		long newPos=int(libwps::readU16(input));
		f << "pos=" << std::hex << newPos << std::dec << ",";
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		if (newPos<input->tell() || !checkFilePosition(newPos))
		{
			WPS_DEBUG_MSG(("MultiplanParser::readZones: bad position\n"));
			ascii().addPos(pos);
			ascii().addNote("###");
			return false;
		}
		while (!input->isEnd())
		{
			pos=input->tell();
			if (pos>=newPos) break;
			ascii().addPos(pos);
			ascii().addNote("Entries(ZoneD):");
			input->seek(pos+50, librevenge::RVNG_SEEK_SET);
		}
		input->seek(newPos, librevenge::RVNG_SEEK_SET);
		if (!readZonesListV2())
			return false;
		input=getInput();
		if (vers==2)
			readFunctionNamesList();
		else if (checkFilePosition(input->tell()+12*92))
		{
			// unsure, seems like a list of block of size 12 mixed
			// with other block of different sizes...
			//
			// this zone probably contains many junk data
			ascii().addPos(input->tell());
			ascii().addNote("Entries(ZoneE):");
			for (int i=0; i<93; ++i)
			{
				pos=input->tell();
				f.str("");
				f << "ZoneE" << i << ":";
				ascii().addPos(pos);
				ascii().addNote(f.str().c_str());
				input->seek(pos+12, librevenge::RVNG_SEEK_SET);
			}
		}
		if (!input->isEnd())
		{
			ascii().addPos(input->tell());
			ascii().addNote("Entries(Unknown):###extra");
			WPS_DEBUG_MSG(("MultiplanParser::readZones: find extra data\n"));
		}
		return !m_state->m_cellPositionsMap.empty();
	}
	if (!readZonesList())
		return false;
	if (!input->isEnd())
	{
		WPS_DEBUG_MSG(("MultiplanParser::readZones: find extra data\n"));
		ascii().addPos(input->tell());
		ascii().addNote("Entries(Unknown):###extra");
	}
	return true;
}

bool MultiplanParser::readColumnsWidth()
{
	auto input = getInput();
	long pos = input->tell();
	int const vers=version();
	int const numCols=vers==1 ? 63 : 255;
	if (!checkFilePosition(pos+numCols))
	{
		WPS_DEBUG_MSG(("MultiplanParser::readColumnsWidth: the zone seems too short\n"));
		return false;
	}
	libwps::DebugStream f;
	f << "Entries(ColWidth):width=[";
	for (int i=0; i<numCols; ++i)
	{
		int val=int(libwps::readU8(input));
		if (vers<=2 && val==0) val=255;
		m_state->m_columnWidths.push_back(val);
		if (val!=255)
			f << m_state->m_columnWidths.back() << ",";
		else
			f << "_,";
	}
	f << "],";
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

bool MultiplanParser::readZonesList()
{
	auto input = getInput();
	long pos = input->tell();
	if (!checkFilePosition(pos+16))
	{
		WPS_DEBUG_MSG(("MultiplanParser::readZonesList: the zone seems too short\n"));
		return false;
	}
	libwps::DebugStream f;
	f << "Entries(ZonesList):";
	int lastPos=0;
	f << "zones=[";
	WPSEntry cellPosEntry;
	for (int i=0; i<8; ++i)
	{
		int newPos=int(libwps::readU16(input));
		if (i==4) newPos+=lastPos; // length
		if (i==5)
		{
			lastPos=newPos;
			continue;
		}
		if (newPos>lastPos)
		{
			if (!checkFilePosition(pos+16+newPos))
			{
				WPS_DEBUG_MSG(("MultiplanParser::readZonesList: find a bad position"));
				f << "###";
			}
			else
			{
				WPSEntry entry;
				entry.setBegin(pos+16+lastPos);
				entry.setEnd(pos+16+newPos);
				MultiplanParserInternal::Zone::Type const what[]=
				{
					MultiplanParserInternal::Zone::Zone0,
					MultiplanParserInternal::Zone::CellPosition,
					MultiplanParserInternal::Zone::Link,
					MultiplanParserInternal::Zone::FileName,
					MultiplanParserInternal::Zone::CellData,
					MultiplanParserInternal::Zone::Undef, // beg of shared data
					MultiplanParserInternal::Zone::SharedData,
					MultiplanParserInternal::Zone::Name,
				};
				auto w=what[i];
				if (w==MultiplanParserInternal::Zone::CellData)
				{
					m_state->m_zonesList[0].m_entry=entry;
					ascii().addPos(pos+16+lastPos);
					ascii().addNote("Entries(CellData)");
				}
				else if (w==MultiplanParserInternal::Zone::CellPosition)
					cellPosEntry=entry;
				else
					m_state->m_entries[w]=entry;
			}
			f << std::hex << lastPos << "<->" << newPos << std::dec << ",";
			lastPos=newPos;
		}
		else
			f << "_,";
	}
	f << "],";
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	bool ok=readCellDataPosition(cellPosEntry);
	for (int i=0; i<5; ++i)
	{
		if (!m_state->m_entries[i].valid()) continue;
		f.str("");
		f << "Entries(" << MultiplanParserInternal::Zone::getName(i) << "):";
		ascii().addPos(m_state->m_entries[i].begin());
		ascii().addNote(f.str().c_str());
		ascii().addPos(m_state->m_entries[i].end());
		ascii().addNote("_");
		input->seek(m_state->m_entries[i].end(), librevenge::RVNG_SEEK_SET);
	}
	return ok;
}

bool MultiplanParser::readZonesListV2()
{
	auto input = getInput();
	long pos = input->tell();
	if (!checkFilePosition(pos+2+5*6))
	{
		WPS_DEBUG_MSG(("MultiplanParser::readZonesListV2: the zone seems too short\n"));
		return false;
	}
	libwps::DebugStream f;
	f << "Entries(ZonesList):";
	int N=int(libwps::readU16(input));
	f << "N[row]=" << N << ",";
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	// time to decode the different crypted zones
	if (m_state->m_hash || m_state->m_checksum)
	{
		if (!m_state->m_password || !checkPassword(m_state->m_password))
		{
			if (!retrievePasswordKeys())
			{
				WPS_DEBUG_MSG(("MultiplanParser::readZonesListV2: can not find the password to decode data\n"));
				throw (libwps::PasswordException());
			}
		}
		input->seek(pos+2, librevenge::RVNG_SEEK_SET);
		auto newInput=decodeStream(input);
		if (!newInput) throw (libwps::ParseException());
		input=newInput;
		resetInput(newInput);
		ascii().setStream(newInput);
	}

	for (int i=0; i<6; ++i)
	{
		pos=input->tell();
		if (!checkFilePosition(pos+6))
		{
			WPS_DEBUG_MSG(("MultiplanParser::readZonesListV2: can not find zone %d\n", i));
			return false;
		}
		int val=int(libwps::readU16(input));
		if (val<7 || val>12)
		{
			input->seek(pos, librevenge::RVNG_SEEK_SET);
			break;
		}
		int type=val-7;
		MultiplanParserInternal::Zone::Type const what[]=
		{
			MultiplanParserInternal::Zone::Zone0,
			MultiplanParserInternal::Zone::Link,
			MultiplanParserInternal::Zone::FileName,
			MultiplanParserInternal::Zone::Name,
			MultiplanParserInternal::Zone::SharedData,
			MultiplanParserInternal::Zone::CellData
		};
		auto wh=what[type];
		auto &entry=wh==MultiplanParserInternal::Zone::CellData ?
		            m_state->m_zonesList[0].m_entry : m_state->m_entries[wh];
		if (entry.valid())
		{
			input->seek(pos, librevenge::RVNG_SEEK_SET);
			break;
		}
		f.str("");
		f << "Entries(" << MultiplanParserInternal::Zone::getName(wh) << "):";
		val=int(libwps::readU16(input));
		f << "f0=" << std::hex << val << std::dec << ",";
		int dSz=int(libwps::readU16(input));
		if (dSz<6 || !checkFilePosition(pos+dSz))
		{
			WPS_DEBUG_MSG(("MultiplanParser::readZonesListV2: can not find zone %d\n", i));
			return false;
		}
		entry.setBegin(pos);
		entry.setLength(dSz);
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		input->seek(pos+dSz, librevenge::RVNG_SEEK_SET);
		if (type==5) break;
	}
	while (checkFilePosition(input->tell()+6))
	{
		pos=input->tell();
		if (libwps::readU16(input)!=0xc)
		{
			input->seek(pos, librevenge::RVNG_SEEK_SET);
			break;
		}
		f.str("");
		f << "Entries(CellData)[extra]:";
		int val=int(libwps::readU16(input));
		f << "f0=" << std::hex << val << std::dec << ",";
		int dSz=int(libwps::readU16(input));
		if (dSz<6 || !checkFilePosition(pos+dSz))
		{
			WPS_DEBUG_MSG(("MultiplanParser::readZonesListV2: can not find extra data cell zone\n"));
			return false;
		}
		m_state->m_zonesList.push_back(MultiplanParserInternal::Zone());
		auto &entry=m_state->m_zonesList.back().m_entry;
		entry.setBegin(pos);
		entry.setLength(dSz);
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		input->seek(pos+dSz, librevenge::RVNG_SEEK_SET);
	}
	pos=input->tell();
	if (!checkFilePosition(pos+2)||int(libwps::readU16(input))!=6)
	{
		WPS_DEBUG_MSG(("MultiplanParser::readZonesListV2: can not find row header\n"));
		return false;
	}
	f.str("");
	f << "Entries(Row):";
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	for (int i=0; i<N; ++i)
	{
		pos=input->tell();
		if (!checkFilePosition(pos+6))
		{
			WPS_DEBUG_MSG(("MultiplanParser::readZonesListV2: can not find row %d\n", i));
			return false;
		}
		f.str("");
		f << "Row" << i << ":";
		int val=int(libwps::readU16(input));
		if (val!=6) f << "f0=" << val << ",";
		int dSz=int(libwps::readU16(input));
		if (dSz<4 || !checkFilePosition(pos+2*dSz))
		{
			WPS_DEBUG_MSG(("MultiplanParser::readZonesListV2: can not find row %d\n", i));
			return false;
		}
		int num=int(libwps::readU16(input));
		if (8+3*num>2*dSz)
		{
			WPS_DEBUG_MSG(("MultiplanParser::readZonesListV2: can not find row %d\n", i));
			return false;
		}
		int r=int(libwps::readU16(input));
		if (m_state->m_cellPositionsMap.find(r)!=m_state->m_cellPositionsMap.end())
		{
			WPS_DEBUG_MSG(("MultiplanParser::readZonesListV2: oops, row %d already exists\n", r));
			return !m_state->m_cellPositionsMap.empty();
		}
		f << "row=" << r << ",";
		std::vector<int> cellPos;
		f << "data=[";
		for (int d=0; d<num; ++d)
		{
			val=int(libwps::readU16(input));
			int z=int(libwps::readU8(input));
			val+=0x10000*z;
			if (z >= int(m_state->m_zonesList.size()))
			{
				WPS_DEBUG_MSG(("MultiplanParser::readZonesListV2: oops, find bad cell pos for row%d\n", r));
				f << "##" << std::hex << val << std::dec << ",";
				cellPos.push_back(0);
				continue;
			}
			cellPos.push_back(val);
			if (val==0)
			{
				f << "_,";
				continue;
			}
			f << std::hex << val << std::dec << ",";
			m_state->m_zonesList[size_t(z)].m_positionsSet.insert(val&0xffff);
		}
		m_state->m_cellPositionsMap[r]=cellPos;
		if (input->tell()!=pos+2*dSz)
			ascii().addDelimiter(input->tell(),'|');
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		input->seek(pos+2*dSz, librevenge::RVNG_SEEK_SET);
	}
	return true;
}

bool MultiplanParser::readFunctionNamesList()
{
	auto input = getInput();
	if (input->isEnd()) return true;
	long pos=input->tell();
	libwps::DebugStream f;
	f << "Entries(NamFunctions):";
	while (!input->isEnd())
	{
		long actPos=input->tell();
		int cSz=int(libwps::readU8(input));
		if (cSz==0 || !checkFilePosition(actPos+1+cSz))
		{
			input->seek(actPos, librevenge::RVNG_SEEK_SET);
			break;
		}
		std::string name;
		for (int i=0; i<cSz; ++i) name+=char(libwps::readU8(input));
		f << name << ",";
	}
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

////////////////////////////////////////////////////////////
// double
////////////////////////////////////////////////////////////
bool MultiplanParser::readDouble(double &value)
{
	auto input = getInput();
	long pos=input->tell();
	value=0;
	if (!checkFilePosition(pos+8))
	{
		WPS_DEBUG_MSG(("MultiplanParser::readDouble: the zone is too short\n"));
		return false;
	}
	if (version()>=2)
	{
		bool isNan;
		if (!libwps::readDouble8(input,value,isNan))
		{
			value=0;
			input->seek(pos+8, librevenge::RVNG_SEEK_SET);
		}
		return true;
	}
	int exponant=int(libwps::readU8(input));
	double sign=1;
	if (exponant&0x80)
	{
		exponant&=0x7f;
		sign=-1;
	}
	bool ok=true;
	double factor=1;
	for (int i=0; i<7; ++i)
	{
		int val=int(libwps::readU8(input));
		for (int d=0; d<2; ++d)
		{
			int v= d==0 ? (val>>4) : (val&0xf);
			if (v>=10)
			{
				WPS_DEBUG_MSG(("MultiplanParser::readDouble: oops find a bad digits\n"));
				ok=false;
				break;
			}
			factor/=10.;
			value+=factor*v;
		}
		if (!ok) break;
	}
	value *= sign*std::pow(10.,exponant-0x40);
	input->seek(pos+8, librevenge::RVNG_SEEK_SET);
	return ok;
}

////////////////////////////////////////////////////////////
// formula
////////////////////////////////////////////////////////////
namespace MultiplanParserInternal
{
struct Functions
{
	char const *m_name;
	int m_arity;
};

static Functions const s_listOperators[] =
{
	// 0
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	// 1
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	// 2
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { ":", 2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { ":", 2}, { "", -2}, { "", -2},
	// 3
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	// 4
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { ":", 2}, { "", -2}, { "", -2},
	// 5
	{ "&", 2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	// 6
	{ "<", 2}, { "", -2}, { "<=", 2}, { "", -2},
	{ "=", 2}, { "", -2}, { ">=", 2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	// 7
	{ ">", 2}, { "", -2}, { "<>", 2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	// 8
	{ "", -2}, { "", -2}, { "+", 2}, { "", -2},
	{ "-", 2}, { "", -2}, { "*", 2}, { "", -2},
	{ "/", 2}, { "", -2}, { "^", 2}, { "", -2},
	{ "+", 1}, { "", -2}, { "-", 1}, { "", -2},
	// 9
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "%", 1}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
};

static char const *s_listFunctions[]=
{
	// 0
	"Count", "If", "IsNA", "IsError",
	"Sum", "Average", "Min", "Max",
	"Row", "Column", "NA", "NPV",
	"Stdev", "Dollar", "Fixed", "Sin",
	// 1
	"Cos", "Tan", "Atan", "Pi",
	"Sqrt", "Exp", "Ln", "Log",
	"Abs", "Int", "Sign", "Round",
	"Lookup", "Index", "Rept", "Mid",
	// 2
	"Len", "Value", "True", "False",
	"And", "Or", "Not", "Mod",
	// mac v1 new functions
	"IterCnt", "Delta", "PV", "FV",
	"NPer", "PMT", "Rate", "MIRR",
	// 3
	"Irr", nullptr, nullptr, nullptr,
	nullptr, nullptr, nullptr, nullptr,
	nullptr, nullptr, nullptr, nullptr,
	nullptr, nullptr, nullptr, nullptr,
};
}

bool MultiplanParser::readFormula(Vec2i const &cellPos, std::vector<WKSContentListener::FormulaInstruction> &formula, long endPos, std::string &error)
{
	formula.clear();
	auto input = getInput();
	std::vector<std::vector<WKSContentListener::FormulaInstruction> > stack;
	auto const &listOperators=MultiplanParserInternal::s_listOperators;
	int const numOperators=int(WPS_N_ELEMENTS(listOperators));
	auto const &listFunctions=MultiplanParserInternal::s_listFunctions;
	int const numFunctions=int(WPS_N_ELEMENTS(listFunctions));
	bool ok=true;
	int closeDelayed=0;
	bool checkForClose=false;
	while (input->tell()<=endPos)
	{
		long pos=input->tell();
		int wh=pos==endPos ? -1 : int(libwps::readU8(input));
		bool needCloseParenthesis=closeDelayed && (checkForClose || pos==endPos);
		ok=true;
		if (closeDelayed && !needCloseParenthesis && wh!=0x3c)
			needCloseParenthesis=wh>=numOperators || listOperators[wh].m_arity!=2;
		while (needCloseParenthesis && closeDelayed>0)
		{
			auto len=stack.size();
			if (len<2)
			{
				error="##closedParenthesis,";
				ok=false;
				break;
			}
			auto &dParenthesisFunc=stack[len-2];
			if (dParenthesisFunc.size()!=1 || dParenthesisFunc[0].m_type!=dParenthesisFunc[0].F_Operator ||
			        dParenthesisFunc[0].m_content!="(")
			{
				error="##closedParenthesis,";
				ok=false;
				break;
			}
			dParenthesisFunc.insert(dParenthesisFunc.end(),stack.back().begin(), stack.back().end());
			WKSContentListener::FormulaInstruction instr;
			instr.m_type=instr.F_Operator;
			instr.m_content=")";
			dParenthesisFunc.push_back(instr);
			stack.resize(len-1);
			--closeDelayed;
		}
		if (!ok || pos==endPos)
			break;
		int arity=0;
		WKSContentListener::FormulaInstruction instr;
		ok=false;
		bool noneInstr=false, closeFunction=false;
		switch (wh)
		{
		case 0:
		{
			if (pos+3>endPos)
				break;
			ok=readLink(int(libwps::readU16(input)), instr);
			break;
		}
		case 0x12:
		{
			if (pos+2>endPos)
				break;
			ok=true;
			instr.m_type=instr.F_Function;
			int id=int(libwps::readU8(input));
			if (id<numFunctions && listFunctions[id])
				instr.m_content=listFunctions[id];
			else
			{
				std::stringstream s;
				s << "Funct" << std::hex << id << std::dec;
				instr.m_content=s.str();
			}
			std::vector<WKSContentListener::FormulaInstruction> child;
			child.push_back(instr);
			stack.push_back(child);
			instr.m_type=instr.F_Operator;
			instr.m_content="(";
			break;
		}
		case 0x67:
		case 0x87:
		case 0xc7:
			closeFunction=ok=true;
			break;
		case 0x1c: // use before %
		case 0x1e: // use for <> A 1e B "code <>"
		case 0x34: // use for <=,>= ... A 34 B "code <=,..."
		case 0x36: // use before -unary
		case 0x38: // use before +unary
			noneInstr=ok=true;
			break;
		case 0x3a:
			ok=true;
			instr.m_type=instr.F_Operator;
			instr.m_content=";";
			break;
		case 0x3c:
			noneInstr=ok=true;
			++closeDelayed;
			break;
		case 0x3e:
			ok=true;
			instr.m_type=instr.F_Operator;
			instr.m_content="(";
			break;
		case 0x56:
		{
			int dSz=int(libwps::readU8(input));
			if (pos+2+dSz>endPos)
				break;
			instr.m_type=instr.F_Text;
			instr.m_content=libwps_tools_win::Font::unicodeString(input.get(), unsigned(dSz), m_state->getDefaultFontType()).cstr();
			ok=true;
			break;
		}
		case 0x2d:
		case 0xed: // example C2
			if (pos+5>endPos)
				break;
			WPS_DEBUG_MSG(("MultiplanParser::readFormula: find column/row solitary definition\n"));
			error="###RorC";
			ok=false;
			break;
		case 0xe1: // example C2 R1
			WPS_DEBUG_MSG(("MultiplanParser::readFormula: find union operator\n"));
			error="###union";
			ok=false;
			break;
		case 0x37: // use for list cell
		case 0x53:
		case 0x73:
		case 0x93: // basic cell
		case 0xf3:
		{
			if (pos+3>endPos)
				break;
			instr.m_type=instr.F_Cell;
			instr.m_positionRelative[0]=Vec2b(true,true);
			int val=int(libwps::readU16(input));
			auto &newPos=instr.m_position[0];
			if (val&0x8000)
				newPos[1]=cellPos[1]-(val&0xff);
			else
				newPos[1]=cellPos[1]+(val&0xff);
			if (val&0x4000)
				newPos[0]=cellPos[0]-((val>>8)&0x3f);
			else
				newPos[0]=cellPos[0]+((val>>8)&0x3f);
			ok=newPos[0]>=0 && newPos[1]>=0;
			break;
		}
		case 0x94:
			if (pos+9>endPos || !readDouble(instr.m_doubleValue))
				break;
			instr.m_type=instr.F_Double;
			ok=true;
			break;
		case 0x35:
		case 0x8f:
		case 0xef:
			if (pos+3>endPos)
				break;
			instr.m_type=instr.F_Cell;
			instr.m_positionRelative[0]=Vec2b(false,false);
			instr.m_position[0][1]=int(libwps::readU8(input));
			instr.m_position[0][0]=int(libwps::readU8(input));
			ok=(instr.m_position[0][0]<63) && (instr.m_position[0][1]<255);
			if (!ok)
			{
				error="###RorC";
				WPS_DEBUG_MSG(("MultiplanParser::readFormula: find only row/column reference\n"));
			}
			break;
		case 0xeb:
			if (pos+3>endPos || !readName(int(libwps::readU16(input)), instr))
				break;
			ok=true;
			break;
		default:
			if (wh<numOperators && listOperators[wh].m_arity!=-2)
			{
				instr.m_content=listOperators[wh].m_name;
				instr.m_type=instr.F_Function;
				arity=listOperators[wh].m_arity;
			}
			if (instr.m_content.empty())
			{
				WPS_DEBUG_MSG(("MultiplanParser::readFormula: find unknown type %x\n", wh));
				std::stringstream s;
				s << "##unkn[func]=" << std::hex << wh << std::dec << ",";
				error=s.str();
				break;
			}
			ok=true;
			break;
		}
		if (!ok)
		{
			input->seek(pos, librevenge::RVNG_SEEK_SET);
			break;
		}
		checkForClose=!noneInstr && closeDelayed>0;
		if (noneInstr) continue;
		if (closeFunction)
		{
			ok=false;
			if (stack.empty())
			{
				error="##closed,";
				break;
			}
			auto it=stack.end();
			--it;
			for (; it!=stack.begin(); --it)
			{
				if (it->size()!=1) continue;
				auto const &dInstr=(*it)[0];
				if (dInstr.m_type!=dInstr.F_Operator || dInstr.m_content!="(") continue;
				auto fIt=it;
				--fIt;
				auto &functionStack=*fIt;
				if (functionStack.size()!=1 || functionStack[0].m_type!=functionStack[0].F_Function) continue;
				ok=true;
				for (; it!=stack.end(); ++it)
					functionStack.insert(functionStack.end(), it->begin(), it->end());
				++fIt;
				stack.erase(fIt, stack.end());
				break;
			}
			if (!ok)
			{
				error="##closed";
				break;
			}
			instr.m_type=instr.F_Operator;
			instr.m_content=")";
			stack.back().push_back(instr);
			continue;
		}
		std::vector<WKSContentListener::FormulaInstruction> child;
		if (instr.m_type!=WKSContentListener::FormulaInstruction::F_Function)
		{
			child.push_back(instr);
			stack.push_back(child);
			continue;
		}
		size_t numElt = stack.size();
		if (static_cast<int>(numElt) < arity)
		{
			std::stringstream s;
			s << instr.m_content << "[##" << arity << "]";
			error=s.str();
			input->seek(pos, librevenge::RVNG_SEEK_SET);
			ok=false;
			break;
		}
		if (arity==1)
		{
			instr.m_type=WKSContentListener::FormulaInstruction::F_Operator;
			if (instr.m_content=="%")
				stack[numElt-1].push_back(instr);
			else
				stack[numElt-1].insert(stack[numElt-1].begin(), instr);
			continue;
		}
		if (arity==2)
		{
			instr.m_type=WKSContentListener::FormulaInstruction::F_Operator;
			stack[numElt-2].push_back(instr);
			stack[numElt-2].insert(stack[numElt-2].end(), stack[numElt-1].begin(), stack[numElt-1].end());
			stack.resize(numElt-1);
			continue;
		}
		ok=false;
		error = "### unexpected arity";
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		break;
	}
	long pos=input->tell();
	if (pos!=endPos || !ok || closeDelayed || stack.size()!=1 || stack[0].empty())
	{
		WPS_DEBUG_MSG(("MultiplanParser::readFormula: can not read a formula\n"));
		ascii().addDelimiter(pos, '|');
		input->seek(endPos, librevenge::RVNG_SEEK_SET);

		std::stringstream s;
		if (!error.empty())
			s << error;
		else
			s << "##unknownError";
		s << "[";
		for (auto const &i : stack)
		{
			for (auto const &j : i)
				s << j << ",";
		}
		s << "],";
		error=s.str();
		return true;
	}
	formula=stack[0];
	return true;
}

namespace MultiplanParserInternal
{
static Functions const s_listOperatorsV2[] =
{
	// 0
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "-", 1}, { "+", 1}, { "%", 1},
	// 1
	{ "", -2}, { "+", 2}, { "-", 2}, { "*", 2},
	{ "/", 2}, { "^", 2}, { "<", 2}, { ">", 2},
	{ "<=", 2}, { ">=", 2}, { "=", 2}, { "<>", 2},
	{ "&", 2}, { "", -2}, { "", -2}, { "", -2},
	// 2
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
	{ "", -2}, { "", -2}, { "", -2}, { "", -2},
};

static char const *s_listFunctionsV2[]=
{
	// 0
	nullptr /* special Row */, nullptr/* special Column*/, nullptr/* checkme: unknown*/, nullptr/* checkme: unknown*/,
	nullptr /* special Index */, nullptr/* special LookUp*/, "Count", "If",
	"IsNa", "IsError", "Sum", "Average",
	"Min", "Max", "NA", "NPV",
	// 1
	"StDev", "Dollar", "Fixed", "Sin",
	"Cos", "Tan", "Atan", "Pi",
	"Sqrt", "Exp", "Ln", "Log10",
	"Abs", "Int", "Sign", "Round",
	// 2
	"Rept", "Mid", "Len", "Value",
	"True", "False", "And", "Or",
	"Not", "Mod", "PV", "NV",
	"NPER", "PMT", "Rate", "Mirr",
	// 3
	"Irr", nullptr /* special Now */, nullptr /* checkme: unknown */, "Date",
	"Time", "Day", "Month", "WeekDay",
	"Year", "Hour", "Minute", "Second",
	nullptr, nullptr, nullptr, nullptr,
};
}

bool MultiplanParser::readFormulaV2(Vec2i const &cellPos, std::vector<WKSContentListener::FormulaInstruction> &formula, long endZonePos, std::string &error)
{
	formula.clear();
	auto input = getInput();
	long pos=input->tell();
	int dSz=int(libwps::readU8(input));
	long endPos=pos+2*dSz;
	if (dSz==0 || endPos>endZonePos)
	{
		WPS_DEBUG_MSG(("MultiplanParser::readFormulaV2: the zone seems bad\n"));
		return false;
	}
	if (endZonePos!=endPos)
		ascii().addDelimiter(endPos, '|');
	long dataPos=-1;
	long dataSize=0;
	if (endZonePos>endPos+2)
	{
		input->seek(endPos+1, librevenge::RVNG_SEEK_SET);
		dSz=int(libwps::readU8(input));
		if (dSz && endPos+2+2*dSz<=endZonePos)
		{
			dataSize=dSz;
			dataPos=endPos+2;
		}
	}
	std::vector<std::vector<WKSContentListener::FormulaInstruction> > stack;
	auto const &listOperators=MultiplanParserInternal::s_listOperatorsV2;
	int const numOperators=int(WPS_N_ELEMENTS(listOperators));
	auto const &listFunctions=MultiplanParserInternal::s_listFunctionsV2;
	int const numFunctions=int(WPS_N_ELEMENTS(listFunctions));
	bool ok=true;
	input->seek(pos+1, librevenge::RVNG_SEEK_SET);
	while (input->tell()<endPos)
	{
		pos=input->tell();
		int code=int(libwps::readU8(input));
		if (code==0) break;
		WKSContentListener::FormulaInstruction instr;
		ok=false;
		int arity=0, cellId=-1;
		bool noneInstr=false;
		switch (code)
		{
		case 1:
		{
			if (pos+1>endPos)
				break;
			dSz=int(libwps::readU8(input));
			if (dSz==0 || pos+1+dSz>endPos)
				break;
			double mantisse = 0;
			for (int i = 0; i < dSz-2; i++)
				mantisse = mantisse/256 + double(readU8(input));
			auto mantExp = dSz>=2 ? int(readU8(input)) : 0;
			mantisse = (mantisse/256 + double(0x10+(mantExp & 0x0F)))/16;
			int exp = ((mantExp&0xF0)>>4)+int(readU8(input)<<4);
			int sign = 1;
			if (exp & 0x800)
			{
				exp &= 0x7ff;
				sign = -1;
			}
			exp -= 0x3ff;
			instr.m_type=instr.F_Double;
			instr.m_doubleValue = sign * std::ldexp(mantisse, exp);
			ok=true;
			break;
		}
		case 2:
		{
			if (pos+1>endPos)
				break;
			dSz=int(libwps::readU8(input));
			if (pos+1+dSz>endPos)
				break;
			instr.m_type=instr.F_Text;
			instr.m_content=libwps_tools_win::Font::unicodeString(input.get(), unsigned(dSz), m_state->getDefaultFontType()).cstr();
			ok=true;
			break;
		}
		case 0x4: // cell in function
		case 0x6: // cell in basic formula
			if (pos+2>endPos)
				break;
			cellId=libwps::readU8(input);
			break;
		case 0x7: // function without argument
		case 0x8: // function with argument
		case 0x9: // special
		case 0xa:   // special with arg
		{
			if (pos+2>endPos)
				break;
			instr.m_type=instr.F_Function;
			int id=int(libwps::readU8(input));
			if (code==9 || code==0xa)
			{
				int fId=-1;
				if (id+1<=dataSize)
				{
					input->seek(dataPos+2*id, librevenge::RVNG_SEEK_SET);
					int val=int(libwps::readU8(input));
					if (val==6) fId=int(libwps::readU8(input));
					input->seek(pos+2, librevenge::RVNG_SEEK_SET);
				}
				char const *wh[]= { "Row", "Column", nullptr, nullptr, "Index", "LookUp" };
				if (fId>=0 && fId<int(WPS_N_ELEMENTS(wh)) && wh[fId])
					instr.m_content=wh[fId];
				else if (fId==0x31)
					instr.m_content="Now";
				else
				{
					WPS_DEBUG_MSG(("MultiplanParser::readFormulaV2: can not find a function id\n"));
					error += "###fId";
					std::stringstream s;
					s << "FunctId" << fId;
					instr.m_content=s.str();
				}
			}
			else if (id<numFunctions && listFunctions[id])
				instr.m_content=listFunctions[id];
			else
			{
				WPS_DEBUG_MSG(("MultiplanParser::readFormulaV2: can not find a function %d\n", id));
				error += "###fId";
			}
			if (instr.m_content.empty())
			{
				std::stringstream s;
				s << "Funct" << std::hex << id << std::dec;
				instr.m_content=s.str();
			}
			std::vector<WKSContentListener::FormulaInstruction> child;
			child.push_back(instr);
			instr.m_type=instr.F_Operator;
			instr.m_content="(";
			if ((code%2)==0)
			{
				stack.push_back(child);
				ok=true;
				break;
			}
			child.push_back(instr);
			instr.m_content=")";
			child.push_back(instr);
			stack.push_back(child);
			ok=noneInstr=true;
			break;
		}
		case 0xb: // close 0x8: 1 arg
		case 0xc:   // close 0x8: >=2 arg
		{
			if (code==0xc)   // ignore the num arg
			{
				if (pos+2>endPos)
					break;
				input->seek(1, librevenge::RVNG_SEEK_CUR);
			}
			if (stack.empty())
			{
				error="##closed,";
				break;
			}
			auto it=stack.end();
			--it;
			for (; it!=stack.begin(); --it)
			{
				if (it->size()!=1) continue;
				auto const &dInstr=(*it)[0];
				if (dInstr.m_type!=dInstr.F_Operator || dInstr.m_content!="(") continue;
				auto fIt=it;
				--fIt;
				auto &functionStack=*fIt;
				if (functionStack.size()!=1 || functionStack[0].m_type!=functionStack[0].F_Function) continue;
				ok=true;
				for (; it!=stack.end(); ++it)
					functionStack.insert(functionStack.end(), it->begin(), it->end());
				++fIt;
				stack.erase(fIt, stack.end());
				break;
			}
			if (!ok)
			{
				error="##closed";
				break;
			}
			instr.m_type=instr.F_Operator;
			instr.m_content=")";
			stack.back().push_back(instr);
			noneInstr=true;
			break;
		}
		case 0x21:
			instr.m_type=instr.F_Operator;
			instr.m_content=";";
			ok=true;
			break;
		case 0x22:
			ok=true;
			arity=1;
			instr.m_type=instr.F_Function;
			instr.m_content=")";
			break;
		case 3: // begin of parenthesis
		case 0x23: // associate with -
		case 0x24: // associate with +
		case 0x25:
		case 0x26:
			noneInstr=ok=true;
			break;
		case 0x28:
			if (pos+3>endPos)
				break;
			ok=readLink(int(libwps::readU16(input)), instr);
			break;
		default:
		{
			if (code<numOperators && listOperators[code].m_arity!=-2)
			{
				instr.m_content=listOperators[code].m_name;
				instr.m_type=instr.F_Function;
				arity=listOperators[code].m_arity;
			}
			if (!instr.m_content.empty())
			{
				ok=true;
				break;
			}
			if (code>=0x80)
			{
				cellId=code-0x80;
				break;
			}
			WPS_DEBUG_MSG(("MultiplanParser::readFormulaV2: find unknown type %x\n", code));
			std::stringstream s;
			s << "##unkn[func]=" << std::hex << code << std::dec << ",";
			error=s.str();
			break;
		}
		}
		if (!ok && cellId>=0 && cellId+2<=dataSize)
		{
			long actPos=input->tell();
			input->seek(dataPos+2*cellId, librevenge::RVNG_SEEK_SET);
			int headerVal=int(libwps::readU8(input));
			int numCell=(headerVal&5)==5 ? 2 : 1;
			if (numCell==2 && actPos+6>=endPos)
			{
				WPS_DEBUG_MSG(("MultiplanParser::readFormulaV2: oops can not reserve extra space\n"));
				error="###extraSz,";
			}
			else
			{
				if (numCell==2)
					endPos-=6;
				int actCellId=2*cellId+1;
				if (numCell==2)   // small number maybe rel/abs
				{
					input->seek(1, librevenge::RVNG_SEEK_CUR);
					++actCellId;
				}
				instr.m_type=numCell==1 ? instr.F_Cell : instr.F_CellList;
				for (int c=0; c<numCell; ++c)
				{
					int val=headerVal;
					if (numCell==2)
					{
						val=int(libwps::readU8(input));
						++actCellId;
					}
					if ((val&3)==1 || (val&3)==2)   // 1: row, 2:col
					{
						error="##RorC";
						WPS_DEBUG_MSG(("MultiplanParser::readFormulaV2: find only row/column reference\n"));
						break;
					}
					if (val&4)
					{
						if (actCellId+3>2*dataSize) break;
						input->seek(1, librevenge::RVNG_SEEK_CUR);
						WKSContentListener::FormulaInstruction nameInstr;
						if (!readName(int(libwps::readU16(input)), nameInstr))
							break;
						if (nameInstr.m_type==nameInstr.F_Text)
							instr=nameInstr;
						else
						{
							for (int d=0; d+c<2; ++d)
							{
								if (d==1)
								{
									if (nameInstr.m_type!=nameInstr.F_CellList || numCell==2)
										break;
									instr.m_type=instr.F_CellList;
								}
								instr.m_position[c+d]=nameInstr.m_position[d];
								instr.m_positionRelative[c+d]=nameInstr.m_positionRelative[d];
							}
						}
						ok=c+1==numCell;
						continue;
					}
					if (actCellId+3>2*dataSize) break;
					int col=int(libwps::readU8(input));
					val=int(libwps::readU16(input));
					if (val&0x4000)
					{
						if (val&0x1000)
							instr.m_position[c][0]=cellPos[0]-col;
						else
							instr.m_position[c][0]=cellPos[0]+col;
						instr.m_positionRelative[c][0]=true;
					}
					else
					{
						instr.m_position[c][0]=col;
						instr.m_positionRelative[c][0]=false;
					}
					int row=val&0xfff;
					if (val&0x8000)
					{
						if (val&0x2000)
							instr.m_position[c][1]=cellPos[1]-row;
						else
							instr.m_position[c][1]=cellPos[1]+row;
						instr.m_positionRelative[c][1]=true;
					}
					else
					{
						instr.m_position[c][1]=row;
						instr.m_positionRelative[c][1]=false;
					}
					if (instr.m_position[c][0]<0 || instr.m_position[c][0]>=255 ||
					        instr.m_position[c][1]<0 || instr.m_position[c][1]>=4095)
						break;
					ok=c+1==numCell;
				}
			}
			input->seek(actPos, librevenge::RVNG_SEEK_SET);
			if (!ok)
			{
				WPS_DEBUG_MSG(("MultiplanParser::readFormulaV2: can not find a cell id[%d]\n", cellId));
				std::stringstream s;
				s << "###cell" << std::hex << code << std::dec;
				error+=s.str();
				break;
			}
		}
		if (!ok)
		{
			input->seek(pos,librevenge::RVNG_SEEK_SET);
			break;
		}
		if (noneInstr) continue;
		std::vector<WKSContentListener::FormulaInstruction> child;
		if (instr.m_type!=WKSContentListener::FormulaInstruction::F_Function)
		{
			child.push_back(instr);
			stack.push_back(child);
			continue;
		}
		size_t numElt = stack.size();
		if (static_cast<int>(numElt) < arity)
		{
			std::stringstream s;
			s << instr.m_content << "[##" << arity << "]";
			error=s.str();
			input->seek(pos, librevenge::RVNG_SEEK_SET);
			ok=false;
			break;
		}
		if (arity==1)
		{
			instr.m_type=WKSContentListener::FormulaInstruction::F_Operator;
			if (instr.m_content=="%")
				stack[numElt-1].push_back(instr);
			else if (instr.m_content==")")
			{
				stack[numElt-1].push_back(instr);
				instr.m_content="(";
				stack[numElt-1].insert(stack[numElt-1].begin(), instr);
			}
			else
				stack[numElt-1].insert(stack[numElt-1].begin(), instr);
			continue;
		}
		if (arity==2)
		{
			instr.m_type=WKSContentListener::FormulaInstruction::F_Operator;
			stack[numElt-2].push_back(instr);
			stack[numElt-2].insert(stack[numElt-2].end(), stack[numElt-1].begin(), stack[numElt-1].end());
			stack.resize(numElt-1);
			continue;
		}
		ok=false;
		error = "### unexpected arity";
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		break;
	}
	pos=input->tell();
	if (pos!=endPos || !ok || stack.size()!=1 || stack[0].empty())
	{
		WPS_DEBUG_MSG(("MultiplanParser::readFormula: can not read a formula\n"));
		ascii().addDelimiter(pos, '|');
		input->seek(endPos, librevenge::RVNG_SEEK_SET);

		std::stringstream s;
		if (!error.empty())
			s << error;
		else
			s << "##unknownError";
		s << "[";
		for (auto const &i : stack)
		{
			for (auto const &j : i)
				s << j << ",";
		}
		s << "],";
		error=s.str();
		return true;
	}
	formula=stack[0];
	return true;
}
////////////////////////////////////////////////////////////
//   generic
////////////////////////////////////////////////////////////
bool MultiplanParser::readCellDataPosition(WPSEntry const &entry)
{
	if (m_state->m_maximumCell[0]<=0 || m_state->m_maximumCell[1]<=0 ||
	        entry.length()/2/m_state->m_maximumCell[0]<m_state->m_maximumCell[1])
	{
		WPS_DEBUG_MSG(("MultiplanParser::readCellDataPosition: the zone seems bad\n"));
		return false;
	}
	auto input = getInput();
	input->seek(entry.begin(), librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f;
	f << "Entries(CellPos):";
	auto &posSet=m_state->m_zonesList[0].m_positionsSet;
	for (int i=0; i<m_state->m_maximumCell[0]; ++i)
	{
		f << "[" << std::hex;
		bool hasValues=false;
		std::vector<int> cellPos;
		for (int j=0; j<m_state->m_maximumCell[1]; ++j)
		{
			cellPos.push_back(int(libwps::readU16(input)));
			posSet.insert(cellPos.back());
			if (cellPos.back())
			{
				hasValues=true;
				f << cellPos.back() << ",";
			}
			else
				f << "_,";
		}
		f << std::dec << "],";
		if (hasValues)
			m_state->m_cellPositionsMap[i]=cellPos;
	}
	if (input->tell()!=entry.end())
	{
		WPS_DEBUG_MSG(("MultiplanParser::readCellDataPosition: find extra data\n"));
		f << "###extra";
		ascii().addDelimiter(input->tell(),'|');
	}
	ascii().addPos(entry.begin());
	ascii().addNote(f.str().c_str());
	return true;
}

bool MultiplanParser::readLink(int pos, WKSContentListener::FormulaInstruction &instruction)
{
	auto it=m_state->m_posToLinkMap.find(pos);
	if (it!=m_state->m_posToLinkMap.end())
	{
		instruction=it->second;
		return true;
	}
	int const vers=version();
	auto const &entry=m_state->m_entries[MultiplanParserInternal::Zone::Link];
	if (!entry.valid() || pos<0 || pos+9>entry.length())
	{
		WPS_DEBUG_MSG(("MultiplanParser::readLink: the pos %d seems bad\n", pos));
		return false;
	}
	auto input = getInput();
	long actPos=input->tell();
	long begPos=entry.begin()+pos;
	input->seek(begPos, librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f;
	f << "Link-" << std::hex << pos << std::dec << "[pos]:";
	int val=libwps::readU16(input);
	int dSz=(val&0x1f);
	int type=(val>>5);
	if ((val>>5)!=2) f << "type=" << (val>>5) << ",";
	val=libwps::readU16(input);
	librevenge::RVNGString filename;
	if (pos+(vers==1 ? 8 : 10)+dSz>entry.length() || !readFilename(val, filename))
	{
		WPS_DEBUG_MSG(("MultiplanParser::readLink: the pos %d seems bad\n", pos));
		input->seek(actPos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	int rows[2];
	for (auto &r : rows) r=vers==1 ? int(libwps::readU8(input)) : int(libwps::readU16(input));
	int cols[2];
	for (auto &c : cols) c=int(libwps::readU8(input));
	if (rows[0]==rows[1] && cols[0]==cols[1]) // used position?
		f << "pos=" << Vec2i(cols[0],rows[0]) << ",";
	else
		f << "pos=" << WPSBox2i(Vec2i(cols[0],rows[0]),Vec2i(cols[1],rows[1])) << ",";
	bool ok=true;
	switch (type)
	{
	case 2: // file + cell's name
		// we can not a named reference to a link, so store the reference in a text zone
		filename.append(':');
		filename.append(libwps_tools_win::Font::unicodeString(input.get(), unsigned(dSz), m_state->getDefaultFontType()));
		instruction.m_type=instruction.F_Text;
		instruction.m_content=filename.cstr();
		break;
	case 3: // file + cell's reference
		if ((vers==1 && dSz!=4) || (vers>1 && dSz!=6))
		{
			WPS_DEBUG_MSG(("MultiplanParser::readLink: unexpected size\n"));
			f << "##";
			ok=false;
			break;
		}
		instruction.m_fileName=filename;
		instruction.m_sheetName[0]=instruction.m_sheetName[1]="Sheet0";
		for (auto &r : rows) r=vers==1 ? int(libwps::readU8(input)) : int(libwps::readU16(input));
		for (auto &c : cols) c=int(libwps::readU8(input));
		for (int i=0; i<2; ++i)
		{
			instruction.m_position[i]=Vec2i(cols[i], rows[i]);
			instruction.m_positionRelative[i]=Vec2b(false,false);
		}
		instruction.m_type=instruction.m_position[0]==instruction.m_position[1] ? instruction.F_Cell : instruction.F_CellList;
		f << instruction << ",";
		break;
	default:
		WPS_DEBUG_MSG(("MultiplanParser::readFilename: unknown type\n"));
		f << "##";
		ok=false;
		break;
	}
	if (ok)
		m_state->m_posToLinkMap[pos]=instruction;
	ascii().addPos(begPos);
	ascii().addNote(f.str().c_str());
	input->seek(actPos, librevenge::RVNG_SEEK_SET);
	return true;
}

bool MultiplanParser::readFilename(int pos, librevenge::RVNGString &filename)
{
	filename.clear();
	auto const &entry=m_state->m_entries[MultiplanParserInternal::Zone::FileName];
	if (!entry.valid() || pos<0 || pos+3>entry.length())
	{
		WPS_DEBUG_MSG(("MultiplanParser::readFilename: the pos %d seems bad\n", pos));
		return false;
	}
	auto input = getInput();
	long actPos=input->tell();
	long begPos=entry.begin()+pos;
	input->seek(begPos, librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f;
	f << "FileName-" << std::hex << pos << std::dec << ":";
	int val=libwps::readU16(input);
	int dSz=(val&0x1f);
	if (pos+2+dSz>entry.length())
	{
		WPS_DEBUG_MSG(("MultiplanParser::readFilename: the pos %d seems bad\n", pos));
		input->seek(actPos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	if (val>>5) f << "f0=" << (val>>5) << ",";
	filename=libwps_tools_win::Font::unicodeString(input.get(), unsigned(dSz), m_state->getDefaultFontType());
	ascii().addDelimiter(input->tell(),'|');
	ascii().addPos(begPos);
	ascii().addNote(f.str().c_str());
	input->seek(actPos, librevenge::RVNG_SEEK_SET);
	return !filename.empty();
}

bool MultiplanParser::readName(int pos, WKSContentListener::FormulaInstruction &instruction)
{
	int const vers=version();
	auto it=m_state->m_posToNameMap.find(pos);
	if (it!=m_state->m_posToNameMap.end())
	{
		instruction=it->second;
		return true;
	}
	auto const &entry=m_state->m_entries[MultiplanParserInternal::Zone::Name];
	if (!entry.valid() || pos<0 || pos+9>=entry.length())
	{
		WPS_DEBUG_MSG(("MultiplanParser::readName: the pos %d seeems bad\n", pos));
		return false;
	}
	auto input = getInput();
	long actPos=input->tell();
	long begPos=entry.begin()+pos;
	input->seek(begPos, librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f;
	f << "Names-" << std::hex << pos << std::dec << ":";
	int val;
	if (vers==1)
	{
		val=int(libwps::readU16(input));
		f << "unkn=" << std::hex << val << std::dec << ",";
	}
	val=int(libwps::readU8(input));
	int dSz=(val&0x1f);
	if (dSz<=0 || begPos+8+dSz>entry.end())
	{
		input->seek(actPos, librevenge::RVNG_SEEK_SET);
		WPS_DEBUG_MSG(("MultiplanParser::readName: the pos %d seeems bad\n", pos));
		return false;
	}
	if (val>>5) f << "f0=" << (val>>5) << ",";
	val=int(libwps::readU8(input)); // v1: 0|40, v2: 0|4
	if (val) f << "f1=" << std::hex << val << std::dec << ",";
	int type=6;
	if (vers>=2)
	{
		type=int(libwps::read16(input));
		if (type!=6) f << "type=" << type << ",";
		f << "unk=[";
		for (int i=0; i<2; ++i)   // _,_
		{
			val=int(libwps::read16(input));
			if (val)
				f << val << ",";
			else
				f << "_,";
		}
		f << "],";
	}
	bool ok=true;
	if (type==6)
	{
		if (input->tell()+(vers==1 ? 4 : 6)+dSz<=entry.end())
		{
			int rows[2];
			for (auto &r : rows) r=vers==1 ? int(libwps::readU8(input)) : int(libwps::readU16(input));
			int cols[2];
			for (auto &c : cols) c=int(libwps::readU8(input));
			for (int i=0; i<2; ++i)
			{
				instruction.m_position[i]=Vec2i(cols[i], rows[i]);
				instruction.m_positionRelative[i]=Vec2b(false,false);
			}
			instruction.m_type=instruction.m_position[0]==instruction.m_position[1] ? instruction.F_Cell : instruction.F_CellList;
			f << instruction << ",";
		}
		else
		{
			f << "###";
			WPS_DEBUG_MSG(("MultiplanParser::readName: the pos %d seeems bad\n", pos));
		}
	}
	std::string name;
	for (int c=0; c<dSz; ++c) name+=char(libwps::readU8(input));
	f << name << ",";
	switch (type)
	{
	case 0: // bad name?
		instruction.m_type=instruction.F_Text;
		instruction.m_content=name;
		break;
	case 6: // cell list
		break;
	default:
		f << "###";
		WPS_DEBUG_MSG(("MultiplanParser::readName: unknown type for pos %d\n", pos));
		ok=false;
		break;
	}
	if (ok)
		m_state->m_posToNameMap[pos]=instruction;
	ascii().addPos(begPos);
	ascii().addNote(f.str().c_str());

	input->seek(actPos, librevenge::RVNG_SEEK_SET);
	return true;
}

bool MultiplanParser::readSharedData(int pos, int cellType, Vec2i const &cellPos, WKSContentListener::CellContent &content)
{
	int const vers=version();
	auto const &entry=m_state->m_entries[MultiplanParserInternal::Zone::SharedData];
	if (!entry.valid() || pos<0 || pos+(vers==1 ? 3 : 4) >entry.length())
	{
		WPS_DEBUG_MSG(("MultiplanParser::readSharedData: the pos %d seems bad\n", pos));
		return false;
	}
	auto input = getInput();
	long actPos=input->tell();
	long begPos=entry.begin()+pos;
	input->seek(begPos, librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f;
	f << "SharedData-" << std::hex << pos << std::dec << ":";
	int val=int(libwps::readU16(input)), N=val;
	int type=0;
	if (vers==1)
	{
		type=(val>>15);
		N &= 0x7fff;
	}
	if (type) f << "type=" << std::hex << type << std::dec << ",";
	if (N!=2) f << "used=" << N << ",";
	if (vers>=2)
	{
		// checkme
		val=int(libwps::readU8(input));
		f << "type=" << std::hex << val << std::dec << ",";
		type = (val&4) ? 1 : 0;
	}
	int dSz=int(libwps::readU8(input));
	if (vers>=2) dSz*=2;
	long endPos=input->tell()+dSz;
	if (endPos>entry.end())
	{
		WPS_DEBUG_MSG(("MultiplanParser::readSharedData: the pos %d seems bad\n", pos));
		input->seek(actPos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	bool ok=true;
	switch (type)
	{
	case 0:
		switch (cellType&3)
		{
		case 0:
		{
			double value;
			if (dSz!=8 || !readDouble(value))
				ok=false;
			else
			{
				content.m_contentType=content.C_NUMBER;
				content.setValue(value);
				f << value << ",";
			}
			break;
		}
		case 1:
		{
			content.m_contentType=content.C_TEXT;
			content.m_textEntry.setBegin(input->tell());
			content.m_textEntry.setLength(dSz);
			std::string name;
			for (int c=0; c<dSz; ++c)
			{
				char ch=char(libwps::readU8(input));
				if (vers>=2 && ch==0 && c+1==dSz)
				{
					content.m_textEntry.setEnd(input->tell()-1);
					break;
				}
				name+=ch;
			}
			f << name << ",";
			break;
		}
		case 2:
			if (dSz!=8)
				ok=false;
			else
			{
				f << "Nan" << libwps::readU8(input) << ",";
				input->seek(7, librevenge::RVNG_SEEK_CUR);
				content.m_contentType=content.C_NUMBER;
				content.setValue(std::nan(""));
			}
			break;
		case 3:
		default:
			if (dSz!=8)
				ok=false;
			else
			{
				val=int(libwps::readU8(input));
				content.m_contentType=content.C_NUMBER;
				content.setValue(val);
				if (val==0)
					f << "false,";
				else if (val==1)
					f << "true,";
				else
					f << "##bool=" << val << ",";
				input->seek(7, librevenge::RVNG_SEEK_CUR);
			}
			break;
		}
		break;
	case 1:
	{
		std::string err;
		if ((vers==1 && !readFormula(cellPos, content.m_formula, endPos, err)) ||
		        (vers>=2 && !readFormulaV2(cellPos, content.m_formula, endPos, err)))
			f << "###";
		else
			content.m_contentType=content.C_FORMULA;
		for (auto const &fo : content.m_formula) f << fo;
		f << ",";
		f << err;
		break;
	}
	default:
		ok=false;
		break;
	}
	if (!ok)
	{
		WPS_DEBUG_MSG(("MultiplanParser::readSharedData: can not read data for the pos %d\n", pos));
		f << "###";
	}
	if (m_state->m_posToSharedDataSeen.find(pos)==m_state->m_posToSharedDataSeen.end())
	{
		m_state->m_posToSharedDataSeen.insert(pos);
		if (input->tell()!=endPos)
			ascii().addDelimiter(input->tell(),'|');
		ascii().addPos(begPos);
		ascii().addNote(f.str().c_str());
	}
	input->seek(actPos, librevenge::RVNG_SEEK_SET);
	return true;
}
////////////////////////////////////////////////////////////
//   Unknown
////////////////////////////////////////////////////////////
bool MultiplanParser::readZoneB()
{
	auto input = getInput();
	long pos = input->tell();
	int const vers=version();
	long endPos=pos+(vers==1 ? 0x84 : 0xb9);
	if (!checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("MultiplanParser::readZoneB: the zone seems too short\n"));
		return false;
	}
	libwps::DebugStream f;
	f << "Entries(ZoneB):";
	int dim[2];
	for (auto &d : dim) d=int(libwps::readU16(input));
	m_state->m_maximumCell=Vec2i(dim[0],dim[1]);
	f << "cell[max]=" << m_state->m_maximumCell << ",";
	int val;
	for (int i=0; i<11; ++i)
	{
		if (i==2 && vers==1) continue;
		val=int(libwps::read16(input));
		int const expected[]= {0,0,0xfff, 0xff, 0, 0, 5, 6, 0x46, 0x36, 0x42};
		if (val!=expected[i])
			f << "f" << i << "=" << val << ",";
	}
	for (int i=0; i<(vers==1 ? 11 : 16); ++i)
	{
		val=int(libwps::read16(input));
		if (!val) continue;
		f << "g" << i << "=" << val << ",";
	}
	if (vers>=2)
	{
		val=int(libwps::read8(input));
		if (val)
			f << "h0=" << val << ",";
	}
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	pos=input->tell();
	f.str("");
	f << "ZoneB[II]:";
	for (int i=0; i<8; ++i)
	{
		if ((i==3||i==5) && vers==1) continue;
		val=int(libwps::read8(input));
		int const expected[]= {1, 0, 0, 0, -2, 0xf, 0, 0x3e};
		if (val!=expected[i]) f << "f" << i << "=" << val << ",";
	}
	for (int i=0; i<(vers==1 ? 40 : 59); ++i)   // 0
	{
		val=int(libwps::read16(input));
		if (val) f << "g" << i << "=" << val << ",";
	}
	if (input->tell()!=endPos)
	{
		ascii().addDelimiter(input->tell(),'|');
		input->seek(endPos, librevenge::RVNG_SEEK_SET);
	}
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

////////////////////////////////////////////////////////////
// send data
////////////////////////////////////////////////////////////
bool MultiplanParser::sendCell(Vec2i const &cellPos, int p)
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("MultiplanParser::sendCell: I can not find the m_listener\n"));
		return false;
	}
	int const vers=version();
	int zoneId=(p>>16);
	if (zoneId<0 || zoneId>=int(m_state->m_zonesList.size()))
	{
		WPS_DEBUG_MSG(("MultiplanParser::sendCell: can not find the zone data zone for position %d\n", p));
		return false;
	}
	p&=0xffff;
	auto const &zone=m_state->m_zonesList[size_t(zoneId)];
	auto const &entry=zone.m_entry;
	auto it=zone.m_positionsSet.find(p);
	if (it!=zone.m_positionsSet.end()) ++it;
	long endPos=it!=zone.m_positionsSet.end() ? entry.begin()+*it : entry.end();
	if (p<=0 || p>entry.length())
	{
		WPS_DEBUG_MSG(("MultiplanParser::sendCell: unexpected position %d\n", p));
		return false;
	}
	libwps::DebugStream f;
	f << "CellData[C" << cellPos[0]+1 << "R" << cellPos[1]+1 << "]:";
	MultiplanParserInternal::Cell cell;
	WKSContentListener::CellContent content;
	cell.setPosition(cellPos);
	cell.setFont(m_state->getDefaultFont());
	long pos=entry.begin()+p;
	if (endPos-pos<4)
	{
		WPS_DEBUG_MSG(("MultiplanParser::sendCell: a cell %d seems to short\n", p));
		f << "###";
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		return false;
	}
	auto input=getInput();
	input->seek(pos, librevenge::RVNG_SEEK_SET);
	int formSize=int(libwps::readU8(input));
	if (vers>=2) formSize*=2;
	if (formSize) f << "form[size]=" << std::hex << formSize << std::dec << ",";
	int val=int(libwps::readU8(input));
	int digits=(val>>4);
	if (digits) f << "decimal=" << digits << ",";
	int form=(val>>1)&7;
	int subFormat=0;
	switch (form)
	{
	case 0: // default
		break;
	case 1: // cont
		subFormat=1;
		cell.setDigits(digits);
		f << "decimal,";
		break;
	case 2:
		subFormat=2;
		cell.setDigits(digits);
		f << "scientific,";
		break;
	case 3:
		subFormat=6;
		cell.setDigits(digits);
		f << "fixed,";
		break;
	case 4: // generic
		break;
	case 5:
		subFormat=4;
		cell.setDigits(digits);
		f << "currency,";
		break;
	case 6: // a bar
		f << "bar,";
		break;
	case 7:
		subFormat=3;
		cell.setDigits(digits);
		f << "percent,";
		break;
	default:
		f << "format=" << form << ",";
		break;
	}
	if (val&1)
	{
		cell.setProtected(true);
		f << "protected,";
	}
	int dSz=0, extraSize=(vers>=2 && formSize) ? 4 : 0;
	if (vers==1)
	{
		val=int(libwps::readU8(input));
		dSz=int(libwps::readU8(input));
	}
	else
	{
		dSz=int(libwps::readU8(input));
		val=int(libwps::readU8(input));
	}
	if (endPos<pos+4+dSz+extraSize)
	{
		WPS_DEBUG_MSG(("MultiplanParser::sendCell: a cell seems to short\n"));
		f << "###";
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		return false;
	}
	int type=(val>>6)&3;
	switch (type)
	{
	case 0:
		f << "double,";
		cell.setFormat(WPSCell::F_NUMBER, subFormat);
		content.m_contentType=content.C_NUMBER;
		break;
	case 1:
		cell.setFormat(WPSCell::F_TEXT);
		content.m_contentType=content.C_TEXT;
		f << "text,";
		break;
	case 2:
		cell.setFormat(WPSCell::F_NUMBER);
		content.m_contentType=content.C_NUMBER;
		f << "nan,";
		break;
	case 3: // or nothing
		cell.setFormat(WPSCell::F_BOOLEAN);
		content.m_contentType=content.C_NUMBER;
		f << "bool,";
		break;
	default: // impossible
		break;
	}
	int align=(val>>3)&7;
	bool hasTimeDate=false;
	switch (align)
	{
	case 1:
		cell.setHAlignment(cell.HALIGN_CENTER);
		f << "center,";
		break;
	case 0: // default
	case 2: // generic
		break;
	case 3:
		cell.setHAlignment(cell.HALIGN_LEFT);
		f << "left,";
		break;
	case 4:
		cell.setHAlignment(cell.HALIGN_RIGHT);
		f << "right,";
		break;
	case 7:
		if (vers>=3)
		{
			f << "timeDate,";
			hasTimeDate=true;
			break;
		}
		WPS_FALLTHROUGH;
	default:
		f << "#align=" << align << ",";
		break;
	}
	bool hasShared=false;
	if (val&2)
	{
		f << "shared,";
		hasShared=true;
	}
	if ((val&4)==0) f << "no[4],";
	val &= 0x1;
	if (val) f << "f1=" << std::hex << val << std::dec << ",";
	if (hasTimeDate && pos+4+dSz+extraSize+4 <= endPos)
	{
		endPos-=4;
		long actPos=input->tell();
		input->seek(endPos, librevenge::RVNG_SEEK_SET);
		f << "dateTime=[";
		val=int(libwps::readU8(input));
		if (val!=2) f << "f0=" << val << ",";
		val=int(libwps::readU8(input));
		switch (val&7)
		{
		case 1:
			cell.setHAlignment(cell.HALIGN_CENTER);
			f << "center,";
			break;
		case 0: // default
		case 2: // generic
			break;
		case 3:
			cell.setHAlignment(cell.HALIGN_LEFT);
			f << "left,";
			break;
		case 4:
			cell.setHAlignment(cell.HALIGN_RIGHT);
			f << "right,";
			break;
		default:
			f << "#align=" << (val&7) << ",";
			break;
		}
		val &=0xf8;
		if (val!=0x90) f << "f1=" << std::hex << val << std::dec << ",";
		int format=int(libwps::readU16(input));
		char const *dtFormat[]=
		{
			nullptr, "%m/%d/%y", "%m/%d", "%d-%b-%y",
			"%d-%b", "%b-%y", "%I:%M%p", "%I:%M:%S%p",
			"%H:%M", "%H:%M:%S", "%m/%d/%y %H:%M"
		};
		switch (format)
		{
		case 1:
		case 2:
		case 3:
		case 4:
		case 5:
		case 10:
			cell.setDTFormat(WPSCell::F_DATE, dtFormat[format]);
			f << dtFormat[format] << ",";
			break;
		case 6:
		case 7:
		case 8:
		case 9:
			cell.setDTFormat(WPSCell::F_TIME, dtFormat[format]);
			f << dtFormat[format] << ",";
			break;
		default:
			WPS_DEBUG_MSG(("MultiplanParser::sendCell: unknown data format\n"));
			f << "###format=" << format << ",";
		}
		f << "],";
		input->seek(actPos, librevenge::RVNG_SEEK_SET);
	}
	else if (hasTimeDate)
	{
		f << "###";
		WPS_DEBUG_MSG(("MultiplanParser::sendCell: can not find the time value data\n"));
	}
	if (vers>=2 && formSize)
	{
		f << "form=[";
		for (int i=0; i<2; ++i) f << libwps::read16(input) << ",";
		f << "],";
	}
	if (type==0 && dSz==8)
	{
		double value;
		if (!readDouble(value))
			f << "###";
		else
			content.setValue(value);
		f << value << ",";
	}
	else if (type==1 && dSz && pos+4+dSz+extraSize+(hasShared ? 2 : 0) <= endPos)
	{
		content.m_textEntry.setBegin(input->tell());
		content.m_textEntry.setLength(dSz);
		std::string name;
		for (int c=0; c<dSz; ++c) name+=char(libwps::readU8(input));
		f << name << ",";
	}
	else if (type==2 && dSz==8)
	{
		content.setValue(std::nan(""));
		f << "Nan" << int(libwps::readU8(input)) << ",";
		input->seek(7, librevenge::RVNG_SEEK_CUR);
	}
	else if (type==3 && dSz==8)
	{
		val=int(libwps::readU8(input));
		content.setValue(val);
		if (val==0)
			f << "false,";
		else if (val==1)
			f << "true,";
		else
			f << "##bool=" << val << ",";
		input->seek(7, librevenge::RVNG_SEEK_CUR);
	}
	if (hasShared && input->tell()+2<=endPos && (formSize==0 || formSize==2))
	{
		if ((input->tell()-pos)%2)
			input->seek(1, librevenge::RVNG_SEEK_CUR);
		int nPos=int(libwps::readU16(input));
		if (!readSharedData(nPos, type, cellPos, content))
			f << "###";
		f << "sharedData-" << std::hex << nPos << std::dec << ",";
	}
	else if (!hasShared && formSize && input->tell()+formSize<=endPos)
	{
		auto endFPos=input->tell()+formSize;
		std::string err;
		if ((vers==1 && !readFormula(cellPos, content.m_formula, endFPos, err)) ||
		        (vers>=2 && !readFormulaV2(cellPos, content.m_formula, endFPos, err)))
		{
			ascii().addDelimiter(input->tell(),'|');
			f << "###";
		}
		else
			content.m_contentType=content.C_FORMULA;

		for (auto const &fo : content.m_formula) f << fo;
		f << ",";
		f << err;
		input->seek(endFPos, librevenge::RVNG_SEEK_SET);
	}
	else if (formSize)
	{
		WPS_DEBUG_MSG(("MultiplanParser::sendCell: can not read a formula\n"));
		f << "###form";
	}
	m_listener->openSheetCell(cell, content);
	if (content.m_textEntry.valid())
	{
		auto fontType=m_state->getDefaultFontType();
		m_listener->setFont(cell.getFont());
		input->seek(content.m_textEntry.begin(), librevenge::RVNG_SEEK_SET);
		std::string text;
		while (input->tell()<=content.m_textEntry.end())
		{
			bool last=input->isEnd() || input->tell()>=content.m_textEntry.end();
			auto c=last ? '\0' : char(libwps::readU8(input));
			if ((c==0 || c==0x9 || c==0xa || c==0xd) && !text.empty())
			{
				m_listener->insertUnicodeString(libwps_tools_win::Font::unicodeString(text, fontType));
				text.clear();
			}
			if (last) break;
			if (c==0x9)
				m_listener->insertTab();
			else if (c==0xa || c==0xd)
				m_listener->insertEOL();
			else if (c)
				text.push_back(c);
		}
	}
	m_listener->closeSheetCell();
	if (vers==1 && input->tell()!=endPos)
		ascii().addDelimiter(input->tell(),'|');

	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

bool MultiplanParser::sendSpreadsheet()
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("MultiplanParser::sendSpreadsheet: I can not find the m_listener\n"));
		return false;
	}
	for (auto &z : m_state->m_zonesList)
		z.m_positionsSet.insert(int(z.m_entry.length()));
	m_listener->openSheet(m_state->getColumnsWidth(), "Sheet0");
	WPSRowFormat rowFormat(16);
	rowFormat.m_isMinimalHeight=true;
	int lastRow=0;
	for (auto it : m_state->m_cellPositionsMap)
	{
		int r=it.first;
		auto const &row = it.second;
		if (r>lastRow)
		{
			m_listener->openSheetRow(rowFormat,r-lastRow);
			m_listener->closeSheetRow();
		}
		lastRow=r+1;
		m_listener->openSheetRow(rowFormat);
		for (size_t col=0; col<row.size(); ++col)
		{
			int zId=(row[col]>>24);
			int p=row[col]&0xffff;
			if (zId<0 || zId>=int(m_state->m_zonesList.size()) ||
			        p<0 || p>=m_state->m_zonesList[size_t(zId)].m_entry.length())
			{
				WPS_DEBUG_MSG(("MultiplanParser::sendSpreadsheet: find some bad data\n"));
				continue;
			}
			if (!p) continue;
			Vec2i cellPos(static_cast<int>(col), static_cast<int>(r));
			sendCell(cellPos, row[col]);
		}
		m_listener->closeSheetRow();
	}
	m_listener->closeSheet();
	return true;
}

////////////////////////////////////////////////////////////
// password
////////////////////////////////////////////////////////////
bool MultiplanParser::checkPassword(char const *password) const
{
	char const endPassword[]=
	{
		0x0A, 0x4E, 0x51, 0x6F, 0x6E, 0x61, 0x70, 0x32, 0x33, 0x71, 0x5B,
		0x30, 0x23, 0x7A
	};
	if (!password || !password[0])
	{
		WPS_DEBUG_MSG(("MultiplanParser::checkPassword: can not find the password\n"));
		return false;
	}
	// set password length to 15 by adding endPassword's data if needed
	char pw[16];
	int w=0;
	while (*password)
	{
		if (w>=15) break;
		pw[w++]=*(password++);
	}
	for (int r=0; w<15;) pw[w++]=char(endPassword[r++]);
	pw[15]=char(0);

	// permutate the solution
	int const which=(m_state->m_hash&0xf);
	if (which!=15) ++pw[which]; // never seem which=15, maybe a special case
	int const perm[]= { 9, 4, 1, 3, 14, 11, 6, 0, 12, 7, 2, 10, 8, 13, 5 };
	uint8_t res[16];
	for (int i=0; i<15; ++i) res[i]=uint8_t(pw[perm[(i+which)%15]]);
	res[15]=0;

	int len=0;
	while (res[len]) ++len;
	if (len!=15)
	{
		WPS_DEBUG_MSG(("MultiplanParser::checkPassword: unexpected size for the password\n"));
		return false;
	}

	// time to check if the checksum is ok or not
	int const data[]=
	{
		0x4ec3, 0xaefc, 0x4dd9, 0x9bb2, 0x2745, 0x4e8a, 0x9d14, 0x2a09,
		0x7b61, 0xf6c2, 0xfda5, 0xeb6b, 0xc6f7, 0x9dcf, 0x2bbf, 0x4563,
		0x8ac6, 0x05ad, 0x0b5a, 0x16b4, 0x2d68, 0x5ad0, 0x0375, 0x06ea,
		0x0dd4, 0x1ba8, 0x3750, 0x6ea0, 0xdd40, 0xd849, 0xa0b3, 0x5147,
		0xa28e, 0x553d, 0xaa7a, 0x44d5, 0x6f45, 0xde8a, 0xad35, 0x4a4b,
		0x9496, 0x390d, 0x721a, 0xeb23, 0xc667, 0x9cef, 0x29ff, 0x53fe,
		0xa7fc, 0x5fd9, 0x47d3, 0x8fa6, 0x0f6d, 0x1eda, 0x3db4, 0x7b68,
		0xf6d0, 0xb861, 0x60e3, 0xc1c6, 0x93ad, 0x377b, 0x6ef6, 0xddec,
		0x45a0, 0x8b40, 0x06a1, 0x0d42, 0x1a84, 0x3508, 0x6a10, 0xaa51,
		0x4483, 0x8906, 0x022d, 0x045a, 0x08b4, 0x1168, 0x76b4, 0xed68,
		0xcaf1, 0x85c3, 0x1ba7, 0x374e, 0x6e9c, 0x3730, 0x6e60, 0xdcc0,
		0xa9a1, 0x4363, 0x86c6, 0x1dad, 0x3331, 0x6662, 0xccc4, 0x89a9,
		0x0373, 0x06e6, 0x0dcc, 0x1021, 0x2042, 0x4084, 0x8108, 0x1231,
		0x2462, 0x48c4
	};
	int const *dataPtr=data;
	int val=*(dataPtr++);
	for (auto r : res)
	{
		for (int byte=1, dec=0; dec<7; byte<<=1, ++dec)
		{
			if (r&byte)
				val^=*dataPtr;
			dataPtr++;
		}
	}
	if (val!=m_state->m_checksum)
	{
		WPS_DEBUG_MSG(("MultiplanParser::checkPassword: can not check the password\n"));
		return false;
	}
	// create the decoding keys
	res[15]=uint8_t(0xbb);
	for (int i=0; i<16; ++i)
	{
		auto v=res[i]^((i%2)==0 ? (m_state->m_checksum&0xFF) : (m_state->m_checksum>>8));
		m_state->m_keys[i]=uint8_t((v>>1)|(v<<7));
	}
	// original password in 14d6:2220, final keys 14d6:49cb
	return true;
}

bool MultiplanParser::retrievePasswordKeys()
{
	auto input=getInput();
	long actPos=input->tell();
	// too short or not the zone0 zone
	if (!checkFilePosition(actPos+6) || libwps::readU16(input)!=7)
	{
		input->seek(actPos,librevenge::RVNG_SEEK_SET);
		return false;
	}
	input->seek(2,librevenge::RVNG_SEEK_CUR);
	int dSz=int(libwps::readU16(input));
	// not enough data to retrieve the keys final block
	if (dSz<22 || !checkFilePosition(actPos+dSz))
	{
		input->seek(actPos,librevenge::RVNG_SEEK_SET);
		return false;
	}
	uint8_t res[16];
	for (int i=0; i<16; ++i)
	{
		auto key=uint8_t(readU8(input));
		if (i==0) key^=8; // first value is normally 8, other 0
		auto r=uint8_t((key<<1)|(key>>7));
		res[(i+6)&0xf]=uint8_t(r^((i%2)==0 ? (m_state->m_checksum&0xFF) : (m_state->m_checksum>>8)));
	}
	if (res[15]!=0xbb)   // last value must be bb
	{
		input->seek(actPos,librevenge::RVNG_SEEK_SET);
		return false;
	}
	// inverse the initial permutation
	char pw[16];
	int const perm[]= { 9, 4, 1, 3, 14, 11, 6, 0, 12, 7, 2, 10, 8, 13, 5 };
	int const which=(m_state->m_hash&0xf);
	for (int i=0; i<15; ++i) pw[perm[(i+which)%15]]=char(res[i]);
	if (which!=15) --pw[which];
	pw[15]=0;
	std::string password;
	for (auto c : pw)
	{
		if (c==0 || c==0xa) break;
		password+=c;
	}
	bool ok=!password.empty() && checkPassword(password.c_str());
	input->seek(actPos,librevenge::RVNG_SEEK_SET);
	return ok;
}

RVNGInputStreamPtr MultiplanParser::decodeStream(RVNGInputStreamPtr input)
{
	if (!input)
	{
		WPS_DEBUG_MSG(("MultiplanParser::decodeStream: the arguments seems bad\n"));
		return RVNGInputStreamPtr();
	}
	long actPos=input->tell();
	input->seek(0,librevenge::RVNG_SEEK_SET);
	librevenge::RVNGBinaryData data;
	if (!libwps::readDataToEnd(input, data) || !data.getDataBuffer())
	{
		WPS_DEBUG_MSG(("MultiplanParser::decodeStream: can not read the original input\n"));
		return RVNGInputStreamPtr();
	}
	auto *buf=const_cast<unsigned char *>(data.getDataBuffer());
	input->seek(actPos,librevenge::RVNG_SEEK_SET);
	auto const &keys=m_state->m_keys;
	while (!input->isEnd())
	{
		long pos=input->tell();
		if (!checkFilePosition(pos+6)) break;
		int type=int(libwps::readU16(input));
		if (type<7 || type>12) break;
		input->seek(2,  librevenge::RVNG_SEEK_CUR);
		int dSz=int(libwps::readU16(input));
		if (dSz<6 || !checkFilePosition(pos+dSz)) break;
		if (dSz==6) continue;
		for (int i=6; i<dSz; ++i)
			buf[pos+i]^=keys[i&0xf];
		input->seek(dSz-6,librevenge::RVNG_SEEK_CUR);
	}
	RVNGInputStreamPtr res(new WPSStringStream(data.getDataBuffer(), static_cast<unsigned int>(data.size())));
	res->seek(actPos, librevenge::RVNG_SEEK_SET);
	return res;
}

}

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
