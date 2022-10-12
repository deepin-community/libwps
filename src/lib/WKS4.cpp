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
#include <sstream>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WKSContentListener.h"
#include "WKSSubDocument.h"

#include "WPSEntry.h"
#include "WPSFont.h"
#include "WPSHeader.h"
#include "WPSPageSpan.h"
#include "WPSStringStream.h"
#include "WPSTable.h"

#include "WKS4Chart.h"
#include "WKS4Format.h"
#include "WKS4Spreadsheet.h"

#include "WKS4.h"

using namespace libwps;

//! Internal: namespace to define internal class of WKS4Parser
namespace WKS4ParserInternal
{
//! the font of a WKS4Parser
struct Font final : public WPSFont
{
	//! constructor
	explicit Font(libwps_tools_win::Font::Type type) : WPSFont(), m_type(type)
	{
	}
	Font(Font const &)=default;
	//! destructor
	~Font() final;
	//! font encoding type
	libwps_tools_win::Font::Type m_type;
};

Font::~Font()
{
}

//! Internal: the subdocument of a WPS4Parser
class SubDocument final : public WKSSubDocument
{
public:
	//! constructor for a text entry
	SubDocument(RVNGInputStreamPtr const &input, WKS4Parser &pars, bool header) :
		WKSSubDocument(input, &pars), m_header(header) {}
	//! destructor
	~SubDocument() final {}

	//! operator==
	virtual bool operator==(std::shared_ptr<WKSSubDocument> const &doc) const final
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
		WPS_DEBUG_MSG(("WKS4ParserInternal::SubDocument::parse: no listener\n"));
		return;
	}
	if (!dynamic_cast<WKSContentListener *>(listener.get()))
	{
		WPS_DEBUG_MSG(("WKS4ParserInternal::SubDocument::parse: bad listener\n"));
		return;
	}

	WKS4Parser *pser = m_parser ? dynamic_cast<WKS4Parser *>(m_parser) : nullptr;
	if (!pser)
	{
		listener->insertCharacter(' ');
		WPS_DEBUG_MSG(("WKS4ParserInternal::SubDocument::parse: bad parser\n"));
		return;
	}
	pser->sendHeaderFooter(m_header);
}

//! the state of WKS4Parser
struct State
{
	//! constructor
	State(libwps_tools_win::Font::Type fontType, char const *password)
		: m_eof(-1)
		, m_creator(libwps::WPS_MSWORKS)
		, m_isSpreadsheet(true)
		, m_fontType(fontType)
		, m_version(-1)
		, m_hasLICSCharacters(false)
		, m_fontsList()
		, m_pageSpan()
		, m_actPage(0)
		, m_numPages(0)
		, m_headerString("")
		, m_footerString("")
		, m_password(password)
		, m_isEncrypted(false)
		, m_isDecoded(false)
	{
	}
	//! returns a color corresponding to an id
	bool getColor(int id, WPSColor &color) const;
	//! return the default font style
	libwps_tools_win::Font::Type getDefaultFontType() const
	{
		if (m_hasLICSCharacters && m_version<=2)
			return libwps_tools_win::Font::LICS;
		if (m_fontType != libwps_tools_win::Font::UNKNOWN)
			return m_fontType;
		if (m_version>2)
			return libwps_tools_win::Font::WIN3_WEUROPE;
		return m_creator==libwps::WPS_MSWORKS ? libwps_tools_win::Font::DOS_850 : libwps_tools_win::Font::CP_437;
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

	//! the last file position
	long m_eof;
	//! the application
	libwps::WPSCreator m_creator;
	//! boolean to know if the file is a spreadsheet file or a database file
	bool m_isSpreadsheet;
	//! the user font type
	libwps_tools_win::Font::Type m_fontType;
	//! the file version
	int m_version;
	//! flag to know if the character
	bool m_hasLICSCharacters;
	//! the fonts list
	std::vector<Font> m_fontsList;
	//! the actual document size
	WPSPageSpan m_pageSpan;
	int m_actPage /** the actual page*/, m_numPages /* the number of pages */;
	//! the header string
	librevenge::RVNGString m_headerString;
	//! the footer string
	librevenge::RVNGString m_footerString;

	//! the password (if known)
	char const *m_password;
	//! true if the file is encrypted
	bool m_isEncrypted;
	//! true if the main stream has been decoded
	bool m_isDecoded;

private:
	State(State const &) = delete;
	State &operator=(State const &) = delete;
};

bool State::getColor(int id, WPSColor &color) const
{
	if (m_version<=2)
	{
		static const uint32_t colorDosMap[]=
		{
			0x0, 0xFF0000, 0x00FF00, 0x0000FF,
			0x00FFFF, 0xFF00FF, 0xFFFF00
		};
		if (id < 0 || id >= 7)
		{
			WPS_DEBUG_MSG(("WKS4ParserInternal::State::getColor(): unknown Dos color id: %d\n",id));
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
		0x008000, 0x800080,	0x808000, 0xC0C0C0
	};
	if (id < 0 || id >= 16)
	{
		WPS_DEBUG_MSG(("WKS4ParserInternal::State::getColor(): unknown color id: %d\n",id));
		return false;
	}
	color = WPSColor(colorMap[id]);
	return true;
}

}

// constructor, destructor
WKS4Parser::WKS4Parser(RVNGInputStreamPtr &input, WPSHeaderPtr &header,
                       libwps_tools_win::Font::Type encoding, char const *password)
	: WKSParser(input, header)
	, m_listener()
	, m_state(new WKS4ParserInternal::State(encoding, password))
	, m_chartParser(new WKS4Chart(*this))
	, m_spreadsheetParser(new WKS4Spreadsheet(*this))
{
}

WKS4Parser::~WKS4Parser()
{
}

int WKS4Parser::version() const
{
	return m_state->m_version;
}

libwps::WPSCreator WKS4Parser::creator() const
{
	return m_state->m_creator;
}

void WKS4Parser::resetMainInput(RVNGInputStreamPtr newInput)
{
	resetInput(newInput);
	ascii().setStream(newInput);
	m_chartParser->resetInput(newInput);
	m_spreadsheetParser->resetInput(newInput);
}

bool WKS4Parser::checkFilePosition(long pos)
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

libwps_tools_win::Font::Type WKS4Parser::getDefaultFontType() const
{
	return m_state->getDefaultFontType();
}

libwps::WPSCreator WKS4Parser::getCreator() const
{
	return m_state->m_creator;
}

//////////////////////////////////////////////////////////////////////
// interface with WKS4Spreadsheet
//////////////////////////////////////////////////////////////////////
bool WKS4Parser::getColor(int id, WPSColor &color) const
{
	return m_state->getColor(id, color);
}

bool WKS4Parser::getFont(int id, WPSFont &font, libwps_tools_win::Font::Type &type) const
{
	if (id < 0 || id>=int(m_state->m_fontsList.size()))
	{
		WPS_DEBUG_MSG(("WKS4Parser::getFont: can not find font %d\n", id));
		return false;
	}
	auto const &ft=m_state->m_fontsList[size_t(id)];
	font=ft;
	type=ft.m_type;
	return true;
}

librevenge::RVNGString WKS4Parser::getSheetName(int id) const
{
	return m_spreadsheetParser->getSheetName(id);
}

// main function to parse the document
void WKS4Parser::parse(librevenge::RVNGSpreadsheetInterface *documentInterface)
{
	RVNGInputStreamPtr input=getInput();
	if (!input)
	{
		WPS_DEBUG_MSG(("WKS4Parser::parse: does not find main ole\n"));
		throw (libwps::ParseException());
	}

	if (!checkHeader(nullptr)) throw(libwps::ParseException());

	bool ok=false;
	try
	{
		ascii().setStream(input);
		ascii().open("MN0");

		if (checkHeader(nullptr) && readZones())
		{
			parseFormatStream();
			m_listener=createListener(documentInterface);
		}
		if (m_listener)
		{
			m_chartParser->setListener(m_listener);
			m_spreadsheetParser->setListener(m_listener);

			m_listener->startDocument();
			int numSheet=m_spreadsheetParser->getNumSpreadsheets();
			if (numSheet==0) ++numSheet;
			for (int i=0; i<numSheet; ++i)
				m_spreadsheetParser->sendSpreadsheet(i);
			if (m_state->m_isSpreadsheet && m_chartParser->getNumCharts())
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
		WPS_DEBUG_MSG(("WKS4Parser::parse: password exception catched when parsing MN0\n"));
		throw (libwps::PasswordException());
	}
	catch (...)
	{
		WPS_DEBUG_MSG(("WKS4Parser::parse: exception catched when parsing MN0\n"));
		throw (libwps::ParseException());
	}

	ascii().reset();
	if (!ok)
		throw(libwps::ParseException());
}

std::shared_ptr<WKSContentListener> WKS4Parser::createListener(librevenge::RVNGSpreadsheetInterface *interface)
{
	std::vector<WPSPageSpan> pageList;
	WPSPageSpan ps(m_state->m_pageSpan);
	if (!m_state->m_headerString.empty())
	{
		WPSSubDocumentPtr subdoc(new WKS4ParserInternal::SubDocument
		                         (getInput(), *this, true));
		ps.setHeaderFooter(WPSPageSpan::HEADER, WPSPageSpan::ALL, subdoc);
	}
	if (!m_state->m_footerString.empty())
	{
		WPSSubDocumentPtr subdoc(new WKS4ParserInternal::SubDocument
		                         (getInput(), *this, false));
		ps.setHeaderFooter(WPSPageSpan::FOOTER, WPSPageSpan::ALL, subdoc);
	}
	pageList.push_back(ps);
	return std::shared_ptr<WKSContentListener>(new WKSContentListener(pageList, interface));
}

////////////////////////////////////////////////////////////
// low level
////////////////////////////////////////////////////////////
// read the header
////////////////////////////////////////////////////////////
bool WKS4Parser::checkHeader(WPSHeader *header, bool strict)
{
	m_state.reset(new WKS4ParserInternal::State(m_state->m_fontType, m_state->m_password));
	libwps::DebugStream f;

	RVNGInputStreamPtr input = getInput();
	if (!checkFilePosition(12))
	{
		WPS_DEBUG_MSG(("WKS4Parser::checkHeader: file is too short\n"));
		return false;
	}

	input->seek(0,librevenge::RVNG_SEEK_SET);
	auto firstOffset = int(libwps::readU8(input));
	auto type = int(libwps::read8(input));
	bool needEncoding=true;
	f << "FileHeader:";
	if ((firstOffset == 0 && type == 0) ||
	        (firstOffset == 0x20 && type == 0x54))
	{
		m_state->m_version=1;
		f << "DOS,";
	}
	else if (firstOffset == 0xff)
	{
		f << "Windows,";
		m_state->m_version=3;
		needEncoding=false;
	}
	else
	{
		WPS_DEBUG_MSG(("WKS4Parser::checkHeader: find unexpected first data\n"));
		return false;
	}
	auto creatorId=libwps::WPS_MSWORKS;
	auto kind=libwps::WPS_SPREADSHEET;
	bool isSpreadsheet=true;
	if (type == 0x54)
	{
		isSpreadsheet=false;
		kind=libwps::WPS_DATABASE;
		f << "database,";
	}
	else if (type == 0)
		f << "spreadsheet,";
	else
	{
		WPS_DEBUG_MSG(("WKS4Parser::checkHeader: find unexpected type file\n"));
		return false;
	}
	auto val=int(libwps::read16(input));
	if (val==2)
	{
		// version
		val=int(libwps::readU16(input));
		if (isSpreadsheet)
		{
			if (val==0x404)
			{
			}
			else if (val==0x405)
			{
				f << "symphony,";
				creatorId=libwps::WPS_SYMPHONY;
			}
			else if (val==0x406)
			{
				m_state->m_version=1;
				f << "lotus,";
				creatorId=libwps::WPS_LOTUS;
			}
			else if (val==0x5120 || val==0x5121)
			{
				WPS_DEBUG_MSG(("WKS4Parser::checkHeader: must not be called with a DOS Quattro file\n"));
				return false;
			}
			else if (val==0x8006)
			{
				WPS_DEBUG_MSG(("WKS4Parser::checkHeader: find lotus file format, sorry parsing this format is not implemented\n"));
				return false;
			}
			else
			{
#ifdef DEBUG
				f << "vers=" << std::hex << val << std::dec << ",";
#else
				WPS_DEBUG_MSG(("WKS4Parser::checkHeader: find unknown file version\n"));
				return false;
#endif
			}
		}
		else if (val)
			return false;
	}
	else
	{
		WPS_DEBUG_MSG(("WKS4Parser::checkHeader: header contain unexpected size field data\n"));
		return false;
	}

	m_state->m_creator=creatorId;
	input->seek(0, librevenge::RVNG_SEEK_SET);
	if (strict && m_state->m_version<1000)
	{
		for (int i=0; i < 4; ++i)
		{
			if (!readZone()) return false;
			if (m_state->m_isEncrypted) break;
		}
	}
	ascii().addPos(0);
	ascii().addNote(f.str().c_str());

	m_state->m_isSpreadsheet=isSpreadsheet;
	if (header)
	{
		header->setMajorVersion(uint8_t(m_state->m_version));
		header->setCreator(creatorId);
		header->setKind(kind);
		header->setNeedEncoding(needEncoding);
		header->setIsEncrypted(m_state->m_isEncrypted);
	}
	return true;
}

bool WKS4Parser::parseFormatStream()
{
	RVNGInputStreamPtr file=getFileInput();
	if (!file || !file->isStructured() || !m_state->m_isSpreadsheet)
		return false;

	RVNGInputStreamPtr formatInput(file->getSubStreamByName("FMT"));
	if (!formatInput)
	{
		WPS_DEBUG_MSG(("WKS4Parser::parseFormatStream: can not find the format stream\n"));
		return false;
	}
	WKS4Format formatManager(*this, formatInput);
	return formatManager.parse();
}

bool WKS4Parser::readZones()
{
	RVNGInputStreamPtr input = getInput();
	input->seek(0, librevenge::RVNG_SEEK_SET);
	if (version()>=1000)
	{
		// error ok, we do no known how to parsed this structure
		while (!input->isEnd())
		{
			if (!readZoneQuattro())
				break;
		}

		ascii().addPos(input->tell());
		ascii().addNote("Entries(UnknownZone):");
		return false;
	}

	while (readZone())
	{
		if (m_state->m_isEncrypted && !m_state->m_isDecoded)
			throw(libwps::PasswordException());
	}

	//
	// look for ending
	//
	input = getInput();
	long pos = input->tell();
	if (!checkFilePosition(pos+4))
	{
		WPS_DEBUG_MSG(("WKS4Parser::readZones: cell header is too short\n"));
		return m_spreadsheetParser->getNumSpreadsheets()>0;
	}
	auto type = int(libwps::readU16(input)); // 1
	auto length = int(libwps::readU16(input));
	if (length)
	{
		WPS_DEBUG_MSG(("WKS4Parser::readZones: parse breaks before ending\n"));
		ascii().addPos(pos);
		ascii().addNote("Entries(BAD):###");
		return m_spreadsheetParser->getNumSpreadsheets()>0;
	}

	ascii().addPos(pos);
	if (type != 1)
	{
		WPS_DEBUG_MSG(("WKS4Parser::readZones: odd end cell type: %d\n", type));
		ascii().addNote("Entries(BAD):###");
	}
	else
		ascii().addNote("__End");

	return true;
}

bool WKS4Parser::readZone()
{
	libwps::DebugStream f;
	RVNGInputStreamPtr input = getInput();
	long pos = input->tell();
	auto id = int(libwps::readU8(input));
	auto type = int(libwps::read8(input));
	auto sz = long(libwps::readU16(input));
	if (sz<0 || !checkFilePosition(pos+4+sz))
	{
		WPS_DEBUG_MSG(("WKS4Parser::readZone: size is bad\n"));
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}

	f << "Entries(Struct";
	if (type == 0x54) f << "A";
	f << std::hex << id << std::dec << "E):";
	bool ok = true, isParsed = false, needWriteInAscii = false;
	int val;
	input->seek(pos, librevenge::RVNG_SEEK_SET);
	switch (type)
	{
	case 0:
		switch (id)
		{
		/* also
		   32: symphony windows settings(144)
		   37: password checksum(4)
		   3c: query(127)
		   3d: query name(16)
		   3e: symphony print record (679)
		   3f: printer name(16)
		   40: symphony graph record (499)
		   42: zoom(9)
		   43: number of split windows(2)
		   44: number of screen row(2)
		   45: number of screen column(2)
		   46: name ruler range(25)
		   47: name sheet range(25)
		   48: autoload comm(65)
		   49: autoexecutute macro adress(8)
		   4a: query parse information
		 */
		case 0:
			if (sz!=2) break;
			f.str("");
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			f << "version=" << std::hex << libwps::readU16(input) << std::dec << ",";
			isParsed=needWriteInAscii=true;
			break;
		case 0x1: // EOF
			ok = false;
			break;
		// boolean
		case 0x2: // Calculation mode 0 or FF
		case 0x3: // Calculation order
		case 0x4: // Split window type
		case 0x5: // Split window syn
		case 0x29: // label format 22|27|5e (spreadsheet)
		case 0x30: // formatted/unformatted print 0|ff (spreadsheet)
		case 0x31: // cursor/location 1|2
		case 0x38: // lock
			f.str("");
			f << "Entries(Byte" << std::hex << id << std::dec << "Z):";
			if (sz!=1) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			val=int(libwps::readU8(input));
			if (id==0x29)
				f << "val=" << std::hex << val << std::dec << ",";
			else if (id==0x31)
			{
				if (val!=1) f << val << ",";
			}
			else
			{
				if (val==0xFF) f << "true,";
				else if (val) f << "#val=" << val << ",";
			}
			isParsed=needWriteInAscii=true;
			break;
		case 0x6: // active worksheet range
			ok = m_spreadsheetParser->readSheetSize();
			isParsed = true;
			break;
		case 0x7: // window 1 record
		case 0x9: // window 2 record
			ok = readWindowRecord();
			isParsed=true;
			break;
		case 0x8: // col width
			ok = m_spreadsheetParser->readColumnSize();
			isParsed = true;
			break;
		case 0xa: // col width (window 2)
			if (sz!=3) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			// varies in this file from 0 to 5
			f << "id=" << libwps::read16(input) << ",";
			// small number 9-13: a dim?
			f << "dim?=" << libwps::read8(input) << ",";
			isParsed=needWriteInAscii=true;
			break;
		case 0xb: // named range
			ok = readFieldName();
			isParsed=true;
			break;
		case 0xc: // blank cell
		case 0xd: // integer cell
		case 0xe: // floating cell
		case 0xf: // label cell
		case 0x10: // formula cell
		case 0x36: // continue label
			ok = m_spreadsheetParser->readCell();
			isParsed = true;
			break;
		case 0x33:  // value of string formula
			ok = m_spreadsheetParser->readCellFormulaResult();
			isParsed = true;
			break;
		// some spreadsheet zone ( mainly flags )
		case 0x18: // data table range
		case 0x19: // query range
		case 0x20: // distribution range
		case 0x27: // print setup
		case 0x2a: // print borders
			ok = readUnknown1();
			isParsed=true;
			break;
		case 0x1a: // print range
		case 0x1b: // sort range
		case 0x1c: // fill range
		case 0x1d: // primary sort key range
		case 0x23: // secondary sort key range
		{
			int expectedSz=8;
			f.str("");
			switch (id)
			{
			case 0x1a: // only in spreadsheet?
				f << "Entries(PrintRange):";
				break;
			case 0x1b: // a dimension or also some big selection? 31999=infinity?, related to report?
				f << "Entries(SortRange):";
				break;
			case 0x1c: // a dimension or also some big selection? only in spreadsheet
				f << "Entries(FillRange):";
				break;
			case 0x1d:
				f << "Entries(PrimSort):";
				expectedSz=9;
				break;
			case 0x23:
				f << "Entries(SecSort):";
				expectedSz=9;
				break;
			default:
				break;
			}
			if (sz!=expectedSz) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			int dim[4];
			for (int &i : dim) i=int(libwps::read16(input));
			// in a spreadsheet, the cell or the cells corresponding to the field
			// in a database, col,0,col,0
			if (dim[0]==-1 && dim[1]==dim[0] && dim[2]==dim[0] && dim[3]==dim[0])
			{
			}
			else if (m_state->m_isSpreadsheet || dim[1] || dim[0]!= dim[2] || dim[3])
			{
				f << "cell=" << dim[0] << "x" << dim[1];
				if (dim[0]!=dim[2] || dim[1]!=dim[3])
					f << "<->" << dim[2] << "x" << dim[3];
				f << ",";
			}
			else
				f << "col=" << dim[0] << ",";
			if (expectedSz==9)
			{
				val=int(libwps::readU8(input)); // 0|1|ff
				if (val==0xFF) f << "true,";
				else if (val) f << "val=" << val << ",";
			}
			isParsed=needWriteInAscii=true;
			break;
		}
		case 0x24: // protection (global)
			f.str("");
			f << "Entries(Protection):global,";
			if (sz!=1) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			val=int(libwps::readU8(input));
			if (val==0)
			{
				f.str("");
				f << "_";
			}
			else if (val==0xFF) f << "protected,";
			else if (val) f << "#protected=" << val << ",";
			isParsed=needWriteInAscii=true;
			break;
		case 0x25: // footer
		case 0x26: // header
			readHeaderFooter(id==0x26);
			isParsed = true;
			break;
		case 0x28: // print margin
			if (sz!=10) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			for (int i=0; i<5; ++i)   // f1=4c|96|ac|f0
			{
				static const int expected[]= {4, 0x4c, 0x42, 2, 2};
				val=int(libwps::read16(input));
				if (val!=expected[i]) f << "f" << i << "=" << val << ",";
			}
			isParsed=needWriteInAscii=true;
			break;
		case 0x2d: // graph setting
		case 0x2e: // named graph setting
			m_chartParser->readChart();
			isParsed = true;
			break;
		case 0x2f: // iteration count: only in dos file Wk1, Wks(dos), Wq[12] ?
			if (sz!=1) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			f.str("");
			val=int(libwps::readU8(input));
			f << "Entries(ItCount):dos";
			if (val!=1) f << "=" << val << ",";
			else if (m_state->m_version==2)
				m_state->m_version=1;
			isParsed = needWriteInAscii = true;
			break;
		case 0x41: // graph record name
			m_chartParser->readChartName();
			isParsed = true;
			break;
		case 0x4b:
			if (sz==2 && m_state->m_creator==libwps::WPS_LOTUS)
			{
				m_state->m_isEncrypted=true;
				input->seek(pos+4, librevenge::RVNG_SEEK_SET);
				f.str("");
				uint16_t fileKey(libwps::readU16(input));
				f << "Entries(Password):pass=" << std::hex << fileKey << std::dec << ",";
				isParsed = needWriteInAscii = true;
				if (!m_state->m_isDecoded)
				{
					static uint8_t const defValues[]=
					{
						0xbb,0xff, 0xff,0xba, 0xff,0xff, 0xb9,0x80,
						0,0x0be, 0xf,0, 0xbf,0xf, 0,0
					};
					uint16_t key;
					std::vector<uint8_t> keys;
					if (m_state->m_password && libwps::encodeLotusPassword(m_state->m_password, key, keys, defValues))
					{
						RVNGInputStreamPtr newInput;
						if (uint16_t(key<<8|key>>8)==fileKey)
							newInput=decodeStream(input, m_state->m_eof, keys);
						if (newInput)
						{
							// let's replace the current input by the decoded input
							m_state->m_isDecoded=true;
							input=newInput;
							resetMainInput(newInput);
						}
						else
						{
							WPS_DEBUG_MSG(("WKS4Parser::parse: the password seems bad\n"));
						}

					}
				}
				break;
			}
			else
			{
				WPS_DEBUG_MSG(("WKS4Parser::parse: find unexpected password field\n"));
			}
			break;
		case 0x64: // hidden column
			isParsed = m_spreadsheetParser->readHiddenColumns();
			break;

		default:
			break;
		}
		break;
	case 0x54:
		switch (id)
		{
		// always empty ?
		case 0x25:
			f.str("");
			f << "Entries(LICS):";
			if (sz)
			{
				f << "###";
				WPS_DEBUG_MSG(("WKS4Parser::readZone: find a not empty LICS encoding zone\n"));
				break;
			}
			m_state->m_hasLICSCharacters = true;
			isParsed = needWriteInAscii = true;
			break;
		// boolean
		case 0x6f: // always 0
			f.str("");
			f << "Entries(ByteA" << std::hex << id << std::dec << "Z):";
			if (sz!=1) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			val=int(libwps::readU8(input));
			if (val==0xFF) f << "true,";
			else if (val) f << "#val=" << val << ",";
			isParsed=needWriteInAscii=true;
			break;
		// small int zone ?
		case 0x12: // sometimes in spreadsheet (with 0)
		case 0x1a: // find at at the end the file, after the reports' definition
			f.str("");
			f << "Entries(IntSmallA" << std::hex << id << std::dec << "Z):";
			if (sz!=1) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			val=int(libwps::readU8(input));
			if (id==0x1a)
			{
				f.str("");
				f << "Entries(Report):act=" << val << ",";
			}
			else
			{
				if (val) f << "#val=" << val << ",";
			}
			isParsed=needWriteInAscii=true;
			break;
		// int zone
		case 0x26: // always with 0
		case 0x6a: // filter definition?
			f.str("");
			if (id==0x6a)
				f << "Entries(Filter)[data1]:";
			else
				f << "Entries(IntA" << std::hex << id << std::dec << "Z):";
			if (sz!=2) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			val=int(libwps::read16(input));
			if (val) f << "f0=" << val << ",";
			isParsed=needWriteInAscii=true;
			break;
		//  zone with 2 ints
		case 0x32: // find with 00000000 (database)
			f.str("");
			f << "Entries(Int2A" << std::hex << id << std::dec << "Z):";
			if (sz!=4) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			for (int i=0; i<2; ++i)
			{
				val=int(libwps::read16(input));
				if (val) f << "f" << i << "=" << val << ",";
			}
			isParsed=needWriteInAscii=true;
			break;
		case 0x1: // the last selected cell
		{
			f.str("");
			f << "Entries(SelectCells):";
			if (sz!=0xc) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			val = int(libwps::read16(input)); // always 0?
			if (val) f << "f0=" << val << ",";
			int dim[4];
			for (int &i : dim) i = int(libwps::read16(input));
			if (dim[2]==dim[0]+1 && dim[3]==dim[1]+1) // almost always true
				f << "cell?=" << dim[0] << "x" << dim[1] << ",";
			else
				f << "cells?=" << dim[0] << "x" << dim[1] << "<->" << dim[2] << "x" << dim[3] << ",";
			val = int(libwps::read16(input)); // always 0|2
			if (val) f << "f1=" << val << ",";
			isParsed = needWriteInAscii = true;
			break;
		}
		case 0x2:
			ok = m_spreadsheetParser->readMsWorksDOSCellProperty();
			isParsed = true;
			break;
		case 0x5:
			if (sz!=2) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			f.str("");
			f << "Entries(Version):vers=" << std::hex << libwps::readU16(input) << std::dec;
			isParsed = needWriteInAscii = true;
			break;
		case 0x6:
			ok = m_spreadsheetParser->readMsWorksDOSFieldProperty();
			isParsed = true;
			break;
		case 0x8: // only in database?, checkme: the structure may be different in dosfile
			if (sz!=0x18) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			for (int i=0; i<6; ++i)   // f0=2|7, f2=0|1|2|4|5|7, f4=0|1|4|5|6|17|19|37|114, f5=3..40
			{
				val=int(libwps::read16(input));
				if (val) f << "f" << i << "=" << val << ",";
			}
			for (int i=0; i<4; ++i)   // g0=0|12|38|59|71, g1=0|1, g3=1|2|3
			{
				int const expected[]= {0,1,0,2};
				val=int(libwps::read8(input));
				if (val!=expected[i]) f << "g" << i << "=" << val << ",";
			}
			for (int i=0; i<4; ++i)   // h0=0|28, h1=0|9
			{
				val=int(libwps::read16(input));
				if (val) f << "h" << i << "=" << val << ",";
			}
			isParsed = needWriteInAscii = true;
			break;
		/* case 9: 000004002f001e000000bccf000005000f0008000000bccf0000060003000f007404bccf01000600
		   1c0000001b000100010007001c0001001e00010000000900300016007404bccf00000b000f0000000000d6ce
		   ( database, find one time)
		*/
		/* case a: (database)
		   CHECKME: a long structure which seems to contain some text, a list of field?
		 */
		case 0x10:
			ok = m_spreadsheetParser->readFilterOpen();
			isParsed = true;
			break;
		case 0x11:
			ok = m_spreadsheetParser->readFilterClose();
			isParsed = true;
			break;
		case 0x13:
			ok = m_spreadsheetParser->readMsWorksPageBreak();
			isParsed = true;
			break;
		case 0x14:
			m_chartParser->readChartAxis();
			isParsed = true;
			break;
		case 0x15:
			m_chartParser->readChartSeries();
			isParsed = true;
			break;
		case 0x16:
			m_chartParser->readChartSeriesStyles();
			isParsed = true;
			break;
		case 0x17:
			ok=m_spreadsheetParser->readReportOpen();
			isParsed = true;
			break;
		case 0x18:
			ok=m_spreadsheetParser->readReportClose();
			isParsed = true;
			break;

		case 0x30: // 30540c00000000000000000000000000 or 30540c00ffff00000000000000000000
			f.str("");
			f << "Entries(ChartUnknA):";
			break;
		case 0x31:
			m_chartParser->readChartSeriesColorMap();
			isParsed = true;
			break;
		case 0x35:
			m_chartParser->readChartDim();
			isParsed = true;
			break;
		case 0x38: // find block of f2 bytes always 0
		case 0x39: // find block of f2 bytes always 0
			f.str("");
			f << "Entries(" << (id==0x38 ? "ChartUnknB" : "ChartUnknC") << "):";
			break;
		case 0x41: // 41540600000000000000
			f.str("");
			f << "Entries(ChartUnknD):";
			break;
		case 0x44:
			m_chartParser->readChart3D();
			isParsed = true;
			break;
		/* find also 83: 835405000[02]ffffffff
		   85: block of 6c bytes always 0
		   86: 86540600010001000100
		   87: 87540c00880c00080a00080a00080a00 or 87540c00880c00880a18880a18880a18
		   related to chart in Works 4 and Works 2000 */

		case 0x19: // list id<->unkn, found after the column definition and Struct 66:54 in report, with f0=0|5|8|9
		case 0x5e: // some time in can be repeated a spreadsheet often with f0=9000
			if (id==0x19)
			{
				f.str("");
				f << "Report[data1]:";
			}
			if (sz!=4) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			f << "id=" << libwps::read16(input) << ",";
			val=int(libwps::readU16(input));
			if (val) f << "f0=" << std::hex << val << std::dec << ",";
			isParsed=needWriteInAscii=true;
			break;
		/* case 1b: 000000000000000000000000000000000000010000000000020000000000000000000000000000000000
		   database v1 */
		case 0x1c:
			m_spreadsheetParser->readMsWorksDOSCellExtraProperty();
			isParsed = true;
			break;
		case 0x1f:
		{
			// frequent field, near the beginning of the file
			// find with 050300000005|058000000005|05000000, so maybe:
			if (sz<4 || (sz%2)!=0) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			val=int(libwps::read8(input)); // always 5?
			if (val!=5) f << "f0=" << val << ",";
			val=int(libwps::readU8(input)); // 0|80
			if (val) f << "f1=" << std::hex << val << ",";
			for (long i=1; i<sz/2; ++i)
			{
				val=int(libwps::read16(input));
				if (val) f << "f" << i+2 << "=" << val << ",";
			}
			isParsed=needWriteInAscii=true;
			break;
		}
		case 0x23: // single page ?
		case 0x37: // multiple page ?
			ok = readPrnt();
			isParsed = true;
			break;
		case 0x24: // font (default)
			f.str("");
			f << "Entries(FontDef):";
			if (sz!=4) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			val=int(libwps::read16(input));
			if (val) f << "fId=" << val << ",";
			f << "fSz=" << libwps::read16(input)/2 << ",";
			isParsed=needWriteInAscii=true;
			break;
		case 0x27:
			ok = m_spreadsheetParser->readMsWorksDOSPageBreak();
			isParsed = true;
			break;
		case 0x33:
			f.str("");
			f << "Entries(Protection)[form]:";
			if (sz!=1) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			val=int(libwps::readU8(input));
			if (val==0)
			{
				f.str("");
				f << "_";
			}
			else if (val==0xFF) f << "protected,";
			else if (val) f << "#protected=" << val << ",";
			isParsed=needWriteInAscii=true;
			break;
		case 0x40:
			m_chartParser->readChartFont();
			isParsed = true;
			break;
		// case 47: big zone, begin by a font name (database)
		// case 50: 010000000000000000000000000000000000
		// case 53: CHECKME: looks like b013cc06d00764000000000001000000 ( database v1)
		case 0x56:
			ok = readFont();
			isParsed = true;
			break;
		case 0x48: // a fontname + 2 ints? (find one time in a spreadsheet file)
		case 0x57:   // int + a fontname + 2 ints? (in database a little after the field name zone)
		{
			int const headerSize= id==0x57 ? 2 : 0;
			f.str("");
			f << "Entries(Prefs)[" << std::hex << id << std::dec << "]:";
			if (sz!=0x24+headerSize) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			if (id==0x57)
			{
				val=int(libwps::read16(input)); // always 0?
				if (val) f << "f0=" << val << ",";
			}
			librevenge::RVNGString name;
			if (!readCString(name,32))
				f << "##name,";
			else if (!name.empty())
				f << name.cstr() << ",";
			input->seek(pos+36+headerSize, librevenge::RVNG_SEEK_SET);
			val=int(libwps::read16(input)); // 10|20|30|50: some flags?
			if (val!=0x10) f << "f1=" << std::hex << val << std::dec << ",";
			val=int(libwps::read16(input)); // 14|18
			if (val!=0x18) f << "f2=" << std::hex << val << std::dec << ",";
			isParsed=needWriteInAscii=true;
			break;
		}
		case 0x58:
		{
			f.str("");
			f << "Entries(Filter)[name]:";
			if (sz!=16) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			librevenge::RVNGString name;
			if (!readCString(name,16))
				f << "##name,";
			else if (!name.empty())
				f << name.cstr() << ",";
			isParsed=needWriteInAscii=true;
			break;
		}
		case 0x5a:
			ok = m_spreadsheetParser->readMsWorksStyle();
			isParsed = true;
			break;
		case 0x5b:
			ok = m_spreadsheetParser->readCell();
			isParsed = true;
			break;
		// case 5c: a small number 0-8 (database)
		case 0x5d: // checkme
			f.str("");
			f << "FldProperties:";
			if (sz!=4) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			f << "col=" << libwps::read16(input) << ",";
			f << "form?=" << std::hex << libwps::readU16(input) << std::dec << ",";
			break;
		case 0x5f:
		{
			// fixme: read end of fields
			f.str("");
			f << "Entries(FormZones):";
			if (sz<0x4d) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			int zType=libwps::read16(input);
			switch (zType)
			{
			case 1:
				f << "field,";
				break;
			case 2:
				f << "textbox,";
				break;
			case 3:
				f << "object,";
				break;
			case 4:
				f << "rectangle,";
				break;
			default:
				WPS_DEBUG_MSG(("WKS4Parser::readZone: find unknown zone type\n"));
				f << "##type=" << zType << ",";
				break;
			}
			if (input->tell()!=pos+4+sz)
				ascii().addDelimiter(input->tell(),'|');
			isParsed=needWriteInAscii=true;
			break;
		}
		case 0x64: // present in database (can to store some block: graphic?)
		{
			if (sz!=4) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			auto dataSz=long(libwps::readU32(input));
			if (!checkFilePosition(pos+8+dataSz)) break;
			if (dataSz) f << "dSz=" << std::hex << dataSz << std::dec << ",";
			ascii().addPos(pos);
			ascii().addNote(f.str().c_str());
			if (dataSz)
			{
				ascii().addPos(pos+8);
				ascii().addNote("Entries(StructA64E)[data]:");
				sz += dataSz;
			}
			isParsed = true;
			break;
		}
		case 0x65:
			ok = m_spreadsheetParser->readMsWorksRowSize();
			isParsed = true;
			break;
		// case 66: ff|12c|13B|1d1, dim/flag? (database)
		case 0x67: // single page ?
		case 0x82: // multiple page ?
			ok = readPrn2();
			isParsed = true;
			break;
		case 0x6b:
			ok = m_spreadsheetParser->readMsWorksColumnSize();
			isParsed = true;
			break;
		case 0x6e: // field(series)
			f.str("");
			f << "Entries(FldSeries):";
			if (sz!=8) break;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
			val=int(libwps::read16(input));
			if (val) f << "col=" << val << ",";
			f << "act[val]=" << libwps::read16(input) << ",";
			val=int(libwps::read16(input)); // always 0 first?
			if (val) f << "first=" << val << ",";
			val=int(libwps::read16(input));
			if (val!=1) f << "increm=" << val << ",";
			isParsed=needWriteInAscii=true;
			break;
		// case 70: id? (database)
		case 0x80:
		case 0x81:
			m_chartParser->readChartLimit();
			isParsed = true;
			break;
		case 0x84:
			m_chartParser->readChart2Font();
			isParsed = true;
			break;
		default:
			break;
		}
		break;
	default:
		ok = false;
		break;
	}

	if (!ok)
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	if (isParsed)
	{
		if (needWriteInAscii)
		{
			ascii().addPos(pos);
			ascii().addNote(f.str().c_str());
		}
		input->seek(pos+4+sz, librevenge::RVNG_SEEK_SET);
		return true;
	}

	if (sz && input->tell()!=pos && input->tell()!=pos+4+sz)
		ascii().addDelimiter(input->tell(),'|');
	input->seek(pos+4+sz, librevenge::RVNG_SEEK_SET);
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

////////////////////////////////////////////////////////////
//   other formats
////////////////////////////////////////////////////////////
bool WKS4Parser::readZoneQuattro()
{
	libwps::DebugStream f;
	RVNGInputStreamPtr input = getInput();
	long pos = input->tell();
	auto id = int(libwps::readU8(input));
	auto type = int(libwps::readU8(input));
	auto sz = long(libwps::readU16(input));
	if (type>5 || sz<0 || !checkFilePosition(pos+4+sz))
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	f << "Entries(Quattro";
	if (type) f << type << "A";
	f << std::hex << id << std::dec << "E):";
	if (sz) ascii().addDelimiter(pos+4,'|');
	input->seek(pos+4+sz, librevenge::RVNG_SEEK_SET);
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

////////////////////////////////////////////////////////////
//   generic
////////////////////////////////////////////////////////////
bool WKS4Parser::readCString(librevenge::RVNGString &string, long maxSize)
{
	RVNGInputStreamPtr input = getInput();
	long pos = input->tell();
	string.clear();
	if (!checkFilePosition(pos+maxSize))
	{
		WPS_DEBUG_MSG(("WKS4Parser::readCString: string's size seems bad\n"));
		return false;
	}
	std::string text;
	for (long i=0; i<maxSize; ++i)
	{
		auto c = char(libwps::readU8(input));
		if (c == '\0') break;
		text.push_back(c);
	}
	if (!text.empty())
		string=libwps_tools_win::Font::unicodeString(text, getDefaultFontType());
	return true;
}

bool WKS4Parser::readFont()
{
	libwps::DebugStream f;
	RVNGInputStreamPtr input = getInput();
	long pos = input->tell();
	auto type = int(libwps::read16(input));

	if (type != 0x5456)
	{
		WPS_DEBUG_MSG(("WKS4Parser::readFont: not a font zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos = pos+4+sz;
	if (sz < 32)
	{
		WPS_DEBUG_MSG(("WKS4Parser::readFont: seems very short\n"));
		ascii().addPos(pos);
		ascii().addNote("Entries(Font)###");
		return true;
	}

	WKS4ParserInternal::Font font(getDefaultFontType());
	auto flags = int(libwps::readU8(input));
	uint32_t attributes = 0;
	if (flags & 1) attributes |= WPS_BOLD_BIT;
	if (flags & 2) attributes |= WPS_ITALICS_BIT;
	if (flags & 4) attributes |= WPS_UNDERLINE_BIT;
	if (flags & 8) attributes |= WPS_STRIKEOUT_BIT;

	font.m_attributes=attributes;
	if (flags & 0xF0)
	{
		if (!m_state->getColor((flags >> 4), font.m_color))
		{
			WPS_DEBUG_MSG(("WKS4Parser::readFont: unknown color\n"));
			f << "##color=" << (flags >> 4) << ",";
		}
	}

	auto val=int(libwps::readU8(input));
	if (val) f << "f0=" << std::hex << val << std::dec << ",";
	librevenge::RVNGString name("");
	while (int(input->tell()) < endPos-4)
	{
		auto c = char(libwps::readU8(input));
		if (c == '\0') break;
		name.append(c);
	}

	font.m_type=libwps_tools_win::Font::getFontType(name);
	if (font.m_type==libwps_tools_win::Font::UNKNOWN)
		font.m_type=getDefaultFontType();
	font.m_name=name;

	input->seek(endPos-4, librevenge::RVNG_SEEK_SET);
	val = int(libwps::readU16(input)); // always 0x20
	if (val!=0x20)  f << "f1=" << std::hex << val << std::dec << ",";
	int fSize = libwps::read16(input)/2;
	if (fSize >= 1 && fSize <= 50)
		font.m_size=double(fSize);
	else
		f << "###fSize=" << fSize << ",";
	if (name.empty())
		f << "###noName,";
	font.m_extra=f.str();

	f.str("");
	f << "Entries(Font):";
	f << "font" << m_state->m_fontsList.size() << "[" << font << "]";
	m_state->m_fontsList.push_back(font);

	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	return true;
}

// ----------------------------------------------------------------------
// Header/Footer
// ----------------------------------------------------------------------
void WKS4Parser::sendHeaderFooter(bool header)
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("WKS4Parser::sendHeaderFooter: can not find the listener\n"));
		return;
	}

	m_listener->setFont(m_state->getDefaultFont());
	m_listener->insertUnicodeString(header ? m_state->m_headerString : m_state->m_footerString);
}

bool WKS4Parser::readHeaderFooter(bool header)
{
	libwps::DebugStream f;
	RVNGInputStreamPtr input = getInput();
	long pos = input->tell();
	auto type = int(libwps::read16(input));
	if (type != 0x0026 && type != 0x0025)
	{
		WPS_DEBUG_MSG(("WKS4Parser::readHeaderFooter: not a header/footer\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos = pos+4+sz;

	f << "Entries(" << (header ? "HeaderText" : "FooterText") << "):";
	if (sz==1)
	{
		// followed with 0
		auto val=int(libwps::read8(input));
		if (val) f << "##f0=" << val << ",";
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		return true;
	}
	if (sz < 0xF2)
	{
		WPS_DEBUG_MSG(("WKS4Parser::readHeaderFooter: the header/footer size seeem odds\n"));
		f << "###";
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		return false;
	}
	librevenge::RVNGString text;
	if (!readCString(text,sz))
		f << "##name,";
	if (header)
		m_state->m_headerString=text;
	else
		m_state->m_footerString=text;
	f << text.cstr();
	if (input->tell()!=endPos)
		ascii().addDelimiter(input->tell(), '|');
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	return true;
}

bool WKS4Parser::readPrnt()
{
	libwps::DebugStream f;
	RVNGInputStreamPtr input = getInput();
	long pos = input->tell();
	auto type = int(libwps::read16(input));
	if (type != 0x5423 && type != 0x5437)
	{
		WPS_DEBUG_MSG(("WKS4Parser::readPrnt: not a prnt zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos = pos+4+sz;

	f << "Entries(PRNT):";
	if (type==0x5437) f << "chart,";
	if (sz >= 12)
	{
		float dim[6];
		for (float &i : dim)
			i = float(libwps::read16(input))/1440.f;
		f << "dim=" << dim[5] << "x" << dim[4] << ",";
		f << "margin=[" << dim[0] << "x" << dim[2] << ","
		  << dim[3] << "x" << dim[1] << "],";
		// check me
		if (type==0x5423)
		{
			m_state->m_pageSpan.setFormWidth(double(dim[5]));
			m_state->m_pageSpan.setFormLength(double(dim[4]));
			m_state->m_pageSpan.setMarginLeft(double(dim[0]));
			m_state->m_pageSpan.setMarginTop(double(dim[2]));
			m_state->m_pageSpan.setMarginRight(double(dim[3]));
			m_state->m_pageSpan.setMarginBottom(double(dim[1]));
		}
	}
	int val = libwps::read16(input);
	if (val!=1) f << "first[pageNumber]=" << val <<",";
	long numElt = (endPos-input->tell())/2;
	for (long i = 0; i < numElt; i++)
	{
		// f2/3=0x2d0 (dim in inches ? )
		val = libwps::read16(input);
		if (!val) continue;
		f << "f" << i << "=" << std::hex << val << std::dec << ",";
	}

	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

bool WKS4Parser::readPrn2()
{
	libwps::DebugStream f;
	RVNGInputStreamPtr input = getInput();
	long pos = input->tell();
	long type = libwps::read16(input);
	if (type != 0x5482 && type != 0x5467)
	{
		WPS_DEBUG_MSG(("WKS4Parser::readPrn2: not a prn2 zone\n"));
		return false;
	}
	long sz = libwps::readU16(input);
	long endPos = pos+4+sz;

	f << "Entries(PRN2):";
	if (sz >= 64)
	{
		float dim[8];
		for (int st = 0; st < 2; st++)
		{
			for (float &i : dim)
				i = float(libwps::read32(input))/1440.f;
			f << "dim" << st << "=" << dim[5] << "x" << dim[4] << ",";
			f << "margin" << st << "=[" << dim[0] << "x" << dim[2] << ","
			  << dim[3] << "x" << dim[1] << "],";
			f << "head/foot" << st << "?=" << dim[7] << "x" << dim[6] << ",";
		}
	}
	long numElt = (endPos-input->tell())/4;
	/*
	  in general only f0=1,
	  but sometime f0=1,f2=1,f8=64,f42=174,f44=1,f46=175,f48=1
	 */
	for (long i = 0; i < numElt; i++)
	{
		auto val = int(libwps::read16(input));
		if (!val) continue;
		f << "f" << i << "=" << std::hex << val << std::dec << ",";
	}

	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

bool WKS4Parser::readFieldName()
{
	libwps::DebugStream f;
	RVNGInputStreamPtr input = getInput();

	long pos = input->tell();
	auto type = long(libwps::readU16(input));
	if (type != 0xb)
	{
		WPS_DEBUG_MSG(("WKS4Parser::readFieldName: not a zoneB type\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	f << "Entries(FldNames):";
	if (sz != 0x18 && sz != 0x1e)
	{
		// find also 0x85 a zone with 4 fldnames ?
		WPS_DEBUG_MSG(("WKS4Parser::readFieldName: size seems bad\n"));
		f << "###";
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		return true;
	}
	librevenge::RVNGString name;
	if (!readCString(name,16))
		f << "##name,";
	else if (!name.empty())
		f << name.cstr() << ",";

	input->seek(pos+20, librevenge::RVNG_SEEK_SET);
	// the position
	int dim[4];
	if (sz==0x18)
	{
		for (int &i : dim) i=int(libwps::read16(input));
	}
	else
	{
		for (int i=0; i<7; ++i)
		{
			auto val=int(libwps::read16(input));
			if (i<2) dim[i]=val;
			else if (i>=3 && i<5) dim[i-1]=val;
			else if (val) f << "f" << i << "=" << val << ",";
		}
	}
	// in a spreadsheet, the cell or the cells corresponding to the field
	// in a database, col,0,col,0xFFF
	if (m_state->m_isSpreadsheet || dim[1] || dim[0]!= dim[2] || dim[3]!=0xFFF)
	{
		f << "cell=" << dim[0] << "x" << dim[1];
		if (dim[0]!=dim[2] || dim[1]!=dim[3])
			f << "<->" << dim[2] << "x" << dim[3];
		f << ",";
	}
	else
		f << "col=" << dim[0] << ",";
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());
	return true;
}

////////////////////////////////////////////////////////////
//   Unknown
////////////////////////////////////////////////////////////


/* the zone 0:7 and 0:9 */
bool WKS4Parser::readWindowRecord()
{
	libwps::DebugStream f;
	RVNGInputStreamPtr input = getInput();

	long pos = input->tell();
	auto type = long(libwps::read16(input));
	if (type != 7 && type != 9)
	{
		WPS_DEBUG_MSG(("WKS4Parser::readWindowRecord: unknown type\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));

	// normally size=0x1f but one time 0x1e
	if (sz < 0x1e)
	{
		WPS_DEBUG_MSG(("WKS4Parser::readWindowRecord: zone seems too short\n"));
		ascii().addPos(pos);
		ascii().addNote("Entries(WindowRecord):###");
		return true;
	}

	f << "Entries(WindowRecord)[" << type << "]:";
	// f0=0-a|1f|21, f1=1-3c|1da, f2=0-6|71|7f|f1, f3=4|9|a|c(size?),
	// f4=0|4|6-11, f5=5-2c, f6=0|3|6|10|1f|20, f7=0-3c|1ca (related to f1?)
	// f8=0|1, f9=0|2, f10=f11=0
	for (int i=0; i<12; ++i)
	{
		auto val=int(libwps::read16(input));
		if (val) f << "f" << i << "=" << val << ",";
	}
	for (int i=0; i<2; ++i)   // g0=4, g1=4|a|b|10
	{
		auto val=int(libwps::read16(input));
		if (val!=4) f << "g" << i << "=" << val << ",";
	}
	auto val=int(libwps::read16(input)); // number between -5 and ad
	f << "g2=" << val << ",";

	if (sz!=0x1e)
		ascii().addDelimiter(input->tell(),'|');
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	return true;
}

/* some spreadsheet zones 0:18, 0:19, 0:20, 0:27, 0:2a */
bool WKS4Parser::readUnknown1()
{
	libwps::DebugStream f;
	RVNGInputStreamPtr input = getInput();

	long pos = input->tell();
	auto type = long(libwps::read16(input));
	int expectedSize=0, extraSize=0;
	switch (type)
	{
	case 0x18:
	case 0x19:
		expectedSize=0x19;
		break;
	case 0x20:
	case 0x2a:
		expectedSize=0x10;
		break;
	case 0x27:
		expectedSize=0x19;
		extraSize=15;
		break;
	default:
		WPS_DEBUG_MSG(("WKS4Parser::readUnknown1: unexpected type ???\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));

	f << "Entries(Flags" << std::hex << type << std::dec << ")]:";
	if (sz != expectedSize+extraSize)
	{
		// find also 270001000[01]
		if (type==0x27 && sz==1)
		{
			f << "f0=" << int(libwps::read8(input)) << ",";
			ascii().addPos(pos);
			ascii().addNote(f.str().c_str());
			return true;
		}
		WPS_DEBUG_MSG(("WKS4Parser::readUnknown1: the zone size seems too bad\n"));
		f << "###";
		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());
		return true;
	}

	// find always 0xff, excepted for zone 18(f0=0), zone 19(f24=0|3), zone 27(f0=..=f23=0|ff, f24=0|3)
	for (int i=0; i<expectedSize; ++i)
	{
		auto val=int(libwps::read8(input));
		if (val!=-1) f << "f" << i << "=" << val << ",";
	}

	if (type==0x27)
	{
		auto val=int(libwps::read8(input)); // always 0
		if (val) f << "g0=" << val << ",";
		for (int i=0; i<7; ++i)   // g1=0|4, g2=0|72, g4=0|20, g5=0|1|4, g6=0|1, g7=0|-1|205|80d8|e9f
		{
			val=int(libwps::read16(input));
			if (val) f << "g" << i+1 << "=" << val << ",";
		}
	}
	ascii().addPos(pos);
	ascii().addNote(f.str().c_str());

	return true;
}

////////////////////////////////////////////////////////////
//   decode
////////////////////////////////////////////////////////////
RVNGInputStreamPtr WKS4Parser::decodeStream(RVNGInputStreamPtr input, long endPos, std::vector<uint8_t> const &key)
{
	if (!input || key.size()!=16)
	{
		WPS_DEBUG_MSG(("WKS4Parser::decodeStream: the arguments seems bad\n"));
		return RVNGInputStreamPtr();
	}
	long actPos=input->tell();
	input->seek(0,librevenge::RVNG_SEEK_SET);
	librevenge::RVNGBinaryData data;
	if (!libwps::readDataToEnd(input, data) || long(data.size())!=endPos || !data.getDataBuffer())
	{
		WPS_DEBUG_MSG(("WKS4Parser::decodeStream: can not read the original input\n"));
		return RVNGInputStreamPtr();
	}
	auto *buf=const_cast<unsigned char *>(data.getDataBuffer());
	input->seek(actPos,librevenge::RVNG_SEEK_SET);
	uint8_t d7=0;
	while (!input->isEnd())
	{
		long pos=input->tell();
		if (pos+4>endPos) break;
		input->seek(2,librevenge::RVNG_SEEK_CUR);
		auto sSz=int(libwps::readU16(input));
		if (pos+4+sSz>endPos)
		{
			input->seek(pos,librevenge::RVNG_SEEK_SET);
			break;
		}
		for (int i=0; i<sSz; ++i)
		{
			auto c=uint8_t(libwps::readU8(input));
			c=uint8_t((c<<1)|(c>>7));
			c=(c^key[(d7++)&0xf]);
			buf[pos+4+i]=uint8_t((c>>6)|(c<<2));
		}
	}
	if (input->tell()!=endPos)
	{
		WPS_DEBUG_MSG(("WKS4Parser::decodeStream: can not decode the end of the file, data may be bad %lx %lx\n", static_cast<unsigned long>(input->tell()), static_cast<unsigned long>(endPos)));
	}
	RVNGInputStreamPtr res(new WPSStringStream(data.getDataBuffer(), static_cast<unsigned int>(endPos)));
	res->seek(actPos, librevenge::RVNG_SEEK_SET);
	return res;
}
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
