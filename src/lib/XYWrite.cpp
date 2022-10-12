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

#include <sstream>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WPSCell.h"
#include "WPSContentListener.h"
#include "WPSEntry.h"
#include "WPSFont.h"
#include "WPSHeader.h"
#include "WPSPageSpan.h"
#include "WPSParagraph.h"
#include "WPSPosition.h"
#include "WPSTable.h"
#include "WPSTextSubDocument.h"

#include "XYWrite.h"

namespace XYWriteParserInternal
{
/** Internal: class to store a basic cell with borders */
struct Cell final : public WPSCell
{
	//! constructor
	explicit Cell(XYWriteParser &parser)
		: WPSCell()
		, m_parser(parser)
		, m_entry()
		, m_style()
	{
	}
	//! virtual destructor
	~Cell() final;

	//! operator<<
	friend std::ostream &operator<<(std::ostream &o, Cell const &cell);
	//! call when a cell must be send
	bool send(WPSListenerPtr &listener) final
	{
		if (!listener) return true;
		auto *listen=dynamic_cast<WPSContentListener *>(listener.get());
		if (!listen)
		{
			WPS_DEBUG_MSG(("XYWriteParserInternal::Cell::send: unexpected listener\n"));
			return true;
		}
		listen->openTableCell(*this);
		sendContent(listener);
		listen->closeTableCell();
		return true;
	}

	//! call when the content of a cell must be send
	bool sendContent(WPSListenerPtr &) final
	{
		auto input=m_parser.getInput();
		if (!input)
		{
			WPS_DEBUG_MSG(("XYWriteParserInternal::Cell::sendContent: can not find the impu\n"));
			return true;
		}
		long pos=input->tell();
		m_parser.parseTextZone(m_entry, m_style);
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return true;
	}

	//! the actual parser
	XYWriteParser &m_parser;
	//! the text entry
	WPSEntry m_entry;
	//! the cell style
	std::string m_style;
};
Cell::~Cell()
{
}

//! a structure used to store a format
struct Format
{
	//! constructor
	explicit Format(bool inDosFile=false)
		: m_inDosFile(inDosFile)
		, m_string()
		, m_args()
		, m_entry()
		, m_isComplex(false)
		, m_listCounter(-1)
		, m_level()
		, m_children()
	{
	}
	//! return true if the format is empty
	bool empty() const
	{
		return m_string.empty() && m_args.empty();
	}
	//! returns the upper case of a string
	static std::string upperCase(std::string const &str)
	{
		std::string res(str);
		for (auto &c : res)
			c=char(std::toupper(c));
		return res;
	}
	//! returns the title in uppercase
	std::string title() const
	{
		return upperCase(m_string);
	}
	//! returns the two/... first character of the main string in uppercase
	std::string shortTitle(size_t sz=2) const
	{
		if (m_string.size()<=sz)
			return upperCase(m_string);
		return upperCase(m_string.substr(0,sz));
	}
	//! update a font if possible
	bool updateFont(WPSFont &font) const
	{
		bool ok=true;
		auto str=title();
		auto sTitle=shortTitle();
		if (sTitle=="MD")
		{
			if (str=="MDNM")
				font.m_attributes&=(unsigned int)(WPS_SMALL_CAPS_BIT|WPS_ALL_CAPS_BIT);
			else if (str=="MDBO" || str=="MD+BO")
				font.m_attributes|=WPS_BOLD_BIT;
			else if (str=="MD-BO")
				font.m_attributes&=~(unsigned int)(WPS_BOLD_BIT);
			else if (str=="MDBR" || str=="MD+BR")
				font.m_attributes|=(WPS_BOLD_BIT|WPS_ITALICS_BIT);
			else if (str=="MD-BR")
				font.m_attributes&=~(unsigned int)(WPS_BOLD_BIT|WPS_ITALICS_BIT);
			else if (str=="MDBU" || str=="MD+BU")
				font.m_attributes|=(WPS_BOLD_BIT|WPS_UNDERLINE_BIT);
			else if (str=="MD-BU")
				font.m_attributes&=~(unsigned int)(WPS_BOLD_BIT|WPS_UNDERLINE_BIT);
			else if (str=="MDDN" || str=="MD+DN")
				font.m_attributes|=WPS_STRIKEOUT_BIT;
			else if (str=="MD-DN")
				font.m_attributes&=~(unsigned int)(WPS_STRIKEOUT_BIT);
			else if (str=="MDRV" || str=="MD+IT")
				font.m_attributes|=WPS_ITALICS_BIT;
			else if (str=="MD-IT")
				font.m_attributes&=~(unsigned int)(WPS_ITALICS_BIT);
			else if (str=="MD+RV")
				font.m_attributes|=WPS_REVERSEVIDEO_BIT;
			else if (str=="MD-RV")
				font.m_attributes&=~(unsigned int)(WPS_REVERSEVIDEO_BIT);
			else if (str=="MDSD" || str=="MD+SD")
				font.m_attributes|=WPS_SUBSCRIPT_BIT;
			else if (str=="MD-SD")
				font.m_attributes&=~(unsigned int)(WPS_SUBSCRIPT_BIT);
			else if (str=="MDSU" || str=="MD+SU")
				font.m_attributes|=WPS_SUPERSCRIPT_BIT;
			else if (str=="MD-SU")
				font.m_attributes&=~(unsigned int)(WPS_SUPERSCRIPT_BIT);
			else if (str=="MDUL" || str=="MD+UL")
				font.m_attributes|=WPS_UNDERLINE_BIT;
			else if (str=="MD-UL")
				font.m_attributes&=~(unsigned int)(WPS_UNDERLINE_BIT);
			else
				ok=false;
		}
		else if (sTitle=="RG")
		{
			if (str=="RG0")
				font.m_attributes&=~(unsigned int)(WPS_SMALL_CAPS_BIT|WPS_ALL_CAPS_BIT);
			else if (str=="RG1")
			{
				font.m_attributes&=~(unsigned int)(WPS_SMALL_CAPS_BIT);
				font.m_attributes|=WPS_ALL_CAPS_BIT;
			}
			else if (str=="RG2")
			{
				font.m_attributes&=~(unsigned int)(WPS_ALL_CAPS_BIT);
				font.m_attributes|=WPS_SMALL_CAPS_BIT;
			}
			else
				ok=false;
		}
		else if (sTitle=="FG")
			ok=readColor(font.m_color);
		else if (sTitle=="SZ")
		{
			double value;
			bool inPoint;
			std::string extra;
			if (readUnit(str, 2, m_inDosFile, value, inPoint, extra) && inPoint)
				font.m_size=value;
		}

		if (!ok && !m_string.empty())
		{
			WPS_DEBUG_MSG(("XYWriteParserInternal::Format::updateFont: unknown format=%s\n", m_string.c_str()));
			return false;
		}
		return true;
	}
	//! update a paragraph if possible
	bool updateParagraph(WPSParagraph &para) const
	{
		bool ok=true;
		auto str=title();
		auto sTitle=shortTitle();
		double value=0;
		bool inPoint;
		std::string tmp;
		if (str=="FC")
			para.m_justify=libwps::JustificationCenter;
		else if (str=="FL")
			para.m_justify=libwps::JustificationLeft;
		else if (str=="FR")
			para.m_justify=libwps::JustificationRight;
		else if (str=="JU")
		{
			if (para.m_justify==libwps::JustificationLeft)
				para.m_justify=libwps::JustificationFull;
		}
		else if (str=="NJ")
		{
			if (para.m_justify==libwps::JustificationFull)
				para.m_justify=libwps::JustificationLeft;
		}
		else if (sTitle=="IP" || sTitle=="RM" || sTitle=="LS")
		{
			size_t p=2;
			if (readUnit(str, p, m_inDosFile, value, inPoint, tmp, sTitle!="LS"))
			{
				if (sTitle=="LS")
					para.setInterline(value, (!inPoint || (m_inDosFile && value<=3)) ? librevenge::RVNG_PERCENT : librevenge::RVNG_POINT, WPSParagraph::AtLeast);
				else if (inPoint)
				{
					if (sTitle=="IP")
						para.m_margins[0]=value/72;
					else if (!m_inDosFile || value<150) // in dos file, size from left
						para.m_margins[2]=value/72;
				}
				if (!tmp.empty())
				{
					WPS_DEBUG_MSG(("XYWriteParserInternal::Paragraph::updateParagraph: find extra data in %s\n", m_string.c_str()));
				}
			}
		}
		else if (sTitle=="AL")
		{
			if (str=="AL0" || str=="AL1")
				para.setInterline(1, librevenge::RVNG_PERCENT);
		}
		else if (sTitle=="NB" || sTitle=="BB")
		{
			if (str=="NB0" || str=="NB1" || str=="BB")
				para.m_breakStatus=0;
			else if (sTitle=="NB")   // no break
			{
				size_t p=2;
				if (readNumber(str, p, value))
					para.m_breakStatus=libwps::NoBreakWithNextBit;
				else
					ok=false;
			}
			else // allow break + data?
				ok=false;
		}
		else if (sTitle=="LL" || sTitle=="EL")
		{
			// some extra spacings
			for (size_t i=0; i<=(sTitle=="LL" ? std::min<size_t>(1,m_args.size()) : 0); ++i)
			{
				size_t p=i==0 ? 2 : 0;
				if (readUnit(i==0 ? str : m_args[i-1], p, m_inDosFile, value, inPoint, tmp))
				{
					if (sTitle=="EL") // after a line
						continue;
					// LL after para
					if (i==1)
						para.m_spacings[2]=inPoint ? value/72 : value*12/72;
					// FIXME update also para spacings 0 when i==0
				}
			}
		}
		else if (sTitle=="TS")
		{
			for (size_t i=0; i<=m_args.size(); ++i)
			{
				size_t p=i==0 ? 2 : 0;
				WPSTabStop tab;
				if (readUnit(i==0 ? m_string : m_args[i-1], p, m_inDosFile, value, inPoint, tmp))
				{
					if (!inPoint)
						continue;
					tmp=upperCase(tmp);
					if (tmp=="R")
						tab.m_alignment=WPSTabStop::RIGHT;
					else if (tmp=="C")
						tab.m_alignment=WPSTabStop::CENTER;
					else if (tmp=="D")
						tab.m_alignment=WPSTabStop::DECIMAL;
					else
						tab.m_alignment=WPSTabStop::LEFT;
					tab.m_position=value/72.;
					para.m_tabs.push_back(tab);
				}
			}
		}
		else if (sTitle=="BG")
			ok=readColor(para.m_backgroundColor);
		else
			ok=false;
		if (!ok && !m_string.empty())
		{
			WPS_DEBUG_MSG(("XYWriteParserInternal::Paragraph::updateParagraph: unknown format=%s\n", m_string.c_str()));
			return false;
		}
		return true;
	}
	//! operator<<
	friend std::ostream &operator<<(std::ostream &o, Format const &format)
	{
		o << format.m_string;
		if (!format.m_children.empty())
		{
			o << "\n";
			for (auto const &child : format.m_children)
				o << "\t" << child << "\n";
			return o;
		}
		if (format.m_isComplex && format.m_entry.valid())
			o << "[dt=" << format.m_entry.length() << "]";
		if (format.m_args.empty())
			return o;
		o << "[";
		for (auto const &arg : format.m_args)
			o << arg << ",";
		o << "]";
		return o;
	}
	//! flag to know if we are in a dos file
	bool m_inDosFile;
	//! the main part
	std::string m_string;
	//! the other arguments
	std::vector<std::string> m_args;
	//! a text zone entry
	WPSEntry m_entry;
	//! a flag to know if this is a complex entry
	bool m_isComplex;
	//! the list counter (if known)
	mutable int m_listCounter;
	//! the list level (if known)
	mutable WPSList::Level m_level;
	//! the list of child (for style, ...)
	std::vector<Format> m_children;

	////////////////////////////////
	// low level
	////////////////////////////////

	//! try to read a color
	bool readColor(WPSColor &color) const
	{
		if (m_args.size()!=2)
		{
			WPS_DEBUG_MSG(("XYWriteParserInternal::Format::readColor: bad number of argument\n"));
			return false;
		}
		unsigned char col[3];
		for (size_t i=0; i<3; ++i)
		{
			size_t p=i==0 ? 2 : 0;
			unsigned int val;
			std::string tmp;
			if (!readUInt(i==0 ? m_string : m_args[i-1], p, val, tmp) || val>255)
			{
				WPS_DEBUG_MSG(("XYWriteParserInternal::Format::readColor: can not read a component\n"));
				return false;
			}
			col[i]=(unsigned char)(val);
		}
		color=WPSColor(col[0],col[1],col[2]);
		return true;
	}
	//! try to read a box of double in points
	static bool readBox2f(std::string const &str, size_t i, bool inDosFile, WPSBox2f &box, std::string &extra)
	{
		Vec2f vec;
		std::string remain;
		for (int c=0; c<2; ++c)
		{
			if ((c==0 && (!readVec2f(str, i, inDosFile, vec, remain) || remain.substr(0,1)!=" ")) ||
			        (c==1 && (!readVec2f(remain, 1, inDosFile, vec, extra))))
				return false;
			if (c==0)
				box.setMin(vec);
			else
				box.setMax(vec);
		}
		return true;
	}
	//! try to read a vector of double in points
	static bool readVec2f(std::string const &str, size_t i, bool inDosFile, Vec2f &vec, std::string &extra)
	{
		double value=0;
		bool inPoint;
		std::string remain;
		for (int c=0; c<2; ++c)
		{
			if ((c==0 && (!readUnit(str, i, inDosFile, value, inPoint, remain) || remain.substr(0,1)!="x")) ||
			        (c==1 && (!readUnit(remain, 1, inDosFile, value, inPoint, extra))) || !inPoint)
				return false;
			vec[c]=float(value);
		}
		return true;
	}
	//! try to read a vector of uint
	static bool readVec2i(std::string const &str, size_t i, Vec2i &vec, std::string &extra)
	{
		unsigned int value=0;
		std::string remain;
		for (int c=0; c<2; ++c)
		{
			size_t p=c==0 ? i : 1;
			if ((c==0 && (!readUInt(str, p, value, remain) || remain.substr(0,1)!="x")) ||
			        (c==1 && (!readUInt(remain, p, value, extra))))
				return false;
			vec[c]=int(value);
		}
		return true;
	}
	//! try to read a int
	static bool readUInt(std::string const &str, size_t &i, unsigned int &value, std::string &extra)
	{
		auto len=str.size();
		value=0;
		size_t p=i;
		while (p<len && str[p]>='0' && str[p]<='9')
			value=value*10+(unsigned int)(str[p++]-'0');
		if (p==i)
			return false;
		i=p;
		if (i<len)
			extra=str.substr(i);
		return true;
	}
	//! try to read a double.
	static bool readNumber(std::string const &str, size_t &i, double &value)
	try
	{
		if (str.size()<=i)
		{
			if (!str.empty())
			{
				WPS_DEBUG_MSG(("XYWriteParserInternal::Format::readNumber: the string %s is too short\n", str.c_str()));
			}
			return false;
		}
		size_t p;
		if (i==0)
			value = std::stod(str,&p);
		else
			value = std::stod(str.substr(i),&p);
		i+=p;
		return true;
	}
	catch (...)
	{
		if (!str.empty())
		{
			WPS_DEBUG_MSG(("XYWriteParserInternal::Format::readNumber: can not extract number in %s\n", str.c_str()));
		}
		return false;
	}
	/** try to read an unit, returns a number in Point or in Line.
		\note if unit=AUTO set value to 72pt=1in
	 */
	static bool readUnit(std::string const &str, size_t i, bool inDosFile, double &value, bool &inPoint, std::string &extra, bool dosInChar=true)
	{
		if (str.size()>=i+4 && str.substr(i,4)=="AUTO")
		{
			value=72;
			inPoint=true;
			if (str.size()>i+4)
				extra=str.substr(i+4);
			return true;
		}
		if (!readNumber(str,i,value))
			return false;
		auto remain=upperCase(str.substr(i,2));
		if (i+2<str.size())
			extra=str.substr(i+2);
		inPoint=true;
		if (inDosFile && remain.empty())
		{
			if (dosInChar) value*=8;
			return true;
		}
		if (remain=="PT")
			return true;
		if (remain=="IN")
		{
			value *= 72;
			return true;
		}
		if (remain=="CM")
		{
			value *= 72/2.54;
			return true;
		}
		if (remain=="MM")
		{
			value *= 72/25.4;
			return true;
		}
		if (remain=="LI")
		{
			inPoint=false;
			return true;
		}
		if (inDosFile)
		{
			if (dosInChar) value*=8;
			extra=str.substr(i);
			return true;
		}
		if (!str.empty())
		{
			WPS_DEBUG_MSG(("XYWriteParserInternal::Format::readUnit: can not extract unit in %s\n", str.c_str()));
		}
		return false;

	}
};

//! the state of XYWrite
struct State
{
	//! constructor
	explicit State(libwps_tools_win::Font::Type fontType)
		: m_isDosFile(false)
		, m_eof(-1)
		, m_fontType(fontType)
		, m_metaData()
		, m_nameToStyleMap()
		, m_counterToTypeMap()
		, m_counterToValueMap()
	{
	}
	//! returns the current font type
	libwps_tools_win::Font::Type getFontType() const
	{
		if (m_fontType!=libwps_tools_win::Font::UNKNOWN)
			return m_fontType;
		// checkme
		return m_isDosFile ? libwps_tools_win::Font::CP_437 :
		       libwps_tools_win::Font::WIN3_WEUROPE;
	}
	//! flag to know if the file is or not a dos file
	bool m_isDosFile;
	//! the last file position
	long m_eof;
	//! the user font type
	libwps_tools_win::Font::Type m_fontType;
	//! the meta data
	librevenge::RVNGPropertyList m_metaData;
	//! map name to style
	std::map<std::string, Format> m_nameToStyleMap;
	//! map counter to type
	std::map<int, libwps::NumberingType> m_counterToTypeMap;
	//! map counter to value
	std::map<int, int> m_counterToValueMap;
};

//! Internal: the subdocument of a XYWriteParser
class SubDocument final : public WPSTextSubDocument
{
public:
	//! constructor for a text entry
	SubDocument(RVNGInputStreamPtr const &input, XYWriteParser &pars, WPSEntry const &entry, std::string const &style="")
		: WPSTextSubDocument(input, &pars), m_entry(entry), m_style(style) {}
	//! destructor
	~SubDocument() final {}

	//! operator==
	bool operator==(std::shared_ptr<WPSSubDocument> const &doc) const final
	{
		if (!doc || !WPSTextSubDocument::operator==(doc))
			return false;
		auto const *sDoc = dynamic_cast<SubDocument const *>(doc.get());
		if (!sDoc) return false;
		return m_entry == sDoc->m_entry && m_style == sDoc->m_style;
	}

	//! the parser function
	void parse(std::shared_ptr<WPSContentListener> &listener, libwps::SubDocumentType subDocumentType) final;
	//! the entry
	WPSEntry m_entry;
	//! the style
	std::string m_style;
};

void SubDocument::parse(std::shared_ptr<WPSContentListener> &listener, libwps::SubDocumentType /*subDocumentType*/)
{
	if (!listener.get())
	{
		WPS_DEBUG_MSG(("XYWriteParserInternal::SubDocument::parse: no listener\n"));
		return;
	}

	if (!m_parser)
	{
		listener->insertCharacter(' ');
		WPS_DEBUG_MSG(("XYWriteParserInternal::SubDocument::parse: bad parser\n"));
		return;
	}
	if (!m_entry.valid() || !m_input)
	{
		listener->insertCharacter(' ');
		return;
	}

	auto *mnParser = dynamic_cast<XYWriteParser *>(m_parser);
	if (!mnParser)
	{
		WPS_DEBUG_MSG(("XYWriteParserInternal::SubDocument::parse: bad parser...\n"));
		listener->insertCharacter(' ');
		return;
	}
	long pos=m_input->tell();
	mnParser->parseTextZone(m_entry, m_style);
	m_input->seek(pos, librevenge::RVNG_SEEK_SET);
}
}

// constructor, destructor
XYWriteParser::XYWriteParser(RVNGInputStreamPtr &input, WPSHeaderPtr &header,
                             libwps_tools_win::Font::Type encoding)
	: WPSParser(input, header)
	, m_listener()
	, m_state()
{
	m_state.reset(new XYWriteParserInternal::State(encoding));
}

XYWriteParser::~XYWriteParser()
{
}

bool XYWriteParser::checkFilePosition(long pos) const
{
	if (m_state->m_eof < 0)
	{
		RVNGInputStreamPtr input = const_cast<XYWriteParser *>(this)->getInput();
		long actPos = input->tell();
		input->seek(0, librevenge::RVNG_SEEK_END);
		m_state->m_eof=input->tell();
		input->seek(actPos, librevenge::RVNG_SEEK_SET);
	}
	return pos>=0 && pos <= m_state->m_eof;
}

// listener
std::shared_ptr<WPSContentListener> XYWriteParser::createListener(librevenge::RVNGTextInterface *interface)
{
	auto input=getInput();
	if (!input)
		throw "Can not find input";
	std::vector<WPSPageSpan> pageList;
	WPSPageSpan ps;
	ps.setMarginLeft(0.1);
	ps.setMarginRight(0.1);
	ps.setMarginTop(0.1);
	ps.setMarginBottom(0.1);
	input->seek(0, librevenge::RVNG_SEEK_SET);
	while (!input->isEnd() && input->tell()<m_state->m_eof)
	{
		uint8_t c=libwps::readU8(input);
		if (c==0x1a)
			break;
		if (c!=0xae)
			continue;
		XYWriteParserInternal::Format format;
		if (!parseFormat(format))
			break;
		std::string const str=format.title();
		std::string const sTitle=format.shortTitle();
		if (str=="PG")
			pageList.push_back(ps);
		else if (sTitle=="RH" || sTitle=="RF")
		{
			if (str.size()==2)
			{
				WPSEntry fEntry=format.m_entry;
				long fEnd=fEntry.end();
				fEntry.setBegin(fEntry.begin()+2);
				fEntry.setEnd(fEnd);
				WPSSubDocumentPtr subdoc(new XYWriteParserInternal::SubDocument
				                         (getInput(), *this,fEntry));
				ps.setHeaderFooter(sTitle=="RH" ? WPSPageSpan::HEADER : WPSPageSpan::FOOTER,
				                   WPSPageSpan::ALL, subdoc);
			}
			else if (str.size()>=3 && (str[2]=='A' || str[2]=='E' || str[2]=='O'))
			{
				WPSEntry fEntry=format.m_entry;
				long fEnd=fEntry.end();
				fEntry.setBegin(fEntry.begin()+3);
				fEntry.setEnd(fEnd);
				WPSSubDocumentPtr subdoc(new XYWriteParserInternal::SubDocument
				                         (getInput(), *this,fEntry));
				ps.setHeaderFooter(sTitle=="RH" ? WPSPageSpan::HEADER : WPSPageSpan::FOOTER,
				                   str[2]=='A' ? WPSPageSpan::ALL : str[2]=='E' ? WPSPageSpan::EVEN : WPSPageSpan::ODD, subdoc);
			}
		}
		else if (sTitle=="PW" || sTitle=="FD")
		{
			// PB is probably related to the page type A4 paysage==82?
			double val;
			bool inPoint;
			std::string tmp;
			if (format.readUnit(str, 2, m_state->m_isDosFile, val, inPoint, tmp) && inPoint && val>0 && val<100*72)
			{
				if (sTitle=="FD")
					ps.setFormLength(val/72.);
				else
					ps.setFormWidth(val/72.);
			}
			else
			{
				WPS_DEBUG_MSG(("XYWriteParser::createListener: can not parse %s\n", format.m_string.c_str()));
			}
		}
		else if (sTitle=="OF" || sTitle=="TP" || sTitle=="BT")
		{
			for (size_t j=0; j<=format.m_args.size(); ++j)
			{
				size_t p=j==0 ? 2 : 0;
				double value;
				bool inPoint;
				std::string tmp;
				if (format.readUnit(j==0 ? str : format.m_args[j-1], p, m_state->m_isDosFile, value, inPoint, tmp))
				{
					if (sTitle=="OF" && inPoint)
					{
						if (j==0)
							ps.setMarginLeft(value/72);
						else
							ps.setMarginRight(value/72);
					}
					else if (sTitle=="TP" && inPoint)
					{
						if (j==1) // i==0 bef header
							ps.setMarginTop(value/72);
					}
					else if (sTitle=="BT" && inPoint)
					{
						if (j==2) // i==0 aft footer, i==1 min, 3 max
							ps.setMarginBottom(value/72);
					}
				}
				else
				{
					WPS_DEBUG_MSG(("XYWriteParser::createListener: can not parse %s\n", format.m_string.c_str()));
				}
			}

		}
	}
	pageList.push_back(ps);
	auto listener=std::make_shared<WPSContentListener>(pageList, interface);
	listener->setMetaData(m_state->m_metaData);
	return listener;
}

////////////////////////////////////////////////////////////
// main funtions to parse a document
////////////////////////////////////////////////////////////

// main function to parse the document
void XYWriteParser::parse(librevenge::RVNGTextInterface *documentInterface)
{
	RVNGInputStreamPtr input=getInput();
	if (!input)
	{
		WPS_DEBUG_MSG(("XYWriteParser::parse: does not find main input\n"));
		throw (libwps::ParseException());
	}
	if (!checkHeader(nullptr, true))
		throw (libwps::ParseException());

	ascii().setStream(input);
	ascii().open("MN0");
	try
	{
		if (!m_state->m_isDosFile && !findAllZones())
			throw (libwps::ParseException());
		m_listener=createListener(documentInterface);
		if (!m_listener)
		{
			WPS_DEBUG_MSG(("XYWriteParser::parse: can not create the listener\n"));
			throw (libwps::ParseException());
		}
		m_listener->startDocument();
		WPSEntry entry;
		entry.setBegin(0);
		entry.setEnd(m_state->m_eof);
		parseTextZone(entry);
		m_listener->endDocument();
	}
	catch (...)
	{
		WPS_DEBUG_MSG(("XYWriteParser::parse: exception catched when parsing the main document\n"));
		throw (libwps::ParseException());
	}

	m_listener.reset();

	ascii().reset();
}

bool XYWriteParser::findAllZones()
{
	RVNGInputStreamPtr input = getInput();
	if (!input)
	{
		WPS_DEBUG_MSG(("XYWriteParser::findAllZones: can not find the input\n"));
		return false;
	}
	input->seek(0, librevenge::RVNG_SEEK_SET);
	bool ok=false;
	while (!input->isEnd())
	{
		if (libwps::readU8(input)==0x1a)
		{
			ok=true;
			break;
		}
	}
	if (!ok)
	{
		WPS_DEBUG_MSG(("XYWriteParser::findAllZones: can not find the end main zone marker\n"));
		return false;
	}
	long endZone1=input->tell();
	while (!input->isEnd())
	{
		if (libwps::readU8(input)==0x1a)
		{
			WPSEntry entry;
			entry.setBegin(endZone1);
			entry.setEnd(input->tell());
			parseMetaData(entry);
			input->seek(entry.end(), librevenge::RVNG_SEEK_SET);
			break;
		}
	}
	// now normally 22 02 fe fc fe 01 00
	if (m_state->m_eof!=input->tell()+7)
	{
		WPS_DEBUG_MSG(("XYWriteParser::findAllZones: end of file seems bad\n"));
	}
	// ok, update the end of the main zone and returns
	m_state->m_eof=endZone1-1;
	return true;
}

bool XYWriteParser::update(XYWriteParserInternal::Format const &format, libwps_tools_win::Font::Type &fontType) const
{
	if (!m_listener)
		throw "no listener";
	std::string const str=format.title();
	std::string const sTitle=format.shortTitle();
	if (sTitle=="MD" || sTitle=="RG" || sTitle=="SZ" ||
	        sTitle=="FG")
	{
		auto font=m_listener->getFont();
		if (!format.updateFont(font))
			return false;
		m_listener->setFont(font);
	}
	else if (sTitle=="UF" && str.size()>2)
	{
		auto font=m_listener->getFont();
		font.m_name=libwps_tools_win::Font::unicodeString(format.m_string.substr(2), m_state->getFontType());
		auto newType=libwps_tools_win::Font::getFontType(font.m_name);
		if (newType!=libwps_tools_win::Font::UNKNOWN)
			fontType=newType;
		m_listener->setFont(font);
	}
	else if (str=="FC" || str=="FL" || str=="FR" ||
	         str=="JU" || str=="NJ" ||
	         sTitle=="IP" || sTitle=="RM" ||
	         sTitle=="AL" || sTitle=="LS" || sTitle=="BB" || sTitle=="NB" ||
	         sTitle=="EL" || sTitle=="LL" || sTitle=="TS" ||
	         sTitle=="BG")
	{
		auto paragraph=m_listener->getParagraph();
		if (!format.updateParagraph(paragraph))
			return false;
		m_listener->setParagraph(paragraph);
	}
	else if (sTitle=="LM")
	{
		size_t p=2;
		double value;
		std::string tmp;
		bool inPoint=false;
		if (!format.readUnit(str, p, m_state->m_isDosFile, value, inPoint, tmp))
			return false;
		auto paragraph=m_listener->getParagraph();
		if (inPoint)
			paragraph.m_margins[1]=value/72;
		if (format.m_listCounter==-1 && !tmp.empty()&&format.m_entry.valid())
		{
			format.m_listCounter=-2;
			// look for a counter
			auto THIS=const_cast<XYWriteParser *>(this);
			auto input=THIS->getInput();
			if (!input)
				throw "XYWriteParser::update: no input";
			long actPos=input->tell();
			bool ok=false;
			input->seek(format.m_entry.begin(), librevenge::RVNG_SEEK_SET);
			while (!input->isEnd() && input->tell()<format.m_entry.end())
			{
				if (libwps::readU8(input)==';')
				{
					ok=true;
					break;
				}
			}
			if (ok)
			{
				ok=false;
				auto defType=m_state->getFontType();
				while (!input->isEnd() && input->tell()<format.m_entry.end())
				{
					auto ch=libwps::readU8(input);
					if (ch!=0xae)
					{
						if (ch=='\t') ch=' '; // replace tabulations by space
						libwps::appendUnicode(uint32_t(libwps_tools_win::Font::unicode(ch, defType)),
						                      format.m_listCounter >= 0 ? format.m_level.m_suffix :
						                      format.m_level.m_prefix);
						continue;
					}
					XYWriteParserInternal::Format newFormat;
					if (!THIS->parseFormat(newFormat))
						break;
					p=1;
					unsigned int val;
					std::string const newFStr=newFormat.title();
					std::string extra;
					if (!newFStr.empty() && newFStr[0]=='C' &&
					        newFormat.readUInt(newFStr, p, val, extra))
					{
						if (format.m_listCounter>=0)
						{
							format.m_level.m_prefix.clear();
							format.m_level.m_suffix.clear();
						}
						format.m_listCounter=int(val);
						auto it=m_state->m_counterToTypeMap.find(format.m_listCounter);
						if (it!=m_state->m_counterToTypeMap.end())
						{
							format.m_level.m_type=it->second;
						}
						else
						{
							WPS_DEBUG_MSG(("XYWriteParser::update: can not find counter %d\n", format.m_listCounter));
						}
						ok=true;
					}
					else
					{
#ifdef DEBUG
						std::cerr << "XYWriteParser::update[LM]: unused\n";
						std::cerr << "\t" << newFormat << "\n";
#endif
					}
				}
			}
			if (!ok)
			{
#ifdef DEBUG
				std::cerr << "XYWriteParser::update[LM]: unused\n";
				std::cerr << "\t" << format << "\n";
#endif
			}
			input->seek(actPos, librevenge::RVNG_SEEK_SET);
		}
		if (!format.m_level.isDefault())
		{
			format.m_level.m_labelIndent=paragraph.m_margins[1];
			paragraph.m_listLevel=format.m_level;
			paragraph.m_listLevelIndex=1;
			paragraph.m_margins[1]=0;
		}
		m_listener->setParagraph(paragraph);
	}
	else
		return false;
	return true;
}

bool XYWriteParser::parseTextZone(WPSEntry const &entry, std::string const &styleName)
{
	RVNGInputStreamPtr input = getInput();
	if (!input || !m_listener)
		throw (libwps::ParseException());
	if (!entry.valid())
		return true;
	input->seek(entry.begin(), librevenge::RVNG_SEEK_SET);
	auto fontType=m_state->getFontType();
	WPSFont defFont;
	defFont.m_name="Courier New";
	defFont.m_size=10;
	m_listener->setFont(defFont);
	if (!styleName.empty())
	{
		auto it=m_state->m_nameToStyleMap.find(styleName);
		if (it!=m_state->m_nameToStyleMap.end())
		{
			for (auto const &child : it->second.m_children)
			{
				std::string const sTitle=child.shortTitle();
				if (!update(child, fontType) && sTitle!="FT" && sTitle!="BF") // checkme: FT and BF related to footnote
				{
#ifdef DEBUG
					std::cerr << "XYWriteParser::parseTextZone[child]: unused\n";
					std::cerr << "\t" << child << "\n";
#endif
				}
			}
		}
	}
	while (!input->isEnd() && input->tell()<entry.end())
	{
		uint8_t c=libwps::readU8(input);
		if (c==0x1a)
		{
			if (input->tell()<entry.end())
			{
				WPS_DEBUG_MSG(("XYWriteParser::parseTextZone: find unexpected end zone\n"));
			}
			break;
		}
		if (c==0xae)
		{
			XYWriteParserInternal::Format format;
			if (!parseFormat(format))
				throw (libwps::ParseException());
			std::string const str=format.title();
			std::string const sTitle=format.shortTitle();
			bool done=true;
			if (update(format, fontType))
				;
			else if (sTitle=="DC" && str.size()>2) // DCxxx=[1iIaA]...
			{
				size_t p=2;
				unsigned int val;
				std::string extra;
				if (!format.readUInt(format.m_string, p, val, extra) || !format.m_entry.valid() ||
				        extra.size()<=1 || extra[0]!='=' ||
				        m_state->m_counterToTypeMap.find(int(val)) != m_state->m_counterToTypeMap.end())
					done=false;
				else
				{
					std::map<char, libwps::NumberingType> const cToTypeMap=
					{
						{'1', libwps::ARABIC}, {'a', libwps::LOWERCASE}, {'A', libwps::UPPERCASE}, {'i', libwps::LOWERCASE_ROMAN}, {'I', libwps::UPPERCASE_ROMAN}
					};
					auto it=cToTypeMap.find(extra[1]);
					if (it==cToTypeMap.end())
					{
						WPS_DEBUG_MSG(("XYWriteParserInternal::Format::parseTextZone: can not decode counter format in %s\n", extra.c_str()));
					}
					else
						m_state->m_counterToTypeMap[int(val)]=it->second;
				}
			}
			else if (sTitle=="SS")
			{
				if (createFormatChildren(format,format.m_string.size()+1) && format.m_string.size()>2)
					m_state->m_nameToStyleMap[format.m_string.substr(2)]=format;
				else
					done=false;
			}
			else if (sTitle=="FM" && format.m_string.size()>2 &&
			         (format.m_string[2]>='1' && format.m_string[2]<='3'))
			{
				if (createFormatChildren(format,3))
				{
					std::string name("__");
					name+=format.m_string.substr(0,3);
					m_state->m_nameToStyleMap[name]=format;
				}
				else
					done=false;
			}
			else if (sTitle=="US" && format.m_string.size()>2)
			{
				auto it=m_state->m_nameToStyleMap.find(format.m_string.substr(2));
				if (it==m_state->m_nameToStyleMap.end())
				{
					WPS_DEBUG_MSG(("XYWriteParser::parseTextZone: can not find style %s\n", format.m_string.c_str()));
				}
				else
				{
					if (format.m_string.size()>4)
					{
						// unsure, when we need to reset the style
						m_listener->setFont(defFont);
						m_listener->setParagraph(WPSParagraph());
					}
					for (auto const &child : it->second.m_children)
					{
						if (!update(child, fontType))
						{
#ifdef DEBUG
							std::cerr << "XYWriteParser::parseTextZone[child]: unused\n";
							std::cerr << "\t" << child << "\n";
#endif
						}
					}
				}
			}
			else if (sTitle=="FA")
			{
				if (createFormatChildren(format))
					parseFrameZone(format);
				else
					done=false;
			}
			else if (sTitle=="IG")
			{
				if (createFormatChildren(format,format.m_string.size()+1))
					parsePictureZone(format);
				else
					done=false;
			}
			else if (sTitle=="NT" && format.m_string.size()>2 && format.m_entry.valid())
			{
				WPSEntry fEntry=format.m_entry;
				long fEnd=fEntry.end();
				fEntry.setBegin(fEntry.begin()+2);
				fEntry.setEnd(fEnd);
				WPSSubDocumentPtr subdoc(new XYWriteParserInternal::SubDocument
				                         (getInput(), *this,fEntry));
				m_listener->insertComment(subdoc);
			}
			else if (sTitle=="FN")
			{
				bool hasId=format.m_string.size()>2 &&
				           (format.m_string[2]>='1' && format.m_string[2]<='3');
				WPSEntry fEntry=format.m_entry;
				long fEnd=fEntry.end();
				fEntry.setBegin(fEntry.begin()+(hasId ? 3 : 2));
				fEntry.setEnd(fEnd);
				std::string sName("__FM");
				sName+=format.m_string[2];
				WPSSubDocumentPtr subdoc(new XYWriteParserInternal::SubDocument
				                         (getInput(), *this, fEntry, hasId ? sName : ""));
				m_listener->insertNote(WPSContentListener::FOOTNOTE, subdoc);
			}
			else if (str=="PG")
				m_listener->insertBreak(WPS_PAGE_BREAK);
			else if (str=="PN")
				m_listener->insertField(WPSField(WPSField::PageNumber));
			else if (str=="DA" || str=="TM") // TM date custom format ?
				m_listener->insertField(WPSField(WPSField::Date));
			else if (str=="TI")
				m_listener->insertField(WPSField(WPSField::Time));
			else if (sTitle=="CT")
			{
				long actPos=input->tell();
				if (!createTable(format, entry.end()))
					input->seek(actPos, librevenge::RVNG_SEEK_SET);
			}
			else if ((sTitle=="RH" || sTitle=="RF") && // header/footer already parsed
			         (str.size()==2 ||
			          (str.size()>=3 && (str[2]=='A'||str[2]=='E'||str[2]=='O'))))
				;
			else if (sTitle=="PW" || sTitle=="FD" || sTitle=="PB" || sTitle=="OF" || sTitle=="TP" || sTitle=="BT" || // page dimension
			         format.shortTitle(3)=="UBN" || sTitle=="GU" || sTitle=="EE" || sTitle=="ET" || // related to table (checkme)
			         sTitle=="SY" || // other method to define a font
			         sTitle=="NF") // footnote number
				;
			else if (sTitle=="LB" || sTitle=="RE")
			{
				// LBname: label
				// RE[PCF]name: ref to page, chapter, reference
				static bool first=true;
				if (first)
				{
					WPS_DEBUG_MSG(("XYWriteParser::parseTextZone: retrieving label/cross ref is not implemented\n"));
				}
			}
			else if (format.shortTitle(1)=="C")
			{
				unsigned int value;
				size_t p=1;
				std::string tmp;
				if (!format.readUInt(str, p, value, tmp))
					done=false;
				else
				{
					int actValue=0;
					auto it=m_state->m_counterToValueMap.find(int(value));
					if (it!=m_state->m_counterToValueMap.end())
						actValue=it->second;
					++actValue;
					std::stringstream s;
					s << actValue;
					for (auto res : s.str())
						m_listener->insertCharacter(uint8_t(res));
					m_state->m_counterToValueMap[int(value)]=actValue;
				}
			}
			else
				done=false;
			if (!done)
			{
#ifdef DEBUG
				std::cerr << "XYWriteParser::parseTextZone: unused\n";
				std::cerr << "\t" << format << "\n";
#endif
			}
			continue;
		}
		if (c==0xff && input->tell()+2<entry.end())
		{
			// special case a char in binary...
			long pos=input->tell();
			c=0;
			bool ok=true;
			for (int i=0; i<2; ++i)
			{
				auto ch=libwps::readU8(input);
				if (ch>='A' && ch<='F')
					c=(unsigned char)(c*16+(ch-'A'+10));
				else if (ch>='0' && ch<='9')
					c=(unsigned char)(c*16+(ch-'0'));
				else
				{
					ok=false;
					break;
				}
			}
			if (!ok)
			{
				input->seek(pos, librevenge::RVNG_SEEK_SET);
				WPS_DEBUG_MSG(("XYWriteParser::parseTextZone: find bad char FF in pos=%lx\n", (unsigned long)(input->tell())));
				continue;
			}
		}
		switch (c)
		{
		case 0x9:
			m_listener->insertTab();
			break;
		case 0xa:
			break;
		case 0xd:
		{
			// we must reset the list level to 0
			if (m_listener->getCurrentList())
			{
				auto paragraph=m_listener->getParagraph();
				if (paragraph.m_listLevelIndex>0)
				{
					paragraph.m_margins[1]=paragraph.m_listLevel.m_labelIndent;
					paragraph.m_listLevelIndex=0;
					m_listener->setParagraph(paragraph);
				}
			}
			m_listener->insertEOL();
			break;
		}
		default:
			if (c<0x1f || ((c==0xaf || c==0xfa) && input->tell() != entry.begin()+1 && input->tell() != entry.end()))
			{
				WPS_DEBUG_MSG(("XYWriteParser::parseTextZone: find bad char %x in pos=%lx\n", unsigned(c), (unsigned long)(input->tell())));
			}
			else
				m_listener->insertUnicode(uint32_t(libwps_tools_win::Font::unicode(c, fontType)));
		}
	}
	return true;
}

bool XYWriteParser::parseFrameZone(XYWriteParserInternal::Format const &frameFormat)
{
	RVNGInputStreamPtr input = getInput();
	if (!input || !m_listener ||!frameFormat.m_entry.valid())
		throw (libwps::ParseException());
	Vec2f dim;
	WPSEntry textEntry;
	for (auto const &child : frameFormat.m_children)
	{
		std::string const sTitle=child.shortTitle();
		bool done=false;
		if (sTitle=="SI")
		{
			std::string tmp;
			done=child.readVec2f(child.m_string, 2, m_state->m_isDosFile, dim, tmp);
		}
		else if (sTitle=="LB") // label are not implement
			done=true;
		else if (sTitle=="PO")   // find POTMxPC;data
		{
			textEntry=child.m_entry;
			done=true;
		}
		if (!done)
		{
#ifdef DEBUG
			std::cerr << "XYWriteParser::parseFrameZone: unused\n";
			std::cerr << "\t" << child << "\n";
#endif
		}
	}
	if (dim[0]<=0 || dim[1]<=0 || !textEntry.valid())
	{
		WPS_DEBUG_MSG(("XYWriteParser::parseFrameZone: can not find frame data\n"));
		return false;
	}
	long begPos=input->tell();

	long endPos=textEntry.end();
	input->seek(textEntry.begin(), librevenge::RVNG_SEEK_SET);
	while (!input->isEnd() && input->tell()<endPos)
	{
		if (libwps::readU8(input)==';')
			break;
	}
	textEntry.setBegin(input->tell());
	textEntry.setEnd(endPos);
	WPSPosition fPos(Vec2f(), dim, librevenge::RVNG_POINT);
	fPos.setRelativePosition(WPSPosition::Char);
	WPSSubDocumentPtr subdoc(new XYWriteParserInternal::SubDocument
	                         (getInput(), *this, textEntry));
	m_listener->insertTextBox(fPos, subdoc);
	input->seek(begPos, librevenge::RVNG_SEEK_SET);
	return true;
}

bool XYWriteParser::parsePictureZone(XYWriteParserInternal::Format const &pictureFormat)
{
	RVNGInputStreamPtr input = getInput();
	if (!input || !m_listener || !pictureFormat.m_entry.valid())
		throw (libwps::ParseException());
	WPSBox2f box;
	Vec2i scaleArray(100,100);
	for (auto const &child : pictureFormat.m_children)
	{
		std::string const sTitle=child.shortTitle();
		bool done=false;
		if (sTitle=="CR")
		{
			std::string tmp;
			done=child.readBox2f(child.m_string, 2, m_state->m_isDosFile, box, tmp);
		}
		else if (sTitle=="TY" || child.m_string=="IML" || sTitle=="RV") // type, unknown, revision?
			done=true;
		else if (sTitle=="SC")
		{
			std::string tmp;
			done=child.readVec2i(child.m_string, 2, scaleArray, tmp);
		}
		if (!done)
		{
#ifdef DEBUG
			std::cerr << "XYWriteParser::parsePictureZone: unused\n";
			std::cerr << "\t" << child << "\n";
#endif
		}
	}
	Vec2f dim(static_cast<float>(scaleArray[0])/100.f*box.size()[0], float(scaleArray[1])/100.f*box.size()[1]);
	if (dim[0]<=0 || dim[1]<=0)
	{
		WPS_DEBUG_MSG(("XYWriteParser::parsePictureZone: can not find picture dimension\n"));
		return false;
	}

	long begPos=input->tell();
	long endPos=pictureFormat.m_entry.end();
	input->seek(pictureFormat.m_entry.begin()+2, librevenge::RVNG_SEEK_SET);
	while (!input->isEnd() && input->tell()<endPos)
	{
		if (libwps::readU8(input)==',')
			break;
	}
	WPSEntry textEntry;
	textEntry.setBegin(pictureFormat.m_entry.begin()+2);
	textEntry.setEnd(input->tell()-1);
	WPSPosition fPos(Vec2f(), dim, librevenge::RVNG_POINT);
	fPos.setRelativePosition(WPSPosition::Char);
	WPSSubDocumentPtr subdoc(new XYWriteParserInternal::SubDocument
	                         (getInput(), *this, textEntry));
	m_listener->insertTextBox(fPos, subdoc);
	input->seek(begPos, librevenge::RVNG_SEEK_SET);
	return true;
}

bool XYWriteParser::createTable(XYWriteParserInternal::Format const &tableFormat, long endPos)
{
	RVNGInputStreamPtr input = getInput();
	if (!input || !m_listener || tableFormat.shortTitle()!="CT")
		throw (libwps::ParseException());
	long begPos=input->tell();
	if (begPos>=endPos)
	{
		WPS_DEBUG_MSG(("XYWriteParser::createTable: the zone seems too short\n"));
		return false;
	}
	std::vector<float> colWidth;
	std::vector<std::string> colStyle;
	for (size_t j=0; j<=tableFormat.m_args.size(); ++j)
	{
		size_t p=j==0 ? 2 : 0;
		double value;
		bool inPoint;
		std::string tmp;
		if (tableFormat.readUnit(j==0 ? tableFormat.m_string : tableFormat.m_args[j-1], p, m_state->m_isDosFile, value, inPoint, tmp) && inPoint)
		{
			if (j==0) // left pos
				;
			else
			{
				colWidth.push_back(float(value));
				if (tmp.size()>1)
					colStyle.push_back(tmp.substr(1));
				else
					colStyle.push_back("");
			}
		}
		else
		{
			WPS_DEBUG_MSG(("XYWriteParser::createTable: can not read some colum size %s\n", j==0 ? tableFormat.m_string.c_str() : tableFormat.m_args[j-1].c_str()));
			if (j)
			{
				colWidth.push_back(0);
				colStyle.push_back("");
			}
		}
	}
	if (colWidth.empty())
	{
		WPS_DEBUG_MSG(("XYWriteParser::createTable: can not find any columns\n"));
		return false;
	}
	int numColumns=int(colWidth.size());

	std::vector<XYWriteParserInternal::Cell> cells;
	int cRow=0, cCol=0;
	cells.push_back(XYWriteParserInternal::Cell(*this));
	cells.back().m_entry.setBegin(input->tell());
	cells.back().setPosition(Vec2i(cCol,cRow));

	bool ok=false;
	while (!input->isEnd())
	{
		long pos=input->tell();
		if (pos>=endPos)
			break;
		uint8_t c=libwps::readU8(input);
		if (c==0x1a)
		{
			if (input->tell()<endPos)
			{
				WPS_DEBUG_MSG(("XYWriteParser::createTable: find unexpected end zone\n"));
			}
			break;
		}
		if (c!=0xae)
			continue;
		XYWriteParserInternal::Format format;
		if (!parseFormat(format))
			throw (libwps::ParseException());
		std::string const str=format.title();
		std::string const sTitle=format.shortTitle();
		if (str=="EC")
		{
			ok=true;
			break;
		}
		if (sTitle=="CT")
			break;
		else if (sTitle=="CO")
		{
			unsigned int value;
			size_t p=2;
			std::string tmp;
			if (!format.readUInt(format.m_string, p, value, tmp))
				continue;
			if (numColumns==0 || cells.empty() || value==0) // the table is not created
				break;
			cells.back().m_entry.setEnd(pos);
			int newCol=int(value)-1;
			if (newCol>numColumns)
				break;
			if (newCol<=cCol)
				++cRow;
			cCol=newCol;
			cells.push_back(XYWriteParserInternal::Cell(*this));
			cells.back().m_entry.setBegin(input->tell());
			cells.back().setPosition(Vec2i(cCol,cRow));
			if (cCol<int(colStyle.size()))
				cells.back().m_style=colStyle[size_t(cCol)];
		}
	}
	if (ok)
	{
		m_listener->openTable(colWidth, librevenge::RVNG_POINT);
		cRow=-1;
		WPSListenerPtr listener(m_listener);
		for (auto &cell : cells)
		{
			auto const cPos=cell.position();
			while (cRow<cPos[1])
			{
				if (cRow!=-1)
					m_listener->closeTableRow();
				m_listener->openTableRow(-10, librevenge::RVNG_POINT);
				++cRow;
				cCol=0;
			}
			if (cCol<cPos[0])
				m_listener->addEmptyTableCell(Vec2i(cCol,cRow),Vec2i(cPos[0]-cCol,1));
			m_listener->openTableCell(cell);
			cell.sendContent(listener);
			m_listener->closeTableCell();
			cCol=cPos[0]+1;
		}
		if (cRow!=-1)
			m_listener->closeTableRow();
		m_listener->closeTable();
	}
	else
		input->seek(begPos, librevenge::RVNG_SEEK_SET);
	return ok;
}

bool XYWriteParser::parseMetaData(WPSEntry const &entry)
{
	RVNGInputStreamPtr input = getInput();
	if (!input)
		throw (libwps::ParseException());
	if (!entry.valid())
		return true;
	input->seek(entry.begin(), librevenge::RVNG_SEEK_SET);
	XYWriteParserInternal::Format format;
	std::string actualString;
	auto fontType=m_state->getFontType();
	while (!input->isEnd() && input->tell()+1<entry.end())
	{
		uint8_t c=libwps::readU8(input);
		if (c==0x1a)
		{
			WPS_DEBUG_MSG(("XYWriteParser::parseMetaData: find end of zone\n"));
			return false;
		}
		if (c!=0xae)
		{
			actualString+=char(c);
			continue;
		}
		const auto strEnd = actualString.find_last_not_of(" ");
		actualString=actualString.substr(0, strEnd+1);
		if (!actualString.empty())
		{
			// also LBLG:40 another author
			// find also LBCD:20, LBCT:15, LBMD:20, LBMT:15 with some checksum?
			//      and LBRP:4, LBPJ:20, LBCM:44, LBKY:~250 empty
			auto const finalStr=libwps_tools_win::Font::unicodeString(actualString, fontType);
			std::string const str=format.title();
			if (str=="LBAU") // sz:40
				m_state->m_metaData.insert("dc:creator", finalStr);
			else if (str=="LBRV") // revision sz:4
				m_state->m_metaData.insert("librevenge:version-number", finalStr);
		}
		actualString.clear();
		if (!parseFormat(format))
			return false;
	}
	return true;
}

////////////////////////////////////////////////////////////
// low level
////////////////////////////////////////////////////////////

bool XYWriteParser::parseFormat(XYWriteParserInternal::Format &format)
{
	RVNGInputStreamPtr input = getInput();
	if (!input)
		throw (libwps::ParseException());
	format=XYWriteParserInternal::Format(m_state->m_isDosFile);
	format.m_entry.setBegin(input->tell());
	while (!input->isEnd())
	{
		if (input->tell()>=m_state->m_eof)
		{
			WPS_DEBUG_MSG(("XYWriteParser::parseFormat: can not find end of format\n"));
			return false;
		}
		uint8_t c=libwps::readU8(input);
		if (c==0xaf)
		{
			format.m_entry.setEnd(input->tell()-1);
			return true;
		}
		if (c==0xfa || c==0xae)
		{
			format.m_isComplex=true;
			// normally ae XXX fa ... af
			// but sometimes I found ae XXX ae ... af fa ... af
			if (format.m_entry.valid())
			{
				WPS_DEBUG_MSG(("XYWriteParser::parseFormat: oops an entry is already defined\n"));
			}
			int depth=c==0xae ? 1 : 0;
			while (!input->isEnd())
			{
				if (input->tell()>=m_state->m_eof)
				{
					WPS_DEBUG_MSG(("XYWriteParser::parseFormat: can not find end of entry\n"));
					return false;
				}
				c=libwps::readU8(input);
				if (c==0xae)
					++depth;
				else if (c==0xaf)
				{
					if (depth==0)
					{
						format.m_entry.setEnd(input->tell()-1);
						return true;
					}
					--depth;
				}
			}
		}
		if (c==0x1a)
		{
			WPS_DEBUG_MSG(("XYWriteParser::parseFormat: find end of zone\n"));
			return false;
		}
		if (c==',')
			format.m_args.push_back(std::string());
		else if (format.m_args.empty())
			format.m_string+=char(c);
		else
			format.m_args.back()+=char(c);
	}
	WPS_DEBUG_MSG(("XYWriteParser::parseFormat: find end of file\n"));
	return false;
}

bool XYWriteParser::createFormatChildren(XYWriteParserInternal::Format &format, size_t fPos)
{
	RVNGInputStreamPtr input = getInput();
	if (!input || !format.m_entry.valid())
		throw (libwps::ParseException());
	long endPos=format.m_entry.end();
	if (endPos>m_state->m_eof)
	{
		WPS_DEBUG_MSG(("XYWriteParser::createFormatChildren: end entry seems bad\n"));
		return false;
	}
	long prevPos=input->tell();
	input->seek(format.m_entry.begin(), librevenge::RVNG_SEEK_SET);
	// skip header
	for (size_t p=0; p<fPos; ++p)
	{
		if (input->isEnd() || input->tell()>=endPos)
			break;
		if (libwps::readU8(input)==',')
			break;
	}
	if (input->tell()>=endPos)
	{
		input->seek(prevPos, librevenge::RVNG_SEEK_SET);
		return true;
	}
	std::string actString;
	long begPos=input->tell();
	bool isComplex=false;
	while (!input->isEnd() && input->tell()<=endPos)
	{
		uint8_t c=input->tell()==endPos ? ',' : libwps::readU8(input);
		auto *child=format.m_children.empty() ? nullptr : &format.m_children.back();
		if (c==0xfa || c==0xae)
		{
			isComplex=true;
			// normally ae XXX fa ... af
			// but sometimes I found ae XXX ae ... af fa ... af
			int depth=c==0xae ? 1 : 0;
			while (!input->isEnd() && input->tell()<endPos)
			{
				c=libwps::readU8(input);
				if (c==0xae)
					++depth;
				else if (c==',' && depth==0)
					break;
				else if (c==0xaf)
				{
					if (depth==0)
						break;
					--depth;
				}
			}
			if (c!=',')
				continue;
		}
		if (c==0x1a)
		{
			WPS_DEBUG_MSG(("XYWriteParser::createFormatChildren: find end of zone\n"));
			input->seek(prevPos, librevenge::RVNG_SEEK_SET);
			return false;
		}
		if (c==',')
		{
			if (!actString.empty())
			{
				if (actString[0]=='.' || (actString[0]>='0' && actString[0]<='9'))
				{
					if (child)
					{
						child->m_isComplex=isComplex;
						child->m_args.push_back(actString);
					}
				}
				else
				{
					if (child)
						child->m_entry.setEnd(begPos-1);
					format.m_children.push_back(XYWriteParserInternal::Format());
					format.m_children.back().m_entry.setBegin(begPos);
					format.m_children.back().m_inDosFile = m_state->m_isDosFile;
					format.m_children.back().m_string=actString;
					format.m_children.back().m_isComplex=isComplex;
					isComplex=false;
				}
			}
			actString.clear();
			begPos=input->tell();
			if (!format.m_children.empty())
				format.m_children.back().m_entry.setEnd(begPos==endPos ? endPos : begPos-1);
			if (begPos==endPos)
				break;
			continue;
		}
		if (c!='=')
			actString+=char(c);
	}
	input->seek(prevPos, librevenge::RVNG_SEEK_SET);
	return true;
}

// basic function to check if the header is ok
bool XYWriteParser::checkHeader(WPSHeader *header, bool /*strict*/)
{
	RVNGInputStreamPtr input = getInput();
	if (!input || !checkFilePosition(10)) // too small for containing any usefull format
	{
		WPS_DEBUG_MSG(("XYWriteParser::checkHeader: file is too short\n"));
		return false;
	}

	/* check sequence 0xae ...[,*] 0xaf */
	input->seek(0x0, librevenge::RVNG_SEEK_SET);
	bool ok=false, inFormat=false;
	int depth=0;
	int numBadChar=0, numCurrentChar=0;
	while (!input->isEnd())
	{
		if (numBadChar>10)
			break;
		uint8_t c=libwps::readU8(input);
		if (c<=0x1f && c!=0x9 && c!=0xa && c!=0xd && c!=0x1b)
			++numBadChar;
		else if (!depth)
		{
			if (c==0xae)
			{
				inFormat=true;
				depth=1;
			}
			else if (c==0xaf) // end before begin
				break;
		}
		else
		{
			if (c==0xaf)
			{
				if (--depth==0)   // find the end of a sequence, ok
				{
					ok=true;
					break;
				}
				inFormat=false;
			}
			else if (c==0xae)
			{
				++depth;
				inFormat=true;
			}
			else if (c==',')
				numCurrentChar=0;
			else if (c==0xfa)
				inFormat=false;
			else if (inFormat)
			{
				if (++numCurrentChar>256)
					break;
			}
		}
	}
	if (!ok)
	{
		WPS_DEBUG_MSG(("XYWriteParser::checkHeader: can not find any sequence\n"));
		return false;
	}
	input->seek(-1, librevenge::RVNG_SEEK_END);
	int val=libwps::readU8(input);
	if (val==0x1a)
		m_state->m_isDosFile=true;
	else if (val)
	{
		WPS_DEBUG_MSG(("XYWriteParser::checkHeader: oops unexpected last character\n"));
		return false;
	}
	if (header)
		header->setMajorVersion(m_state->m_isDosFile ? 0 : 1);
	return true;
}

// read the document structure ...

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
