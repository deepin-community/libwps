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

#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WPSEntry.h"
#include "WPSContentListener.h"
#include "WPSFont.h"
#include "WPSList.h"
#include "WPSParagraph.h"

#include "WPS8.h"
#include "WPS8Struct.h"
#include "WPS8Text.h"

#include "WPS8TextStyle.h"

namespace WPS8TextStyleInternal
{
/** Internal: class to store  font properties */
struct Font final : public WPSFont
{
	//! constructor
	Font()
		: WPSFont()
		, m_special() { }
	Font(Font const &)=default;
	//! destructor
	~Font() final;
	static Font def()
	{
		Font res;
		res.m_name = "Times New Roman"; // checkme
		res.m_size = 10;
		return res;
	}

	//! operator<<
	friend std::ostream &operator<<(std::ostream &o, Font const &ft);

	//! return the auxilliary data
	WPS8TextStyle::FontData &special()
	{
		return m_special;
	}
	//! return the auxilliary data
	WPS8TextStyle::FontData const &special() const
	{
		return m_special;
	}

protected:
	/** the auxialliary data */
	WPS8TextStyle::FontData m_special;
};

Font::~Font()
{
}

//! operator<< for a font
std::ostream &operator<<(std::ostream &o, Font const &ft)
{
	o << reinterpret_cast<WPSFont const &>(ft);
	o << ft.m_special;
	return o;
}

/** Internal: the state of a WPS4Text */
struct State
{
	//! constructor
	State()
		: m_fontNames()
		, m_defaultFont(Font::def())
		, m_fontList()
		, m_defaultParagraph()
		, m_paragraphList()
		, m_fontTypes()
		, m_paragraphTypes()
	{
		initTypeMaps();
	}

	//! initializes the type map
	void initTypeMaps();

	//! the font names
	std::vector<librevenge::RVNGString> m_fontNames;

	//! the default font
	Font m_defaultFont;
	//! a list of all font properties
	std::vector<Font> m_fontList;
	//! the default paragraph
	WPSParagraph m_defaultParagraph;
	//! a list of all paragraph properties
	std::vector<WPSParagraph> m_paragraphList;

	//! the character type
	std::map<int,int> m_fontTypes;
	//! the paragraph type
	std::map<int,int> m_paragraphTypes;
};

void State::initTypeMaps()
{
	static int const fontTypes[] =
	{
		0, 0x12, 2, 0x2, 3, 0x2, 4, 0x2, 5, 0x2,
		0xc, 0x22, 0xf, 0x12,
		0x10, 0x2, 0x12, 0x22, 0x13, 0x2, 0x14, 0x2, 0x15, 0x2, 0x16, 0x2, 0x17, 0x2,
		0x18, 0x22, 0x1a, 0x22, 0x1b, 0x22, 0x1e, 0x12,
		0x22, 0x22, 0x23, 0x22, 0x24, 0x8A,
		0x2d, 0x2, 0x2e, 0x22,
	};
	for (int i = 0; i+1 < int(WPS_N_ELEMENTS(fontTypes)); i+=2)
		m_fontTypes[fontTypes[i]] = fontTypes[i+1];
	static int const paragraphTypes[] =
	{
		2, 0x22, 3, 0x1A, 4, 0x12, 6, 0x22,
		0xc, 0x22, 0xd, 0x22, 0xe, 0x22,
		0x12, 0x22, 0x13, 0x22, 0x14, 0x22, 0x15, 0x22, 0x17, 0x2,
		0x18, 0x2, 0x19, 0x1A, 0x1b, 0x2, 0x1c, 0x2, 0x1d, 0x2, 0x1e, 0x12, 0x1f, 0x22,
		0x20, 0x12, 0x21, 0x22, 0x22, 0x22, 0x23, 0x22, 0x24, 0x22, 0x25, 0x12,
		0x2a, 0x12,
		0x31, 0x12, 0x32, 0x82, 0x33, 0x12, 0x34, 0x22
	};
	for (int i = 0; i+1 < int(WPS_N_ELEMENTS(paragraphTypes)); i+=2)
		m_paragraphTypes[paragraphTypes[i]] = paragraphTypes[i+1];
}
}

////////////////////////////////////////////////////////////
// constructor destructor
////////////////////////////////////////////////////////////
WPS8TextStyle::WPS8TextStyle(WPS8Text &parser)
	: m_mainParser(parser)
	, m_input(parser.getInput())
	, m_listener()
	, m_state()
	, m_asciiFile(parser.ascii())
{
	m_state.reset(new WPS8TextStyleInternal::State);
}

WPS8TextStyle::~WPS8TextStyle()
{
}

////////////////////////////////////////////////////////////
// called to create the data:
////////////////////////////////////////////////////////////
bool WPS8TextStyle::readStructures()
{
	auto &nameTable = m_mainParser.getNameEntryMap();

	/* read fonts table */
	auto pos = nameTable.find("FONT");
	if (nameTable.end() == pos)
	{
		WPS_DEBUG_MSG(("WPS8TextStyle::parse: error: no FONT in header index table\n"));
		return false;
	}
	readFontNames(pos->second);

	// find the FDDP and FDPC positions
	for (int st = 0; st < 2; st++)
	{
		std::vector<WPSEntry> zones;
		if (!findFDPStructures(st, zones))
			findFDPStructuresByHand(st, zones);

		std::vector<WPS8Text::DataFOD> fdps;
		WPS8Text::FDPParser parser = st==0  ? WPS8Text::FDPParser(&WPS8Text::readParagraph)
		                             : WPS8Text::FDPParser(&WPS8Text::readFont);
		for (auto const &z : zones)
			m_mainParser.readFDP(z, fdps, parser);
		m_mainParser.m_FODList = m_mainParser.mergeSortedFODLists(m_mainParser.m_FODList, fdps);
	}
	// read SGP zone
	pos = nameTable.lower_bound("SGP ");
	while (nameTable.end() != pos)
	{
		WPSEntry const &entry = pos++->second;
		if (!entry.hasName("SGP ")) break;
		if (!entry.hasType("SGP ")) continue;

		readSGP(entry);
	}

	return true;
}

////////////////////////////////////////////////////////////
// the fontname:
////////////////////////////////////////////////////////////
bool WPS8TextStyle::readFontNames(WPSEntry const &entry)
{
	if (!entry.hasType(entry.name()))
	{
		WPS_DEBUG_MSG(("WPS8TextStyle::readFontNames: name=%s, type=%s\n",
		               entry.name().c_str(), entry.type().c_str()));
		return false;
	}

	if (entry.length() < 20)
	{
		WPS_DEBUG_MSG(("WPS8TextStyle::readFontNames: length=0x%ld\n", entry.length()));
		return false;
	}

	long debPos = entry.begin();
	m_input->seek(debPos, librevenge::RVNG_SEEK_SET);

	auto len = long(libwps::readU32(m_input)); // len + 0x14 = size
	auto n_fonts = size_t(libwps::readU32(m_input));

	if (long(4*n_fonts) > len)
	{
		WPS_DEBUG_MSG(("WPS8TextStyle::readFontNames: number=%d\n", int(n_fonts)));
		return false;
	}
	if (len > entry.length() - 20)
	{
		WPS_DEBUG_MSG(("WPS8TextStyle::readFontNames: invalid lenth=%ld\n", len));
		return false;
	}
	libwps::DebugStream f;

	entry.setParsed();
	f << "N=" << n_fonts;
	if (len+20 != entry.length()) f << ", ###L=" << std::hex << len+0x14;

	f << ", unkn=(" << std::hex;
	for (int i = 0; i < 3; i++) f << libwps::readU32(m_input) << ", ";
	f << "), dec=[";
	for (int i = 0; i < int(n_fonts); i++) f << ", " << libwps::read32(m_input);
	f << "]" << std::dec;

	ascii().addPos(debPos);
	ascii().addNote(f.str().c_str());

	long pageEnd = entry.end();

	/* read each font in the table */
	while (m_input->tell() > 0 && m_state->m_fontNames.size() < n_fonts)
	{
		debPos = m_input->tell();
		if (debPos+6 > long(pageEnd)) break;

		auto string_size = int(libwps::readU16(m_input));
		if (debPos+2*string_size+6 > long(pageEnd)) break;

		librevenge::RVNGString s;
		for (; string_size>0; string_size--) // checkme
			s.append(char(libwps::readU16(m_input)));

		f.str("");
		f << "FONT("<<m_state->m_fontNames.size()<<"): " << s.cstr();
		f << ", unkn=(";
		for (int i = 0; i < 4; i++) f << int(libwps::read8(m_input)) << ", ";
		f << ")";
		ascii().addPos(debPos);
		ascii().addNote(f.str().c_str());

		m_state->m_fontNames.push_back(s);
	}

	if (m_state->m_fontNames.size() != n_fonts)
	{
		WPS_DEBUG_MSG(("WPS8TextStyle::readFontNames: expected %i fonts but only found %i\n",
		               int(n_fonts), int(m_state->m_fontNames.size())));
	}
	return true;
}

////////////////////////////////////////////////////////////
// style general property
////////////////////////////////////////////////////////////
bool WPS8TextStyle::readSGP(WPSEntry const &entry)
{
	if (!entry.hasType(entry.name()))
	{
		WPS_DEBUG_MSG(("WPS8TextStyle::readSGP: warning: SGP name=%s, type=%s\n",
		               entry.name().c_str(), entry.type().c_str()));
		return false;
	}

	long page_offset = entry.begin();
	long length = entry.length();
	long endPage = entry.end();

	if (length < 2)
	{
		WPS_DEBUG_MSG(("WPS8TextStyle::readSGP: warning: SGP length=0x%lx\n", static_cast<unsigned long>(length)));
		return false;
	}

	entry.setParsed();
	m_input->seek(page_offset, librevenge::RVNG_SEEK_SET);

	libwps::DebugStream f;
	if (libwps::read16(m_input) != length)
	{
		WPS_DEBUG_MSG(("WPS8TextStyle::readSGP: invalid length=%ld\n", length));
		return false;
	}

	WPS8Struct::FileData mainData;
	std::string error;
	bool readOk=readBlockData(m_input, endPage,mainData, error);
	for (auto const &dt : mainData.m_recursData)
	{
		if (dt.isBad()) continue;
		if (dt.id() == 0)
			f << "tabSep[default]=" <<  float(dt.m_value)/914400.f << "(inches),";
		else
			f << "###" << dt << ",";
	}

	if (!readOk) f << "###or [" << mainData << "]";

	ascii().addPos(page_offset);
	ascii().addNote(f.str().c_str());

	return true;
}

////////////////////////////////////////////////////////////
// font
////////////////////////////////////////////////////////////
bool WPS8TextStyle::readFont(long endPos, int &id, std::string &mess)
{
	auto font=WPS8TextStyleInternal::Font::def();

	long actPos = m_input->tell();
	long size = endPos - actPos;

	/* other than blank, the shortest should be 2 bytes */
	if (size && size < 2)
	{
		WPS_DEBUG_MSG(("WPS8TextStyle::readFont: error: charProperty size < 2\n"));
		return false;
	}
	if (size && (size%2) == 1)
	{
		WPS_DEBUG_MSG(("WPS8TextStyle::readFont: error: charProperty size is odd\n"));
		return false;
	}

	WPS8Struct::FileData mainData;
	std::string error;

	bool readOk= size ? readBlockData(m_input, endPos, mainData, error) : true;

	libwps::DebugStream f;
	if (mainData.m_value) f << "unk=" << mainData.m_value << ",";

	uint32_t textAttributeBits = 0;
	for (auto const &data : mainData.m_recursData)
	{
		if (data.isBad()) continue;
		if (m_state->m_fontTypes.find(data.id())==m_state->m_fontTypes.end())
		{
			WPS_DEBUG_MSG(("WPS8TextStyle::readFont: unexpected id %d\n", data.id()));
			f << "##" << data << ",";
			continue;
		}
		if (m_state->m_fontTypes.find(data.id())->second != data.type())
		{
			WPS_DEBUG_MSG(("WPS8TextStyle::readFont: unexpected type for %d\n", data.id()));
			f << "###" << data << ",";
			continue;
		}

		switch (data.id())
		{
		case 0x0:
			font.special().m_type = int(data.m_value);
			break;
		case 0x02:
			if (data.isTrue())
				textAttributeBits |= WPS_BOLD_BIT;
			else
				f << "#bold=false,";
			break;
		case 0x03:
			if (data.isTrue())
				textAttributeBits |= WPS_ITALICS_BIT;
			else
				f << "#it=false,";
			break;
		case 0x04:
			if (data.isTrue())
				textAttributeBits |= WPS_OUTLINE_BIT;
			else
				f << "#outline=false,";
			break;
		case 0x05:
			if (data.isTrue())
				textAttributeBits |= WPS_SHADOW_BIT;
			else
				f << "#shadow=false,";
			break;
		case 0x0c:
			font.m_size = int(data.m_value/12700);
			break;
		case 0x0f:
			if (data.m_value == 1) textAttributeBits |= WPS_SUPERSCRIPT_BIT;
			else if (data.m_value == 2) textAttributeBits |= WPS_SUBSCRIPT_BIT;
			else f << "###sub/supScript=" << data.m_value << ",";
			break;
		case 0x10:
			if (data.isTrue())
				textAttributeBits |= WPS_STRIKEOUT_BIT;
			else
				f << "#strikeout=false,";
			break;
		case 0x12:
			font.m_languageId = int(data.m_value);
			break;
		case 0x13:
			if (data.isTrue())
				textAttributeBits |= WPS_SMALL_CAPS_BIT;
			else
				f << "#smallbit=false,";
			break;
		case 0x14:
			if (data.isTrue())
				textAttributeBits |= WPS_ALL_CAPS_BIT;
			else
				f << "#allcaps=false,";
			break;
		case 0x16:
			if (data.isTrue())
				textAttributeBits |= WPS_EMBOSS_BIT;
			else
				f << "#emboss=false,";
			break;
		case 0x17:
			if (data.isTrue())
				textAttributeBits |= WPS_ENGRAVE_BIT;
			else
				f << "#engrave=false,";
			break;
		case 0x18: // 0 or 0.25
			f << "##f24(inches)=" << double(data.m_value)/914400 << ",";
			break;
		case 0x1b:   // -3175
			if (data.m_value == -3175) f << "##f27,";
			else f << "##f27=" << data.m_value << ",";
			break;
		case 0x1e:
		{
			bool single = true;
			switch (data.m_value)
			{
			case 1:
				break; // normal
			case 2:
				f << "underl[word],";
				break;
			case 3:
				single = false;
				break; // double
			case 4:
				f << "underl[dot],";
				break;
			case 6:
				f << "underl[w=2],";
				break;
			case 7:
				f << "underl[dashed],";
				break;
			case 9:
				f << "underl[style=.-],";
				break;
			case 10:
				f << "underl[style=..-],";
				break;
			case 11:
				f << "underl[curve],";
				break;
			case 16:
				f << "underl[curve2],";
				break;
			case 17:
				f << "underl[dot,w=2],";
				break;
			case 18:
				f << "underl[dashed,w=2],";
				break;
			case 19:
				f << "underl[style=.-,w=2],";
				break;
			case 20:
				f << "underl[style=..-,w=2],";
				break;
			case 21:
				f << "underl[dashed,l*=2],";
				break; // length*=2
			case 22:
				f << "underl[dashed,w=2,l*=2],";
				break;
			case 23:
				f << "underl[curve]";
				single = false;
				break; // and double
			default:
				f << "###underlFlag=" << data.m_value << ",";
			}
			if (single)
				textAttributeBits |= WPS_UNDERLINE_BIT;
			else
				textAttributeBits |= WPS_DOUBLE_UNDERLINE_BIT;
			break;
		}
		case 0x22:
			font.special().m_fieldType = int(data.m_value);
			break;
		case 0x23:
			font.special().m_fieldFormat = int(data.m_value);
			break;
		case 0x24:
		{
			if ((!data.isRead() && !data.readArrayBlock() && data.m_recursData.size() == 0) ||
			        !data.isArray())
			{
				WPS_DEBUG_MSG(("WPS8TextStyle::readFont: can not read font array\n"));
				f << "###fontPb";
				break;
			}

			if (data.m_recursData.empty() || data.m_recursData[0].isBad() || data.m_recursData[0].type() != 0x18)
			{
				WPS_DEBUG_MSG(("WPS8TextStyle::readFont: can not read font id\n"));
				f << "###f24=[" << data << "]";
				break;
			}
			auto fontId = uint8_t(data.m_recursData[0].m_value);
			if (fontId < m_state->m_fontNames.size())

				font.m_name = m_state->m_fontNames[fontId];
			else
			{
				WPS_DEBUG_MSG(("WPS8TextStyle::readFont: can not read find font %d\n", int(fontId)));
			}
			std::vector<int> formats;
			for (auto const &subD : data.m_recursData)
			{
				if (subD.isBad()) continue;
				int formId = subD.id() >> 3;
				int sId = subD.id() & 0x7;
				if (sId == 0)
				{
					formats.resize(size_t(formId)+1,-1);
					formats[size_t(formId)] = int(subD.m_value);
				}
				else
					f << "###formats"<<formId<<"." << sId << "=" << subD << ",";
			}
			// CHECKME: list of font ids ....
			f << "formats=[" << std::hex;
			for (size_t i = 0; i < formats.size(); i++)
			{
				if (formats[i] != -1)
					f << "f" << i << "=" << formats[i] << ",";
			}
			f << "],";
			break;
		}
		case 0x2e:
			font.m_color=data.getRGBColor();
			break;
		default:
			f << "##" << data << ",";
			break;
		}
	}

	font.m_attributes = textAttributeBits;
	if (!readOk) f << ", ###or mainData=[" << mainData << "]";
	font.m_extra = f.str();
	font.m_extra += error;

	id = int(m_state->m_fontList.size());
	m_state->m_fontList.push_back(font);
	f.str("");
	f << font;
	mess = f.str();
	return true;
}

void WPS8TextStyle::sendFont(int fId, WPS8TextStyle::FontData &data)
{
	if (fId >= int(m_state->m_fontList.size()))
	{
		data = WPS8TextStyle::FontData();
		WPS_DEBUG_MSG(("WPS8TextStyle::sendFont: can not find font id %d\n", fId));
		return;
	}
	auto const &font = fId < 0 ? m_state->m_defaultFont : m_state->m_fontList[size_t(fId)];
	data = font.special();
	if (m_listener)
		m_listener->setFont(font);
}

////////////////////////////////////////////////////////////
// font
////////////////////////////////////////////////////////////
bool WPS8TextStyle::readParagraph(long endPos, int &id, std::string &mess)
{
	long actPos = m_input->tell();
	long size = endPos - actPos;

	/* other than blank, the shortest should be 2 bytes */
	if (size && (size%2) == 1)
	{
		WPS_DEBUG_MSG(("WPS8TextStyle::readParagraph: paraProperty size is odd\n"));
		return false;
	}

	libwps::DebugStream f;
	WPS8Struct::FileData mainData;
	std::string error;

	bool readOk= size ? readBlockData(m_input, endPos, mainData, error) : true;
	if (mainData.m_value) f << "unk=" << mainData.m_value << ",";

	WPSParagraph para;
	WPSColor paraColor[] = { WPSColor::black(), WPSColor::white() };
	for (auto const &data : mainData.m_recursData)
	{
		if (data.isBad()) continue;
		if (m_state->m_paragraphTypes.find(data.id())==m_state->m_paragraphTypes.end())
		{
			WPS_DEBUG_MSG(("WPS8TextStyle::readParagraph: unexpected id %d\n", data.id()));
			f << "###" << data << ",";
			continue;
		}
		if (m_state->m_paragraphTypes.find(data.id())->second != data.type())
		{
			WPS_DEBUG_MSG(("WPS8TextStyle::readParagraph: unexpected type for %d\n", data.id()));
			f << "###" << data << ",";
			continue;
		}
		bool ok = true;
		switch (data.id())
		{
		/* case 0x2: what?=data.m_value/914400.; */
		case 0x3:
			switch (data.m_value)
			{
			case 2: // not frequent: only found one time
				f << "#type[bullet]=2,";
				WPS_FALLTHROUGH;
			case 1: // the normal case
				para.m_listLevelIndex = 1;
				para.m_listLevel.m_type = libwps::BULLET;
				para.m_listLevel.m_bullet.clear();
				libwps::appendUnicode(0x2022, para.m_listLevel.m_bullet);
				break;
			default:
				f << "###bullet=" << data.m_value << ",";
				break;
			}
			break;
		case 0x4:
			switch (data.m_value)
			{
			case 0:
				para.m_justify = libwps::JustificationLeft;
				break;
			case 1:
				para.m_justify = libwps::JustificationRight;
				break;
			case 2:
				para.m_justify = libwps::JustificationCenter;
				break;
			case 3:
				para.m_justify = libwps::JustificationFull;
				break;
			default:
				f << "#just=" << std::hex << data.m_value << std::dec << ",";
				para.m_justify = libwps::JustificationLeft;
			}
			break;
		case 0x6:
			para.m_listLevel.m_labelIndent = double(data.m_value)/914400;
			break;
		case 0xc: // first line indentation (6*152400 unit = 1 inches)
			para.m_margins[0] = double(data.m_value)/914400;
			break;
		case 0xd: // left indentation
			para.m_margins[1] = double(data.m_value)/914400;
			break;
		case 0xe: // right indentation
			para.m_margins[2] = double(data.m_value)/914400;
			break;

		case 0x12: // before line spacing 152400 -> 1 line
			para.m_spacings[1] = double(data.m_value)/152400;
			break;
		case 0x13: // after line spacing 152400 -> 1 line
			para.m_spacings[2] = double(data.m_value)/152400;
			break;
		// case 0x15(type22): one time with value 0x29
		case 0x14:
		{
			// link to bullet or numbering

			// first check if this can be a numbering level
			auto suffixId = int(data.m_value >> 16); // 0 -> . 2->) 3-> ??
			int type = data.m_value & 0xFFFF;

			if (data.m_value && suffixId >= 0 && suffixId < 5 && type >= 0 && type <= 6)
			{
				para.m_listLevelIndex = 1;
				/* this seems to implies that we restart a list */
				if (para.m_listLevel.m_type == libwps::BULLET)
					para.m_listLevel.m_startValue=1;

				switch (type)
				{
				case 0:
					para.m_listLevel.m_type = libwps::NONE;
					break;
				case 3:
					para.m_listLevel.m_type = libwps::LOWERCASE;
					break;
				case 4:
					para.m_listLevel.m_type = libwps::UPPERCASE;
					break;
				case 5:
					para.m_listLevel.m_type = libwps::LOWERCASE_ROMAN;
					break;
				case 6:
					para.m_listLevel.m_type = libwps::UPPERCASE_ROMAN;
					break;
				default:
					f << "#bullet/type=" << type << ",";
					WPS_FALLTHROUGH;
				case 2:
					para.m_listLevel.m_type = libwps::ARABIC;
					break;
				}
				switch (suffixId)
				{
				case 0:
					para.m_listLevel.m_suffix = ")";
					break;
				case 2:
					para.m_listLevel.m_suffix = ".";
					break;
				case 3:
					para.m_listLevel.m_suffix = "-";
					break; // checkme
				default:
					f << "#bullet/suffix=" << suffixId << ",";
				}
			}
			/** Note: I also find a val 0x2D which seems to have some sense
				in correspondance with field 15 (val 0x29) and field 1c */
			else if (para.m_listLevel.m_type == libwps::BULLET)
			{
				para.m_listLevel.m_bullet.clear();
				uint32_t code = 0;
				if (data.m_value <= 255)
					code = uint32_t(libwps_tools_win::Font::unicode(static_cast<unsigned char>(data.m_value), libwps_tools_win::Font::WIN3_WINGDINGS));
				if (!code)
				{
					WPS_DEBUG_MSG(("WPS8TextStyle::readParagraph: can not find bullet code\n"));
					f << "#bullet[code]=" << std::hex << data.m_value << "," << std::dec;
					code = 0x2022;
				}
				else if (code >= 0x10000)
					code = 0x2022; // this unicode char may not print ok
				libwps::appendUnicode(code, para.m_listLevel.m_bullet);
			}
			else
				f << "##bullet/numb=" << std::hex << data.m_value << "," << std::dec;
			break;
		}
		case 0x15:
			if (para.m_listLevel.isNumeric() && data.m_value >= 0)
				para.m_listLevel.m_startValue = int(data.m_value);
			else
				// can also be present in the lines preceding/following a line's list : ok
				// do we need to use this value as to compute the next starting value ?
				f << "#bullet/startValue?=" << std::hex << data.m_value << std::dec << ",";
			break;
		case 0x17:
			f << "modBord,";
			break;
		case 0x18:
			f << "modTabs,";
			break;
		// case 0x19(type1a): number between 1 and 6 : stylesheet index ?
		case 0x1b:
			if (data.m_value == 1) f << "bColType=rgb?,";
			else f << "#bColType=" << std::hex << data.m_value << std::dec << ",";
			break;
		case 0x1d:
			f << "##f29Set,";
			break; // present if f42(2a) exist ?
		case 0x1e: // not filled by word 2000 ?
			// 1 -> Left, 2 -> right, 4-> top, 8->bottom
			if (data.m_value & 1) para.m_border |= WPSBorder::TopBit;
			if (data.m_value & 2) para.m_border |= WPSBorder::BottomBit;
			if (data.m_value & 4) para.m_border |= WPSBorder::LeftBit;
			if (data.m_value & 8) para.m_border |= WPSBorder::RightBit;
			if (data.m_value & 0xF0)
				f << "#border=" << std::hex << (data.m_value & 0xF0) << std::dec << ",";
			break;
		case 0x1f:
			para.m_borderStyle.m_color=data.getRGBColor();
			break;
		case 0x20:
		{
			std::string styleMessage("");
			data.getBorderStyles(para.m_borderStyle.m_style, para.m_borderStyle.m_type, styleMessage);
			f << styleMessage;
			break;
		}
		case 0x21:
			para.m_borderStyle.m_width = int(data.m_value)/12700;
			break;
		case 0x22:
			f << "#bordSzY=" << float(data.m_value)/12700.f << ",";
			break;

		case 0x23: // color used to define background: col1*pat+col2*(1-pat)
		case 0x24:
			// color1/2 : default color1=black and color2=white
			paraColor[data.id()-0x23]=data.getRGBColor();
			break;
		case 0x25:
		{
			float percent=0.5;
			if (data.m_value == 0) // checkme: no motif
				break;

			if (data.m_value >= 3 && data.m_value <= 9)
				percent = float(data.m_value)*0.1f; // gray motif
			else
				f << "backMotif=" << data.m_value << ",";
			para.m_backgroundColor=WPSColor::barycenter(percent,paraColor[0],1.f-percent,paraColor[1]);
			break;
		}
		case 0x2a: // exists with f29(1d) in style sheet, find 0|1|3
			f << "##f42=" << data.m_value << ",";
			break;
		// case 0x31(typ12) : always 1 ?
		// case 0x33(typ12) : always 2 ?
		case 0x32:
		{
			if (!data.isRead() && !data.readArrayBlock() && data.m_recursData.size() == 0)
			{
				WPS_DEBUG_MSG(("WPS8TextStyle::readParagraph can not find tabs array\n"));
				ok = false;
				break;
			}
			size_t nChild = data.m_recursData.size();
			if (nChild < 1 ||
			        data.m_recursData[0].isBad() || data.m_recursData[0].id() != 0x27)
			{
				WPS_DEBUG_MSG(("WPS8TextStyle::readParagraph can not find first child\n"));
				ok = false;
				break;
			}
			if (nChild == 1) break;

			auto numTabs = int(data.m_recursData[0].m_value);
			if (numTabs == 0 || nChild < 2 ||
			        data.m_recursData[1].isBad() || data.m_recursData[1].id() != 0x28)
			{
				WPS_DEBUG_MSG(("WPS8TextStyle::readParagraph can not find second child\n"));
				ok = false;
				break;
			}

			WPS8Struct::FileData const &mData = data.m_recursData[1];
			size_t lastParsed = 0;
			if (mData.id() == 0x28 && mData.isArray() &&
			        (mData.isRead() || mData.readArrayBlock() || mData.m_recursData.size() != 0))
			{
				lastParsed = 1;
				int actTab = 0;
				para.m_tabs.resize(size_t(numTabs));

				for (auto const &mDataC : mData.m_recursData)
				{
					if (mDataC.isBad()) continue;
					int value = mDataC.id();
					int wTab = value/8;
					int what = value%8;

					// the first tab can be skipped
					// so this may happens only one time
					if (wTab > actTab && actTab < numTabs)
					{
						para.m_tabs[size_t(actTab)].m_alignment = WPSTabStop::LEFT;
						para.m_tabs[size_t(actTab)].m_position =  0.;

						actTab++;
					}

					if (mDataC.isNumber() && wTab==actTab && what == 0
					        && actTab < numTabs)
					{
						para.m_tabs[size_t(actTab)].m_alignment = WPSTabStop::LEFT;
						para.m_tabs[size_t(actTab)].m_position =  double(mDataC.m_value)/914400;

						actTab++;
						continue;
					}
					if (mDataC.isNumber() && wTab == actTab-1 && what == 1)
					{
						auto actVal = int(mDataC.m_value);
						switch ((actVal & 0x3))
						{
						case 0:
							para.m_tabs[size_t(actTab-1)].m_alignment = WPSTabStop::LEFT;
							break;
						case 1:
							para.m_tabs[size_t(actTab-1)].m_alignment = WPSTabStop::RIGHT;
							break;
						case 2:
							para.m_tabs[size_t(actTab-1)].m_alignment = WPSTabStop::CENTER;
							break;
						case 3:
							para.m_tabs[size_t(actTab-1)].m_alignment = WPSTabStop::DECIMAL;
							break;
						default:
							break;
						}
						if (actVal&0xC)
							f << "###tabFl" << actTab<<":low=" << (actVal&0xC) << ",";
						actVal = (actVal>>8);
						/* not frequent:
						   but fl1:high=db[C], fl2:high=b7[R] appear relatively often*/
						if (actVal)
						{
							f << ", fl" << actTab<<":high="  << std::hex
							  << actVal << std::dec;
							switch (para.m_tabs[size_t(actTab-1)].m_alignment)
							{
							case WPSTabStop::LEFT:
								break;
							case WPSTabStop::RIGHT:
								f << "[R]";
								break;
							case WPSTabStop::CENTER:
								f << "[C]";
								break;
							case WPSTabStop::DECIMAL:
								f << "[D]";
								break;
							case WPSTabStop::BAR:
							default:
								f << "[?]";
								break;
							}
						}
						continue;
					}
					if (mDataC.isNumber() && wTab == actTab-1 && what == 2)
					{
						para.m_tabs[size_t(actTab-1)].m_leaderCharacter = uint16_t(mDataC.m_value);
						continue;
					}
					f << "###tabData:fl" << actTab << "=" << mDataC << ",";
				}
				if (actTab != numTabs)
				{
					f << "NTabs[###founds]="<<actTab << ",";
					para.m_tabs.resize(size_t(actTab));
				}
			}
			for (size_t ch =lastParsed+1; ch < nChild; ch++)
			{
				if (data.m_recursData[ch].isBad()) continue;
				f << "extra[tabs]=[" << data.m_recursData[ch] << "],";
			}
		}
		WPS_FALLTHROUGH; // checkme

		case 0x34: // interline line spacing 8*152400 -> normal, sinon *2
		{
			float lines = float(data.m_value)/1219200.f;
			if (lines < 0.5f)
			{
				// find in one file some bogus line spacing between 0.2 and 0.3
				f << "###lineSpacing = " << lines << ",";
				lines = 1;
			}
			para.setInterline(double(lines), librevenge::RVNG_PERCENT);
			break;
		}

		default:
			ok = false;
		}

		if (ok) continue;
		f << "###" << data << ",";
	}

	if (!readOk)
		f << "###or [" << mainData << "]";
	para.m_extra = f.str();
	para.m_extra += error;
	if (para.m_listLevelIndex)
	{
		para.m_margins[0]+=para.m_margins[1];
		para.m_listLevel.m_labelIndent = para.m_margins[0];
	}
	id = int(m_state->m_paragraphList.size());
	m_state->m_paragraphList.push_back(para);

	f.str("");
	f << para;
	mess = f.str();
	return true;
}


void WPS8TextStyle::sendParagraph(int pId)
{
	if (pId >= int(m_state->m_paragraphList.size()))
	{
		WPS_DEBUG_MSG(("WPS8TextStyle::sendParagraph: can not find paragraph id %d\n", pId));
		return;
	}
	if (!m_listener) return;
	auto const &para=pId < 0 ? m_state->m_defaultParagraph : m_state->m_paragraphList[size_t(pId)];
	m_listener->setParagraph(para);
}

////////////////////////////////////////////////////////////
// StyleSheet: STSH Zone (Checkme)
////////////////////////////////////////////////////////////
bool WPS8TextStyle::readSTSH(WPSEntry const &entry)
{
	if (!entry.hasType(entry.name()))
	{
		WPS_DEBUG_MSG(("WPS8TextStyle::readSTSH: warning: STSH name=%s, type=%s\n",
		               entry.name().c_str(), entry.type().c_str()));
		return false;
	}
	long page_offset = entry.begin();
	long length = entry.length();
	long endPage = entry.end();

	if (length < 20)
	{
		WPS_DEBUG_MSG(("WPS8TextStyle::readSTSH: warning: STSH length=0x%lx\n", static_cast<unsigned long>(length)));
		return false;
	}

	entry.setParsed();
	m_input->seek(page_offset, librevenge::RVNG_SEEK_SET);

	libwps::DebugStream f;

	if (libwps::read32(m_input) != length-20) return false;
	int N = libwps::read32(m_input);

	if (N < 0) return false;
	f << "N=" << N; // 1 or 2

	f << std::hex << ", unk1=" << libwps::read32(m_input) << std::dec;
	int type = libwps::read32(m_input);
	f << ", type=" << type; // 4 : string ? 1 : unknown
	f << std::hex << ", unk2=" << libwps::read32(m_input); // "HST"
	f << ", pos=(";

	long debZone = m_input->tell();
	std::vector<long> ptrs;
	if (debZone + 4*N > endPage) return false;

	bool ok = true;
	for (int i = 0; i < N; i++)
	{
		long val = libwps::read32(m_input) + debZone;
		if (val >= endPage)
		{
			ok = false;
			break;
		}
		f << val << ",";
		ptrs.push_back(val);
	}
	f << ")";

	ascii().addPos(page_offset);
	ascii().addNote(f.str().c_str());

	if (!ok) return false;

	for (size_t i = 0; i < size_t(N); i++)
	{
		long pos = ptrs[i];
		long endZPos = (i+1 == size_t(N)) ? endPage : ptrs[i+1];
		length = endZPos-pos;
		if (length < 2)
		{
			ok = false;
			continue;
		}

		f.str("");
		f << std::dec << "STSH(";
		if (type == 4) f << i;
		else
		{
			if (i%2) f << "P" << i/2;
			else f << "C" << i/2;
		}
		f << "):";

		m_input->seek(pos, librevenge::RVNG_SEEK_SET);
		auto size = int(libwps::readU16(m_input));
		bool correct = true;
		if (2*size + 2 + type != length) correct = false;
		else
		{
			switch (type)
			{
			case 4:
			{
				librevenge::RVNGString str;
				if (size) WPS8Text::readString(m_input, 2*size, str);
				f << "'" << str.cstr() << "'";
				m_input->seek(pos+2+2*size, librevenge::RVNG_SEEK_SET);
				f << ", unkn=" << libwps::read32(m_input);
				break;
			}
			case 0:
			{
				WPS8Struct::FileData mainData;
				auto dataSz = int(libwps::readU16(m_input));
				if (dataSz+2 != 2*size)
				{
					correct = false;
					break;
				}

				int id;
				std::string mess;
				if (i%2 == 0)
				{
					readFont(pos+2*size, id, mess);
					f << "Font" << id << "=" << mess;
				}
				else
				{
					readParagraph(pos+2*size, id, mess);
					f << "Paragraph" << id << "=" << mess;
				}
				break;
			}
			default:
				correct = false;
			}
		}
		if (!correct)
		{
			f << "###";
			ok = false;
		}

		ascii().addPos(pos);
		ascii().addNote(f.str().c_str());

	}

	return ok;
}

////////////////////////////////////////////////////////////
// code to find the fdpc and fdpp entries; normal then by hand
////////////////////////////////////////////////////////////
bool WPS8TextStyle::findFDPStructures(int which, std::vector<WPSEntry> &zones)
{
	zones.resize(0);

	char const *indexName = which ? "BTEC" : "BTEP";
	char const *sIndexName = which ? "FDPC" : "FDPP";

	auto pos = m_mainParser.getNameEntryMap().lower_bound(indexName);

	std::vector<WPSEntry const *> listIndexed;
	while (pos != m_mainParser.getNameEntryMap().end())
	{
		WPSEntry const &entry = pos++->second;
		if (!entry.hasName(indexName)) break;
		if (!entry.hasType("PLC ")) continue;
		listIndexed.push_back(&entry);
	}

	size_t nFind = listIndexed.size();
	if (nFind==0) return false;

	// can nFind be > 1 ?
	for (size_t i = 0; i+1 < nFind; i++)
	{
		bool ok = true;
		for (size_t j = 0; j+i+1 < nFind; j++)
		{
			if (listIndexed[j]->id() <= listIndexed[j+1]->id())
				continue;
			WPSEntry const *tmp = listIndexed[j];
			listIndexed[j] = listIndexed[j+1];
			listIndexed[j+1] = tmp;
			ok = false;
		}
		if (ok) break;
	}

	for (size_t i = 0; i+1 < nFind; i++)
		if (listIndexed[i]->id() == listIndexed[i+1]->id()) return false;

	// create a map offset -> entry
	std::map<long, WPSEntry const *> offsetMap;
	std::map<long, WPSEntry const *>::iterator offsIt;
	pos = m_mainParser.getNameEntryMap().lower_bound(sIndexName);
	while (pos != m_mainParser.getNameEntryMap().end())
	{
		WPSEntry const &entry = pos++->second;
		if (!entry.hasName(sIndexName)) break;
		offsetMap.insert(std::map<long, WPSEntry const *>::value_type
		                 (entry.begin(), &entry));
	}

	for (size_t i = 0; i < nFind; i++)
	{
		WPSEntry const &entry = *(listIndexed[i]);
		std::vector<long> textPtrs;
		std::vector<long> listValues;

		if (!m_mainParser.readPLC(entry, textPtrs, listValues)) return false;

		if (textPtrs.size() != listValues.size()+1) return false;

		for (auto position : listValues)
		{
			if (position <= 0) return false;

			offsIt = offsetMap.find(position);
			if (offsIt == offsetMap.end() || !offsIt->second->hasName(sIndexName))
				return false;

			zones.push_back(*offsIt->second);
		}
	}

	return true;
}

bool WPS8TextStyle::findFDPStructuresByHand(int which, std::vector<WPSEntry> &zones)
{
	char const *indexName = which ? "FDPC" : "FDPP";
	WPS_DEBUG_MSG(("WPS8TextStyle::findFDPStructuresByHand: error: need to create %s list by hand \n", indexName));
	zones.resize(0);

	auto pos = m_mainParser.getNameEntryMap().lower_bound(indexName);
	while (pos != m_mainParser.getNameEntryMap().end())
	{
		WPSEntry const &entry = pos++->second;
		if (!entry.hasName(indexName)) break;
		if (!entry.hasType(indexName)) continue;

		zones.push_back(entry);
	}
	return zones.size() != 0;
}

////////////////////////////////////////////////////////////
// FontData functions
////////////////////////////////////////////////////////////
std::ostream &operator<<(std::ostream &o, WPS8TextStyle::FontData const &fData)
{
	switch (fData.m_type)
	{
	case WPS8TextStyle::FontData::T_None:
		break;
	case WPS8TextStyle::FontData::T_Object:
		o << "object,";
		break;
	case WPS8TextStyle::FontData::T_Footnote:
		o << "footnote,";
		break;
	case WPS8TextStyle::FontData::T_Endnote:
		o << "endnote,";
		break;
	case WPS8TextStyle::FontData::T_Field:
		o << "field,";
		break;
	case WPS8TextStyle::FontData::T_Comment: // or maybe link
		o << "comment,";
		break;
	default:
		o << "#type=" << fData.m_type << ",";
		break;
	}
	switch (fData.m_fieldType)
	{
	case WPS8TextStyle::FontData::F_None:
		break;
	case WPS8TextStyle::FontData::F_PageNumber:
		o << "pNumber,";
		break;
	case WPS8TextStyle::FontData::F_Date:
		o << "date,";
		break;
	case WPS8TextStyle::FontData::F_Time:
		o << "time,";
		break;
	default:
		o << "#fType=" << fData.m_fieldType << ",";
		break;
	}
	if (fData.m_fieldFormat) o << "fFormat=" << std::hex << fData.m_fieldFormat << std::dec << ",";
	return o;
}

std::string WPS8TextStyle::FontData::format() const
{
	switch (m_fieldFormat)
	{
	case 0x75: // 13/8/12
		return "%m/%d/%y";
	case 0x77: // 12 aout 12
		return "%d %B %y";
	case 0x78: // 2 aout 2012
	case 0x79: // 02 aout 2012
		return "%d %B %Y";
	case 0x7b: // aout 12
		return "%B %y";
	case 0x7c: // aout 2012,
		return "%B %Y";
	case 0x84: // lundi 13 aout 2012)
		return "%A %d %B %Y";
	case 0x87: // 09:06 p.m
		return "%I:%M %p";
	case 0x89: // 21:06
		return "%H:%M";
	case 0x8e: // 13/08/2012
	case 0x8f: // 13/8/2012
		return "%m/%d/%Y";
	case 0x90: // 13/8
		return "%d/%m";
	case 0x91: // 8/12
		return "%m/%y";
	case 0x102: // aout 2012
		return "%B %Y";
	case 0x125: //2012/8/13
		return "%Y/%m/%d";
	default:
		break;
	}
	if (m_fieldType == F_Date) return "%m/%d/%y";
	if (m_fieldType == F_Time) return "%I:%M:%S %p";
	WPS_DEBUG_MSG(("WPS8TextStyle::FontData::format: called with no date/time field\n"));
	return "";
}

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */

