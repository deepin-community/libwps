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
#include <limits>
#include <map>
#include <stack>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"
#include "libwps_tools_win.h"

#include "WKSContentListener.h"
#include "WKSSubDocument.h"

#include "WPSEntry.h"
#include "WPSFont.h"
#include "WPSGraphicShape.h"
#include "WPSGraphicStyle.h"
#include "WPSOLEObject.h"
#include "WPSParagraph.h"
#include "WPSPosition.h"
#include "WPSStream.h"

#include "Quattro9.h"

#include "Quattro9Graph.h"

namespace Quattro9GraphInternal
{
//! Internal: a shape of a Quattro9Graph
struct Shape
{
	//! constructor
	Shape()
		: m_box()
		, m_listShapes()
		, m_child()
	{
	}
	//! returns true if the shape is empty
	bool empty() const
	{
		return m_listShapes.empty() && !m_child;
	}
	//! returns true if the bounding box
	WPSBox2f getBdBox() const
	{
		if (m_box.size()[0]>0 || m_box.size()[1]>0) return m_box;
		WPSBox2f box;
		bool first=true;
		for (auto const &sh : m_listShapes)
		{
			if (first)
			{
				box=sh.first.getBdBox();
				first=false;
			}
			else
				box=box.getUnion(sh.first.getBdBox());
		}
		if (m_child)
		{
			if (first)
			{
				box=m_child->getBdBox();
				first=false;
			}
			else
				box=box.getUnion(m_child->getBdBox());
		}
		if (first)
		{
			WPS_DEBUG_MSG(("QuattroGraphInternal::Shape:getBdBox() can not find any box\n"));
		}
		return box;
	}
	//! the box(if this is a group)
	WPSBox2f m_box;
	//! the list of shape and style
	std::vector<std::pair<WPSGraphicShape, WPSGraphicStyle> > m_listShapes;
	//! the child (if this is a group)
	std::shared_ptr<Shape> m_child;
};

//! Internal: a shape of a Quattro9Graph
struct Textbox
{
	//! constructor
	Textbox()
		: m_font()
		, m_paragraph()
		, m_style(WPSGraphicStyle::emptyStyle())
		, m_text()
		, m_stream()
	{
	}
	//! the font
	WPSFont m_font;
	//! the paragraph
	WPSParagraph m_paragraph;
	//! the textbox style
	WPSGraphicStyle m_style;
	//! the text
	Quattro9ParserInternal::TextEntry m_text;
	//! the text stream
	std::shared_ptr<WPSStream> m_stream;
};

//! Internal: a graph of a QuattroGraph
struct Graph
{
	//! the posible type
	enum Type { Button, Chart, Frame, OLE/* or bitmap */, Image, Shape, Textbox, Unknown };
	//! constructor
	explicit Graph(std::shared_ptr<WPSStream> const &stream, Type type=Unknown)
		: m_type(type)
		, m_size()
		, m_cellBox()
		, m_cellBoxDecal()
		, m_OLEName()
		, m_shape()
		, m_textbox()
		, m_stream(stream)
	{
	}
	//! operator<<
	friend std::ostream &operator<<(std::ostream &o, Graph const &gr);
	//! the type
	Type m_type;
	//! the size
	Vec2f m_size;
	//! the cell's position
	WPSBox2i m_cellBox;
	//! the decal position(LT, RB)
	WPSBox2f m_cellBoxDecal;
	//! the OLE name
	librevenge::RVNGString m_OLEName;
	//! the graphic shape
	std::shared_ptr<Quattro9GraphInternal::Shape> m_shape;
	//! the textbox shape
	std::shared_ptr<Quattro9GraphInternal::Textbox> m_textbox;

	//! the main stream
	std::shared_ptr<WPSStream> m_stream;
};

std::ostream &operator<<(std::ostream &o, Graph const &gr)
{
	if (gr.m_size!=Vec2f()) o << "size=" << gr.m_size << ",";
	if (gr.m_cellBox!=WPSBox2i()) o << "cellBox=" << gr.m_cellBox << ",";
	if (gr.m_cellBoxDecal!=WPSBox2f()) o << "cellBox[decal]=" << gr.m_cellBoxDecal << ",";
	return o;
}

//! the state of Quattro9Graph
struct State
{
	//! constructor
	State()
		: m_version(-1)
		, m_actualSheet(-1)
		, m_zoneDepth(0)
		, m_actualGraph()
		, m_actualGraphDepth(-1)
		, m_colorsList()
		, m_patterns32List()
		, m_sheetIdToGraphMap()
		, m_linkNameToObjectMap()
	{
	}
	//! store a graph
	void storeGraph(std::shared_ptr<Graph> graph)
	{
		if (!graph)
		{
			WPS_DEBUG_MSG(("QuattroGraphInternal::storeGraph: no graph\n"));
			return;
		}
		m_actualGraph=graph;
		m_actualGraphDepth=m_zoneDepth;
		if (m_actualSheet<0)
		{
			WPS_DEBUG_MSG(("QuattroGraphInternal::storeGraph: can not find the current sheet\n"));
			return;
		}
		m_sheetIdToGraphMap.insert(std::multimap<int, std::shared_ptr<Graph> >::value_type(m_actualSheet, graph));
	}
	//! returns the color corresponding to an id
	bool getColor(int id, WPSColor &color) const;
	//! returns the pattern corresponding to a pattern id between 0 and 24
	static bool getPattern24(int id, WPSGraphicStyle::Pattern &pattern);
	//! returns the pattern corresponding to a pattern id between 0 and 32
	bool getPattern32(int id, WPSGraphicStyle::Pattern &pattern);
	//! the file version
	int m_version;
	//! the actual sheet id
	int m_actualSheet;
	//! current zone begin/end depth
	int m_zoneDepth;
	//! the actual graph
	std::shared_ptr<Graph> m_actualGraph;
	//! the depth which correspond to the creation of the graph
	int m_actualGraphDepth;
	//! the color list
	std::vector<WPSColor> m_colorsList;
	//! the pattern 32 list
	std::vector<WPSGraphicStyle::Pattern> m_patterns32List;
	//! a multimap sheetId to graph
	std::multimap<int, std::shared_ptr<Graph> > m_sheetIdToGraphMap;
	//! a map link name to object
	std::map<librevenge::RVNGString,WPSEmbeddedObject> m_linkNameToObjectMap;
};

bool State::getColor(int id, WPSColor &color) const
{
	if (m_colorsList.empty())
	{
		auto *THIS=const_cast<State *>(this);
		static const uint32_t quattroColorMap[]=
		{
			0xFFFFFF/*none*/, 0xB3B3B3, 0x6D6D6D, 0x000000,
			0xE12728, 0x71F504, 0x0029FA, 0x69FAFA,
			0xDC3BFF, 0xFFFA28, 0x5E196D, 0x306801,
			0x6F6B11, 0x00116B, 0x601111, 0x2D6B6B,
			// 10
			0xFFFFFF, 0x000830, 0xF9D9B6,0xB0B6FD,
			0xA277FD, 0x5A44FB, 0x5B22B0, 0x143030,
			0x1C73FA, 0xB9D9FD, 0xC4FDFD, 0x74FAFA,
			0x162F00, 0xC6FAB3, 0xFFFB74, 0xFBD628,
			// 20
			0x2B0707, 0xEB95B6, 0xF3B474, 0xE97228,
			0xE55728, 0x2A0B31, 0xEFB9FF, 0xEB95B6,
			0xE677B6, 0xE02C74, 0x808080, 0xAD8F70,
			0xA35239, 0xFFFFFF/*CheckMe*/, 0xFFFFFF/*CheckMe*/, 0xFFFFFF/*CheckMe*/,
			// 30
			0x633434, 0x150303, 0x2B0707, 0x601111,
			0x601111/*CheckMe*/, 0x9E1C1C, 0x9E1C1C, 0xE12728,
			0xE12728/*CheckMe*/, 0xE22F28, 0xE34040, 0xE55740,
			0xE55740, 0xED9374, 0xF1B6B6, 0xF9D9B6,
			// 40
			0x000000/*CheckMe*/, 0x2D1907, 0x2D1907, 0x2D1907,
			0x643311, 0x643311, 0xA3511C, 0xA3511C,
			0xE97228/*CheckMe*/, 0xED9141, 0xED9141, 0xED9141,
			0xED9141, 0xF3B474, 0xFFFDB6/*Checkme*/, 0xF9D9B6,
			// 50
			0x6E6B34, 0x2D1907, 0x323007, 0x394B08,
			0x6F6B11/*CheckMe*/, 0xAE8D1C, 0xB6AF1C, 0xFBD628,
			0xFFFA28/*CheckMe*/, 0xFFFA28, 0xFFFA41, 0xFFFA41,
			0xFFFB74, 0xFFFB74, 0xFFFDB6, 0xFFFDB6,
			// 60
			0x6E6B34, 0x0B1700, 0x323007, 0x394B08,
			0x436908, 0x508909, 0x81AD11, 0x8DD111,
			0x9AF612, 0xC9F83A, 0x9AF634, 0xC9F83A,
			0xC8F970, 0xC8F970, 0xFFFDB6, 0xFFFDB6,
			// 70
			0x436931, 0x0B1700, 0x162F00, 0x224A00,
			0x306801/*CheckMe*/, 0x3F8901, 0x4FAB02, 0x60D003,
			0x71F504, 0x71F504, 0x7CF531, 0x7CF531,
			0x99F76D, 0xC6FAB3, 0xC6FAB3, 0xC6FAB3,
			// 80
			0x416B6B, 0x0B1700, 0x143030, 0x214B30,
			0x306930, 0x3F8930, 0x4FAC30, 0x5ED16B,
			0x70F66B, 0x70F66B, 0x70F66B, 0x97F9B1,
			0x97F9B1, 0xC6FAB3, 0xC6FAB3, 0xC6FAB3,
			// 90
			0x416B6B, 0x071930, 0x143030, 0x214B30,
			0x2D6B6B, 0x388DAF, 0x49AFAF, 0x56D6FA,
			0x69FAFA/*CheckMe*/, 0x74FAFA, 0x74FAFA, 0x74FAFA,
			0x94FBFB, 0xC4FDFD, 0xC4FDFD, 0xC4FDFD,
			// a0
			0x364E6B, 0x071930, 0x071930, 0x0E336B,
			0x0E336B, 0x0E336B, 0x1451AF, 0x1C73FA,
			0x1C73FA/*CheckMe*/, 0x1C73FA, 0x4392FA, 0x7AB4FB,
			0x7AB4FB, 0xB9D9FD, 0xB9D9FD, 0xB9D9FD,
			// b0
			0x2E346B, 0x000830, 0x000830, 0x00116B,
			0x00116B/*CheckMe*/, 0x00116B, 0x001CAF, 0x001CAF,
			0x0029FA/*CheckMe*/, 0x1E41FA, 0x1E41FA, 0x5F5AFB,
			0x6675FB, 0xA895FD, 0xB0B6FD, 0xF6DCFF,
			// c0
			0x62376D, 0x62376D/*CheckMe*/, 0x000830, 0x27136B,
			0x27136B, 0x27136B, 0x562DFB/*checkme*/, 0x562DFB,
			0x562DFB, 0x5734FB, 0x5A44FB/*checkme*/, 0x9D5DFD,
			0xA277FD, 0xA895FD, 0xEFB9FF, 0xF6DCFF,
			// d0
			0x62376D, 0x2A0B31, 0x2A0B31, 0x2A0B31,
			0x5E196D, 0x9A29B3, 0x9A29B3, 0x9A29B3,
			0xDC3BFF/*checkme*/, 0xDD41FF, 0xDE4EFF, 0xE062FF,
			0xE47CFF, 0xE999FF, 0xEFB9FF, 0xF6DCFF,
			// e0
			0x633434, 0x2A0B31, 0x2A0B31, 0x601334,
			0x601334, 0x9C2270, 0x9D1D39, 0xE02C74,
			0xE02C74/*checkme*/, 0xE03374, 0xE24374, 0xE35DB6,
			0xE677B6/*checkme*/, 0xEB95B6/*checkme*/, 0xF1B6B6, 0xF6DCFF,
			// f0
			0xffffff/*checkme*/, 0xD9D9D9, 0xC5C5C5, 0xB3B3B3,
			0x8F8F8F, 0x6D6D6D, 0x4D4D4D, 0x3F3F3F,
			0x242424, 0x181818, 0/*checkme*/, 0/*checkme*/,
			0/*checkme*/, 0/*checkme*/, 0/*checkme*/, 0/*checkme*/,
		};
		for (uint32_t i : quattroColorMap)
			THIS->m_colorsList.push_back(WPSColor(i));
	}
	if (id < 0 || id >= int(m_colorsList.size()))
	{
		WPS_DEBUG_MSG(("Quattro9GraphInternal::State::getColor(): unknown Quattro9 Pro color id: %d\n",id));
		return false;
	}
	color = m_colorsList[size_t(id)];
	return true;
}

bool State::getPattern24(int id, WPSGraphicStyle::Pattern &pat)
{
	if (id<0 || id>24)
	{
		WPS_DEBUG_MSG(("Quattro9Internal::State::getPattern24(): unknown pattern id: %d\n",id));
		return false;
	}
	static const uint16_t  patterns[]=
	{
		0xffff, 0xffff, 0xffff, 0xffff, 0x0000, 0x0000, 0x0000, 0x0000, 0x00ff, 0x0000, 0x00ff, 0x0000, 0x0101, 0x0101, 0x0101, 0x0101, // 0-3
		0x8844, 0x2211, 0x8844, 0x2211, 0x8811, 0x2244, 0x8811, 0x2244, 0xff01, 0x0101, 0x0101, 0x0101, 0x040a, 0x11a0, 0x40a0, 0x110a, // 4-7
		0x44aa, 0x1100, 0x44aa, 0x1100, 0xffff, 0x0000, 0xffff, 0x0000, 0xaaaa, 0xaaaa, 0xaaaa, 0xaaaa, 0x060c, 0x1830, 0x60c0, 0x8103, // 8-9
		0xc060, 0x3018, 0x0c06, 0x0381, 0xc864, 0x3219, 0x8c46, 0x2391, 0xff11, 0xff11, 0xff11, 0xff11, 0xcccc, 0x3333, 0xcccc, 0x3333,
		0xcc33, 0xcc33, 0xcc33, 0xcc33, 0x0110, 0x0110, 0x0110, 0x0110, 0x1144, 0x1144, 0x1144, 0x1144, 0x070e, 0x9ee9, 0xe070, 0xb99b,
		0x0101, 0x01ff, 0x1010, 0x10ff, 0x4080, 0x0103, 0x8448, 0x3020, 0x2011, 0x0204, 0x0811, 0x8040, 0x00aa, 0x00aa, 0x00aa, 0x00aa,
		0xaa55, 0xaa55, 0xaa55, 0xaa55
	};
	pat.m_dim=Vec2i(8,8);
	uint16_t const *ptr=&patterns[4*id];
	pat.m_data.resize(8);
	for (size_t i=0; i < 8; i+=2)
	{
		uint16_t val=*(ptr++);
		pat.m_data[i]=static_cast<unsigned char>((val>>8) & 0xFF);
		pat.m_data[i+1]=static_cast<unsigned char>(val & 0xFF);
	}
	return true;
}

bool State::getPattern32(int id, WPSGraphicStyle::Pattern &pat)
{
	if (m_patterns32List.empty())
	{
		static const uint16_t  patterns[]=
		{
			0x0000, 0x0000, 0x0000, 0x0000, 0xaa88, 0xaa88, 0xaa88, 0xaa88,
			0x2288, 0x2288, 0x2288, 0x2288, 0x0022, 0x0088, 0x0022, 0x0088,
			0xff22, 0x2222, 0xff22, 0x2222, 0xff02, 0x0202, 0x0202, 0x0202,
			0x1028, 0x4482, 0x0182, 0x4428, 0x0000, 0x0000, 0x0000, 0x0000, /*7:16bit*/
			0xff00, 0x0000, 0xff00, 0x0000, 0xff00, 0x0000, 0x0000, 0x0000,
			0x1111, 0x1111, 0x1111, 0x1111, 0x1010, 0x1010, 0x1010, 0x1010,
			0x0102, 0x0408, 0x1020, 0x4080, 0x0000, 0x0000, 0x0000, 0x0000, /*13:16bit*/
			0x8040, 0x2010, 0x0804, 0x0201, 0x0000, 0x0000, 0x0000, 0x0000, /*15:16bit*/
			0x6666, 0x9999, 0x6666, 0x9999, 0xf0f0, 0xf0f0, 0x0f0f, 0x0f0f,
			0x0000, 0x0000, 0x0000, 0x0000, 0x2254, 0x8815, 0x2245, 0x8850, /*18:16bit*/
			0x0000, 0x0000, 0x0000, 0x0000, 0x3844, 0x8744, 0x3844, 0x8744, /*20:16bit*/
		};
		m_patterns32List.reserve(32);
		uint16_t const *ptr=&patterns[0];
		for (int p=0; p<22; ++p)
		{
			pat.m_dim=Vec2i(8,8);
			pat.m_data.resize(8);
			for (size_t i=0; i < 8; i+=2)
			{
				uint16_t val=*(ptr++);
				pat.m_data[i]=static_cast<unsigned char>((val>>8) & 0xFF);
				pat.m_data[i+1]=static_cast<unsigned char>(val & 0xFF);
			}
			m_patterns32List.push_back(pat);
		}
		static const uint16_t patterns16[]=
		{
			// 7
			0x0001, 0x8002, 0x4004, 0x2008, 0x1010, 0x0820, 0x0440, 0x0280,
			0x0100, 0x0280, 0x0440, 0x0820, 0x1010, 0x2008, 0x4004, 0x8002,
			// 13
			0x0001, 0x0002, 0x0004, 0x0008, 0x0010, 0x0020, 0x0040, 0x0080,
			0x0100, 0x0200, 0x0400, 0x0800, 0x1000, 0x2000, 0x4000, 0x8000,
			// 15
			0x0001, 0x8000, 0x4000, 0x2000, 0x1000, 0x0800, 0x0400, 0x0200,
			0x0100, 0x0080, 0x0040, 0x0020, 0x0010, 0x0008, 0x0004, 0x0002,
			// 18
			0xffff, 0xffff, 0x3300, 0x3300, 0x3300, 0x3300, 0x3300, 0x3300,
			0xffff, 0xffff, 0x0033, 0x0033, 0x0033, 0x0033, 0x0033, 0x0033,
			// 20
			0xa073, 0xffe1, 0x7f80, 0x0c00, 0x0c00, 0x0c00, 0x1e00, 0x3f00,
			0xf3c0, 0xe8ff, 0x807f, 0x000c, 0x000c, 0x000c, 0x001e, 0x003f,
			// 22
			0x8610, 0x6960, 0x1080, 0x1080, 0x1086, 0x6069, 0x8010, 0x8010,
			0x8610, 0x6960, 0x1080, 0x1080, 0x1086, 0x6069, 0x8010, 0x8010,
			// 23
			0x1212, 0x2121, 0x8080, 0x4040, 0xc0c0, 0xc0c0, 0x8080, 0x4040,
			0x2121, 0x1212, 0x0404, 0x0808, 0x0c0c, 0x0c0c, 0x0404, 0x0808,
			// 24
			0x1111, 0x8b8b, 0xc7c7, 0xa3a3, 0x1111, 0x3a3a, 0x7c7c, 0xb8b8,
			0x1111, 0x8b8b, 0xc7c7, 0xa3a3, 0x1111, 0x3a3a, 0x7c7c, 0xb8b8,
			// 25
			0xffff, 0x2a00, 0xffff, 0x2a00, 0xffff, 0x2a00, 0x2a00, 0x2a00,
			0x2a00, 0x2a00, 0x2a00, 0x2a00, 0x2a00, 0x2a00, 0x2a00, 0x2a00,
			// 26
			0x0ff0, 0x0ff0, 0x07e0, 0x03c0, 0xc003, 0x6004, 0xf008, 0xf00f,
			0xf00f, 0xf00f, 0xe007, 0xc003, 0x03c0, 0x0460, 0x08f0, 0x0ff0,
			// 27
			0x8686, 0x8686, 0x8686, 0x8686, 0xfefe, 0x0000, 0xfefe, 0xfefe,
			0x8686, 0x8686, 0x8686, 0x8686, 0xfefe, 0x0000, 0xfefe, 0xfefe,
			// 28
			0xe070, 0x0070, 0x0070, 0x07fc, 0x07fc, 0x07fc, 0x0070, 0x0070,
			0xe070, 0xe000, 0xe000, 0xfc03, 0xfc03, 0xfc03, 0xe000, 0xe000,
			// 29
			0x7f7f, 0x3f3f, 0x1f1f, 0x0f0f, 0x0707, 0x0303, 0x0101, 0x0101,
			0x7f7f, 0x3f3f, 0x1f1f, 0x0f0f, 0x0707, 0x0303, 0x0101, 0x0101,
			// 30
			0xc003, 0x600c, 0x8010, 0x0021, 0x00c1, 0x8300, 0x7c00, 0x0000,
			0x03c0, 0x0c60, 0x0180, 0x2100, 0xc100, 0x0083, 0x007c, 0x0000,
			// 31
			0xffff, 0xff03, 0xfe00, 0x7c00, 0x3800, 0x3800, 0x1000, 0x1000,
			0xffff, 0x03ff, 0x00fe, 0x007c, 0x0038, 0x0038, 0x0010, 0x0010,
		};
		int const idPatterns16[]= {7,13,15,18,20,22,23,24,25,26,27,28,29,30,31};
		m_patterns32List.resize(32);
		ptr=&patterns16[0];
		for (auto pId : idPatterns16)
		{
			pat.m_dim=Vec2i(16,16);
			pat.m_data.resize(32);
			for (size_t i=0; i < 32; i+=2)
			{
				uint16_t val=*(ptr++);
				pat.m_data[i]=static_cast<unsigned char>((val>>8) & 0xFF);
				pat.m_data[i+1]=static_cast<unsigned char>(val & 0xFF);
			}
			m_patterns32List[size_t(pId)]=pat;
		}
	}
	if (id<0 || id>=int(m_patterns32List.size()))
	{
		WPS_DEBUG_MSG(("Quattro9Internal::State::getPattern32(): unknown pattern id: %d\n",id));
		return false;
	}
	pat=m_patterns32List[size_t(id)];
	return true;
}

//! Internal: the subdocument of a Quattro9GraphInternal
class SubDocument final : public WKSSubDocument
{
public:
	//! constructor for a textbox document
	SubDocument(Quattro9Graph &graphParser, std::shared_ptr<Textbox> const &textbox, libwps_tools_win::Font::Type fontType)
		: WKSSubDocument(RVNGInputStreamPtr(), &graphParser.m_mainParser)
		, m_textbox(textbox)
		, m_fontType(fontType)
	{}
	//! destructor
	~SubDocument() final {}

	//! operator==
	bool operator==(std::shared_ptr<WPSSubDocument> const &doc) const final
	{
		if (!doc || !WKSSubDocument::operator==(doc))
			return false;
		auto const *sDoc = dynamic_cast<SubDocument const *>(doc.get());
		if (!sDoc) return false;
		return m_textbox.get() == sDoc->m_textbox.get() && m_fontType==sDoc->m_fontType;
	}

	//! the parser function
	void parse(std::shared_ptr<WKSContentListener> &listener, libwps::SubDocumentType subDocumentType) final;
	//! the textbox data
	std::shared_ptr<Textbox> m_textbox;
	//! the font type
	libwps_tools_win::Font::Type m_fontType;
};

void SubDocument::parse(std::shared_ptr<WKSContentListener> &listener, libwps::SubDocumentType)
{
	if (!listener.get())
	{
		WPS_DEBUG_MSG(("QuattroGraphInternal::SubDocument::parse: no listener\n"));
		return;
	}
	if (!dynamic_cast<WKSContentListener *>(listener.get()))
	{
		WPS_DEBUG_MSG(("QuattroGraphInternal::SubDocument::parse: bad listener\n"));
		return;
	}
	if (!m_textbox || !m_textbox->m_stream)
	{
		WPS_DEBUG_MSG(("QuattroGraphInternal::SubDocument::parse: can not find the textbox\n"));
		return;
	}
	listener->setParagraph(m_textbox->m_paragraph);
	WPSFont font=m_textbox->m_font;
	auto fontType=m_fontType;
	if (!font.m_name.empty())
	{
		fontType=libwps_tools_win::Font::getFontType(font.m_name);
		if (fontType==libwps_tools_win::Font::UNKNOWN)
			fontType=m_fontType;
	}
	listener->setFont(font);
	m_textbox->m_text.send(m_textbox->m_stream, font, fontType, listener);
}

}

// constructor, destructor
Quattro9Graph::Quattro9Graph(Quattro9Parser &parser)
	: m_listener()
	, m_mainParser(parser)
	, m_state(new Quattro9GraphInternal::State())
{
}

Quattro9Graph::~Quattro9Graph()
{
}

void Quattro9Graph::cleanState()
{
	m_state.reset(new Quattro9GraphInternal::State());
}

void Quattro9Graph::updateState()
{
}

int Quattro9Graph::version() const
{
	if (m_state->m_version<0)
		m_state->m_version=m_mainParser.version();
	return m_state->m_version;
}

void Quattro9Graph::storeObjects(std::map<librevenge::RVNGString,WPSEmbeddedObject> const &nameToObjectMap)
{
	m_state->m_linkNameToObjectMap=nameToObjectMap;
}

bool Quattro9Graph::getColor(int id, WPSColor &color) const
{
	return m_state->getColor(id, color);
}
bool Quattro9Graph::getPattern(int id, WPSGraphicStyle::Pattern &pattern) const
{
	return m_state->getPattern24(id, pattern);
}

////////////////////////////////////////////////////////////
// low level

////////////////////////////////////////////////////////////
// zones
////////////////////////////////////////////////////////////
bool Quattro9Graph::readBeginEnd(std::shared_ptr<WPSStream> stream, int sheetId)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type != 0x1401 && type != 0x1402)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readBeginEnd: not a begin/end zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	int const expectedSize=(type==0x1401 ? 6 : 0);
	m_state->m_actualGraph.reset();
	m_state->m_actualSheet=type==0x1401 ? sheetId : -1;
	if (sz!=expectedSize)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readBeginEnd: size seems very bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	if (type==0x1401)
	{
		f << "size=" << std::hex << libwps::readU32(input) << std::dec << ",";
		f << "type=" << std::hex << libwps::readU16(input) << std::dec << ","; // 79|b4
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool Quattro9Graph::readBeginEndZone(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type != 0x2001 && type != 0x2002)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readBeginEndZone: not a begin/end zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	int const expectedSize=(type==0x2001 ? 10 : 0);
	m_state->m_zoneDepth+=(type==0x2001 ? 1 : -1);
	if (m_state->m_zoneDepth<0)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readBeginEndZone: the zone depth seems bad\n"));
		m_state->m_zoneDepth=0;
	}
	if (type==0x2002 && m_state->m_actualGraphDepth>=m_state->m_zoneDepth)
		m_state->m_actualGraph.reset();
	if (sz!=expectedSize)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readBeginEndZone: size seems very bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	if (type==0x2001)
	{
		f << "size=" << std::hex << libwps::readU32(input) << std::dec << ",";
		f << "id=" << libwps::readU16(input) << ",";
		ascFile.addDelimiter(input->tell(), '|');
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool Quattro9Graph::readGraphHeader(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto id = int(libwps::readU16(input));
	if (id!=0x2051)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readGraphHeader: unknown shape type\n"));
		return false;
	}
	long sz = long(libwps::readU16(input));
	long endPos=input->tell()+sz;
	if (sz<0x3d || !stream->checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readGraphHeader: bad size\n"));
		return false;
	}
	auto frame=std::make_shared<Quattro9GraphInternal::Graph>(stream);
	m_state->m_actualGraph.reset();
	int dim[4];
	for (auto &d: dim) d=int(libwps::readU32(input));
	frame->m_cellBox=WPSBox2i(Vec2i(dim[0],dim[1]),Vec2i(dim[2],dim[3]));
	float fDim[4];
	for (auto &d: fDim) d=float(libwps::read32(input))/20.f;
	frame->m_cellBoxDecal=WPSBox2f(Vec2f(fDim[0],fDim[1]),Vec2f(fDim[2],fDim[3]));
	for (int i=0; i<2; ++i) fDim[i]=float(libwps::read32(input))/20.f;
	frame->m_size=Vec2f(fDim[0],fDim[1]);
	ascFile.addDelimiter(input->tell(),'|');
	input->seek(pos+60, librevenge::RVNG_SEEK_SET);
	ascFile.addDelimiter(input->tell(),'|');
	int fl=int(libwps::readU16(input));
	if ((fl&0x2000)==0) f << "protected=no,";
	fl &= 0xdfff;
	if (fl) f << "flags=" << std::hex << fl << std::dec << ",";
	ascFile.addDelimiter(input->tell(),'|');

	f << *frame << ",";
	m_state->storeGraph(frame);
	if (input->tell()!=endPos)
		ascFile.addDelimiter(input->tell(),'|');
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	return true;
}

bool Quattro9Graph::readShape(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;

	auto id = int(libwps::readU16(input));
	bool const bigBlock=id&0x8000;
	id &= 0x7fff;
	if (id!=0x2221 && id!=0x23d1)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readShape: unknown shape type\n"));
		return false;
	}
	long sz = bigBlock ? long(libwps::readU32(input)) : long(libwps::readU16(input));
	long endPos=input->tell()+sz;
	if (sz<4 || !stream->checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readShape: bad size\n"));
		return false;
	}
	Quattro9GraphInternal::Shape shape;
	if (readShapeRec(stream, endPos, shape, WPSGraphicStyle::emptyStyle()) && id==0x2221)
	{
		auto graph=m_state->m_actualGraph;
		if (!graph)
		{
			WPS_DEBUG_MSG(("Quattro9Graph::readShape: can not find the graphic frame\n"));
		}
		else
		{
			graph->m_type=graph->Shape;
			graph->m_shape.reset(new Quattro9GraphInternal::Shape(shape));
		}
	}
	return true;
}

bool Quattro9Graph::readShapeRec(std::shared_ptr<WPSStream> const &stream, long endPos, Quattro9GraphInternal::Shape &zone, WPSGraphicStyle const &actualStyle)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	WPSGraphicShape shape;
	WPSGraphicStyle style=actualStyle;
	WPSColor surfColor[2]= {WPSColor::white(), WPSColor::black()}; // fill, pattern/gradiant
	int patId=-1;
	while (input->tell()+4<endPos)
	{
		long pos=input->tell();
		int type1=int(libwps::readU8(input)); // normally 4/6
		int type2=int(libwps::readU16(input));
		f.str("");
		if (type1==4)
			f << "ShapeMain-" << std::hex << type2 << std::dec << ":";
		else if (type1==6)
			f << "ShapeShadow-" << std::hex << type2 << std::dec << ":";
		else
			f << "Shape-Data" << type1 << "-" << std::hex << type2 << std::dec << ":";
		int dSz=int(libwps::readU8(input));
		if (dSz==0xFF) dSz=int(libwps::readU16(input));
		long endFieldPos=input->tell()+dSz;
		if (endFieldPos>endPos)
		{
			input->seek(pos, librevenge::RVNG_SEEK_SET);
			break;
		}
		if (type1==4)
		{
			int val;
			if (type2>=0x12 && type2<=0x19 && patId>=0)
			{
				// we must update the style here
				if (patId==0)
					style.setSurfaceColor(surfColor[0]);
				else if (patId>0)
				{
					WPSGraphicStyle::Pattern pattern;
					if (m_state->getPattern32(patId, pattern))
					{
						pattern.m_colors[0]=surfColor[1];
						pattern.m_colors[1]=surfColor[0];
						WPSColor color;
						if (pattern.getUniqueColor(color))
							style.setSurfaceColor(color);
						else
							style.setPattern(pattern);
					}
					else
					{
						WPS_DEBUG_MSG(("Quattro9Graph::readShapeRec: can not find the graphic pattern=%d\n", patId));
					}
				}
				patId=-1;
			}
			switch (type2)
			{
			case 0x12:
			{
				if (dSz<2) break;
				val=int(libwps::readU16(input)); // 0
				if (val) f << "f0=" << val << ",";
				if (!zone.m_child) zone.m_child.reset(new Quattro9GraphInternal::Shape);
				readShapeRec(stream, endFieldPos, *zone.m_child, style);
				break;
			}
			case 0x15:   // list of points of shape
			{
				if (dSz<4) break;
				int fl=int(libwps::readU16(input)); // 2000|8000
				f << "fl=" << std::hex << fl << std::dec << ",";
				int N=int(libwps::readU16(input));
				if (N<1 || 4+4*N!=dSz)
				{
					f << "###N=" << N << ",";
					break;
				}
				f << "pts=[";
				std::vector<Vec2f> vertices;
				vertices.resize(size_t(N));
				for (auto &pt : vertices)
				{
					float coord[2];
					for (auto &c : coord) c=float(libwps::readU16(input))/20.f;
					pt=Vec2f(coord[0],coord[1]);
					f << pt << ",";
				}
				f << "],";
				if (N==2)
				{
					shape=WPSGraphicShape::line(vertices[0], vertices[1]);
					zone.m_listShapes.push_back(std::make_pair(shape,style));
				}
				else if (N>2)
				{
					WPSBox2f box(vertices[0], vertices[0]);
					for (size_t pt=1; pt<size_t(N); ++pt)
						box=box.getUnion(WPSBox2f(vertices[pt],vertices[pt]));
					if (fl&0x2000)
						shape=WPSGraphicShape::polygon(WPSBox2f(box));
					else
						shape=WPSGraphicShape::polyline(WPSBox2f(box));
					shape.m_vertices=vertices;
					zone.m_listShapes.push_back(std::make_pair(shape,style));
				}
				break;
			}
			case 0x17:   // list of spline points of shape
			{
				if (dSz<4) break;
				int fl=int(libwps::readU16(input)); // 2000|8000
				f << "fl=" << std::hex << fl << std::dec << ",";
				int N=int(libwps::readU16(input));
				if (4+12*N!=dSz)
				{
					f << "###N=" << N << ",";
					break;
				}
				f << "pts=[";
				std::vector<Vec2f> vertices;
				vertices.reserve(3*size_t(N));
				WPSBox2f box;
				for (int pt=0; pt<N; ++pt)
				{
					f << "[";
					for (int co=0; co<3; ++co)
					{
						float coord[2];
						for (auto &c : coord) c=float(libwps::readU16(input))/20.f;
						vertices.push_back(Vec2f(coord[0],coord[1]));
						f << vertices.back() << ",";
						if (pt==0 && co==0)
							box=WPSBox2f(vertices.back(), vertices.back());
						else
							box=box.getUnion(WPSBox2f(vertices.back(), vertices.back()));
					}
					f << "],";
				}
				f << "],";
				if (N<=1) break;
				shape=WPSGraphicShape::path(WPSBox2f(box));
				shape.m_path.push_back(WPSGraphicShape::PathData('M', vertices[1]));
				size_t numPts=vertices.size()/3;
				for (size_t pt=1; pt<numPts; ++pt)
				{
					if (vertices[3*pt-2]==vertices[3*pt-1] && vertices[3*pt]==vertices[3*pt+1])
						shape.m_path.push_back(WPSGraphicShape::PathData('L', vertices[3*pt+1]));
					else
						shape.m_path.push_back(WPSGraphicShape::PathData('C', vertices[3*pt+1], vertices[3*pt-1], vertices[3*pt]));
				}
				if (fl&0x2000)
					shape.m_path.push_back(WPSGraphicShape::PathData('Z'));
				zone.m_listShapes.push_back(std::make_pair(shape,style));
				break;
			}
			case 0x18:   // rectangle
			{
				if (dSz!=0xe) break;
				int fl=int(libwps::readU16(input)); // a000
				f << "fl=" << std::hex << fl << std::dec << ",";
				f << "pts=[";
				Vec2f pts[2];
				for (auto &pt : pts)
				{
					float coord[2];
					for (auto &c : coord) c=float(libwps::readU16(input))/20.f;
					pt=Vec2f(coord[0],coord[1]);
					f << pt << ",";
				}
				f << "],";
				float corner[2];
				for (auto &d : corner) d=float(libwps::readU16(input))/20.f;
				if (corner[0]>0 || corner[1]>0) f << "corner=" << Vec2f(corner[0],corner[1]) << ",";
				shape=WPSGraphicShape::rectangle(WPSBox2f(pts[0],pts[1]), Vec2f(corner[0],corner[1]));
				zone.m_listShapes.push_back(std::make_pair(shape,style));
				break;
			}
			case 0x19:   // oval
			{
				if (dSz!=0x14) break;
				int fl=int(libwps::readU16(input)); // a000
				f << "fl=" << std::hex << fl << std::dec << ",";
				f << "pts=[";
				Vec2f pts[2];
				for (auto &pt : pts)
				{
					float coord[2];
					for (auto &c : coord) c=float(libwps::readU16(input))/20.f;
					pt=Vec2f(coord[0],coord[1]);
					f << pt << ",";
				}
				f << "],";
				f << "unkn=["; // -17.5, 0, -17.5, 0, 0
				for (int i=0; i<5; ++i) f << float(libwps::read16(input))/20.f << ",";
				f << "],";
				shape=WPSGraphicShape::circle(WPSBox2f(pts[0]-pts[1],pts[0]+pts[1]));
				zone.m_listShapes.push_back(std::make_pair(shape,style));
				break;
			}
			case 0x25: // line
			case 0x33: // second color
			{
				if (dSz!=4) break;
				unsigned char col[4];
				for (auto &c : col) c=static_cast<unsigned char>(libwps::readU8(input));
				WPSColor color(col[0],col[1],col[2]);
				if (!color.isBlack())
					f << (type2==0x25 ? "line" : "color[fill2]") << "=" << color << ",";
				if (type2==0x25)
					style.m_lineColor=color;
				else
					surfColor[1]=color;
				break;
			}
			case 0x29:
			{
				val=int(libwps::readU16(input));
				switch (val)
				{
				case 0: // default
					break;
				case 1:
					style.m_lineDashWidth= {4,1};
					break;
				case 2:
					style.m_lineDashWidth= {3,1};
					break;
				case 3:
					style.m_lineDashWidth= {2,2};
					break;
				case 4:
					style.m_lineDashWidth= {2,1};
					break;
				case 5:
					style.m_lineDashWidth= {1,1};
					break;
				case 6:
					style.m_lineDashWidth= {1,2};
					break;
				case 7:
					style.m_lineDashWidth= {1,1,1,1,1,4};
					break;
				case 8:
					style.m_lineDashWidth= {4,1,1,1};
					break;
				case 9:
					style.m_lineDashWidth= {3,1,1,1,1,1};
					break;
				case 10:
					style.m_lineDashWidth= {2,1,2,1,1,1};
					break;
				case 11:
					style.m_lineDashWidth= {4,1,4,1,1,1};
					break;
				case 12:
					style.m_lineDashWidth= {4,1,2,1};
					break;
				case 13:
					style.m_lineDashWidth= {4,1,1,1,1,1};
					break;
				case 14:
					style.m_lineDashWidth= {6,1};
					break;
				default:
					f << "#dash=" << std::hex << val << std::dec << ",";
					break;
				}
				if (!style.m_lineDashWidth.empty())
				{
					f << "dash=";
					for (auto const &d : style.m_lineDashWidth) f << d << ":";
					f << ",";
				}
				break;
			}
			case 0x2a: // 0
				if (dSz!=2) break;
				val=int(libwps::readU16(input));
				if (val) f << "f0=" << std::hex << val << std::dec << ",";
				break;
			case 0x2b:
			{
				if (dSz!=4) break;
				int values[2];
				for (auto &v : values) v=int(libwps::readU16(input));
				if (values[0]!=values[1])
					f << "pen[size]=" << Vec2i(values[0],values[1]) << ",";
				else if (values[0])
					f << "pen[size]=" << values[0] << ",";
				if (values[0]+values[1])
					style.m_lineWidth=float(values[0]+values[1])/2.f/13.f;
				else
					style.m_lineWidth=1;
				break;
			}
			case 0x2d:
				if (dSz!=2) break;
				for (int wh=0; wh<2; ++wh)
				{
					val=int(libwps::readU8(input));
					if (!val) continue;
					f << "arrow[" << (wh==0 ? "start" : "end") << "=" << val << ",";
					style.m_arrows[wh]=true;
				}
				break;
			case 0x2e: // 2
				if (dSz!=1) break;
				val=int(libwps::readU8(input));
				switch (val)
				{
				case 1: // checkme
					style.m_lineJoin=style.J_Bevel;
					f << "bevel,";
					break;
				case 2: // default
					break;
				case 3:
					style.m_lineJoin=style.J_Round;
					f << "round,";
					break;
				default:
					f << "#join=" << val << ",";
					break;
				}
				break;
			case 0x2f:
			{
				if (dSz!=9) break;
				val=int(libwps::readU16(input));
				if (val) f << "f0=" << val << ",";
				style.m_gradientAngle=int(libwps::read16(input));
				f << "angle[grad]=" << style.m_gradientAngle << ",";
				float center[2];
				for (auto &c: center) c=float(libwps::readU16(input))/65535.f;
				style.m_gradientPercentCenter=Vec2f(center[0],center[1]);
				f << "center=" << style.m_gradientPercentCenter << ",";
				val=int(libwps::readU8(input)); // radius?
				if (val!=100) f << "f1=" << val << ",";
				break;
			}
			case 0x31:   // fill color
			{
				if (dSz==0) break;
				int type=int(libwps::readU8(input));
				if (dSz==2)
				{
					f << "inherit[" << type << "],";
					break;
				}
				if (dSz==5 && type==0)
				{
					unsigned char col[4];
					for (auto &c : col) c=static_cast<unsigned char>(libwps::readU8(input));
					surfColor[0]=WPSColor(col[0],col[1],col[2]);
					if (!surfColor[0].isWhite())
						f << "color[fill]=" << surfColor[0] << ",";
				}
				else if (dSz==0xd && (type==1 || type==3 || type==9))
				{
					f << "gradient[" << type << "],";
					val=int(libwps::readU16(input));
					if (val!=2) f << "f0=" << val << ",";
					for (int wh=0; wh<2; ++wh)
					{
						unsigned char col[4];
						for (auto &c : col) c=static_cast<unsigned char>(libwps::readU8(input));
						surfColor[wh]=WPSColor(col[0],col[1],col[2]);
						if ((wh==0 && !surfColor[wh].isWhite()) || (wh==1 && !surfColor[wh].isBlack()))
							f << "color[grad" << wh << "]=" << surfColor[wh] << ",";
					}

					style.m_gradientStopList.clear();
					if (type==1)
					{
						float mPos=style.m_gradientPercentCenter[0]+style.m_gradientPercentCenter[1];
						if (style.m_gradientAngle>40 && style.m_gradientAngle<50) mPos*=0.5f;
						else if (style.m_gradientAngle>-50 && style.m_gradientAngle<-40) mPos=2*style.m_gradientPercentCenter[1];
						style.m_gradientType=style.G_Linear;
						if (mPos<0.05f)
						{
							style.m_gradientStopList.push_back(WPSGraphicStyle::GradientStop(0.0,surfColor[0]));
							style.m_gradientStopList.push_back(WPSGraphicStyle::GradientStop(1.0,surfColor[1]));
						}
						else if (mPos>0.95f)
						{
							style.m_gradientStopList.push_back(WPSGraphicStyle::GradientStop(0.0,surfColor[1]));
							style.m_gradientStopList.push_back(WPSGraphicStyle::GradientStop(1.0,surfColor[0]));
						}
						else if (mPos>0.45f && mPos<0.55f)
						{
							style.m_gradientType=style.G_Axial;
							style.m_gradientStopList.push_back(WPSGraphicStyle::GradientStop(0.0,surfColor[0]));
							style.m_gradientStopList.push_back(WPSGraphicStyle::GradientStop(1.0,surfColor[1]));
						}
						else
						{
							style.m_gradientStopList.push_back(WPSGraphicStyle::GradientStop(0.0,surfColor[1]));
							style.m_gradientStopList.push_back(WPSGraphicStyle::GradientStop(mPos,surfColor[0]));
							style.m_gradientStopList.push_back(WPSGraphicStyle::GradientStop(1.0,surfColor[1]));
						}
					}
					else
					{
						style.m_gradientType=type==9 ? style.G_Square : style.G_Radial;
						style.m_gradientStopList.push_back(WPSGraphicStyle::GradientStop(0.0,surfColor[0]));
						style.m_gradientStopList.push_back(WPSGraphicStyle::GradientStop(1.0,surfColor[1]));
					}
					val=int(libwps::readU16(input));
					if (val!=1) f << "f1=" << val << ",";
				}
				break;
			}
			case 0x35: // 0
				if (dSz!=2) break;
				patId=int(libwps::readU16(input));
				if (patId) f << "pat[id]=" << patId << ",";
				break;
			case 0x42:   // main bdbox
			{
				// child bdbox
				if (dSz!=0xe) break;
				val=int(libwps::readU16(input)); // a000
				if (val) f << "fl=" << std::hex << val << std::dec << ",";
				float fDim[4];
				for (auto &d: fDim) d=float(libwps::read16(input))/20.f;
				zone.m_box=WPSBox2f(Vec2f(fDim[0],fDim[1]),Vec2f(fDim[2],fDim[3]));
				f << "box=" << zone.m_box << ",";
				break;
			}
			case 0x120: //
			case 0x620: //
			{
				if (dSz!=0xa) break;
				val=int(libwps::readU16(input)); // a000
				if (val) f << "fl=" << std::hex << val << std::dec << ",";
				float fDim[4];
				for (auto &d: fDim) d=float(libwps::read16(input))/20.f;
				zone.m_box=WPSBox2f(Vec2f(fDim[0],fDim[1]),Vec2f(fDim[2],fDim[3]));
				f << "box=" << zone.m_box << ",";
				break;
			}
			// 0x3d: sz=7, related to picture filled
			// 0x3e: sz=9, related to picture filled
			case 0x221:
			{
				if (dSz!=0x1c) break;
				val=int(libwps::readU16(input)); // a000
				if (val) f << "fl=" << std::hex << val << std::dec << ",";
				float fDim[4];
				for (auto &d: fDim) d=float(libwps::read16(input))/20.f;
				zone.m_box=WPSBox2f(Vec2f(fDim[0],fDim[1]),Vec2f(fDim[2],fDim[3]));
				f << "box=" << zone.m_box << ",";
				for (int j=0; j<9; ++j)
				{
					val=int(libwps::readU16(input));
					int const expected[]= {2, 0x102, 0xa0, 0x4b0, 0x4b0, 2, 0x80, 0x4b0, 0x4b0};
					if (val==expected[j]) continue;
					f << "f" << j << "=" << val << ",";
				}
				break;
			}
			case 0x1020: // related to shadow: e000 maybe the type
				if (dSz!=2) break;
				val=int(libwps::readU16(input));
				if (val) f << "f0=" << std::hex << val << std::dec << ",";
				break;
			default:
				break;
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
		WPS_DEBUG_MSG(("Quattro9Graph::readShapeRec: find extra data\n"));
		ascFile.addPos(input->tell());
		ascFile.addNote("Shape:###extra");
	}
	return true;
}

bool Quattro9Graph::readFrameHeader(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto id = int(libwps::readU16(input));
	if (id!=0x2171)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readFrameHeader: unknown shape type\n"));
		return false;
	}
	long sz = long(libwps::readU16(input));
	long endPos=input->tell()+sz;
	if (sz<0x2a || !stream->checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readFrameHeader: bad size\n"));
		return false;
	}
	int val=int(libwps::readU16(input)); // 102, 288, 36a
	if (val) f << "f0=" << std::hex << val << std::dec << ",";
	float dim[4];
	for (auto &d: dim) d=float(libwps::readU32(input))/20.f;
	f << "dim=" << WPSBox2f(Vec2f(dim[0],dim[1]),Vec2f(dim[2],dim[3])) << ",";
	ascFile.addDelimiter(input->tell(),'|');
	input->seek(pos+38, librevenge::RVNG_SEEK_SET);
	ascFile.addDelimiter(input->tell(),'|');
	for (int wh=0; wh<2; ++wh)
	{
		unsigned char col[4];
		for (auto &c: col) c=static_cast<unsigned char>(libwps::readU8(input));
		WPSColor color(col[0],col[1],col[2]);
		if ((wh==0 && !color.isWhite()) || (wh==1 && !color.isBlack()))
			f << (wh==0 ? "surf" : "line") << "[color]=" << color << ",";
	}
	if (input->tell()!=endPos)
		ascFile.addDelimiter(input->tell(),'|');
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	return true;
}

bool Quattro9Graph::readFramePattern(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto id = int(libwps::readU16(input));
	if (id!=0x2141)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readFramePattern: unknown shape type\n"));
		return false;
	}
	long sz = long(libwps::readU16(input));
	long endPos=input->tell()+sz;
	if (sz<0x8 || !stream->checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readFramePattern: bad size\n"));
		return false;
	}
	int val=int(libwps::readU16(input));
	if (val) f << "pat[id]=" << val << ",";
	for (int i=0; i<3; ++i)   // f1=0,1,6
	{
		val=int(libwps::readU16(input));
		if (val) f << "f" << i << "=" << val << ",";
	}
	if (input->tell()!=endPos)
		ascFile.addDelimiter(input->tell(),'|');
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	return true;
}

bool Quattro9Graph::readFrameStyle(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto id = int(libwps::readU16(input));
	if (id!=0x2131)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readFrameStyle: unknown shape type\n"));
		return false;
	}
	long sz = long(libwps::readU16(input));
	long endPos=input->tell()+sz;
	if (sz<0xc || !stream->checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readFrameStyle: bad size\n"));
		return false;
	}
	unsigned char col[4];
	for (auto &c: col) c=static_cast<unsigned char>(libwps::readU8(input));
	WPSColor color(col[0],col[1],col[2]);
	if (color!=WPSColor(128,128,128))
		f << "surf2[color]=" << color << ",";
	int val=int(libwps::readU16(input)); // 0|1
	if (val) f << "line[width]=" << val << ",";
	val=int(libwps::readU16(input));
	switch (val)
	{
	case 0: // none
		break;
	case 1:
		f << "pattern,";
		break;
	case 5:
		f << "gradient,";
		break;
	case 0x1001:
		f << "bitmap,";
		break;
	default:
		f << "type=" << val << ",";
	}
	val=int(libwps::readU16(input)); // 0|1
	if (val) f << "f0=" << val << ",";
	if (input->tell()!=endPos)
		ascFile.addDelimiter(input->tell(),'|');
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	return true;
}

bool Quattro9Graph::readTextboxStyle(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto id = int(libwps::readU16(input));
	if (id!=0x2371)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readTextboxStyle: unknown zone's type\n"));
		return false;
	}
	long sz = long(libwps::readU16(input));
	long endPos=input->tell()+sz;
	if (sz<0x67 || !stream->checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readTextboxStyle: bad size\n"));
		return false;
	}
	int val=int(libwps::readU8(input)); // f,19,1f
	if (val) f << "fl=" << std::hex << val << std::dec << ",";
	WPSFont font;
	WPSParagraph para;
	WPSGraphicStyle style(WPSGraphicStyle::emptyStyle());
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

	auto const fontType = m_mainParser.getDefaultFontType();
	std::string name;
	for (long i=0; i<32; ++i)
	{
		auto c = char(libwps::readU8(input));
		if (c == '\0') break;
		name.push_back(c);
	}
	if (!name.empty())
		font.m_name=libwps_tools_win::Font::unicodeString(name, fontType);
	input->seek(pos+37, librevenge::RVNG_SEEK_SET);
	for (int i=0; i<2; ++i)
	{
		val=int(libwps::read16(input));
		if (val) f << "f" << i << "=" << val << ",";
	}
	unsigned char col[4];
	for (auto &c: col) c=static_cast<unsigned char>(libwps::readU8(input));
	font.m_color=WPSColor(col[0],col[1],col[2]);
	f << font;
	for (auto &c: col) c=static_cast<unsigned char>(libwps::readU8(input));
	WPSColor color(col[0],col[1],col[2]);
	if (!color.isWhite())
	{
		style.setBackgroundColor(color);
		f << "background[color]=" << color << ",";
	}
	for (int i=0; i<4; ++i)   // f2=0|8505,f4=1
	{
		val=int(libwps::readU16(input));
		if (!val) continue;
		if (i==2)
		{
			f << "line[style]=" << val << ",";
			style.m_lineWidth=1;
		}
		else
			f << "f" << i+2 << "=" << std::hex << val << std::dec << ",";
	}
	val=int(libwps::read16(input));
	switch (val)
	{
	case 0:
		break;
	case 1:
		para.m_justify=libwps::JustificationCenter;
		f << "center,";
		break;
	case 2:
		para.m_justify=libwps::JustificationRight;
		f << "right,";
		break;
	default:
		f << "##align=" << val << ",";
		break;
	}
	for (int i=0; i<3; ++i)   // f7=1
	{
		val=int(libwps::read16(input));
		if (val) f << "f" << i+6 << "=" << val << ",";
	}
	f << "n[current]=" << int(libwps::readU16(input)) << ",";
	ascFile.addDelimiter(input->tell(),'|');
	input->seek(10, librevenge::RVNG_SEEK_CUR);
	ascFile.addDelimiter(input->tell(),'|');
	val=int(libwps::read16(input));
	if (val!=0x12c) f << "tabs=" << float(val)/300.f << ","; // useme: unit inch
	auto graph=m_state->m_actualGraph;
	if (!graph)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readTextboxStyle: can not find the graphic frame\n"));
	}
	else
	{
		if (!graph->m_textbox) graph->m_textbox.reset(new Quattro9GraphInternal::Textbox);
		graph->m_textbox->m_font=font;
		graph->m_textbox->m_paragraph=para;
		graph->m_textbox->m_style=style;
	}

	if (input->tell()!=endPos)
		ascFile.addDelimiter(input->tell(),'|');
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	return true;
}

bool Quattro9Graph::readOLEName(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto id = int(libwps::readU16(input));
	if (id!=0x21d1)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readOLEName: unknown zone's type\n"));
		return false;
	}
	long sz = long(libwps::readU16(input));
	long endPos=input->tell()+sz;
	if (sz<2 || !stream->checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readOLEName: bad size\n"));
		return false;
	}
	Quattro9ParserInternal::TextEntry entry;
	std::string name;
	if (m_mainParser.readPString(stream, endPos, entry))
	{
		name=entry.getDebugString(stream);
		f << name << ",";
	}
	else
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readOLEName: can not read a string\n"));
		f << "###";
		ascFile.addDelimiter(input->tell(),'|');
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}

	auto graph=m_state->m_actualGraph;
	if (!graph)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readTextboxText: can not find the graphic frame\n"));
	}
	else
	{
		if (!graph->m_textbox) graph->m_textbox.reset(new Quattro9GraphInternal::Textbox);
		graph->m_type=graph->OLE;
		graph->m_OLEName=name.c_str();
	}

	if (input->tell()!=endPos)
		ascFile.addDelimiter(input->tell(),'|');
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool Quattro9Graph::readTextboxText(std::shared_ptr<WPSStream> const &stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto id = int(libwps::readU16(input));
	if ((id&0x7fff)!=0x2372)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readTextboxText: unknown zone's type\n"));
		return false;
	}
	long sz = ((id&0x8000) ? long(libwps::readU32(input)) : long(libwps::readU16(input)));
	long endPos=input->tell()+sz;
	if (sz<2 || !stream->checkFilePosition(endPos))
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readTextboxText: bad size\n"));
		return false;
	}
	Quattro9ParserInternal::TextEntry entry;
	if (m_mainParser.readPString(stream, endPos, entry))
		f << entry.getDebugString(stream) << ",";
	else
	{
		WPS_DEBUG_MSG(("Quattro9Parser::readTextboxText: can not read a string\n"));
		f << "###";
		ascFile.addDelimiter(input->tell(),'|');
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}

	auto graph=m_state->m_actualGraph;
	if (!graph)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::readTextboxText: can not find the graphic frame\n"));
	}
	else
	{
		if (!graph->m_textbox) graph->m_textbox.reset(new Quattro9GraphInternal::Textbox);
		graph->m_type=graph->Textbox;
		graph->m_textbox->m_stream=stream;
		graph->m_textbox->m_text=entry;
	}

	if (input->tell()!=endPos)
		ascFile.addDelimiter(input->tell(),'|');
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

////////////////////////////////////////////////////////////
// send data
////////////////////////////////////////////////////////////
bool Quattro9Graph::sendShape(Quattro9GraphInternal::Graph const &graph, int sheetId) const
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::sendShape: can not find the listener\n"));
		return false;
	}
	if (graph.m_type!=graph.Shape || !graph.m_shape)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::sendShape: can not find the shape\n"));
		return false;
	}
	auto const &shape=*graph.m_shape;
	if (shape.empty())
	{
		WPS_DEBUG_MSG(("Quattro9Graph::sendShape: the list of shape is empty\n"));
		return false;
	}
	Vec2f mainOrigin=graph.m_cellBoxDecal[0]+m_mainParser.getCellPosition(sheetId, graph.m_cellBox[0]);
	// rescale (Y axis is inverted) and translate the points so that the origin is preserved
	auto const bdbox=shape.getBdBox();
	auto scale=WPSTransformation::scale(Vec2f(bdbox.size()[0]>0 ? float(graph.m_size[0])/bdbox.size()[0] : 1.f,
	                                          bdbox.size()[1]>0 ? -float(graph.m_size[1])/bdbox.size()[1] : -1.f));
	WPSTransformation transf=WPSTransformation::translation(mainOrigin-scale*Vec2f(bdbox[0][0],bdbox[1][1]))*scale;
	sendShape(shape, transf);
	return true;
}

bool Quattro9Graph::sendShape(Quattro9GraphInternal::Shape const &shape, WPSTransformation const &transf) const
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::sendShape: can not find the listener\n"));
		return false;
	}
	for (auto const &sh : shape.m_listShapes)
		sendShape(sh.first, sh.second, transf);
	if (shape.m_child)
	{
		auto const bdbox=shape.getBdBox();
		WPSPosition pos(transf*bdbox[0], transf.multiplyDirection(bdbox.size()), librevenge::RVNG_POINT);
		pos.m_anchorTo = WPSPosition::Page;
		m_listener->openGroup(pos);
		sendShape(*shape.m_child, transf);
		m_listener->closeGroup();
	}
	return true;
}

bool Quattro9Graph::sendShape(WPSGraphicShape const &shape, WPSGraphicStyle const &style, WPSTransformation const &transf) const
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::sendShape: can not find the listener\n"));
		return false;
	}
	auto const bdbox=shape.getBdBox();
	WPSPosition pos(transf*bdbox[0], transf.multiplyDirection(bdbox.size()), librevenge::RVNG_POINT);
	pos.m_anchorTo = WPSPosition::Page;
	m_listener->insertPicture(pos, shape.transform(transf), style);
	return true;
}

bool Quattro9Graph::sendOLE(Quattro9GraphInternal::Graph const &graph, int sheetId) const
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::sendOLE: can not find the listener\n"));
		return false;
	}
	if (graph.m_type!=graph.OLE || graph.m_OLEName.empty())
	{
		WPS_DEBUG_MSG(("Quattro9Graph::sendOLE: can not find the OLE\n"));
		return false;
	}
	WPSPosition pos(graph.m_cellBoxDecal[0]+m_mainParser.getCellPosition(sheetId, graph.m_cellBox[0]), graph.m_size, librevenge::RVNG_POINT);
	pos.m_anchorTo = WPSPosition::Page;
	auto it=m_state->m_linkNameToObjectMap.find(graph.m_OLEName);
	if (it==m_state->m_linkNameToObjectMap.end() || it->second.isEmpty())
	{
		WPS_DEBUG_MSG(("Quattro9Graph::sendOLE: can not find ole %s\n", graph.m_OLEName.cstr()));
	}
	else
		m_listener->insertObject(pos, it->second);
	return true;
}

bool Quattro9Graph::sendTextbox(Quattro9GraphInternal::Graph const &graph, int sheetId) const
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::sendTextbox: can not find the listener\n"));
		return false;
	}
	if (graph.m_type!=graph.Textbox || !graph.m_textbox)
	{
		WPS_DEBUG_MSG(("Quattro9Graph::sendTextbox: can not find the textbox\n"));
		return false;
	}
	WPSPosition pos(graph.m_cellBoxDecal[0]+m_mainParser.getCellPosition(sheetId, graph.m_cellBox[0]), graph.m_size, librevenge::RVNG_POINT);
	pos.m_anchorTo = WPSPosition::Page;
	auto doc=std::make_shared<Quattro9GraphInternal::SubDocument>(const_cast<Quattro9Graph &>(*this),graph.m_textbox, m_mainParser.getDefaultFontType());
	m_listener->insertTextBox(pos, doc, graph.m_textbox->m_style);
	return true;
}

bool Quattro9Graph::sendPageGraphics(int sheetId) const
{
	for (auto it=m_state->m_sheetIdToGraphMap.lower_bound(sheetId); it!=m_state->m_sheetIdToGraphMap.upper_bound(sheetId); ++it)
	{
		auto &graph=it->second;
		if (!graph) continue;
		if (graph->m_type==graph->OLE)
			sendOLE(*graph, sheetId);
		else if (graph->m_type==graph->Shape)
			sendShape(*graph, sheetId);
		else if (graph->m_type==graph->Textbox)
			sendTextbox(*graph, sheetId);
	}
	return true;
}

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
