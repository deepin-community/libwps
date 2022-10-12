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

#include "Quattro.h"

#include "QuattroGraph.h"

namespace QuattroGraphInternal
{
//! a dialog header
struct Dialog
{
	//! constructor
	explicit Dialog()
		: m_cellBox()
	{
		for (auto &f : m_flags1) f=0;
		for (auto &f : m_flags2) f=0;
	}
	//! operator<<
	friend std::ostream &operator<<(std::ostream &o, Dialog const &gr);
	//! the cell's position
	WPSBox2i m_cellBox;
	//! some flags
	int m_flags1[5];
	//! final flag
	int m_flags2[9];
};

std::ostream &operator<<(std::ostream &o, Dialog const &dlg)
{
	if (dlg.m_cellBox!=WPSBox2i()) o << "cellBox=" << dlg.m_cellBox << ",";
	o << "fl1=[";
	for (auto f : dlg.m_flags1)
	{
		if (f)
			o << std::hex << f << std::dec << ",";
		else
			o << "_,";
	}
	o << "],";
	o << "fl2=[";
	for (auto f : dlg.m_flags2)
	{
		if (f)
			o << std::hex << f << std::dec << ",";
		else
			o << "_,";
	}
	o << "],";
	return o;
}

//! a shape header of a QuattroGraph
struct ShapeHeader
{
	//! constructor
	ShapeHeader()
		: m_type(0)
		, m_box()
		, m_id(0)
		, m_style()
		, m_extra()
	{
		for (auto &f : m_flags) f=0;
		for (auto &v : m_values) v=0;
		for (auto &v : m_values2) v=0;
	}
	//! destructor
	virtual ~ShapeHeader();
	//! returns true if the shape is a textbox
	virtual bool isTextbox() const
	{
		return false;
	}
	//! operator<<
	friend std::ostream &operator<<(std::ostream &o, ShapeHeader const &sh);
	//! the type
	int m_type;
	//! the bdbox
	WPSBox2i m_box;
	//! an id?
	int m_id;
	//! the graphic style
	WPSGraphicStyle m_style;
	//! optional values
	int m_values[5];
	//! some flags
	int m_flags[14];

	//
	// style
	//

	//! style values
	int m_values2[4];
	//! error message
	std::string m_extra;
};

ShapeHeader::~ShapeHeader()
{
}

std::ostream &operator<<(std::ostream &o, ShapeHeader const &sh)
{
	o << "type=" << sh.m_type << ",";
	o << "box=" << sh.m_box << ",";
	if (sh.m_id) o << "id=" << sh.m_id << ",";
	o << sh.m_style << ",";
	int wh=0;
	for (auto v : sh.m_values)
	{
		if (v) o << "f" << wh << "=" << v << ",";
		++wh;
	}
	o << "unkn[";
	for (auto f : sh.m_flags)
	{
		if (f)
			o << std::hex << f << std::dec << ",";
		else
			o << ",";
	}
	o << "],";
	wh=0;
	for (auto v : sh.m_values2)
	{
		if (v) o << "g" << wh << "=" << v << ",";
		++wh;
	}
	o << sh.m_extra << ",";
	return o;
}

//! Internal: a shape of a QuattroGraph
struct Shape final : public ShapeHeader
{
	//! constructor
	Shape()
		: ShapeHeader()
		, m_shape()
	{
	}
	//! destructor
	~Shape() final;
	//! the graphic shape
	WPSGraphicShape m_shape;
};

Shape::~Shape()
{
}

//! Internal: a shape of a QuattroGraph
struct Textbox final : public ShapeHeader
{
	//! constructor
	Textbox()
		: ShapeHeader()
		, m_entry()
		, m_font()
		, m_paragraph()
	{
	}
	//! destructor
	~Textbox() final;
	//! returns true
	bool isTextbox() const final
	{
		return true;
	}
	//! the text entry
	WPSEntry m_entry;
	//! the font
	WPSFont m_font;
	//! the paragraph style
	WPSParagraph m_paragraph;
};

Textbox::~Textbox()
{
}

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
		, m_label()
		, m_ole()
		, m_linkName()
		, m_shape()
		, m_textbox()
		, m_stream(stream)
	{
		for (auto &f : m_flags1) f=0;
		for (auto &f : m_flags2) f=0;
		for (auto &v : m_values) v=0;
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
	//! some flags
	int m_flags1[4];
	//! final flag
	int m_flags2[7];
	//! some values
	int m_values[5];

	//! the label(button)
	librevenge::RVNGString m_label;

	//! the OLE's data
	WPSEmbeddedObject m_ole;
	//! the OLE's link name
	librevenge::RVNGString m_linkName;

	//! the graphic shape
	std::shared_ptr<QuattroGraphInternal::Shape> m_shape;
	//! the textbox
	std::shared_ptr<QuattroGraphInternal::Textbox> m_textbox;

	//! the main stream
	std::shared_ptr<WPSStream> m_stream;
};

std::ostream &operator<<(std::ostream &o, Graph const &gr)
{
	if (gr.m_size!=Vec2f()) o << "size=" << gr.m_size << ",";
	if (gr.m_cellBox!=WPSBox2i()) o << "cellBox=" << gr.m_cellBox << ",";
	if (gr.m_cellBoxDecal!=WPSBox2f()) o << "cellBox[decal]=" << gr.m_cellBoxDecal << ",";
	o << "fl1=[";
	for (auto f : gr.m_flags1)
	{
		if (f)
			o << std::hex << f << std::dec << ",";
		else
			o << "_,";
	}
	o << "],";
	o << "fl2=[";
	for (auto f : gr.m_flags2)
	{
		if (f)
			o << std::hex << f << std::dec << ",";
		else
			o << "_,";
	}
	o << "],";
	for (int i=0; i<5; ++i)
	{
		if (gr.m_values[i])
			o << "f" << i << "=" << gr.m_values[i] << ",";
	}
	return o;
}

//! the state of QuattroGraph
struct State
{
	//! constructor
	State()
		: m_version(-1)
		, m_actualSheet(-1)
		, m_sheetIdToGraphMap()
		, m_actualGraph()
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
		if (m_actualSheet<0)
		{
			WPS_DEBUG_MSG(("QuattroGraphInternal::storeGraph: can not find the current sheet\n"));
			return;
		}
		m_sheetIdToGraphMap.insert(std::multimap<int, std::shared_ptr<Graph> >::value_type(m_actualSheet, graph));
	}
	//! returns the pattern corresponding to a pattern id between 0 and 24
	static bool getPattern(int id, WPSGraphicStyle::Pattern &pattern);
	//! the file version
	int m_version;
	//! the actual sheet id
	int m_actualSheet;
	//! a multimap sheetId to graph
	std::multimap<int, std::shared_ptr<Graph> > m_sheetIdToGraphMap;
	//! the actual graph
	std::shared_ptr<Graph> m_actualGraph;
	//! a map link name to object
	std::map<librevenge::RVNGString,WPSEmbeddedObject> m_linkNameToObjectMap;
};

bool State::getPattern(int id, WPSGraphicStyle::Pattern &pat)
{
	if (id<0 || id>24)
	{
		WPS_DEBUG_MSG(("QuattroInternal::State::getPattern(): unknown pattern id: %d\n",id));
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

//! Internal: the subdocument of a QuattroGraphInternal
class SubDocument final : public WKSSubDocument
{
public:
	//! constructor for a textbox document
	SubDocument(QuattroGraph &graphParser, std::shared_ptr<Textbox> const &textbox, std::shared_ptr<WPSStream> const &stream)
		: WKSSubDocument(RVNGInputStreamPtr(), &graphParser.m_mainParser)
		, m_graphParser(graphParser)
		, m_textbox(textbox)
		, m_stream(stream)
		, m_text() {}
	//! constructor for a text entry
	SubDocument(QuattroGraph &graphParser, librevenge::RVNGString const &text)
		: WKSSubDocument(RVNGInputStreamPtr(), &graphParser.m_mainParser)
		, m_graphParser(graphParser)
		, m_textbox()
		, m_stream()
		, m_text(text) {}
	//! destructor
	~SubDocument() final {}

	//! operator==
	bool operator==(std::shared_ptr<WPSSubDocument> const &doc) const final
	{
		if (!doc || !WKSSubDocument::operator==(doc))
			return false;
		auto const *sDoc = dynamic_cast<SubDocument const *>(doc.get());
		if (!sDoc) return false;
		if (&m_graphParser != &sDoc->m_graphParser) return false;
		if (m_textbox.get() != sDoc->m_textbox.get()) return false;
		if (m_stream.get() != sDoc->m_stream.get()) return false;
		return m_text == sDoc->m_text;
	}

	//! the parser function
	void parse(std::shared_ptr<WKSContentListener> &listener, libwps::SubDocumentType subDocumentType) final;
	//! the graph parser
	QuattroGraph &m_graphParser;
	//! the textbox data
	std::shared_ptr<Textbox> m_textbox;
	//! the file stream
	std::shared_ptr<WPSStream> m_stream;
	//! the main text
	librevenge::RVNGString m_text;
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
	if (m_textbox && m_stream)
	{
		auto input=m_stream->m_input;
		long actPos=input->tell();
		m_graphParser.send(*m_textbox, m_stream);
		input->seek(actPos, librevenge::RVNG_SEEK_SET);
		return;
	}
	WPSParagraph para;
	para.m_justify=libwps::JustificationCenter;
	listener->setParagraph(para);
	if (!m_text.empty())
		listener->insertUnicodeString(m_text);
}

}

// constructor, destructor
QuattroGraph::QuattroGraph(QuattroParser &parser)
	: m_listener()
	, m_mainParser(parser)
	, m_state(new QuattroGraphInternal::State())
{
}

QuattroGraph::~QuattroGraph()
{
}

void QuattroGraph::cleanState()
{
	m_state.reset(new QuattroGraphInternal::State());
}

void QuattroGraph::updateState()
{
}

int QuattroGraph::version() const
{
	if (m_state->m_version<0)
		m_state->m_version=m_mainParser.version();
	return m_state->m_version;
}

void QuattroGraph::storeObjects(std::map<librevenge::RVNGString,WPSEmbeddedObject> const &nameToObjectMap)
{
	m_state->m_linkNameToObjectMap=nameToObjectMap;
}

std::vector<Vec2i> QuattroGraph::getGraphicCellsInSheet(int sheetId) const
{
	std::vector<Vec2i> list;
	auto it=m_state->m_sheetIdToGraphMap.find(sheetId);
	while (it!=m_state->m_sheetIdToGraphMap.end() && it->first==sheetId)
	{
		auto const &graph=it++->second;
		if (graph && graph->m_type!=graph->Shape && graph->m_type!=graph->Textbox)
			list.push_back(graph->m_cellBox[0]);
	}
	return list;
}

////////////////////////////////////////////////////////////
// low level

////////////////////////////////////////////////////////////
// zones
////////////////////////////////////////////////////////////
bool QuattroGraph::readHeader(QuattroGraphInternal::Graph &header, std::shared_ptr<WPSStream> stream, long endPos)
{
	auto input=stream->m_input;
	long pos = input->tell();
	if (endPos-pos<49)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readHeader: the zone is too short\n"));
		return false;
	}
	for (auto &fl : header.m_flags1) fl=int(libwps::readU16(input));
	int dim[4];
	for (auto &d: dim) d=libwps::readU16(input);
	header.m_cellBox=WPSBox2i(Vec2i(dim[0],dim[1]),Vec2i(dim[2],dim[3]));
	float fDim[4];
	for (auto &d: fDim) d=float(libwps::read16(input))/20.f;
	header.m_cellBoxDecal=WPSBox2f(Vec2f(fDim[0],fDim[1]),Vec2f(fDim[2],fDim[3]));
	for (int i=0; i<2; ++i) fDim[i]=float(libwps::read32(input))/20.f;
	header.m_size=Vec2f(fDim[0],fDim[1]);
	for (auto &fl : header.m_flags2) fl=int(libwps::readU8(input));
	for (auto &v : header.m_values) v=int(libwps::read16(input));
	return true;
}

bool QuattroGraph::readShapeHeader(QuattroGraphInternal::ShapeHeader &shape, std::shared_ptr<WPSStream> stream, long endPos)
{
	int const vers=version();
	auto input=stream->m_input;
	long pos = input->tell();
	int const endSize=15+(vers==1003 ? 3 : 0);
	if (endPos-pos<42+(vers==1003 ? 4 : 0))
	{
		WPS_DEBUG_MSG(("QuattroGraph::readShapeHeader: the zone is too short\n"));
		return false;
	}
	libwps::DebugStream f;
	shape.m_type=int(libwps::readU16(input)); // v1: 6f, v2: 73, v3: 9e
	int wFl=0;
	for (int i=0; i<4; ++i)
		shape.m_flags[wFl++]=int(libwps::readU16(input));
	int dim[4];
	for (auto &d: dim) d=int(libwps::read16(input));
	shape.m_box=WPSBox2i(Vec2i(dim[0],dim[1]),Vec2i(dim[2],dim[3]));
	for (int i=0; i<(vers>=1003 ? 7 : 5); ++i)
		shape.m_flags[wFl++]=int(libwps::readU16(input));
	shape.m_id=int(libwps::readU16(input));
	auto &style=shape.m_style;
	WPSColor surfaceColor[2];
	unsigned char col[4];
	for (auto &c : col) c=libwps::readU8(input);
	surfaceColor[0]=WPSColor(col[0],col[1],col[2]);
	for (auto &c : col) c=libwps::readU8(input);
	style.m_lineColor=WPSColor(col[0],col[1],col[2]);
	shape.m_flags[wFl++]=int(libwps::readU16(input));
	int hasData[2];
	for (auto &d : hasData) d=int(libwps::readU8(input));
	if (hasData[0]==1 && endPos-input->tell()>=3+endSize)
	{
		shape.m_values[0]=int(libwps::read8(input));
		shape.m_values[1]=int(libwps::read16(input));
	}
	else if (hasData[0])
	{
		WPS_DEBUG_MSG(("QuattroGraph::readShapeHeader: find unexpected data0 %d\n", hasData[0]));
		return false;
	}
	if (hasData[1]==1 && endPos-input->tell()>=6+endSize)
	{
		for (int i=0; i<3; ++i)
			shape.m_values[i+2]=int(libwps::read16(input));
	}
	else if (hasData[1])
	{
		WPS_DEBUG_MSG(("QuattroGraph::readShapeHeader: find unexpected data1 %d\n", hasData[1]));
		return false;
	}

	// end data
	shape.m_values2[0]=int(libwps::readU8(input));
	if (vers>=1003)
	{
		for (int i=0; i<2; ++i)
		{
			int val=i==1 ? int(libwps::read8(input)) : int(libwps::read16(input));
			shape.m_values2[1+i]=val;
		}
	}
	int patternId=int(libwps::readU16(input)); // 154: solid
	int lineStyle=int(libwps::readU16(input));
	switch (lineStyle)
	{
	case 1: // solid
		break;
	case 2: // dash
		style.m_lineDashWidth.push_back(float(4));
		style.m_lineDashWidth.push_back(float(1));
		break;
	case 3: // dots
		style.m_lineDashWidth.resize(2,float(1));
		break;
	case 4: // dash/dots
		style.m_lineDashWidth.resize(4,float(1));
		style.m_lineDashWidth[0]=float(4);
		break;
	case 5: //
		style.m_lineDashWidth.resize(6,float(1));
		style.m_lineDashWidth[0]=float(4);
		break;
	case 6: // empty
		style.m_lineWidth=0;
		break;
	default:
		WPS_DEBUG_MSG(("QuattroGraph::readShapeHeader: find unknown line style\n"));
		f << "line[style]=##" << lineStyle << ",";
		break;
	}
	for (auto &c : col) c=libwps::readU8(input);
	surfaceColor[1]=WPSColor(col[0],col[1],col[2]);
	int lineWidth=int(libwps::readU16(input));
	if (style.m_lineWidth>0) style.m_lineWidth=float(lineWidth);
	// 0: none, 1: pattern, 5: gradient(wb3), 1001: bitmap
	int fillType=int(libwps::readU16(input));
	shape.m_values2[3]=int(libwps::readU16(input));
	bool isTextbox=shape.isTextbox();
	if (fillType==0)
	{
		WPSGraphicStyle::Pattern pattern;
		if (patternId==0) // none
			;
		else if (patternId==1)
		{
			if (isTextbox)
				style.setBackgroundColor(surfaceColor[1]);
			else
				style.setSurfaceColor(surfaceColor[1]);
		}
		else if (patternId==154)
		{
			if (isTextbox)
				style.setBackgroundColor(surfaceColor[0]);
			else
				style.setSurfaceColor(surfaceColor[0]);
		}
		else if (m_state->getPattern(patternId, pattern))
		{
			for (int i=0; i<2; ++i)
				pattern.m_colors[i]=surfaceColor[i];
			if (isTextbox)
			{
				WPSColor finalColor;
				if (pattern.getAverageColor(finalColor))
					style.setBackgroundColor(finalColor);
			}
			else
				style.setPattern(pattern);
		}
		else
			f << "###pat[id]=" << patternId << ",";
	}
	else
	{
		if (!readFillData(shape.m_style, fillType, stream, endPos))
			return false;
		if (fillType>=1 && fillType<=6)
		{
			f << "gradient=" << fillType << ",";
			if (isTextbox)
				style.setBackgroundColor(WPSColor::barycenter(0.5f, surfaceColor[0], 0.5f, surfaceColor[1]));
			else
			{
				style.m_gradientType=fillType<=4 ? WPSGraphicStyle::G_Linear : WPSGraphicStyle::G_Axial;
				style.m_gradientStopList.clear();
				style.m_gradientStopList.push_back(WPSGraphicStyle::GradientStop(0.0, surfaceColor[1]));
				style.m_gradientStopList.push_back(WPSGraphicStyle::GradientStop(1.0, surfaceColor[0]));
				int const rot[]= {0, 90, -90, 0, 180, 90, 0};
				style.m_gradientAngle=float(rot[fillType]);
			}
		}
		else
		{
			if (!surfaceColor[0].isBlack()) f << "surf[col0]=" << surfaceColor[0] << ",";
			if (!surfaceColor[1].isWhite()) f << "surf[col1]=" << surfaceColor[1] << ",";
			f << "fill[type]=" << fillType << ",";
			f << "pat[id]=" << patternId << ",";
			if ((fillType&0xfff)==1)
			{
				f << "###bitmap[" << std::hex << fillType << std::dec << "],";
				f << "crop[type]=" << (fillType>>12) << ",";
				stream->m_ascii.addDelimiter(input->tell(),'|');
				shape.m_extra=f.str();
				WPS_DEBUG_MSG(("QuattroGraph::readShapeHeader: find a bitmap, unimplemented!!!\n"));
				return false;
			}
			f << "###fill[type]=" << std::hex << fillType << std::dec << ",";
			WPS_DEBUG_MSG(("QuattroGraph::readShapeHeader: unexpected fill type %d\n", fillType));
		}
	}
	shape.m_extra=f.str();
	return true;
}

bool QuattroGraph::readFillData(WPSGraphicStyle &/*style*/, int fillType, std::shared_ptr<WPSStream> stream, long endPos)
{
	if (fillType==0) return true;
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	if (fillType<0)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readFillData: unexpected fillType\n"));
		return false;
	}
	if (pos+4>endPos || libwps::readU16(input)!=0x2e4)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readFillData: the zone length seems bad\n"));
		return false;
	}
	f << "Entries(FillData)[" << std::hex << fillType << std::dec << ":";
	int dSz=int(libwps::readU16(input));
	if (pos+4+dSz>endPos)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readFillData: can not read the data size\n"));
		return false;
	}
	if (dSz)
	{
		ascFile.addDelimiter(input->tell(),'|');
		input->seek(pos+4+dSz, librevenge::RVNG_SEEK_SET);
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	if ((fillType&0xf000)==0 || (fillType&0xfff)!=1)
		return true;
	pos=input->tell();
	if (pos+68>endPos)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readFillData: can not read the bitmap name\n"));
		return false;
	}
	f.str("");
	f << "FillData[bitmap]:";
	int val=int(libwps::readU16(input));
	if (val!=10) f << "f0=" << val << ",";
	val=int(libwps::readU16(input));
	if (val) f << "crop[type]=" << val << ",";
	librevenge::RVNGString name;
	if (!m_mainParser.readCString(stream,name,64))
		f << "###name,";
	else
		f << name.cstr() << ",";
	input->seek(pos+68, librevenge::RVNG_SEEK_SET);
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	pos=input->tell();
	if (pos+10>endPos)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readFillData: can not read the bitmap data\n"));
		return false;
	}
	f.str("");
	f << "FillData[extra]:";
	val=int(libwps::readU16(input));
	if (val!=0x4000) f << "f0=" << val << ",";
	val=int(libwps::readU16(input));
	if (val!=0x1c93) f << "f1=" << std::hex << val << std::dec << ",";
	int dim[2];
	for (auto &d:dim) d=int(libwps::readU16(input));
	f << "dim=" << Vec2i(dim[0],dim[1]) << ",";
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	return true;
}

bool QuattroGraph::readBeginEnd(std::shared_ptr<WPSStream> stream, int sheetId)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type != 0x321 && type != 0x322)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readBeginEnd: not a begin/end zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	int const expectedSize=(type==0x321 ? 0 : 2);
	m_state->m_actualGraph.reset();
	m_state->m_actualSheet=type==0x321 ? sheetId : -1;
	if (sz!=expectedSize)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readBeginEnd: size seems very bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	if (type==0x322)   // always 0
	{
		auto val=int(libwps::read16(input));
		if (val) f << "f0=" << val << ",";
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroGraph::readFrame(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type != 0x385)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readFrame: not a frame zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	auto frame=std::make_shared<QuattroGraphInternal::Graph>(stream,QuattroGraphInternal::Graph::Frame);
	m_state->m_actualGraph.reset();
	if (sz<57 || !readHeader(*frame,stream,endPos))
	{
		if (sz)
		{
			WPS_DEBUG_MSG(("QuattroGraph::readFrame: size seems very bad\n"));
			f << "###";
		}
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << *frame;
	m_state->storeGraph(frame);
	auto sSz=int(libwps::readU16(input));
	librevenge::RVNGString text;
	if (input->tell()+sSz+6>endPos || !m_mainParser.readCString(stream,text,sSz))
	{
		WPS_DEBUG_MSG(("QuattroGraph::readFrame: can not read string1\n"));
		f << "##sSz,";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << "name=" << text.cstr() << ",";
	for (int i=0; i<3; ++i)   // g0=1, g2=2001
	{
		auto val=int(libwps::readU16(input));
		if (val)
			f << "g" << i << "=" << std::hex << val << std::dec << ",";
	}
	if (input->tell()!=endPos)
	{
		ascFile.addDelimiter(input->tell(),'|');
		WPS_DEBUG_MSG(("QuattroGraph::readFrame: find extra data\n"));
		f << "##extra,";
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroGraph::readFrameOLE(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type != 0x381)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readFrameOLE: not a frame zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	auto frame=std::make_shared<QuattroGraphInternal::Graph>(stream,QuattroGraphInternal::Graph::OLE);
	m_state->m_actualGraph.reset();
	if (sz<59 || !readHeader(*frame,stream,endPos))
	{
		if (sz)
		{
			WPS_DEBUG_MSG(("QuattroGraph::readFrameOLE: size seems very bad\n"));
			f << "###";
		}
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << *frame;
	m_state->storeGraph(frame);
	auto sSz=int(libwps::readU16(input));
	librevenge::RVNGString text;
	if (input->tell()+sSz+4>endPos || !m_mainParser.readCString(stream,text,sSz))
	{
		WPS_DEBUG_MSG(("QuattroGraph::readFrameOLE: can not read string1\n"));
		f << "##sSz,";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	frame->m_linkName=text;
	f << "name=" << text.cstr() << ",";
	for (int i=0; i<4; ++i)   // g0=11, g2=d00, g4=7500
	{
		auto val=int(libwps::readU16(input));
		if (val)
			f << "g" << i << "=" << std::hex << val << std::dec << ",";
	}
	if (input->tell()!=endPos)
	{
		ascFile.addDelimiter(input->tell(),'|');
		WPS_DEBUG_MSG(("QuattroGraph::readFrameOLE: find extra data\n"));
		f << "##extra,";
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroGraph::readOLEData(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type != 0x38b)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readOLEData: not a OLE zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=sz<0xFF00 ? pos+4+sz : stream->m_eof;
	if (sz<38)
	{
		if (sz)
		{
			WPS_DEBUG_MSG(("QuattroGraph::readOLEData: size seems very bad\n"));
			f << "###";
		}
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto frame=m_state->m_actualGraph;
	if (frame && frame->m_type!=frame->Frame)
		frame.reset();
	if (frame)
		frame->m_type=frame->OLE;
	else
	{
		WPS_DEBUG_MSG(("QuattroGraph::readOLEData: can not find current frame\n"));
	}
	for (int i=0; i<5; ++i)
	{
		int val=int(libwps::readU16(input));
		int const expected[]= {0x1a,0x8068,0x2001,0,0};
		if (val!=expected[i])
			f << "f" << i << "=" << std::hex << val << std::dec << ",";
	}
	long actPos=input->tell();
	auto sSz=int(libwps::readU16(input));
	librevenge::RVNGString text;
	if (actPos+2+sSz+12+1+12>endPos || !m_mainParser.readCString(stream,text,sSz))
	{
		WPS_DEBUG_MSG(("QuattroGraph::readOLEData: can not read the name\n"));
		f << "##sSz,";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << "name=" << text.cstr() << ",";
	input->seek(actPos+2+sSz, librevenge::RVNG_SEEK_SET);
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	pos=input->tell();
	WPSEmbeddedObject dummyObject;
	if (!WPSOLEObject::readOLE(stream, frame ? frame->m_ole : dummyObject,endPos))
		input->seek(pos, librevenge::RVNG_SEEK_SET);
	if (input->tell()!=endPos)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readOLEData: find extra data\n"));
		ascFile.addPos(input->tell());
		ascFile.addNote("Object:###extra");
	}
	return true;
}

bool QuattroGraph::readButton(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type != 0x386)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readButton: not a button zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	m_state->m_actualGraph.reset();
	auto button=std::make_shared<QuattroGraphInternal::Graph>(stream,QuattroGraphInternal::Graph::Button);
	if (sz<67 || !readHeader(*button,stream,endPos))
	{
		if (sz)
		{
			WPS_DEBUG_MSG(("QuattroGraph::readButton: size seems very bad\n"));
			f << "###";
		}
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << *button; // fl1=[7f|81,0x8063,0x2000,0];
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	pos=input->tell();
	f.str("");
	f << "Object-A:";
	auto sSz=int(libwps::readU16(input));
	librevenge::RVNGString text;
	if (pos+2+sSz>endPos || !m_mainParser.readCString(stream,text,sSz))
	{
		WPS_DEBUG_MSG(("QuattroGraph::readButton: can not read string1 bad\n"));
		f << "##sSz,";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	input->seek(pos+2+sSz, librevenge::RVNG_SEEK_SET);
	f << "name=" << text.cstr() << ",";
	for (int i=0; i<5; ++i)
	{
		auto val=int(libwps::readU16(input));
		if (val) f << "f" << i << "=" << val << ",";
	}
	auto val=int(libwps::readU8(input));
	if (val) f << "f5=" << val << ",";
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	pos=input->tell();
	f.str("");
	f << "Object-B:";
	auto dType=int(libwps::readU8(input));
	if (dType==1) f << "complex,";
	else if (dType)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readButton: find unknown type\n"));
		f << "##dType=" << dType << ",";
	}
	for (int st=0; st<2; ++st)
	{
		sSz=int(libwps::readU16(input));
		if (pos+2+sSz>endPos || !m_mainParser.readCString(stream,text,sSz))
		{
			WPS_DEBUG_MSG(("QuattroGraph::readButton: can not read string2 bad\n"));
			f << "##sSz,";
			ascFile.addPos(pos);
			ascFile.addNote(f.str().c_str());
			return true;
		}
		if (text.empty()) continue;
		f << (st==0 ? "macros" : "label") << "=" << text.cstr() << ",";
		if (st==1)
			button->m_label=text;
	}
	if (dType==0)
	{
		if (input->tell()!=endPos)
		{
			WPS_DEBUG_MSG(("QuattroGraph::readButton: find extra data\n"));
			f << "##extra,";
			ascFile.addDelimiter(input->tell(),'|');
		}
		m_state->storeGraph(button);
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	if (dType && input->tell()!=endPos)
	{
		ascFile.addPos(input->tell());
		ascFile.addNote("Object-C:");
	}
	return true;
}

bool QuattroGraph::readImage(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type != 0x382)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readImage: unknown id\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	auto zone382=std::make_shared<QuattroGraphInternal::Graph>(stream,QuattroGraphInternal::Graph::Image);
	m_state->m_actualGraph.reset();
	if (sz<53 || !readHeader(*zone382,stream,endPos))
	{
		if (sz)
		{
			WPS_DEBUG_MSG(("QuattroGraph::readImage: size seems very bad\n"));
			f << "###";
		}
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << *zone382;
	auto sSz=int(libwps::readU16(input));
	librevenge::RVNGString text;
	if (input->tell()+sSz+2>endPos || !m_mainParser.readCString(stream,text,sSz))
	{
		WPS_DEBUG_MSG(("QuattroGraph::readImage: can not read string1\n"));
		f << "##sSz,";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << text.cstr() << ","; // find Bitmap5, followed by
	// 20000000070001000080000000ffffff0000000100000000070064e96729a1e87328af2b0700628597334100
	// c767040064e96729b7e87328af2b04008bce67354800
	// c7670500673575ce0400
	// 2d002d00: a picture dimension ?
	// eb010000: the gif's size
	// the gif
	if (input->tell()!=endPos)
	{
		ascFile.addDelimiter(input->tell(),'|');
	}
	static bool first=true;
	if (first)
	{
		first=false;
		WPS_DEBUG_MSG(("QuattroGraph::readImage: this file contains a zone 382, there will not be recovered\n"));
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroGraph::readBitmap(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type != 0x383)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readBitmap: unknown id\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	auto bitmap=std::make_shared<QuattroGraphInternal::Graph>(stream,QuattroGraphInternal::Graph::OLE);
	m_state->m_actualGraph.reset();
	if (sz<67 || !readHeader(*bitmap,stream,endPos))
	{
		if (sz)
		{
			WPS_DEBUG_MSG(("QuattroGraph::readBitmap: size seems very bad\n"));
			f << "###";
		}
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << *bitmap;
	auto sSz=int(libwps::readU16(input));
	librevenge::RVNGString text;
	if (input->tell()+sSz+16>endPos || !m_mainParser.readCString(stream,text,sSz))
	{
		WPS_DEBUG_MSG(("QuattroGraph::readBitmap: can not read string1\n"));
		f << "##sSz,";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << text.cstr() << ",";
	f << "unkn=[";
	for (int i=0; i<8; ++i)
	{
		auto val=int(libwps::readU16(input));
		if (val)
			f << std::hex << val << std::dec << ",";
		else
			f << "_,";
	}
	f << "],";
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	WPSEmbeddedObject object;
	pos=input->tell();
	if (!WPSOLEObject::readWMF(stream, bitmap->m_ole, endPos))
	{
		WPS_DEBUG_MSG(("QuattroGraph::readBitmap: can not find the wmf file\n"));
		ascFile.addPos(pos);
		ascFile.addNote("Object:###");
	}
	else
		m_state->storeGraph(bitmap);

	return true;
}

bool QuattroGraph::readChart(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type != 0x384)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readChart: unknown id\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	auto chart=std::make_shared<QuattroGraphInternal::Graph>(stream,QuattroGraphInternal::Graph::Chart);
	m_state->m_actualGraph.reset();
	if (sz<57 || !readHeader(*chart,stream,endPos))
	{
		if (sz)
		{
			WPS_DEBUG_MSG(("QuattroGraph::readChart: size seems very bad\n"));
			f << "###";
		}
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << *chart;
	auto sSz=int(libwps::readU16(input));
	librevenge::RVNGString text;
	if (input->tell()+sSz+6>endPos || !m_mainParser.readCString(stream,text,sSz))
	{
		WPS_DEBUG_MSG(("QuattroGraph::readChart: can not read string1\n"));
		f << "##sSz,";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << text.cstr() << ","; // find Inserted1-21
	for (int i=0; i<2; ++i)   // f0=1|30
	{
		auto val=int(libwps::read16(input));
		if (val) f << "f" << i << "=" << val << ",";
	}
	sSz=int(libwps::readU16(input));
	if (input->tell()+sSz>endPos || !m_mainParser.readCString(stream,text,sSz))
	{
		WPS_DEBUG_MSG(("QuattroGraph::readChart: can not read string1\n"));
		f << "##sSz2,";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << "name=" << text.cstr() << ",";
	if (input->tell()!=endPos)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readChart: find extra data\n"));
		f << "##extra,";
		ascFile.addDelimiter(input->tell(),'|');
	}
	static bool first=true;
	if (first)
	{
		first=false;
		WPS_DEBUG_MSG(("QuattroGraph::readChart: this file contains some charts, there will not be recovered\n"));
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

////////////////////////////////////////////////////////////
// shape, line, ...
////////////////////////////////////////////////////////////
bool QuattroGraph::readShape(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type != 0x4d3)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readShape: not a shape zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	auto shape=std::make_shared<QuattroGraphInternal::Graph>(stream,QuattroGraphInternal::Graph::Shape);
	m_state->m_actualGraph.reset();
	if (sz<59 || !readHeader(*shape,stream,endPos))
	{
		if (sz)
		{
			WPS_DEBUG_MSG(("QuattroGraph::readShape: size seems very bad\n"));
			f << "###";
		}
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << *shape;
	auto sSz=int(libwps::readU16(input));
	librevenge::RVNGString text;
	if (input->tell()+sSz+8>endPos || !m_mainParser.readCString(stream,text,sSz))
	{
		WPS_DEBUG_MSG(("QuattroGraph::readShape: can not read string1\n"));
		f << "##sSz,";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	shape->m_linkName=text;
	f << "name=" << text.cstr() << ",";
	for (int i=0; i<4; ++i)   // f2=1b(oval) 37(line),41(rect), 56(rect oval), 58(arrow)
	{
		int val=int(libwps::read16(input));
		int const expected[]= {0x20,0,0,0x2001};
		if (val!=expected[i])
			f << "f" << i << "=" << val << ",";
	}
	m_state->storeGraph(shape);
	if (input->tell()!=endPos)
	{
		ascFile.addDelimiter(input->tell(),'|');
		f << "##extra,";
		WPS_DEBUG_MSG(("QuattroGraph::readShape: find extra data\n"));
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroGraph::readLine(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);
	if (type!=0x35a && type!=0x37b)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readLine: not a line zone\n"));
		return false;
	}
	auto parent=m_state->m_actualGraph;
	m_state->m_actualGraph.reset();
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	auto shape=std::make_shared<QuattroGraphInternal::Shape>();
	if (sz<58 || !readShapeHeader(*shape, stream, endPos-1) || input->tell()+1>endPos)
	{
		if (sz)
		{
			WPS_DEBUG_MSG(("QuattroGraph::readLine: the size seems very bad\n"));
			f << shape << ",";
			f << "###";
		}
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << *shape << ",";
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	pos=input->tell();
	f.str("");
	f << "GrLine:";
	int val=int(libwps::readU8(input));
	f << "type=" << val << ",";
	if (input->tell()!=endPos)
	{
		ascFile.addDelimiter(input->tell(),'|');
		f << "##extra,";
		WPS_DEBUG_MSG(("QuattroGraph::readLine: find extra data\n"));
	}
	shape->m_style.m_arrows[1]=type==0x37b;
	switch (val&3)
	{
	case 1:
		shape->m_shape=WPSGraphicShape::line(Vec2f(float(shape->m_box[1][0]),float(shape->m_box[0][1])),
		                                     Vec2f(float(shape->m_box[0][0]),float(shape->m_box[1][1])));
		break;
	case 2:
		shape->m_shape=WPSGraphicShape::line(Vec2f(shape->m_box[1]),Vec2f(shape->m_box[0]));
		break;
	case 3:
		shape->m_shape=WPSGraphicShape::line(Vec2f(float(shape->m_box[0][0]),float(shape->m_box[1][1])),
		                                     Vec2f(float(shape->m_box[1][0]),float(shape->m_box[0][1])));
		break;
	case 0:
	default:
		shape->m_shape=WPSGraphicShape::line(Vec2f(shape->m_box[0]),Vec2f(shape->m_box[1]));
		break;
	}
	if (parent && parent->m_type==parent->Shape)
		parent->m_shape=shape;

	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroGraph::readRect(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	auto parent=m_state->m_actualGraph;
	m_state->m_actualGraph.reset();
	if (type!=0x33e && type!=0x364 && type!=0x379)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readRect: not a rect zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	auto shape=std::make_shared<QuattroGraphInternal::Shape>();
	if (sz<57 || !readShapeHeader(*shape, stream, endPos))
	{
		if (sz)
		{
			WPS_DEBUG_MSG(("QuattroGraph::readRect: the size seems very bad\n"));
			f << *shape << ",";
			f << "###";
		}
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << *shape << ",";
	switch (type)
	{
	case 0x33e:
		shape->m_shape=WPSGraphicShape::circle(WPSBox2f(shape->m_box));
		break;
	case 0x364:
		shape->m_shape=WPSGraphicShape::rectangle(WPSBox2f(shape->m_box));
		break;
	default:
	case 0x379:
		shape->m_shape=WPSGraphicShape::rectangle(WPSBox2f(shape->m_box), Vec2f(20,20));
		break;
	}
	if (parent && parent->m_type==parent->Shape)
		parent->m_shape=shape;
	if (input->tell()!=endPos)
	{
		ascFile.addDelimiter(input->tell(),'|');
		f << "##extra,";
		WPS_DEBUG_MSG(("QuattroGraph::readRect: find extra data\n"));
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroGraph::readPolygon(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);
	if (type!=0x35b && type!=0x35c && type!=0x37c && type!=0x388)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readPolygon: not a polygon zone\n"));
		return false;
	}
	auto parent=m_state->m_actualGraph;
	m_state->m_actualGraph.reset();
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	auto shape=std::make_shared<QuattroGraphInternal::Shape>();
	if (sz<57 || !readShapeHeader(*shape, stream, endPos-6) || input->tell()+6>endPos)
	{
		if (sz)
		{
			WPS_DEBUG_MSG(("QuattroGraph::readPolygon: the size seems very bad\n"));
			f << *shape << ",";
			f << "###";
		}
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << *shape << ",";
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	pos=input->tell();
	f.str("");
	f << "GrPolygon:";
	int N=int(libwps::readU16(input));
	if ((endPos-pos-2)/4!=N+1)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readPolygon: the number of points seems very bad\n"));
		f << "###N=" << N << ",";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	if (type==0x35c || type==0x37c) // FIXME
		shape->m_shape=WPSGraphicShape::polyline(WPSBox2f(shape->m_box));
	else
		shape->m_shape=WPSGraphicShape::polygon(WPSBox2f(shape->m_box));
	f << "pts=[";
	for (int i=0; i<=N; ++i)
	{
		int pt[2];
		for (auto &p : pt) p=int(libwps::read16(input));
		shape->m_shape.m_vertices.push_back(Vec2f(float(pt[0]),float(pt[1])));
		f << shape->m_shape.m_vertices.back() << ",";
	}
	f << "],";
	if (parent && parent->m_type==parent->Shape)
		parent->m_shape=shape;
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroGraph::readTextBox(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type!=0x36f)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readTextBox: not a text box zone\n"));
		return false;
	}
	auto parent=m_state->m_actualGraph;
	m_state->m_actualGraph.reset();
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	auto textbox=std::make_shared<QuattroGraphInternal::Textbox>();
	if (sz<57 || !readShapeHeader(*textbox, stream, endPos-3) || input->tell()+3>endPos)
	{
		if (sz)
		{
			WPS_DEBUG_MSG(("QuattroGraph::readTextBox: the size seems very bad\n"));
			f << *textbox << ",";
			f << "###";
		}
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << *textbox << ",";
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	pos=input->tell();
	f.str("");
	f << "GrTextBox[text]:";
	int val=int(libwps::readU8(input));
	if (val) f << "f0=" << std::hex << val << std::dec << ",";
	int N=int(libwps::readU16(input));
	if (pos+3+N+10>endPos)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readTextBox: can not read the text zone\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	textbox->m_entry.setBegin(pos+3);
	textbox->m_entry.setLength(N);
	std::string text;
	for (int i=0; i<N; ++i) text+=char(libwps::readU8(input));
	f << text;
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	pos=input->tell();
	f.str("");
	f << "GrTextBox[format]:";
	if (pos+80>endPos)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readTextBox: can not read the format zone\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto &font = textbox->m_font;;
	font.m_size=double(libwps::readU16(input));
	int flags=int(libwps::readU16(input));
	uint32_t attributes = 0;
	if (flags & 1) attributes |= WPS_BOLD_BIT;
	if (flags & 2) attributes |= WPS_ITALICS_BIT;
	if (flags & 4) attributes |= WPS_UNDERLINE_BIT;
	if (flags & 0x20) attributes |= WPS_STRIKEOUT_BIT;
	font.m_attributes=attributes;
	flags&=0xffd8;
	if (flags)
		f << "#font[flag]=" << std::hex << flags << std::dec << ",";
	librevenge::RVNGString name;
	if (!m_mainParser.readCString(stream, name, 32))
		f << "###name,";
	else
		font.m_name=name;
	input->seek(pos+35, librevenge::RVNG_SEEK_SET);
	val=int(libwps::readU8(input)); // 0|5|8|58|ff
	if (val) f << "f0=" << std::hex << val << std::dec << ",";
	unsigned char col[4];
	WPSColor colors[2];
	for (auto &color : colors)
	{
		for (auto &c : col) c=libwps::readU8(input);
		color=WPSColor(col[0],col[1],col[2]);
	}
	int fillType=int(libwps::readU16(input));
	if (fillType==0)
	{
		font.m_color=colors[0];
	}
	else if (fillType>=1 && fillType<=6)
		font.m_color=WPSColor::barycenter(0.5f, colors[0], 0.5f, colors[1]);
	else if ((fillType&0x8060)==0x8060)
	{
		font.m_color=colors[0];
		f << "#3d[effect]=" << (fillType&0x7f9f) << ",";
		WPS_DEBUG_MSG(("QuattroGraph::readTextBox: find unimplemented 3d color\n"));
	}
	else
	{
		WPS_DEBUG_MSG(("QuattroGraph::readTextBox: find unimplemented fillType color\n"));
		f << "###fill[type]=" << std::hex << fillType << std::dec << ",";
	}
	f << font;
	val=int(libwps::read16(input));
	if (val) f << "f1=" << val << ",";
	val=int(libwps::read16(input));
	if (val) f << "box[width]=" << val << ","; // maybe type
	val=int(libwps::read16(input));
	switch (val)
	{
	case 0: // left
		break;
	case 1:
		textbox->m_paragraph.m_justify=libwps::JustificationCenter;
		f << "center,";
		break;
	case 2:
		textbox->m_paragraph.m_justify=libwps::JustificationRight;
		f << "right,";
		break;
	default:
		WPS_DEBUG_MSG(("QuattroGraph::readTextBox: unknown alignment\n"));
		f << "###align=" << val << ",";
		break;
	}
	for (int i=0; i<4; ++i) // f3=1, f2=small number, f3=-1,f4=-1
	{
		val=int(libwps::read16(input));
		if (!val) continue;
		f << "f" << i+2 << "=" << val << ",";
	}
	val=int(libwps::read16(input));
	if (val!=300) f << "tab[width]=" << float(val)/10.f << ","; // unsure about the unit 300=1inch, 390=1,3125 inch
	val=int(libwps::read16(input));
	if (val) f << "g1=" << val << ",";
	int sSz=int(libwps::readU16(input));
	if (input->tell()+sSz+14>endPos)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readTextBox: can not read the last string\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	text="";
	for (int i=0; i<sSz; ++i) text+=char(libwps::readU8(input));
	if (!text.empty()) f << text << ",";
	for (int i=0; i<7; ++i)   // h0=0|2|1d,h3=0|2, h5=0|16a, h6=0|922c
	{
		val=int(libwps::read16(input));
		if (val) f << "h" << i << "=" << val << ",";
	}
	long actPos=input->tell();
	if (fillType && !readFillData(textbox->m_style, fillType, stream, endPos))
		input->seek(actPos, librevenge::RVNG_SEEK_SET);
	if (input->tell()!=endPos)
	{
		ascFile.addDelimiter(input->tell(),'|');
		f << "##extra,";
		WPS_DEBUG_MSG(("QuattroGraph::readTextBox: find extra data\n"));
	}
	if (parent && parent->m_type==parent->Shape)
	{
		parent->m_type=parent->Textbox;
		parent->m_textbox=textbox;
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}
////////////////////////////////////////////////////////////
// dialog
////////////////////////////////////////////////////////////
bool QuattroGraph::readHeader(QuattroGraphInternal::Dialog &header, std::shared_ptr<WPSStream> stream, long endPos)
{
	auto input=stream->m_input;
	long pos = input->tell();
	if (endPos-pos<22)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readHeader: the zone is too short\n"));
		return false;
	}
	for (auto &fl : header.m_flags1) fl=int(libwps::readU16(input));
	int dim[4];
	for (auto &d: dim) d=libwps::readU16(input);
	header.m_cellBox=WPSBox2i(Vec2i(dim[0],dim[1]),Vec2i(dim[2],dim[3]));
	for (auto &fl : header.m_flags2) fl=int(libwps::readU8(input));
	return true;
}

bool QuattroGraph::readDialog(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type != 0x35e)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readDialog: unknown id\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	QuattroGraphInternal::Dialog dialog;
	if (sz<65 || !readHeader(dialog,stream,endPos))
	{
		if (sz)
		{
			WPS_DEBUG_MSG(("QuattroGraph::readDialog: size seems very bad\n"));
			f << "###";
		}
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << dialog << ",";
	int val;
	for (int i=0; i<3; ++i)   // f1=0|-256, f2=-1
	{
		val=int(libwps::read16(input));
		if (val)
			f << "f" << i << "=" << val <<",";
	}
	for (int i=0; i<3; ++i)   // 0
	{
		val=int(libwps::read16(input));
		if (val)
			f << "f" << i+4 << "=" << val <<",";
	}
	val=int(libwps::readU16(input));
	if (val!=0x100)
		f << "f7=" << std::hex << val << std::dec << ",";
	auto sSz=int(libwps::readU16(input));
	librevenge::RVNGString text;
	if (input->tell()+sSz+7+15>endPos || !m_mainParser.readCString(stream,text,sSz))
	{
		WPS_DEBUG_MSG(("QuattroGraph::readDialog: can not read string1\n"));
		f << "##sSz,";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << text.cstr() << ","; // find Dialog1...
	val=int(libwps::readU8(input));
	if (val!=0x1)
		f << "f9=" << val << ",";
	sSz=int(libwps::readU16(input));
	if (sSz<4 || input->tell()+sSz+15>endPos || (sz>4 && !m_mainParser.readCString(stream,text,sSz-4)))
	{
		WPS_DEBUG_MSG(("QuattroGraph::readDialog: can not read string1\n"));
		f << "##sSz2,";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	if (!text.empty()) // find Enter ...
		f << text.cstr() << ",";
	for (int i=0; i<2; ++i)   // 0
	{
		val=int(libwps::read16(input));
		if (!val) continue;
		f << "g" << i << "=" << val << ",";
	}
	if (input->tell()!=endPos) // then some flags
		ascFile.addDelimiter(input->tell(),'|');
	static bool first=true;
	if (first)
	{
		first=false;
		WPS_DEBUG_MSG(("QuattroGraph::readDialog: this file contains some dialogs, there will not be recovered\n"));
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroGraph::readDialogUnknown(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type<0x330 || type>0x380)
	{
		WPS_DEBUG_MSG(("QuattroGraph::readDialogUnknown: unknown id\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	QuattroGraphInternal::Dialog dialog;
	if (sz<38 || !readHeader(dialog,stream,endPos))
	{
		if (sz)
		{
			WPS_DEBUG_MSG(("QuattroGraph::readDialogUnknown: size seems very bad\n"));
			f << "###";
		}
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	if (dialog.m_flags2[8]&0x80)
	{
		f << "select,";
		dialog.m_flags2[8]&=0x7f;
	}
	f << dialog << ",";
	auto fl=int(libwps::readU8(input)); // [0128][08ac]
	if (fl&1) f << "has[frame],";
	fl &= 0xfe;
	if (fl) f << "flag=" << std::hex << fl << std::dec << ",";
	auto id=int(libwps::readU16(input));
	f << "id=" << id << ",";
	unsigned char col[3];
	for (auto &c : col) c=libwps::readU8(input);
	f << "col=" << WPSColor(col[0],col[1],col[2]) << ",";
	f << "fl3=[";
	for (int i=0; i<5; ++i)   // 0
	{
		auto val=int(libwps::readU8(input));
		if (val)
			f << std::hex << val << std::dec << ",";
		else
			f << "_,";
	}
	f << "],";
	if (input->tell()!=endPos) // then some flags
		ascFile.addDelimiter(input->tell(),'|');
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

////////////////////////////////////////////////////////////
// send data
////////////////////////////////////////////////////////////
bool QuattroGraph::sendGraphics(int sheetId, Vec2i const &cell) const
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("QuattroGraph::sendGraphics: can not find the listener\n"));
		return false;
	}
	auto it=m_state->m_sheetIdToGraphMap.find(sheetId);
	bool find=false;
	while (it!=m_state->m_sheetIdToGraphMap.end() && it->first==sheetId)
	{
		auto &graph=it++->second;
		if (!graph || graph->m_cellBox[0]!=cell) continue;
		sendGraphic(*graph);
		find=true;
	}
	if (!find)
	{
		WPS_DEBUG_MSG(("QuattroGraph::sendGraphics: sorry, can not find any graph\n"));
	}
	return find;
}

bool QuattroGraph::sendPageGraphics(int sheetId) const
{
	for (auto it=m_state->m_sheetIdToGraphMap.lower_bound(sheetId); it!=m_state->m_sheetIdToGraphMap.upper_bound(sheetId); ++it)
	{
		auto &graph=it->second;
		if (!graph) continue;
		if (graph->m_type==graph->Shape)
			sendShape(*graph, sheetId);
		if (graph->m_type==graph->Textbox)
			sendTextbox(*graph, sheetId);
	}
	return true;
}

bool QuattroGraph::sendGraphic(QuattroGraphInternal::Graph const &graph) const
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("QuattroGraph::sendGraphic: can not find the listener\n"));
		return false;
	}
	WPSPosition pos(graph.m_cellBoxDecal[0], graph.m_size, librevenge::RVNG_POINT);
	pos.m_anchorTo = WPSPosition::Cell;
	pos.m_anchorCellName = libwps::getCellName(graph.m_cellBox[1]+Vec2i(1,1)).c_str();
	if (graph.m_type==graph.OLE)
	{
		if (!graph.m_linkName.empty())
		{
			auto it=m_state->m_linkNameToObjectMap.find(graph.m_linkName);
			if (it==m_state->m_linkNameToObjectMap.end() || it->second.isEmpty())
			{
				WPS_DEBUG_MSG(("QuattroGraph::sendGraphic: can not find ole %s\n", graph.m_linkName.cstr()));
			}
			else
				m_listener->insertObject(pos, it->second);
		}
		else if (graph.m_ole.isEmpty())
		{
			WPS_DEBUG_MSG(("QuattroGraph::sendGraphic: find an empty ole\n"));
		}
		else
			m_listener->insertObject(pos, graph.m_ole);
		return true;
	}
	if (graph.m_type==graph.Button)
	{
		if (graph.m_label.empty())
		{
			WPS_DEBUG_MSG(("QuattroGraph::sendGraphic: find an empty label\n"));
		}
		else
		{
			WPSGraphicStyle style;
			style.setBackgroundColor(WPSColor(128,128,128));
			auto doc=std::make_shared<QuattroGraphInternal::SubDocument>(const_cast<QuattroGraph &>(*this),graph.m_label);
			m_listener->insertTextBox(pos, doc, style);
		}
		return true;
	}
	static bool first=true;
	if (first)
	{
		first=false;
		WPS_DEBUG_MSG(("QuattroGraph::sendGraphic: sorry, unexpected graph type\n"));
	}
	return true;
}

bool QuattroGraph::sendShape(QuattroGraphInternal::Graph const &graph, int sheetId) const
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("QuattroGraph::sendShape: can not find the listener\n"));
		return false;
	}
	if (graph.m_type!=graph.Shape || !graph.m_shape)
	{
		WPS_DEBUG_MSG(("QuattroGraph::sendShape: can not find the shape\n"));
		return false;
	}
	auto const &shape=*graph.m_shape;
	WPSPosition pos(graph.m_cellBoxDecal[0]+m_mainParser.getCellPosition(sheetId, graph.m_cellBox[0]), graph.m_size, librevenge::RVNG_POINT);
	pos.m_anchorTo = WPSPosition::Page;
	auto gShape=shape.m_shape;
	auto bdBoxSize=gShape.getBdBox().size();
	gShape.scale(Vec2f(bdBoxSize[0]>0 ? float(graph.m_size[0])/bdBoxSize[0] : 1.f, bdBoxSize[1]>0 ? float(graph.m_size[1])/bdBoxSize[1] : 1.f));
	m_listener->insertPicture(pos, gShape, shape.m_style);
	return true;
}

bool QuattroGraph::sendTextbox(QuattroGraphInternal::Graph const &graph, int sheetId) const
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("QuattroGraph::sendTextbox: can not find the listener\n"));
		return false;
	}
	if (graph.m_type!=graph.Textbox || !graph.m_textbox)
	{
		WPS_DEBUG_MSG(("QuattroGraph::sendTextbox: can not find the textbox\n"));
		return false;
	}
	auto const &textbox=*graph.m_textbox;
	WPSPosition pos(graph.m_cellBoxDecal[0]+m_mainParser.getCellPosition(sheetId, graph.m_cellBox[0]), graph.m_size, librevenge::RVNG_POINT);
	pos.m_anchorTo = WPSPosition::Page;
	auto doc=std::make_shared<QuattroGraphInternal::SubDocument>(const_cast<QuattroGraph &>(*this),graph.m_textbox, graph.m_stream);
	m_listener->insertTextBox(pos, doc, textbox.m_style);
	return true;
}

bool QuattroGraph::send(QuattroGraphInternal::Textbox const &textbox, std::shared_ptr<WPSStream> stream) const
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("QuattroGraph::send: can not find the listener\n"));
		return false;
	}
	if (!stream || !textbox.m_entry.valid())
	{
		WPS_DEBUG_MSG(("QuattroGraph::send: can not find the file stream\n"));
		return false;
	}
	m_listener->setFont(textbox.m_font);
	m_listener->setParagraph(textbox.m_paragraph);
	auto input=stream->m_input;
	input->seek(textbox.m_entry.begin(), librevenge::RVNG_SEEK_SET);
	auto fontType = m_mainParser.getDefaultFontType();
	std::string text;
	for (long l=0; l<=textbox.m_entry.length(); ++l)
	{
		auto c=l==textbox.m_entry.length() ? '\0' : char(libwps::readU8(input));
		if ((c==0 || c==0x9 || c==0xd) && !text.empty())
		{
			m_listener->insertUnicodeString(libwps_tools_win::Font::unicodeString(text, fontType));
			text.clear();
		}
		if (l==textbox.m_entry.length()) break;
		if (c==0xd)
			m_listener->insertEOL();
		else if (c==0x9)
			m_listener->insertTab();
		else if (c)
			text.push_back(c);
	}
	return true;
}

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
