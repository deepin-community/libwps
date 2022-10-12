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
#include "WPSParagraph.h"
#include "WPSPosition.h"
#include "WPSStream.h"

#include "Lotus.h"
#include "LotusStyleManager.h"

#include "LotusGraph.h"

namespace LotusGraphInternal
{
//! the graphic zone of a LotusGraph for 123 mac
struct ZoneMac
{
	//! the different type
	enum Type { Arc, Frame, Line, Rect, Unknown };
	//! constructor
	explicit ZoneMac(std::shared_ptr<WPSStream> const &stream)
		: m_type(Unknown)
		, m_subType(0)
		, m_stream(stream)
		, m_box()
		, m_ordering(0)
		, m_lineId(0)
		, m_graphicId(0)
		, m_surfaceId(0)
		, m_hasShadow(false)
		, m_chartId(0)
		, m_pictureEntry()
		, m_textBoxEntry()
		, m_extra("")
	{
		for (int &value : m_values) value=0;
	}
	//! returns a graphic shape corresponding to the main form (and the origin)
	bool getGraphicShape(WPSGraphicShape &shape, WPSPosition &pos) const;
	//! operator<<
	friend std::ostream &operator<<(std::ostream &o, ZoneMac const &z)
	{
		switch (z.m_type)
		{
		case Arc:
			o << "arc,";
			break;
		case Frame:
			// subType is a small number 2, 3 ?
			o << "frame[" << z.m_subType << "],";
			break;
		case Line:
			o << "line,";
			break;
		case Rect:
			if (z.m_subType==1)
				o << "rect,";
			else if (z.m_subType==2)
				o << "rectOval,";
			else if (z.m_subType==3)
				o << "oval,";
			else
				o << "rect[#" << z.m_subType << "],";
			break;
		case Unknown:
		default:
			break;
		}
		o << z.m_box << ",";
		o << "order=" << z.m_ordering << ",";
		if (z.m_lineId)
			o << "L" << z.m_lineId << ",";
		if (z.m_surfaceId)
			o << "Co" << z.m_surfaceId << ",";
		if (z.m_graphicId)
			o << "G" << z.m_graphicId << ",";
		if (z.m_hasShadow)
			o << "shadow,";
		for (int i=0; i<4; ++i)
		{
			if (z.m_values[i])
				o << "val" << i << "=" << z.m_values[i] << ",";
		}
		o << z.m_extra << ",";
		return o;
	}
	//! the zone type
	Type m_type;
	//! the file modifier type
	int m_subType;
	//! the stream
	std::shared_ptr<WPSStream> m_stream;
	//! the bdbox
	WPSBox2i m_box;
	//! the ordering
	int m_ordering;
	//! the line style id
	int m_lineId;
	//! the graphic style id, used with rect shadow
	int m_graphicId;
	//! the surface color style id
	int m_surfaceId;
	//! a flag to know if we need to add shadow
	bool m_hasShadow;
	//! the chart id(for chart)
	int m_chartId;
	//! the picture entry
	WPSEntry m_pictureEntry;
	//! the text box entry
	WPSEntry m_textBoxEntry;
	//! unknown other value
	int m_values[4];
	//! extra data
	std::string m_extra;
};

bool ZoneMac::getGraphicShape(WPSGraphicShape &shape, WPSPosition &pos) const
{
	pos=WPSPosition(Vec2f(m_box[0]),Vec2f(m_box.size()), librevenge::RVNG_POINT);
	pos.setRelativePosition(WPSPosition::Page);
	WPSBox2f box(Vec2f(0,0), Vec2f(m_box.size()));
	switch (m_type)
	{
	case Line:
	{
		// we need to recompute the bdbox
		int bounds[4]= {m_box[0][0],m_box[0][1],m_box[1][0],m_box[1][1]};
		for (int i=0; i<2; ++i)
		{
			if (bounds[i]<=bounds[i+2]) continue;
			bounds[i]=bounds[i+2];
			bounds[i+2]=m_box[0][i];
		}
		WPSBox2i realBox(Vec2i(bounds[0],bounds[1]),Vec2i(bounds[2],bounds[3]));
		pos=WPSPosition(Vec2f(realBox[0]),Vec2f(realBox.size()), librevenge::RVNG_POINT);
		pos.setRelativePosition(WPSPosition::Page);
		shape=WPSGraphicShape::line(Vec2f(m_box[0]-realBox[0]), Vec2f(m_box[1]-realBox[0]));
		return true;
	}
	case Rect:
		switch (m_subType)
		{
		case 2:
			shape=WPSGraphicShape::rectangle(box, Vec2f(5,5));
			return true;
		case 3:
			shape=WPSGraphicShape::circle(box);
			return true;
		default:
		case 1:
			shape=WPSGraphicShape::rectangle(box);
			return true;
		}
	case Frame:
		shape=WPSGraphicShape::rectangle(box);
		return true;
	case Arc:
		// changeme if the shape box if defined with different angle
		shape=WPSGraphicShape::arc(box, WPSBox2f(Vec2f(-box[1][0],0),Vec2f(box[1][0],2*box[1][1])), Vec2f(0,90));
		return true;
	case Unknown:
	default:
		break;
	}
	return false;
}

//! the graphic zone of a LotusGraph : wk4
struct ZoneWK4
{
	//! the different type
	enum Type { Border, Chart, Group, Picture, Shape, TextBox, Unknown };
	//! constructor
	explicit ZoneWK4(std::shared_ptr<WPSStream> const &stream)
		: m_type(Unknown)
		, m_subType(-1)
		, m_id(-1)
		, m_cell(0,0)
		, m_cellPosition(0,0)
		, m_frameSize(0,0)
		, m_pictureDim()
		, m_pictureName()
		, m_shape()
		, m_graphicStyle(WPSGraphicStyle::emptyStyle())
		, m_hasShadow(false)
		, m_textEntry()
		, m_stream(stream)
	{
	}
	//! the zone type
	int m_type;
	//! the sub type
	int m_subType;
	//! the zone id
	int m_id;
	//! the cell
	Vec2i m_cell;
	//! the position in the cell
	Vec2f m_cellPosition;
	//! the frame dimension
	Vec2i m_frameSize;
	//! the picture dimension
	WPSBox2i m_pictureDim;
	//! the picture name
	std::string m_pictureName;
	//! the graphic shape
	WPSGraphicShape m_shape;
	//! the graphic style wk4
	WPSGraphicStyle m_graphicStyle;
	//! a flag to know if we need to add shadow
	bool m_hasShadow;
	//! the text entry (for textbox and button)
	WPSEntry m_textEntry;
	//! the stream
	std::shared_ptr<WPSStream> m_stream;
};

//! the graphic zone of a LotusGraph for 123 pc
struct ZonePc
{
	//! the different type
	enum Type { Arc, Chart, Ellipse, FreeHand, Line, Picture, Polygon, Rect, Set, TextBox, Unknown };
	//! constructor
	explicit ZonePc(std::shared_ptr<WPSStream> const &stream)
		: m_type(Unknown)
		, m_isGroup(false)
		, m_groupLastPosition(0)
		, m_numPoints(0)
		, m_vertices()
		, m_isRoundRect(false)
		, m_stream(stream)
		, m_box()
		, m_translate(0,0)
		, m_rotate(0)
		, m_arrows(0)
		, m_textBoxEntry()
		, m_pictureData()
		, m_pictureHeaderRead(0)
		, m_isSent(false)
		, m_extra("")
	{
		for (int &i : m_graphicId) i=-1;
	}
	//! returns a graphic shape corresponding to the main form (and the origin)
	bool getGraphicShape(WPSGraphicShape &shape, WPSPosition &pos) const;
	//! returns a transformation corresponding to the shape
	WPSTransformation getTransformation() const
	{
		WPSTransformation res;
		if (m_rotate<0||m_rotate>0)
			res = WPSTransformation::rotation(-m_rotate, m_box.center());
		if (m_translate!=Vec2f(0,0))
			res = WPSTransformation::translation(m_translate) * res;
		return res;
	}
	//! operator<<
	friend std::ostream &operator<<(std::ostream &o, ZonePc const &z)
	{
		switch (z.m_type)
		{
		case Arc:
			o << "arc,";
			break;
		case Chart:
			o << "chart,";
			break;
		case Ellipse:
			o << "ellipse,";
			break;
		case FreeHand:
			o << "freeHand,";
			if (z.m_numPoints) o << "N=" << z.m_numPoints << ",";
			break;
		case Line:
			o << "line,";
			if (z.m_numPoints!=2) o << "N=" << z.m_numPoints << ",";
			break;
		case Picture:
			o << "picture,";
			break;
		case Polygon:
			o << "polygon,";
			if (z.m_numPoints) o << "N=" << z.m_numPoints << ",";
			break;
		case Rect:
			if (z.m_isRoundRect)
				o << "rect[round],";
			else
				o << "rect,";
			break;
		case Set:
			if (z.m_isGroup)
				o << "group,";
			else
				o << "set,";
			break;
		case TextBox:
			o << "textBox,";
			break;
		case Unknown:
		default:
			break;
		}
		o << "dim=" << z.m_box << ",";
		if (z.m_translate!=Vec2f(0,0))
			o << "translation=" << z.m_translate << ",";
		if (z.m_rotate<0||z.m_rotate>0)
			o << "rotation=" << z.m_rotate << ",";
		for (int i=0; i<2; ++i)
		{
			if (z.m_graphicId[i]<0) continue;
			o << (i==0 ? "style" : "shadow") << "=GS" << z.m_graphicId[i] << ",";
		}
		if (z.m_arrows&1)
			o << "arrows[beg],";
		if (z.m_arrows&2)
			o << "arrows[end],";
		o << z.m_extra << ",";
		return o;
	}
	//! the zone type
	Type m_type;
	//! true if the set is a group
	bool m_isGroup;
	//! the group last position
	size_t m_groupLastPosition;
	//! the number of points of a line or a polygon
	int m_numPoints;
	//! the list of points for a polygon
	std::vector<Vec2f> m_vertices;
	//! true if the rect is a round rect
	bool m_isRoundRect;
	//! the stream
	std::shared_ptr<WPSStream> m_stream;
	//! the bdbox
	WPSBox2f m_box;
	//! the translation
	Vec2f m_translate;
	//! the rotation
	float m_rotate;
	//! the graphic style id and the shadow style
	int m_graphicId[2];
	//! the line arrows
	int m_arrows;
	//! the text box entry
	WPSEntry m_textBoxEntry;
	//! the picture data
	librevenge::RVNGBinaryData m_pictureData;
	//! the number of read byte of the header
	int m_pictureHeaderRead;
	//! a flag to know if the zone has already be sent
	mutable bool m_isSent;
	//! extra data
	std::string m_extra;
};

bool ZonePc::getGraphicShape(WPSGraphicShape &shape, WPSPosition &pos) const
{
	pos=WPSPosition(m_box[0],m_box.size(), librevenge::RVNG_POINT);
	pos.setRelativePosition(WPSPosition::Page);
	WPSBox2f box(Vec2f(0,0), m_box.size());
	switch (m_type)
	{
	case Line:
	{
		// we need to recompute the bdbox
		float bounds[4]= {m_box[0][0],m_box[0][1],m_box[1][0],m_box[1][1]};
		for (int i=0; i<2; ++i)
		{
			if (bounds[i]<=bounds[i+2]) continue;
			bounds[i]=bounds[i+2];
			bounds[i+2]=m_box[0][i];
		}
		WPSBox2f realBox(Vec2f(bounds[0],bounds[1]),Vec2f(bounds[2],bounds[3]));
		pos=WPSPosition(realBox[0],realBox.size(), librevenge::RVNG_POINT);
		pos.setRelativePosition(WPSPosition::Page);
		shape=WPSGraphicShape::line(m_box[0]-realBox[0], m_box[1]-realBox[0]);
		return true;
	}
	case Ellipse:
		shape=WPSGraphicShape::circle(box);
		return true;
	case Rect:
		if (m_isRoundRect)
			shape=WPSGraphicShape::rectangle(box, Vec2f(5,5));
		else
			shape=WPSGraphicShape::rectangle(box);
		return true;
	case Arc: // checkme only work if no flip
		shape=WPSGraphicShape::arc(box, WPSBox2f(Vec2f(-box[1][0],0),Vec2f(box[1][0],2*box[1][1])), Vec2f(0,90));
		return true;
	case FreeHand:
	case Polygon:
		if (m_vertices.empty())
		{
			WPS_DEBUG_MSG(("LotusGraphInternal::ZonePc::getGraphicShape: sorry, can not find the polygon vertices\n"));
			return false;
		}
		shape=WPSGraphicShape::polygon(box);
		shape.m_vertices=m_vertices;
		shape.m_vertices.push_back(m_vertices[0]);
		return true;
	case Chart:
	case Set:
	case TextBox:
		return true;
	case Picture:
	case Unknown:
	default:
		break;
	}
	static bool first=true;
	if (first)
	{
		first=false;
		WPS_DEBUG_MSG(("LotusGraphInternal::ZonePc::getGraphicShape: sorry, sending some graph types is not implemented\n"));
	}
	return false;
}
//! a list of ZonePc of a LotusGraph for 123 pc
struct ZonePcList
{
	//! constructor
	ZonePcList()
		: m_zones()
		, m_groupBeginStack()
	{
	}
	//! returns true if the number of graphic zone is empty
	bool empty() const
	{
		for (auto const &zone : m_zones)
		{
			if (zone)
				return false;
		}
		return true;
	}
	//! the list of zones
	std::vector<std::shared_ptr<ZonePc> > m_zones;
	//! the stack indicating the beginning of each group
	std::stack<size_t> m_groupBeginStack;
};

//! the state of LotusGraph
struct State
{
	//! constructor
	State()
		:  m_version(-1)
		, m_actualSheetId(-1)
		, m_sheetIdZoneMacMap()
		, m_actualZoneMac()
		, m_sheetIdZoneWK4Map()
		, m_actualZoneWK4()
		, m_sheetIdZonePcListMap()
		, m_actualZonePc()
		, m_zIdToSheetIdMap()
		, m_nameToChartIdMap()
	{
	}

	//! the file version
	int m_version;
	//! the actual sheet id
	int m_actualSheetId;
	//! a map sheetid to zone
	std::multimap<int, std::shared_ptr<ZoneMac> > m_sheetIdZoneMacMap;
	//! a pointer to the actual zone
	std::shared_ptr<ZoneMac> m_actualZoneMac;
	//! a map sheetid to zone
	std::multimap<int, std::shared_ptr<ZoneWK4> > m_sheetIdZoneWK4Map;
	//! a pointer to the actual zone
	std::shared_ptr<ZoneWK4> m_actualZoneWK4;
	//! a map sheetid to zone
	std::map<int, ZonePcList> m_sheetIdZonePcListMap;
	//! a pointer to the actual zone
	std::shared_ptr<ZonePc> m_actualZonePc;
	//! a map sheet zone id to final sheet id map
	std::map<int,int> m_zIdToSheetIdMap;
	//! a map chart name to id chart map
	std::map<std::string,int> m_nameToChartIdMap;
};

//! Internal: the subdocument of a LotusGraphc
class SubDocument final : public WKSSubDocument
{
public:
	//! constructor for a text entry
	SubDocument(std::shared_ptr<WPSStream> const &stream, LotusGraph &graphParser, WPSEntry &entry, int version)
		: WKSSubDocument(RVNGInputStreamPtr(), &graphParser.m_mainParser)
		, m_stream(stream)
		, m_graphParser(graphParser)
		, m_entry(entry)
		, m_version(version) {}
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
		if (m_stream.get() != sDoc->m_stream.get()) return false;
		if (m_version != sDoc->m_version) return false;
		return m_entry == sDoc->m_entry;
	}

	//! the parser function
	void parse(std::shared_ptr<WKSContentListener> &listener, libwps::SubDocumentType subDocumentType) final;
	//! the stream
	std::shared_ptr<WPSStream> m_stream;
	//! the graph parser
	LotusGraph &m_graphParser;
	//! a flag to known if we need to send the entry or the footer
	WPSEntry m_entry;
	//! the textbox version
	int m_version;
};

void SubDocument::parse(std::shared_ptr<WKSContentListener> &listener, libwps::SubDocumentType)
{
	if (!listener.get())
	{
		WPS_DEBUG_MSG(("LotusGraphInternal::SubDocument::parse: no listener\n"));
		return;
	}
	if (!dynamic_cast<WKSContentListener *>(listener.get()))
	{
		WPS_DEBUG_MSG(("LotusGraphInternal::SubDocument::parse: bad listener\n"));
		return;
	}
	if (m_version==0)
		m_graphParser.sendTextBox(m_stream, m_entry);
	else if (m_version==1 || m_version==2)
		m_graphParser.sendTextBoxWK4(m_stream, m_entry, m_version==2);
	else
	{
		WPS_DEBUG_MSG(("LotusGraphInternal::SubDocument::parse: unknown version=%d\n", m_version));
		return;
	}
}

}

// constructor, destructor
LotusGraph::LotusGraph(LotusParser &parser)
	: m_listener()
	, m_mainParser(parser)
	, m_styleManager(parser.m_styleManager)
	, m_state(new LotusGraphInternal::State)
{
}

LotusGraph::~LotusGraph()
{
}

void LotusGraph::cleanState()
{
	m_state.reset(new LotusGraphInternal::State);
}

void LotusGraph::updateState(std::map<int,int> const &zIdToSheetIdMap,
                             std::map<std::string,int> const &nameToChartIdMap)
{
	m_state->m_zIdToSheetIdMap=zIdToSheetIdMap;
	m_state->m_nameToChartIdMap=nameToChartIdMap;
}

int LotusGraph::version() const
{
	if (m_state->m_version<0)
		m_state->m_version=m_mainParser.version();
	return m_state->m_version;
}

bool LotusGraph::hasGraphics(int sheetId) const
{
	if (m_state->m_sheetIdZoneMacMap.find(sheetId)!=m_state->m_sheetIdZoneMacMap.end() ||
	        m_state->m_sheetIdZoneWK4Map.find(sheetId)!=m_state->m_sheetIdZoneWK4Map.end())
		return true;
	if (m_state->m_zIdToSheetIdMap.find(sheetId)!=m_state->m_zIdToSheetIdMap.end())
	{
		int finalId=m_state->m_zIdToSheetIdMap.find(sheetId)->second;
		if (m_state->m_sheetIdZonePcListMap.find(finalId)!=m_state->m_sheetIdZonePcListMap.end()
		        && !m_state->m_sheetIdZonePcListMap.find(finalId)->second.empty())
			return true;
	}
	return false;
}

bool LotusGraph::setChartId(int chartId)
{
	auto zone=m_state->m_actualZoneMac;
	if (!zone || zone->m_type != LotusGraphInternal::ZoneMac::Frame)
	{
		WPS_DEBUG_MSG(("LotusGraph::setChartId: Oops can not find the parent frame\n"));
		return false;
	}
	zone->m_chartId = chartId;
	m_state->m_actualZoneMac.reset();
	return true;
}

////////////////////////////////////////////////////////////
// low level

////////////////////////////////////////////////////////////
// zones
////////////////////////////////////////////////////////////
bool LotusGraph::readZoneBegin(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	f << "Entries(GraphBegin):";
	long pos = input->tell();
	if (endPos-pos!=4)
	{
		WPS_DEBUG_MSG(("LotusParser::readZoneBegin: the zone seems bad\n"));
		f << "###";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());

		return true;
	}
	m_state->m_actualSheetId=int(libwps::readU8(input));
	f << "sheet[id]=" << m_state->m_actualSheetId << ",";
	for (int i=0; i<3; ++i)   // f0=1
	{
		auto val=int(libwps::readU8(input));
		if (val)
			f << "f" << i << "=" << val << ",";
	}
	m_state->m_actualZoneMac.reset();
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	return true;

}

bool LotusGraph::readZoneData(std::shared_ptr<WPSStream> stream, long endPos, int type)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	long sz=endPos-pos;

	std::shared_ptr<LotusGraphInternal::ZoneMac> zone(new LotusGraphInternal::ZoneMac(stream));
	m_state->m_actualZoneMac=zone;

	switch (type)
	{
	case 0x2332:
		zone->m_type=LotusGraphInternal::ZoneMac::Line;
		break;
	case 0x2346:
		zone->m_type=LotusGraphInternal::ZoneMac::Rect;
		break;
	case 0x2350:
		zone->m_type=LotusGraphInternal::ZoneMac::Arc;
		break;
	case 0x2352:
		zone->m_type=LotusGraphInternal::ZoneMac::Frame;
		zone->m_hasShadow=true;
		break;
	case 0x23f0:
		zone->m_type=LotusGraphInternal::ZoneMac::Frame;
		break;
	default:
		WPS_DEBUG_MSG(("LotusGraph::readZoneData: find unexpected graph data\n"));
		f << "###";
	}
	if (sz<24)
	{
		WPS_DEBUG_MSG(("LotusGraph::readZoneData: the zone seems too short\n"));
		f << "Entries(GraphMac):" << *zone << "###";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	zone->m_ordering=int(libwps::readU8(input));
	for (int i=0; i<4; ++i)   // always 0?
	{
		auto val=int(libwps::read8(input));
		if (val)
			f << "f" << i << "=" << val << ",";
	}
	int dim[4];
	for (int i=0; i<4; ++i)   // dim3[high]=0|100
	{
		dim[i]=int(libwps::read16(input));
		if (i==3)
			break;
		auto val=int(libwps::read16(input));
		if (val) f << "dim" << i << "[high]=" << std::hex << val << std::dec << ",";
	}
	zone->m_box=WPSBox2i(Vec2i(dim[1],dim[0]),Vec2i(dim[3],dim[2]));
	auto val=int(libwps::read8(input));
	if (val) // always 0
		f << "f4=" << val << ",";
	int fl;
	switch (zone->m_type)
	{
	case LotusGraphInternal::ZoneMac::Line:
		val=int(libwps::readU8(input));
		fl=int(libwps::readU8(input));
		if (val)
		{
			if (fl!=0x10)
				f << "#line[fl]=" << std::hex << fl << std::dec << ",";
			zone->m_lineId=val;
		}
		val=int(libwps::readU8(input)); // always 1?
		if (val!=1)
			f << "g0=" << val << ",";
		// the arrows &1 means end, &2 means begin
		zone->m_values[0]=int(libwps::readU8(input));
		if (sz<26)
		{
			WPS_DEBUG_MSG(("LotusGraph::readZoneData: the line zone seems too short\n"));
			f << "###sz,";
			break;
		}
		for (int i=0; i<2; ++i)   // always g1=0, g2=3 ?
		{
			val=int(libwps::readU8(input));
			if (val!=3*i)
				f << "g" << i+1 << "=" << val << ",";
		}
		break;
	case LotusGraphInternal::ZoneMac::Rect:
		val=int(libwps::readU8(input)); // always 1?
		if (val!=1)
			f << "g0=" << val << ",";
		zone->m_subType=int(libwps::readU8(input));
		if (sz<28)
		{
			WPS_DEBUG_MSG(("LotusGraph::readZoneData: the rect zone seems too short\n"));
			f << "###sz,";
			break;
		}
		for (int i=0; i<2; ++i)
		{
			val=int(libwps::readU8(input));
			fl=int(libwps::readU8(input));
			if (!val) continue;
			if (i==0)
			{
				if (fl!=0x10)
					f << "#line[fl]=" << std::hex << fl << std::dec << ",";
				zone->m_lineId=val;
				continue;
			}
			if (fl!=0x20)
				f << "#surface[fl]=" << std::hex << fl << std::dec << ",";
			zone->m_surfaceId=val;
		}
		val=int(libwps::read16(input)); // always 3?
		if (val!=3)
			f << "g1=" << val << ",";
		break;
	case LotusGraphInternal::ZoneMac::Frame:
		val=int(libwps::readU8(input)); // always 1?
		if (val!=1)
			f << "g0=" << val << ",";
		// small value 1-4
		zone->m_subType=int(libwps::readU8(input));
		val=int(libwps::readU8(input));
		fl=int(libwps::readU8(input));
		if (!val) break;;
		if (fl!=0x40)
			f << "#graphic[fl]=" << std::hex << fl << std::dec << ",";
		zone->m_graphicId=val;
		// can be followed by 000000000100 : some way to determine the content ?
		break;
	case LotusGraphInternal::ZoneMac::Arc:
		val=int(libwps::readU8(input)); // always 1?
		if (val!=1)
			f << "g0=" << val << ",";
		// always 3
		zone->m_subType=int(libwps::readU8(input));
		val=int(libwps::readU8(input));
		fl=int(libwps::readU8(input));
		if (val)
		{
			if (fl!=0x10)
				f << "#line[fl]=" << std::hex << fl << std::dec << ",";
			zone->m_lineId=val;
		}
		if (sz<26)
		{
			WPS_DEBUG_MSG(("LotusGraph::readZoneData: the arc zone seems too short\n"));
			f << "###sz,";
			break;
		}
		val=int(libwps::read16(input)); // always 0? maybe the angle
		if (val)
			f << "g1=" << val << ",";
		break;
	case LotusGraphInternal::ZoneMac::Unknown:
	default:
		break;
	}

	if (m_state->m_actualSheetId<0)
	{
		WPS_DEBUG_MSG(("LotusGraph::readZoneData: oops no sheet zone is opened\n"));
		f << "###sheetId,";
	}
	else
		m_state->m_sheetIdZoneMacMap.insert(std::multimap<int, std::shared_ptr<LotusGraphInternal::ZoneMac> >::value_type
		                                    (m_state->m_actualSheetId, zone));
	zone->m_extra=f.str();
	f.str("");
	f << "Entries(GraphMac):" << *zone;
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusGraph::readTextBoxData(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	long sz=endPos-pos;
	f << "Entries(GraphTextBox):";
	if (sz<1)
	{
		WPS_DEBUG_MSG(("LotusGraph::readTextBoxData: Oops the zone seems too short\n"));
		f << "###";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}

	if (!m_state->m_actualZoneMac || m_state->m_actualZoneMac->m_type != LotusGraphInternal::ZoneMac::Frame)
	{
		WPS_DEBUG_MSG(("LotusGraph::readTextBoxData: Oops can not find the parent frame\n"));
	}
	else
	{
		m_state->m_actualZoneMac->m_textBoxEntry.setBegin(input->tell());
		m_state->m_actualZoneMac->m_textBoxEntry.setEnd(endPos);
		m_state->m_actualZoneMac.reset();
	}

	m_state->m_actualZoneMac.reset();
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	input->seek(endPos, librevenge::RVNG_SEEK_SET);
	return true;
}

bool LotusGraph::readPictureDefinition(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	long sz=endPos-pos;

	f << "Entries(PictDef):";
	if (sz!=13)
	{
		WPS_DEBUG_MSG(("LotusGraph::readPictureDefinition: the picture def seems bad\n"));
		f << "###";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	if (!m_state->m_actualZoneMac || m_state->m_actualZoneMac->m_type != LotusGraphInternal::ZoneMac::Frame)
	{
		WPS_DEBUG_MSG(("LotusGraph::readPictureDefinition: Oops can not find the parent frame\n"));
	}
	auto val=int(libwps::readU8(input)); // always 0?
	if (val)
		f << "f0=" << val << ",";
	int dim[2];
	dim[0]=int(libwps::readU16(input));
	for (int i=0; i<2; ++i)
	{
		val=int(libwps::readU8(input));
		if (val)
			f << "f" << i+1 << "=" << val << ",";
	}
	dim[1]=int(libwps::readU16(input));
	f << "dim=" << Vec2i(dim[0], dim[1]) << ",";
	val=int(libwps::readU8(input));
	if (val)
		f << "f3=" << val << ",";
	auto pictSz=int(libwps::readU16(input)); // maybe 32bits
	f << "pict[sz]=" << std::hex << pictSz << std::dec << ",";
	for (int i=0; i<3; ++i)   // always 0,0,1
	{
		val=int(libwps::readU8(input));
		if (val)
			f << "g" << i << "=" << val << ",";
	}
	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusGraph::readPictureData(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	long sz=endPos-pos;

	f << "Entries(PictData):";
	if (sz<=1)
	{
		WPS_DEBUG_MSG(("LotusGraph::readPictureData: the picture def seems bad\n"));
		f << "###";
		ascFile.addPos(pos-6);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto val=int(libwps::readU8(input)); // always 1?
	if (val!=1)
		f << "type?=" << val << ",";
	if (!m_state->m_actualZoneMac || m_state->m_actualZoneMac->m_type != LotusGraphInternal::ZoneMac::Frame)
	{
		WPS_DEBUG_MSG(("LotusGraph::readPictureData: Oops can not find the parent frame\n"));
	}
	else
	{
		m_state->m_actualZoneMac->m_pictureEntry.setBegin(input->tell());
		m_state->m_actualZoneMac->m_pictureEntry.setEnd(endPos);
		m_state->m_actualZoneMac.reset();
	}
#ifdef DEBUG_WITH_FILES
	ascFile.skipZone(input->tell(), endPos-1);
	librevenge::RVNGBinaryData data;
	if (!libwps::readData(input, static_cast<unsigned long>(endPos-input->tell()), data))
		f << "###";
	else
	{
		std::stringstream s;
		static int fileId=0;
		s << "Pict" << ++fileId << ".pct";
		libwps::Debug::dumpFile(data, s.str().c_str());
	}
#endif

	ascFile.addPos(pos-6);
	ascFile.addNote(f.str().c_str());
	return true;
}

////////////////////////////////////////////////////////////
// send data
////////////////////////////////////////////////////////////
void LotusGraph::sendPicture(LotusGraphInternal::ZoneMac const &zone)
{
	if (!m_listener || !zone.m_stream || !zone.m_stream->m_input || !zone.m_pictureEntry.valid())
	{
		WPS_DEBUG_MSG(("LotusGraph::sendPicture: I can not find the listener/picture entry\n"));
		return;
	}
	RVNGInputStreamPtr input=zone.m_stream->m_input;
	librevenge::RVNGBinaryData data;
	input->seek(zone.m_pictureEntry.begin(), librevenge::RVNG_SEEK_SET);
	if (!libwps::readData(input, static_cast<unsigned long>(zone.m_pictureEntry.length()), data))
	{
		WPS_DEBUG_MSG(("LotusGraph::sendPicture: I can not find the picture\n"));
		return;
	}
	WPSGraphicShape shape;
	WPSPosition pos;
	if (!zone.getGraphicShape(shape, pos))
		return;
	WPSGraphicStyle style;
	if (zone.m_graphicId)
		m_styleManager->updateGraphicStyle(zone.m_graphicId, style);
	m_listener->insertPicture(pos, data, "image/pict", style);
}

void LotusGraph::sendTextBox(std::shared_ptr<WPSStream> stream, WPSEntry const &entry)
{
	if (!stream || !m_listener || entry.length()<1)
	{
		WPS_DEBUG_MSG(("LotusGraph::sendTextBox: I can not find the listener/textbox entry\n"));
		return;
	}
	RVNGInputStreamPtr input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = entry.begin();
	long sz=entry.length();
	f << "GraphTextBox[data]:";
	input->seek(pos, librevenge::RVNG_SEEK_SET);
	auto val=int(libwps::readU8(input)); // always 1?
	if (val!=1) f << "f0=" << val << ",";
	auto fontType = m_mainParser.getDefaultFontType();
	WPSFont font=WPSFont::getDefault();
	m_listener->setFont(font);
	bool actualFlags[7]= {false, false, false, false, false, false, false };
	std::string text;
	for (long i=1; i<=sz; ++i)
	{
		auto c=i+1==sz ? '\0' : char(libwps::readU8(input));
		if ((c==0 || c==0xe || c==0xf) && !text.empty())
		{
			m_listener->insertUnicodeString(libwps_tools_win::Font::unicodeString(text, fontType));
			text.clear();
		}
		if (c==0)
		{
			if (int(i+2)<int(sz)) // g++ v7 warms if we test i+2<sz :-~
			{
				WPS_DEBUG_MSG(("LotusGraph::sendTextBox: find a 0 char\n"));
				f << "[###0]";
			}
			continue;
		}
		if (c!=0xe && c!=0xf)
		{
			f << c;
			text.push_back(c);
			continue;
		}
		if (i+1>=sz)
		{
			WPS_DEBUG_MSG(("LotusGraph::sendTextBox: find modifier in last pos\n"));
			f << "[###" << int(c) << "]";
		}
		auto mod=int(libwps::readU8(input));
		++i;
		if (c==0xf)
		{
			if (mod==45)
			{
				f << "[break]";
				m_listener->insertEOL();
			}
			else
			{
				WPS_DEBUG_MSG(("LotusGraph::sendTextBox: find unknown modifier f\n"));
				f << "[###f:" << mod << "]";
			}
			continue;
		}
		int szParam=(mod==0x80) ? 4 : (mod>=0x40 && mod<=0x44) ? 2 : 0;
		if (i+1+2*szParam>=sz)
		{
			WPS_DEBUG_MSG(("LotusGraph::sendTextBox: the param size seems bad\n"));
			f << "[##e:" << std::hex << mod << std::dec << "]";
			continue;
		}
		int param=0;
		long actPos=input->tell();
		bool ok=true;
		for (int d=0; d<szParam; ++d)
		{
			auto mod1=int(libwps::readU8(input));
			val=int(libwps::readU8(input));
			static int const decal[]= {1,0,3,2};
			if (mod1==0xe && (val>='0'&&val<='9'))
			{
				param += (val-'0')<<(4*decal[d]);
				continue;
			}
			else if (mod1==0xe && (val>='A'&&val<='F'))
			{
				param += (val-'A'+10)<<(4*decal[d]);
				continue;
			}
			WPS_DEBUG_MSG(("LotusGraph::sendTextBox: something when bad when reading param\n"));
			f << "[##e:" << std::hex << mod << ":" << param << std::dec << "]";
			ok=false;
		}
		if (!ok)
		{
			input->seek(actPos, librevenge::RVNG_SEEK_SET);
			continue;
		}
		i+=2*szParam;
		switch (mod)
		{
		case 1:
		case 2:
		case 3:
		case 4:
		case 5:
		case 6:
		case 7:
		{
			bool newFlag=actualFlags[mod-1]=!actualFlags[mod-1];
			static char const *wh[]= {"b", "it", "outline", "underline", "shadow", "condensed", "extended"};
			f << "[";
			if (!newFlag) f << "/";
			f << wh[mod-1] << "]";
			if (mod<=5)
			{
				static uint32_t const attrib[]= { WPS_BOLD_BIT, WPS_ITALICS_BIT, WPS_UNDERLINE_BIT, WPS_OUTLINE_BIT, WPS_SHADOW_BIT };
				if (newFlag)
					font.m_attributes |= attrib[mod-1];
				else
					font.m_attributes &= ~attrib[mod-1];
			}
			else
			{
				font.m_spacing=0;
				if (actualFlags[5]) font.m_spacing-=2;
				if (actualFlags[6]) font.m_spacing+=2;
			}
			m_listener->setFont(font);
			break;
		}
		case 0x40:
		{
			WPSFont newFont;
			f << "[FN" << param<< "]";
			if (m_mainParser.getFont(param, newFont, fontType))
			{
				font.m_name=newFont.m_name;
				m_listener->setFont(font);
			}
			else
				f << "###";
			break;
		}
		case 0x41:
		{
			f << "[color=" << param << "]";
			WPSColor color;
			if (m_styleManager->getColor256(param, color))
			{
				font.m_color=color;
				m_listener->setFont(font);
			}
			else
				f << "###";
			break;
		}
		case 0x44:
		{
			WPSParagraph para;
			switch (param)
			{
			case 1:
				f << "align[left]";
				para.m_justify=libwps::JustificationLeft;
				break;
			case 2:
				f << "align[right]";
				para.m_justify=libwps::JustificationRight;
				break;
			case 3:
				f << "align[center]";
				para.m_justify=libwps::JustificationCenter;
				break;
			default:
				f << "#align=" << param << ",";
				break;
			};
			m_listener->setParagraph(para);
			break;
		}
		case 0x80:
		{
			f << "[fSz=" << param/32. << "]";
			font.m_size=param/32.;
			m_listener->setFont(font);
			break;
		}
		default:
			WPS_DEBUG_MSG(("LotusGraph::sendTextBox: Oops find unknown modifier e\n"));
			f << "[##e:" << std::hex << mod << "=" << param << std::dec << "]";
			break;
		}
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
}

void LotusGraph::sendGraphics(int sheetId)
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("LotusGraph::sendGraphics: I can not find the listener\n"));
		return;
	}
	auto it=m_state->m_sheetIdZoneMacMap.lower_bound(sheetId);
	while (it!=m_state->m_sheetIdZoneMacMap.end() && it->first==sheetId)
	{
		auto zone=(it++)->second;
		if (!zone) continue;
		if (zone->m_pictureEntry.valid())
		{
			sendPicture(*zone);
			continue;
		}
		WPSGraphicShape shape;
		WPSPosition pos;
		if (!zone->getGraphicShape(shape, pos))
			continue;
		WPSGraphicStyle style;
		if (zone->m_lineId)
			m_styleManager->updateLineStyle(zone->m_lineId, style);
		if (zone->m_surfaceId)
			m_styleManager->updateSurfaceStyle(zone->m_surfaceId, style);
		if (zone->m_graphicId)
			m_styleManager->updateGraphicStyle(zone->m_graphicId, style);
		if (zone->m_textBoxEntry.valid())
		{
			std::shared_ptr<LotusGraphInternal::SubDocument> doc
			(new LotusGraphInternal::SubDocument(zone->m_stream, *this, zone->m_textBoxEntry, 0));
			m_listener->insertTextBox(pos, doc, style);
			continue;
		}
		if (zone->m_chartId)
		{
			m_mainParser.sendChart(zone->m_chartId, pos, style);
			continue;
		}
		if (zone->m_type==zone->Line)
		{
			if (zone->m_values[0]&1)
				style.m_arrows[1]=true;
			if (zone->m_values[0]&2)
				style.m_arrows[0]=true;
		}
		m_listener->insertPicture(pos, shape, style);
	}
	auto it4=m_state->m_sheetIdZoneWK4Map.lower_bound(sheetId);
	while (it4!=m_state->m_sheetIdZoneWK4Map.end() && it4->first==sheetId)
	{
		auto zone=it4++->second;
		if (!zone) continue;
		Vec2f decal;
		if (!m_mainParser.getLeftTopPosition(zone->m_cell, sheetId, decal))
			decal=Vec2f(float(72*zone->m_cell[0]),float(16*zone->m_cell[1]));
		Vec2f dimension(zone->m_frameSize);
		if (zone->m_type==zone->Shape)
			dimension=zone->m_shape.getBdBox().size();
		else if (zone->m_type==zone->Picture)
			dimension=Vec2f(zone->m_pictureDim.size());
		WPSPosition pos(decal+zone->m_cellPosition, dimension, librevenge::RVNG_POINT);
		pos.setRelativePosition(WPSPosition::Page);
		if (zone->m_type==zone->Shape)
			m_listener->insertPicture(pos, zone->m_shape, zone->m_graphicStyle);
		else if (zone->m_type==zone->TextBox)
		{
			std::shared_ptr<LotusGraphInternal::SubDocument> doc
			(new LotusGraphInternal::SubDocument(zone->m_stream, *this, zone->m_textEntry, zone->m_subType==0xd ? 2 : 1));
			m_listener->insertTextBox(pos, doc, zone->m_graphicStyle);
		}
		else if (zone->m_type==zone->Chart)
		{
			if (zone->m_pictureName.empty())
			{
				WPS_DEBUG_MSG(("LotusGraph::sendGraphics: sorry, can not find the chart name\n"));
			}
			else
			{
				auto nIt=m_state->m_nameToChartIdMap.find(zone->m_pictureName);
				if (nIt==m_state->m_nameToChartIdMap.end())
				{
					WPS_DEBUG_MSG(("LotusGraph::sendGraphics: sorry, can not find the chart id for %s\n", zone->m_pictureName.c_str()));
				}
				else
					m_mainParser.sendChart(nIt->second, pos, zone->m_graphicStyle);
			}
		}
		else if (zone->m_type==zone->Picture)
		{
			WPSEmbeddedObject object;
			if (m_mainParser.updateEmbeddedObject(zone->m_id, object) && !object.isEmpty())
				m_listener->insertObject(pos, object);
		}
		else
		{
			static bool first=true;
			if (first)
			{
				first=false;
				WPS_DEBUG_MSG(("LotusGraph::sendGraphics: sorry, sending some graph types is not implemented\n"));
			}
		}
	}
	if (m_state->m_zIdToSheetIdMap.find(sheetId)!=m_state->m_zIdToSheetIdMap.end())
	{
		int finalId=m_state->m_zIdToSheetIdMap.find(sheetId)->second;
		if (m_state->m_sheetIdZonePcListMap.find(finalId)!=m_state->m_sheetIdZonePcListMap.end())
		{
			auto const &zoneList=m_state->m_sheetIdZonePcListMap.find(finalId)->second;
			WPSTransformation transform;
			for (size_t i=0; i<zoneList.m_zones.size(); ++i)
				sendZone(zoneList, i, transform);
		}
	}
}

bool LotusGraph::readZoneBeginC9(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::read16(input));
	if (type!=0xc9)
	{
		WPS_DEBUG_MSG(("LotusGraph::readZoneBeginC9: not a sheet header\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	f << "Entries(FMTGraphBegin):";
	if (sz != 1)
	{
		WPS_DEBUG_MSG(("LotusGraph::readZoneBeginC9: the zone size seems bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	m_state->m_actualSheetId=int(libwps::readU8(input));
	f << "sheet[id]=" << m_state->m_actualSheetId << ",";
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusGraph::readFMTPictName(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = int(libwps::read16(input));
	if (type!=0xb7)
	{
		WPS_DEBUG_MSG(("LotusGraph::readFMTPictName: not a font name definition\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	if (sz!=0x68)
	{
		WPS_DEBUG_MSG(("LotusGraph::readFMTPictName: the zone size seems bad\n"));
		ascFile.addPos(pos);
		ascFile.addNote("Entries(FMTPictName):###");
		return true;
	}
	f << "Entries(FMTPictName):";
	std::string name;
	for (int i=0; i<16; ++i)
	{
		auto c=char(libwps::readU8(input));
		if (c==0) break;
		name+= c;
	}
	f << "name=" << name << ",";
	// <METAFILE> means data in Pictbf?
	if (version()==3)
	{
		if (!m_state->m_actualZoneWK4)
		{
			// rare, find also this zone isolated in the header zone...
			WPS_DEBUG_MSG(("LotusGraph::readFMTPictName: can not find the current chart\n"));
		}
		else
			m_state->m_actualZoneWK4->m_pictureName=name;
	}
	input->seek(pos+4+16, librevenge::RVNG_SEEK_SET);
	for (int i=0; i<2; ++i)
	{
		// seems ok in wk3 files but not in wk4 files
		auto col=int(libwps::readU8(input));
		auto table=int(libwps::readU8(input));
		auto row=int(libwps::readU16(input));
		f << "C" << col << "-" << row;
		if (table) f << "[" << table << "]";
		if (i==0)
			f << "<->";
		else
			f << ",";
	}
	for (int i=0; i<5; ++i)   // f1=0|1|3
	{
		auto val=int(libwps::readU16(input));
		if (!val) continue;
		f << "f" << i << "=" << val << ",";
	}
	int dim[2];
	for (auto &d : dim) d=int(libwps::readU16(input));
	if (dim[0]||dim[1]) f << "dim=" << Vec2i(dim[0],dim[1]) << ",";
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	pos=input->tell();
	f.str("");
	f << "FMTPictName-A:" << ",";
	name.clear();
	for (int i=0; i<16; ++i)
	{
		auto c=char(libwps::readU8(input));
		if (c==0) break;
		name+= c;
	}
	if (!name.empty())
		// if file name is not empty, we will not retrieve the chart, ...
		f << "fileName=" << name << ",";
	input->seek(pos+16, librevenge::RVNG_SEEK_SET);
	for (int i=0; i<25; ++i)
	{
		auto val=int(libwps::readU16(input));
		if (!val) continue;
		f << "f" << i << "=" << val << ",";
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusGraph::readGraphic(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::read16(input));
	if (type!=0xca)
	{
		WPS_DEBUG_MSG(("LotusGraph::readGraphic: not a sheet header\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	f << "Entries(FMTGraphic):";
	if (sz < 0x23)
	{
		WPS_DEBUG_MSG(("LotusGraph::readGraphic: the zone size seems bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto mType=int(libwps::readU8(input));
	if (mType==2) f << "graph,";
	else if (mType==4) f << "group,";
	else if (mType==5) f << "chart,";
	else if (mType==0xa) f << "textbox,";
	else if (mType==0xb) f << "cell[border],";
	else if (mType==0xc) f << "graph,";
	else f << "type[main]=" << mType << ",";

	std::shared_ptr<LotusGraphInternal::ZoneWK4> zone(new LotusGraphInternal::ZoneWK4(stream));
	zone->m_subType=int(libwps::readU8(input));
	switch (zone->m_subType)
	{
	case 1:
		zone->m_type=LotusGraphInternal::ZoneWK4::Shape;
		f << "line,";
		break;
	case 2:
		zone->m_type=LotusGraphInternal::ZoneWK4::Shape;
		f << "poly,";
		break;
	case 4:
		zone->m_type=LotusGraphInternal::ZoneWK4::Shape;
		f << "arc,";
		break;
	case 5:
		zone->m_type=LotusGraphInternal::ZoneWK4::Shape;
		f << "spline,";
		break;
	case 6:
		zone->m_type=LotusGraphInternal::ZoneWK4::Shape;
		f << "rect,";
		break;
	case 7:
		zone->m_type=LotusGraphInternal::ZoneWK4::Shape;
		f << "rect[round],";
		break;
	case 8:
		zone->m_type=LotusGraphInternal::ZoneWK4::Shape;
		f << "oval,";
		break;
	case 9:
		zone->m_type=LotusGraphInternal::ZoneWK4::Chart;
		f << "chart,";
		break;
	case 0xa:
		zone->m_type=LotusGraphInternal::ZoneWK4::Group;
		f << "group,";
		break;
	case 0xd:
		zone->m_type=LotusGraphInternal::ZoneWK4::TextBox;
		f << "button,";
		break;
	case 0xe:
		zone->m_type=LotusGraphInternal::ZoneWK4::TextBox;
		f << "textbox,";
		break;
	case 0x10:
		zone->m_type=LotusGraphInternal::ZoneWK4::Border;
		f << "cell[border],";
		break;
	case 0x11:
		zone->m_type=LotusGraphInternal::ZoneWK4::Picture;
		f << "picture,";
		break;
	default:
		WPS_DEBUG_MSG(("LotusGraph::readGraphic: find unknown graphic type=%d\n", zone->m_subType));
		f << "##type[local]=" << zone->m_subType << ",";
		break;
	}
	auto val=int(libwps::readU8(input)); // 0|80
	if (val) f << "fl0=" << std::hex << val << std::dec << ",";
	zone->m_id=int(libwps::readU16(input));
	f << "id=" << zone->m_id << ",";

	f << "line=[";
	val=int(libwps::readU8(input));
	WPSGraphicStyle &style=zone->m_graphicStyle;
	if (!m_styleManager->getColor256(val, style.m_lineColor))
	{
		WPS_DEBUG_MSG(("LotusGraph::readGraphic: can not read a color\n"));
		f << "###colId=" << val << ",";
	}
	else if (!style.m_lineColor.isBlack()) f << style.m_lineColor << ",";
	val=int(libwps::readU8(input));
	bool noLine=false;
	if (val<8)
	{
		switch (val)
		{
		case 0:
			f << "none,";
			noLine=true;
			break;
		case 2:
			style.m_lineDashWidth.push_back(7);
			style.m_lineDashWidth.push_back(3);
			f << "dash7x3";
			break;
		case 3:
			style.m_lineDashWidth.push_back(4);
			style.m_lineDashWidth.push_back(4);
			f << "dot4x4";
			break;
		case 4:
			style.m_lineDashWidth.push_back(6);
			style.m_lineDashWidth.push_back(2);
			style.m_lineDashWidth.push_back(4);
			style.m_lineDashWidth.push_back(2);
			f << "dash6x2:4x2";
			break;
		case 5:
			style.m_lineDashWidth.push_back(4);
			style.m_lineDashWidth.push_back(2);
			style.m_lineDashWidth.push_back(2);
			style.m_lineDashWidth.push_back(2);
			style.m_lineDashWidth.push_back(2);
			style.m_lineDashWidth.push_back(2);
			f << "dash4x2:2x2:2x2";
			break;
		case 6:
			style.m_lineDashWidth.push_back(2);
			style.m_lineDashWidth.push_back(2);
			f << "dot2x2";
			break;
		case 7:
			style.m_lineDashWidth.push_back(1);
			style.m_lineDashWidth.push_back(1);
			f << "dot1x1";
			break;
		default:
			break;
		}
	}
	else
	{
		WPS_DEBUG_MSG(("LotusGraph::readGraphic: can not read the line's style\n"));
		f << "###style=" << val << ",";
	}
	val=int(libwps::readU8(input));
	if (val<8)
	{
		style.m_lineWidth = noLine ? 0.f : float(val+1);
		if (val) f << "w=" << val+1 << ",";
	}
	else
	{
		style.m_lineWidth= noLine ? 0 : 1;
		WPS_DEBUG_MSG(("LotusGraph::readGraphic: can not read the line's width\n"));
		f << "###width=" << val << ",";
	}
	f << "],";
	f << "surf=[";
	int colId[2];
	for (int i=0; i<2; ++i)
	{
		colId[1-i]=int(libwps::readU8(input));
		f << colId[1-i] << ",";
	}
	auto patternId=int(libwps::readU8(input));
	f << patternId << ",";
	if (!m_styleManager->updateSurfaceStyle(colId[0], colId[1], patternId, style))
		f << "###";
	f << "],";
	f << "shadow=["; // border design
	val=int(libwps::readU8(input));
	WPSColor color;
	if (!m_styleManager->getColor256(val, color))
	{
		WPS_DEBUG_MSG(("LotusGraph::readGraphic: can not read a color\n"));
		f << "###colId=" << val << ",";
	}
	else if (!color.isBlack()) f << color << ",";
	val=int(libwps::readU8(input)); // 0=none
	if (val)
	{
		zone->m_hasShadow=true;
		f << "type=" << val << ",";
	}
	f << "],";
	float matrix[6];
	for (float &i : matrix) i=float(libwps::read32(input))/65536.f;
	WPSTransformation transform(WPSVec3f(matrix[0],matrix[1],matrix[4]), WPSVec3f(matrix[2],matrix[3],matrix[5]));
	if (!transform.isIdentity())
		f << "mat=[" << matrix[0] << "," << matrix[1] << "," << matrix[4]
		  << " ," << matrix[2] << "," << matrix[3] << "," << matrix[5] << "],";
	unsigned long lVal=libwps::readU32(input); // unsure maybe related to some angle
	if (lVal) f << "unkn=" << std::hex << lVal << std::dec << ",";
	for (int i=0; i<2; ++i)   // f0=0|1|2, f1=0|5|6
	{
		val=int(libwps::readU8(input));
		if (val) f << "f" << i << "=" << val << ",";
	}
	switch (zone->m_subType)
	{
	case 1:   // line
	{
		if (sz!=0x37)
			break;
		val=int(libwps::readU16(input));
		if (val!=2) f << "g0=" << val << ",";
		val=int(libwps::readU16(input));
		if (val&1)
		{
			style.m_arrows[0]=true;
			f << "arrow[beg],";
		}
		if (val&2)
		{
			style.m_arrows[1]=true;
			f << "arrow[end],";
		}
		val&=0xFFFC;
		if (val) f << "g1=" << std::hex << val << std::dec << ",";
		int pts[4];
		for (int &pt : pts) pt=int(libwps::readU16(input));
		f << "pts=" << Vec2i(pts[0],pts[1]) << "<->" << Vec2i(pts[2],pts[3]) << ",";
		zone->m_shape=WPSGraphicShape::line(Vec2f(float(pts[0]),float(pts[1])), Vec2f(float(pts[2]),float(pts[3])));
		break;
	}
	case 4:   // arc
	{
		if (sz!=0x3b)
			break;
		val=int(libwps::readU16(input));
		if (val!=3) f << "g0=" << val << ",";
		val=int(libwps::readU16(input)); // 0208
		if (val) f << "g1=" << std::hex << val << std::dec << ",";
		f << "pts=[";
		std::vector<Vec2f> vertices;
		vertices.resize(3);
		WPSBox2f box;
		for (size_t i=0; i<3; ++i)
		{
			int pts[2];
			for (int &pt : pts) pt=int(libwps::readU16(input));
			vertices[i]=Vec2f(float(pts[0]),float(pts[1]));
			if (i==0) box=WPSBox2f(vertices[i],vertices[i]);
			else box=box.getUnion(WPSBox2f(vertices[i],vertices[i]));
			f << vertices[i] << ",";
		}
		f << "],";
		// not frequent, approximate it by a Bezier's curve
		zone->m_shape=WPSGraphicShape::path(box);
		zone->m_shape.m_path.push_back(WPSGraphicShape::PathData('M', vertices[0]));
		zone->m_shape.m_path.push_back(WPSGraphicShape::PathData('Q', vertices[2], vertices[1]));
		break;
	}
	case 2:
	case 5:
	{
		auto N=int(libwps::readU16(input));
		if (sz!=4*N+0x2f) break;
		val=int(libwps::readU16(input)); // 0
		if (val) f << "g0=" << val << ",";
		std::vector<Vec2f> vertices;
		vertices.resize(size_t(N));
		f << "pts=[";
		WPSBox2f box;
		for (size_t i=0; i<size_t(N); ++i)
		{
			int pts[2];
			for (int &pt : pts) pt=int(libwps::readU16(input));
			vertices[i]=Vec2f(float(pts[0]),float(pts[1]));
			if (i==0) box=WPSBox2f(vertices[i],vertices[i]);
			else box=box.getUnion(WPSBox2f(vertices[i],vertices[i]));
			f << vertices[i] << ",";
		}
		f << "],";
		if (zone->m_subType==2 || vertices.size()<=1)
		{
			zone->m_shape=WPSGraphicShape::polygon(box);
			zone->m_shape.m_vertices=vertices;
		}
		else
		{
			zone->m_shape=WPSGraphicShape::path(box);
			zone->m_shape.m_path.push_back(WPSGraphicShape::PathData('M', vertices[0]));
			for (size_t i=0; i+1<vertices.size(); ++i)
				zone->m_shape.m_path.push_back(WPSGraphicShape::PathData('Q', 0.5f*(vertices[i]+vertices[i+1]),
				                                                         vertices[i]));
			zone->m_shape.m_path.push_back(WPSGraphicShape::PathData('T', vertices.back()));
		}
		break;
	}
	case 6: // rect
	case 7: // rectround
	case 8:   // oval
	{
		if (sz!=0x3f)
			break;
		val=int(libwps::readU16(input));
		if (val!=4) f << "g0=" << val << ",";
		for (int i=0; i<2; ++i)   // g1=4, g2=0|1
		{
			val=int(libwps::readU8(input));
			if (i==1)
			{
				if (val&1)
					f << "round,";
				else if (val&2)
					f << "oval,";
				val &= 0xFC;
			}
			if (val)
				f << "g" << i+1 << "=" << val << ",";
		}
		WPSBox2f box;
		f << "pts=[";
		for (int i=0; i<4; ++i)
		{
			int pts[2];
			for (int &pt : pts) pt=int(libwps::readU16(input));
			Vec2f pt=Vec2f(float(pts[0]), float(pts[1]));
			f << pt << ",";
			if (i==0)
				box=WPSBox2f(pt,pt);
			else
				box=box.getUnion(WPSBox2f(pt,pt));
		}
		f << "],";
		if (zone->m_subType==8)
			zone->m_shape=WPSGraphicShape::circle(box);
		else
			zone->m_shape=WPSGraphicShape::rectangle(box, zone->m_subType==6 ? Vec2f(0,0) : Vec2f(5,5));
		break;
	}
	case 9: // chart, readme
		if (sz!=0x33)
			break;
		f << "dim=";
		for (int i=0; i<2; ++i)
		{
			int pts[2];
			for (int &pt : pts) pt=int(libwps::readU16(input));
			f << Vec2i(pts[0],pts[1]) << (i==0 ? "<->" : ",");
		}
		break;
	case 10: // group, useme
		if (sz!=0x35)
			break;
		f << "pts=[";
		for (int i=0; i<2; ++i)
		{
			int pts[2];
			for (int &pt : pts) pt=int(libwps::readU16(input));
			f << Vec2i(pts[0],pts[1]) << ",";
		}
		f << "],";
		val=int(libwps::readU16(input));
		if (val!=1) f << "g0=" << val << ",";
		break;
	case 0xd:
	case 0xe:
		if (sz!=0x35) break;
		f << "pts=";
		for (int i=0; i<4; ++i)
		{
			f << int(libwps::readU16(input));
			if (i==1) f << "<->";
			else if (i==3) f << ",";
			else f << "x";
		}
		val=int(libwps::readU16(input)); // small number
		f << "g0=" << val << ",";
		break;
	case 0x10:
		if (sz!=0x34) break;
		for (int i=0; i<2; ++i)   // 0
		{
			val=int(libwps::readU16(input));
			if (val) f << "g" << i << "=" << val << ",";
		}
		for (int i=0; i<3; ++i)   // g2=[0-f][0|8], g3=0-2 g4=??
		{
			val=int(libwps::readU8(input));
			if (val) f << "g" << i+2 << "=" << std::hex << val << std::dec << ",";
		}
		val=int(libwps::readU16(input)); // 0-2
		if (val) f << "g5=" << val << ",";
		break;
	case 0x11:
	{
		if (sz!=0x43) break;
		Vec2i dim[2];
		for (auto &i : dim)
		{
			int pts[2];
			for (int &pt : pts) pt=int(libwps::readU16(input));
			i=Vec2i(pts[0],pts[1]);
		}
		zone->m_pictureDim=WPSBox2i(dim[0], dim[1]);
		f << "dim=" << zone->m_pictureDim << ",";
		val=int(libwps::readU16(input)); // big number 2590..3262
		if (val) f << "g0=" << std::hex << val << std::dec << ",";
		val=int(libwps::readU16(input));
		if (val!=0x3cf7) f << "g1=" << std::hex << val << std::dec << ",";
		for (int i=0; i<6; ++i)   // 0
		{
			val=int(libwps::readU16(input));
			if (val) f << "g" << i+2 << "=" << val << ",";
		}
		break;
	}
	default:
		break;
	}
	if (zone->m_shape.m_type!=WPSGraphicShape::ShapeUnknown && !transform.isIdentity())
		zone->m_shape=zone->m_shape.transform(transform);
	if (m_state->m_actualZoneWK4)
	{
		WPS_DEBUG_MSG(("LotusGraph::readGraphic: oops an zone is already defined\n"));
	}
	m_state->m_actualZoneWK4=zone;
	if (input->tell() != pos+4+sz)
		ascFile.addDelimiter(input->tell(), '|');
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool LotusGraph::readFrame(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::read16(input));
	if (type!=0xcc)
	{
		WPS_DEBUG_MSG(("LotusGraph::readFrame: not a frame header\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	f << "Entries(FMTFrame):";
	if (sz != 0x13)
	{
		WPS_DEBUG_MSG(("LotusGraph::readFrame: the zone size seems bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto zone=m_state->m_actualZoneWK4;
	if (!zone)
	{
		WPS_DEBUG_MSG(("LotusGraph::readFrame: can not find the original shape\n"));
		f << "##noShape,";
	}
	/* the positions are relative to a cell ; the first cell stores the LT position,
	   while the second stores the RB position.

	   fixme: if we want precise position, we must add graphic with anchor-type=cell
	     instead of using the first cell to find the page's LT position...
	 */
	for (int c=0; c<2; ++c)
	{
		auto row=int(libwps::readU16(input));
		auto col=int(libwps::readU8(input));
		int pts[2]; // first in 100.*char, second in point
		for (int &pt : pts) pt=int(libwps::readU16(input));
		Vec2f decal(8.f*float(pts[0])/100.f, float(pts[1]));
		Vec2i cell(col, row);
		if (c==0 && zone)
		{
			zone->m_cell=cell;
			zone->m_cellPosition=decal;
		}
		f << "C" << cell << "[" << decal << "]";
		if (c==0) f << "<->";
		else f << ",";
	}
	// this is the initial dimension, which may become invalid when a column/row shrinks/grows...
	int dim[2];
	for (int &i : dim) i=int(libwps::readU16(input));
	f << "dim=" << Vec2i(dim[0], dim[1]) << ",";
	if (zone) zone->m_frameSize=Vec2i(dim[0], dim[1]);
	auto val=int(libwps::readU8(input)); // 1|2
	if (val&0x80) f << "in[group],";
	val &= 0x7F;
	if (val!=1) f << "fl0=" << val << ",";
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	if (zone)
	{
		if (m_state->m_actualSheetId<0)
		{
			WPS_DEBUG_MSG(("LotusGraph::readFrame: oops no sheet zone is opened\n"));
			f << "###sheetId,";
		}
		else
			m_state->m_sheetIdZoneWK4Map.insert(std::multimap<int, std::shared_ptr<LotusGraphInternal::ZoneWK4> >::value_type
			                                    (m_state->m_actualSheetId, zone));

	}
	m_state->m_actualZoneWK4.reset();

	return true;
}

bool LotusGraph::readTextBoxDataD1(std::shared_ptr<WPSStream> stream)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::read16(input));
	if (type!=0xd1)
	{
		WPS_DEBUG_MSG(("LotusGraph::readTextBoxDataD1: not a textbox header\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	f << "Entries(FMTTextBox):";
	if (!m_state->m_actualZoneWK4 || m_state->m_actualZoneWK4->m_type!=LotusGraphInternal::ZoneWK4::TextBox)
	{
		WPS_DEBUG_MSG(("LotusGraph::readTextBoxDataD1: find unexpected textbox data\n"));
		f << "###";
	}
	else
	{
		m_state->m_actualZoneWK4->m_textEntry.setBegin(input->tell());
		m_state->m_actualZoneWK4->m_textEntry.setLength(sz);
		input->seek(sz, librevenge::RVNG_SEEK_CUR);
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

void LotusGraph::sendTextBoxWK4(std::shared_ptr<WPSStream> stream, WPSEntry const &entry, bool isButton)
{
	if (!stream || !m_listener || (entry.length() && entry.length()<3))
	{
		WPS_DEBUG_MSG(("LotusGraph::sendTextBoxWK4: I can not find the listener/textbox entry\n"));
		return;
	}
	RVNGInputStreamPtr input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long endPos=entry.end();
	input->seek(entry.begin(), librevenge::RVNG_SEEK_SET);

	auto fontType = m_mainParser.getDefaultFontType();
	WPSFont font=WPSFont::getDefault();
	m_listener->setFont(font);
	while (!input->isEnd())
	{
		long pos = input->tell();
		if (pos+3>endPos) break;
		f.str("");
		f << "FMTTextBox[data]:";
		auto dSz=int(libwps::readU16(input));
		if (pos+2+dSz>endPos)
		{
			input->seek(pos, librevenge::RVNG_SEEK_SET);
			break;
		}
		bool send=!isButton || pos==entry.begin();
		std::string text;
		for (int i=0; i<dSz; ++i)
		{
			auto c=i+1==dSz ? '\0' : char(libwps::readU8(input));
			if ((c==0 || c==1) && !text.empty())
			{
				if (send)
					m_listener->insertUnicodeString(libwps_tools_win::Font::unicodeString(text, fontType));
				text.clear();
			}
			if (c==0)
			{
				if (i+2<dSz)
				{
					WPS_DEBUG_MSG(("LotusGraph::sendTextBox: find a 0 char\n"));
					f << "[###0]";
				}
				continue;
			}
			if (c==1)
			{
				if (i+2>=dSz)
				{
					WPS_DEBUG_MSG(("LotusGraph::sendTextBox: find unexpected 1 char\n"));
					f << "[###1]";
					continue;
				}
				++i;
				c=char(libwps::readU8(input));
				f << "[1-" << int(c) << "]";
				// find 010d010d010a
				if (send)
				{
					if (c==0xd) m_listener->insertEOL(false);
					else if (c==0xa) m_listener->insertEOL();
					else
					{
						WPS_DEBUG_MSG(("LotusGraph::sendTextBox: find unexpected 1 char\n"));
						f << "###";
					}
				}
				continue;
			}
			f << c;
			text.push_back(c);
		}
		f << ",unk=" << int(libwps::readU8(input));
		if (input->tell()<endPos)
			m_listener->insertEOL();
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
	}
	if (isButton && input->tell()+1==endPos)
	{
		f.str("");
		f << "FMTTextBox[data]:button=" << int(libwps::readU8(input)) << ",";
		ascFile.addPos(endPos-1);
		ascFile.addNote(f.str().c_str());
	}
	if (input->tell()!=endPos)
	{
		WPS_DEBUG_MSG(("LotusGraph::sendTextBoxWK4: find extra data\n"));
		ascFile.addPos(input->tell());
		ascFile.addNote("FMTTextBox[data]:###extra");
	}
}

void LotusGraph::sendZone(LotusGraphInternal::ZonePcList const &zoneList, size_t id, WPSTransformation &transf)
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("LotusGraph::sendZone: can not find the listener\n"));
		return;
	}
	std::vector<std::shared_ptr<LotusGraphInternal::ZonePc> > const &zones=zoneList.m_zones;
	if (id>=zones.size())
	{
		WPS_DEBUG_MSG(("LotusGraph::sendZone: can not find the sub zone %d\n", int(id)));
		return;
	}
	auto zone=zones[id];
	if (!zone || zone->m_isSent) return;
	zone->m_isSent=true;
	WPSTransformation finalTrans=transf*zone->getTransformation();
	if (zone->m_type==zone->Set)
	{
		if (!zone->m_isGroup || zone->m_groupLastPosition+1<=id) return;
		WPSPosition pos(zone->m_box[0],zone->m_box.size(), librevenge::RVNG_POINT);
		pos.setRelativePosition(WPSPosition::Page);
		if (!m_listener->openGroup(pos)) return;
		for (size_t i=id+1; i<zone->m_groupLastPosition; ++i)
			sendZone(zoneList, i, finalTrans);
		m_listener->closeGroup();
		return;
	}
	if (zone->m_type==zone->Picture)
	{
		if (!zone->m_pictureData.empty())
		{
			WPSPosition pos(zone->m_box[0],zone->m_box.size(), librevenge::RVNG_POINT);
			pos.setRelativePosition(WPSPosition::Page);
			m_listener->insertPicture(pos, zone->m_pictureData);
#ifdef DEBUG_WITH_FILES
			std::stringstream s;
			static int fileId=0;
			s << "Pict" << ++fileId << ".emf";
			libwps::Debug::dumpFile(zone->m_pictureData, s.str().c_str());
#endif
		}
		return;
	}

	WPSGraphicShape shape;
	WPSPosition pos;
	if (!zone->getGraphicShape(shape, pos))
		return;
	WPSGraphicStyle style;
	if (zone->m_graphicId[0]>=0)
		m_styleManager->updateGraphicStyle(zone->m_graphicId[0], style);
	if (zone->m_type==zone->TextBox)
	{
		std::shared_ptr<LotusGraphInternal::SubDocument> doc
		(new LotusGraphInternal::SubDocument(zone->m_stream, *this, zone->m_textBoxEntry, 1));
		m_listener->insertTextBox(pos, doc, style);
		return;
	}
	if (zone->m_type==zone->Line)
	{
		if (zone->m_arrows&1)
			style.m_arrows[1]=true;
		if (zone->m_arrows&2)
			style.m_arrows[0]=true;
	}
	if (finalTrans.isIdentity())
		m_listener->insertPicture(pos, shape, style);
	else
	{
		// checkme: ok for translation but not for rotation...
		WPSGraphicShape finalShape=shape.transform(finalTrans);
		pos.setOrigin(finalTrans*pos.origin());
		pos.setSize(finalShape.getBdBox().size());
		m_listener->insertPicture(pos, finalShape, style);
	}
}

bool LotusGraph::readGraphZone(std::shared_ptr<WPSStream> stream, int zId)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	float const unit=version()>=5 ? 1.f/16.f : 1.f/256.f;
	long pos = input->tell();
	auto id = int(libwps::readU8(input));
	if (libwps::readU8(input) != 3)
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
	if (zId)
		f << "Entries(GraphZone)[Z" << zId << "]:";
	else
		f << "Entries(GraphZone)[_]:";
	if (id<0x80)
		m_state->m_actualZonePc.reset();
	int val;
	switch (id)
	{
	case 0: // rare, when it exists, present in sheet zone
		f << "zoneA0,";
		if (sz!=16)
		{
			WPS_DEBUG_MSG(("LotusGraph::readGraphZone: the size seems bad for zone 0\n"));
			f << "###";
			break;
		}
		for (int i=0; i<8; ++i)   // always f0=1, other 0
		{
			val=int(libwps::readU16(input));
			if (val) f << "f" << i << "=" << val << ",";
		}
		break;
	case 4: // seems link to the ref zone, unsure
		f << "ref,";
		if (sz!=20)
		{
			WPS_DEBUG_MSG(("LotusGraph::readGraphZone: the size seems bad for zone 4\n"));
			f << "###";
			break;
		}
		for (int i=0; i<10; ++i)   // always f0=1, other 0
		{
			val=int(libwps::readU16(input));
			if (val) f << "f" << i << "=" << val << ",";
		}
		break;

	case 0x83: // relative to graphic
	case 0x84: // rare, when it exists before chartName
	case 0x86: // rare, when it exists, end of a set. TODO useme
		f << (id==0x83 ? "GraphBeg" : id==0x84 ? "chartBeg" : "endSet") << ",";
		if (id==0x86)
		{
			if (zId<0 || m_state->m_sheetIdZonePcListMap.find(zId)==m_state->m_sheetIdZonePcListMap.end() ||
			        m_state->m_sheetIdZonePcListMap.find(zId)->second.m_groupBeginStack.empty())
			{
				WPS_DEBUG_MSG(("LotusGraph::readGraphZone: oops can not find the begin of the group\n"));
				f << "###group,";
			}
			else
			{
				auto &current=m_state->m_sheetIdZonePcListMap.find(zId)->second;
				size_t begPos=current.m_groupBeginStack.top();
				if (begPos<current.m_zones.size() && current.m_zones[begPos])
					current.m_zones[begPos]->m_groupLastPosition=current.m_zones.size();
				current.m_groupBeginStack.pop();
			}
		}
		if (sz!=0)
		{
			WPS_DEBUG_MSG(("LotusGraph::readGraphZone: the size seems bad for zone %d\n", id));
			f << "###";
			break;
		}
		break;
	case 0x81: // rare, when it exists, in sheetA4 zones (followed by sheetB5, sheetB6)
		f << "zoneB1,";
		if (sz!=6)
		{
			WPS_DEBUG_MSG(("LotusGraph::readGraphZone: the size seems bad for zone 81\n"));
			f << "###";
			break;
		}
		for (int i=0; i<3; ++i)   // always f0=1, other 0
		{
			val=int(libwps::read16(input));
			if (val) f << "f" << i << "=" << val << ",";
		}
		break;
	case 0x95: // relative to graphic
		f << "zoneB15,";
		if ((sz%4)!=0)
		{
			WPS_DEBUG_MSG(("LotusGraph::readGraphZone: the size seems bad for zone 815\n"));
			f << "###";
			break;
		}
		if (sz==0) break;
		val=int(libwps::read16(input));
		if (m_state->m_actualZonePc)
			m_state->m_actualZonePc->m_arrows=int(val&3);
		switch (val&3)
		{
		case 1:
			f << "arrow[beg],";
			break;
		case 2:
			f << "arrow[end],";
			break;
		case 3:
			f << "arrow[beg,end],";
			break;
		case 0:
		default:
			break;
		}
		val &=0xFFFC;
		if (val) f << "f0=" << val << ",";
		for (long i=1; i<sz/2; ++i)   // always f0=1, other 0
		{
			val=int(libwps::read16(input));
			if (val) f << "f" << i << "=" << val << ",";
		}
		break;
	case 0x85: // group or ???, after sheetB6.
	case 0x88: // line after sheetB4
	case 0x89: // polygon
	case 0x8a: // freehand
	case 0x8b: // rect
	case 0x8c: // ellipse
	case 0x8d: // arc
	case 0x8e: // picture
	case 0x90: // textbox
	case 0x9a:   // chart, after sheetB4
	{
		std::shared_ptr<LotusGraphInternal::ZonePc> zone(new LotusGraphInternal::ZonePc(stream));
		if (zId<0)
		{
			WPS_DEBUG_MSG(("LotusGraph::readGraphZone: oops can not find the sheet zone id\n"));
			f << "###sheetId,";
		}
		else
		{
			// CHECKME: sometimes this does not work, ie. zId is not the spreadssheet data id
			if (m_state->m_sheetIdZonePcListMap.find(zId)==m_state->m_sheetIdZonePcListMap.end())
				m_state->m_sheetIdZonePcListMap.insert(std::map<int, LotusGraphInternal::ZonePcList>::value_type(zId, LotusGraphInternal::ZonePcList()));
			auto &current=m_state->m_sheetIdZonePcListMap.find(zId)->second;
			if (id==0x85) current.m_groupBeginStack.push(current.m_zones.size());
			current.m_zones.push_back(zone);
		}

		m_state->m_actualZonePc=zone;
		int expectedSz=0;
		switch (id)
		{
		case 0x85:
			zone->m_type=zone->Set;
			expectedSz=80;
			break;
		case 0x88:
			zone->m_type=zone->Line;
			expectedSz=70;
			break;
		case 0x89:
			zone->m_type=zone->Polygon;
			expectedSz=70;
			break;
		case 0x8a:
			zone->m_type=zone->FreeHand;
			expectedSz=70;
			break;
		case 0x8b:
			zone->m_type=zone->Rect;
			expectedSz=110;
			break;
		case 0x8c:
			zone->m_type=zone->Ellipse;
			expectedSz=80;
			break;
		case 0x8d:
			zone->m_type=zone->Arc;
			expectedSz=84;
			break;
		case 0x8e:
			zone->m_type=zone->Picture;
			expectedSz=116;
			break;
		case 0x90:
			zone->m_type=zone->TextBox;
			expectedSz=112;
			break;
		case 0x9a:
			zone->m_type=zone->Chart;
			expectedSz=126;
			break;
		default:
			break;
		}

		if (sz<expectedSz)
		{
			WPS_DEBUG_MSG(("LotusGraph::readGraphZone: the size seems bad for zone %d\n", id));
			f << *zone << "###";
			break;
		}
		libwps::DebugStream f2;
		f2 << "id=" << libwps::read32(input) << ","; // small number or -1
		for (int i=0; i<5; ++i)   // f2=-1|4
		{
			val=int(libwps::readU16(input));
			if (val==0xFFFF) continue;
			switch (i)
			{
			case 0:
				if ((val&0x10)==0x10) zone->m_isGroup=true;
				if ((val&0x40)==0) f2 << "locked,";
				if ((val&0x200)==0) zone->m_isRoundRect=true;
				if ((val&0x2000)==0) f2 << "hidden,";
				val &=0xDDAF;
				if (val!=0x4d01)
					f2 << "f0=" << std::hex << val << std::dec << ",";
				break;
			case 1:
				if (val!=0x94)
					f2 << "f1=" << std::hex << val << std::dec << ",";
				break;
			case 2:
				if ((val>>8)==0x40)
					zone->m_graphicId[0]=(val&0xFF);
				else
				{
					WPS_DEBUG_MSG(("LotusGraph::readGraphZone: find unexpected graphic style id\n"));
					f2 << "###GS" << std::hex << val << std::dec << ",";
				}
				break;
			case 3:
				switch (val)
				{
				case 0: // TL and BR
					break;
				case 1:
					f2 << "fasten[TL],";
					break;
				case 2:
					f2 << "no[fasten],";
					break;
				default:
					f2 << "f3=" << val << ",";
				}
				break;
			default:
			{
				if (val)
					f2 << "f" << i << "=" << val << ",";
				break;
			}
			}
		}
		auto sSz=int(libwps::readU16(input));
		if ((id!=0x8e && expectedSz+sSz!=sz) || (id==0x8e && expectedSz+sSz>sz))
		{
			WPS_DEBUG_MSG(("LotusGraph::readGraphZone: the size seems bad for zone %d\n", id));
			f << *zone << "###" << f2.str();
			break;
		}
		std::string name;
		for (int i=0; i<sSz; ++i)
		{
			auto c=char(libwps::readU8(input));
			if (c) name+=c;
			else if (i+1!=sSz)
			{
				WPS_DEBUG_MSG(("LotusGraph::readGraphZone: find odd char in zone %d\n", id));
				f2 << "###";
			}
		}
		if (!name.empty()) f2 << name << ",";
		f2 << "unkn=[";
		for (int i=0; i<4; ++i) // often _,3ff00000,_,3ff00000
		{
			unsigned long lVal=libwps::readU32(input);
			if (lVal)
				f2 << std::hex << lVal << std::dec << ",";
			else
				f2 << "_,";
		}
		f2 << "],";
		bool hasFlip=false;
		for (int i=0; i<10; ++i)
		{
			val=int(libwps::readU16(input));
			if (!val) continue;
			if (i==0)
				zone->m_rotate=float(val)/10.f;
			else if (i==5 && val==1)
			{
				hasFlip=true;
				f2 << "has[flip],";
			}
			else
				f2 << "g" << i << "=" << std::hex << val << std::dec << ",";
		}
		float translate[2];
		for (float &i : translate) i=unit*float(libwps::read32(input));
		zone->m_translate=Vec2f(translate[0],translate[1]);
		for (int i=0; i<2; ++i)
		{
			val=int(libwps::readU16(input));
			if (!val) continue;
			if (i==0 && hasFlip && val<3)   // useme
			{
				if (val&1) f2 << "flipX,";
				if (val&2) f2 << "flipY,";
				continue;
			}
			f2 << "g" << i+10 << "=" << std::hex << val << std::dec << ",";
		}
		if (id>=0x88 && id<=0x8a)
		{
			val=int(libwps::read16(input)); // 0
			if (val) f2 << "h0=" << val << ",";
			zone->m_numPoints=int(libwps::readU16(input));
			val=int(libwps::read16(input)); // 0
			if (val) f2 << "h1=" << val << ",";
		}
		else
		{
			float dim[4];
			for (float &i : dim) i=unit*float(libwps::read32(input));
			zone->m_box=WPSBox2f(Vec2f(dim[0],dim[1]),Vec2f(dim[2],dim[3]));
			if (id==0x8b || id==0x8e || id==0x9a)
			{
				for (int i=0; i<2; ++i) // h0 small number
				{
					val=int(libwps::read16(input));
					if (val)
						f2 << "h" << i << "=" << val << ",";
				}
				val=int(libwps::readU16(input));
				if ((val>>8)==0x40)
					zone->m_graphicId[1]=(val&0xFF);
				else if (val)
				{
					WPS_DEBUG_MSG(("LotusGraph::readGraphZone: find unexpected graphic style id\n"));
					f2 << "###GS" << std::hex << val << std::dec << ",";
				}
			}
			if (id==0x8d)
			{
				for (int i=0; i<2; ++i) // 0
				{
					val=int(libwps::read16(input));
					if (val) f2 << "h" << i+3 << "=" << val << ",";
				}
			}
			else if (id==0x8e || id==0x9a)
			{
				for (int i=0; i<11; ++i)   // h3=h4=h6=0|1, h8=h10=0|2d, h13=-1
				{
					val=int(libwps::read16(input));
					if (val) f2 << "h" << i+3 << "=" << val << ",";
				}
				if (id==0x8e)
				{
					auto dSz=int(libwps::readU16(input));
					if (expectedSz+sSz+dSz!=sz)
					{
						WPS_DEBUG_MSG(("LotusGraph::readGraphZone: the size seems bad for zone %d\n", id));
						f << *zone << "###" << f2.str();
						break;
					}
					std::string dir;
					for (int i=0; i<dSz; ++i)
					{
						auto c=char(libwps::readU8(input));
						if (c) dir+=c;
						else if (i+1!=dSz)
						{
							WPS_DEBUG_MSG(("LotusGraph::readGraphZone: find odd char in zone %d\n", id));
							f2 << "###";
						}
					}
					if (!dir.empty()) f2 << dir << ",";
					for (int i=0; i<2; ++i)   // l0=19, l1=24: maybe some zoneId
					{
						val=int(libwps::read16(input));
						if (val) f2 << "l" << i << "=" << val << ",";
					}
				}
				else
				{
					f2 << "prev[id]=Z" << int(libwps::read32(input)) << ",";
					for (int i=0; i<5; ++i)
					{
						val=int(libwps::read16(input));
						int const expected[]= {0,0,1,1,0};
						if (val!=expected[i]) f2 << "l" << i << "=" << val << ",";
					}
					f2 << "act[id]=Z" << int(libwps::read32(input)) << ",";
				}
			}
		}
		zone->m_extra=f2.str();
		f << *zone;
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

bool LotusGraph::readGraphDataZone(std::shared_ptr<WPSStream> stream, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	float const unit=version()>=5 ? 1.f/16.f : 1.f/256.f;
	f << "Entries(GraphZone)[data]:";
	long pos = input->tell();
	auto sz=int(endPos-pos);
	if (m_state->m_actualZonePc && m_state->m_actualZonePc->m_type==LotusGraphInternal::ZonePc::Line && sz==16)
	{
		f << "line,";
		float dim[4];
		for (float &i : dim) i=unit*float(libwps::read32(input));
		m_state->m_actualZonePc->m_box= WPSBox2f(Vec2f(dim[0],dim[1]),Vec2f(dim[2],dim[3]));
		f << "dim=" << m_state->m_actualZonePc->m_box << ",";
	}
	else if (m_state->m_actualZonePc && sz==8*m_state->m_actualZonePc->m_numPoints &&
	         (m_state->m_actualZonePc->m_type==LotusGraphInternal::ZonePc::FreeHand || m_state->m_actualZonePc->m_type==LotusGraphInternal::ZonePc::Polygon))
	{
		f << "poly,pts=[";
		float dim[2];
		for (int n=0; n<m_state->m_actualZonePc->m_numPoints; ++n)
		{
			for (float &i : dim) i=unit*float(libwps::read32(input));
			m_state->m_actualZonePc->m_vertices.push_back(Vec2f(dim[0],dim[1]));
			f << m_state->m_actualZonePc->m_vertices.back() << ",";
		}
		f << "],";
	}
	else if (m_state->m_actualZonePc && m_state->m_actualZonePc->m_type==LotusGraphInternal::ZonePc::TextBox)
	{
		m_state->m_actualZonePc->m_textBoxEntry.setBegin(pos-2);
		m_state->m_actualZonePc->m_textBoxEntry.setEnd(endPos);
		f << "texbox,";
		std::string text;
		for (int i=0; i<sz; ++i)
		{
			auto c=char(libwps::readU8(input));
			if (c)
				text+=c;
			else if (i+1!=sz)
			{
				WPS_DEBUG_MSG(("LotusGraph::readGraphDataZone: find unexpected 0 char\n"));
				f << "###";
			}
		}
		f << text;
	}
	else if (m_state->m_actualZonePc && m_state->m_actualZonePc->m_type==LotusGraphInternal::ZonePc::Picture)
	{
		/* checkme: the picture is stored in a list on consecutif data zone
		   and seems preceded by 0100000008000000da0a0000da0a0000de380000
		 */
		if (endPos!=pos)
		{
			const unsigned char *readData;
			unsigned long sizeRead, expectedSize=static_cast<unsigned long>(endPos-pos);
			if ((readData=input->read(expectedSize, sizeRead)) == nullptr || sizeRead!=expectedSize)
			{
				WPS_DEBUG_MSG(("LotusGraph::readGraphDataZone: can not read the data\n"));
				f << "###";
			}
			else
			{
				f << "picture,";
				if (m_state->m_actualZonePc->m_pictureHeaderRead<20)   // skip the header? zone
				{
					int headerRead=m_state->m_actualZonePc->m_pictureHeaderRead+int(expectedSize) > 20 ?
					               20-m_state->m_actualZonePc->m_pictureHeaderRead : int(expectedSize);
					m_state->m_actualZonePc->m_pictureHeaderRead+=headerRead;
					if (headerRead<int(expectedSize))
					{
						m_state->m_actualZonePc->m_pictureData.append(readData+headerRead, static_cast<unsigned long>(endPos-pos-headerRead));
						ascFile.skipZone(pos+headerRead, endPos-1);
					}
				}
				else
				{
					m_state->m_actualZonePc->m_pictureData.append(readData, sizeRead);
					ascFile.skipZone(pos, endPos-1);
				}
			}
		}
	}
	else
	{
		WPS_DEBUG_MSG(("LotusGraph::readGraphDataZone: find unknown data zone\n"));
		f << "###";
	}
	ascFile.addPos(pos-4);
	ascFile.addNote(f.str().c_str());
	if (input->tell()!=endPos)
		ascFile.addDelimiter(input->tell(),'|');
	input->seek(endPos, librevenge::RVNG_SEEK_SET);
	return true;
}

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
