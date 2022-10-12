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

//#include "QuattroChart.h"
#include "QuattroFormula.h"
#include "QuattroGraph.h"
#include "QuattroSpreadsheet.h"

#include "Quattro.h"

using namespace libwps;

//! Internal: namespace to define internal class of QuattroParser
namespace QuattroParserInternal
{
//! the font of a QuattroParser
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
	SubDocument(RVNGInputStreamPtr const &input, QuattroParser &pars, bool header) :
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
		WPS_DEBUG_MSG(("QuattroParserInternal::SubDocument::parse: no listener\n"));
		return;
	}
	if (!dynamic_cast<WKSContentListener *>(listener.get()))
	{
		WPS_DEBUG_MSG(("QuattroParserInternal::SubDocument::parse: bad listener\n"));
		return;
	}

	QuattroParser *pser = m_parser ? dynamic_cast<QuattroParser *>(m_parser) : nullptr;
	if (!pser)
	{
		listener->insertCharacter(' ');
		WPS_DEBUG_MSG(("QuattroParserInternal::SubDocument::parse: bad parser\n"));
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

//! the state of QuattroParser
struct State
{
	//! constructor
	explicit State(libwps_tools_win::Font::Type fontType, char const *password)
		: m_fontType(fontType)
		, m_version(-1)
		, m_metaData()
		, m_actualSheet(-1)
		, m_fontsList()
		, m_colorsList()
		, m_idToExternalFileMap()
		, m_idToExternalNameMap()
		, m_idToFieldMap()
		, m_pageSpan()
		, m_actPage(0)
		, m_numPages(0)
		, m_headerString()
		, m_footerString()
		, m_password(password)
		, m_isEncrypted(false)
		, m_isDecoded(false)
		, m_idToZoneNameMap()
		, m_readingZone341(false)
	{
	}
	//! return the default font style
	libwps_tools_win::Font::Type getDefaultFontType() const
	{
		if (m_fontType != libwps_tools_win::Font::UNKNOWN)
			return m_fontType;
		return libwps_tools_win::Font::WIN3_WEUROPE;
	}

	//! returns a color corresponding to an id
	bool getColor(int id, WPSColor &color) const;
	//! returns a default font (Courier12) with file's version to define the default encoding */
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
	//! the meta data
	librevenge::RVNGPropertyList m_metaData;
	//! the actual sheet
	int m_actualSheet;
	//! the font list
	std::vector<Font> m_fontsList;
	//! the color list
	std::vector<WPSColor> m_colorsList;
	//! map id to filename
	std::map<int, librevenge::RVNGString> m_idToExternalFileMap;
	//! map id to external name
	std::map<int, librevenge::RVNGString> m_idToExternalNameMap;
	//! map id to field
	std::map<int, std::pair<librevenge::RVNGString,QuattroFormulaInternal::CellReference> >m_idToFieldMap;
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
	//! map zone id to zone name
	std::map<int, ZoneName> m_idToZoneNameMap;
	//! true if zone 341 is being read. Avoids recursion.
	bool m_readingZone341;
private:
	State(State const &)=delete;
	State &operator=(State const &)=delete;
};

bool State::getColor(int id, WPSColor &color) const
{
	if (m_colorsList.empty())
	{
		auto *THIS=const_cast<State *>(this);
		static const uint32_t quattroColorMap[]=
		{
			0xFFFFFF, 0xC0C0C0, 0x808080, 0x000000,
			0xFF0000, 0x00FF00, 0x0000FF, 0x00FFFF,
			0xFF00FF, 0xFFFF00, 0x800080, 0x000080,
			0x808000, 0x008000, 0x800000, 0x008080
		};
		for (uint32_t i : quattroColorMap)
			THIS->m_colorsList.push_back(WPSColor(quattroColorMap[i]));
	}
	if (id < 0 || id >= int(m_colorsList.size()))
	{
		WPS_DEBUG_MSG(("QuattroParserInternal::State::getColor(): unknown Quattro Pro color id: %d\n",id));
		return false;
	}
	color = m_colorsList[size_t(id)];
	return true;
}

void State::initZoneNameMap()
{
	if (!m_idToZoneNameMap.empty())
		return;
	m_idToZoneNameMap=std::map<int,ZoneName>
	{
		{ 0x0, ZoneName("File", "begin") },
		{ 0x1, ZoneName("File", "end") },
		{ 0x2, ZoneName("Recalculation", "mode") },
		{ 0x3, ZoneName("Recalculation", "order") },
		{ 0x6, ZoneName("Sheet", "size") },
		{ 0xb, ZoneName("FldName") },
		{ 0xc, ZoneName("Cell", "blank") },
		{ 0xd, ZoneName("Cell", "int") },
		{ 0xe, ZoneName("Cell", "float") },
		{ 0xf, ZoneName("Cell", "label") },
		{ 0x10, ZoneName("Cell", "formula") },
		{ 0x18, ZoneName("Range", "table") },
		{ 0x19, ZoneName("Range", "query") },
		{ 0x1a, ZoneName("Print", "block") },
		{ 0x1b, ZoneName("Range", "sort,block") },
		{ 0x1c, ZoneName("Range", "fill") },
		{ 0x1d, ZoneName("Range", "sort,firstKey") },
		{ 0x20, ZoneName("Range", "frequency") },
		{ 0x23, ZoneName("Range", "sort,secondKey") },
		{ 0x24, ZoneName("Protection") },
		{ 0x25, ZoneName("Print", "footer") },
		{ 0x26, ZoneName("Print", "header") },
		{ 0x27, ZoneName("Print", "setup") },
		{ 0x28, ZoneName("Print", "margins") },
		{ 0x2f, ZoneName("Recalculation", "iteration,count") },
		{ 0x30, ZoneName("Print", "pagebreak") },
		{ 0x33, ZoneName("Cell", "string,value") },
		{ 0x4b, ZoneName("Password", "data") },
		{ 0x4c, ZoneName("Password", "level") },
		{ 0x4d, ZoneName("System", "property") },
		{ 0x66, ZoneName("Range", "parse") },
		{ 0x67, ZoneName("Range", "regression") },
		{ 0x69, ZoneName("Range", "matrix") },
		{ 0x96, ZoneName("Column", "size") },
		{ 0x97, ZoneName("External", "link") },
		{ 0x98, ZoneName("External", "name") },
		{ 0x99, ZoneName("Macro", "library") },
		// 0x9d Cell style (never seen)
		// 0x9e Conditional border color (never seen)
		{ 0x9f, ZoneName("Range", "sort,thirdKey") },
		{ 0xa0, ZoneName("Range", "sort,fourstKey") },
		{ 0xa1, ZoneName("Range", "sort,fifthKey") },
		{ 0xb7, ZoneName("Range", "solve for") },
		{ 0xc9, ZoneName("Version") },
		{ 0xca, ZoneName("Sheet", "begin") },
		{ 0xcb, ZoneName("Sheet", "end") },
		{ 0xcc, ZoneName("Sheet", "name") },
		{ 0xce, ZoneName("Cell", "style") },
		{ 0xcf, ZoneName("FontDef") },
		{ 0xd0, ZoneName("StyleName") },
		{ 0xd1, ZoneName("Sheet", "attribute") },
		{ 0xd2, ZoneName("Pane", "row,default") },
		{ 0xd3, ZoneName("Pane", "row2,default") },
		{ 0xd4, ZoneName("Pane", "col,default") },
		{ 0xd5, ZoneName("Pane", "col2,default") },
		{ 0xd6, ZoneName("Pane", "row") },
		{ 0xd7, ZoneName("Pane", "row2") },
		{ 0xd8, ZoneName("Pane", "col") },
		{ 0xd9, ZoneName("Pane", "col2") },
		{ 0xda, ZoneName("Pane", "font,max") },
		{ 0xdb, ZoneName("Pane", "font2,max") },
		{ 0xdc, ZoneName("Pane", "row,hidden") },
		{ 0xdd, ZoneName("Pane", "row2,hidden") },
		{ 0xde, ZoneName("Pane", "col,hidden") },
		{ 0xdf, ZoneName("Pane", "col2,hidden") },
		{ 0xe0, ZoneName("Pane", "style") },
		{ 0xe1, ZoneName("Pane", "style2") },
		{ 0xe2, ZoneName("PageGroup", "on") },
		{ 0xe3, ZoneName("PageGroup") },
		{ 0xe4, ZoneName("DLLIdFunct", "e5") },
		{ 0xe5, ZoneName("DLLIdFunct", "e6") },
		{ 0xe6, ZoneName("UserFormat") },
		{ 0xe7, ZoneName("Column", "def,attr") },
		{ 0xe8, ZoneName("ColorList") },
		{ 0xe9, ZoneName("Collection") }, // list of cell reference data, never seens
		{ 0xed, ZoneName("Print", "beg,names") },
		{ 0xee, ZoneName("Print", "formula") },
		{ 0xef, ZoneName("Print", "block,delimiter") },
		{ 0xf0, ZoneName("Print", "page,delimiter") },
		{ 0xf1, ZoneName("Print", "copies") },
		{ 0xf2, ZoneName("Print", "pages") },
		{ 0xf3, ZoneName("Print", "density") },
		{ 0xf4, ZoneName("Print", "tofit") },
		{ 0xf5, ZoneName("Print", "scaling") },
		{ 0xf6, ZoneName("Print", "paper,type") },
		{ 0xf7, ZoneName("Print", "orientation") },
		{ 0xf8, ZoneName("Print", "left,border") },
		{ 0xf9, ZoneName("Print", "top,border") },
		{ 0xfa, ZoneName("Print", "center,blocks") },
		{ 0xfb, ZoneName("Print", "end") },
		{ 0xfc, ZoneName("Print", "header,font") },

		{ 0x101, ZoneName("Print", "headings") },
		{ 0x102, ZoneName("Print", "gridlines") },
		{ 0x103, ZoneName("Optimizer") },
		{ 0x104, ZoneName("Optimizer", "constraint") },
		{ 0x105, ZoneName("Pane", "row,range") },
		{ 0x106, ZoneName("Pane", "row2,range") },
		{ 0x107, ZoneName("Pane", "font,max,range") },
		{ 0x108, ZoneName("Pane", "font2,max,range") },
		{ 0x109, ZoneName("Print", "beg,record") },
		{ 0x10a, ZoneName("Print", "beg,graph") },
		{ 0x10c, ZoneName("Print", "draft,margins") },
		{ 0x10d, ZoneName("Show", "compatible") },
		{ 0x110, ZoneName("Print", "footer,font") },
		{ 0x111, ZoneName("Print", "area") },

		{ 0x12e, ZoneName("Object", "number") },
		{ 0x12f, ZoneName("Query", "table,command") },
		{ 0x132, ZoneName("Formula", "compile") },
		{ 0x133, ZoneName("Formula", "audit") },
		{ 0x134, ZoneName("Sheet", "tab,color") },
		{ 0x135, ZoneName("Sheet", "zoom") },
		{ 0x136, ZoneName("Show", "notebook,object") },
		{ 0x137, ZoneName("Sheet", "protection") },

		{ 0x154, ZoneName("UserFormat", "complete") },

		{ 0x191, ZoneName("View", "begin") },
		{ 0x192, ZoneName("View", "end") },
		{ 0x193, ZoneName("View", "window") },
		{ 0x194, ZoneName("View", "location") },
		{ 0x195, ZoneName("View", "split") },
		{ 0x196, ZoneName("View", "synchronize") },
		{ 0x197, ZoneName("View", "pane,info") },
		{ 0x198, ZoneName("View", "pane2,info") },
		{ 0x199, ZoneName("View", "page") },
		{ 0x19a, ZoneName("View", "page2") },
		{ 0x19b, ZoneName("View", "current") },
		{ 0x19c, ZoneName("View", "display,settings") },
		{ 0x19e, ZoneName("View", "zoom") },

		{ 0x259, ZoneName("Graph", "begin,name") },
		{ 0x25a, ZoneName("Graph", "end") },
		{ 0x25d, ZoneName("Graph", "icon,coord") },
		{ 0x25e, ZoneName("Slide", "begin") },
		{ 0x25f, ZoneName("Slide", "end") },
		{ 0x260, ZoneName("Slide", "icon,coord") },
		{ 0x262, ZoneName("Slide", "time") },
		{ 0x263, ZoneName("Slide", "effect0") },
		{ 0x264, ZoneName("Graph", "version") },
		{ 0x265, ZoneName("Slide", "speed") },
		{ 0x266, ZoneName("Slide", "effect1") },
		{ 0x267, ZoneName("Slide", "level") },
		{ 0x26a, ZoneName("Slide", "type") },
		{ 0x26b, ZoneName("Slide", "comment") },
		{ 0x26c, ZoneName("Slide", "master,name") },
		{ 0x2bc, ZoneName("Graph", "beg,record") },
		{ 0x2bd, ZoneName("Chart", "beg,serie") },
		{ 0x2be, ZoneName("Chart", "end,serie") },
		{ 0x2bf, ZoneName("Serie", "Xlabel") },
		{ 0x2c0, ZoneName("Serie", "Zlabel") },
		{ 0x2c1, ZoneName("Serie", "legend") },
		{ 0x2c2, ZoneName("Serie", "number") },
		{ 0x2c3, ZoneName("Serie", "beg,data") },
		{ 0x2c4, ZoneName("Serie", "end,data") },
		{ 0x2c6, ZoneName("Serie", "data") },
		{ 0x2c7, ZoneName("Serie", "label") },
		{ 0x2c8, ZoneName("Serie", "legend") },
		{ 0x2c9, ZoneName("Chart", "beg,record") },
		{ 0x2ca, ZoneName("Chart", "end,record") },
		{ 0x2cb, ZoneName("Graph", "extension") },
		{ 0x2cd, ZoneName("Chart", "beg,save") },
		{ 0x2ce, ZoneName("Chart", "end,save") },
		{ 0x2db, ZoneName("Graph", "display,order") },
		{ 0x2dc, ZoneName("Serie", "extension") },

		{ 0x31f, ZoneName("Graph", "end,record") },
		{ 0x321, ZoneName("Object", "begin") },
		{ 0x322, ZoneName("Object", "end") },

		{ 0x335, ZoneName("GrDialog","textbox")},
		{ 0x337, ZoneName("GrDialog","37")},
		{ 0x33e, ZoneName("GrRect", "circle")},
		{ 0x33f, ZoneName("GrDialog","3f")},
		{ 0x342, ZoneName("GrDialog","42")},
		{ 0x343, ZoneName("GrDialog","button")},
		{ 0x345, ZoneName("GrDialog","bitmap")},
		{ 0x349, ZoneName("GrDialog","49")},
		{ 0x34a, ZoneName("GrDialog","4a")},
		{ 0x34e, ZoneName("GrDialog","4e")},
		{ 0x34f, ZoneName("GrDialog","4f")},
		{ 0x351, ZoneName("GrDialog","51")},
		{ 0x35a, ZoneName("GrLine")},
		{ 0x35b, ZoneName("GrPolygon")},
		{ 0x35c, ZoneName("GrPolygon","line")},
		{ 0x35d, ZoneName("GrDialog","5d")},
		{ 0x35e, ZoneName("GrDialog","main")},
		{ 0x364, ZoneName("GrRect")},
		{ 0x36d, ZoneName("GrDialog","6d")},
		{ 0x36f, ZoneName("GrTextBox")},
		{ 0x379, ZoneName("GrRect", "round")},
		{ 0x37b, ZoneName("GrLine", "arrow")},
		{ 0x37c, ZoneName("GrPolygon","line,bezier")},

		{ 0x381, ZoneName("Object", "frame,ole") },
		{ 0x382, ZoneName("Object", "image") },
		{ 0x383, ZoneName("Object", "bitmap") },
		{ 0x384, ZoneName("Object", "chart") },
		{ 0x385, ZoneName("Object", "frame") },
		{ 0x386, ZoneName("Object", "button") },
		{ 0x388, ZoneName("GrPolygon","bezier")},
		{ 0x38b, ZoneName("Object", "ole") },

		{ 0x4d3, ZoneName("Object", "shape") }
	};
}

}

// constructor, destructor
QuattroParser::QuattroParser(RVNGInputStreamPtr &input, WPSHeaderPtr &header,
                             libwps_tools_win::Font::Type encoding, char const *password)
	: WKSParser(input, header)
	, m_listener()
	, m_state(new QuattroParserInternal::State(encoding, password))
	  //, m_chartParser(new QuattroChart(*this))
	, m_graphParser(new QuattroGraph(*this))
	, m_spreadsheetParser(new QuattroSpreadsheet(*this))
{
}

QuattroParser::~QuattroParser()
{
}

int QuattroParser::version() const
{
	return m_state->m_version;
}

libwps_tools_win::Font::Type QuattroParser::getDefaultFontType() const
{
	return m_state->getDefaultFontType();
}

bool QuattroParser::getExternalFileName(int fId, librevenge::RVNGString &fName) const
{
	auto it = m_state->m_idToExternalFileMap.find(fId);
	if (it!=m_state->m_idToExternalFileMap.end())
	{
		fName=it->second;
		return true;
	}
	WPS_DEBUG_MSG(("QuattroParser::getExternalFileName: can not find %d name\n", fId));
	return false;
}

bool QuattroParser::getField(int fId, librevenge::RVNGString &text,
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
		WPS_DEBUG_MSG(("QuattroParser::getField: can not find %d name\n", fId&0xbfff));
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
	WPS_DEBUG_MSG(("QuattroParser::getField: can not find %d field\n", fId));
	return false;
}

//////////////////////////////////////////////////////////////////////
// interface with QuattroGraph
//////////////////////////////////////////////////////////////////////

bool QuattroParser::sendPageGraphics(int sheetId) const
{
	return m_graphParser->sendPageGraphics(sheetId);
}

bool QuattroParser::sendGraphics(int sheetId, Vec2i const &cell) const
{
	return m_graphParser->sendGraphics(sheetId, cell);
}


//////////////////////////////////////////////////////////////////////
// interface with QuattroSpreadsheet
//////////////////////////////////////////////////////////////////////

Vec2f QuattroParser::getCellPosition(int sheetId, Vec2i const &cell) const
{
	return m_spreadsheetParser->getPosition(sheetId, cell);
}

bool QuattroParser::getColor(int id, WPSColor &color) const
{
	return m_state->getColor(id, color);
}

bool QuattroParser::getFont(int id, WPSFont &font, libwps_tools_win::Font::Type &type) const
{
	if (id < 0 || id>=int(m_state->m_fontsList.size()))
	{
		WPS_DEBUG_MSG(("QuattroParser::getFont: can not find font %d\n", id));
		return false;
	}
	auto const &ft=m_state->m_fontsList[size_t(id)];
	font=ft;
	type=ft.m_type;
	return true;
}

// main function to parse the document
void QuattroParser::parse(librevenge::RVNGSpreadsheetInterface *documentInterface)
{
	RVNGInputStreamPtr input=getInput();
	if (!input)
	{
		WPS_DEBUG_MSG(("QuattroParser::parse: does not find main ole\n"));
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
			//m_chartParser->setListener(m_listener);
			m_graphParser->setListener(m_listener);
			m_spreadsheetParser->setListener(m_listener);

			m_graphParser->updateState();
			m_spreadsheetParser->updateState();

			m_listener->startDocument();
			int numSheet=m_spreadsheetParser->getNumSpreadsheets();
			if (numSheet==0) ++numSheet;
			for (int i=0; i<numSheet; ++i)
				m_spreadsheetParser->sendSpreadsheet(i, m_graphParser->getGraphicCellsInSheet(i));
			m_listener->endDocument();
			m_listener.reset();
			ok = true;
		}
	}
	catch (...)
	{
		WPS_DEBUG_MSG(("QuattroParser::parse: exception catched when parsing MN0\n"));
		throw (libwps::ParseException());
	}

	ascii().reset();
	if (!ok)
		throw(libwps::ParseException());
}

std::shared_ptr<WKSContentListener> QuattroParser::createListener(librevenge::RVNGSpreadsheetInterface *interface)
{
	std::vector<WPSPageSpan> pageList;
	WPSPageSpan ps(m_state->m_pageSpan);
	int numSheet=m_spreadsheetParser->getNumSpreadsheets();
	if (numSheet<=0) numSheet=1;
	if (!m_state->m_headerString.empty())
	{
		WPSSubDocumentPtr subdoc(new QuattroParserInternal::SubDocument
		                         (getInput(), *this, true));
		ps.setHeaderFooter(WPSPageSpan::HEADER, WPSPageSpan::ALL, subdoc);
	}
	if (!m_state->m_footerString.empty())
	{
		WPSSubDocumentPtr subdoc(new QuattroParserInternal::SubDocument
		                         (getInput(), *this, false));
		ps.setHeaderFooter(WPSPageSpan::FOOTER, WPSPageSpan::ALL, subdoc);
	}
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
bool QuattroParser::checkHeader(WPSHeader *header, bool strict)
{
	m_state.reset(new QuattroParserInternal::State(m_state->m_fontType, m_state->m_password));
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

bool QuattroParser::checkHeader(std::shared_ptr<WPSStream> stream, bool strict)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	if (!stream || !stream->checkFilePosition(12))
	{
		WPS_DEBUG_MSG(("QuattroParser::checkHeader: file is too short\n"));
		return false;
	}

	input->seek(0,librevenge::RVNG_SEEK_SET);
	auto firstOffset = int(libwps::readU8(input));
	auto type = int(libwps::read8(input));
	f << "FileHeader:";
	if (firstOffset == 0 && type == 0)
		m_state->m_version=1000;
	else
	{
		WPS_DEBUG_MSG(("QuattroParser::checkHeader: find unexpected first data\n"));
		return false;
	}
	auto val=int(libwps::read16(input));
	if (val==2)
	{
		// version
		val=int(libwps::readU16(input));
		if (val==0x1001)
		{
			m_state->m_version=1001;
			f << "quattropro[wb1],";
		}
		else if (val==0x1002)
		{
			m_state->m_version=1002;
			f << "quattropro[wb2],";
		}
		else if (val==0x1007)
		{
			m_state->m_version=1003;
			f << "quattropro[wb3],";
		}
		else
		{
			WPS_DEBUG_MSG(("QuattroParser::checkHeader: find unknown file version\n"));
			return false;
		}
	}
	else
	{
		WPS_DEBUG_MSG(("QuattroParser::checkHeader: header contain unexpected size field data\n"));
		return false;
	}
	input->seek(0, librevenge::RVNG_SEEK_SET);
	if (strict)
	{
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

bool QuattroParser::readZones()
{
	int const vers=version();
	m_graphParser->cleanState();
	m_spreadsheetParser->cleanState();
	m_state->initZoneNameMap();

	std::shared_ptr<WPSStream> stream(new WPSStream(getInput(), ascii()));
	RVNGInputStreamPtr &input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	input->seek(0, librevenge::RVNG_SEEK_SET);
	while (true)
	{
		if (!stream->checkFilePosition(input->tell()+4))
			break;
		if (!readZone(stream))
			break;
		if (m_state->m_isEncrypted && !m_state->m_isDecoded)
			throw(libwps::PasswordException());
	}

	//
	// look for ending
	//
	long pos = input->tell();
	if (!stream->checkFilePosition(pos+4))
	{
		WPS_DEBUG_MSG(("QuattroParser::readZones: cell header is too short\n"));
		return m_spreadsheetParser->getNumSpreadsheets()>0;
	}
	auto type = int(libwps::readU16(input)); // 1
	auto length = int(libwps::readU16(input));
	if (length)
	{
		WPS_DEBUG_MSG(("QuattroParser::readZones: parse breaks before ending\n"));
		ascFile.addPos(pos);
		ascFile.addNote("Entries(BAD):###");
		return m_spreadsheetParser->getNumSpreadsheets()>0;
	}

	ascFile.addPos(pos);
	if (type != 1)
	{
		WPS_DEBUG_MSG(("QuattroParser::readZones: odd end cell type: %d\n", type));
		ascFile.addNote("Entries(BAD):###");
		return m_spreadsheetParser->getNumSpreadsheets();
	}
	ascFile.addNote("Entries(EndSpreadsheet)");

	// checkme: crypted .wb3 files also contain an OLE zone, but it seems empty...
	if (vers==1002 || (vers==1003 && m_state->m_isEncrypted))
		readOLEZones(stream);
	if (vers==1003)
	{
		// object in PerfectOffice_OBJECTS/_1507007992_c/, ... *
		parseOLEStream(getFileInput(), "PerfectOffice_MAIN");
	}
	return m_spreadsheetParser->getNumSpreadsheets();
}

bool QuattroParser::parseOLEStream(RVNGInputStreamPtr input, std::string const &avoid)
{
	if (!input || !input->isStructured())
	{
		WPS_DEBUG_MSG(("QuattroParser::parseOLEStream: oops, can not find the input stream\n"));
		return false;
	}
	std::map<std::string,size_t> dirToIdMap;
	WPSOLEParser oleParser(avoid, getDefaultFontType(),
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
		for (int wh=0; wh<2; ++wh)
		{
			std::string name=it.first+"/"+(wh==0 ? "LinkInfo" : "BOlePart");
			RVNGInputStreamPtr cOle(input->getSubStreamByName(name.c_str()));
			if (!cOle)
			{
				WPS_DEBUG_MSG(("QuattroParser::parseOLEStream: oops, can not find link info for dir %s\n", name.c_str()));
				continue;
			}
			libwps::DebugFile asciiFile(cOle);
			asciiFile.open(libwps::Debug::flattenFileName(name));
			if (wh==1)
				readOleBOlePart(std::make_shared<WPSStream>(cOle,asciiFile));
			else
			{
				librevenge::RVNGString linkName;
				if (readOleLinkInfo(std::make_shared<WPSStream>(cOle,asciiFile),linkName) && !linkName.empty())
				{
					if (objectMap.find(int(it.second))==objectMap.end())
					{
						WPS_DEBUG_MSG(("QuattroParser::parseOLEStream: oops, can not find embedded data for %s\n", name.c_str()));
						continue;
					}
					nameToObjectsMap[linkName]=objectMap.find(int(it.second))->second;
				}
			}
		}
	}
	if (!nameToObjectsMap.empty())
		m_graphParser->storeObjects(nameToObjectsMap);
	return true;
}

bool QuattroParser::readOLEZones(std::shared_ptr<WPSStream> &stream)
{
	if (!stream)
		return false;
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	f << "Entries(OLEData)[header]:";
	long pos = input->tell();
	long endPos = stream->m_eof;
	if (!stream->checkFilePosition(pos+18))
	{
		WPS_DEBUG_MSG(("QuattroParser::readOLEZones: the zone seems to short\n"));
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return false;
	}
	for (int i=0; i<4; ++i)   // 0
	{
		auto val=int(libwps::read16(input));
		if (val) f << "f" << i << "=" << val << ",";
	}
	auto sSz=long(libwps::readU32(input));
	librevenge::RVNGString text; // QPW$ExtendedStorage$6.0
	if (sSz<=0 || sSz > endPos-input->tell()-6 || !readCString(stream,text,sSz))
	{
		WPS_DEBUG_MSG(("QuattroParser::readOLEZones: can not read header's type\n"));
		f << "##sSz,";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	f << "type=" << text.cstr() << ",";
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	while (input->tell()+6<=endPos)
	{
		pos=input->tell();
		f.str("");
		f << "OLEData:";
		auto type=int(libwps::read16(input));
		sSz=long(libwps::readU32(input));
		if (sSz < 0 || sSz > endPos-pos-6 || type<1 || type>2 || (sSz==0 && type==2))
		{
			input->seek(pos, librevenge::RVNG_SEEK_SET);
			break;
		}
		if (type==1)
		{
			if (sSz)
				f << "###sz=" << sSz << ",";
			f << "end,";
			ascFile.addPos(pos);
			ascFile.addNote(f.str().c_str());
			return true;
		}
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		unsigned long numRead;
		const unsigned char *data=input->read(static_cast<unsigned long>(sSz), numRead);
		if (data && long(numRead)==sSz)
		{
			auto ole=libwps_OLE::getOLEInputStream(std::make_shared<WPSStringStream>(data, unsigned(numRead)));
			if (!ole)
			{
				WPS_DEBUG_MSG(("QuattroParser::readOLEZones::readOLE: oops, can not decode the ole\n"));
			}
			else
			{
				/* normally /_Date_XXX/ Where Date~1507005964 and XXX is an hexadecimal data.
				   Moreever, a file can be saved many times, so must read /_Date_XXX/ to retrieve the correspondance
				 */
				ascFile.skipZone(pos+6, pos+6+sSz-1);
				parseOLEStream(ole);
			}
		}
		else
		{
			WPS_DEBUG_MSG(("QuattroParser::readOLEZones::readOLE: I can not find the data\n"));
			input->seek(pos, librevenge::RVNG_SEEK_SET);
			break;
		}
		input->seek(pos+6+sSz, librevenge::RVNG_SEEK_SET);
	}
	if (input->tell()<endPos)
	{
		WPS_DEBUG_MSG(("QuattroParser::readOLEZones: find extra data\n"));
		ascFile.addPos(input->tell());
		ascFile.addNote("OLEData:###extra");
	}
	return true;
}

bool QuattroParser::readZone(std::shared_ptr<WPSStream> &stream)
{
	if (!stream)
		return false;
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto id = int(libwps::readU16(input));
	auto sz = long(libwps::readU16(input));
	if (sz<0 || !stream->checkFilePosition(pos+4+sz))
	{
		WPS_DEBUG_MSG(("QuattroParser::readZone: size is bad\n"));
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}

	if (id&0x8000)
	{
		// very rare, unsure what this means ; is it possible to have
		// other flags here ?
		WPS_DEBUG_MSG(("QuattroParser::readZone: find type[8] flags\n"));
		ascFile.addPos(pos);
		ascFile.addNote("#flag8000,");
		id &= 0x7fff;
	}

	if (id>=0x800) // I never found anything biggest than 47d, so must be ok
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}

	if (sz>=0xFF00 && stream->checkFilePosition(pos+4+sz+4))
	{
		input->seek(pos+4+sz, librevenge::RVNG_SEEK_SET);
		if (libwps::readU16(input)==0x10f)
		{
			// incomplete block, we must rebuild it
			input->seek(pos, librevenge::RVNG_SEEK_SET);
			unsigned long numRead;
			const unsigned char *data=input->read(4+static_cast<unsigned long>(sz), numRead);
			if (data && long(numRead)==4+sz)
			{
				ascFile.skipZone(pos,pos+4+sz-1);
				auto newInput=std::make_shared<WPSStringStream>(data, unsigned(numRead));
				bool ok=true;
				while (true)
				{
					long actPos=input->tell();
					if (!stream->checkFilePosition(actPos+4) || libwps::readU16(input)!=0x10f)
					{
						input->seek(actPos, librevenge::RVNG_SEEK_SET);
						break;
					}
					auto extraSize=long(libwps::readU16(input));
					if (!stream->checkFilePosition(actPos+4+extraSize))
					{
						ok=false;
						break;
					}
					ascFile.addPos(actPos);
					ascFile.addNote("Entries(ExtraData):");
					if (!extraSize)
						break;
					data=input->read(static_cast<unsigned long>(extraSize), numRead);
					if (!data || long(numRead)!=extraSize)
					{
						ok=false;
						break;
					}
					newInput->append(data, unsigned(numRead));
					ascFile.skipZone(actPos+4,actPos+4+extraSize-1);
				}
				if (ok)
				{
					std::stringstream s;
					static int complexDataNum=0;
					s << "Data" << ++complexDataNum;
					auto newStream=std::make_shared<WPSStream>(newInput);
					newStream->m_ascii.open(s.str());
					newStream->m_ascii.setStream(newInput);
					readZone(newStream);
					return true;
				}
			}
			WPS_DEBUG_MSG(("QuattroParser::readZone: can not reconstruct a zone\n"));
			ascFile.addPos(pos);
			ascFile.addNote("Entries(###Bad):");
			input->seek(pos+4+sz, librevenge::RVNG_SEEK_SET);
			return true;
		}
	}
	auto zIt=m_state->m_idToZoneNameMap.find(id);
	if (zIt==m_state->m_idToZoneNameMap.end())
		f << "Entries(Zone" << std::hex << id << std::dec << "A):";
	else if (zIt->second.m_extra.empty())
		f << "Entries(" << zIt->second.m_name << "):";
	else
		f << "Entries(" << zIt->second.m_name << ")[" << zIt->second.m_extra << "]:";
	if (id>1)
	{
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
	}
	f.str("");

	bool ok = true, isParsed = false, needWriteInAscii = false;
	int val;
	input->seek(pos, librevenge::RVNG_SEEK_SET);
	switch (id)
	{
	case 0:
		if (sz!=2) break;
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		f << "version=" << std::hex << libwps::readU16(input) << std::dec << ",";
		isParsed=needWriteInAscii=true;
		break;
	case 0x1: // EOF
		ok = false;
		break;

	// no data
	case 0xfb: // print end record/end
	case 0x191:
	case 0x192:
	case 0x25a:
	case 0x25b:
	case 0x25c:
	case 0x25f:
	case 0x2bc:
	case 0x2bd:
	case 0x2be:
	case 0x2c3:
	case 0x2c4:
	case 0x2c9:
	case 0x2ca:
	case 0x2cd:
	case 0x2ce:
	case 0x31f:
		if (sz!=0) break;
		isParsed=needWriteInAscii=true;
		break;
	// boolean
	case 0x2: // calculation mode 0: manual, 1:background, FF:auto
	case 0x3: // calculation order 0: default, 1:column, ff:row
	case 0x24: // protected if FF
	case 0x2f: // number of iteration
	case 0x30: // print page break 0=yes, 1=no
	case 0x99: // 0: no macro, 1: macro library
	case 0xe2: // 0: off, 1: on
	case 0xee: // 0: print results, 1: print formula
	case 0xf3: // 0: draft, 1: final
	case 0xf4: // 0: no, 1: yes, sometimes with size 5
	case 0xf7: // 0: landscape, 1: portrait
	case 0xfa: // 0: no centering, 1: print centered
	case 0x101: // 0: do not print headings, 1: print headings
	case 0x102: // 0: do not print gridlines, 1: print them
	case 0x109: // in fact, a name always empty
	case 0x10a: // in fact, a name always empty
	case 0x111: // 0: current page, 1: notebook, 2: block selection
	case 0x132: // 0: no formula compiled, 1: formula compiled
	case 0x133: // 0: no audit display, 1: display audit
	case 0x137: // 0: disabled, 1: enabled
	case 0x196: // 0: synchronized, 2: not synchronized
	case 0x19b: // 1: pane 1, 2: pane 2
		if (sz==0)
		{
			f << "##";
			break;
		}
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		val=int(libwps::readU8(input));
		if (val==1) f << "true,";
		else if (val) f << "#val=" << val << ",";
		if (id==0xf4 && sz==5)
		{
			for (int i=0; i<2; ++i)
			{
				val=int(libwps::readU16(input));
				if (val!=1) f << "f" << i << "=" << val << ",";
			}
		}
		isParsed=needWriteInAscii=true;
		break;
	case 0x4c: // with 0: none, 1:low, 2:medium, 3:high
	case 0x4d: // with 0
	case 0xc9: // with [026]{00,01,c8}
	case 0xe0: // default panel style
	case 0xe1: // default panel2 style
	case 0xf1: // number of copies
	case 0xf5: // with scaling in %
	case 0xf6: // paper type: 0|1|9|101
	case 0x12d: // with 1
	case 0x136: // 0: show all objects, 1: show outline objects, 2: hidde object
	case 0x19e: // zoom factor
	case 0x262: // time
	case 0x263: // slide effect
	case 0x265: // speed
	case 0x266: // slide special effect
	case 0x26a: // 0: small, 1: medium, 2: large, 3: list of name
	case 0x2c2: // number of series
		if (sz!=2)
		{
			f << "##";
			break;
		}
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		val=int(libwps::readU16(input));
		if (val) f << "f0=" << std::hex << val << std::dec << ",";
		break;
	case 0xe7: // f0=column, f1>>3=def attritude
	case 0xef: // f0=number of line, f1:0=specified number of line, 1=page feed
	case 0xf0: // f0=number of line, f1:0=specified number of line, 1=page feed
	case 0x12e: // f0-f1=number
	case 0x25d: // f0=x, f1=y
	case 0x260: // f0=x, f1=y
	case 0x264: // version
		if (sz!=4)
		{
			f << "##";
			break;
		}
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		for (int i=0; i<2; ++i)
		{
			val=int(libwps::readU16(input));
			if (val) f << "f" << i << "=" << std::hex << val << std::dec << ",";
		}
		break;
	case 0xf2: // f0=min page, f1=max page, f2:0=print range, 1=print all
		if (sz!=6)
		{
			f << "##";
			break;
		}
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		for (int i=0; i<3; ++i)
		{
			val=int(libwps::readU16(input));
			if (val) f << "f" << i << "=" << std::hex << val << std::dec << ",";
		}
		break;
	// not basic
	case 0x6: // active worksheet range
		ok = m_spreadsheetParser->readSheetSize(stream);
		isParsed = true;
		break;
	case 0xb: // named range
		readFieldName(stream);
		isParsed=true;
		break;
	case 0xc: // blank cell
	case 0xd: // integer cell
	case 0xe: // floating cell
	case 0xf: // label cell
	case 0x10: // formula cell
	case 0x33:  // value of string formula
		// case 10e:  seems relative to cell with formula: list of dependency? can have some text
		ok = m_spreadsheetParser->readCell(stream);
		isParsed=true;
		break;

	case 0x18: // data table range
	case 0x19: // query range
	case 0x1a: // print block
	case 0x1b: // sort range
	case 0x1c: // fill range
	case 0x1d: // primary sort key range
	case 0x20: // distribution range
	case 0x23: // secondary sort key range
	case 0x66: // parse
	case 0x67: // regression
	case 0x69: // matrix
	case 0x9f: // sort third key
	case 0xa0: // sort four key
	case 0xa1: // sort fifth key
	case 0xb7: // solve for
	case 0xf8: // left border
	case 0xf9: // top border
	case 0x10d: //compatible slide show
	case 0x2bf: // X axis label
	case 0x2c0: // Z axis label
	case 0x2c1: // legend
	case 0x2c6: // serie data
	case 0x2c7: // serie legend
		readBlockList(stream);
		isParsed=true;
		break;
	// checkme this nodes appear two times, even/odd page ?
	case 0x25: // footer
	case 0x26: // header
		readHeaderFooter(stream, id==0x26);
		isParsed = true;
		break;
	case 0x27: // print setup
		if (sz<1) break;
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		val=int(libwps::readU8(input));
		if (val+1<sz) break;
		// now data to send to the printer
		break;
	case 0x28: // print margin: one by spreadsheet
		if (sz!=12) break;
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		f << "margs=[";
		for (int i=0; i<4; ++i)   // LRTB
			f << float(libwps::read16(input))/20.f << ",";
		f << "],";
		f << "hf[height]=[";
		for (int i=0; i<2; ++i)   // header, footer height
			f << float(libwps::read16(input))/20.f << ",";
		f << "],";
		isParsed=needWriteInAscii=true;
		break;

	case 0x4b:
		m_state->m_isEncrypted=true;
		if (sz==20)
		{
			m_state->m_isEncrypted=true;
			input->seek(pos+4, librevenge::RVNG_SEEK_SET);
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
			WPS_DEBUG_MSG(("QuattroParser::parse: can not decode the file\n"));
		}
		break;
	case 0x96:
		readCellPosition(stream);
		isParsed=true;
		break;
	case 0x97: // name use to refer to some external spreadsheet
	case 0x98: // external name
		readExternalData(stream);
		isParsed=true;
		break;

	case 0xca:
	case 0xcb:
		isParsed=m_spreadsheetParser->readBeginEndSheet(stream, m_state->m_actualSheet);
		break;
	case 0xcc:
		isParsed=m_spreadsheetParser->readSheetName(stream);
		break;
	case 0xce:
		ok = m_spreadsheetParser->readCellStyle(stream);
		isParsed = true;
		break;
	case 0xcf:
	case 0xfc:
	case 0x110:
		isParsed=readFontDef(stream);
		break;
	case 0xd0:
		isParsed=readStyleName(stream);
		break;
	case 0xd1:
		isParsed=readPaneAttribute(stream);
		break;
	case 0xd6:
	case 0xd7:
		isParsed = m_spreadsheetParser->readRowSize(stream);
		break;
	case 0xd8:
	case 0xd9:
		isParsed = m_spreadsheetParser->readColumnSize(stream);
		break;
	case 0xd2: // default row height
	case 0xd3:
	case 0xd4: // default col width
	case 0xd5:
		isParsed=m_spreadsheetParser->readColumnRowDefaultSize(stream);
		break;
	case 0xda:
	case 0xdb:
		if (sz!=3 && sz!=4)
		{
			f << "##";
			break;
		}
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		f << "row=" << libwps::readU16(input) << ",";
		f << "F" << int(libwps::readU8(input)) << ",";
		if (sz==4)
		{
			val=int(libwps::readU8(input)); // 0, 5b, da
			if (val) f << "f0=" << std::hex << val << std::dec << ",";
		}
		break;
	case 0xdc:
	case 0xdd:
	case 0xde:
	case 0xdf:
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		f << "hidden=[";
		for (long i=0; i<sz; ++i)
		{
			val=int(libwps::readU8(input));
			if (val==0) continue;
			for (int d=0, b=1; d<8; ++d, b<<=1)   // check byte order
			{
				if (val&b) f << 8*i+d << ",";
			}
		}
		f << "],";
		break;
	case 0xe3:
	case 0xe4:
	case 0xe5:
	case 0xe6: // user format, FIXME: parse me and use me
	{
		if (sz<3) break;
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		val=int(libwps::readU16(input));
		if (id==0xe3)
			f << "sheet" << (val&0xff) << "=>" << (val>>8) << ",";
		else
			f << "id=" << val << ",";
		librevenge::RVNGString text;
		if (!readCString(stream, text,sz-2))
			f << "###";
		else
		{
			if (id==0xe4 || id==0xe5)
				m_spreadsheetParser->addDLLIdName(val, text, id==0xe4);
			else if (id==0xe6)
				m_spreadsheetParser->addUserFormat(val, text);
			f << text.cstr() << ",";
		}
		break;
	}
	case 0xe8:
		isParsed=readColorList(stream);
		break;
	case 0xed:
	case 0x259:
	case 0x25e:
	case 0x261:
	case 0x26b:
	case 0x26c:
	case 0x2c8:
	{
		if (sz<1) break;
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		librevenge::RVNGString text;
		if (!readCString(stream, text,sz))
			f << "###";
		else
			f << text.cstr() << ",";
		break;
	}
	// case 0xfe: some counter
	case 0x103:
		isParsed=readOptimizer(stream);
		break;
	// 0x104: optimizer constraint: readme
	case 0x105:
	case 0x106:
		isParsed=m_spreadsheetParser->readRowRangeSize(stream);
		break;
	case 0x107: // never seems
	case 0x108:
		if (sz!=5)
		{
			f << "##";
			break;
		}
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		f << "rows=" << libwps::readU16(input) << ",";
		f << "x" << libwps::readU16(input) << ",";
		f << "F" << int(libwps::readU8(input)) << ",";
		break;
	case 0x10c: // draft margin: never seens
		if (sz!=12) break;
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		f << "margs=[";
		for (int i=0; i<4; ++i)   // LRTB
			f << float(libwps::read16(input))/20.f << ",";
		f << "],";
		f << "height=" << float(libwps::read16(input))/20.f << ",";
		f << "units=" << libwps::read16(input) << ","; // 0: char, 1: inches, 2: centimters
		isParsed=needWriteInAscii=true;
		break;
	case 0x12f:
		isParsed=readQueryCommand(stream);
		break;
	case 0x134:
	{
		if (sz!=4)
		{
			f << "###";
			break;
		}
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		unsigned char colors[4];
		for (auto &c : colors) c=static_cast<unsigned char>(libwps::readU8(input));
		f << WPSColor(colors[0],colors[1],colors[2]) << ",";
		break;
	}
	case 0x135:
	{
		if (sz!=4)
		{
			f << "###";
			break;
		}
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		int values[2];
		for (auto &v : values) v=int(libwps::readU16(input));
		if (values[0]==100)
			f << values[1] << "%,";
		else if (values[0]!=1 || values[1]!=1)
			f << values[1] << "/" << values[0] << ",";
		break;
	}
	case 0x193:
	{
		if (sz!=6)
		{
			f << "###";
			break;
		}
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		f << "size=" << libwps::readU16(input) << "x";
		f << libwps::readU16(input) << ",";
		f << "state=" << libwps::readU16(input) << ","; // 0:normal, 1:max, 2:min, 3:hidden
		break;
	}
	case 0x194:
	{
		if (sz!=4)
		{
			f << "###";
			break;
		}
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		f << "pos=" << libwps::read16(input) << "x";
		f << libwps::read16(input) << ",";
		break;
	}
	case 0x195:
	{
		if (sz!=6)
		{
			f << "###";
			break;
		}
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		f << "type=" << libwps::readU16(input) << ","; // 0: no split, 1: horizontal, 2:vertical
		f << "split=" << libwps::readU16(input) << "%,";
		f << libwps::readU16(input) << "%],";
		break;
	}
	case 0x197: // one by sheet
	case 0x198:
		isParsed=m_spreadsheetParser->readViewInfo(stream);
		break;
	case 0x19c:
	{
		if (sz!=8)
		{
			f << "###";
			break;
		}
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		for (int i=0; i<4; ++i)   // f0: group mode, f1: hori bar vis, f2: vert bar vis, f3: vis tabs
		{
			val=int(libwps::readU16(input));
			if (val!=1) f << "f" << i << "=" << val << ",";
		}
		break;
	}
	case 0x267:
		if (sz!=2)
		{
			f << "###";
			break;
		}
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		val=int(libwps::readU16(input));
		f << "ident=" << (val&0xff) << ",";
		if (val&0x100) f << "no[master],";
		if (val&0x1000) f << "skip[slide],";
		break;
	case 0x2db: // find with size 0 or 2
		if (sz==0) break;
		if (sz%2)
		{
			f << "###";
			break;
		}
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		for (long i=0; i<sz; ++i)
		{
			val=int(libwps::readU8(input));
			if (val) // small number 0-c
				f << "f" << i << "=" << std::hex << val << std::dec << ",";
		}
		break;
	case 0x2dc:
		isParsed=readSerieExtension(stream);
		break;
	case 0x321:
	case 0x322:
		isParsed=m_graphParser->readBeginEnd(stream, m_state->m_actualSheet);
		break;
	// case 0x324: a list of XX:20?: seems relative to graph/chart
	case 0x33e: // oval
		isParsed = m_graphParser->readRect(stream);
		break;
	case 0x341: // maybe chart ?
		isParsed = readZone341(stream);
		break;
	case 0x335:
	case 0x337:
	case 0x33f:
	case 0x342:
	case 0x343:
	case 0x345:
	case 0x349:
	case 0x34a:
	case 0x34e:
	case 0x34f:
	case 0x351:
	case 0x35d:
	case 0x36d:
		isParsed = m_graphParser->readDialogUnknown(stream);
		break;
	case 0x35a:
		isParsed = m_graphParser->readLine(stream);
		break;
	case 0x35b: // polygon
	case 0x35c: // polyline
	case 0x37c: // free polyline
	case 0x388: // free polygon
		isParsed = m_graphParser->readPolygon(stream);
		break;
	case 0x35e:
		isParsed = m_graphParser->readDialog(stream);
		break;
	case 0x364:
		isParsed = m_graphParser->readRect(stream);
		break;
	case 0x36f:
		isParsed = m_graphParser->readTextBox(stream);
		break;
	case 0x379: // round rect
		isParsed = m_graphParser->readRect(stream);
		break;
	case 0x37b: // arrow
		isParsed = m_graphParser->readLine(stream);
		break;
	case 0x381: // frame wb2
		isParsed = m_graphParser->readFrameOLE(stream);
		break;
	case 0x382: // only wb2?
		isParsed = m_graphParser->readImage(stream);
		break;
	case 0x383: // only wb2?
		isParsed = m_graphParser->readBitmap(stream);
		break;
	case 0x384:
		isParsed = m_graphParser->readChart(stream);
		break;
	case 0x385:
		isParsed = m_graphParser->readFrame(stream);
		break;
	case 0x386:
		isParsed = m_graphParser->readButton(stream);
		break;
	case 0x38b:
		isParsed = m_graphParser->readOLEData(stream);
		break;

	case 0x4d3: // wb2 and wb3
		isParsed = m_graphParser->readShape(stream);
		break;
	default:
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
			ascFile.addPos(pos);
			ascFile.addNote(f.str().c_str());
		}
		input->seek(pos+4+sz, librevenge::RVNG_SEEK_SET);
		return true;
	}

	if (sz && input->tell()!=pos && input->tell()!=pos+4+sz)
		ascFile.addDelimiter(input->tell(),'|');
	input->seek(pos+4+sz, librevenge::RVNG_SEEK_SET);
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

////////////////////////////////////////////////////////////
//   generic
////////////////////////////////////////////////////////////

bool QuattroParser::readCString(std::shared_ptr<WPSStream> stream, librevenge::RVNGString &string, long maxSize)
{
	RVNGInputStreamPtr input = stream->m_input;
	long pos = input->tell();
	string.clear();
	if (!stream->checkFilePosition(pos+maxSize))
	{
		WPS_DEBUG_MSG(("QuattroParser::readCString: string's size seems bad\n"));
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

bool QuattroParser::readFieldName(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	if (type != 0xb)
	{
		WPS_DEBUG_MSG(("QuattroParser::readFieldName: not a zoneB type\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	if (sz < 4)
	{
		WPS_DEBUG_MSG(("QuattroParser::readFieldName: size seems bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto id=int(libwps::readU16(input));
	f << "id=" << id << ",";
	auto val=int(libwps::readU8(input)); // always 1?
	bool hasRef=(val&1);
	if ((val&1)==0) f << "no[ref],";
	if (val&2) f << "deleted,";
	// val&4 must be zero, other reserved
	librevenge::RVNGString name;
	auto sSz=int(libwps::readU8(input));
	if (4+sSz+(hasRef ? 6 : 0)>sz || !readCString(stream,name,sSz))
	{
		f << "##sSz,";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	else if (!name.empty())
		f << name.cstr() << ',';
	input->seek(pos+4+4+sSz, librevenge::RVNG_SEEK_SET);
	if (!hasRef)
	{
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	QuattroFormulaInternal::CellReference instr;
	if (!m_spreadsheetParser->readCellReference(stream, endPos, instr))
	{
		WPS_DEBUG_MSG(("QuattroParser::readFieldName: can not read some reference\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	else if (!instr.empty())
	{
		f << instr;
		if (m_state->m_idToFieldMap.find(id)!=m_state->m_idToFieldMap.end())
		{
			WPS_DEBUG_MSG(("QuattroParser::readFieldName: oops a field with id=%d already exists\n", id));
		}
		else
			m_state->m_idToFieldMap[id]=std::pair<librevenge::RVNGString,QuattroFormulaInternal::CellReference>(name, instr);
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroParser::readExternalData(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type != 0x97 && type != 0x98)
	{
		WPS_DEBUG_MSG(("QuattroParser::readExternalData: not a font zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	if (sz<3)
	{
		WPS_DEBUG_MSG(("QuattroParser::readExternalData: seems very short\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto id=int(libwps::readU16(input));
	f << "id=" << id << ",";
	librevenge::RVNGString name;
	auto &map=type==0x98? m_state->m_idToExternalNameMap : m_state->m_idToExternalFileMap;
	if (!readCString(stream,name,sz-2))
	{
		f << "##name,";
	}
	else if (map.find(id)!=map.end())
	{
		WPS_DEBUG_MSG(("QuattroParser::readExternalData: id=%d already found\n", id));
		f << "##duplicated,";
	}
	else if (!name.empty() || type==0x97) // external[file]=="" means current
		map[id]=name;
	if (!name.empty())
		f << name.cstr() << ',';
	/* if name not empty,
	   0x97 is followed by first page/second page (or 00000000e97f9409)
	   0x98 is followed by index of owner + a potential cell ref?
	 */
	if (input->tell()!=pos+4+sz) ascFile.addDelimiter(input->tell(),'|');
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroParser::readFontDef(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type != 0xcf && type != 0xfc && type!=0x110)
	{
		WPS_DEBUG_MSG(("QuattroParser::readFontDef: not a font zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	if (type==0xcf)
		f << "[F" << m_state->m_fontsList.size() << "],";
	QuattroParserInternal::Font font(getDefaultFontType());
	if (sz!=0x24)
	{
		WPS_DEBUG_MSG(("QuattroParser::readFontDef: seems very bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		m_state->m_fontsList.push_back(font);
		return true;
	}
	auto fSize = int(libwps::readU16(input));
	if (fSize >= 1 && fSize <= 50) // in fact maximum is 500
		font.m_size=double(fSize);
	else
		f << "###fSize=" << fSize << ",";
	auto flags = int(libwps::readU16(input));
	uint32_t attributes = 0;
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
	flags &= 0xfe00;
	if (flags)
		f << "##fl=" << std::hex << flags << std::dec << ",";
	librevenge::RVNGString name;
	if (!readCString(stream,name,32))
	{
		f << "##name,";
	}
	else
		font.m_name=name;
	if (type==0xcf)
		m_state->m_fontsList.push_back(font);

	f << font;
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroParser::readColorList(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type != 0xe8)
	{
		WPS_DEBUG_MSG(("QuattroParser::readColorList: not a font zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	if (sz<0x40 || (sz%4))
	{
		WPS_DEBUG_MSG(("QuattroParser::readColorList: seems very bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto N=size_t(sz/4);
	m_state->m_colorsList.resize(N);
	for (auto &c: m_state->m_colorsList)
	{
		uint8_t cols[4];
		for (auto &co : cols) co=uint8_t(libwps::readU8(input));
		c=WPSColor(cols[0],cols[1],cols[2],cols[3]);
		f << c << ",";
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroParser::readStyleName(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input))&0x7fff;

	if (type != 0xd0)
	{
		WPS_DEBUG_MSG(("QuattroParser::readStyleName: not a font zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	if (sz<4)
	{
		WPS_DEBUG_MSG(("QuattroParser::readStyleName: seems very bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto id=int(libwps::readU16(input)); // current style id
	f << "St" << id << ",";
	auto val=int(libwps::readU16(input));
	if ((val&0x3fff)!=id) f << "attrib[id]=" << val << ",";
	if (sz!=4)   // no name seems ok
	{
		librevenge::RVNGString name;
		if (!readCString(stream,name,sz-4))
		{
			f << "##name,";
		}
		else if (!name.empty())
			f << name.cstr() << ",";
		if (input->tell()!=pos+4+sz)
			ascFile.addDelimiter(input->tell(),'|');
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

// ----------------------------------------------------------------------
// Header/Footer
// ----------------------------------------------------------------------
void QuattroParser::sendHeaderFooter(bool header)
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("QuattroParser::sendHeaderFooter: can not find the listener\n"));
		return;
	}

	m_listener->setFont(m_state->getDefaultFont());
	auto const &text = header ? m_state->m_headerString : m_state->m_footerString;
	m_listener->insertUnicodeString(text);
}

bool QuattroParser::readHeaderFooter(std::shared_ptr<WPSStream> stream, bool header)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);
	if (type != 0x0026 && type != 0x0025)
	{
		WPS_DEBUG_MSG(("QuattroParser::readHeaderFooter: not a header/footer\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos = pos+4+sz;

	librevenge::RVNGString text;
	if (!readCString(stream,text,sz))
	{
		f << "##sSz,";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	if (!text.empty())
	{
		if (header)
			m_state->m_headerString=text;
		else
			m_state->m_footerString=text;
		f << text.cstr();
	}
	if (input->tell()!=endPos)
		ascFile.addDelimiter(input->tell(), '|');
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	return true;
}

bool QuattroParser::readOptimizer(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	if (type != 0x103)
	{
		WPS_DEBUG_MSG(("QuattroParser::readOptimizer: not an optimizer zone"));
		return false;
	}
	if (sz<84)
	{
		WPS_DEBUG_MSG(("QuattroParser::readOptimizer: seems very bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	for (int i=0; i<2; ++i)
	{
		long actPos=input->tell();
		QuattroFormulaInternal::CellReference instr;
		if (!m_spreadsheetParser->readCellReference(stream, endPos, instr))
		{
			WPS_DEBUG_MSG(("QuattroParser::readOptimizer: can not read some reference\n"));
			f << "###";
			ascFile.addPos(pos);
			ascFile.addNote(f.str().c_str());
			return true;
		}
		else if (!instr.empty())
			f << "cell" << i << "=" << instr << ",";
		input->seek(actPos+10, librevenge::RVNG_SEEK_SET);
	}
	int val=int(libwps::readU16(input));
	if (val==1) f << "goal=min,";
	else if (val==2) f << "goal=max,";
	else if (val==3) f << "goal=targeted,";
	else if (val) f << "##goal=" << val << ",";
	double value;
	bool isNaN;
	long actPos=input->tell();
	if (libwps::readDouble8(input, value, isNaN))
		f << "reached=" << value << ",";
	else
	{
		f << "###reached,";
		input->seek(actPos+10, librevenge::RVNG_SEEK_SET);
	}

	val=int(libwps::readU16(input));
	if (val==1) f << "quadratic,";
	else if (val) f << "##estimate=" << val << ",";
	val=int(libwps::readU16(input));
	if (val==1) f << "derivated=central,";
	else if (val) f << "##derivated=" << val << ",";
	val=int(libwps::readU16(input));
	if (val==1) f << "search=conjugate,";
	else if (val) f << "##search=" << val << ",";
	val=int(libwps::readU16(input));
	if (val==1) f << "linear,";
	else if (val) f << "##linear=" << val << ",";
	val=int(libwps::readU16(input));
	if (val==1) f << "show[result],";
	else if (val) f << "##show[result]=" << val << ",";

	val=int(libwps::readU16(input));
	if (val!=100) f << "max[time]=" << val << ",";
	val=int(libwps::readU16(input));
	if (val!=100) f << "max[iteration]=" << val << ",";
	actPos=input->tell();
	if (libwps::readDouble8(input, value, isNaN))
		f << "precision=" << value << ",";
	else
	{
		f << "###precision,";
		input->seek(actPos+8, librevenge::RVNG_SEEK_SET);
	}
	for (int i=0; i<3; ++i)
	{
		actPos=input->tell();
		QuattroFormulaInternal::CellReference instr;
		if (!m_spreadsheetParser->readCellReference(stream, endPos, instr))
		{
			WPS_DEBUG_MSG(("QuattroParser::readOptimizer: can not read some reference\n"));
			f << "###";
			ascFile.addPos(pos);
			ascFile.addNote(f.str().c_str());
			return true;
		}
		else if (!instr.empty())
			f << "cell" << i+3 << "=" << instr << ",";
		input->seek(actPos+10, librevenge::RVNG_SEEK_SET);
	}
	input->seek(2, librevenge::RVNG_SEEK_CUR); // unused
	if (sz>=94)
	{
		actPos=input->tell();
		if (libwps::readDouble8(input, value, isNaN))
			f << "tolerance=" << value << ",";
		else
		{
			f << "###tolerance,";
			input->seek(actPos+8, librevenge::RVNG_SEEK_SET);
		}
		val=int(libwps::readU16(input));
		if (val) f << "autoScale=" << val << ",";
	}
	if (input->tell()!=endPos)
		ascFile.addDelimiter(input->tell(), '|');
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	return true;
}

bool QuattroParser::readQueryCommand(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	if (type != 0x12f)
	{
		WPS_DEBUG_MSG(("QuattroParser::readQueryCommand: not an queryCommand zone"));
		return false;
	}
	if (sz<22)
	{
		WPS_DEBUG_MSG(("QuattroParser::readQueryCommand: seems very bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	for (int i=0; i<2; ++i)
	{
		long actPos=input->tell();
		QuattroFormulaInternal::CellReference instr;
		if (!m_spreadsheetParser->readCellReference(stream, endPos, instr))
		{
			WPS_DEBUG_MSG(("QuattroParser::readQueryCommand: can not read some reference\n"));
			f << "###";
			ascFile.addPos(pos);
			ascFile.addNote(f.str().c_str());
			return true;
		}
		else if (!instr.empty())
			f << "cell" << i << "=" << instr << ",";
		input->seek(actPos+10, librevenge::RVNG_SEEK_SET);
	}
	int val=int(libwps::readU16(input));
	if (val) f << "id=" << val << ",";
	// then a name?
	if (input->tell()!=endPos)
		ascFile.addDelimiter(input->tell(), '|');
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	return true;
}

////////////////////////////////////////////////////////////
//   Unknown
////////////////////////////////////////////////////////////

bool QuattroParser::readBlockList(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	auto type = long(libwps::readU16(input)&0x7fff);
	int N=0;
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	int extraSize=0;
	switch (type)
	{
	case 0x18: // table, extra value: 0,1,2: 0,1,2 free variables
	case 0x19: // query, extra value: 0-4: meaning none,find,extract,delete, unique
		N=3;
		extraSize=2;
		break;
	case 0x1a: // print block
		N=1;
		break;
	case 0x1b: // sort block
		N=1;
		extraSize=4; // f0=0(numbers first),1(labels first), f1=0(dictionnary),1(ascii) order
		break;
	case 0x1c: // fill range
		N=1;
		break;
	case 0x1d: // key
	case 0x23:
	case 0x9f:
	case 0xa0:
	case 0xa1:
		N=1;
		extraSize=2; // 0: ascending, 1: descending when sz==12
		break;
	case 0x20: // frequency
	case 0x66: // parse
		N=2;
		break;
	case 0x67: // regression
		N=3;
		extraSize=2; // 0: compute-y-intersect, 1: y-intersect is 0
		break;
	case 0x69: // matrix
		N=5;
		break;
	case 0xb7: // solve
		N=2;
		extraSize=18;
		break;
	case 0xf8: // print left
	case 0xf9: // print top
		N=1;
		break;
	case 0x10d: // show compatible
		N=1;
		break;
	case 0x2bf: // X label
	case 0x2c0: // Z label
	case 0x2c1: // legend
	case 0x2c6: // serie data
	case 0x2c7: // serie legend
		N=1;
		break;
	default:
		break;
	}
	bool fixedSize10=10*N+extraSize==sz;
	for (int i=0; i<N; ++i)
	{
		QuattroFormulaInternal::CellReference instr;
		long actPos=input->tell();
		if (!m_spreadsheetParser->readCellReference(stream, endPos, instr))
		{
			WPS_DEBUG_MSG(("QuattroParser::readBlockList: can not read a reference\n"));
			f << "###";
			input->seek(actPos, librevenge::RVNG_SEEK_SET);
			ascFile.addPos(pos);
			ascFile.addNote(f.str().c_str());

			return true;
		}
		if (!instr.empty())
			f << "cell" << i << "=" << instr << ",";
		if (fixedSize10)
			input->seek(actPos+10, librevenge::RVNG_SEEK_SET);
	}
	long remainSize=endPos-input->tell();
	if (type==0xb7 && (remainSize==2 || remainSize==18))
	{
		for (int i=0; i<(remainSize==2 ? 0 : 2); ++i)   // target, accuracy
		{
			long actPos=input->tell();
			double val;
			bool isNaN;
			if (libwps::readDouble8(input, val, isNaN))
				f << "f" << i << "=" << val << ",";
			else
				f << "###f" << i << ",";
			input->seek(actPos+8, librevenge::RVNG_SEEK_SET);
		}
		f << "max[iter]=" << libwps::readU16(input) << ",";
	}
	else if (remainSize!=extraSize)
	{
		ascFile.addDelimiter(input->tell(),'|');
		WPS_DEBUG_MSG(("QuattroParser::readBlockList: unexpected extra data\n"));
		f << "###";
	}
	else
	{
		for (int i=0; i<extraSize/2; ++i)
		{
			auto val=int(libwps::read16(input));
			if (val) f << "f" << i << "=" << val << ",";
		}
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	return true;
}

bool QuattroParser::readCellPosition(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type != 0x96)
	{
		WPS_DEBUG_MSG(("QuattroParser::readCellPosition: not a cell position zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	if (sz%6)
	{
		WPS_DEBUG_MSG(("QuattroParser::readCellPosition: size seems very bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	auto N=int(sz/6);
	for (int i=0; i<N; ++i)
	{
		int cellPos[3]; // col, rowMin, rowMax
		for (auto &c : cellPos) c=int(libwps::readU16(input));
		f <<  "C" << cellPos[0] << "[" << cellPos[1] << "->" << cellPos[2] << "],";
	}
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroParser::readPaneAttribute(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type != 0xd1)
	{
		WPS_DEBUG_MSG(("QuattroParser::readPaneAttribute: not a attribute zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	if (sz<30)
	{
		WPS_DEBUG_MSG(("QuattroParser::readPaneAttribute: size seems very bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	int val=int(libwps::readU8(input));
	if (val==0)
		f << "diplay[no],";
	else if (val!=1)
		f << "##display=" << val << ",";
	val=int(libwps::readU8(input));
	WPSColor color;
	if (!getColor(val, color))
		f << "##lineColor=" << val << ",";
	else if (!color.isBlack())
		f << "lineColor=" << color << ",";
	val=int(libwps::readU16(input));
	switch (val)
	{
	case 0:
		f << "lab[align]=default,";
		break;
	case 1:
		break;
	case 2:
		f << "lab[align]=center,";
		break;
	case 3:
		f << "lab[align]=right,";
		break;
	default:
		f << "##lab[align]=" << val << ",";
		break;
	}
	val=int(libwps::readU16(input));
	switch (val)
	{
	case 0:
		f << "number[align]=default,";
		break;
	case 1:
		f << "number[align]=left,";
		break;
	case 2:
		f << "number[align]=center,";
		break;
	case 3:
		break;
	default:
		f << "##number[align]=" << val << ",";
		break;
	}
	val=int(libwps::readU16(input));
	if (val==1)
	{
		f << "has[cond],";
		for (int i=0; i<2; ++i)
		{
			double value;
			bool isNaN;
			if (libwps::readDouble8(input, value, isNaN))
				f << "cond" << i << "=" << value << ",";
			else
				f << "###cond" << i << ",";
		}
	}
	else if (val)
		f << "##has[cond]=" << val << ",";
	input->seek(pos+4+8+16, librevenge::RVNG_SEEK_SET);
	f << "colors=[";
	for (int i=0; i<4; ++i)   // color < cond0, normal, color> cond1, error
	{
		val=int(libwps::readU8(input));
		int const expected[]= {4,3,5,4};
		if (val==expected[i])
			f << "_,";
		else if (!getColor(val, color))
			f << "##" << val << ",";
		else
			f << color << ",";
	}
	f << "],";
	input->seek(2, librevenge::RVNG_SEEK_CUR); // reserved
	if (sz!=30) // find sometimes +4|8 bytes
		ascFile.addDelimiter(input->tell(),'|');
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroParser::readSerieExtension(std::shared_ptr<WPSStream> stream)
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);

	if (type != 0x2dc)
	{
		WPS_DEBUG_MSG(("QuattroParser::readSerieExtension: not a extension zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	if (sz<6)
	{
		WPS_DEBUG_MSG(("QuattroParser::readSerieExtension: size seems very bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	int val=int(libwps::readU16(input));
	if (val==1)
		f << "ysecondary,";
	else if (val)
		f << "##ysecondaty=" << val << ",";
	val=int(libwps::readU16(input));
	if (val>=1 && val<5)
	{
		char const *wh[]= {nullptr, "bar", "line", "area", "high-low"};
		f << wh[val] << ",";
	}
	else if (val)
		f << "#override=" << val << ",";
	input->seek(2, librevenge::RVNG_SEEK_CUR);
	if (sz<10)
	{
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	val=int(libwps::readU16(input));
	if (val) f << "f0=" << val << ",";
	int dSz=int(libwps::readU16(input));
	if (dSz+5>sz || dSz<4)
	{
		f << "###dSz=" << dSz << ",";
		WPS_DEBUG_MSG(("QuattroParser::readSerieExtension: can not read the size extension\n"));
		ascFile.addDelimiter(input->tell(),'|');
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	val=int(libwps::readU16(input));
	if (val) f << "type=" << val << ","; // 0: aggregation, 1: moving aggregation, 2: linear, 3: expon
	val=int(libwps::readU16(input));
	if (val&1) f << "filter[in legend],";
	if (val&2) f << "filter[in table],";
	if (val&4) f << "table[can be increased],";
	val &= 0xfff8;
	if (val) f << "fl=" << std::hex << val << std::dec << ",";
	if (input->tell()!=endPos)
		ascFile.addDelimiter(input->tell(),'|');
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroParser::readZone341(std::shared_ptr<WPSStream> stream)
{
	if (m_state->m_readingZone341)
	{
		WPS_DEBUG_MSG(("QuattroParser::readZone341: recursive call\n"));
		return false;
	}
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;
	long pos = input->tell();
	auto type = int(libwps::readU16(input)&0x7fff);
	if (type != 0x341)
	{
		WPS_DEBUG_MSG(("QuattroParser::readZone341: not a 341 zone\n"));
		return false;
	}
	auto sz = long(libwps::readU16(input));
	long endPos=pos+4+sz;
	int const headerSize=version()>=1003 ? 82 : 75;
	if (sz<headerSize)
	{
		WPS_DEBUG_MSG(("QuattroParser::readZone341: size seems very bad\n"));
		f << "###";
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
		return true;
	}
	ascFile.addDelimiter(input->tell(),'|');
	input->seek(pos+4+headerSize, librevenge::RVNG_SEEK_SET);
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());

	while (input->tell()+4<=endPos)
	{
		pos=input->tell();
		bool end=(libwps::readU16(input)&0x7fff)==0x31f;
		input->seek(pos,librevenge::RVNG_SEEK_SET);
		m_state->m_readingZone341 = true;
		const bool ok = readZone(stream);
		m_state->m_readingZone341 = false;
		if (!ok || input->tell()>endPos)
		{
			WPS_DEBUG_MSG(("QuattroParser::readZone341: find extra data\n"));
			ascFile.addPos(pos);
			ascFile.addNote("Zone341:###extra");
			return true;
		}
		if (end)
		{
			if (input->tell()<endPos)
			{
				ascFile.addPos(input->tell());
				ascFile.addNote("_");
			}
			return true;
		}
	}
	if (sz!=headerSize)
	{
		WPS_DEBUG_MSG(("QuattroParser::readZone341: oops, does not find end zone\n"));
	}
	return true;
}

////////////////////////////////////////////////////////////
//   ole stream
////////////////////////////////////////////////////////////
bool QuattroParser::readOleLinkInfo(std::shared_ptr<WPSStream> stream, librevenge::RVNGString &link)
{
	if (!stream || !stream->checkFilePosition(4))
	{
		WPS_DEBUG_MSG(("QuattroParser::readLinkInfo: unexpected zone\n"));
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
	if (!readCString(stream, link, stream->m_eof-3))
	{
		WPS_DEBUG_MSG(("QuattroParser::readLinkInfo: can not read the link\n"));
		f << "##link,";
		ascFile.addPos(0);
		ascFile.addNote(f.str().c_str());
		return false;
	}
	if (!link.empty())
		f << "link=" << link.cstr() << ",";
	ascFile.addPos(0);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool QuattroParser::readOleBOlePart(std::shared_ptr<WPSStream> stream)
{
	if (!stream || !stream->checkFilePosition(20))
	{
		WPS_DEBUG_MSG(("QuattroParser::readOleBOlePart: unexpected zone\n"));
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
RVNGInputStreamPtr QuattroParser::decodeStream(RVNGInputStreamPtr input, std::vector<uint8_t> const &key) const
{
	int const vers=version();
	if (!input || key.size()!=16)
	{
		WPS_DEBUG_MSG(("QuattroParser::decodeStream: the arguments seems bad\n"));
		return RVNGInputStreamPtr();
	}
	long actPos=input->tell();
	input->seek(0,librevenge::RVNG_SEEK_SET);
	librevenge::RVNGBinaryData data;
	if (!libwps::readDataToEnd(input, data) || !data.getDataBuffer())
	{
		WPS_DEBUG_MSG(("QuattroParser::decodeStream: can not read the original input\n"));
		return RVNGInputStreamPtr();
	}
	auto *buf=const_cast<unsigned char *>(data.getDataBuffer());
	auto endPos=long(data.size());
	input->seek(actPos,librevenge::RVNG_SEEK_SET);
	uint32_t d7=0;
	std::stack<long> stack;
	stack.push(endPos);
	const int zone341Size=vers<=1002 ? 75 : 82;
	while (!input->isEnd() && !stack.empty())
	{
		long pos=input->tell();
		if (pos+4>stack.top()) break;
		auto id=int(libwps::readU16(input)&0x7fff);
		auto sSz=int(libwps::readU16(input));
		if (pos+4+sSz>stack.top())
		{
			input->seek(pos,librevenge::RVNG_SEEK_SET);
			break;
		}
		if (id==0x341 && sSz>zone341Size)
		{
			// special case a container with header of size 75 + different subzones
			stack.push(pos+4+sSz);
			sSz=zone341Size; // transform only the header
		}
		for (int i=0; i<sSz; ++i)
		{
			auto c=uint8_t(libwps::readU8(input));
			c=(c^key[(d7++)&0xf]);
			buf[pos+4+i]=uint8_t((c>>5)|(c<<3));
		}
		// main zone ends with zone 1, zone 341 ends with zone 31f
		if (id==(stack.size()==1 ? 1 : 0x31f))
		{
			input->seek(stack.top(),librevenge::RVNG_SEEK_SET);
			stack.pop();
		}
	}
	if (input->tell()!=endPos)
	{
		WPS_DEBUG_MSG(("QuattroParser::decodeStream: can not decode the end of the file, data may be bad %lx %lx\n", static_cast<unsigned long>(input->tell()), static_cast<unsigned long>(endPos)));
	}
	RVNGInputStreamPtr res(new WPSStringStream(data.getDataBuffer(), static_cast<unsigned int>(endPos)));
	res->seek(actPos, librevenge::RVNG_SEEK_SET);
	return res;
}
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
