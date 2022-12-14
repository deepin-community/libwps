/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: t; c-basic-offset: 4 -*- */
/* libwps
 * Version: MPL 2.0 / LGPLv2.1+
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *
 * Major Contributor(s):
 * Copyright (C) 2009, 2011 Alonso Laurent (alonso@loria.fr)
 * Copyright (C) 2006, 2007 Andrew Ziem
 * Copyright (C) 2004-2006 Fridrich Strba (fridrich.strba@bluewin.ch)
 * Copyright (C) 2004 Marc Maurer (uwog@uwog.net)
 * Copyright (C) 2003-2005 William Lachance (william.lachance@sympatico.ca)
 *
 * For minor contributions see the git repository.
 *
 * Alternatively, the contents of this file may be used under the terms
 * of the GNU Lesser General Public License Version 2.1 or later
 * (LGPLv2.1+), in which case the provisions of the LGPLv2.1+ are
 * applicable instead of those above.
 *
 * For further information visit http://libwps.sourceforge.net
 */

#include <iomanip>
#include <iostream>

#include <librevenge/librevenge.h>

#include "WPSCell.h"
#include "WPSContentListener.h"
#include "WPSEntry.h"
#include "WPSPosition.h"
#include "WPSTable.h"

#include "WPS8.h"
#include "WPS8Struct.h"

#include "WPS8Table.h"


/** Internal: the structures of a WPS8Table */
namespace WPS8TableInternal
{
/** Internal: class to store a basic cell with borders */
struct Cell final : public WPSCell
{
	//! constructor
	explicit Cell(WPS8Table &parser)
		: WPSCell()
		, m_tableParser(parser)
		, m_id(-1)
		, m_strsId(-1)
		, m_size()
	{
		WPSBorder emptyBorder;
		emptyBorder.m_style = WPSBorder::None;
		m_bordersList.resize(4, emptyBorder);
		for (float &i : m_bordersSep) i = 0.0;
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
			WPS_DEBUG_MSG(("WPS8TableInternal::Cell::send: unexpected listener\n"));
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
		m_tableParser.sendTextInCell(m_strsId, m_id);
		return true;
	}

	//! the actual parser
	WPS8Table &m_tableParser;
	//! the cell id
	int m_id;
	//! the strsId
	mutable int m_strsId;
	//! frame size in inches
	Vec2f m_size;
	//! border text separator T,L,R,B ( checkme, not sure )
	float m_bordersSep[4];
};
Cell::~Cell()
{
}
//! operator<< for a Cell
std::ostream &operator<<(std::ostream &o, Cell const &cell)
{
	o << reinterpret_cast<WPSCell const &>(cell);
	if (cell.m_size.x() > 0 || cell.m_size.y() > 0) o << "size=" << cell.m_size << ",";
	bool hasBSize = false;
	for (float i : cell.m_bordersSep)
	{
		if (i <= 0)
			continue;
		hasBSize = true;
		break;
	}
	if (hasBSize)
	{
		o << "borderSep?=[";
		for (float i : cell.m_bordersSep)
			if (i > 0) o << i << ",";
			else o << "_,";
		o << "],";
	}
	return o;
}

/** Internal: class to store a table: list of cells */
struct Table final : public WPSTable
{
	//! constructor
	Table()
		: m_id(-1)
		, m_parsed(false) {}
	Table &operator=(Table const &)=default;
	//! destructor
	~Table() final;
	//! operator<<
	friend std::ostream &operator<<(std::ostream &o, Table const &table);
	//! the table id
	int m_id;

	//! a bool to know if the table has been parsed
	mutable bool m_parsed;
};

Table::~Table()
{
}

//! operator<< for a Table
std::ostream &operator<<(std::ostream &o, Table const &table)
{
	o << "id=" << table.m_id << ",";
	for (int i = 0; i < table.numCells(); i++)
	{
		WPSCellPtr cell = const_cast<Table &>(table).getCell(i);
		if (!cell) continue;
		o << "cell" << i << "=[" << *reinterpret_cast<Cell const *>(cell.get()) << "],";
	}
	return o;
}

/** Internal: the state of a WPS8Table */
struct State
{
	//! constructor
	State()
		: m_version(-1)
		, m_numPages(0)
		, m_tableMap()
		, m_MCLDTypes()
	{
		initTypeMaps();
	}

	//! initialize the type map
	void initTypeMaps();
	//! the version
	int m_version;
	//! the number page
	int m_numPages;
	//! a map id -> table
	std::map<int, Table> m_tableMap;
	//! the MCLD type
	std::map<int,int> m_MCLDTypes;
};

void State::initTypeMaps()
{
	static int const MCLDTypes[] =
	{
		0, 0x22, 1, 0x22, 2, 0x22, 3, 0x22, 4, 0x22, 5, 0x22, 6, 0x22, 7, 0x22,
		8, 0x22, 9, 0x22, 0xa, 0x22, 0xb, 0x1a, 0xc, 0x2, 0xd, 0x22, 0xe, 0x22,
		0x11, 0x22, 0x12, 0x22, 0x13, 0x12, 0x14, 0x2, 0x15, 0x22, 0x16, 0x22, 0x17, 0x22,
		0x18, 0x22, 0x19, 0x2, 0x1a, 0x2, 0x1d, 0x22, 0x1e, 0x22, 0x1f, 0x12,
		0x20, 0x22, 0x21, 0x12, 0x22, 0x22, 0x23, 0x22, 0x24, 0x12, 0x25, 0x22, 0x26, 0x22, 0x27, 0x12,
		0x28, 0x22, 0x29, 0x22, 0x2a, 0x12, 0x2b, 0x22, 0x2c, 0x12,
		0x31, 0x18
	};
	for (int i = 0; i+1 < int(WPS_N_ELEMENTS(MCLDTypes)); i+=2)
		m_MCLDTypes[MCLDTypes[i]] = MCLDTypes[i+1];

}
}

////////////////////////////////////////////////////////////
// constructor/destructor
////////////////////////////////////////////////////////////
WPS8Table::WPS8Table(WPS8Parser &parser)
	: m_listener()
	, m_mainParser(parser)
	, m_state(new WPS8TableInternal::State)
	, m_asciiFile(parser.ascii())
{
}

WPS8Table::~WPS8Table()
{
}

////////////////////////////////////////////////////////////
// update the positions and send data to the listener
////////////////////////////////////////////////////////////
int WPS8Table::version() const
{
	if (m_state->m_version <= 0)
		m_state->m_version = m_mainParser.version();
	return m_state->m_version;
}

int WPS8Table::numPages() const
{
	return m_state->m_numPages;
}

void WPS8Table::sendTextInCell(int strsId, int cellId)
{
	m_mainParser.sendTextInCell(strsId, cellId);
}

////////////////////////////////////////////////////////////
// update the positions and send data to the listener
////////////////////////////////////////////////////////////
void WPS8Table::computePositions() const
{
	m_state->m_numPages = 0;
}

void WPS8Table::flushExtra()
{
	if (m_listener.get() == nullptr) return;

	for (auto const &table : m_state->m_tableMap)
	{
		if (table.second.m_parsed) continue;
		int strsid = m_mainParser.getTableSTRSId(table.second.m_id);
		if (strsid < 0) continue;
		sendTable(Vec2f(100.,100.),  table.second.m_id, strsid);
	}
}

////////////////////////////////////////////////////////////
// send a table id
////////////////////////////////////////////////////////////
bool WPS8Table::sendTable(Vec2f const &siz, int tableId, int strsid, bool inTextBox)
{
	if (!m_listener)
	{
		WPS_DEBUG_MSG(("WPS8Table::sendTable: listener is not set\n"));
		return true;
	}
	if (strsid <= 0)
	{
		WPS_DEBUG_MSG(("WPS8Table::sendTable: strsid is not set\n"));
		return false;
	}

	auto pos = m_state->m_tableMap.find(tableId);
	if (pos == m_state->m_tableMap.end())
	{
		WPS_DEBUG_MSG(("WPS8Table::sendTable: can not find table with id=%d\n", tableId));
		if (inTextBox)
			m_mainParser.send(strsid);
		else
		{
			// OK, we revert to a textbox inserted as a character
			WPSPosition tablePos(Vec2f(), siz);
			tablePos.m_anchorTo = WPSPosition::CharBaseLine;
			tablePos.m_wrapping = WPSPosition::WDynamic;

			m_mainParser.sendTextBox(tablePos, strsid);
		}
		return true;
	}

	auto &table = pos->second;
	if (table.m_parsed)
		WPS_DEBUG_MSG(("WPS8Table::sendTable: table with id=%d is already parsed\n", tableId));
	else
		table.m_parsed = true;
	// first we need to update the strsid
	for (int c = 0; c < table.numCells(); c++)
	{
		WPSCellPtr cell = table.getCell(c);
		if (!cell) continue;
		reinterpret_cast<WPS8TableInternal::Cell *>(cell.get())->m_strsId = strsid;
	}
	if (!table.sendTable(m_listener))
		table.sendAsText(m_listener);
	return true;
}

////////////////////////////////////////////////////////////
// find all structures which correspond to the table
////////////////////////////////////////////////////////////
bool WPS8Table::readStructures(RVNGInputStreamPtr const &input)
{
	m_state->m_tableMap.clear();

	auto &nameTable = m_mainParser.getNameEntryMap();
	auto pos = nameTable.lower_bound("MCLD");
	while (pos != nameTable.end())
	{
		WPSEntry const &entry = pos++->second;
		if (!entry.hasName("MCLD")) break;
		if (!entry.hasType("MCLD")) continue;

		readMCLD(input, entry);
	}

	return true;
}

////////////////////////////////////////////////////////////
// low level
////////////////////////////////////////////////////////////
bool WPS8Table::readMCLD(RVNGInputStreamPtr input, WPSEntry const &entry)
{
	if (!entry.hasType(entry.name()))
	{
		WPS_DEBUG_MSG(("WPS8Table::readMCLD: warning: MCLD name=%s, type=%s\n",
		               entry.name().c_str(), entry.type().c_str()));
		return false;
	}

	long page_offset = entry.begin();
	long length = entry.length();
	long endPage = entry.end();

	if (length < 24)
	{
		WPS_DEBUG_MSG(("WPS8Table::readMCLD: warning: MCLD length=0x%lx\n", static_cast<unsigned long>(length)));

		return false;
	}

	entry.setParsed();
	input->seek(page_offset, librevenge::RVNG_SEEK_SET);

	libwps::DebugStream f;
	auto mZone = int(libwps::read32(input));
	auto nTables = int(libwps::read32(input));

	f << "maxUnknown=" << mZone << ", nTables= " << nTables;
	if ((6+long(nTables)) * 4 > length) return false;

	f << ", ids=(";
	std::vector<int> listIds;
	for (int i = 0; i < nTables; i++)
	{
		auto val = int(libwps::read32(input));
		listIds.push_back(val);
		f << val << ",";
	}
	f << ")";

	ascii().addPos(page_offset);
	ascii().addNote(f.str().c_str());

	int mCounter = -1;
	bool ok = true;

	static char const *borderNames[] = { "T", "L", "R", "B" };
	static int const borderPos[] =
	{ WPSBorder::Top, WPSBorder::Left, WPSBorder::Right, WPSBorder::Bottom};
	static int const borderBit[] =
	{
		WPSBorder::TopBit, WPSBorder::LeftBit,
		WPSBorder::RightBit, WPSBorder::BottomBit
	};
	long lastPosOk = input->tell();
	while (input->tell() != endPage)
	{
		if (input->tell()+16 > endPage)
		{
			ok = false;
			break;
		}

		mCounter++;
		lastPosOk = input->tell();

		WPS8Struct::FileData tableData;
		auto sz = int(libwps::read16(input));
		if (sz < 2 || lastPosOk+sz > endPage)
		{
			ok = false;
			break;
		}

		std::string error;
		if (!readBlockData(input, lastPosOk+sz,tableData, error))
		{
			f.str("");
			f << tableData;
			error = f.str();
		}

		f.str("");
		int tableId = (mCounter < int(listIds.size())) ? listIds[size_t(mCounter)] : -1;
		if (tableId < 0) f << "MCLD/Table[###unknownId]:";
		else f << "MCLD/Table" << listIds[size_t(mCounter)]<<":";

		auto N = int(libwps::readU32(input));
		f << " nCells=" << N;
		if (N < 0 || N > 100)
		{
			ok = false;
			break;
		}

		WPS8TableInternal::Table table;
		table.m_id = tableId;

		// point dim seems odd, we will need to update them
		Vec2f totalRealDim(0,0), totalDataDim(0,0);

		if (!tableData.m_recursData.empty())
		{
			f << ",(";
			for (auto const &dt : tableData.m_recursData)
			{
				if (dt.isBad()) continue;
				int const expectedTypes[] = { 2, 0x22 };
				if (dt.id() < 0 || dt.id() > 1)
				{
					f << "##" << dt << ",";
					continue;
				}
				if (expectedTypes[dt.id()]!=dt.type())
				{
					WPS_DEBUG_MSG(("WPS8Table::readMCLD: unexpected type for %d=%d\n", dt.id(), dt.type()));
					f << "###" << dt << ",";
					continue;
				}
				bool done = true;
				switch (dt.id())
				{
				case 0: // always set
					f << "f" << dt.id() << ", ";
					break;
				case 1: // always 0
					f << "f" << dt.id() << "=" << dt.m_value << ", ";
					break;
				default:
					done = false;
					break;
				}

				if (!done) f << "###" << dt << ",";
			}
			f << ")";
		}
		if (!error.empty()) f << ", ###err=" << error;

		ascii().addPos(lastPosOk);
		ascii().addNote(f.str().c_str());

		for (size_t actualCell = 0; actualCell < size_t(N); actualCell++)
		{
			lastPosOk = input->tell();
			sz = int(libwps::read16(input));
			if (sz < 2 || lastPosOk+sz > endPage)
			{
				ok = false;
				break;
			}

			WPS8Struct::FileData mainData;
			error = "";

			if (!readBlockData(input, lastPosOk+sz,mainData, error))
			{
				f.str("");
				f << mainData;
				error = f.str();
			}

			// original position in point ( checkme)
			float dim[4] = {  0, 0, 0, 0 };

			f.str("");
			f << "MCLD/Table";
			if (tableId >= 0) f << tableId;
			f << "(Cell" << actualCell <<"):";

			auto cell = std::make_shared<WPS8TableInternal::Cell>(*this);
			WPSCellPtr cellPtr(cell);
			cell->m_id = int(actualCell);

			libwps::DebugStream f2;
			WPSColor cellColor[] = { WPSColor::black(), WPSColor::white() };

			for (auto &dt : mainData.m_recursData)
			{
				if (dt.isBad()) continue;
				if (m_state->m_MCLDTypes.find(dt.id())==m_state->m_MCLDTypes.end())
				{
					f << "##" << dt << ",";
					continue;
				}
				if (m_state->m_MCLDTypes.find(dt.id())->second != dt.type())
				{
					WPS_DEBUG_MSG(("WPS8Table::readMCLD: unexpected type for %d=%d\n", dt.id(), dt.type()));
					f << "###" << dt << ",";
					continue;
				}

				bool done = true;
				switch (dt.id())
				{
				case 0:
				case 1:
				case 2:
				case 3: // min/max pos in point (96 dpi ?)
					dim[dt.id()] = float(dt.m_value);
					break;
				case 4:
				case 5: // look coherent with EOBJ info
					if (dt.id() == 4) cell->m_size.setX(float(dt.m_value)/914400.f);
					else cell->m_size.setY(float(dt.m_value)/914400.f);
					break;
				// border size : unknown dim
				// [ 1, 3, 1, 3, ]
				// [ 0.666667, 1.33333, 0.666667, 1.33333, ]/30
				case 6:
				case 7:
				case 8:
				case 9:
					cell->m_bordersSep[dt.id()-6] = float(dt.m_value)/914400.f;
					break;
				case 0xe: // always 0x20000000
					if (dt.m_value != 0x20000000)
						f2 << "f" << dt.id() << "=" << std::hex << dt.m_value << std::dec << ",";
					break;
				case 0x12: //  0x0 or 0x997aa0 or 0x3fffe238 ?
					f2 << "f" << dt.id() << "=" << std::hex << dt.m_value << std::dec << ",";
					break;
				case 0x13:   //find -1|6 here
					f2 << "f" << dt.id() << "=" << int(int8_t(dt.m_value)) << ",";
					break;
				case 0x1d: // first color
				case 0x1e: // second color
					cellColor[dt.id()-0x1d] = dt.getRGBColor();
					break;
				case 0x1f:
				{
					float percent=0.5;
					if (dt.m_value == 0) // no motif
						break;
					if (dt.m_value >= 3 && dt.m_value <= 9)
						percent = float(dt.m_value)*0.1f; // gray motif
					else
						f2 << "backMotif=" << dt.m_value << ",";
					cell->setBackgroundColor(WPSColor::barycenter(percent,cellColor[0],1.f-percent,cellColor[1]));
					break;
				}
				// the border color
				case 0x20:
				case 0x23:
				case 0x26:
				case 0x29:
				{
					int wh = (dt.id()-0x20)/3;
					WPSBorder border = cell->borders()[size_t(borderPos[wh])];
					border.m_color = dt.getRGBColor();
					cell->setBorders(borderBit[wh], border);
					break;
				}
				case 0x21:
				case 0x24:
				case 0x27:
				case 0x2a:
				{
					std::string mess("");
					int wh = (dt.id()-0x21)/3;
					WPSBorder border = cell->borders()[size_t(borderPos[wh])];
					dt.getBorderStyles(border.m_style, border.m_type, mess);
					cell->setBorders(borderBit[wh], border);
					if (mess.length())
						f2 << "bordStyle" << borderNames[wh] << "=[" << mess << "],";
					break;
				}
				case 0x22:
				case 0x25:
				case 0x28:
				case 0x2b: // often 20
					f2 << "unknBord" << borderNames[(dt.id()-0x22)/3] << "=" << dt.m_value << ",";
					break;
				case 0x2c: // 1, 0, -1
					switch (dt.m_value)
					{
					case 0:
						break; // normal
					case 0xFF: // also unset, diff with value = 1 ?
						f2 << "#f" << dt.id() << "=" << std::hex << dt.m_value << std::dec << ",";
						WPS_FALLTHROUGH;
					case 1:
						cell->setVerticalSet(false);
						break;
					default:
						f2 << "f" << dt.id() << "=" << std::hex << dt.m_value << std::dec << ",";
						break;
					}
					break;
				// always 0,excepted 0x15 hich can be 0 or 1 ?
				case 0xa:
				case 0xb:
				case 0xd:
				case 0x11:
				case 0x15:
				case 0x16:
				case 0x17:
					if (dt.m_value == 0);
					else if (dt.m_value > -10 && dt.m_value < 10)
						f2 << "f" << dt.id() << "=" << dt.m_value << ",";
					else done = false;
					break;
				case 0x18:
					switch (dt.m_value)
					{
					case 0:
						cell->setVAlignment(WPSCellFormat::VALIGN_TOP);
						break;
					case 1:
						cell->setVAlignment(WPSCellFormat::VALIGN_CENTER);
						break;
					case 2:
						cell->setVAlignment(WPSCellFormat::VALIGN_BOTTOM);
						break;
					default:
						if (dt.m_value > -10 && dt.m_value < 10)
							f2 << "f" << dt.id() << "=" << dt.m_value << ",";
						else
							done = false;
						break;
					}
					break;
				case 0xc:
				case 0x14:
				case 0x19:
				case 0x1a: // flag ?
					if (dt.isTrue())
						f2 << "f" << dt.id() << ", ";
					else
						f2 << "f" << dt.id() << "=false, ";
					break;
				default:
					done = false;
					break;
				}

				if (!done)
					f2 << dt << ",";
			}

			cell->setBox(WPSBox2f(Vec2f(dim[1],dim[0]),Vec2f(dim[3],dim[2])));
			totalRealDim += cell->m_size;
			totalDataDim += cell->box().size();
			table.add(cellPtr);

			f << *cell;
			if (!f2.str().empty())
				f << ", unk=(" << f2.str() << ")";

			if (!error.empty()) f << ",###err=" << error;
			input->seek(lastPosOk+sz, librevenge::RVNG_SEEK_SET);

			ascii().addPos(lastPosOk);
			ascii().addNote(f.str().c_str());
		}
		if (!ok) break;

		float factor[2] = { 1, 1};
		if (totalDataDim[0] > 0) factor[0] = 72.f*totalRealDim[0]/totalDataDim[0];
		if (totalDataDim[1] > 0) factor[1] = 72.f*totalRealDim[1]/totalDataDim[1];
		for (int c = 0; c < table.numCells(); c++)
		{
			auto cell = table.getCell(c);
			if (!cell) continue;
			WPSBox2f box=cell->box();
			cell->setBox(WPSBox2f(Vec2f(box[0][0]*factor[0], box[0][1]*factor[1]),
			                      Vec2f(box[1][0]*factor[0], box[1][1]*factor[1])));
		}

		if (tableId >=0) m_state->m_tableMap[tableId] = table;
		else
		{
			WPS_DEBUG_MSG(("WPS8Table::readMCLD: find a table with negative id\n"));
		}
	}

	if (!ok)
	{
		WPS_DEBUG_MSG(("WPS8Table::readMCLD: stopped prematively"));
		ascii().addPos(lastPosOk);
		ascii().addNote("###MCLD");
	}
	return ok;
}
// vim: set filetype=cpp tabstop=2 shiftwidth=2 cindent autoindent smartindent noexpandtab:
