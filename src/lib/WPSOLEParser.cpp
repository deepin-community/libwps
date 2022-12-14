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

/*
 *  freely inspired from istorage :
 * ------------------------------------------------------------
 *      Generic OLE Zones furnished with a copy of the file header
 *
 * Compound Storage (32 bit version)
 * Storage implementation
 *
 * This file contains the compound file implementation
 * of the storage interface.
 *
 * Copyright 1999 Francis Beaudet
 * Copyright 1999 Sylvain St-Germain
 * Copyright 1999 Thuy Nguyen
 * Copyright 2005 Mike McCormack
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
 *
 */

/*
 * NOTES
 *  The compound file implementation of IStorage used for create
 *  and manage substorages and streams within a storage object
 *  residing in a compound file object.
 *
 * ------------------------------------------------------------
 */

#include <cstdlib>
#include <cstring>
#include <map>
#include <sstream>
#include <string>

#include <librevenge/librevenge.h>

#include "WPSHeader.h"
#include "WPSPosition.h"

#include "WPSOLEParser.h"

using namespace libwps;

//////////////////////////////////////////////////
// internal structure
//////////////////////////////////////////////////
namespace WPSOLEParserInternal
{
/** Internal: internal method to compobj definition */
class CompObj
{
public:
	//! the constructor
	CompObj()
		: m_mapCls()
	{
		initCLSMap();
	}

	/** return the CLS Name corresponding to an identifier */
	char const *getCLSName(unsigned long v)
	{
		if (m_mapCls.find(v) == m_mapCls.end()) return nullptr;
		return m_mapCls[v];
	}

protected:
	/** map CLSId <-> name */
	std::map<unsigned long, char const *> m_mapCls;

	/** initialise a map CLSId <-> name */
	void initCLSMap()
	{
		// source: binfilter/bf_so3/source/inplace/embobj.cxx
		m_mapCls[0x00000319]="Picture"; // addon Enhanced Metafile ( find in some file)

		m_mapCls[0x00021290]="MSClipArtGalley2";
		m_mapCls[0x000212F0]="MSWordArt"; // or MSWordArt.2
		m_mapCls[0x00021302]="MSWorksREVENGEoc"; // addon

		// MS Apps
		m_mapCls[0x00030000]= "ExcelWorksheet";
		m_mapCls[0x00030001]= "ExcelChart";
		m_mapCls[0x00030002]= "ExcelMacrosheet";
		m_mapCls[0x00030003]= "WordDocument";
		m_mapCls[0x00030004]= "MSPowerPoint";
		m_mapCls[0x00030005]= "MSPowerPointSho";
		m_mapCls[0x00030006]= "MSGraph";
		m_mapCls[0x00030007]= "MSDraw"; // find also ca003 ?
		m_mapCls[0x00030008]= "Note-It";
		m_mapCls[0x00030009]= "WordArt";
		m_mapCls[0x0003000a]= "PBrush";
		m_mapCls[0x0003000b]= "Equation"; // "Microsoft Equation Editor"
		m_mapCls[0x0003000c]= "Package";
		m_mapCls[0x0003000d]= "SoundRec";
		m_mapCls[0x0003000e]= "MPlayer";
		// MS Demos
		m_mapCls[0x0003000f]= "ServerDemo"; // "OLE 1.0 Server Demo"
		m_mapCls[0x00030010]= "Srtest"; // "OLE 1.0 Test Demo"
		m_mapCls[0x00030011]= "SrtInv"; //  "OLE 1.0 Inv Demo"
		m_mapCls[0x00030012]= "OleDemo"; //"OLE 1.0 Demo"

		// Coromandel / Dorai Swamy / 718-793-7963
		m_mapCls[0x00030013]= "CoromandelIntegra";
		m_mapCls[0x00030014]= "CoromandelObjServer";

		// 3-d Visions Corp / Peter Hirsch / 310-325-1339
		m_mapCls[0x00030015]= "StanfordGraphics";

		// Deltapoint / Nigel Hearne / 408-648-4000
		m_mapCls[0x00030016]= "DGraphCHART";
		m_mapCls[0x00030017]= "DGraphDATA";

		// Corel / Richard V. Woodend / 613-728-8200 x1153
		m_mapCls[0x00030018]= "PhotoPaint"; // "Corel PhotoPaint"
		m_mapCls[0x00030019]= "CShow"; // "Corel Show"
		m_mapCls[0x0003001a]= "CorelChart";
		m_mapCls[0x0003001b]= "CDraw"; // "Corel Draw"

		// Inset Systems / Mark Skiba / 203-740-2400
		m_mapCls[0x0003001c]= "HJWIN1.0"; // "Inset Systems"

		// Mark V Systems / Mark McGraw / 818-995-7671
		m_mapCls[0x0003001d]= "ObjMakerOLE"; // "MarkV Systems Object Maker"

		// IdentiTech / Mike Gilger / 407-951-9503
		m_mapCls[0x0003001e]= "FYI"; // "IdentiTech FYI"
		m_mapCls[0x0003001f]= "FYIView"; // "IdentiTech FYI Viewer"

		// Inventa Corporation / Balaji Varadarajan / 408-987-0220
		m_mapCls[0x00030020]= "Stickynote";

		// ShapeWare Corp. / Lori Pearce / 206-467-6723
		m_mapCls[0x00030021]= "ShapewareVISIO10";
		m_mapCls[0x00030022]= "ImportServer"; // "Spaheware Import Server"

		// test app SrTest
		m_mapCls[0x00030023]= "SrvrTest"; // "OLE 1.0 Server Test"

		// test app ClTest.  Doesn't really work as a server but is in reg db
		m_mapCls[0x00030025]= "Cltest"; // "OLE 1.0 Client Test"

		// Microsoft ClipArt Gallery   Sherry Larsen-Holmes
		m_mapCls[0x00030026]= "MS_ClipArt_Gallery";
		// Microsoft Project  Cory Reina
		m_mapCls[0x00030027]= "MSProject";

		// Microsoft Works Chart
		m_mapCls[0x00030028]= "MSWorksChart";

		// Microsoft Works Spreadsheet
		m_mapCls[0x00030029]= "MSWorksSpreadsheet";

		// AFX apps - Dean McCrory
		m_mapCls[0x0003002A]= "MinSvr"; // "AFX Mini Server"
		m_mapCls[0x0003002B]= "HierarchyList"; // "AFX Hierarchy List"
		m_mapCls[0x0003002C]= "BibRef"; // "AFX BibRef"
		m_mapCls[0x0003002D]= "MinSvrMI"; // "AFX Mini Server MI"
		m_mapCls[0x0003002E]= "TestServ"; // "AFX Test Server"

		// Ami Pro
		m_mapCls[0x0003002F]= "AmiProDocument";

		// WordPerfect Presentations For Windows
		m_mapCls[0x00030030]= "WPGraphics";
		m_mapCls[0x00030031]= "WPCharts";

		// MicroGrafx Charisma
		m_mapCls[0x00030032]= "Charisma";
		m_mapCls[0x00030033]= "Charisma_30"; // v 3.0
		m_mapCls[0x00030034]= "CharPres_30"; // v 3.0 Pres
		// MicroGrafx Draw
		m_mapCls[0x00030035]= "Draw"; //"MicroGrafx Draw"
		// MicroGrafx Designer
		m_mapCls[0x00030036]= "Designer_40"; // "MicroGrafx Designer 4.0"

		// STAR DIVISION
		//m_mapCls[0x000424CA]= "StarMath"; // "StarMath 1.0"
		m_mapCls[0x00043AD2]= "FontWork"; // "Star FontWork"
		//m_mapCls[0x000456EE]= "StarMath2"; // "StarMath 2.0"
	}
};

/** Internal: internal method to keep ole definition */
struct OleDef
{
	OleDef()
		: m_id(-1)
		, m_base("")
		, m_dir("")
		, m_name("") { }
	int m_id /**main id*/;
	std::string m_base/** the base name*/, m_dir/**the directory*/, m_name/**the complete name*/;
};

/** Internal: the state of a WPSOLEParser */
struct State
{
	//! constructor
	State(libwps_tools_win::Font::Type fontType, std::function<int(std::string const &)> const &dirToIdFunc)
		: m_fontType(fontType)
		, m_directoryToIdFunction(dirToIdFunc)
		, m_metaData()
		, m_unknownOLEs()
		, m_idToObjectMap()
		, m_compObjIdName()
	{
	}
	//! the default font type
	libwps_tools_win::Font::Type m_fontType;
	//! the function used to convert a directory name in a id
	std::function<int(std::string const &)> m_directoryToIdFunction;
	//! the meta data
	librevenge::RVNGPropertyList m_metaData;
	//! list of ole which can not be parsed
	std::vector<std::string> m_unknownOLEs;
	//! map id to object
	std::map<int, WPSEmbeddedObject> m_idToObjectMap;

	//! a smart ptr used to stored the list of compobj id->name
	std::shared_ptr<WPSOLEParserInternal::CompObj> m_compObjIdName;
};
}

// constructor/destructor
WPSOLEParser::WPSOLEParser(const std::string &mainName, libwps_tools_win::Font::Type fontType,
                           std::function<int(std::string const &)> const &dirToIdFunc)
	: m_avoidOLE(mainName)
	, m_state(new WPSOLEParserInternal::State(fontType, dirToIdFunc))
{
}

WPSOLEParser::~WPSOLEParser()
{
}


std::vector<std::string> const &WPSOLEParser::getNotParse() const
{
	return m_state->m_unknownOLEs;
}
//! returns the list of data positions which have been read
std::map<int,WPSEmbeddedObject> const &WPSOLEParser::getObjectsMap() const
{
	return m_state->m_idToObjectMap;
}
void WPSOLEParser::updateMetaData(librevenge::RVNGPropertyList &metaData) const
{
	librevenge::RVNGPropertyList::Iter i(m_state->m_metaData);
	for (i.rewind(); i.next();)
	{
		if (!metaData[i.key()])
			metaData.insert(i.key(),i()->clone());
	}
}
// parsing
int WPSOLEParser::getIdFromDirectory(std::string const &dirName)
{
	// try to retrieve the identificator stored in the directory
	//  MatOST/MatadorObject1/ -> 1, -1
	//  Object 2/ -> 2, -1
	auto dir=dirName+'/';
	auto pos = dir.find('/');
	while (pos != std::string::npos)
	{
		if (pos >= 1 && dir[pos-1] >= '0' && dir[pos-1] <= '9')
		{
			auto idP = pos-1;
			while (idP >=1 && dir[idP-1] >= '0' && dir[idP-1] <= '9')
				idP--;
			int val = std::atoi(dir.substr(idP, idP-pos).c_str());
			return val;
		}
		pos = dir.find('/', pos+1);
	}
	WPS_DEBUG_MSG(("WPSOLEParser::getIdFromDirectory: can not find id for %s\n", dirName.c_str()));
	return -1;
}

bool WPSOLEParser::parse(RVNGInputStreamPtr file)
{
	if (!m_state->m_compObjIdName)
		m_state->m_compObjIdName.reset(new WPSOLEParserInternal::CompObj);

	m_state->m_unknownOLEs.resize(0);
	m_state->m_idToObjectMap.clear();

	if (!file || !file->isStructured()) return false;

	//
	// we begin by grouping the Ole by their potential main id
	//
	std::multimap<int, WPSOLEParserInternal::OleDef> listsById;
	std::vector<int> listIds;
	unsigned numSubStreams = file->subStreamCount();
	for (unsigned i = 0; i < numSubStreams; ++i)
	{
		char const *nm=file->subStreamName(i);
		if (!nm) continue;
		std::string name(nm);
		if (name.empty() || name[name.length()-1]=='/') continue;
		// separated the directory and the name
		//    MatOST/MatadorObject1/Ole10Native
		//      -> dir="MatOST/MatadorObject1", base="Ole10Native"
		auto pos = name.find_last_of('/');

		std::string dir, base;
		if (pos == std::string::npos) base = name;
		else if (pos == 0) base = name.substr(1);
		else
		{
			dir = name.substr(0,pos);
			base = name.substr(pos+1);
		}
		if (dir == "" && base == m_avoidOLE) continue;

#define PRINT_OLE_NAME
#ifdef PRINT_OLE_NAME
		WPS_DEBUG_MSG(("WPSOLEParser::parse: find OLEName=%s\n", name.c_str()));
#endif
		WPSOLEParserInternal::OleDef data;
		data.m_name = name;
		data.m_base = base;
		data.m_dir = dir;
		data.m_id = m_state->m_directoryToIdFunction(dir);
		if (listsById.find(data.m_id) == listsById.end())
			listIds.push_back(data.m_id);
		listsById.insert(std::multimap<int, WPSOLEParserInternal::OleDef>::value_type(data.m_id, data));
	}

	for (auto id : listIds)
	{
		auto pos = listsById.lower_bound(id);

		// try to find a representation for each id
		// FIXME: maybe we must also find some for each subid
		WPSEmbeddedObject pict;

		while (pos != listsById.end())
		{
			auto const &dOle = pos->second;
			if (pos->first != id) break;
			++pos;

			RVNGInputStreamPtr ole(file->getSubStreamByName(dOle.m_name.c_str()));
			if (!ole)
			{
				WPS_DEBUG_MSG(("WPSOLEParser: error: can not find OLE part: \"%s\"\n", dOle.m_name.c_str()));
				continue;
			}

			libwps::DebugFile asciiFile(ole);
			asciiFile.open(dOle.m_name);
			bool ok = true;

			try
			{
				librevenge::RVNGPropertyList pList;
				if (readMM(ole, dOle.m_base, asciiFile));
				else if (readSummaryInformation(ole, dOle.m_base, dOle.m_dir.empty() ? m_state->m_metaData : pList, asciiFile));
				else if (readObjInfo(ole, dOle.m_base, asciiFile));
				else if (readOle(ole, dOle.m_base, asciiFile));
				else if (readMN0AndCheckWKS(ole, dOle.m_base, pict, asciiFile));
				else if (isOlePres(ole, dOle.m_base) && readOlePres(ole, pict, asciiFile));
				else if (isOle10Native(ole, dOle.m_base) && readOle10Native(ole, pict, asciiFile));
				else if (readCompObj(ole, dOle.m_base, asciiFile));
				else if (readContents(ole, dOle.m_base, pict, asciiFile));
				else if (readCONTENTS(ole, dOle.m_base, pict, asciiFile));
				else
					ok = false;
			}
			catch (...)
			{
				ok = false;
			}
			if (!ok)
				m_state->m_unknownOLEs.push_back(dOle.m_name);
			asciiFile.reset();
		}

		if (!pict.isEmpty())
			m_state->m_idToObjectMap[id]=pict;
	}

	return true;
}



////////////////////////////////////////
//
// small structure
//
////////////////////////////////////////
bool WPSOLEParser::readOle(RVNGInputStreamPtr &ip, std::string const &oleName,
                           libwps::DebugFile &ascii)
{
	if (!ip.get()) return false;

	if (strcmp("Ole",oleName.c_str()) != 0) return false;

	if (ip->seek(20, librevenge::RVNG_SEEK_SET) != 0 || ip->tell() != 20) return false;

	ip->seek(0, librevenge::RVNG_SEEK_SET);

	int val[20];
	for (int &i : val)
	{
		i = libwps::read8(ip);
		if (i < -10 || i > 10) return false;
	}

	libwps::DebugStream f;
	f << "@@Ole: ";
	// always 1, 0, 2, 0*
	for (int i = 0; i < 20; i++)
		if (val[i]) f << "f" << i << "=" << val[i] << ",";
	ascii.addPos(0);
	ascii.addNote(f.str().c_str());

	if (!ip->isEnd())
	{
		ascii.addPos(20);
		ascii.addNote("@@Ole:###");
	}

	return true;
}

bool WPSOLEParser::readObjInfo(RVNGInputStreamPtr &input, std::string const &oleName,
                               libwps::DebugFile &ascii)
{
	if (strcmp(oleName.c_str(),"ObjInfo") != 0) return false;

	input->seek(14, librevenge::RVNG_SEEK_SET);
	if (input->tell() != 6 || !input->isEnd()) return false;

	input->seek(0, librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f;
	f << "@@ObjInfo:";

	// always 0, 3, 4 ?
	for (int i = 0; i < 3; i++) f << libwps::read16(input) << ",";

	ascii.addPos(0);
	ascii.addNote(f.str().c_str());

	return true;
}

bool WPSOLEParser::readMM(RVNGInputStreamPtr &input, std::string const &oleName,
                          libwps::DebugFile &ascii)
{
	if (strcmp(oleName.c_str(),"MM") != 0) return false;

	input->seek(14, librevenge::RVNG_SEEK_SET);
	if (input->tell() != 14 || !input->isEnd()) return false;

	input->seek(0, librevenge::RVNG_SEEK_SET);
	int entete = libwps::readU16(input);
	if (entete != 0x444e)
	{
		if (entete == 0x4e44)
		{
			WPS_DEBUG_MSG(("WPSOLEParser::readMM: ERROR: endian mode probably bad, potentially bad PC/Mac mode detection.\n"));
		}
		return false;
	}
	libwps::DebugStream f;
	f << "@@MM:";

	int val[6];
	for (int &i : val)
		i = libwps::read16(input);

	switch (val[5])
	{
	case 0:
		f << "conversion,";
		break;
	case 2:
		f << "Works3,";
		break;
	case 4:
		f << "Works4,";
		break;
	default:
		f << "version=unknown,";
		break;
	}

	// 1, 0, 0, 0, 0 : Mac file
	// 0, 1, 0, [0,1,2,4,6], 0 : Pc file
	for (int i = 0; i < 5; i++)
	{
		if ((i%2)!=1 && val[i]) f << "###";
		f << val[i] << ",";
	}

	ascii.addPos(0);
	ascii.addNote(f.str().c_str());

	return true;
}


bool WPSOLEParser::readCompObj(RVNGInputStreamPtr &ip, std::string const &oleName, libwps::DebugFile &ascii)
{
	if (strncmp(oleName.c_str(), "CompObj", 7) != 0) return false;

	// check minimal size
	const int minSize = 12 + 14+ 16 + 12; // size of header, clsid, footer, 3 string size
	if (ip->seek(minSize,librevenge::RVNG_SEEK_SET) != 0 || ip->tell() != minSize) return false;

	libwps::DebugStream f;
	f << "@@CompObj(Header): ";
	ip->seek(0,librevenge::RVNG_SEEK_SET);

	for (int i = 0; i < 6; i++)
	{
		auto val = long(libwps::readU16(ip));
		f << val << ", ";
	}

	ascii.addPos(0);
	ascii.addNote(f.str().c_str());

	ascii.addPos(12);
	// the clsid
	unsigned long clsData[4]; // ushort n1, n2, n3, b8, ... b15
	for (unsigned long &i : clsData)
		i = libwps::readU32(ip);

	f.str("");
	f << "@@CompObj(CLSID):";
	if (clsData[1] == 0 && clsData[2] == 0xC0 && clsData[3] == 0x46000000L)
	{
		// normally, a referenced object
		char const *clsName = m_state->m_compObjIdName->getCLSName(clsData[0]);
		if (clsName)
			f << "'" << clsName << "'";
		else
		{
			WPS_DEBUG_MSG(("WPSOLEParser::readCompObj: unknown clsid=%lx\n", clsData[0]));
			f << "unknCLSID='" << std::hex << clsData[0] << "'";
		}
	}
	else
	{
		/* I found:
		  c1dbcd28e20ace11a29a00aa004a1a72     for MSWorks.Table
		  c2dbcd28e20ace11a29a00aa004a1a72     for Microsoft Works/MSWorksREVENGEoc
		  a3bcb394c2bd1b10a18306357c795b37     for Microsoft Drawing 1.01/MSDraw.1.01
		  b25aa40e0a9ed111a40700c04fb932ba     for Quill96 Story Group Class ( basic MSWorks doc?)
		  796827ed8bc9d111a75f00c04fb9667b     for MSWorks4Sheet
		*/
		f << "data0=(" << std::hex << clsData[0] << "," << clsData[1] << "), "
		  << "data1=(" << clsData[2] << "," << clsData[3] << ")";
	}
	ascii.addNote(f.str().c_str());
	f << std::dec;
	for (int ch = 0; ch < 3; ch++)
	{
		long actPos = ip->tell();
		long sz = libwps::read32(ip);
		bool waitNumber = sz == -1;
		if (waitNumber) sz = 4;
		if (sz < 0 || ip->seek(actPos+4+sz,librevenge::RVNG_SEEK_SET) != 0 ||
		        ip->tell() != actPos+4+sz) return false;
		ip->seek(actPos+4,librevenge::RVNG_SEEK_SET);

		std::string st;
		if (waitNumber)
		{
			f.str("");
			f << libwps::read32(ip) << "[val*]";
			st = f.str();
		}
		else
		{
			for (int i = 0; i < sz; i++)
				st += char(libwps::readU8(ip));
		}

		f.str("");
		f << "@@CompObj:";
		switch (ch)
		{
		case 0:
			f << "UserType=";
			break;
		case 1:
			f << "ClipName=";
			break;
		case 2:
			f << "ProgIdName=";
			break;
		default:
			break;
		}
		f << st;
		ascii.addPos(actPos);
		ascii.addNote(f.str().c_str());
	}

	if (ip->isEnd()) return true;

	long actPos = ip->tell();
	long nbElt = 4;
	if (ip->seek(actPos+16,librevenge::RVNG_SEEK_SET) != 0 ||
	        ip->tell() != actPos+16)
	{
		if ((ip->tell()-actPos)%4) return false;
		nbElt = (ip->tell()-actPos)/4;
	}

	f.str("");
	f << "@@CompObj(Footer): " << std::hex;
	ip->seek(actPos,librevenge::RVNG_SEEK_SET);
	for (int i = 0; i < nbElt; i++)
		f << libwps::readU32(ip) << ",";
	ascii.addPos(actPos);
	ascii.addNote(f.str().c_str());

	ascii.addPos(ip->tell());

	return true;
}

//////////////////////////////////////////////////
//
// OlePres001 seems to contained standart picture file and size
//    extract the picture if it is possible
//
//////////////////////////////////////////////////

bool WPSOLEParser::isOlePres(RVNGInputStreamPtr &ip, std::string const &oleName)
{
	if (!ip.get()) return false;

	if (strncmp("OlePres",oleName.c_str(),7) != 0) return false;

	if (ip->seek(40, librevenge::RVNG_SEEK_SET) != 0 || ip->tell() != 40) return false;

	ip->seek(0, librevenge::RVNG_SEEK_SET);
	for (int i= 0; i < 2; i++)
	{
		long val = libwps::read32(ip);
		if (val < -10 || val > 10) return false;
	}

	long actPos = ip->tell();
	int hSize = libwps::read32(ip);
	if (hSize < 4) return false;
	if (ip->seek(actPos+hSize+28, librevenge::RVNG_SEEK_SET) != 0
	        || ip->tell() != actPos+hSize+28)
		return false;

	ip->seek(actPos+hSize, librevenge::RVNG_SEEK_SET);
	for (int i= 3; i < 7; i++)
	{
		long val = libwps::read32(ip);
		if (val < -10 || val > 10)
		{
			if (i != 5 || val > 256) return false;
		}
	}

	ip->seek(8, librevenge::RVNG_SEEK_CUR);
	long size = libwps::read32(ip);

	if (size <= 0) return ip->isEnd();

	actPos = ip->tell();
	if (ip->seek(actPos+size, librevenge::RVNG_SEEK_SET) != 0
	        || ip->tell() != actPos+size)
		return false;

	return true;
}

bool WPSOLEParser::readOlePres(RVNGInputStreamPtr &ip, WPSEmbeddedObject &obj, libwps::DebugFile &ascii)
{
	if (!isOlePres(ip, "OlePres")) return false;

	libwps::DebugStream f;
	f << "@@OlePress(header): ";
	ip->seek(0,librevenge::RVNG_SEEK_SET);
	for (int i = 0; i < 2; i++)
	{
		long val = libwps::read32(ip);
		f << val << ", ";
	}

	long actPos = ip->tell();
	int hSize = libwps::read32(ip);
	if (hSize < 4) return false;
	f << "hSize = " << hSize;
	ascii.addPos(0);
	ascii.addNote(f.str().c_str());

	long endHPos = actPos+hSize;
	if (hSize > 4)   // CHECKME
	{
		bool ok = true;
		f.str("");
		f << "@@OlePress(headerA): ";
		if (hSize < 14) ok = false;
		else
		{
			// 12,21,32|48,0
			for (int i = 0; i < 4; i++) f << libwps::read16(ip) << ",";
			// 3 name of creator
			for (int ch=0; ch < 3; ch++)
			{
				std::string str;
				bool find = false;
				while (ip->tell() < endHPos)
				{
					unsigned char c = libwps::readU8(ip);
					if (c == 0)
					{
						find = true;
						break;
					}
					str += char(c);
				}
				if (!find)
				{
					ok = false;
					break;
				}
				f << ", name" <<  ch << "=" << str;
			}
			if (ok) ok = ip->tell() == endHPos;
		}
		// FIXME, normally they remain only a few bits (size unknown)
		if (!ok) f << "###";
		ascii.addPos(actPos);
		ascii.addNote(f.str().c_str());
	}
	if (ip->seek(endHPos+28, librevenge::RVNG_SEEK_SET) != 0
	        || ip->tell() != endHPos+28)
		return false;

	ip->seek(endHPos, librevenge::RVNG_SEEK_SET);

	actPos = ip->tell();
	f.str("");
	f << "@@OlePress(headerB): ";
	for (int i = 3; i < 7; i++)
	{
		long val = libwps::read32(ip);
		f << val << ", ";
	}
	// dim in TWIP ?
	auto extendX = long(libwps::readU32(ip));
	auto extendY = long(libwps::readU32(ip));
	if (extendX > 0 && extendY > 0 && obj.m_size!=Vec2f())
		obj.m_size=Vec2f(float(extendX)/1440.f, float(extendY)/1440.f);
	long fSize = libwps::read32(ip);
	f << "extendX="<< extendX << ", extendY=" << extendY << ", fSize=" << fSize;

	ascii.addPos(actPos);
	ascii.addNote(f.str().c_str());

	if (fSize == 0) return ip->isEnd();

	librevenge::RVNGBinaryData data;
	if (!libwps::readData(ip, static_cast<unsigned long>(fSize), data)) return false;
	obj.add(data);
#ifdef DEBUG_WITH_FILES
	std::stringstream s;
	static int num=0;
	s << "OlePress" << num++;
	libwps::Debug::dumpFile(data, s.str().c_str());
#endif
	if (!ip->isEnd())
	{
		ascii.addPos(ip->tell());
		ascii.addNote("@@OlePress###");
	}

	ascii.skipZone(36+hSize,36+hSize+fSize-1);
	return true;
}

//////////////////////////////////////////////////
//
//  Ole10Native: basic Windows picture, with no size
//          - in general used to store a bitmap
//
//////////////////////////////////////////////////

bool WPSOLEParser::isOle10Native(RVNGInputStreamPtr &ip, std::string const &oleName)
{
	if (strncmp("Ole10Native",oleName.c_str(),11) != 0) return false;

	if (ip->seek(4, librevenge::RVNG_SEEK_SET) != 0 || ip->tell() != 4) return false;

	ip->seek(0, librevenge::RVNG_SEEK_SET);
	long size = libwps::read32(ip);

	if (size <= 0) return false;
	if (ip->seek(4+size, librevenge::RVNG_SEEK_SET) != 0 || ip->tell() != 4+size)
		return false;

	return true;
}

bool WPSOLEParser::readOle10Native(RVNGInputStreamPtr &ip, WPSEmbeddedObject &obj, libwps::DebugFile &ascii)
{
	if (!isOle10Native(ip, "Ole10Native")) return false;

	libwps::DebugStream f;
	f << "@@Ole10Native(Header): ";
	ip->seek(0,librevenge::RVNG_SEEK_SET);
	long fSize = libwps::read32(ip);
	f << "fSize=" << fSize;

	ascii.addPos(0);
	ascii.addNote(f.str().c_str());

	librevenge::RVNGBinaryData data;
	if (!libwps::readData(ip, static_cast<unsigned long>(fSize), data)) return false;
	obj.add(data);
#ifdef DEBUG_WITH_FILES
	std::stringstream s;
	static int num=0;
	s << "Ole10_" << num++ << ".bmp";
	libwps::Debug::dumpFile(data, s.str().c_str());
#endif

	if (!ip->isEnd())
	{
		ascii.addPos(ip->tell());
		ascii.addNote("@@Ole10Native###");
	}
	ascii.skipZone(4,4+fSize-1);
	return true;
}

////////////////////////////////////////////////////////////////
//
// In general a picture : a PNG, an JPEG, a basic metafile,
//    find also a MSDraw.1.01 picture (with first bytes 0x78563412="xV4") or WordArt,
//    ( with first bytes "WordArt" )  which are not sucefull read
//    (can probably contain a list of data, but do not know how to
//     detect that)
//
// To check: does this is related to MSO_BLIPTYPE ?
//        or OO/filter/sources/msfilter/msdffimp.cxx ?
//
////////////////////////////////////////////////////////////////
bool WPSOLEParser::readContents(RVNGInputStreamPtr &input,
                                std::string const &oleName,
                                WPSEmbeddedObject &obj,
                                libwps::DebugFile &ascii)
{
	if (strcmp(oleName.c_str(),"Contents") != 0) return false;

	libwps::DebugStream f;
	input->seek(0, librevenge::RVNG_SEEK_SET);
	f << "@@Contents:";

	bool ok = true;
	// bdbox 0 : size in the file ?
	int dim[2];
	dim[0] = libwps::read32(input);
	if (dim[0]==0x12345678)
	{
		WPS_DEBUG_MSG(("WPSOLEParser: warning: find a MSDraw picture, ignored\n"));
		ascii.addPos(0);
		ascii.addNote("Entries(MSDraw):");
		return false;
	}
	dim[1] = libwps::read32(input);
	f << "bdbox0=(" << dim[0] << "," << dim[1]<<"),";
	for (int i = 0; i < 3; i++)
	{
		/* 0,{10|21|75|101|116}x2 */
		auto val = long(libwps::readU32(input));
		if (val < 1000)
			f << val << ",";
		else
			f << std::hex << "0x" << val << std::dec << ",";
		if (val > 0x10000) ok=false;
	}
	// new bdbox : size of the picture ?
	int naturalSize[2];
	naturalSize[0] = libwps::read32(input);
	naturalSize[1] = libwps::read32(input);
	f << std::dec << "bdbox1=(" << naturalSize[0] << "," << naturalSize[1]<<"),";
	f << "unk=" << libwps::readU32(input) << ","; // 24 or 32
	if (input->isEnd())
	{
		WPS_DEBUG_MSG(("WPSOLEParser: warning: Contents header length\n"));
		return false;
	}
	long actPos = input->tell();
	auto size = long(libwps::readU32(input));
	if (size <= 0) ok = false;
	if (ok)
	{
		input->seek(actPos+size+4, librevenge::RVNG_SEEK_SET);
		if (input->tell() != actPos+size+4 || !input->isEnd())
		{
			ok = false;
			WPS_DEBUG_MSG(("WPSOLEParser: warning: Contents unexpected file size=%ld\n",
			               size));
		}
	}

	if (!ok) f << "###";
	f << "dataSize=" << size;

	ascii.addPos(0);
	ascii.addNote(f.str().c_str());

	input->seek(actPos+4, librevenge::RVNG_SEEK_SET);

	if (ok)
	{
		librevenge::RVNGBinaryData data;
		if (libwps::readData(input,static_cast<unsigned long>(size), data))
		{
			obj.add(data);
			ascii.skipZone(actPos+4, actPos+size+4-1);
#ifdef DEBUG_WITH_FILES
			std::stringstream s;
			static int fileId=0;
			s << oleName << ++fileId << "cntents.pict";
			libwps::Debug::dumpFile(data, s.str().c_str());
#endif
		}
		else
		{
			input->seek(actPos+4, librevenge::RVNG_SEEK_SET);
			ok = false;
		}
	}
	if (ok)
	{
		if (dim[0] > 0 && dim[0] < 3000 && dim[1] > 0 && dim[1] < 3000 && obj.m_size!=Vec2f())
			obj.m_size=Vec2f(float(dim[0])/72.f,float(dim[1])/72.f);
		else
		{
			WPS_DEBUG_MSG(("WPSOLEParser: warning: Contents odd size : %d %d\n", dim[0], dim[1]));
		}
		if (naturalSize[0] > 0 && naturalSize[0] < 5000 &&
		        naturalSize[1] > 0 && naturalSize[1] < 5000 && obj.m_size!=Vec2f())
			obj.m_size=Vec2f(float(dim[0])/72.f,float(dim[1])/72.f);
		else
		{
			WPS_DEBUG_MSG(("WPSOLEParser: warning: Contents odd naturalsize : %d %d\n", naturalSize[0], naturalSize[1]));
		}
	}
	if (!input->isEnd())
	{
		ascii.addPos(actPos);
		ascii.addNote("@@Contents:###");
	}

	if (!ok)
	{
		WPS_DEBUG_MSG(("WPSOLEParser: warning: read ole Contents: failed\n"));
	}
	return ok;
}

////////////////////////////////////////////////////////////////
//
// Another different type of contents (this time in majuscule)
// we seem to contain the header of a EMF and then the EMF file
//
////////////////////////////////////////////////////////////////
bool WPSOLEParser::readCONTENTS(RVNGInputStreamPtr &input,
                                std::string const &oleName,
                                WPSEmbeddedObject &obj,
                                libwps::DebugFile &ascii)
{
	if (strcmp(oleName.c_str(),"CONTENTS") != 0) return false;

	libwps::DebugStream f;

	input->seek(0, librevenge::RVNG_SEEK_SET);
	f << "@@CONTENTS:";

	auto hSize = long(libwps::readU32(input));
	if (input->isEnd()) return false;
	f << "hSize=" << std::hex << hSize << std::dec;

	if (hSize <= 52 || input->seek(hSize+8,librevenge::RVNG_SEEK_SET) != 0
	        || input->tell() != hSize+8)
	{
		WPS_DEBUG_MSG(("WPSOLEParser: warning: CONTENTS headerSize=%ld\n",
		               hSize));
		return false;
	}

	// minimal checking of the "copied" header
	input->seek(4, librevenge::RVNG_SEEK_SET);
	auto type = long(libwps::readU32(input));
	if (type < 0 || type > 4) return false;
	auto newSize = long(libwps::readU32(input));

	f << ", type=" << type;
	if (newSize < 8) return false;

	if (newSize != hSize) // can sometimes happen, pb after a conversion ?
		f << ", ###newSize=" << std::hex << newSize << std::dec;

	// checkme: two bdbox, in document then data : units ?
	//     Maybe first in POINT, second in TWIP ?
	for (int st = 0; st < 2 ; st++)
	{
		int dim[4];
		for (int &i : dim) i = libwps::read32(input);

		bool ok = dim[0] >= 0 && dim[2] > dim[0] && dim[1] >= 0 && dim[3] > dim[2];
		if (ok && st==0 && obj.m_size==Vec2f())
			obj.m_size=Vec2f(float(dim[2]-dim[0])/72.f, float(dim[3]-dim[1])/72.f);
		if (st==0) f << ", bdbox(Text)";
		else f << ", bdbox(Data)";
		if (!ok) f << "###";
		f << "=(" << dim[0] << "x" << dim[1] << "<->" << dim[2] << "x" << dim[3] << ")";
	}
	char dataType[5];
	for (int i = 0; i < 4; i++) dataType[i] = char(libwps::readU8(input));
	dataType[4] = '\0';
	f << ",typ=\""<<dataType<<"\""; // always " EMF" ?

	for (int i = 0; i < 2; i++)   // always id0=0, id1=1 ?
	{
		int val = libwps::readU16(input);
		if (val) f << ",id"<< i << "=" << val;
	}
	auto dataLength = long(libwps::readU32(input));
	f << ",length=" << dataLength+hSize;

	ascii.addPos(0);
	ascii.addNote(f.str().c_str());

	ascii.addPos(input->tell());
	f.str("");
	f << "@@CONTENTS(2)";
	for (int i = 0; i < 12 && 4*i+52 < hSize; i++)
	{
		// f0=7,f1=1,f5=500,f6=320,f7=1c4,f8=11a
		// or f0=a,f1=1,f2=2,f3=6c,f5=480,f6=360,f7=140,f8=f0
		// or f0=61,f1=1,f2=2,f3=58,f5=280,f6=1e0,f7=a9,f8=7f
		// f3=some header sub size ? f5/f6 and f7/f8 two other bdbox ?
		auto val = long(libwps::readU32(input));
		if (val) f << std::hex << ",f" << i << "=" << val;
	}
	for (int i = 0; 2*i+100 < hSize; i++)
	{
		// g0=e3e3,g1=6,g2=4e6e,g3=4
		// g0=e200,g1=4,g2=a980,g3=3,g4=4c,g5=50
		// ---
		long val = libwps::readU16(input);
		if (val) f << std::hex << ",g" << i << "=" << val;
	}
	ascii.addNote(f.str().c_str());

	if (dataLength <= 0 || input->seek(hSize+4+dataLength,librevenge::RVNG_SEEK_SET) != 0
	        || input->tell() != hSize+4+dataLength || !input->isEnd())
	{
		WPS_DEBUG_MSG(("WPSOLEParser: warning: CONTENTS unexpected file length=%ld\n",
		               dataLength));
		return false;
	}

	input->seek(4+hSize, librevenge::RVNG_SEEK_SET);
	librevenge::RVNGBinaryData data;
	if (!libwps::readDataToEnd(input, data)) return false;
	obj.add(data);
#ifdef DEBUG_WITH_FILES
	std::stringstream s;
	static int fileId=0;
	s << oleName << ++fileId << "Contents.pict";
	libwps::Debug::dumpFile(data, s.str().c_str());
#endif

	ascii.skipZone(hSize+4, input->tell()-1);
	return true;
}

////////////////////////////////////////////////////////////////
//
// the MN0 subdirectory
//
////////////////////////////////////////////////////////////////
bool WPSOLEParser::readMN0AndCheckWKS(RVNGInputStreamPtr &input, std::string const &oleName,
                                      WPSEmbeddedObject &obj, libwps::DebugFile &/*ascii*/)
{
	if (strcmp(oleName.c_str(),"MN0") != 0) return false;
	std::unique_ptr<WPSHeader> header(WPSHeader::constructHeader(input));
	if (!header) return false;
	bool ok=header->getKind()==WPS_SPREADSHEET;
	if (!ok) return false;

	input->seek(0, librevenge::RVNG_SEEK_SET);
	librevenge::RVNGBinaryData data;
	if (!libwps::readDataToEnd(input, data))
		return false;
	obj.add(data,"image/wks-ods");
#ifdef DEBUG_WITH_FILES
	std::stringstream s;
	static int fileId=0;
	s << oleName << ++fileId << ".wks";
	libwps::Debug::dumpFile(data, s.str().c_str());
#endif
	return true;
}

////////////////////////////////////////////////////////////
//
// Summary function
//
////////////////////////////////////////////////////////////
bool WPSOLEParser::readSummaryInformation(RVNGInputStreamPtr input, std::string const &oleName,
                                          librevenge::RVNGPropertyList &pList, libwps::DebugFile &ascii) const
{
	if (oleName!="SummaryInformation") return false;
	input->seek(0, librevenge::RVNG_SEEK_END);
	long endPos=input->tell();
	input->seek(0, librevenge::RVNG_SEEK_SET);
	libwps::DebugStream f;
	f << "Entries(SumInfo):";
	auto val=int(libwps::readU16(input));
	if (endPos<48 || val!=0xfffe)
	{
		WPS_DEBUG_MSG(("WPSOLEParser::readSummaryInformation: header seems bad\n"));
		f << "###";
		ascii.addPos(0);
		ascii.addNote(f.str().c_str());
		return true;
	}
	for (int i=0; i<11; ++i)   // f1=1, f2=0-2
	{
		val=int(libwps::readU16(input));
		if (val) f << "f" << i << "=" << val << ",";
	}
	unsigned long lVal=libwps::readU32(input);
	if (lVal==0 || lVal>15)   // find 1 or 2 sections, unsure about the maximum numbers
	{
		WPS_DEBUG_MSG(("WPSOLEParser::readSummaryInformation: summary info is bad\n"));
		f << "###sumInfo=" << std::hex << lVal << std::dec << ",";
		ascii.addPos(0);
		ascii.addNote(f.str().c_str());
		return true;
	}
	auto numSection=int(lVal);
	if (numSection!=1)
		f << "num[section]=" << numSection << ",";
	for (int i=0; i<4; ++i)
	{
		val=int(libwps::readU32(input));
		static int const expected[]= {int(0xf29f85e0),0x10684ff9,0x891ab,int(0xd9b3272b)};
		if (val==expected[i]) continue;
		f << "#fmid" << i << "=" << std::hex << val << std::dec << ",";
		static bool first=true;
		if (first)
		{
			WPS_DEBUG_MSG(("WPSOLEParser::readSummaryInformation: fmid is bad\n"));
			first=false;
		}
	}
	auto decal=int(libwps::readU32(input));
	if (decal<0x30 || endPos<decal)
	{
		WPS_DEBUG_MSG(("WPSOLEParser::readSummaryInformation: decal is bad\n"));
		f << "decal=" << val << ",";
		ascii.addPos(0);
		ascii.addNote(f.str().c_str());
		return true;

	}
	ascii.addPos(0);
	ascii.addNote(f.str().c_str());
	if (decal!=0x30)
	{
		ascii.addPos(0x30);
		ascii.addNote("_");
		input->seek(decal, librevenge::RVNG_SEEK_SET);
	}

	for (int sect=0; sect<numSection; ++sect)
	{
		long pos=input->tell();
		f.str("");
		f << "SumInfo-A:";
		auto pSectSize=long(libwps::readU32(input));
		long endSect=pos+pSectSize;
		auto N=int(libwps::readU32(input));
		f << "N=" << N << ",";
		if (pSectSize<0 || endPos-pos<pSectSize || (pSectSize-8)/8<N)
		{
			WPS_DEBUG_MSG(("WPSOLEParser::readSummaryInformation: psetstruct is bad\n"));
			f << "###";
			ascii.addPos(pos);
			ascii.addNote(f.str().c_str());
			return true;
		}
		f << "[";
		std::map<long,int> posToTypeMap;
		for (int i=0; i<N; ++i)
		{
			auto type=int(libwps::readU32(input));
			auto depl=int(libwps::readU32(input));
			if (depl<=0) continue;
			f << std::hex << depl << std::dec << ":" << type << ",";
			if ((depl-8)/8<N || depl>pSectSize-4 || posToTypeMap.find(pos+depl)!=posToTypeMap.end())
			{
				f << "###";
				continue;
			}
			posToTypeMap[pos+depl]=type;
		}
		f << "],";
		ascii.addPos(pos);
		ascii.addNote(f.str().c_str());

		for (auto it=posToTypeMap.begin(); it!=posToTypeMap.end(); ++it)
		{
			pos=it->first;
			auto nextIt=it;
			if (++nextIt!=posToTypeMap.end()) endPos=nextIt->first;
			input->seek(pos, librevenge::RVNG_SEEK_SET);
			f.str("");
			f << "SumInfo-B" << it->second << ":";
			auto type=int(libwps::readU32(input));
			if (sect==0 && it->second==1 && type==2)
			{
				long value=-1;
				if (readSummaryPropertyLong(input,endPos,type,value,f) && value>=0 && value<10000) // 10000 is mac
					f << "encoding,"; // USEME: never seems actually
			}
			else if (sect==0 && type==0x1e && ((it->second>=2 && it->second<=6) || it->second==8))
			{
				librevenge::RVNGString text;
				if (readSummaryPropertyString(input, endPos, type, text, f) && !text.empty())
				{
					static char const *attribNames[] =
					{
						"", "", "dc:title", "dc:subject", "meta:initial-creator",
						"meta:keywords", "dc:description"/*comment*/, "", "dc:creator"
					};
					pList.insert(attribNames[it->second], text);
				}
			}
			else if (!readSummaryProperty(input, endPos, type, ascii, f))
			{
				WPS_DEBUG_MSG(("WPSOLEParser::readSummaryInformation: find unknown type\n"));
				f << "##type=" << std::hex << type << std::dec << ",";
			}
			if (input->tell()!=endPos && input->tell()!=pos)
				ascii.addDelimiter(input->tell(),'|');
			ascii.addPos(pos);
			ascii.addNote(f.str().c_str());
		}
		input->seek(endSect, librevenge::RVNG_SEEK_SET);
	}
	return true;
}

bool WPSOLEParser::readSummaryPropertyString(RVNGInputStreamPtr input, long endPos, int type,
                                             librevenge::RVNGString &string, libwps::DebugStream &f) const
{
	if (!input) return false;
	long pos=input->tell();
	string.clear();
	auto sSz=long(libwps::readU32(input));
	if (sSz<0 || (endPos-pos-4)<sSz || pos+4+sSz>endPos)
	{
		WPS_DEBUG_MSG(("WPSOLEParser::readSummaryPropertyString: string size is bad\n"));
		f << "##stringSz=" << sSz << ",";
		return false;
	}
	std::string text("");
	for (long c=0; c < sSz; ++c)
	{
		auto ch=static_cast<unsigned char>(libwps::readU8(input));
		if (ch)
			text+=char(ch);
		else if (c+1!=sSz)
			text+="##";
	}
	f << text;
	if (!text.empty())
		string=libwps_tools_win::Font::unicodeString(text, m_state->m_fontType);
	if (type==0x1f && (sSz%4))
		input->seek(sSz%4, librevenge::RVNG_SEEK_CUR);
	return true;
}

bool WPSOLEParser::readSummaryPropertyLong(RVNGInputStreamPtr input, long endPos, int type, long &value,
                                           libwps::DebugStream &f)
{
	if (!input) return false;
	long pos=input->tell();
	switch (type)
	{
	case 2: // int
	case 0x12: // uint
		if (pos+2>endPos)
			return false;
		value=type==2 ? long(libwps::read16(input)) : long(libwps::readU16(input));
		break;
	case 3: // int
	case 9: // uint
		if (pos+4>endPos)
			return false;
		value=type==3 ? long(libwps::read32(input)) : long(libwps::readU32(input));
		break;
	default:
		return false;
	}
	f << "val=" << value << ",";
	return true;
}

bool WPSOLEParser::readSummaryProperty(RVNGInputStreamPtr input, long endPos, int type,
                                       libwps::DebugFile &ascii, libwps::DebugStream &f) const
{
	if (!input) return false;
	long pos=input->tell();
	// see propread.cxx
	if (type&0x1000)
	{
		auto N=int(libwps::readU32(input));
		f << "N=" << N << ",";
		f << "[";
		for (int n=0; n<N; ++n)
		{
			pos=input->tell();
			f << "[";
			if (!readSummaryProperty(input, endPos, type&0xFFF, ascii, f))
			{
				input->seek(pos, librevenge::RVNG_SEEK_SET);
				return false;
			}
			f << "],";
		}
		f << "],";
		return true;
	}
	switch (type)
	{
	case 0x10: // int1
	case 0x11: // uint1
		if (pos+1>endPos)
			return false;
		f << "val=" << char(libwps::readU8(input));
		break;
	case 2: // int
	case 0xb: // bool
	case 0x12: // uint
		if (pos+2>endPos)
			return false;
		if (type==2)
			f << "val=" << int(libwps::read16(input)) << ",";
		else if (type==0x12)
			f << "val=" << int(libwps::readU16(input)) << ",";
		else if (libwps::readU16(input))
			f << "true,";
		break;
	case 3: // int
	case 4: // float
	case 9: // uint
		if (pos+4>endPos)
			return false;
		if (type==3)
			f << "val=" << int(libwps::read32(input)) << ",";
		else if (type==9)
			f << "val=" << int(libwps::readU32(input)) << ",";
		else
			f << "val[fl4]=" << std::hex << libwps::readU32(input) << std::dec << ",";
		break;
	case 5: // double
	case 6:
	case 7:
	case 20:
	case 21:
	case 0x40:
		if (pos+8>endPos)
			return false;
		ascii.addDelimiter(input->tell(),'|');
		if (type==5)
			f << "double,";
		else if (type==6)
			f << "cy,";
		else if (type==7)
			f << "date,";
		else if (type==20)
			f << "long,";
		else if (type==21)
			f << "ulong,";
		else
			f << "fileTime,"; // readme 8 byte
		input->seek(pos+8, librevenge::RVNG_SEEK_SET);
		break;
	case 0xc: // variant
		if (pos+4>endPos)
			return false;
		type=int(libwps::readU32(input));
		return readSummaryProperty(input, endPos, type, ascii, f);
	// case 20: int64
	// case 21: uint64
	case 8:
	case 0x1e:
	case 0x1f:
	{
		librevenge::RVNGString string;
		if (!readSummaryPropertyString(input, endPos, type, string, f))
			return false;
		break;
	}
	case 0x41:
	case 0x46:
	case 0x47:
	{
		if (pos+4>endPos)
			return false;
		f << (type==0x41 ? "blob" : type==0x46 ? "blob[object]" : "clipboard") << ",";
		auto dSz=long(libwps::readU32(input));
		if (dSz<0 || pos+4+dSz>endPos)
			return false;
		if (dSz)
		{
			ascii.skipZone(pos+4, pos+4+dSz-1);
			input->seek(dSz, librevenge::RVNG_SEEK_CUR);
		}
		break;
	}
	/* todo type==0x47, vtcf clipboard */
	default:
		return false;
	}
	return true;
}
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
