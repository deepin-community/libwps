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

#include <sstream>
#include <string>

#include <librevenge/librevenge.h>

#include "WPSOLEObject.h"

#include "WPSDebug.h"
#include "WPSStream.h"
#include "WPSStringStream.h"

using namespace libwps;


////////////////////////////////////////////////////////////
// read the OLE object
////////////////////////////////////////////////////////////
bool WPSOLEObject::readMetafile(std::shared_ptr<WPSStream> stream, WPSEmbeddedObject &object, long endPos, bool strict)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	if (endPos<=0) endPos=stream->m_eof;
	else if (endPos>stream->m_eof) endPos=stream->m_eof;

	long pos = input->tell();
	if (pos+8+4>endPos)
		return false;
	f << "Entries(Metafile):";
	f << "type=" << libwps::readU16(input) << ",";
	float fDim[2];
	for (auto &d : fDim) d=float(libwps::read16(input)/1440.);
	if (fDim[0]<0||fDim[1]<0)
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	if (object.m_size==Vec2f() && fDim[0]>0 && fDim[1]>0)
	{
		object.m_size=Vec2f(fDim[0],fDim[1]);
		f << "sz=" << object.m_size << ",";
	}
	input->seek(2, librevenge::RVNG_SEEK_CUR); // seek handle
	if (strict)
	{
		if (!checkIsWMF(stream, endPos))
		{
			input->seek(pos, librevenge::RVNG_SEEK_SET);
			return false;
		}
		input->seek(pos+8, librevenge::RVNG_SEEK_SET);
	}
	librevenge::RVNGBinaryData data;
	if (!libwps::readData(input, static_cast<unsigned long>(endPos-pos-8), data))
	{
		WPS_DEBUG_MSG(("WPSOLEObject::readMetafile: I can not find the picture\n"));
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	object.add(data,"application/x-wmf");
#ifdef DEBUG_WITH_FILES
	ascFile.skipZone(pos+8, endPos-1);
	std::stringstream s;
	static int fileId=0;
	s << "PictMeta" << ++fileId << ".wmf";
	libwps::Debug::dumpFile(data, s.str().c_str());
#endif

	input->seek(endPos, librevenge::RVNG_SEEK_SET);
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool WPSOLEObject::readWMF(std::shared_ptr<WPSStream> stream, WPSEmbeddedObject &object,long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	long pos=input->tell();

	long lastPos=endPos<=0 ? stream->m_eof : endPos;
	if (lastPos>stream->m_eof) lastPos=stream->m_eof;
	if (!checkIsWMF(stream,lastPos)) return false;

	input->seek(pos+6, librevenge::RVNG_SEEK_SET);
	auto fSize = long(libwps::read32(input));
	input->seek(pos, librevenge::RVNG_SEEK_SET);
	librevenge::RVNGBinaryData data;
	if (!libwps::readData(input, static_cast<unsigned long>(2*fSize), data))
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	object.add(data,"application/x-wmf");
#ifdef DEBUG_WITH_FILES
	std::stringstream f;
	static volatile int actPict = 0;
	f << "WMF" << actPict++ << ".wmf";
	libwps::Debug::dumpFile(data, f.str().c_str());
	ascFile.skipZone(pos,pos+2*fSize-1);
#endif
	if (endPos>0 && input->tell()!=endPos)
	{
		ascFile.addPos(input->tell());
		ascFile.addNote("_");
		input->seek(endPos, librevenge::RVNG_SEEK_SET);
	}
	return true;
}

bool WPSOLEObject::readOLE(std::shared_ptr<WPSStream> stream, WPSEmbeddedObject &object, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	if (endPos<=0) endPos=stream->m_eof;
	else if (endPos>stream->m_eof) endPos=stream->m_eof;

	long pos = input->tell();
	if (pos+24>endPos || libwps::read32(input)!=0x501)
	{
		WPS_DEBUG_MSG(("WPSOLEObject::readOLE: not a picture header\n"));
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	f.str("");
	f << "Entries(OLEObject):";
	auto type=int(readU32(input));
	f << "type=" << type << ",";
	input->seek(pos, librevenge::RVNG_SEEK_SET);
	bool ok=false;
	switch (type)
	{
	case 1:
		WPS_DEBUG_MSG(("WPSOLEObject::readOLE: find a link ole\n"));
		f << "link,";
		break;
	case 2:
		ok=readEmbeddedOLE(stream, object, endPos);
		break;
	case 3:
	case 5:
		ok=readStaticOLE(stream, object, endPos);
		break;
	default:
		WPS_DEBUG_MSG(("WPSOLEObject::readOLE: find a unknown type\n"));
		f << "unknown,";
		break;
	}
	if (!ok)
	{
		f << "###";
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		ascFile.addPos(pos);
		ascFile.addNote(f.str().c_str());
	}
	return true;
}

bool WPSOLEObject::readEmbeddedOLE(std::shared_ptr<WPSStream> stream, WPSEmbeddedObject &object, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	if (pos+24+4>endPos || libwps::readU32(input)!=0x501)
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	auto type=int(libwps::readU32(input));
	if (type!=2)
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	f << "Entries(OLEObject)[embedded]:";
	std::string names[3];
	for (auto &name : names)
	{
		if (!readString(stream, name, endPos) || input->tell()+4>endPos)
		{
			WPS_DEBUG_MSG(("WPSOLEObject::readEmbeddedOLE: can not read the name\n"));
			input->seek(pos, librevenge::RVNG_SEEK_SET);
			return false;
		}
		if (name.empty()) continue;
		f << name << ",";
	}
	// find Paint.Picture, WangImage.Document
	unsigned long dSz=libwps::readU32(input);
	long actPos=input->tell();
	if (dSz>0x40000000 || dSz<10 || dSz>static_cast<unsigned long>(endPos-actPos))
	{
		WPS_DEBUG_MSG(("WPSOLEObject::readOLE: pict size seems bad\n"));
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}

	bool ok=true;
	if (names[0]=="METAFILEPICT")
		ok=readMetafile(stream, object, actPos+long(dSz));
	else
	{
		librevenge::RVNGBinaryData data;
		if (!libwps::readData(input, dSz, data))
			ok=false;
		else
		{
			object.add(data);
#ifdef DEBUG_WITH_FILES
			ascFile.skipZone(actPos, actPos+long(dSz)-1);
			std::stringstream s;
			static int fileId=0;
			s << "PictOLEEmbedded" << ++fileId << ".pct";
			libwps::Debug::dumpFile(data, s.str().c_str());
#endif
		}
	}
	if (!ok)
	{
		WPS_DEBUG_MSG(("WPSOLEObject::readEmbeddedOLE: I can not find the picture\n"));
		f << "###";
	}
	input->seek(actPos+long(dSz), librevenge::RVNG_SEEK_SET);
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	if (input->tell()<endPos) // normally followed by a static OLE
		readStaticOLE(stream, object, endPos);
	return true;
}

bool WPSOLEObject::readStaticOLE(std::shared_ptr<WPSStream> stream, WPSEmbeddedObject &object, long endPos)
{
	if (!stream) return false;
	RVNGInputStreamPtr &input=stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	libwps::DebugStream f;

	long pos = input->tell();
	if (pos+24+4>endPos || libwps::readU32(input)!=0x501)
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	auto type=int(libwps::readU32(input));
	if (type!=3 && type!=5)
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	f << "Entries(OLEObject)[static]:";
	f << "type=" << type << ",";
	std::string name;
	if (!readString(stream, name, endPos) || input->tell()+12>endPos)
	{
		WPS_DEBUG_MSG(("WPSOLEObject::readStaticOLE: can not read the name\n"));
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	// find METAFILEPICT
	f << name << ",";
	for (int i=0; i<2; ++i)   // g0+g1~0, some application id?
	{
		auto val=long(libwps::read32(input));
		if (val) f << "g" << i << "=" << val << ",";
	}
	unsigned long dSz=libwps::readU32(input);
	long actPos=input->tell();
	if (dSz>0x40000000 || dSz<10 || dSz>static_cast<unsigned long>(endPos-actPos))
	{
		WPS_DEBUG_MSG(("WPSOLEObject::readOLE: pict size seems bad\n"));
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}

	bool ok=true;
	if (name=="METAFILEPICT")
		ok=readMetafile(stream, object, actPos+long(dSz));
	else
	{
		librevenge::RVNGBinaryData data;
		if (!libwps::readData(input, dSz, data))
			ok=false;
		else
		{
			object.add(data);
#ifdef DEBUG_WITH_FILES
			ascFile.skipZone(actPos, actPos+long(dSz)-1);
			std::stringstream s;
			static int fileId=0;
			s << "PictOLEStatic" << ++fileId << ".pct";
			libwps::Debug::dumpFile(data, s.str().c_str());
#endif
		}
	}
	if (!ok)
	{
		WPS_DEBUG_MSG(("WPSOLEObject::readStaticOLE: I can not find the picture\n"));
		f << "###";
	}
	input->seek(actPos+long(dSz), librevenge::RVNG_SEEK_SET);
	ascFile.addPos(pos);
	ascFile.addNote(f.str().c_str());
	return true;
}

bool WPSOLEObject::readString(std::shared_ptr<WPSStream> stream, std::string &name, long endPos)
{
	if (!stream) return false;
	name="";
	RVNGInputStreamPtr &input=stream->m_input;
	long pos=input->tell();
	if (pos+4>endPos) return false;
	auto sSz=long(libwps::readU32(input));
	if (sSz<0 || sSz>endPos-pos-4)
	{
		WPS_DEBUG_MSG(("WPSOLEObject::readString: name size seems bad\n"));
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	if (sSz==0) return true;
	for (long i=0; i+1<sSz; ++i)
	{
		auto c=char(libwps::readU8(input));
		if (c)
			name+=c;
		else
			return false;
	}
	return libwps::readU8(input)==0;
}

bool WPSOLEObject::checkIsWMF(std::shared_ptr<WPSStream> stream, long endPos)
{
	RVNGInputStreamPtr &input=stream->m_input;
	long pos=input->tell();
	if (pos+18>endPos) return false;
	auto fType = int(libwps::read16(input));
	if ((fType!=1 && fType!=2) || libwps::read16(input)<9)
	{
		input->seek(pos, librevenge::RVNG_SEEK_SET);
		return false;
	}
	// seek version
	input->seek(2, librevenge::RVNG_SEEK_CUR);
	auto fSize = long(libwps::read32(input)); // size in words
	input->seek(pos, librevenge::RVNG_SEEK_SET);
	return (2*fSize>18 && 2*fSize<=endPos-pos);
}

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
