/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: t; c-basic-offset: 4 -*- */
/* libwps
* Version: MPL 2.0 / LGPLv2+
*
* The contents of this file are subject to the Mozilla Public License Version
* 2.0 (the "License"); you may not use this file except in compliance with
* the License or as specified alternatively below. You may obtain a copy of
* the License at http://www.mozilla.org/MPL/
*
* Software distributed under the License is distributed on an "AS IS" basis,
* WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
* for the specific language governing rights and limitations under the
* License.
*
* Major Contributor(s):
* Copyright (C) 2002 William Lachance (wrlach@gmail.com)
* Copyright (C) 2002,2004 Marc Maurer (uwog@uwog.net)
* Copyright (C) 2004-2006 Fridrich Strba (fridrich.strba@bluewin.ch)
* Copyright (C) 2006, 2007 Andrew Ziem
* Copyright (C) 2011, 2012 Alonso Laurent (alonso@loria.fr)
*
*
* All Rights Reserved.
*
* For minor contributions see the git repository.
*
* Alternatively, the contents of this file may be used under the terms of
* the GNU Lesser General Public License Version 2 or later (the "LGPLv2+"),
* in which case the provisions of the LGPLv2+ are applicable
* instead of those above.
*/
#include <stdarg.h>
#include <string.h>

#include <algorithm>
#include <cstring>
#include <cstdio>
#include <iostream>
#include <map>
#include <vector>

#include "helper.h"

#include <librevenge/librevenge.h>
#include <libwps/libwps.h>

#ifndef __EMSCRIPTEN__
#  include <sys/stat.h>
#  include <sys/types.h>
#endif

namespace libwpsHelper
{
////////////////////////////////////////////////////////////
// debug
////////////////////////////////////////////////////////////
#ifdef DEBUG
void printDebugMsg(const char *format, ...)
{
	va_list args;
	va_start(args, format);
	std::vfprintf(stderr, format, args);
	va_end(args);
}
#endif

#ifndef __EMSCRIPTEN__
////////////////////////////////////////////////////////////
// static class to create a RVNGInputStream for some data
////////////////////////////////////////////////////////////

/** internal class used to create a structrured RVNGInputStream from two files
 */
class FolderStream final: public librevenge::RVNGInputStream
{
public:
	//! constructor
	FolderStream()
		: librevenge::RVNGInputStream()
		, m_nameToPathMap()
	{
	}

	//! destructor
	~FolderStream() override
	{
	}

	//! add a file
	void addFile(std::string const &path, std::string const &shortName)
	{
		m_nameToPathMap[shortName]=path;
	}
	/**! reads numbytes data.

	 * \return a pointer to the read elements
	 */
	const unsigned char *read(unsigned long, unsigned long &) final
	{
		return nullptr;
	}
	//! returns actual offset position
	long tell() final
	{
		return 0;
	}
	/*! \brief seeks to a offset position, from actual, beginning or ending position
	 * \return 0 if ok
	 */
	int seek(long, librevenge::RVNG_SEEK_TYPE) final
	{
		return 1;
	}
	//! returns true if we are at the end of the section/file
	bool isEnd() final
	{
		return true;
	}

	/** returns true if the stream is ole

	 \sa returns always false*/
	bool isStructured() final
	{
		return true;
	}
	/** returns the number of sub streams.

	 \sa returns always 2*/
	unsigned subStreamCount() final
	{
		return unsigned(m_nameToPathMap.size());
	}
	/** returns the ith sub streams name */
	const char *subStreamName(unsigned id) final
	{
		std::map<std::string, std::string>::const_iterator it=m_nameToPathMap.begin();
		for (unsigned i=0; i<id; ++i)
		{
			if (it==m_nameToPathMap.end()) return nullptr;
			++it;
		}
		if (it==m_nameToPathMap.end()) return nullptr;
		return it->first.c_str();
	}
	/** returns true if a substream with name exists */
	bool existsSubStream(const char *name) final
	{
		return name && m_nameToPathMap.find(name)!= m_nameToPathMap.end();
	}
	/** return a new stream for a ole zone */
	librevenge::RVNGInputStream *getSubStreamByName(const char *name) final;
	/** return a new stream for a ole zone */
	librevenge::RVNGInputStream *getSubStreamById(unsigned id) final
	{
		char const *name=subStreamName(id);
		if (name==nullptr) return nullptr;
		return getSubStreamByName(name);
	}
private:
	/// the map short name to path
	std::map<std::string, std::string> m_nameToPathMap;
	FolderStream(const FolderStream &) = delete; // copy is not allowed
	FolderStream &operator=(const FolderStream &) = delete; // assignment is not allowed
};

librevenge::RVNGInputStream *FolderStream::getSubStreamByName(const char *name)
{
	if (m_nameToPathMap.find(name)== m_nameToPathMap.end()) return nullptr;
	return new librevenge::RVNGFileStream(m_nameToPathMap.find(name)->second.c_str());
}

////////////////////////////////////////////////////////////
// main functions
////////////////////////////////////////////////////////////

/* check if the file is a lotus123 file and a .fm3 file exists or
   if the file is a dos lotus file and a .fmt file exists.
   If yes, try to convert it in a structured input which can be parsed by libwps */
static std::shared_ptr<librevenge::RVNGInputStream> createMergeInput(char const *fName, librevenge::RVNGInputStream &input)
try
{
	std::shared_ptr<FolderStream> res;

	/* we do not want to compress already compressed file.
	   So check if the file is structured, is a binhex file
	 */
	if (!fName || input.isStructured()) return res;

	// first check
	std::string name(fName);
	size_t len=name.length();
	if (len<=4 || name[len-4]!='.') return res;
	std::string extension=name.substr(len-3, 2);
	if (extension!="wk" && extension!="WK")
		return res;

	// check the file header
	if (input.seek(0, librevenge::RVNG_SEEK_SET)!=0) return res;
	unsigned long numBytesRead;
	const unsigned char *data=input.read(6, numBytesRead);
	if (!data || numBytesRead!=6 || data[0]!=0 || data[1]!=0 || data[3]!=0) return res;
	bool oldFile=false;
	if (data[2]==2 && data[4]==6 && data[5]==4)
		oldFile=true;
	else if (data[2]!=0x1a || data[4]>=2 || data[5]!=0x10) return res;

	// check if the .fm3 file exists
	std::string fmName=name.substr(0, len-3);
	if (extension=="wk")
		fmName+=oldFile ? "fmt" : "fm3";
	else
		fmName+=oldFile ? "FMT" : "FM3";
	struct stat status;
	if (stat(fmName.c_str(), &status) || !S_ISREG(status.st_mode))
		return res;

	res.reset(new FolderStream());
	if (oldFile)
	{
		res->addFile(name, "WK1");
		res->addFile(fmName, "FMT");
	}
	else
	{
		res->addFile(name, "WK3");
		res->addFile(fmName, "FM3");
	}
	return res;
}
catch (...)
{
	return std::shared_ptr<librevenge::RVNGInputStream>();
}
#endif

////////////////////////////////////////////////////////////
// main functions
////////////////////////////////////////////////////////////

std::shared_ptr<librevenge::RVNGInputStream> isSupported
(char const *filename, libwps::WPSConfidence &confidence, libwps::WPSKind &kind, bool &needEncoding)
{
	std::shared_ptr<librevenge::RVNGInputStream> input(new librevenge::RVNGFileStream(filename));
	libwps::WPSCreator creator;
#ifndef __EMSCRIPTEN__
	try
	{
		auto mergeInput=createMergeInput(filename, *input);
		if (mergeInput)
		{
			confidence=libwps::WPSDocument::isFileFormatSupported(mergeInput.get(), kind, creator, needEncoding);
			if (confidence != libwps::WPS_CONFIDENCE_NONE)
				return mergeInput;
		}
	}
	catch (...)
	{
	}
#endif
	try
	{
		confidence = libwps::WPSDocument::isFileFormatSupported(input.get(), kind, creator, needEncoding);
		if (confidence != libwps::WPS_CONFIDENCE_NONE)
			return input;
	}
	catch (...)
	{
	}
	return std::shared_ptr<librevenge::RVNGInputStream>();
}

bool checkErrorAndPrintMessage(libwps::WPSResult result)
{
	if (result == libwps::WPS_ENCRYPTION_ERROR)
		fprintf(stderr, "ERROR: Encrypted file, bad Password!\n");
	else if (result == libwps::WPS_FILE_ACCESS_ERROR)
		fprintf(stderr, "ERROR: File Exception!\n");
	else if (result == libwps::WPS_PARSE_ERROR)
		fprintf(stderr, "ERROR: Parse Exception!\n");
	else if (result == libwps::WPS_OLE_ERROR)
		fprintf(stderr, "ERROR: File is an OLE document, but does not contain a Microsoft Works stream!\n");
	else if (result != libwps::WPS_OK)
		fprintf(stderr, "ERROR: Unknown Error!\n");
	else
		return false;
	return true;
}

}
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
