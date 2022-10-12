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
 * Copyright (C) 2003-2005 William Lachance (william.lachance@sympatico.ca)
 * Copyright (C) 2003 Marc Maurer (uwog@uwog.net)
 *
 * For minor contributions see the git repository.
 *
 * Alternatively, the contents of this file may be used under the terms
 * of the GNU Lesser General Public License Version 2.1 or later
 * (LGPLv2.1+), in which case the provisions of the LGPLv2.1+ are
 * applicable instead of those above.
 */

#ifndef QUATTRO9_GRAPH_H
#define QUATTRO9_GRAPH_H

#include <ostream>
#include <map>
#include <vector>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"

#include "WPSDebug.h"
#include "WKSContentListener.h"
#include "WPSGraphicStyle.h"

namespace Quattro9GraphInternal
{
struct Graph;
struct Shape;
struct State;

class SubDocument;
}

class Quattro9Parser;

/**
 * This class parses Quattro9Pro graph file
 *
 */
class Quattro9Graph
{
public:
	friend class Quattro9Parser;
	friend class Quattro9GraphInternal::SubDocument;

	//! constructor
	explicit Quattro9Graph(Quattro9Parser &parser);
	//! destructor
	~Quattro9Graph();
	//! sets the listener
	void setListener(WKSContentListenerPtr &listen)
	{
		m_listener = listen;
	}
	//! clean internal state
	void cleanState();
	//! update the state (need to be called before sending data)
	void updateState();

	//! returns the color corresponding to an id
	bool getColor(int id, WPSColor &color) const;
	//! returns the pattern corresponding to a pattern id between 0 and 24
	bool getPattern(int id, WPSGraphicStyle::Pattern &pattern) const;
protected:
	//! return the file version
	int version() const;
	//! stores the OLE objets
	void storeObjects(std::map<librevenge::RVNGString,WPSEmbeddedObject> const &nameToObjectMap);

	//! send the page graphic corresponding to a sheet
	bool sendPageGraphics(int sheetId) const;
	//! send a shape
	bool sendShape(Quattro9GraphInternal::Graph const &graph, int sheetId) const;
	//! send a shape (recursif)
	bool sendShape(Quattro9GraphInternal::Shape const &shape, WPSTransformation const &transf) const;
	//! send a shape
	bool sendShape(WPSGraphicShape const &shape, WPSGraphicStyle const &style, WPSTransformation const &transf) const;
	//! send a OLE
	bool sendOLE(Quattro9GraphInternal::Graph const &graph, int sheetId) const;
	//! send a textbox
	bool sendTextbox(Quattro9GraphInternal::Graph const &graph, int sheetId) const;

	//
	// low level
	//

	//! read the begin/end graph zone: 1401/1402
	bool readBeginEnd(std::shared_ptr<WPSStream> stream, int sheetId);
	//! try to read a begin/end  zone: 2051
	bool readBeginEndZone(std::shared_ptr<WPSStream> const &stream);
	//! try to read a graph header zone: 2051
	bool readGraphHeader(std::shared_ptr<WPSStream> const &stream);
	//! try to read a frame style zone: 2131
	static bool readFrameStyle(std::shared_ptr<WPSStream> const &stream);
	//! try to read a frame style zone: 2141
	static bool readFramePattern(std::shared_ptr<WPSStream> const &stream);
	//! try to read a frame header zone: 2171
	static bool readFrameHeader(std::shared_ptr<WPSStream> const &stream);
	//! try to read the OLE name zone: 21d1
	bool readOLEName(std::shared_ptr<WPSStream> const &stream);
	//! try to read a shape zone: 2221, 23d1
	bool readShape(std::shared_ptr<WPSStream> const &stream);
	//! try to read a shape sub zone: 2221, 23d1
	bool readShapeRec(std::shared_ptr<WPSStream> const &stream, long endPos, Quattro9GraphInternal::Shape &shape, WPSGraphicStyle const &actStyle);

	//! try to read a textbox style zone: 2371
	bool readTextboxStyle(std::shared_ptr<WPSStream> const &stream);
	//! try to read a textbox text zone: 2372
	bool readTextboxText(std::shared_ptr<WPSStream> const &stream);

private:
	Quattro9Graph(Quattro9Graph const &orig) = delete;
	Quattro9Graph &operator=(Quattro9Graph const &orig) = delete;
	std::shared_ptr<WKSContentListener> m_listener; /** the listener (if set)*/
	//! the main parser
	Quattro9Parser &m_mainParser;
	//! the internal state
	std::shared_ptr<Quattro9GraphInternal::State> m_state;
};

#endif /* QUATTRO9_GRAPH_H */
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
