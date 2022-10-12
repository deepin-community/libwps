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

#ifndef QUATTRO_GRAPH_H
#define QUATTRO_GRAPH_H

#include <ostream>
#include <map>
#include <vector>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"

#include "WPSDebug.h"
#include "WKSContentListener.h"

namespace QuattroGraphInternal
{
struct Dialog;
struct Graph;
struct ShapeHeader;
struct Textbox;

struct State;

class SubDocument;
}

class QuattroParser;
class QuattroStyleManager;

/**
 * This class parses QuattroPro graph file
 *
 */
class QuattroGraph
{
public:
	friend class QuattroParser;
	friend class QuattroGraphInternal::SubDocument;

	//! constructor
	explicit QuattroGraph(QuattroParser &parser);
	//! destructor
	~QuattroGraph();
	//! sets the listener
	void setListener(WKSContentListenerPtr &listen)
	{
		m_listener = listen;
	}
	//! clean internal state
	void cleanState();
	//! update the state (need to be called before sending data)
	void updateState();

protected:
	//! return the file version
	int version() const;

	//! stores the OLE objets
	void storeObjects(std::map<librevenge::RVNGString,WPSEmbeddedObject> const &nameToObjectMap);
	//! return the list of cells containing some graphics in a sheet
	std::vector<Vec2i> getGraphicCellsInSheet(int sheetId) const;
	//! send the graphic corresponding to a cell
	bool sendGraphics(int sheetId, Vec2i const &cell) const;
	//! send the page graphic corresponding to a sheet
	bool sendPageGraphics(int sheetId) const;

	//! send a graphic
	bool sendGraphic(QuattroGraphInternal::Graph const &graph) const;
	//! send a shape
	bool sendShape(QuattroGraphInternal::Graph const &graph, int sheetId) const;
	//! send a textbox
	bool sendTextbox(QuattroGraphInternal::Graph const &graph, int sheetId) const;
	//! send the textbox content
	bool send(QuattroGraphInternal::Textbox const &textbox,
	          std::shared_ptr<WPSStream> stream) const;
	//
	// low level
	//

	// ////////////////////// zone //////////////////////////////

	//! read the begin/end graph zone: 321/322
	bool readBeginEnd(std::shared_ptr<WPSStream> stream, int sheetId);
	//! read a new OLE frame zone: zone 381(wb2)
	bool readFrameOLE(std::shared_ptr<WPSStream> stream);
	//! read a image zone: zone 382(wb2)
	bool readImage(std::shared_ptr<WPSStream> stream);
	//! read a bitmap zone: zone 383(wb2)
	bool readBitmap(std::shared_ptr<WPSStream> stream);
	//! read a chart zone: zone 384
	bool readChart(std::shared_ptr<WPSStream> stream);
	//! read an frame: zone 385
	bool readFrame(std::shared_ptr<WPSStream> stream);
	//! read a button: zone 386
	bool readButton(std::shared_ptr<WPSStream> stream);
	//! read a OLE data: zone 38b
	bool readOLEData(std::shared_ptr<WPSStream> stream);
	//! try to read a graph header
	static bool readHeader(QuattroGraphInternal::Graph &header, std::shared_ptr<WPSStream> stream, long endPos);

	// ////////////////////// shape //////////////////////////////

	//! read a shape header: zone 4d3
	bool readShape(std::shared_ptr<WPSStream> stream);
	/** read a shape header
	 \note the serie header is pretty similar */
	bool readShapeHeader(QuattroGraphInternal::ShapeHeader &shape, std::shared_ptr<WPSStream> stream,  long endPos);
	//! read a line/arrow data: zone 35a/37b
	bool readLine(std::shared_ptr<WPSStream> stream);
	//! read a oval/rect/round rect data: zone 33e/364/379
	bool readRect(std::shared_ptr<WPSStream> stream);
	//! read a polygon/polyline data: zone 35c/37c/35b/388
	bool readPolygon(std::shared_ptr<WPSStream> stream);
	//! read a textbox data: zone 36f
	bool readTextBox(std::shared_ptr<WPSStream> stream);

	// ////////////////////// dialog zone //////////////////////////////

	//! try to read a dialog header
	static bool readHeader(QuattroGraphInternal::Dialog &header, std::shared_ptr<WPSStream> stream, long endPos);
	//! read a dialog zone: 35e
	bool readDialog(std::shared_ptr<WPSStream> stream);
	//! read a unknown dialog zone: 335,343,345
	static bool readDialogUnknown(std::shared_ptr<WPSStream> stream);

	// ////////////////////// style //////////////////////////////
	bool readFillData(WPSGraphicStyle &style, int fillType, std::shared_ptr<WPSStream> stream, long endPos);
private:
	QuattroGraph(QuattroGraph const &orig) = delete;
	QuattroGraph &operator=(QuattroGraph const &orig) = delete;
	std::shared_ptr<WKSContentListener> m_listener; /** the listener (if set)*/
	//! the main parser
	QuattroParser &m_mainParser;
	//! the internal state
	std::shared_ptr<QuattroGraphInternal::State> m_state;
};

#endif /* QUATTRO_GRAPH_H */
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
