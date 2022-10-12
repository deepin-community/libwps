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

#ifndef LOTUS_GRAPH_H
#define LOTUS_GRAPH_H

#include <ostream>
#include <map>
#include <string>
#include <vector>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"

#include "WPSDebug.h"
#include "WKSContentListener.h"

namespace LotusGraphInternal
{
struct ZoneMac; // 123 mac
struct ZonePc; // 123 pc
struct ZonePcList; // 123 pc
struct ZoneWK4; // lotus 4
struct State;

class SubDocument;
}

class LotusParser;
class LotusStyleManager;

/**
 * This class parses Microsoft Works graph file
 *
 */
class LotusGraph
{
public:
	friend class LotusParser;
	friend class LotusGraphInternal::SubDocument;

	//! constructor
	explicit LotusGraph(LotusParser &parser);
	//! destructor
	~LotusGraph();
	//! clean internal state
	void cleanState();
	//! sets the listener
	void setListener(WKSContentListenerPtr &listen)
	{
		m_listener = listen;
	}
	/** update the state (need to be called before sending data)

		\param zIdToSheetIdMap the correspondance between a zone and a sheet: defined in .123
		\param nameToChartIdMap the name and the id of each chart's: used to retrieve the
		correspondance betweeen a graphic's chart and the chart in .wk4
	 */
	void updateState(std::map<int,int> const &zIdToSheetIdMap,
	                 std::map<std::string,int> const &nameToChartIdMap);
protected:
	//! return the file version
	int version() const;

	//! return true if the sheet sheetId has some graphic
	bool hasGraphics(int sheetId) const;
	//! send the graphics corresponding to a sheetId
	void sendGraphics(int sheetId);
	//! try to send a shape: 123 pc
	void sendZone(LotusGraphInternal::ZonePcList const &zoneList, size_t id, WPSTransformation &transf);
	//! try to send a picture: 123 mac
	void sendPicture(LotusGraphInternal::ZoneMac const &zone);
	//! try to send a textbox content's
	void sendTextBox(std::shared_ptr<WPSStream> stream, WPSEntry const &entry);
	//! try to send a textbox content's
	void sendTextBoxWK4(std::shared_ptr<WPSStream> stream, WPSEntry const &entry, bool isButton);

	//! sets the current chart id(interface with LotusChart)
	bool setChartId(int chartId);

	//
	// low level
	//

	// ////////////////////// zone //////////////////////////////

	// zone 1b

	//! reads a begin graphic zone: 2328 (wk3mac)
	bool readZoneBegin(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a graphic zone: 2332, 2346, 2350, 2352, 23f0 (wk3mac)
	bool readZoneData(std::shared_ptr<WPSStream> stream, long endPos, int type);
	//! reads a graphic textbox data: 23f0 (wk3mac)
	bool readTextBoxData(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a picture definition: 240e (wk3mac)
	bool readPictureDefinition(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a picture data: 2410 (wk3mac)
	bool readPictureData(std::shared_ptr<WPSStream> stream, long endPos);

	// fmt

	//! try to read the pict filename, ...: b7 (wk3-wk4 pc)
	bool readFMTPictName(std::shared_ptr<WPSStream> stream);
	//! try to read the sheet id: 0xc9 (wk4)
	bool readZoneBeginC9(std::shared_ptr<WPSStream> stream);
	//! try to read a graphic: 0xca (wk4)
	bool readGraphic(std::shared_ptr<WPSStream> stream);
	//! try to read a graph's frame: 0xcc (wk4)
	bool readFrame(std::shared_ptr<WPSStream> stream);
	//! reads a graphic textbox data: 0xd1 (wk4)
	bool readTextBoxDataD1(std::shared_ptr<WPSStream> stream);

	// 123 zone 3XX

	//! try to read the graphic zone: 1XXX
	bool readGraphZone(std::shared_ptr<WPSStream> stream, int zId);
	//! try to read the graphic data zone: 010d
	bool readGraphDataZone(std::shared_ptr<WPSStream> stream, long endPos);

private:
	LotusGraph(LotusGraph const &orig) = delete;
	LotusGraph &operator=(LotusGraph const &orig) = delete;
	std::shared_ptr<WKSContentListener> m_listener; /** the listener (if set)*/
	//! the main parser
	LotusParser &m_mainParser;
	//! the style manager
	std::shared_ptr<LotusStyleManager> m_styleManager;
	//! the internal state
	std::shared_ptr<LotusGraphInternal::State> m_state;
};

#endif /* LOTUS_GRAPH_H */
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
