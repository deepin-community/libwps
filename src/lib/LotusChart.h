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

#ifndef LOTUS_CHART_H
#define LOTUS_CHART_H

#include <ostream>
#include <string>
#include <map>
#include <vector>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"

#include "WPSDebug.h"
#include "WKSChart.h"
#include "WKSContentListener.h"

#include "Lotus.h"

namespace LotusChartInternal
{
class Chart;
struct State;
}

class LotusStyleManager;
class WPSGraphicStyle;

/**
 * This class parses Microsoft Works chart file
 *
 */
class LotusChart
{
public:
	friend class LotusParser;
	friend class LotusChartInternal::Chart;

	//! constructor
	explicit LotusChart(LotusParser &parser);
	//! clean internal state
	void cleanState();
	//! update internal state (must be called one time before sending data)
	void updateState();
	//! destructor
	~LotusChart();
	//! sets the listener
	void setListener(WKSContentListenerPtr &listen)
	{
		m_listener = listen;
	}

protected:
	//! return the file version
	int version() const;

	//! return the number of chart
	int getNumCharts() const;
	//! returns a map chart name to chart id map
	std::map<std::string,int> getNameToChartIdMap() const;
	//! update a chart, so that it can be send
	void updateChart(LotusChartInternal::Chart &chart, int id);
	//! try to send the charts(for Windows .wk3 file)
	bool sendCharts();
	//! try to send a chart
	bool sendChart(int cId, WPSPosition const &pos, WPSGraphicStyle const &style);
	//! try to send the text
	bool sendText(std::shared_ptr<WPSStream> stream, WPSEntry const &entry);

	//! reads a structure which define a chart: 11
	bool readChart(std::shared_ptr<WPSStream> stream);
	//! reads the chart name or title: 12
	bool readChartName(std::shared_ptr<WPSStream> stream);

	// zone 1b

	//! reads a chart data: 2710 (wk3mac)
	bool readMacHeader(std::shared_ptr<WPSStream> stream, long endPos, int &chartId);
	//! reads a placement position(wk3mac): 2774
	bool readMacPlacement(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a legend position(wk3mac): 277e
	bool readMacLegend(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a plot area position(wk3mac): 2788
	bool readMacPlotArea(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads an axis style(wk3mac): 27d8
	bool readMacAxis(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a serie style(wk3mac): 27e2
	bool readMacSerie(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a 3D floor style(wk3mac): 2846
	bool readMacFloor(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a manual position(wk3mac): 2904
	bool readMacPosition(std::shared_ptr<WPSStream> stream, long endPos);

	//! reads a plot area style: 2a30 (unsure)
	bool readPlotArea(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a serie style: 2a31
	bool readSerie(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a serie name: 2a32 (serie 6-...)
	bool readSerieName(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a serie width style: 2a33
	static bool readSerieWidth(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a list of font style: 2a34
	static bool readFontsStyle(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a some frame styles: 2a35
	bool readFramesStyle(std::shared_ptr<WPSStream> stream, long endPos);

	//! convert a link zone in a chart position(if possible)
	bool convert(LotusParser::Link const &link, WKSChart::Position(&positions)[2]) const;
private:
	LotusChart(LotusChart const &orig) = delete;
	LotusChart &operator=(LotusChart const &orig) = delete;
	std::shared_ptr<WKSContentListener> m_listener; /** the listener (if set)*/
	//! the main parser
	LotusParser &m_mainParser;
	//! the style manager
	std::shared_ptr<LotusStyleManager> m_styleManager;
	//! the internal state
	std::shared_ptr<LotusChartInternal::State> m_state;
};

#endif /* WPS4_H */
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
