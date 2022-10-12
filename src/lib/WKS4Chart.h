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

#ifndef WKS4_CHART_H
#define WKS4_CHART_H

#include <ostream>
#include <map>
#include <vector>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"

#include "WPSDebug.h"
#include "WKSContentListener.h"

namespace WKS4ChartInternal
{
class Chart;
struct State;

}

class WKS4Parser;

/**
 * This class parses Microsoft Works chart file
 *
 */
class WKS4Chart
{
public:
	friend class WKS4Parser;
	friend class WKS4ChartInternal::Chart;

	//! constructor
	explicit WKS4Chart(WKS4Parser &parser);
	//! destructor
	~WKS4Chart();
	//! sets the listener
	void setListener(WKSContentListenerPtr &listen)
	{
		m_listener = listen;
	}

protected:
	//! return true if the pos is in the file, update the file size if need
	bool checkFilePosition(long pos);
	//! return the file version
	int version() const;

	/// reset the main input
	void resetInput(RVNGInputStreamPtr const &input);

	//! return the number of chart
	int getNumCharts() const;
	//! update a chart, so that it can be send
	void updateChart(WKS4ChartInternal::Chart &chart);

	//! try to send the charts
	bool sendCharts();
	//! try to send the text
	bool sendText(WPSEntry const &entry);

	//! reads a structure which define a chart: 2d(default), 2e(name + value)
	bool readChart();

	//! reads the axis(or second yaxis) data: zone 0x14
	bool readChartAxis();

	//! reads a list of series definition: zone 0x15
	bool readChartSeries();

	//! reads the series style: zone 0x16
	bool readChartSeriesStyles();

	//! reads the chart's series color map: zone 0x31
	bool readChartSeriesColorMap();
	//! reads the chart name or title: 41
	bool readChartName();

	//! reads a structure which seems to define some dimension (only present in windows file): 35
	bool readChartDim();
	//! reads a structure which seems to define two chart font (only present in windows file): 40
	bool readChartFont();
	//! reads a structure which stores zfront, zfar: 44
	bool readChart3D();

	//! reads a structure which seems to define four chart font (only present in windows file): 84
	bool readChart2Font();

	//! reads end/begin of chart (only present in windows file): 80,81
	bool readChartLimit();

private:
	WKS4Chart(WKS4Chart const &orig) = delete;
	WKS4Chart &operator=(WKS4Chart const &orig) = delete;
	//! returns the debug file
	libwps::DebugFile &ascii()
	{
		return m_asciiFile;
	}
	/** the input */
	RVNGInputStreamPtr m_input;
	std::shared_ptr<WKSContentListener> m_listener; /** the listener (if set)*/
	//! the main parser
	WKS4Parser &m_mainParser;
	//! the internal state
	std::shared_ptr<WKS4ChartInternal::State> m_state;
	//! the ascii file
	libwps::DebugFile &m_asciiFile;
};

#endif /* WPS4_H */
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
