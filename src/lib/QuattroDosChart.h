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

#ifndef QUATTRO_DOS_CHART_H
#define QUATTRO_DOS_CHART_H

#include <ostream>
#include <map>
#include <vector>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"

#include "WPSDebug.h"
#include "WKSContentListener.h"

namespace QuattroDosChartInternal
{
class Chart;
struct State;

}

class QuattroDosParser;

/**
 * This class parses Quattro Pro DOS chart file
 *
 */
class QuattroDosChart
{
public:
	friend class QuattroDosParser;
	friend class QuattroDosChartInternal::Chart;

	//! constructor
	explicit QuattroDosChart(QuattroDosParser &parser);
	//! destructor
	~QuattroDosChart();
	//! sets the listener
	void setListener(WKSContentListenerPtr &listen)
	{
		m_listener = listen;
	}
	//! reads the chart type: b8(3d) or ca(bubble)
	bool readChartSetType();
	//! reads the chart name: b9
	bool readChartName();
	//! reads a structure which define a chart: 2d(default), 2e(name + value)
	bool readChart();

	//! returns the number of spreadsheet
	int getNumSpreadsheets() const;
	//! returns the list of cell's begin to end corresponding to a chart
	void getChartPositionMap(int sheetId, std::map<Vec2i,Vec2i> &cellMap) const;
	//! try to send the chart corresponding to sheetId and a position
	bool sendChart(int sheetId, Vec2i const &cell, Vec2f const &chartSize);
protected:
	//! return true if the pos is in the file, update the file size if need
	bool checkFilePosition(long pos);
	//! return the file version
	int version() const;

	//! try to send the text
	bool sendText(WPSEntry const &entry);
private:
	QuattroDosChart(QuattroDosChart const &orig) = delete;
	QuattroDosChart &operator=(QuattroDosChart const &orig) = delete;
	//! returns the debug file
	libwps::DebugFile &ascii()
	{
		return m_asciiFile;
	}
	/** the input */
	RVNGInputStreamPtr m_input;
	std::shared_ptr<WKSContentListener> m_listener; /** the listener (if set)*/
	//! the main parser
	QuattroDosParser &m_mainParser;
	//! the internal state
	std::shared_ptr<QuattroDosChartInternal::State> m_state;
	//! the ascii file
	libwps::DebugFile &m_asciiFile;
};

#endif /* WPS4_H */
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
