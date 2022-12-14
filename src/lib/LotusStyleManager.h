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

#ifndef LOTUS_STYLE_MANAGER_H
#define LOTUS_STYLE_MANAGER_H

#include <ostream>
#include <set>
#include <vector>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"

#include "WPSDebug.h"

namespace LotusStyleManagerInternal
{
struct CellStyleEntry;
struct State;
}

class LotusParser;

/**
 * This class parses the Lotus style
 *
 */
class LotusStyleManager
{
public:
	friend class LotusParser;

	//! constructor
	explicit LotusStyleManager(LotusParser &parser);
	//! destructor
	~LotusStyleManager();
	//! clean internal state
	void cleanState();
	//! update the state (need to be called before asking for style)
	void updateState();

	//! returns if possible the color(id between 0 and 7)
	bool getColor8(int cId, WPSColor &color) const;
	//! returns if possible the color(id between 0 and 15)
	bool getColor16(int cId, WPSColor &color) const;
	//! returns if possible the color(id between 0 and 255)
	bool getColor256(int cId, WPSColor &color) const;

	//! returns the pattern corresponding to a pattern id (id between 1 and 48)
	bool getPattern48(int id, WPSGraphicStyle::Pattern &pattern) const;
	//! returns the pattern corresponding to a pattern id (id between 1 and 64)
	bool getPattern64(int id, WPSGraphicStyle::Pattern &pattern) const;

	//! update a cell format using the cell id
	bool updateCellStyle(int cellId, WPSCellFormat &format,
	                     WPSFont &font, libwps_tools_win::Font::Type &fontType);
	//! update a font using the font id
	bool updateFontStyle(int fontId, WPSFont &font, libwps_tools_win::Font::Type &fontType);
	//! update style using line id
	bool updateLineStyle(int lineId, WPSGraphicStyle &style) const;
	//! update style using color id
	bool updateSurfaceStyle(int colorId, WPSGraphicStyle &style) const;
	//! update style using graphic id
	bool updateGraphicStyle(int graphicId, WPSGraphicStyle &style) const;
	//! update style using front/back color and pattern id
	bool updateSurfaceStyle(int fColorId, int bColorId, int patternId, WPSGraphicStyle &style) const;
protected:
	//! return the file version
	int version() const;

	//
	// low level
	//

	//! reads a color style
	bool readColorStyle(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a format style
	bool readFormatStyle(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a line style
	bool readLineStyle(std::shared_ptr<WPSStream> stream, long endPos, int vers);
	//! reads a graphic style
	bool readGraphicStyle(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a graphic style: fc9, lotus123
	bool readGraphicStyleC9(std::shared_ptr<WPSStream> stream, long endPos);

	// 1b style

	//! reads a font style: fa0
	bool readFontStyleA0(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a font style: ff0 (wk4)
	bool readFontStyleF0(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a cell style: fd2 (mac 123 or 123)
	bool readCellStyleD2(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a cell style: fe6 (wk4)
	bool readCellStyleE6(std::shared_ptr<WPSStream> stream, long endPos);

	//! reads the list of style: 32e7 (wk4)
	static bool readMenuStyleE7(std::shared_ptr<WPSStream> stream, long endPos);
	//! reads a style: fe6 (123)
	bool readStyleE6(std::shared_ptr<WPSStream> stream, long endPos);

	// old fmt style

	//! reads a format font name: zones 0xae
	bool readFMTFontName(std::shared_ptr<WPSStream> stream);
	//! reads a format font sizes zones 0xaf and 0xb1
	bool readFMTFontSize(std::shared_ptr<WPSStream> stream);
	//! reads a format font id zone: 0xb0
	bool readFMTFontId(std::shared_ptr<WPSStream> stream);

	//! update style using color id for defining shadow
	bool updateShadowStyle(int colorId, WPSGraphicStyle &style) const;

	//
	// low level
	//

	/** really reads the cell style: fd2 (123)

		\note this function normally called when all the styles are read, so it can reliably recover its data
	 */
	bool readCellStyleD2Data(LotusStyleManagerInternal::CellStyleEntry const &entry, std::set<int> &seen);
private:
	LotusStyleManager(LotusStyleManager const &orig) = delete;
	LotusStyleManager &operator=(LotusStyleManager const &orig) = delete;
	//! the main parser
	LotusParser &m_mainParser;
	//! the internal state
	std::shared_ptr<LotusStyleManagerInternal::State> m_state;
};

#endif /* LOTUS_STYLE_MANAGER_H */
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
