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

#ifndef QUATTRO_FORMULA_H
#define QUATTRO_FORMULA_H

#include <functional>
#include <ostream>
#include <string>
#include <vector>

#include <librevenge-stream/librevenge-stream.h>

#include "libwps_internal.h"

#include "WPSDebug.h"
#include "WKSContentListener.h"

namespace QuattroFormulaInternal
{
struct State;

/** small class use to store Quattro Pro cell reference (.wb1-3 and qpw) */
class CellReference
{
public:
	/// constructor
	CellReference()
		: m_cells()
	{
	}
	/// add an instruction
	void addInstruction(WKSContentListener::FormulaInstruction const &instr)
	{
		if (!m_cells.empty() && instr.m_type!=instr.F_Operator && m_cells.back().m_type!=instr.F_Operator)
		{
			WKSContentListener::FormulaInstruction sep;
			sep.m_type=sep.F_Operator;
			sep.m_content=";";
			m_cells.push_back(sep);
		}
		m_cells.push_back(instr);
	}
	/// return true if we have not read any reference
	bool empty() const
	{
		return m_cells.empty();
	}
	/// friend operator<<
	friend std::ostream &operator<<(std::ostream &o, CellReference const &ref);
	//! the list of instruction coding each cell's block
	std::vector<WKSContentListener::FormulaInstruction> m_cells;
};

}

/** a class to read formula in a wb1-wb3, qpw file */
class QuattroFormulaManager
{
public:
	typedef std::function<bool(std::shared_ptr<WPSStream> const &stream, long endPos,
	                           QuattroFormulaInternal::CellReference &ref,
	                           Vec2i const &cPos, int cSheet)> CellReferenceFunction;

	//! constructor, version=1: means .wb1-3, version=2 means .qpw
	QuattroFormulaManager(QuattroFormulaManager::CellReferenceFunction const &readCellReference,
	                      int version);

	//! add a dll's correspondance between an id and a name
	void addDLLIdName(int id, librevenge::RVNGString const &name, bool func1);
	/** reads a formula */
	bool readFormula(std::shared_ptr<WPSStream> const &stream, long endPos, Vec2i const &pos, int sheetId,
	                 std::vector<WKSContentListener::FormulaInstruction> &formula, std::string &error) const;
private:
	/** the formula state */
	std::shared_ptr<QuattroFormulaInternal::State> m_state;
};
#endif /* QUATTRO_FORMULA_H */
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
