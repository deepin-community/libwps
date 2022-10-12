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
 * Copyright (C) 2004 Marc Maurer (uwog@uwog.net)
 * Copyright (C) 2004-2006 Fridrich Strba (fridrich.strba@bluewin.ch)
 *
 * For minor contributions see the git repository.
 *
 * Alternatively, the contents of this file may be used under the terms
 * of the GNU Lesser General Public License Version 2.1 or later
 * (LGPLv2.1+), in which case the provisions of the LGPLv2.1+ are
 * applicable instead of those above.
 */

#include <map>
#include <sstream>

#include "WPSStream.h"

#include "QuattroFormula.h"

/** namespace to regroup data used to read QuattroPro .wb1-3, .qpw formula */
namespace QuattroFormulaInternal
{
struct Functions
{
	char const *m_name;
	int m_arity;
};

struct State
{
	/** constructor */
	State(QuattroFormulaManager::CellReferenceFunction const &readCellReference, int version)
		: m_readCellReferenceFunction(readCellReference)
		, m_version(version)
		, m_idFunctionsMap()
		, m_idToDLLName1Map()
		, m_actDLLName1Id(-1)
		, m_idToDLLName2Map()
	{
		if (m_version>=2)
		{
			// in .qpw, H/VLookUp have four arguments
			m_idFunctionsMap =
			{
				{0x55, {"VLookUp", 4}},
				{0x5a, {"HLookup", 4}}
			};
		}
	}

	/** function to call to read a cell reference*/
	QuattroFormulaManager::CellReferenceFunction m_readCellReferenceFunction;
	/** the file version: 1: .wb1-3, 2: .qpw*/
	int m_version;
	/** the function which differs from default */
	std::map<int, Functions> m_idFunctionsMap;
	//! map id to DLL name 1
	std::map<int, librevenge::RVNGString> m_idToDLLName1Map;
	//! the current id DLL name 1
	int m_actDLLName1Id;
	//! map id to DLL name2
	std::map<Vec2i, librevenge::RVNGString> m_idToDLLName2Map;
};
}

// constructor
QuattroFormulaManager::QuattroFormulaManager(QuattroFormulaManager::CellReferenceFunction const &readCellReference, int version)
	: m_state(new QuattroFormulaInternal::State(readCellReference, version))
{
}

void QuattroFormulaManager::addDLLIdName(int id, librevenge::RVNGString const &name, bool func1)
{
	if (name.empty())
	{
		WPS_DEBUG_MSG(("QuattroFormulaManager::addDLLIdName: called with empty name for id=%d\n", id));
		return;
	}
	if (func1)
	{
		m_state->m_actDLLName1Id=id;
		auto &map = m_state->m_idToDLLName1Map;
		if (map.find(id) != map.end())
		{
			WPS_DEBUG_MSG(("QuattroFormulaManager::addDLLIdName: called with dupplicated id=%d\n", id));
		}
		else
			map[id]=name;
		return;
	}
	if (m_state->m_actDLLName1Id<0)
	{
		WPS_DEBUG_MSG(("QuattroFormulaManager::addDLLIdName: oops, unknown name1 id for %d\n", id));
		return;
	}
	auto &map = m_state->m_idToDLLName2Map;
	Vec2i fId(m_state->m_actDLLName1Id, id);
	if (map.find(fId) != map.end())
	{
		WPS_DEBUG_MSG(("QuattroFormulaManager::addDLLIdName: called with dupplicated id=%d,%d\n", m_state->m_actDLLName1Id, id));
	}
	else
		map[fId]=name;
	return;
}

//------------------------------------------------------------
// read a formula
//------------------------------------------------------------
namespace QuattroFormulaInternal
{
static Functions const s_listFunctions[] =
{
	// 0
	{ "", 0} /*SPEC: double*/, {"", 0}/*SPEC: cell*/, {"", 0}/*SPEC: cells*/, {"=", 1} /*SPEC: end of formula*/,
	{ "(", 1} /* SPEC: () */, {"", 0}/*SPEC: int*/, { "", -2} /*SPEC: text*/, {"", -2} /*SPEC: default argument*/,
	{ "-", 1}, {"+", 2}, {"-", 2}, {"*", 2},
	{ "/", 2}, { "^", 2}, {"=", 2}, {"<>", 2},

	// 1
	{ "<=", 2},{ ">=", 2},{ "<", 2},{ ">", 2},
	{ "And", 2},{ "Or", 2}, { "Not", 1}, { "+", 1},
	{ "&", 2}, { "", -2} /*halt*/, { "DLL", 0} /*DLL*/,{ "", -2} /*extended noop: 1b00011c020400020000000 means A*/,
	{ "", -2} /*extended op*/,{ "", -2} /*reserved*/,{ "", -2} /*reserved*/,{ "NA", 0} /*checkme*/,

	// 2
	{ "NA", 0} /* Error*/,{ "Abs", 1},{ "Int", 1},{ "Sqrt", 1},
	{ "Log10", 1},{ "Ln", 1},{ "Pi", 0},{ "Sin", 1},
	{ "Cos", 1},{ "Tan", 1},{ "Atan2", 2},{ "Atan", 1},
	{ "Asin", 1},{ "Acos", 1},{ "Exp", 1},{ "Mod", 2},

	// 3
	{ "Choose", -1},{ "IsNa", 1},{ "IsError", 1},{ "False", 0},
	{ "True", 0},{ "Rand", 0},{ "Date", 3},{ "Now", 0},
	{ "PMT", 3} /*BAD*/,{ "QPRO_PV", 3} /*BAD*/,{ "QPRO_FV", 3} /*BAD*/,{ "IF", 3},
	{ "Day", 1},{ "Month", 1},{ "Year", 1},{ "Round", 2},

	// 4
	{ "Time", 3},{ "Hour", 1},{ "Minute", 1},{ "Second", 1},
	{ "IsNumber", 1},{ "IsText", 1},{ "Len", 1},{ "Value", 1},
	{ "Fixed", 2}, { "Mid", 3}, { "Char", 1},{ "Ascii", 1},
	{ "Find", 3},{ "DateValue", 1} /*checkme*/,{ "TimeValue", 1} /*checkme*/,{ "CellPointer", 1} /*checkme*/,

	// 5
	{ "Sum", -1},{ "Average", -1},{ "COUNT", -1},{ "Min", -1},
	{ "Max", -1},{ "VLookUp", 3},{ "NPV", 2}, { "Var", -1},
	{ "StDev", -1},{ "IRR", 2} /*BAD*/, { "HLookup", 3},{ "DSum", 3},
	{ "DAverage", 3},{ "DCount", 3},{ "DMin", 3},{ "DMax", 3},

	// 6
	{ "DVar", 3},{ "DStd", 3},{ "Index", 3} /* index2d*/, { "Columns", 1},
	{ "Rows", 1},{ "Rept", 2},{ "Upper", 1},{ "Lower", 1},
	{ "Left", 2},{ "Right", 2},{ "Replace", 4}, { "Proper", 1},
	{ "Cell", 2},{ "Trim", 1},{ "Clean", 1},{ "IsText", 1},

	// 7
	{ "IsNonText", 1},{ "Exact", 2},{ "QPRO_Call", -2} /*UNKN*/,{ "Indirect", 1},
	{ "RRI", 3}, { "TERM", 3}, { "CTERM", 3}, { "SLN", 3},
	{ "SYD", 4},{ "DDB", 4}, { "StDevP", -1}, { "VarP", -1},
	{ "DBStdDevP", 3}, { "DBVarP", 3}, { "PV", 5}, { "PMT", 5},

	// 8
	{ "FV", 5}, { "Nper", 5}, { "Rate", 5}/*IRate*/, { "Ipmt", 6},
	{ "Ppmt", 6}, { "SumProduct", 2}, { "QPRO_MemAvail", 0}, { "QPRO_MememsAvail", 0},
	{ "QPRO_FileExist", 1}, { "QPRO_CurValue", 2}, { "Degrees", 1},{ "Radians", 1},
	{ "QPRO_Hex", 1},{ "QPRO_Num", 1},{ "Today", 0},{ "NPV", 2},

	// 9
	{ "QPRO_CellIndex", 4}, { "QPRO_Version", 0}, { "", -2} /*UNKN*/,{ "", -2} /*UNKN*/,
	{ "QPRO_Dhol", 3} /* fixme name: DHOL ?*/, { "", -2} /*UNKN*/, { "", -2} /*UNKN*/,{ "", -2} /*UNKN*/,
	{ "", -2} /*UNKN*/,{ "", -2} /*UNKN*/, { "Sheet", 1}, { "", -2} /*UNKN*/,
	{ "", -2} /*UNKN*/,{ "Index", 4}, { "QPRO_CellIndex3d", -2} /*UNKN*/,{ "QPRO_property", 1},

	// a
	{"QPRO_DDE", 4}, {"QPRO_Command", 1}, {"QPRO_Gerlinie", 3} /* fixme: name GERLINIE? */
};

}

bool QuattroFormulaManager::readFormula(std::shared_ptr<WPSStream> const &stream, long endPos,
                                        Vec2i const &position, int sheetId,
                                        std::vector<WKSContentListener::FormulaInstruction> &formula, std::string &error) const
{
	RVNGInputStreamPtr input = stream->m_input;
	libwps::DebugFile &ascFile=stream->m_ascii;
	formula.resize(0);
	error = "";
	long pos = input->tell();
	if (endPos - pos < 4) return false;
	auto sz = int(libwps::readU16(input)); // max 1024
	if (endPos-pos-4 != sz) return false;

	std::vector<QuattroFormulaInternal::CellReference> listCellsPos;
	auto fieldPos= int(libwps::readU16(input)); // ref begin
	if (fieldPos<0||fieldPos>sz)
	{
		WPS_DEBUG_MSG(("QuattroFormulaManager::readFormula: can not find the field header\n"));
		error="###fieldPos";
		return false;
	}
	if (fieldPos!=sz)
	{
		input->seek(pos+4+fieldPos, librevenge::RVNG_SEEK_SET);
		ascFile.addDelimiter(pos+4+fieldPos,'|');
		while (!input->isEnd())
		{
			long actPos=input->tell();
			if (actPos+4>endPos) break;
			QuattroFormulaInternal::CellReference cell;
			if (!m_state->m_readCellReferenceFunction(stream, endPos, cell, position, sheetId) || input->tell()<actPos+2)
			{
				input->seek(actPos, librevenge::RVNG_SEEK_SET);
				break;
			}
			if (cell.empty())
			{
				WPS_DEBUG_MSG(("QuattroFormulaManager::readFormula: find some deleted cells\n"));
			}
			else
				listCellsPos.push_back(cell);
			continue;
		}
		if (input->tell() !=endPos)
		{
			ascFile.addDelimiter(input->tell(),'@');
			static bool first=true;
			if (first)
			{
				WPS_DEBUG_MSG(("QuattroFormulaManager::readFormula: potential formula codes\n"));
				first=false;
			}
			error="###codes,";
		}
		input->seek(pos+4, librevenge::RVNG_SEEK_SET);
		endPos=pos+4+fieldPos;
	}
	std::stringstream f;
	std::vector<std::vector<WKSContentListener::FormulaInstruction> > stack;
	bool ok = true;
	size_t actCellId=0;
	int numDefault=0;
	while (long(input->tell()) != endPos)
	{
		double val;
		bool isNaN;
		pos = input->tell();
		if (pos > endPos) return false;
		auto wh = int(libwps::readU8(input));
		int arity = 0;
		WKSContentListener::FormulaInstruction instr;
		bool noInstr=false;
		switch (wh)
		{
		case 0x0:
			if (endPos-pos<9 || !libwps::readDouble8(input, val, isNaN))
			{
				f.str("");
				f << "###number";
				error=f.str();
				ok = false;
				break;
			}
			instr.m_type=WKSContentListener::FormulaInstruction::F_Double;
			instr.m_doubleValue=val;
			break;
		case 0x1:
			if (actCellId>=listCellsPos.size())
			{
				f.str("");
				f << "###unknCell" << actCellId;
				error=f.str();
				ok = false;
				break;
			}
			stack.push_back(listCellsPos[actCellId++].m_cells);
			noInstr=true;
			break;
		case 0x2:
			if (actCellId>=listCellsPos.size())
			{
				f.str("");
				f << "###unknListCell" << actCellId;
				error=f.str();
				ok = false;
				break;
			}
			stack.push_back(listCellsPos[actCellId++].m_cells);
			noInstr=true;
			break;
		case 0x5:
			instr.m_type=WKSContentListener::FormulaInstruction::F_Long;
			instr.m_longValue=long(libwps::read16(input));
			break;
		case 0x6:
			instr.m_type=WKSContentListener::FormulaInstruction::F_Text;
			while (!input->isEnd())
			{
				if (input->tell() >= endPos)
				{
					ok=false;
					break;
				}
				auto c = char(libwps::readU8(input));
				if (c==0) break;
				instr.m_content += c;
			}
			break;
		case 0x7: // maybe default parameter
			++numDefault;
			noInstr=true;
			break;
		case 0x1a:
		{
			if (input->tell()+4 >= endPos)
			{
				ok=false;
				break;
			}
			static bool first=true;
			if (first)
			{
				WPS_DEBUG_MSG(("QuattroFormulaManager::readFormula: this file contains some DLL functions, the result can be bad\n"));
				first=false;
			}
			arity= int(libwps::read8(input));
			std::stringstream s;
			s << "DLL";
			int ids[2];
			for (auto &id : ids) id=int(libwps::readU16(input));
			s << "_";
			auto it1 = m_state->m_idToDLLName1Map.find(ids[0]);
			if (it1!= m_state->m_idToDLLName1Map.end())
				s << it1->second.cstr();
			else
			{
				WPS_DEBUG_MSG(("QuattroFormulaManager::readFormula: can not find DLL function0 name for id=%d\n", ids[0]));
				s << "F" << ids[0];
				f << "##DLLFunc0=" << ids[0] << ",";
			}
			s << "_";
			auto it2 = m_state->m_idToDLLName2Map.find(Vec2i(ids[0],ids[1]));
			if (it2!= m_state->m_idToDLLName2Map.end())
				s << it2->second.cstr();
			else
			{
				WPS_DEBUG_MSG(("QuattroFormulaManager::readFormula: can not find DLL function1 name for id=%d\n", ids[1]));
				s << "F" << ids[1];
				f << "##DLLFunc1=" << ids[1] << ",";
			}
			instr.m_type=WKSContentListener::FormulaInstruction::F_Function;
			instr.m_content=s.str();
			break;
		}
		default:
		{
			auto fIt=m_state->m_idFunctionsMap.find(wh);
			if (fIt!=m_state->m_idFunctionsMap.end())
			{
				instr.m_type=WKSContentListener::FormulaInstruction::F_Function;
				instr.m_content=fIt->second.m_name;
				arity = fIt->second.m_arity;
			}
			else if (unsigned(wh) >= WPS_N_ELEMENTS(QuattroFormulaInternal::s_listFunctions) || QuattroFormulaInternal::s_listFunctions[wh].m_arity == -2)
			{
				f.str("");
				f << "##Funct" << std::hex << wh;
				error=f.str();
				ok = false;
				break;
			}
			else
			{
				instr.m_type=WKSContentListener::FormulaInstruction::F_Function;
				instr.m_content=QuattroFormulaInternal::s_listFunctions[wh].m_name;
				arity = QuattroFormulaInternal::s_listFunctions[wh].m_arity;
			}
			ok=!instr.m_content.empty();
			if (arity == -1) arity = int(libwps::read8(input));
			break;
		}
		}

		if (!ok) break;
		if (noInstr) continue;
		std::vector<WKSContentListener::FormulaInstruction> child;
		if (instr.m_type!=WKSContentListener::FormulaInstruction::F_Function)
		{
			child.push_back(instr);
			stack.push_back(child);
			continue;
		}
		size_t numElt = stack.size();
		arity-=numDefault;
		numDefault=0;
		if (arity<0 || int(numElt) < arity)
		{
			f.str("");
			f << instr.m_content << "[##" << arity << "]";
			error=f.str();
			ok = false;
			break;
		}
		//
		// first treat the special cases
		//
		if (arity==3 && instr.m_type==WKSContentListener::FormulaInstruction::F_Function && instr.m_content=="TERM")
		{
			// @TERM(pmt,pint,fv) -> NPER(pint,-pmt,pv=0,fv)
			auto pmt=stack[size_t(int(numElt)-3)];
			auto pint=stack[size_t(int(numElt)-2)];
			auto fv=stack[size_t(int(numElt)-1)];

			stack.resize(size_t(++numElt));
			// pint
			stack[size_t(int(numElt)-4)]=pint;
			//-pmt
			auto &node=stack[size_t(int(numElt)-3)];
			instr.m_type=WKSContentListener::FormulaInstruction::F_Operator;
			instr.m_content="-";
			node.resize(0);
			node.push_back(instr);
			instr.m_content="(";
			node.push_back(instr);
			node.insert(node.end(), pmt.begin(), pmt.end());
			instr.m_content=")";
			node.push_back(instr);
			//pv=zero
			instr.m_type=WKSContentListener::FormulaInstruction::F_Long;
			instr.m_longValue=0;
			stack[size_t(int(numElt)-2)].resize(0);
			stack[size_t(int(numElt)-2)].push_back(instr);
			//fv
			stack[size_t(int(numElt)-1)]=fv;
			arity=4;
			instr.m_type=WKSContentListener::FormulaInstruction::F_Function;
			instr.m_content="NPER";
		}
		else if (arity==3 && instr.m_type==WKSContentListener::FormulaInstruction::F_Function && instr.m_content=="CTERM")
		{
			// @CTERM(pint,fv,pv) -> NPER(pint,pmt=0,-pv,fv)
			auto pint=stack[size_t(int(numElt)-3)];
			auto fv=stack[size_t(int(numElt)-2)];
			auto pv=stack[size_t(int(numElt)-1)];
			stack.resize(size_t(++numElt));
			// pint
			stack[size_t(int(numElt)-4)]=pint;
			// pmt=0
			instr.m_type=WKSContentListener::FormulaInstruction::F_Long;
			instr.m_longValue=0;
			stack[size_t(int(numElt)-3)].resize(0);
			stack[size_t(int(numElt)-3)].push_back(instr);
			// -pv
			auto &node=stack[size_t(int(numElt)-2)];
			instr.m_type=WKSContentListener::FormulaInstruction::F_Operator;
			instr.m_content="-";
			node.resize(0);
			node.push_back(instr);
			instr.m_content="(";
			node.push_back(instr);
			node.insert(node.end(), pv.begin(), pv.end());
			instr.m_content=")";
			node.push_back(instr);

			//fv
			stack[size_t(int(numElt)-1)]=fv;
			arity=4;
			instr.m_type=WKSContentListener::FormulaInstruction::F_Function;
			instr.m_content="NPER";
		}

		if ((instr.m_content[0] >= 'A' && instr.m_content[0] <= 'Z') || instr.m_content[0] == '(')
		{
			if (instr.m_content[0] != '(')
				child.push_back(instr);

			instr.m_type=WKSContentListener::FormulaInstruction::F_Operator;
			instr.m_content="(";
			child.push_back(instr);
			for (int i = 0; i < arity; i++)
			{
				if (i)
				{
					instr.m_content=";";
					child.push_back(instr);
				}
				auto const &node=stack[size_t(int(numElt)-arity+i)];
				child.insert(child.end(), node.begin(), node.end());
			}
			instr.m_content=")";
			child.push_back(instr);

			stack.resize(size_t(int(numElt)-arity+1));
			stack[size_t(int(numElt)-arity)] = child;
			continue;
		}
		if (arity==1)
		{
			instr.m_type=WKSContentListener::FormulaInstruction::F_Operator;
			stack[numElt-1].insert(stack[numElt-1].begin(), instr);
			if (wh==3)
				break;
			continue;
		}
		if (arity==2)
		{
			instr.m_type=WKSContentListener::FormulaInstruction::F_Operator;
			stack[numElt-2].push_back(instr);
			stack[numElt-2].insert(stack[numElt-2].end(), stack[numElt-1].begin(), stack[numElt-1].end());
			stack.resize(numElt-1);
			continue;
		}
		ok=false;
		error = "### unexpected arity";
		break;
	}

	if (!ok) ;
	else if (stack.size()==1 && stack[0].size()>1 && stack[0][0].m_content=="=")
	{
		formula.insert(formula.begin(),stack[0].begin()+1,stack[0].end());
		if (input->tell()!=endPos)
		{
			// unsure, find some text here, maybe some note
			static bool first=true;
			if (first)
			{
				WPS_DEBUG_MSG(("QuattroFormulaManager::readFormula: find some extra data\n"));
				first=false;
			}
			error="##extra data";
			ascFile.addDelimiter(input->tell(),'#');
		}
		return true;
	}
	else
		error = "###stack problem";

	static bool first = true;
	if (first)
	{
		WPS_DEBUG_MSG(("QuattroFormulaManager::readFormula: I can not read some formula\n"));
		first = false;
	}

	f.str("");
	for (auto const &i : stack)
	{
		for (auto const &j : i)
			f << j << ",";
		f << "@";
	}
	f << error << "###";
	error = f.str();
	return false;
}
////////////////////////////////////////////////////////////
// cell reference
////////////////////////////////////////////////////////////
namespace QuattroFormulaInternal
{
std::ostream &operator<<(std::ostream &o, CellReference const &ref)
{
	if (ref.m_cells.size()==1)
	{
		o << ref.m_cells[0];
		return o;
	}
	o << "[";
	for (auto const &r: ref.m_cells) o << r;
	o << "]";
	return o;
}
}

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */

