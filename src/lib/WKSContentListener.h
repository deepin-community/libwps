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
 * Copyright (C) 2006 Fridrich Strba (fridrich.strba@bluewin.ch)
 * Copyright (C) 2003-2005 William Lachance (william.lachance@sympatico.ca)
 * Copyright (C) 2003 Marc Maurer (uwog@uwog.net)
 *
 * For minor contributions see the git repository.
 *
 * Alternatively, the contents of this file may be used under the terms
 * of the GNU Lesser General Public License Version 2.1 or later
 * (LGPLv2.1+), in which case the provisions of the LGPLv2.1+ are
 * applicable instead of those above.
 *
 * For further information visit http://libwps.sourceforge.net
 */

#ifndef WKSCONTENTLISTENER_H
#define WKSCONTENTLISTENER_H

#include <vector>

#include <librevenge/librevenge.h>

#include "libwps_internal.h"

#include "WPSEntry.h"
#include "WPSGraphicStyle.h"

#include "WPSListener.h"

class WPSCellFormat;
class WKSChart;
struct WPSColumnFormat;
class WPSGraphicShape;
class WPSGraphicStyle;
class WPSList;
class WPSPageSpan;
struct WPSParagraph;
struct WPSRowFormat;
struct WPSTabStop;

struct WKSContentParsingState;
struct WKSDocumentParsingState;

class WKSContentListener final : public WPSListener
{
public:
	//! small class use to define a formula instruction
	struct FormulaInstruction
	{
		enum What { F_Operator, F_Function, F_Cell, F_CellList, F_Long, F_Double, F_Text };
		//! constructor
		FormulaInstruction()
			: m_type(F_Text)
			, m_content()
			, m_longValue(0)
			, m_doubleValue(0)
			, m_fileName()
		{
			for (auto &pos : m_position) pos=Vec2i(0,0);
			for (auto &pos : m_positionRelative) pos=Vec2b(false,false);
			for (auto &id : m_sheetId) id=-1;
		}
		//! return a proplist corresponding to a instruction
		librevenge::RVNGPropertyList getPropertyList() const;
		//! operator<<
		friend std::ostream &operator<<(std::ostream &o, FormulaInstruction const &inst);
		//! the type
		What m_type;
		//! the content ( if type == F_Operator or type = F_Function or type==F_Text)
		std::string m_content;
		//! value ( if type==F_Long )
		double m_longValue;
		//! value ( if type==F_Double )
		double m_doubleValue;
		//! cell position ( if type==F_Cell or F_CellList )
		Vec2i m_position[2];
		//! relative cell position ( if type==F_Cell or F_CellList )
		Vec2b m_positionRelative[2];
		//! the sheet name
		librevenge::RVNGString m_sheetName[2];
		/** the sheet id

			\note local field which can be used to store the sheet id
			before setting the sheet name */
		int m_sheetId[2];
		//! the file name (external reference)
		librevenge::RVNGString m_fileName;
	};
	//! small class use to define a sheet cell content
	struct CellContent
	{
		/** the different types of cell's field */
		enum ContentType { C_NONE, C_TEXT, C_NUMBER, C_FORMULA, C_UNKNOWN };
		/// constructor
		CellContent()
			: m_contentType(C_UNKNOWN)
			, m_value(0.0)
			, m_valueSet(false)
			, m_textEntry()
			, m_formula() { }
		CellContent(CellContent const &)=default;
		CellContent &operator=(CellContent const &)=default;
		/// destructor
		~CellContent() {}
		//! operator<<
		friend std::ostream &operator<<(std::ostream &o, CellContent const &cell);

		//! returns true if the cell has no content
		bool empty() const
		{
			if (m_contentType == C_NUMBER) return false;
			if (m_contentType == C_TEXT && !m_textEntry.valid()) return false;
			if (m_contentType == C_FORMULA && (m_formula.size() || isValueSet())) return false;
			return true;
		}
		//! sets the double value
		void setValue(double value)
		{
			m_value = value;
			m_valueSet = true;
		}
		//! returns true if the value has been setted
		bool isValueSet() const
		{
			return m_valueSet;
		}
		//! returns true if the text is set
		bool hasText() const
		{
			return m_textEntry.valid();
		}
		/** conversion beetween double days since 1900 and date */
		static bool double2Date(double val, int &Y, int &M, int &D);
		/** conversion beetween double: second since 0:00 and time */
		static bool double2Time(double val, int &H, int &M, int &S);

		//! the content type ( by default unknown )
		ContentType m_contentType;
		//! the cell value
		double m_value;
		//! true if the value has been set
		bool m_valueSet;
		//! the cell string
		WPSEntry m_textEntry;
		//! the formula list of instruction
		std::vector<FormulaInstruction> m_formula;
	};

	WKSContentListener(std::vector<WPSPageSpan> const &pageList, librevenge::RVNGSpreadsheetInterface *documentInterface);
	~WKSContentListener() final;

	void setDocumentLanguage(int lcid) final;
	void setMetaData(const librevenge::RVNGPropertyList &list);

	void startDocument();
	void endDocument();
	void handleSubDocument(WPSSubDocumentPtr &subDocument, libwps::SubDocumentType subDocumentType);

	// ------ text data -----------

	//! adds a basic character, ..
	void insertCharacter(uint8_t character) final;
	/** adds an unicode character
	 *
	 * by convention if \a character=0xfffd(undef), no character is added */
	void insertUnicode(uint32_t character) final;
	//! adds a unicode string
	void insertUnicodeString(librevenge::RVNGString const &str) final;

	void insertTab() final;
	void insertEOL(bool softBreak=false) final;
	void insertBreak(const uint8_t breakType) final;

	// ------ text format -----------
	//! set the actual font
	void setFont(const WPSFont &font) final;
	//! returns the actual font
	WPSFont const &getFont() const final;

	// ------ paragraph format -----------
	//! returns true if a paragraph or a list is opened
	bool isParagraphOpened() const final;
	//! sets the actual paragraph
	void setParagraph(const WPSParagraph &para) final;
	//! returns the actual paragraph
	WPSParagraph const &getParagraph() const final;

	// ------- fields ----------------
	//! adds a field
	void insertField(WPSField const &field) final;

	// ------- subdocument -----------------
	/** adds comment */
	void insertComment(WPSSubDocumentPtr &subDocument);
	/** adds a picture in given position */
	void insertPicture(WPSPosition const &pos, const librevenge::RVNGBinaryData &binaryData,
	                   std::string type="image/pict", WPSGraphicStyle const &style=WPSGraphicStyle::emptyStyle());
	/** adds an object with replacement picture in given position */
	void insertObject(WPSPosition const &pos, const WPSEmbeddedObject &obj,
	                  WPSGraphicStyle const &style=WPSGraphicStyle::emptyStyle());
	/** adds a picture in given position */
	void insertPicture(WPSPosition const &pos, WPSGraphicShape const &shape, WPSGraphicStyle const &style);
	/** adds a textbox in given position */
	void insertTextBox(WPSPosition const &pos, WPSSubDocumentPtr subDocument,
	                   WPSGraphicStyle const &frameStyle=WPSGraphicStyle::emptyStyle());
	/** open a group (not implemented) */
	bool openGroup(WPSPosition const &pos) final;
	/** close a group (not implemented) */
	void closeGroup() final;

	// ------- sheet -----------------
	/** open a sheet*/
	void openSheet(std::vector<WPSColumnFormat> const &columns, librevenge::RVNGString const &name="");
	/** closes this sheet */
	void closeSheet();
	/** open a row */
	void openSheetRow(WPSRowFormat const &f, int numRepeated=1);
	/** closes this row */
	void closeSheetRow();
	/** low level function to define a cell.
		\param cell the cell position, alignement, ...
		\param content the cell content
		\param numRepeated the cell columns repeatition*/
	void openSheetCell(WPSCell const &cell, CellContent const &content, int numRepeated=1);
	/** close a cell */
	void closeSheetCell();

	// ------- chart -----------------
	/** adds a chart in given position */
	void insertChart(WPSPosition const &pos, WKSChart const &chart,
	                 WPSGraphicStyle const &style=WPSGraphicStyle::emptyStyle());
protected:
	void _openPageSpan();
	void _closePageSpan();

	void _handleFrameParameters(librevenge::RVNGPropertyList &propList, WPSPosition const &pos);
	bool _openFrame(WPSPosition const &pos, WPSGraphicStyle const &style);
	void _closeFrame();

	void _startSubDocument();
	void _endSubDocument();

	void _openParagraph();
	void _closeParagraph();
	void _appendParagraphProperties(librevenge::RVNGPropertyList &propList, const bool isListElement=false);
	void _resetParagraphState(const bool isListElement=false);

	void _openSpan();
	void _closeSpan();

	void _flushText();
	void _flushDeferredTabs();

	void _insertBreakIfNecessary(librevenge::RVNGPropertyList &propList);

	/** creates a new parsing state (copy of the actual state)
	 *
	 * \return the old one */
	std::shared_ptr<WKSContentParsingState> _pushParsingState();
	//! resets the previous parsing state
	void _popParsingState();

protected:
	std::shared_ptr<WKSDocumentParsingState> m_ds; // main parse state
	std::shared_ptr<WKSContentParsingState> m_ps; // parse state
	std::vector<std::shared_ptr<WKSContentParsingState> > m_psStack;
	librevenge::RVNGSpreadsheetInterface *m_documentInterface;

private:
	WKSContentListener(const WKSContentListener &) = delete;
	WKSContentListener &operator=(const WKSContentListener &) = delete;
};

#endif
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
