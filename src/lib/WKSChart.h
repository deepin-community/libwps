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

/*
 * Structure to store and construct a chart
 *
 */

#ifndef WKS_CHART
#  define WKS_CHART

#include <iostream>
#include <vector>
#include <map>

#include "libwps_internal.h"

#include "WPSEntry.h"
#include "WPSFont.h"
#include "WPSGraphicStyle.h"

namespace WKSChartInternal
{
class SubDocument;
}
/** a class used to store a chart associated to a spreadsheet .... */
class WKSChart
{
	friend class WKSChartInternal::SubDocument;
public:
	//! a cell position
	struct Position
	{
		//! constructor
		explicit Position(Vec2i pos=Vec2i(-1,-1), librevenge::RVNGString const &sheetName="")
			: m_pos(pos)
			, m_sheetName(sheetName)
		{
		}
		//! return true if the position is valid
		bool valid() const
		{
			return m_pos[0]>=0 && m_pos[1]>=0 && !m_sheetName.empty();
		}
		//! return true if the position is valid
		bool valid(Position const &maxPos) const
		{
			return valid() && maxPos.valid() && maxPos.m_pos[0]>=m_pos[0] && maxPos.m_pos[1]>=m_pos[1];
		}
		//! return the cell name
		librevenge::RVNGString getCellName() const;
		//! operator<<
		friend std::ostream &operator<<(std::ostream &o, Position const &pos);
		//! operator==
		bool operator==(Position const &pos) const
		{
			return m_pos==pos.m_pos && m_sheetName==pos.m_sheetName;
		}
		//! operator!=
		bool operator!=(Position const &pos) const
		{
			return !(operator==(pos));
		}
		//! the cell column and row
		Vec2i m_pos;
		//! the cell sheet name
		librevenge::RVNGString m_sheetName;
	};
	//! a axis in a chart
	struct Axis
	{
		//! the axis content
		enum Type { A_None, A_Numeric, A_Logarithmic, A_Sequence, A_Sequence_Skip_Empty };
		//! constructor
		Axis();
		//! destructor
		~Axis();
		//! add content to the propList
		void addContentTo(int coord, librevenge::RVNGPropertyList &propList) const;
		//! add style to the propList
		void addStyleTo(librevenge::RVNGPropertyList &propList) const;
		//! operator<<
		friend std::ostream &operator<<(std::ostream &o, Axis const &axis);
		//! the sequence type
		Type m_type;
		//! automatic scaling (or manual)
		bool m_automaticScaling;
		//! the minimum, maximum scaling(if manual)
		Vec2f m_scaling;
		//! show or not the grid
		bool m_showGrid;
		//! show or not the label
		bool m_showLabel;
		//! the label range if defined
		Position m_labelRanges[2];

		//! show or not the title/subtitle
		bool m_showTitle;
		//! the title cell range
		Position m_titleRange;
		//! the title label
		librevenge::RVNGString m_title;
		//! the subtitle label
		librevenge::RVNGString m_subTitle;
		//! the graphic style
		WPSGraphicStyle m_style;
	};
	//! a legend in a chart
	struct Legend
	{
		//! constructor
		Legend()
			: m_show(false)
			, m_autoPosition(true)
			, m_relativePosition(WPSBorder::RightBit)
			, m_position(0,0)
			, m_font()
			, m_style()
		{
		}
		//! add content to the propList
		void addContentTo(librevenge::RVNGPropertyList &propList) const;
		//! add style to the propList
		void addStyleTo(librevenge::RVNGPropertyList &propList) const;
		//! operator<<
		friend std::ostream &operator<<(std::ostream &o, Legend const &legend);
		//! show or not the legend
		bool m_show;
		//! automatic position
		bool m_autoPosition;
		//! the automatic position libwps::LeftBit|...
		int m_relativePosition;
		//! the position in points
		Vec2f m_position;
		//! the font
		WPSFont m_font;
		//! the graphic style
		WPSGraphicStyle m_style;
	};
	//! a serie in a chart
	struct Serie
	{
		//! the series type
		enum Type { S_Area, S_Bar, S_Bubble, S_Circle, S_Column, S_Gantt, S_Line, S_Radar, S_Ring, S_Scatter, S_Stock, S_Surface };
		//! the point type
		enum PointType
		{
			P_None=0, P_Automatic, P_Square, P_Diamond, P_Arrow_Down,
			P_Arrow_Up, P_Arrow_Right, P_Arrow_Left, P_Bow_Tie, P_Hourglass,
			P_Circle, P_Star, P_X, P_Plus, P_Asterisk,
			P_Horizontal_Bar, P_Vertical_Bar
		};
		//! constructor
		Serie();
		Serie(Serie const &)=default;
		Serie(Serie &&)=default;
		Serie &operator=(Serie const &)=default;
		Serie &operator=(Serie &&)=default;
		//! destructor
		virtual ~Serie();
		//! return true if the serie style is 1D
		bool is1DStyle() const
		{
			return m_type==S_Line || m_type==S_Radar || (m_type==S_Scatter && m_pointType==P_None);
		}
		//! set the primary color
		void setPrimaryColor(WPSColor const &color, float opacity = 1, bool force1D=false)
		{
			if (force1D || is1DStyle())
				m_style.m_lineColor=color;
			else
				m_style.setSurfaceColor(color, opacity);
		}
		//! set the primary pattern
		void setPrimaryPattern(WPSGraphicStyle::Pattern const &pattern, bool force1D=false);
		//! set the secondary color
		void setSecondaryColor(WPSColor const &color)
		{
			if (!is1DStyle())
				m_style.m_lineColor=color;
		}
		//! return true if the serie is valid
		bool valid() const
		{
			return m_ranges[0].valid(m_ranges[0]);
		}
		//! add content to the propList
		void addContentTo(librevenge::RVNGPropertyList &propList) const;
		//! add style to the propList
		void addStyleTo(librevenge::RVNGPropertyList &propList) const;
		//! returns a string corresponding to a series type
		static std::string getSerieTypeName(Type type);
		//! operator<<
		friend std::ostream &operator<<(std::ostream &o, Serie const &series);
		//! the type
		Type m_type;
		//! the data range
		Position m_ranges[2];
		//! use or not the secondary y axis
		bool m_useSecondaryY;
		//! the label font
		WPSFont m_font;
		//! the label ranges if defined(unused)
		Position m_labelRanges[2];
		//! the legend range if defined
		Position m_legendRange;
		//! the legend name if defined
		librevenge::RVNGString m_legendText;
		//! the graphic style
		WPSGraphicStyle m_style;
		//! the point type
		PointType m_pointType;
	};
	//! a text zone a chart
	struct TextZone
	{
		//! the text type
		enum Type { T_Title, T_SubTitle, T_Footer };
		//! the text content type
		enum ContentType { C_Cell, C_Text };

		//! constructor
		explicit TextZone(Type type);
		TextZone(TextZone const &)=default;
		//! destructor
		~TextZone();
		//! returns true if the textbox is valid
		bool valid() const
		{
			if (!m_show) return false;
			if (m_contentType==C_Cell)
				return m_cell.valid();
			if (m_contentType!=C_Text)
				return false;
			for (auto &e : m_textEntryList)
			{
				if (e.valid()) return true;
			}
			return false;
		}
		//! add content to the propList
		void addContentTo(librevenge::RVNGPropertyList &propList) const;
		//! add to the propList
		void addStyleTo(librevenge::RVNGPropertyList &propList) const;
		//! operator<<
		friend std::ostream &operator<<(std::ostream &o, TextZone const &zone);
		//! the zone type
		Type m_type;
		//! the content type
		ContentType m_contentType;
		//! true if the zone is visible
		bool m_show;
		//! the position in the zone
		Vec2f m_position;
		//! the cell position ( or title and subtitle)
		Position m_cell;
		//! the text entry (or the list of text entry)
		std::vector<WPSEntry> m_textEntryList;
		//! the zone format
		WPSFont m_font;
		//! the graphic style
		WPSGraphicStyle m_style;
	};

	//! the constructor
	explicit WKSChart(Vec2f const &dim=Vec2f());
	//! the destructor
	virtual ~WKSChart();
	//! send the chart to the listener
	void sendChart(WKSContentListenerPtr &listener, librevenge::RVNGSpreadsheetInterface *interface) const;
	//! send the zone content (called when the zone is of text type)
	virtual void sendContent(TextZone const &zone, WPSListenerPtr &listener) const=0;

	//! set the grid color
	void setGridColor(WPSColor const &color)
	{
		m_gridColor=color;
	}
	//! return an axis (corresponding to a coord)
	Axis &getAxis(int coord);
	//! return an axis (corresponding to a coord)
	Axis const &getAxis(int coord) const;

	//! returns the legend
	Legend const &getLegend() const
	{
		return m_legend;
	}
	//! returns the legend
	Legend &getLegend()
	{
		return m_legend;
	}

	//! return a serie
	Serie *getSerie(int id, bool create);
	//! returns the list of defined series
	std::map<int, Serie> const &getIdSerieMap() const
	{
		return m_serieMap;
	}
	//! returns a textzone content
	TextZone *getTextZone(TextZone::Type type, bool create=false);

protected:
	//! sends a textzone content
	void sendTextZoneContent(TextZone::Type type, WPSListenerPtr listener) const;

public:
	//! the chart dimension in point
	Vec2f m_dimension;
	//! the chart type (if no series)
	Serie::Type m_type;
	//! a flag to know if the data are stacked or not
	bool m_dataStacked;
	//! a flag to know if the data are percent stacked or not
	bool m_dataPercentStacked;
	//! a flag to know if the data are vertical (for bar)
	bool m_dataVertical;
	//! a flag to know if the graphic is 3D
	bool m_is3D;
	//! a flag to know if real 3D or 2D-extended
	bool m_is3DDeep;

	// main

	//! the chart style
	WPSGraphicStyle m_style;
	//! the chart name
	librevenge::RVNGString m_name;

	// plot area

	//! the plot area dimension in percent
	WPSBox2f m_plotAreaPosition;
	//! the ploat area style
	WPSGraphicStyle m_plotAreaStyle;

	// legend

	//! the legend dimension in percent
	WPSBox2f m_legendPosition;

	//! floor
	WPSGraphicStyle m_floorStyle;
	//! wall
	WPSGraphicStyle m_wallStyle;

protected:
	//! the grid color
	WPSColor m_gridColor;
	//! the x,y,y-second,z and a bad axis
	Axis m_axis[5];
	//! the legend
	Legend m_legend;
	//! the list of series
	std::map<int, Serie> m_serieMap;
	//! a map text zone type to text zone
	std::map<TextZone::Type, TextZone> m_textZoneMap;
private:
	explicit WKSChart(WKSChart const &orig) = delete;
	WKSChart &operator=(WKSChart const &orig) = delete;
};

#endif
/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
