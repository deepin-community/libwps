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
 * Structure to store and construct a chart from an unstructured list
 * of cell
 *
 */

#include <algorithm>
#include <iomanip>
#include <iostream>
#include <map>
#include <sstream>

#include <librevenge/librevenge.h>

#include "WPSListener.h"
#include "WPSPosition.h"
#include "WKSContentListener.h"
#include "WKSSubDocument.h"

#include "WKSChart.h"

/** Internal: the structures of a WKSChart */
namespace WKSChartInternal
{
////////////////////////////////////////
//! Internal: the subdocument of a WKSChart
class SubDocument final : public WKSSubDocument
{
public:
	SubDocument(WKSChart const *chart, WKSChart::TextZone::Type textZone)
		: WKSSubDocument(RVNGInputStreamPtr(), nullptr)
		, m_chart(chart)
		, m_textZone(textZone)
	{
	}

	//! destructor
	~SubDocument() final {}

	//! operator==
	bool operator==(std::shared_ptr<WPSSubDocument> const &doc) const final
	{
		if (!WKSSubDocument::operator==(doc) || !doc)
			return false;
		auto const *subDoc=dynamic_cast<SubDocument const *>(doc.get());
		if (!subDoc) return false;
		return m_chart==subDoc->m_chart && m_textZone==subDoc->m_textZone;
	}

	//! the parser function
	void parse(std::shared_ptr<WKSContentListener> &listener, libwps::SubDocumentType type) override;
protected:
	//! the chart
	WKSChart const *m_chart;
	//! the textzone type
	WKSChart::TextZone::Type m_textZone;
private:
	SubDocument(SubDocument const &orig) = delete;
	SubDocument &operator=(SubDocument const &orig) = delete;
};

void SubDocument::parse(std::shared_ptr<WKSContentListener> &listener, libwps::SubDocumentType /*type*/)
{
	if (!listener.get())
	{
		WPS_DEBUG_MSG(("WKSChartInternal::SubDocument::parse: no listener\n"));
		return;
	}

	if (!m_chart)
	{
		WPS_DEBUG_MSG(("WKSChartInternal::SubDocument::parse: can not find the chart\n"));
		return;
	}
	m_chart->sendTextZoneContent(m_textZone, listener);
}

}

////////////////////////////////////////////////////////////
// WKSChart
////////////////////////////////////////////////////////////
WKSChart::WKSChart(Vec2f const &dim)
	: m_dimension(dim)
	, m_type(WKSChart::Serie::S_Bar)
	, m_dataStacked(false)
	, m_dataPercentStacked(false)
	, m_dataVertical(false)
	, m_is3D(false)
	, m_is3DDeep(false)

	, m_style(WPSGraphicStyle::emptyStyle())
	, m_name()

	, m_plotAreaPosition()
	, m_plotAreaStyle(WPSGraphicStyle::emptyStyle())

	, m_legendPosition()

	, m_floorStyle()
	, m_wallStyle()

	, m_gridColor(179,179,179)
	, m_legend()
	, m_serieMap()
	, m_textZoneMap()
{
	m_wallStyle.m_lineColor=m_floorStyle.m_lineColor=WPSColor(0xb3,0xb3, 0xb3);
}

WKSChart::~WKSChart()
{
}

WKSChart::Axis &WKSChart::getAxis(int coord)
{
	if (coord<0 || coord>3)
	{
		WPS_DEBUG_MSG(("WKSChart::getAxis: called with bad coord\n"));
		return  m_axis[4];
	}
	return m_axis[coord];
}

WKSChart::Axis const &WKSChart::getAxis(int coord) const
{
	if (coord<0 || coord>3)
	{
		WPS_DEBUG_MSG(("WKSChart::getAxis: called with bad coord\n"));
		return  m_axis[4];
	}
	return m_axis[coord];
}

WKSChart::Serie *WKSChart::getSerie(int id, bool create)
{
	if (m_serieMap.find(id)!=m_serieMap.end())
		return &m_serieMap.find(id)->second;
	if (!create)
		return nullptr;
	m_serieMap[id]=WKSChart::Serie();
	return &m_serieMap.find(id)->second;
}

WKSChart::TextZone *WKSChart::getTextZone(WKSChart::TextZone::Type type, bool create)
{
	if (m_textZoneMap.find(type)!=m_textZoneMap.end())
		return &m_textZoneMap.find(type)->second;
	if (!create)
		return nullptr;
	m_textZoneMap.insert(std::map<TextZone::Type, TextZone>::value_type(type,WKSChart::TextZone(type)));
	return &m_textZoneMap.find(type)->second;
}

void WKSChart::sendTextZoneContent(WKSChart::TextZone::Type type, WPSListenerPtr listener) const
{
	if (m_textZoneMap.find(type)==m_textZoneMap.end())
	{
		WPS_DEBUG_MSG(("WKSChart::sendTextZoneContent: called with unknown zone(%d)\n", int(type)));
		return;
	}
	sendContent(m_textZoneMap.find(type)->second, listener);
}

void WKSChart::sendChart(WKSContentListenerPtr &listener, librevenge::RVNGSpreadsheetInterface *interface) const
{
	if (!listener || !interface)
	{
		WPS_DEBUG_MSG(("WKSChart::sendChart: can not find listener or interface\n"));
		return;
	}
	if (m_serieMap.empty())
	{
		WPS_DEBUG_MSG(("WKSChart::sendChart: can not find the series\n"));
		return;
	}
	std::shared_ptr<WPSListener> genericListener=listener;
	int styleId=0;

	librevenge::RVNGPropertyList style;
	style.insert("librevenge:chart-id", styleId);
	m_style.addTo(style);
	interface->defineChartStyle(style);

	librevenge::RVNGPropertyList chart;
	if (m_dimension[0]>0 && m_dimension[1]>0)
	{
		// set the size if known
		chart.insert("svg:width", double(m_dimension[0]), librevenge::RVNG_POINT);
		chart.insert("svg:height", double(m_dimension[1]), librevenge::RVNG_POINT);
	}
	if (!m_serieMap.empty())
		chart.insert("chart:class", Serie::getSerieTypeName(m_serieMap.begin()->second.m_type).c_str());
	else
		chart.insert("chart:class", Serie::getSerieTypeName(m_type).c_str());
	chart.insert("librevenge:chart-id", styleId++);
	interface->openChart(chart);

	// legend
	if (m_legend.m_show)
	{
		bool autoPlace=m_legendPosition==WPSBox2f()||m_dimension==Vec2f();
		style=librevenge::RVNGPropertyList();
		m_legend.addStyleTo(style);
		style.insert("librevenge:chart-id", styleId);
		style.insert("chart:auto-position",autoPlace);
		interface->defineChartStyle(style);
		librevenge::RVNGPropertyList legend;
		m_legend.addContentTo(legend);
		legend.insert("librevenge:chart-id", styleId++);
		legend.insert("librevenge:zone-type", "legend");
		if (!autoPlace)
		{
			legend.insert("svg:x", double(m_legendPosition[0][0])*double(m_dimension[0]), librevenge::RVNG_POINT);
			legend.insert("svg:y", double(m_legendPosition[0][1])*double(m_dimension[1]), librevenge::RVNG_POINT);
			legend.insert("svg:width", double(m_legendPosition.size()[0])*double(m_dimension[0]), librevenge::RVNG_POINT);
			legend.insert("svg:height", double(m_legendPosition.size()[1])*double(m_dimension[1]), librevenge::RVNG_POINT);
		}
		interface->openChartTextObject(legend);
		interface->closeChartTextObject();
	}
	for (auto const &textIt : m_textZoneMap)
	{
		TextZone const &zone= textIt.second;
		if (!zone.valid()) continue;
		style=librevenge::RVNGPropertyList();
		zone.addStyleTo(style);
		style.insert("librevenge:chart-id", styleId);
		interface->defineChartStyle(style);
		librevenge::RVNGPropertyList textZone;
		zone.addContentTo(textZone);
		textZone.insert("librevenge:chart-id", styleId++);
		textZone.insert("librevenge:zone-type",
		                zone.m_type==TextZone::T_Title ? "title": zone.m_type==TextZone::T_SubTitle ?"subtitle" : "footer");
		interface->openChartTextObject(textZone);
		if (zone.m_contentType==TextZone::C_Text)
		{
			std::shared_ptr<WPSSubDocument> doc(new WKSChartInternal::SubDocument(this, zone.m_type));
			listener->handleSubDocument(doc, libwps::DOC_CHART_ZONE);
		}
		interface->closeChartTextObject();
	}
	// plot area
	style=librevenge::RVNGPropertyList();
	bool autoPlace=m_plotAreaPosition==WPSBox2f()||m_dimension==Vec2f();
	m_plotAreaStyle.addTo(style);
	style.insert("librevenge:chart-id", styleId);
	style.insert("chart:include-hidden-cells","false");
	style.insert("chart:auto-position",autoPlace);
	style.insert("chart:auto-size",autoPlace);
	style.insert("chart:treat-empty-cells","leave-gap");
	style.insert("chart:right-angled-axes","true");
	style.insert("chart:stacked", m_dataStacked);
	style.insert("chart:percentage", m_dataPercentStacked);
	if (m_dataVertical)
		style.insert("chart:vertical", true);
	if (m_is3D)
	{
		style.insert("chart:three-dimensional", true);
		style.insert("chart:deep", m_is3DDeep);
	}
	interface->defineChartStyle(style);

	librevenge::RVNGPropertyList plotArea;
	if (!autoPlace)
	{
		plotArea.insert("svg:x", double(m_plotAreaPosition[0][0])*double(m_dimension[0]), librevenge::RVNG_POINT);
		plotArea.insert("svg:y", double(m_plotAreaPosition[0][1])*double(m_dimension[1]), librevenge::RVNG_POINT);
		plotArea.insert("svg:width", double(m_plotAreaPosition.size()[0])*double(m_dimension[0]), librevenge::RVNG_POINT);
		plotArea.insert("svg:height", double(m_plotAreaPosition.size()[1])*double(m_dimension[1]), librevenge::RVNG_POINT);
	}
	plotArea.insert("librevenge:chart-id", styleId++);

	librevenge::RVNGPropertyList floor, wall;
	librevenge::RVNGPropertyListVector vect;
	style=librevenge::RVNGPropertyList();
	// add floor
	m_floorStyle.addTo(style);
	style.insert("librevenge:chart-id", styleId);
	interface->defineChartStyle(style);
	floor.insert("librevenge:type", "floor");
	floor.insert("librevenge:chart-id", styleId++);
	vect.append(floor);

	// add wall
	style=librevenge::RVNGPropertyList();
	m_wallStyle.addTo(style);
	style.insert("librevenge:chart-id", styleId);
	interface->defineChartStyle(style);
	wall.insert("librevenge:type", "wall");
	wall.insert("librevenge:chart-id", styleId++);
	vect.append(wall);

	plotArea.insert("librevenge:childs", vect);

	interface->openChartPlotArea(plotArea);
	// axis : x, y, second, z
	for (int i=0; i<4; ++i)
	{
		if (m_axis[i].m_type==Axis::A_None) continue;
		style=librevenge::RVNGPropertyList();
		m_axis[i].addStyleTo(style);
		style.insert("librevenge:chart-id", styleId);
		interface->defineChartStyle(style);
		librevenge::RVNGPropertyList axis;
		m_axis[i].addContentTo(i, axis);
		axis.insert("librevenge:chart-id", styleId++);
		interface->insertChartAxis(axis);
	}
	// series
	for (auto it : m_serieMap)
	{
		auto const &serie = it.second;
		if (!serie.valid()) continue;
		style=librevenge::RVNGPropertyList();
		serie.addStyleTo(style);
		style.insert("librevenge:chart-id", styleId);
		interface->defineChartStyle(style);
		librevenge::RVNGPropertyList series;
		serie.addContentTo(series);
		series.insert("librevenge:chart-id", styleId++);
		interface->openChartSerie(series);
		interface->closeChartSerie();
	}
	interface->closeChartPlotArea();

	interface->closeChart();
}

////////////////////////////////////////////////////////////
// Position
////////////////////////////////////////////////////////////
librevenge::RVNGString WKSChart::Position::getCellName() const
{
	if (!valid())
	{
		WPS_DEBUG_MSG(("WKSChart::Position::getCellName: called on invalid cell\n"));
		return librevenge::RVNGString();
	}
	std::string cellName=libwps::getCellName(m_pos);
	if (cellName.empty())
		return librevenge::RVNGString();
	std::stringstream o;
	o << m_sheetName.cstr() << "." << cellName;
	return librevenge::RVNGString(o.str().c_str());
}
std::ostream &operator<<(std::ostream &o, WKSChart::Position const &pos)
{
	if (pos.valid())
		o << pos.m_pos << "[" << pos.m_sheetName.cstr() << "]";
	else
		o << "_";
	return o;
}
////////////////////////////////////////////////////////////
// Axis
////////////////////////////////////////////////////////////
WKSChart::Axis::Axis()
	: m_type(WKSChart::Axis::A_None)
	, m_automaticScaling(true)
	, m_scaling()
	, m_showGrid(true)
	, m_showLabel(true)
	, m_showTitle(true)
	, m_titleRange()
	, m_title()
	, m_subTitle()
	, m_style()
{
	m_style.m_lineWidth=0;
}

WKSChart::Axis::~Axis()
{
}

void WKSChart::Axis::addContentTo(int coord, librevenge::RVNGPropertyList &propList) const
{
	std::string axis("");
	axis += coord==0 ? 'x' : coord==3 ? 'z' : 'y';
	propList.insert("chart:dimension",axis.c_str());
	axis = (coord==2 ? "secondary-y" : "primary-"+axis);
	propList.insert("chart:name",axis.c_str());
	librevenge::RVNGPropertyListVector childs;
	if (m_showGrid && (m_type==A_Numeric || m_type==A_Logarithmic))
	{
		librevenge::RVNGPropertyList grid;
		grid.insert("librevenge:type", "grid");
		grid.insert("chart:class", "major");
		childs.append(grid);
	}
	if (m_labelRanges[0].valid(m_labelRanges[1]) && m_showLabel)
	{
		librevenge::RVNGPropertyList range;
		range.insert("librevenge:sheet-name", m_labelRanges[0].m_sheetName);
		range.insert("librevenge:start-row", m_labelRanges[0].m_pos[1]);
		range.insert("librevenge:start-column", m_labelRanges[0].m_pos[0]);
		if (m_labelRanges[0].m_sheetName!=m_labelRanges[1].m_sheetName)
			range.insert("librevenge:end-sheet-name", m_labelRanges[1].m_sheetName);
		range.insert("librevenge:end-row", m_labelRanges[1].m_pos[1]);
		range.insert("librevenge:end-column", m_labelRanges[1].m_pos[0]);
		librevenge::RVNGPropertyListVector vect;
		vect.append(range);
		librevenge::RVNGPropertyList categories;
		categories.insert("librevenge:type", "categories");
		categories.insert("table:cell-range-address", vect);
		childs.append(categories);
	}
	if (m_showTitle && (!m_title.empty() || !m_subTitle.empty()))
	{
		auto finalString(m_title);
		if (!m_title.empty() && !m_subTitle.empty()) finalString.append(" - ");
		finalString.append(m_subTitle);
		librevenge::RVNGPropertyList title;
		title.insert("librevenge:type", "title");
		title.insert("librevenge:text", finalString);
		childs.append(title);
	}
	else if (m_showTitle && m_titleRange.valid())
	{
		librevenge::RVNGPropertyList title;
		title.insert("librevenge:type", "title");
		librevenge::RVNGPropertyList range;
		range.insert("librevenge:sheet-name", m_titleRange.m_sheetName);
		range.insert("librevenge:start-row", m_titleRange.m_pos[1]);
		range.insert("librevenge:start-column", m_titleRange.m_pos[0]);
		librevenge::RVNGPropertyListVector vect;
		vect.append(range);
		title.insert("table:cell-range", vect);
		childs.append(title);
	}
	if (!childs.empty())
		propList.insert("librevenge:childs", childs);
}

void WKSChart::Axis::addStyleTo(librevenge::RVNGPropertyList &propList) const
{
	propList.insert("chart:display-label", m_showLabel);
	propList.insert("chart:axis-position", 0, librevenge::RVNG_GENERIC);
	propList.insert("chart:reverse-direction", false);
	propList.insert("chart:logarithmic", m_type==WKSChart::Axis::A_Logarithmic);
	propList.insert("text:line-break", false);
	if (!m_automaticScaling)
	{
		propList.insert("chart:minimum", double(m_scaling[0]), librevenge::RVNG_GENERIC);
		propList.insert("chart:maximum", double(m_scaling[1]), librevenge::RVNG_GENERIC);
	}
	m_style.addTo(propList, true);
}

std::ostream &operator<<(std::ostream &o, WKSChart::Axis const &axis)
{
	switch (axis.m_type)
	{
	case WKSChart::Axis::A_None:
		o << "none,";
		break;
	case WKSChart::Axis::A_Numeric:
		o << "numeric,";
		break;
	case WKSChart::Axis::A_Logarithmic:
		o << "logarithmic,";
		break;
	case WKSChart::Axis::A_Sequence:
		o << "sequence,";
		break;
	case WKSChart::Axis::A_Sequence_Skip_Empty:
		o << "sequence[noEmpty],";
		break;
#if !defined(__clang__)
	default:
		o << "###type,";
		WPS_DEBUG_MSG(("WKSChart::Axis: unexpected type\n"));
		break;
#endif
	}
	if (axis.m_showGrid) o << "show[grid],";
	if (axis.m_showLabel) o << "show[label],";
	if (axis.m_labelRanges[0].valid(axis.m_labelRanges[1]))
		o << "label[range]=" << axis.m_labelRanges[0] << ":" << axis.m_labelRanges[1] << ",";
	if (axis.m_showTitle)
	{
		if (axis.m_titleRange.valid()) o << "title[range]=" << axis.m_titleRange << ",";
		if (!axis.m_title.empty()) o << "title=" << axis.m_title.cstr() << ",";
		if (!axis.m_subTitle.empty()) o << "subTitle=" << axis.m_subTitle.cstr() << ",";
	}
	if (!axis.m_automaticScaling && axis.m_scaling!=Vec2f())
		o << "scaling=manual[" << axis.m_scaling[0] << "->" << axis.m_scaling[1] << ",";
	o << axis.m_style;
	return o;
}

////////////////////////////////////////////////////////////
// Legend
////////////////////////////////////////////////////////////
void WKSChart::Legend::addContentTo(librevenge::RVNGPropertyList &propList) const
{
	if (m_position[0]>0 && m_position[1]>0)
	{
		propList.insert("svg:x", double(m_position[0]), librevenge::RVNG_POINT);
		propList.insert("svg:y", double(m_position[1]), librevenge::RVNG_POINT);
	}
	if (!m_autoPosition || !m_relativePosition)
		return;
	std::stringstream s;
	if (m_relativePosition&WPSBorder::TopBit)
		s << "top";
	else if (m_relativePosition&WPSBorder::BottomBit)
		s << "bottom";
	if (s.str().length() && (m_relativePosition&(WPSBorder::LeftBit|WPSBorder::RightBit)))
		s << "-";
	if (m_relativePosition&WPSBorder::LeftBit)
		s << "start";
	else if (m_relativePosition&WPSBorder::RightBit)
		s << "end";
	propList.insert("chart:legend-position", s.str().c_str());
}

void WKSChart::Legend::addStyleTo(librevenge::RVNGPropertyList &propList) const
{
	propList.insert("chart:auto-position", m_autoPosition);
	m_font.addTo(propList);
	m_style.addTo(propList);
}

std::ostream &operator<<(std::ostream &o, WKSChart::Legend const &legend)
{
	if (legend.m_show)
		o << "show,";
	if (legend.m_autoPosition)
	{
		o << "automaticPos[";
		if (legend.m_relativePosition&WPSBorder::TopBit)
			o << "t";
		else if (legend.m_relativePosition&WPSBorder::RightBit)
			o << "b";
		else
			o << "c";
		if (legend.m_relativePosition&WPSBorder::LeftBit)
			o << "L";
		else if (legend.m_relativePosition&WPSBorder::BottomBit)
			o << "R";
		else
			o << "C";
		o << "]";
	}
	else
		o << "pos=" << legend.m_position << ",";
	o << legend.m_style;
	return o;
}

////////////////////////////////////////////////////////////
// Serie
////////////////////////////////////////////////////////////
WKSChart::Serie::Serie()
	: m_type(WKSChart::Serie::S_Bar)
	, m_useSecondaryY(false)
	, m_font()
	, m_legendRange()
	, m_legendText()
	, m_style()
	, m_pointType(P_None)
{
	m_style.m_lineWidth=0;
	m_style.setSurfaceColor(WPSColor(0x80,0x80,0xFF));
}

WKSChart::Serie::~Serie()
{
}

std::string WKSChart::Serie::getSerieTypeName(Type type)
{
	switch (type)
	{
	case S_Area:
		return "chart:area";
	case S_Bar:
		return "chart:bar";
	case S_Bubble:
		return "chart:bubble";
	case S_Circle:
		return "chart:circle";
	case S_Column:
		return "chart:column";
	case S_Gantt:
		return "chart:gantt";
	case S_Line:
		return "chart:line";
	case S_Radar:
		return "chart:radar";
	case S_Ring:
		return "chart:ring";
	case S_Scatter:
		return "chart:scatter";
	case S_Stock:
		return "chart:stock";
	case S_Surface:
		return "chart:surface";
#if !defined(__clang__)
	default:
		break;
#endif
	}
	return "chart:bar";
}

void WKSChart::Serie::addContentTo(librevenge::RVNGPropertyList &serie) const
{
	serie.insert("chart:class",getSerieTypeName(m_type).c_str());
	if (m_useSecondaryY)
		serie.insert("chart:attached-axis","secondary-y");
	librevenge::RVNGPropertyList datapoint;
	librevenge::RVNGPropertyListVector vect;
	if (m_ranges[0].valid(m_ranges[1]))
	{
		librevenge::RVNGPropertyList range;
		range.insert("librevenge:sheet-name", m_ranges[0].m_sheetName);
		range.insert("librevenge:start-row", m_ranges[0].m_pos[1]);
		range.insert("librevenge:start-column", m_ranges[0].m_pos[0]);
		if (m_ranges[0].m_sheetName != m_ranges[1].m_sheetName)
			range.insert("librevenge:end-sheet-name", m_ranges[1].m_sheetName);
		range.insert("librevenge:end-row", m_ranges[1].m_pos[1]);
		range.insert("librevenge:end-column", m_ranges[1].m_pos[0]);
		vect.append(range);
		serie.insert("chart:values-cell-range-address", vect);
		vect.clear();
	}

	// to do use labelRanges

	// USEME: set the font here
	if (m_legendRange.valid())
	{
		librevenge::RVNGPropertyList label;
		label.insert("librevenge:sheet-name", m_legendRange.m_sheetName);
		label.insert("librevenge:start-row", m_legendRange.m_pos[1]);
		label.insert("librevenge:start-column", m_legendRange.m_pos[0]);
		vect.append(label);
		serie.insert("chart:label-cell-address", vect);
		vect.clear();
	}
	if (!m_legendText.empty())
	{
		// replace ' ' and non basic caracters by _ because this causes LibreOffice's problem
		std::string basicString(m_legendText.cstr());
		for (auto &c : basicString)
		{
			if (c==' ' || static_cast<unsigned char>(c)>=0x80)
				c='_';
		}
		serie.insert("chart:label-string", basicString.c_str());
	}
	datapoint.insert("librevenge:type", "data-point");
	Vec2i dataSize=m_ranges[1].m_pos-m_ranges[0].m_pos;
	datapoint.insert("chart:repeated", 1+std::max(dataSize[0],dataSize[1]));
	vect.append(datapoint);
	serie.insert("librevenge:childs", vect);
}

void WKSChart::Serie::addStyleTo(librevenge::RVNGPropertyList &propList) const
{
	m_style.addTo(propList);
	if (m_pointType != WKSChart::Serie::P_None)
	{
		char const *what[] =
		{
			"none", "automatic", "square", "diamond", "arrow-down",
			"arrow-up", "arrow-right", "arrow-left", "bow-tie", "hourglass",
			"circle", "star", "x", "plus", "asterisk",
			"horizontal-bar", "vertical-bar"
		};
		if (m_pointType == WKSChart::Serie::P_Automatic)
			propList.insert("chart:symbol-type", "automatic");
		else if (m_pointType < WPS_N_ELEMENTS(what))
		{
			propList.insert("chart:symbol-type", "named-symbol");
			propList.insert("chart:symbol-name", what[m_pointType]);
		}
	}
	//propList.insert("chart:data-label-number","value");
	//propList.insert("chart:data-label-text","false");
}

void WKSChart::Serie::setPrimaryPattern(WPSGraphicStyle::Pattern const &pattern, bool force1D)
{
	WPSColor finalColor;
	if (pattern.getUniqueColor(finalColor))
		setPrimaryColor(finalColor, 1, force1D);
	else if (!force1D && !is1DStyle())
		m_style.setPattern(pattern);
	else if (pattern.getAverageColor(finalColor))
		setPrimaryColor(finalColor);
}

std::ostream &operator<<(std::ostream &o, WKSChart::Serie const &serie)
{
	switch (serie.m_type)
	{
	case WKSChart::Serie::S_Area:
		o << "area,";
		break;
	case WKSChart::Serie::S_Bar:
		o << "bar,";
		break;
	case WKSChart::Serie::S_Bubble:
		o << "bubble,";
		break;
	case WKSChart::Serie::S_Circle:
		o << "circle,";
		break;
	case WKSChart::Serie::S_Column:
		o << "column,";
		break;
	case WKSChart::Serie::S_Gantt:
		o << "gantt,";
		break;
	case WKSChart::Serie::S_Line:
		o << "line,";
		break;
	case WKSChart::Serie::S_Radar:
		o << "radar,";
		break;
	case WKSChart::Serie::S_Ring:
		o << "ring,";
		break;
	case WKSChart::Serie::S_Scatter:
		o << "scatter,";
		break;
	case WKSChart::Serie::S_Stock:
		o << "stock,";
		break;
	case WKSChart::Serie::S_Surface:
		o << "surface,";
		break;
#if !defined(__clang__)
	default:
		o << "###type,";
		WPS_DEBUG_MSG(("WKSChart::Serie: unexpected type\n"));
		break;
#endif
	}
	o << "range=" << serie.m_ranges[0] << ":" << serie.m_ranges[1] << ",";
	o << serie.m_style;
	if (serie.m_labelRanges[0].valid(serie.m_labelRanges[1]))
		o << "label[range]=" << serie.m_labelRanges[0] << "<->" << serie.m_labelRanges[1] << ",";
	if (serie.m_legendRange.valid())
		o << "legend[range]=" << serie.m_legendRange << ",";
	if (!serie.m_legendText.empty())
		o << "label[text]=" << serie.m_legendText.cstr() << ",";
	if (serie.m_pointType != WKSChart::Serie::P_None)
	{
		char const *what[] =
		{
			"none", "automatic", "square", "diamond", "arrow-down",
			"arrow-up", "arrow-right", "arrow-left", "bow-tie", "hourglass",
			"circle", "star", "x", "plus", "asterisk",
			"horizontal-bar", "vertical-bar"
		};
		if (serie.m_pointType>0 && serie.m_pointType < WPS_N_ELEMENTS(what))
		{
			o << "point=" << what[serie.m_pointType] << ",";
		}
		else if (serie.m_pointType>0)
			o << "#point=" << serie.m_pointType << ",";
	}
	return o;
}

////////////////////////////////////////////////////////////
// TextZone
////////////////////////////////////////////////////////////
WKSChart::TextZone::TextZone(Type type)
	: m_type(type)
	, m_contentType(WKSChart::TextZone::C_Text)
	, m_show(true)
	, m_position(-1,-1)
	, m_cell()
	, m_textEntryList()
	, m_font()
	, m_style()
{
	m_style.m_lineWidth=0;
}

WKSChart::TextZone::~TextZone()
{
}

void WKSChart::TextZone::addContentTo(librevenge::RVNGPropertyList &propList) const
{
	if (m_position[0]>0 && m_position[1]>0)
	{
		propList.insert("svg:x", double(m_position[0]), librevenge::RVNG_POINT);
		propList.insert("svg:y", double(m_position[1]), librevenge::RVNG_POINT);
	}
	else
		propList.insert("chart:auto-position",true);
	propList.insert("chart:auto-size",true);
	switch (m_type)
	{
	case T_Footer:
		propList.insert("librevenge:zone-type", "footer");
		break;
	case T_Title:
		propList.insert("librevenge:zone-type", "title");
		break;
	case T_SubTitle:
		propList.insert("librevenge:zone-type", "subtitle");
		break;
#if !defined(__clang__)
	default:
		propList.insert("librevenge:zone-type", "label");
		WPS_DEBUG_MSG(("WKSChart::TextZone:addContentTo: unexpected type\n"));
		break;
#endif
	}
	if (m_contentType==C_Cell && m_cell.valid())
	{
		librevenge::RVNGPropertyList range;
		librevenge::RVNGPropertyListVector vect;
		range.insert("librevenge:sheet-name", m_cell.m_sheetName);
		range.insert("librevenge:row", m_cell.m_pos[1]);
		range.insert("librevenge:column", m_cell.m_pos[0]);
		vect.append(range);
		propList.insert("table:cell-range", vect);
	}
}

void WKSChart::TextZone::addStyleTo(librevenge::RVNGPropertyList &propList) const
{
	m_font.addTo(propList);
	m_style.addTo(propList);
}

std::ostream &operator<<(std::ostream &o, WKSChart::TextZone const &zone)
{
	switch (zone.m_type)
	{
	case WKSChart::TextZone::T_SubTitle:
		o << "sub";
		WPS_FALLTHROUGH;
	case WKSChart::TextZone::T_Title:
		o << "title,";
		break;
	case WKSChart::TextZone::T_Footer:
		o << "footer,";
		break;
#if !defined(__clang__)
	default:
		o << "###type,";
		WPS_DEBUG_MSG(("WKSChart::TextZone: unexpected type\n"));
		break;
#endif
	}
	if (zone.m_contentType==WKSChart::TextZone::C_Text)
		o << "text,";
	else if (zone.m_contentType==WKSChart::TextZone::C_Cell)
		o << "cell=" << zone.m_cell << ",";
	if (zone.m_position[0]>0 || zone.m_position[1]>0)
		o << "pos=" << zone.m_position << ",";
	o << zone.m_style;
	return o;
}

/* vim:set shiftwidth=4 softtabstop=4 noexpandtab: */
