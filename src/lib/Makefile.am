if OS_WIN32
version_info = -version-info $(LT_CURRENT):$(LT_REVISION):$(LT_AGE) -avoid-version
else
version_info = -version-info $(LT_CURRENT):$(LT_REVISION):$(LT_AGE)
endif

lib_LTLIBRARIES = libwps-@WPS_MAJOR_VERSION@.@WPS_MINOR_VERSION@.la $(target_libwps_stream)

AM_CXXFLAGS =  -I$(top_srcdir)/inc $(REVENGE_CFLAGS) $(DEBUG_CXXFLAGS) -DBUILD_WPS=1

libwps_@WPS_MAJOR_VERSION@_@WPS_MINOR_VERSION@_la_LIBADD  = $(REVENGE_LIBS) @LIBWPS_WIN32_RESOURCE@
libwps_@WPS_MAJOR_VERSION@_@WPS_MINOR_VERSION@_la_DEPENDENCIES = @LIBWPS_WIN32_RESOURCE@
libwps_@WPS_MAJOR_VERSION@_@WPS_MINOR_VERSION@_la_LDFLAGS = $(version_info) -export-dynamic -no-undefined
libwps_@WPS_MAJOR_VERSION@_@WPS_MINOR_VERSION@_la_SOURCES = \
	libwps_internal.cpp		\
	libwps_tools_win.h		\
	libwps_tools_win.cpp		\
	libwps_internal.h		\
	DosWord.cpp			\
	DosWord.h			\
	Lotus.cpp			\
	Lotus.h				\
	LotusChart.cpp			\
	LotusChart.h			\
	LotusGraph.cpp			\
	LotusGraph.h			\
	LotusSpreadsheet.cpp		\
	LotusSpreadsheet.h		\
	LotusStyleManager.cpp		\
	LotusStyleManager.h		\
	MSWrite.cpp			\
	MSWrite.h			\
	Multiplan.cpp			\
	Multiplan.h			\
	PocketWord.cpp			\
	PocketWord.h			\
	Quattro.cpp			\
	Quattro.h			\
	QuattroFormula.cpp		\
	QuattroFormula.h		\
	QuattroGraph.cpp		\
	QuattroGraph.h			\
	QuattroSpreadsheet.cpp		\
	QuattroSpreadsheet.h		\
	QuattroDos.cpp			\
	QuattroDos.h			\
	QuattroDosChart.cpp		\
	QuattroDosChart.h		\
	QuattroDosSpreadsheet.cpp	\
	QuattroDosSpreadsheet.h		\
	Quattro9.cpp			\
	Quattro9.h			\
	Quattro9Graph.cpp		\
	Quattro9Graph.h			\
	Quattro9Spreadsheet.cpp		\
	Quattro9Spreadsheet.h		\
	WKS4.cpp			\
	WKS4.h				\
	WKS4Chart.cpp			\
	WKS4Chart.h			\
	WKS4Format.cpp			\
	WKS4Format.h			\
	WKS4Spreadsheet.cpp		\
	WKS4Spreadsheet.h		\
	WKSChart.cpp			\
	WKSChart.h			\
	WKSContentListener.cpp		\
	WKSContentListener.h		\
	WKSParser.cpp			\
	WKSParser.h			\
	WKSSubDocument.cpp		\
	WKSSubDocument.h		\
	WPS4.cpp			\
	WPS4.h				\
	WPS4Graph.cpp			\
	WPS4Graph.h			\
	WPS4Text.cpp			\
	WPS4Text.h			\
	WPS8.cpp			\
	WPS8.h				\
	WPS8Graph.cpp			\
	WPS8Graph.h			\
	WPS8Struct.cpp			\
	WPS8Struct.h			\
	WPS8Table.cpp			\
	WPS8Table.h			\
	WPS8Text.cpp			\
	WPS8Text.h			\
	WPS8TextStyle.cpp		\
	WPS8TextStyle.h			\
	WPSCell.cpp			\
	WPSCell.h			\
	WPSContentListener.cpp		\
	WPSContentListener.h		\
	WPSDebug.cpp			\
	WPSDebug.h			\
	WPSDocument.cpp			\
	WPSEntry.cpp			\
	WPSEntry.h			\
	WPSFont.cpp			\
	WPSFont.h			\
	WPSGraphicShape.cpp		\
	WPSGraphicShape.h		\
	WPSGraphicStyle.cpp		\
	WPSGraphicStyle.h		\
	WPSHeader.cpp			\
	WPSHeader.h			\
	WPSList.cpp			\
	WPSList.h			\
	WPSListener.cpp			\
	WPSListener.h			\
	WPSOLE1Parser.cpp		\
	WPSOLE1Parser.h			\
	WPSOLEObject.cpp		\
	WPSOLEObject.h			\
	WPSOLEParser.cpp		\
	WPSOLEParser.h			\
	WPSOLEStream.cpp		\
	WPSOLEStream.h			\
	WPSPageSpan.cpp			\
	WPSPageSpan.h			\
	WPSParagraph.cpp		\
	WPSParagraph.h			\
	WPSParser.cpp			\
	WPSParser.h			\
	WPSPosition.h			\
	WPSStream.cpp			\
	WPSStream.h			\
	WPSStringStream.cpp		\
	WPSStringStream.h		\
	WPSSubDocument.cpp		\
	WPSSubDocument.h		\
	WPSTable.cpp			\
	WPSTable.h			\
	WPSTextParser.h			\
	WPSTextParser.cpp		\
	WPSTextSubDocument.cpp		\
	WPSTextSubDocument.h		\
	XYWrite.cpp			\
	XYWrite.h

if OS_WIN32

@LIBWPS_WIN32_RESOURCE@ : libwps.rc $(libwps_@WPS_MAJOR_VERSION@_@WPS_MINOR_VERSION@_la_OBJECTS)
	chmod +x $(top_srcdir)/build/win32/*compile-resource
	WINDRES=@WINDRES@ $(top_srcdir)/build/win32/lt-compile-resource libwps.rc @LIBWPS_WIN32_RESOURCE@

endif

EXTRA_DIST = \
	libwps.rc.in

# These may be in the builddir too
BUILD_EXTRA_DIST = \
	libwps.rc
