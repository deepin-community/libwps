if BUILD_TOOLS

bin_PROGRAMS = wps2html

AM_CXXFLAGS = -I$(top_srcdir)/inc -I$(top_srcdir)/src/conv/helper $(REVENGE_CFLAGS) $(REVENGE_GENERATORS_CFLAGS) $(REVENGE_STREAM_CFLAGS) $(DEBUG_CXXFLAGS)

wps2html_DEPENDENCIES = @WPS2HTML_WIN32_RESOURCE@

if STATIC_TOOLS

wps2html_LDADD = \
	../../lib/@WPS_OBJDIR@/libwps-@WPS_MAJOR_VERSION@.@WPS_MINOR_VERSION@.a \
	../helper/@WPS_OBJDIR@/libconvHelper.a $(REVENGE_LIBS) $(REVENGE_GENERATORS_LIBS) $(REVENGE_STREAM_LIBS) @WPS2HTML_WIN32_RESOURCE@
wps2html_LDFLAGS = -all-static

else	

wps2html_LDADD = \
	../../lib/libwps-@WPS_MAJOR_VERSION@.@WPS_MINOR_VERSION@.la \
	../helper/libconvHelper.la $(REVENGE_LIBS) $(REVENGE_GENERATORS_LIBS) $(REVENGE_STREAM_LIBS) @WPS2HTML_WIN32_RESOURCE@

endif

wps2html_SOURCES = 			\
	wps2html.cpp

if OS_WIN32

@WPS2HTML_WIN32_RESOURCE@ : wps2html.rc $(wps2html_OBJECTS)
	chmod +x $(top_srcdir)/build/win32/*compile-resource
	WINDRES=@WINDRES@ $(top_srcdir)/build/win32/lt-compile-resource wps2html.rc @WPS2HTML_WIN32_RESOURCE@
endif

EXTRA_DIST = \
	wps2html.rc.in

# These may be in the builddir too
BUILD_EXTRA_DIST = \
	wps2html.rc	 

endif
