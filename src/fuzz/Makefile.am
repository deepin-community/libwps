noinst_PROGRAMS = \
	wpsfuzzer wksfuzzer wdbfuzzer wrifuzzer docfuzzer 123fuzzer \
	wqfuzzer mpfuzzer

AM_CXXFLAGS = -I$(top_srcdir)/inc \
	$(REVENGE_GENERATORS_CFLAGS) \
	$(REVENGE_CFLAGS) \
	$(REVENGE_STREAM_CFLAGS) \
	$(DEBUG_CXXFLAGS)

commonfuzzer_ldadd = \
	$(top_builddir)/src/lib/libwps-@WPS_MAJOR_VERSION@.@WPS_MINOR_VERSION@.la \
	$(REVENGE_GENERATORS_LIBS) \
	$(REVENGE_LIBS) \
	$(REVENGE_STREAM_LIBS) \
	-lFuzzingEngine

wpsfuzzer_LDADD = $(commonfuzzer_ldadd)

wpsfuzzer_SOURCES = \
	wpsfuzzer.cpp

wksfuzzer_LDADD = $(commonfuzzer_ldadd)

wksfuzzer_SOURCES = \
	wksfuzzer.cpp

wdbfuzzer_LDADD = $(commonfuzzer_ldadd)

wdbfuzzer_SOURCES = \
	wdbfuzzer.cpp

wrifuzzer_LDADD = $(commonfuzzer_ldadd)

wrifuzzer_SOURCES = \
	wrifuzzer.cpp

docfuzzer_LDADD = $(commonfuzzer_ldadd)

docfuzzer_SOURCES = \
	docfuzzer.cpp

123fuzzer_LDADD = $(commonfuzzer_ldadd)

123fuzzer_SOURCES = \
	123fuzzer.cpp

wqfuzzer_LDADD = $(commonfuzzer_ldadd)

wqfuzzer_SOURCES = \
	wqfuzzer.cpp

mpfuzzer_LDADD = $(commonfuzzer_ldadd)

mpfuzzer_SOURCES = \
	mpfuzzer.cpp
