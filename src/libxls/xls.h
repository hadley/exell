/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 *
 * Copyright 2004 Komarov Valery
 * Copyright 2006 Christophe Leitienne
 * Copyright 2008-2017 David Hoerl
 * Copyright 2013-2018 Evan Miller
 *
 * This file is part of libxls -- A multiplatform, C/C++ library for parsing
 * Excel(TM) files.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 *    1. Redistributions of source code must retain the above copyright notice,
 *    this list of conditions and the following disclaimer.
 *
 *    2. Redistributions in binary form must reproduce the above copyright
 *    notice, this list of conditions and the following disclaimer in the
 *    documentation and/or other materials provided with the distribution.
 *
 * THIS SOFTWARE IS PROVIDED BY Evan Miller ''AS IS'' AND ANY EXPRESS OR
 * IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF
 * MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO
 * EVENT SHALL THE COPYRIGHT HOLDERS OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT,
 * INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
 * THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 *
 */

#ifndef XLS_INCLUDE
#define XLS_INCLUDE
 
#ifdef __cplusplus
namespace xls {
extern "C" {
#endif

#include "libxls/xlstypes.h"
#include "libxls/xlsstruct.h"
#include "libxls/xlstool.h"

typedef enum {
    LIBXLS_OK,
    LIBXLS_ERROR_OPEN,
    LIBXLS_ERROR_SEEK,
    LIBXLS_ERROR_READ,
    LIBXLS_ERROR_PARSE,
    LIBXLS_ERROR_MALLOC,
    LIBXLS_ERROR_UNSUPPORTED_ENCRYPTION,
    LIBXLS_ERROR_NULL_ARGUMENT
} xls_error_t;

const char* xls_getVersion(void);
const char* xls_getError(xls_error_t code);

int xls(int debug);	// Set debug. Force library to load?
void xls_set_formula_hander(xls_formula_handler handler);

xls_error_t xls_parseWorkBook(xlsWorkBook* pWB);
xls_error_t xls_parseWorkSheet(xlsWorkSheet* pWS);

// Preferred API
// charset - convert 16bit strings within the spread sheet to this 8-bit encoding (UTF-8 default)
xlsWorkBook *xls_open_file(const char *file, const char *charset, xls_error_t *outError);
xlsWorkBook *xls_open_buffer(const unsigned char *data, size_t data_len,
        const char *charset, xls_error_t *outError);
void xls_close_WB(xlsWorkBook* pWB);

// Historical API
xlsWorkBook* xls_open(const char *file,const char *charset);
#define xls_close xls_close_WB

xlsWorkSheet * xls_getWorkSheet(xlsWorkBook* pWB,int num);
void xls_close_WS(xlsWorkSheet* pWS);

xlsSummaryInfo *xls_summaryInfo(xlsWorkBook* pWB);
void xls_close_summaryInfo(xlsSummaryInfo *pSI);

// utility function
xlsRow *xls_row(xlsWorkSheet* pWS, WORD cellRow);
xlsCell	*xls_cell(xlsWorkSheet* pWS, WORD cellRow, WORD cellCol);

// parse defined name formula
char *xls_parseFormulaBytes(xlsWorkBook* pWB, BYTE *bytes, size_t len);


#ifdef __cplusplus
} // extern c block
} // namespace
#endif

#endif

