/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 *
 * Copyright 2004 Komarov Valery
 * Copyright 2006 Christophe Leitienne
 * Copyright 2008-2017 David Hoerl
 * Copyright 2013 Bob Colbert
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
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS ''AS
 * IS'' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO,
 * THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
 * PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDERS OR
 * CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
 * EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
 * PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS;
 * OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
 * WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR
 * OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 *
 */

#include "libxls/xlsstruct.h"

void verbose(char* str);

char *codepage_decode(const char *s, size_t len, xlsWorkBook *pWB);
char *unicode_decode(const char *s, size_t len, xlsWorkBook *pWB);
char *transcode_utf16_to_utf8(const char *s, size_t len);
char *get_string(const char *s, size_t len, BYTE is2, xlsWorkBook *pWB);
char *strcat_realloc(char *s, char *add, size_t *buflenptr);
DWORD xls_getColor(const WORD color,WORD def);

void xls_showBookInfo(xlsWorkBook* pWB);
void xls_showROW(struct st_row_data* row);
void xls_showColinfo(struct st_colinfo_data* col);
void xls_showCell(struct st_cell_data* cell);
void xls_showFont(struct st_font_data* font);
void xls_showXF(XF8* xf);
void xls_showFormat(struct st_format_data* format);
char* xls_getfcell(xlsWorkBook* pWB, struct st_cell_data* cell, BYTE *label);
char* xls_getCSS(xlsWorkBook* pWB);
void xls_showBOF(BOF* bof);
