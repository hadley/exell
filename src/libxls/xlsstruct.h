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

#ifndef XLS_STRUCT_INC
#define XLS_STRUCT_INC

#include "libxls/ole.h"

#define XLS_RECORD_EOF          0x000A
#define XLS_RECORD_EXTERNSHEET  0x0017
#define XLS_RECORD_DEFINEDNAME  0x0018
#define XLS_RECORD_NOTE         0x001C
#define XLS_RECORD_1904         0x0022
#define XLS_RECORD_FILEPASS     0x002F
#define XLS_RECORD_CONTINUE     0x003C
#define XLS_RECORD_WINDOW1      0x003D
#define XLS_RECORD_CODEPAGE     0x0042
#define XLS_RECORD_OBJ          0x005D
#define XLS_RECORD_MERGEDCELLS  0x00E5
#define XLS_RECORD_DEFCOLWIDTH  0x0055
#define XLS_RECORD_COLINFO      0x007D
#define XLS_RECORD_BOUNDSHEET   0x0085
#define XLS_RECORD_PALETTE      0x0092
#define XLS_RECORD_MULRK        0x00BD
#define XLS_RECORD_MULBLANK     0x00BE
#define XLS_RECORD_RSTRING      0x00D6
#define XLS_RECORD_DBCELL       0x00D7
#define XLS_RECORD_XF           0x00E0
#define XLS_RECORD_MSODRAWINGGROUP   0x00EB
#define XLS_RECORD_MSODRAWING   0x00EC
#define XLS_RECORD_SST          0x00FC
#define XLS_RECORD_LABELSST     0x00FD
#define XLS_RECORD_EXTSST       0x00FF
#define XLS_RECORD_SUPBOOK      0x01AE
#define XLS_RECORD_TXO          0x01B6
#define XLS_RECORD_HYPERREF     0x01B8
#define XLS_RECORD_BLANK        0x0201
#define XLS_RECORD_NUMBER       0x0203
#define XLS_RECORD_LABEL        0x0204
#define XLS_RECORD_BOOLERR      0x0205
#define XLS_RECORD_STRING       0x0207 // only follows a formula
#define XLS_RECORD_ROW          0x0208
#define XLS_RECORD_INDEX        0x020B
#define XLS_RECORD_ARRAY        0x0221 // Array-entered formula
#define XLS_RECORD_DEFAULTROWHEIGHT    0x0225
#define XLS_RECORD_FONT         0x0031 // spec says 0x0231 but Excel expects 0x0031
#define XLS_RECORD_FONT_ALT     0x0231
#define XLS_RECORD_WINDOW2      0x023E
#define XLS_RECORD_RK           0x027E
#define XLS_RECORD_STYLE        0x0293
#define XLS_RECORD_FORMULA      0x0006
#define XLS_RECORD_FORMULA_ALT  0x0406 // Apple Numbers bug
#define XLS_RECORD_FORMAT       0x041E
#define XLS_RECORD_BOF          0x0809

#define BLANK_CELL  XLS_RECORD_BLANK  // compat

#if defined(_AIX) || defined(__sun)
#pragma pack(1)
#else
#pragma pack(push, 1)
#endif

typedef struct BOF
{
    WORD id;
    WORD size;
}
BOF;

typedef struct BIFF
{
    WORD ver;
    WORD type;
    WORD id_make;
    WORD year;
    DWORD flags;
    DWORD min_ver;
}
BIFF;

typedef struct WIND1
{
    WORD xWn;
    WORD yWn;
    WORD dxWn;
    WORD dyWn;
    WORD grbit;
    WORD itabCur;
    WORD itabFirst;
    WORD ctabSel;
    WORD wTabRatio;
}
WIND1;

typedef struct BOUNDSHEET
{
    DWORD	filepos;
    BYTE	type;
    BYTE	visible;
    char	name[1];
}
BOUNDSHEET;

typedef struct ROW
{
    WORD	index;
    WORD	fcell; // first cell, 0-indexed
    WORD	lcell; // last cell, 1-indexed
    WORD	height;
    WORD	notused;
    WORD	notused2; //used only for BIFF3-4
    WORD	flags;
    WORD	xf;
}
ROW;

typedef struct COL
{
    WORD	row;
    WORD	col;
    WORD	xf;
}
COL;


typedef struct FORMULA // BIFF8
{
    WORD	row;
    WORD	col;
    WORD	xf;
	// next 8 bytes either a IEEE double, or encoded on a byte basis
    BYTE	resid;
    BYTE	resdata[5];
    WORD	res;
    WORD	flags;
    BYTE	chn[4]; // BIFF8
    WORD	len;
    BYTE	value[1]; //var
}
FORMULA;

typedef struct FARRAY // BIFF8
{
    WORD	row1;
    WORD	row2;
    BYTE	col1;
    BYTE	col2;
    WORD	flags;
    BYTE	chn[4]; // BIFF8
    WORD	len;
    BYTE	value[1]; //var
}
FARRAY;

typedef struct RK
{
    WORD	row;
    WORD	col;
    WORD	xf;
    DWORD   value;
}
RK;

typedef struct MULRK
{
    WORD	row;
    WORD	col;
	struct {
		WORD	xf;
		DWORD   value;
	}		rk[1];
	//WORD	last_col;
}
MULRK;

typedef struct MULBLANK
{
    WORD	row;
    WORD	col;
    WORD	xf[1];
	//WORD	last_col;
}
MULBLANK;

typedef struct BLANK
{
    WORD	row;
    WORD	col;
    WORD	xf;
}
BLANK;

typedef struct LABEL
{
    WORD	row;
    WORD	col;
    WORD	xf;
    BYTE	value[1]; // var
}
LABEL;

typedef struct BOOLERR
{
    WORD    row;
    WORD    col;
    WORD    xf;
    BYTE    value;
    BYTE    iserror;
}
BOOLERR;

typedef struct SST
{
    DWORD	num;
    DWORD	numofstr;
    BYTE	strings[1];
}
SST;

typedef struct XF5
{
    WORD	font;
    WORD	format;
    WORD	type;
    WORD	align;
    WORD	color;
    WORD	fill;
    WORD	border;
    WORD	linestyle;
}
XF5;

typedef struct XF8
{
    WORD	font;
    WORD	format;
    WORD	type;
    BYTE	align;
    BYTE	rotation;
    BYTE	ident;
    BYTE	usedattr;
    DWORD	linestyle;
    DWORD	linecolor;
    WORD	groundcolor;
}
XF8;

typedef struct BR_NUMBER
{
    WORD	row;
    WORD	col;
    WORD	xf;
    double value;
}
BR_NUMBER;

typedef struct COLINFO
{
    WORD	first;
    WORD	last;
    WORD	width;
    WORD	xf;
    WORD	flags;
/* There should be an unused WORD field at the end here. However, some files in
 * the wild report it as a BYTE, which results in a boundary-check parse error.
 * Since the value is ignored anyway, we'll just pretend it was never there.
 *
 * See issue https://github.com/evanmiller/libxls/issues/27
 */
}
COLINFO;

typedef struct MERGEDCELLS
{
    WORD	rowf;
    WORD	rowl;
    WORD	colf;
    WORD	coll;
}
MERGEDCELLS;

typedef struct FONT
{
    WORD	height;
    WORD	flag;
    WORD	color;
    WORD	bold;
    WORD	escapement;
    BYTE	underline;
    BYTE	family;
    BYTE	charset;
    BYTE	notused;
    char    name[1];
}
FONT;

typedef struct FORMAT
{
    WORD	index;
    char	value[1];
}
FORMAT;


typedef struct SUPBOOK
{
    WORD  ctab;
    WORD  cch;
    BYTE  data[1];
}
SUPBOOK;


typedef struct XTI
{
    WORD  supbook;
    WORD  tabfirst;
    WORD  tablast;
}
XTI;

typedef struct XTIARRAY
{
    WORD  length;
    struct XTI  xti[1];
}
XTIARRAY;

typedef struct LBL
{
    WORD  flags;
    BYTE  shortcutkey;
    BYTE  namelength;
    WORD  formulalength;
    WORD  reserved3;
    WORD  tabindex;
    BYTE  reserved4;
    BYTE  reserved5;
    BYTE  reserved6;
    BYTE  reserved7;
    char  data[1];
}
LBL;

#pragma pack(pop)

//---------------------------------------------------------

typedef	struct st_sheet
{
    DWORD count;        // Count of sheets
    struct st_sheet_data
    {
        DWORD filepos;
        BYTE visibility;
        BYTE type;
        char * name;
    }
    * sheet;
}
st_sheet;

typedef	struct st_font
{
    DWORD count;		// Count of FONT's
    struct st_font_data
    {
        WORD	height;
        WORD	flag;
        WORD	color;
        WORD	bold;
        WORD	escapement;
        BYTE	underline;
        BYTE	family;
        BYTE	charset;
        char *	name;
    }
    * font;
}
st_font;

typedef struct st_format
{
    DWORD count;		// Count of FORMAT's
    struct st_format_data
    {
         WORD index;
         char *value;
    }
    * format;
}
st_format;

typedef	struct st_xf
{
    DWORD count;	// Count of XF
    //	XF** xf;
    struct st_xf_data
    {
        WORD	font;
        WORD	format;
        WORD	type;
        BYTE	align;
        BYTE	rotation;
        BYTE	ident;
        BYTE	usedattr;
        DWORD	linestyle;
        DWORD	linecolor;
        WORD	groundcolor;
    }
    * xf;
}
st_xf;


typedef	struct st_sst
{
    DWORD count;
    DWORD lastid;
    DWORD continued;
    DWORD lastln;
    DWORD lastrt;
    DWORD lastsz;
    struct str_sst_string
    {
        char * str;
    }
    * string;
}
st_sst;


typedef	struct st_cell
{
    DWORD count;
    struct st_cell_data
    {
        WORD	id;
        WORD	row;
        WORD	col;
        WORD	xf;
        char *	str;		// String value;
        double	d;
        int32_t	l;
        WORD	width;		// Width of col
        WORD	colspan;
        WORD	rowspan;
        BYTE	isHidden;	// Is cell hidden
    }
    * cell;
}
st_cell;


typedef	struct st_row
{
    //	DWORD count;
    WORD lastcol;	// numCols - 1
    WORD lastrow;	// numRows - 1
    struct st_row_data
    {
        WORD index;
        WORD fcell;
        WORD lcell;
        WORD height;
        WORD flags;
        WORD xf;
        BYTE xfflags;
        st_cell cells;
    }
    * row;
}
st_row;


typedef	struct st_colinfo
{
    DWORD count;				// Count of COLINFO
    struct st_colinfo_data
    {
        WORD	first;
        WORD	last;
        WORD	width;
        WORD	xf;
        WORD	flags;
    }
    * col;
}
st_colinfo;

typedef enum e_reftype {
  ref_self,
  ref_addin,
  ref_extern
} e_reftype;

typedef struct st_supbook
{
  DWORD count;
  struct st_supbook_data
  {
    enum e_reftype reftype;
  } * supbook;
}
st_supbook;

typedef struct st_externsheet
{
  DWORD count;
  struct st_externsheet_data
  {
    WORD index;
    int16_t tabfirst;
    int16_t tablast;
  } * sheet;
}
st_externsheet;

typedef struct st_definedname
{
  DWORD count;
  struct st_definedname_data
  {
    char *  name;
    int tabindex;
    size_t  formulalen;
    BYTE *  formulabytes;
  } * name;
}
st_definedname;


typedef struct xlsWorkBook
{
    //FILE*		file;
    OLE2Stream*	olestr;
    int32_t		filepos;		// position in file

    //From Header (BIFF)
    BYTE		is5ver;
    BYTE		is1904;
    WORD		type;
    WORD		activeSheetIdx;	// index of the active sheet

    //Other data
    WORD		codepage;		// Charset codepage
    char*		charset;
    st_sheet	sheets;
    st_sst		sst;			// SST table
    st_xf		xfs;			// XF table
    st_font		fonts;
    st_format	formats;		// FORMAT table
    st_supbook  supbooks;
    st_externsheet  externsheets;
    st_definedname  definednames;

	char		*summary;		// ole file
	char		*docSummary;	// ole file

    void        *converter;
    void        *utf16_converter;
    void        *utf8_locale;
}
xlsWorkBook;

typedef struct xlsWorkSheet
{
    DWORD		filepos;
    WORD		defcolwidth;
    st_row		rows;
    xlsWorkBook *workbook;
    st_colinfo	colinfo;
}
xlsWorkSheet;

#ifdef __cplusplus
typedef struct st_cell::st_cell_data xlsCell;
typedef	struct st_row::st_row_data xlsRow;
#else
typedef struct st_cell_data xlsCell;
typedef	struct st_row_data xlsRow;
#endif

typedef struct xls_summaryInfo
{
	BYTE		*title;
	BYTE		*subject;
	BYTE		*author;
	BYTE		*keywords;
	BYTE		*comment;
	BYTE		*lastAuthor;
	BYTE		*appName;
	BYTE		*category;
	BYTE		*manager;
	BYTE		*company;
}
xlsSummaryInfo;

typedef void (*xls_formula_handler)(WORD bof, WORD len, BYTE *formula);

#endif
