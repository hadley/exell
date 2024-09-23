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

#include "config.h"

#include <stdio.h>
#include <stdlib.h>
#include <stddef.h>
#include <errno.h>

#ifdef HAVE_ICONV
#include <iconv.h>
#endif

#include <memory.h>
#include <sys/types.h>
#include <string.h>
#include <wchar.h>

#include "libxls/endian.h"
#include "libxls/locale.h"
#include "libxls/xls.h"

#ifndef min
#define min(a,b) ((a) < (b) ? (a) : (b))
#endif

//#define DEBUG_DRAWINGS
int xls_debug = 0;

static double NumFromRk(DWORD drk);
static int escape_sheetname(char *sheetname);
static xls_formula_handler formula_handler;

static xls_error_t xls_addSST(xlsWorkBook* pWB, SST* sst, DWORD size);
static xls_error_t xls_appendSST(xlsWorkBook* pWB, BYTE* buf, DWORD size);
static xls_error_t xls_addFormat(xlsWorkBook* pWB, FORMAT* format, DWORD size);
static xls_error_t xls_addSupBook(xlsWorkBook* pWB, SUPBOOK* supbook, DWORD size);
static xls_error_t xls_addExternSheet(xlsWorkBook* pWB, XTIARRAY* rgxti, DWORD size);
static xls_error_t xls_addDefinedName(xlsWorkBook* pWB, LBL* lbl, DWORD size);
static xls_error_t xls_addSheet(xlsWorkBook* pWB, BOUNDSHEET* bs, DWORD size);
static xls_error_t xls_addRow(xlsWorkSheet* pWS,ROW* row);
static xls_error_t xls_makeTable(xlsWorkSheet* pWS);
static struct st_cell_data *xls_addCell(xlsWorkSheet* pWS, BOF* bof, BYTE* buf);
static char *xls_addFont(xlsWorkBook* pWB, FONT* font, DWORD size);
static xls_error_t xls_addXF8(xlsWorkBook* pWB, XF8* xf);
static xls_error_t xls_addXF5(xlsWorkBook* pWB, XF5* xf);
static xls_error_t xls_addColinfo(xlsWorkSheet* pWS, COLINFO* colinfo);
static xls_error_t xls_mergedCells(xlsWorkSheet* pWS, BOF* bof, BYTE* buf);
static xls_error_t xls_preparseWorkSheet(xlsWorkSheet* pWS);
static xls_error_t xls_formatColumn(xlsWorkSheet* pWS);
static void xls_dumpSummary(char *buf, int isSummary, xlsSummaryInfo *pSI);

#if defined(_AIX) || defined(__sun)
#pragma pack(1)
#else
#pragma pack(push, 1)
#endif

typedef struct {
	uint32_t		format[4];
	uint32_t		offset;
} sectionList;

typedef struct {
	uint16_t		sig;
	uint16_t		_empty;
	uint32_t		os;
	uint32_t		format[4];
	uint32_t		count;
	sectionList		secList[1];
} header;

typedef struct {
	uint32_t		propertyID;
	uint32_t		sectionOffset;
} propertyList;

typedef struct {
	uint32_t		length;
	uint32_t		numProperties;
	propertyList	properties[1];
} sectionHeader;

typedef struct {
	uint32_t		propertyID;
	uint32_t		data[1];
} property;

typedef struct {
  uint16_t row;
  int16_t col;
} RGCE_LOCREL;

typedef struct {
  uint16_t row_first;
  uint16_t row_last;
  int16_t col_first;
  int16_t col_last;
} RGCE_AREAREL;


#pragma pack(pop)

int xls(int debug)
{
	xls_debug = debug;
    return 1;
}

static xls_error_t xls_addSST(xlsWorkBook* pWB,SST* sst,DWORD size)
{
    verbose("xls_addSST");

    pWB->sst.continued=0;
    pWB->sst.lastln=0;
    pWB->sst.lastid=0;
    pWB->sst.lastrt=0;
    pWB->sst.lastsz=0;

    if (sst->num > (1<<26)) // 64 MB
        return LIBXLS_ERROR_MALLOC;

    if (pWB->sst.string)
        return LIBXLS_ERROR_PARSE;

    if ((pWB->sst.string = calloc(pWB->sst.count = sst->num,
                    sizeof(struct str_sst_string))) == NULL)
        return LIBXLS_ERROR_MALLOC;

    return xls_appendSST(pWB, sst->strings, size - offsetof(SST, strings));
}

static xls_error_t xls_appendSST(xlsWorkBook* pWB, BYTE* buf, DWORD size)
{
    DWORD ln;	// String character count
    DWORD ofs;	// Current offset in SST buffer
    DWORD rt;	// Count of rich text formatting runs
    DWORD sz;	// Size of asian phonetic settings block
    BYTE flag;	// String flags
    char* ret = NULL;

    if (xls_debug) {
	    printf("xls_appendSST %u\n", size);
    }

	sz = rt = ln = 0;	// kch
    ofs=0;

	while(ofs<size)
    {
        int ln_toread;

        // Restore state when we're in a continue record
        // or read string length
        if (pWB->sst.continued) {
            ln=pWB->sst.lastln;
            rt=pWB->sst.lastrt;
            sz=pWB->sst.lastsz;
        } else {
            if (ofs + 2 > size) {
                return LIBXLS_ERROR_PARSE;
            }
            ln = buf[ofs+0] + (buf[ofs+1] << 8);
            rt = 0;
            sz = 0;

            ofs+=2;
        }

		if (xls_debug) {
        	printf("ln=%u\n", ln);
		}

        // Read flags
        if ( !pWB->sst.continued || (pWB->sst.continued && ln != 0) ) {
            if (ofs + sizeof(BYTE) > size) {
                return LIBXLS_ERROR_PARSE;
            }
            flag=*(BYTE *)(buf+ofs);
            ofs++;

            // Count of rich text formatting runs
            if (flag & 0x8) {
                if (ofs + sizeof(WORD) > size) {
                    return LIBXLS_ERROR_PARSE;
                }
                rt = buf[ofs+0] + (buf[ofs+1] << 8);
                ofs+=2;
            }

            // Size of asian phonetic settings block
            if (flag & 0x4) {
                if (ofs + sizeof(DWORD) > size) {
                    return LIBXLS_ERROR_PARSE;
                }
                sz = buf[ofs+0] + (buf[ofs+1] << 8) + (buf[ofs+2] << 16) + ((DWORD)buf[ofs+3] << 24);
                ofs+=4;

				if (xls_debug) {
					printf("sz=%u\n", sz);
				}
            }
        } else {
            flag = 0;
        }

		// Read characters (compressed or not)
        ln_toread = 0;
        if (ln > 0) {
            if (flag & 0x1) {
                size_t new_len = 0;
                ln_toread = min((size-ofs)/2, ln);
                ret=unicode_decode((char *)buf+ofs, ln_toread*2, pWB);

                if (ret == NULL) {
                    ret = strdup("*failed to decode utf16*");
                }

                ln -= ln_toread;
                ofs+=ln_toread*2;

                if (xls_debug) {
                    new_len = strlen(ret);
	                printf("String16SST: %s(%lu)\n", ret, (unsigned long)new_len);
                }
            } else {
                ln_toread = min((size-ofs), ln);

                ret = codepage_decode((char *)buf+ofs, ln_toread, pWB);
                if (ret == NULL) {
                    ret = strdup("*failed to decode BIFF5 string*");
                }

                ln  -= ln_toread;
                ofs += ln_toread;

                if (xls_debug) {
                	printf("String8SST: %s(%u) \n",ret,ln);
                }
            }
        } else {
            ret = strdup("");
        }

        if (ln_toread > 0 || !pWB->sst.continued) {
            // Concat string if it's a continue, or add string in table
            if (!pWB->sst.continued) {
                if (pWB->sst.lastid >= pWB->sst.count) {
                    free(ret);
                    return LIBXLS_ERROR_PARSE;
                }
                pWB->sst.lastid++;
                pWB->sst.string[pWB->sst.lastid-1].str=ret;
            } else {
                char *tmp = pWB->sst.string[pWB->sst.lastid-1].str;
                if (tmp == NULL) {
                    free(ret);
                    return LIBXLS_ERROR_PARSE;
                }
                tmp = realloc(tmp, strlen(tmp)+strlen(ret)+1);
                if (tmp == NULL)  {
                    free(ret);
                    return LIBXLS_ERROR_MALLOC;
                }
                pWB->sst.string[pWB->sst.lastid-1].str=tmp;
                memcpy(tmp+strlen(tmp), ret, strlen(ret)+1);
				free(ret);
            }

			if (xls_debug) {
	            printf("String %4u: %s<end>\n", pWB->sst.lastid-1, pWB->sst.string[pWB->sst.lastid-1].str);
			}
        } else {
            free(ret);
	}

		// Jump list of rich text formatting runs
        if (ofs < size && rt > 0) {
            int rt_toread = min((size-ofs)/4, rt);
            rt -= rt_toread;
            ofs += rt_toread*4;
        }

		// Jump asian phonetic settings block
        if (ofs < size && sz > 0) {
            int sz_toread = min((size-ofs), sz);
            sz -= sz_toread;
            ofs += sz_toread;
        }

        pWB->sst.continued=0;
    }

    // Save current character count and count of rich text formatting runs and size of asian phonetic settings block
	if (ln > 0 || rt > 0 || sz > 0) {
		pWB->sst.continued = 1;
		pWB->sst.lastln = ln;
		pWB->sst.lastrt = rt;
		pWB->sst.lastsz = sz;

		if (xls_debug) {
			printf("continued: ln=%u, rt=%u, sz=%u\n", ln, rt, sz);
		}
	}

    return LIBXLS_OK;
}

static double NumFromRk(DWORD drk)
{
	double ret;

	// What kind of value is this ?
    if (drk & 0x02) {
    	// Integer value
		int tmp = (int)drk >> 2;	// cast to keep it negative in < 0
        ret = (double)tmp;
    } else {
    	// Floating point value;
		unsigned64_t tmp = drk & 0xfffffffc;
		tmp <<= 32;
		memcpy(&ret, &tmp, sizeof(unsigned64_t));
    }
    // Is value multiplied by 100 ?
    if (drk & 0x01) {
        ret /= 100.0;
    }
    return ret;
}


static xls_error_t xls_addSupBook(xlsWorkBook* pWB, SUPBOOK* supbook, DWORD size)
{
  pWB->supbooks.supbook = realloc(pWB->supbooks.supbook,(pWB->supbooks.count+1)*sizeof (struct st_supbook_data));
  if (pWB->supbooks.supbook == NULL) {
    return LIBXLS_ERROR_MALLOC;
  }

  struct st_supbook_data *sb = &(pWB->supbooks.supbook[pWB->supbooks.count]);

  switch(supbook->cch) {
  case 0x0401:
    sb->reftype = ref_self;
    break;
  case 0x3A01:
    sb->reftype = ref_addin;
    break;
  default:
    sb->reftype = ref_extern;
    // not implemented: reading external workbook and sheet names
    break;
  }

  pWB->supbooks.count++;

  return LIBXLS_OK;
}

static xls_error_t xls_addExternSheet(xlsWorkBook* pWB, XTIARRAY *rgxti, DWORD size)
{
  int n = rgxti->length;

  pWB->externsheets.sheet = realloc(pWB->externsheets.sheet,(pWB->externsheets.count+n)*sizeof(struct st_externsheet_data));

  struct st_externsheet_data *ext = &(pWB->externsheets.sheet[pWB->externsheets.count]);
  for(int i = 0; i < n; i++, ext++) {
    ext->index = rgxti->xti[i].supbook;
    ext->tabfirst = rgxti->xti[i].tabfirst;
    ext->tablast = rgxti->xti[i].tablast;
  }

  pWB->externsheets.count += n;

  return LIBXLS_OK;
}


static xls_error_t xls_addDefinedName(xlsWorkBook* pWB, LBL* lbl, DWORD size)
{
  char *name;
  BYTE *formulabytes;

  int is2 = 0;
  int ln = 0;

  is2 = (int)(lbl->data[0]);
  if(is2) {
    ln = lbl->namelength*2;
    name = unicode_decode(lbl->data + 1, ln, pWB);
  } else {
    ln = lbl->namelength;
    name = codepage_decode(lbl->data + 1, ln, pWB);
  }
  if (name == NULL)
    return LIBXLS_ERROR_MALLOC;

  int ofs = 1 + ln;

  is2 = (int)(lbl->data[ofs]);
  formulabytes = malloc(lbl->formulalength);
  if(formulabytes == NULL) {
    free(name);
    return LIBXLS_ERROR_MALLOC;
  }

  memcpy(formulabytes, lbl->data + ofs, lbl->formulalength);

  pWB->definednames.name = realloc(pWB->definednames.name,(pWB->definednames.count+1)*sizeof(struct st_definedname_data));
  if (pWB->definednames.name == NULL) {
    free(name);
    free(formulabytes);
    return LIBXLS_ERROR_MALLOC;
  }

  pWB->definednames.name[pWB->definednames.count].name=name;
  pWB->definednames.name[pWB->definednames.count].tabindex=lbl->tabindex;
  pWB->definednames.name[pWB->definednames.count].formulalen=lbl->formulalength;
  pWB->definednames.name[pWB->definednames.count].formulabytes=formulabytes;
  pWB->definednames.count++;

  return LIBXLS_OK;
}


static xls_error_t xls_addSheet(xlsWorkBook* pWB, BOUNDSHEET *bs, DWORD size)
{
	char * name;
	DWORD filepos;
	BYTE visible, type;

	filepos = bs->filepos;
	visible = bs->visible;
	type = bs->type;

	// printf("charset=%s uni=%d\n", pWB->charset, unicode);
	// printf("bs name %.*s\n", bs->name[0], bs->name+1);
	name = get_string(bs->name, size - offsetof(BOUNDSHEET, name), 0, pWB);
	// printf("name=%s\n", name);

	if(xls_debug) {
		printf ("xls_addSheet[0x%x]\n", type);
		switch (type & 0x0f)
		{
		case 0x00:
			/* worksheet or dialog sheet */
			printf ("85: Worksheet or dialog sheet\n");
			break;
		case 0x01:
			/* Microsoft Excel 4.0 macro sheet */
			printf ("85: Microsoft Excel 4.0 macro sheet\n");
			break;
		case 0x02:
			/* Chart */
			printf ("85: Chart sheet\n");
			break;
		case 0x06:
			/* Visual Basic module */
			printf ("85: Visual Basic sheet\n");
			break;
		default:
			printf ("???\n");
			break;
		}
		printf("visible: %x\n", visible);
		printf("    Pos: %Xh\n",filepos);
		printf("   type: %.4Xh\n",type);
		printf("   name: %s\n", name);
	}

    pWB->sheets.sheet = realloc(pWB->sheets.sheet,(pWB->sheets.count+1)*sizeof (struct st_sheet_data));
    if (pWB->sheets.sheet == NULL) {
        free(name);
        return LIBXLS_ERROR_MALLOC;
    }

    pWB->sheets.sheet[pWB->sheets.count].name=name;
    pWB->sheets.sheet[pWB->sheets.count].filepos=filepos;
    pWB->sheets.sheet[pWB->sheets.count].visibility=visible;
    pWB->sheets.sheet[pWB->sheets.count].type=type;
    pWB->sheets.count++;

	return LIBXLS_OK;
}


static xls_error_t xls_addRow(xlsWorkSheet* pWS,ROW* row)
{
    struct st_row_data* tmp;

    //verbose ("xls_addRow");

    if (row->index > pWS->rows.lastrow)
        return LIBXLS_ERROR_PARSE;

    tmp=&pWS->rows.row[row->index];
    tmp->height=row->height;
    tmp->fcell=row->fcell;
    tmp->lcell=row->lcell;
    tmp->flags=row->flags;
    tmp->xf=row->xf&0xfff;
    tmp->xfflags=(row->xf >> 8)&0xf0;
    if(xls_debug) xls_showROW(tmp);

    return LIBXLS_OK;
}

static xls_error_t xls_makeTable(xlsWorkSheet* pWS)
{
    DWORD i,t;
    struct st_row_data* tmp;
    verbose ("xls_makeTable");

    if ((pWS->rows.row = calloc((pWS->rows.lastrow+1),sizeof(struct st_row_data))) == NULL)
        return LIBXLS_ERROR_MALLOC;

	// printf("ALLOC: rows=%d cols=%d\n", pWS->rows.lastrow, pWS->rows.lastcol);
    for (t=0;t<=pWS->rows.lastrow;t++)
    {
        tmp=&pWS->rows.row[t];
        tmp->index=t;
        tmp->fcell=0;
        tmp->lcell=pWS->rows.lastcol;

		tmp->cells.count = pWS->rows.lastcol+1;
        if ((tmp->cells.cell = calloc(tmp->cells.count, sizeof(struct st_cell_data))) == NULL)
            return LIBXLS_ERROR_MALLOC;

        for (i=0;i<=pWS->rows.lastcol;i++)
        {
            tmp->cells.cell[i].col = i;
            tmp->cells.cell[i].row = t;
            tmp->cells.cell[i].width = pWS->defcolwidth;
            tmp->cells.cell[i].id = XLS_RECORD_BLANK;
        }
    }
    return LIBXLS_OK;
}

int xls_isCellTooSmall(xlsWorkBook* pWB, BOF* bof, BYTE* buf) {
    if (bof->size < sizeof(COL))
        return 1;

    if (bof->id == XLS_RECORD_FORMULA || bof->id == XLS_RECORD_FORMULA_ALT)
        return (bof->size < sizeof(FORMULA));

    if (bof->id == XLS_RECORD_MULRK)
        return (bof->size < offsetof(MULRK, rk));

    if (bof->id == XLS_RECORD_MULBLANK)
        return (bof->size < offsetof(MULBLANK, xf));

    if (bof->id == XLS_RECORD_LABELSST)
        return (bof->size < offsetof(LABEL, value) + (pWB->is5ver ? 2 : 4));

    if (bof->id == XLS_RECORD_LABEL || bof->id == XLS_RECORD_RSTRING) {
        if (bof->size < offsetof(LABEL, value) + 2)
            return 1;

        size_t label_len = ((LABEL*)buf)->value[0] + (((LABEL*)buf)->value[1] << 8);
        if (pWB->is5ver) {
            return (bof->size < offsetof(LABEL, value) + 2 + label_len);
        }

        if (bof->size < offsetof(LABEL, value) + 3)
            return 1;

        if ((((LABEL*)buf)->value[2] & 0x01) == 0) {
            return (bof->size < offsetof(LABEL, value) + 3 + label_len);
        }
        return (bof->size < offsetof(LABEL, value) + 3 + 2 * label_len);
    }

    if (bof->id == XLS_RECORD_RK)
        return (bof->size < sizeof(RK));

    if (bof->id == XLS_RECORD_NUMBER)
        return (bof->size < sizeof(BR_NUMBER));

    if (bof->id == XLS_RECORD_BOOLERR)
        return (bof->size < sizeof(BOOLERR));

    return 0;
}

void xls_cell_set_str(struct st_cell_data *cell, char *str) {
    if (cell->str) {
        free(cell->str);
    }
    cell->str = str;
}

static struct st_cell_data *xls_addCell(xlsWorkSheet* pWS,BOF* bof,BYTE* buf)
{
    struct st_cell_data*	cell;
    struct st_row_data*		row;
    WORD                    col;
    int						i;

	verbose ("xls_addCell");

    if (xls_isCellTooSmall(pWS->workbook, bof, buf))
        return NULL;

	// printf("ROW: %u COL: %u\n", xlsShortVal(((COL*)buf)->row), xlsShortVal(((COL*)buf)->col));
    row=&pWS->rows.row[xlsShortVal(((COL*)buf)->row)];

    col = xlsShortVal(((COL*)buf)->col);
    if (col >= row->cells.count) {
        if (xls_debug) fprintf(stderr, "Error: Column index out of bounds\n");
        return NULL;
    }
    cell = &row->cells.cell[col];

    cell->id=bof->id;
    cell->xf=xlsShortVal(((COL*)buf)->xf);

    switch (bof->id)
    {
    case XLS_RECORD_FORMULA:
    case XLS_RECORD_FORMULA_ALT:
		xlsConvertFormula((FORMULA *)buf);
        cell->id=XLS_RECORD_FORMULA;
        if (((FORMULA*)buf)->res!=0xffff) {
			// if a double, then set double and clear l
			cell->l=0;
			memcpy(&cell->d, &((FORMULA*)buf)->resid, sizeof(double));	// Required for ARM
            cell->id = XLS_RECORD_NUMBER; // hack
            xls_cell_set_str(cell, xls_getfcell(pWS->workbook,cell, NULL));
            cell->id = bof->id;
		} else {
			double d = ((FORMULA*)buf)->resdata[1];
			cell->l = 0xFFFF;
			switch(((FORMULA*)buf)->resid) {
			case 0:		// String
				break;	// cell is half complete, get the STRING next record
			case 1:		// Boolean
				memcpy(&cell->d, &d, sizeof(double)); // Required for ARM
                xls_cell_set_str(cell, strdup("bool"));
				break;
			case 2:		// error
				memcpy(&cell->d, &d, sizeof(double)); // Required for ARM
                xls_cell_set_str(cell, strdup("error"));
				break;
			case 3:		// empty string
                xls_cell_set_str(cell, strdup(""));
				break;
			}
		}
		if(formula_handler) formula_handler(bof->id, bof->size, buf);
        break;
    case XLS_RECORD_MULRK:
        for (i = 0; i < (bof->size - 6)/6; i++)	// 6 == 2 row + 2 col + 2 trailing index
        {
            WORD index = col + i;
            if(index >= row->cells.count) {
                if (xls_debug) fprintf(stderr, "Error: MULTI-RK index out of bounds\n");
                return NULL;
            }
            cell=&row->cells.cell[index];
            cell->id=XLS_RECORD_RK;
            cell->xf=xlsShortVal(((MULRK*)buf)->rk[i].xf);
            cell->d=NumFromRk(xlsIntVal(((MULRK*)buf)->rk[i].value));
            xls_cell_set_str(cell, xls_getfcell(pWS->workbook,cell, NULL));
        }
        break;
    case XLS_RECORD_MULBLANK:
        for (i = 0; i < (bof->size - 6)/2; i++)	// 6 == 2 row + 2 col + 2 trailing index
        {
            WORD index = col + i;
            if(index >= row->cells.count) {
                if (xls_debug) fprintf(stderr, "Error: MULTI-BLANK index out of bounds\n");
                return NULL;
            }
            cell=&row->cells.cell[index];
            cell->id=XLS_RECORD_BLANK;
            cell->xf=xlsShortVal(((MULBLANK*)buf)->xf[i]);
            xls_cell_set_str(cell, xls_getfcell(pWS->workbook,cell, NULL));
        }
        break;
    case XLS_RECORD_LABELSST:
    case XLS_RECORD_LABEL:
    case XLS_RECORD_RSTRING:
        xls_cell_set_str(cell, xls_getfcell(pWS->workbook, cell, ((LABEL*)buf)->value));
        if (cell->str) {
            sscanf((char *)cell->str, "%d", &cell->l);
            sscanf((char *)cell->str, "%lf", &cell->d);
        }
		break;
    case XLS_RECORD_RK:
        cell->d=NumFromRk(xlsIntVal(((RK*)buf)->value));
        xls_cell_set_str(cell, xls_getfcell(pWS->workbook,cell, NULL));
        break;
    case XLS_RECORD_BLANK:
        break;
    case XLS_RECORD_NUMBER:
        xlsConvertDouble((BYTE *)&((BR_NUMBER*)buf)->value);
		memcpy(&cell->d, &((BR_NUMBER*)buf)->value, sizeof(double)); // Required for ARM
        xls_cell_set_str(cell, xls_getfcell(pWS->workbook,cell, NULL));
        break;
    case XLS_RECORD_BOOLERR:
        cell->d = ((BOOLERR *)buf)->value;
        if (((BOOLERR *)buf)->iserror) {
            xls_cell_set_str(cell, strdup("error"));
        } else {
            xls_cell_set_str(cell, strdup("bool"));
        }
        break;
    default:
        xls_cell_set_str(cell, xls_getfcell(pWS->workbook,cell, NULL));
        break;
    }
    if (xls_debug) xls_showCell(cell);

	return cell;
}

static char *xls_addFont(xlsWorkBook* pWB, FONT* font, DWORD size)
{
    struct st_font_data* tmp;

    verbose("xls_addFont");

    pWB->fonts.font = realloc(pWB->fonts.font,(pWB->fonts.count+1)*sizeof(struct st_font_data));
    if (pWB->fonts.font == NULL)
        return NULL;

    tmp=&pWB->fonts.font[pWB->fonts.count];

    tmp->name = get_string(font->name, size - offsetof(FONT, name), 0, pWB);

    tmp->height=font->height;
    tmp->flag=font->flag;
    tmp->color=font->color;
    tmp->bold=font->bold;
    tmp->escapement=font->escapement;
    tmp->underline=font->underline;
    tmp->family=font->family;
    tmp->charset=font->charset;

    //	xls_showFont(tmp);
    pWB->fonts.count++;

	return tmp->name;
}

static xls_error_t xls_addFormat(xlsWorkBook* pWB, FORMAT* format, DWORD size)
{
    struct st_format_data* tmp;

    verbose("xls_addFormat");
    pWB->formats.format = realloc(pWB->formats.format, (pWB->formats.count+1)*sizeof(struct st_format_data));
    if (pWB->formats.format == NULL)
        return LIBXLS_ERROR_MALLOC;

    tmp = &pWB->formats.format[pWB->formats.count];
    tmp->index = format->index;
    tmp->value = get_string(format->value, size - offsetof(FORMAT, value), (BYTE)!pWB->is5ver, pWB);
    if(xls_debug) xls_showFormat(tmp);
    pWB->formats.count++;

    return LIBXLS_OK;
}

static xls_error_t xls_addXF8(xlsWorkBook* pWB,XF8* xf)
{
    struct st_xf_data* tmp;

    verbose("xls_addXF");
    pWB->xfs.xf= realloc(pWB->xfs.xf, (pWB->xfs.count+1)*sizeof(struct st_xf_data));
    if (pWB->xfs.xf == NULL)
        return LIBXLS_ERROR_MALLOC;

    tmp=&pWB->xfs.xf[pWB->xfs.count];

    tmp->font=xf->font;
    tmp->format=xf->format;
    tmp->type=xf->type;
    tmp->align=xf->align;
    tmp->rotation=xf->rotation;
    tmp->ident=xf->ident;
    tmp->usedattr=xf->usedattr;
    tmp->linestyle=xf->linestyle;
    tmp->linecolor=xf->linecolor;
    tmp->groundcolor=xf->groundcolor;

    //	xls_showXF(tmp);
    pWB->xfs.count++;

    return LIBXLS_OK;
}

static xls_error_t xls_addXF5(xlsWorkBook* pWB,XF5* xf)
{
    struct st_xf_data* tmp;

    verbose("xls_addXF");
    pWB->xfs.xf = realloc(pWB->xfs.xf, (pWB->xfs.count+1)*sizeof(struct st_xf_data));
    if (pWB->xfs.xf == NULL)
        return LIBXLS_ERROR_MALLOC;

    tmp=&pWB->xfs.xf[pWB->xfs.count];

    tmp->font=xf->font;
    tmp->format=xf->format;
    tmp->type=xf->type;
    tmp->align=(BYTE)xf->align;
/*
    tmp->rotation=xf->rotation;
    tmp->ident=xf->ident;
    tmp->usedattr=xf->usedattr;
    tmp->linestyle=xf->linestyle;
    tmp->linecolor=xf->linecolor;
    tmp->groundcolor=xf->groundcolor;
*/

    //	xls_showXF(tmp);
    pWB->xfs.count++;
    return LIBXLS_OK;
}

static xls_error_t xls_addColinfo(xlsWorkSheet* pWS,COLINFO* colinfo)
{
    struct st_colinfo_data* tmp;

    verbose("xls_addColinfo");
    pWS->colinfo.col =  realloc(pWS->colinfo.col,(pWS->colinfo.count+1)*sizeof(struct st_colinfo_data));
    if (pWS->colinfo.col == NULL)
        return LIBXLS_ERROR_MALLOC;

    tmp=&pWS->colinfo.col[pWS->colinfo.count];
    tmp->first=colinfo->first;
    tmp->last=colinfo->last;
    tmp->width=colinfo->width;
    tmp->xf=colinfo->xf;
    tmp->flags=colinfo->flags;

    if(xls_debug) xls_showColinfo(tmp);
    pWS->colinfo.count++;

    return LIBXLS_OK;
}

static xls_error_t xls_mergedCells(xlsWorkSheet* pWS,BOF* bof,BYTE* buf)
{
    if (bof->size < sizeof(WORD))
        return LIBXLS_ERROR_PARSE;

    int count = buf[0] + (buf[1] << 8);
    DWORD limit = sizeof(WORD)+count*sizeof(struct MERGEDCELLS);
    if(limit > (DWORD)bof->size) {
        verbose("Merged Cells Count out of range");
        return LIBXLS_ERROR_PARSE;
    }
    int i,c,r;
    struct MERGEDCELLS *span;
    verbose("Merged Cells");
    for (i=0;i<count;i++)
    {
        span=(struct MERGEDCELLS*)(buf+(2+i*sizeof(struct MERGEDCELLS)));
        xlsConvertMergedcells(span);
        //		printf("Merged Cells: [%i,%i] [%i,%i] \n",span->colf,span->rowf,span->coll,span->rowl);
        // Sanity check:
        if(!(   span->rowf <= span->rowl &&
                span->rowl <= pWS->rows.lastrow &&
                span->colf <= span->coll &&
                span->coll <= pWS->rows.lastcol
        )) {
            return LIBXLS_ERROR_PARSE;
        }

        for (r=span->rowf;r<=span->rowl;r++)
            for (c=span->colf;c<=span->coll;c++)
                pWS->rows.row[r].cells.cell[c].isHidden=1;
        pWS->rows.row[span->rowf].cells.cell[span->colf].colspan=(span->coll-span->colf+1);
        pWS->rows.row[span->rowf].cells.cell[span->colf].rowspan=(span->rowl-span->rowf+1);
        pWS->rows.row[span->rowf].cells.cell[span->colf].isHidden=0;
    }
    return LIBXLS_OK;
}

int xls_isRecordTooSmall(xlsWorkBook *pWB, BOF *bof1) {
    switch (bof1->id) {
        case XLS_RECORD_BOF:	// BIFF5-8
            return (bof1->size < 2 * sizeof(WORD));
        case XLS_RECORD_CODEPAGE:
            return (bof1->size < sizeof(WORD));
		case XLS_RECORD_WINDOW1:
            return (bof1->size < sizeof(WIND1));
        case XLS_RECORD_SST:
            return (bof1->size < offsetof(SST, strings));
        case XLS_RECORD_BOUNDSHEET:
            return (bof1->size < offsetof(BOUNDSHEET, name));
        case XLS_RECORD_XF:
			if(pWB->is5ver) {
                return (bof1->size < sizeof(XF5));
            }
            return (bof1->size < sizeof(XF8));
        case XLS_RECORD_FONT:
        case XLS_RECORD_FONT_ALT:
            return (bof1->size < offsetof(FONT, name));
        case XLS_RECORD_FORMAT:
            return (bof1->size < offsetof(FORMAT, value));
		case XLS_RECORD_1904:
            return (bof1->size < sizeof(BYTE));
        default:
            break;
    }
    return 0;
}

xls_error_t xls_parseWorkBook(xlsWorkBook* pWB)
{
    if(!pWB) return LIBXLS_ERROR_NULL_ARGUMENT;

    BOF bof1 = { .id = 0, .size = 0 };
    BOF bof2 = { .id = 0, .size = 0 };
    BYTE* buf = NULL;
	BYTE once = 0;
    xls_error_t retval = LIBXLS_OK;

    verbose ("xls_parseWorkBook");
    do {
		if(xls_debug > 10) {
			printf("READ WORKBOOK filePos=%ld\n",  (long)pWB->filepos);
			printf("  OLE: start=%d pos=%u size=%u fatPos=%u\n",
                    pWB->olestr->start, (unsigned int)pWB->olestr->pos,
                    (unsigned int)pWB->olestr->size, (unsigned int)pWB->olestr->fatpos); 
		}

        if (ole2_read(&bof1, 1, 4, pWB->olestr) != 4) {
            retval = LIBXLS_ERROR_READ;
            goto cleanup;
        }
        xlsConvertBof(&bof1);
 		if(xls_debug) xls_showBOF(&bof1);

        if (bof1.size) {
            if ((buf = realloc(buf, bof1.size)) == NULL) {
                if (xls_debug) fprintf(stderr, "Error: failed to allocate buffer of size %d\n", (int)bof1.size);
                retval = LIBXLS_ERROR_MALLOC;
                goto cleanup;
            }
            if (ole2_read(buf, 1, bof1.size, pWB->olestr) != bof1.size) {
                if (xls_debug) fprintf(stderr, "Error: failed to read OLE block\n");
                retval = LIBXLS_ERROR_READ;
                goto  cleanup;
            }
        }

        if (xls_isRecordTooSmall(pWB, &bof1)) {
            retval = LIBXLS_ERROR_PARSE;
            goto cleanup;
        }

        switch (bof1.id) {
        case XLS_RECORD_EOF:
            //verbose("EOF");
            break;
        case XLS_RECORD_BOF:	// BIFF5-8
            pWB->is5ver = (buf[0] + (buf[1] << 8) != 0x600);
            pWB->type = buf[2] + (buf[3] << 8);
            if(xls_debug) {
                printf("version: %s\n", pWB->is5ver ? "BIFF5" : "BIFF8" );
                printf("   type: %.2X\n", pWB->type);
            }
            break;

        case XLS_RECORD_CODEPAGE:
            pWB->codepage = buf[0] + (buf[1] << 8);
			if(xls_debug) printf("codepage: %d\n", pWB->codepage);
            break;

        case XLS_RECORD_CONTINUE:
			if(once) {
				if (bof2.id==XLS_RECORD_SST) {
					if ((retval = xls_appendSST(pWB,buf,bof1.size)) != LIBXLS_OK)
                        goto cleanup;
                }
				bof1=bof2;
			}
            break;

		case XLS_RECORD_WINDOW1:
			{
				WIND1 *w = (WIND1*)buf;
                xlsConvertWindow(w);
				pWB->activeSheetIdx = w->itabCur;
				if(xls_debug) {
					printf("WINDOW1: ");
					printf("xWn    : %d\n", w->xWn/20);
					printf("yWn    : %d\n", w->yWn/20);
					printf("dxWn   : %d\n", w->dxWn/20);
					printf("dyWn   : %d\n", w->dyWn/20);
					printf("grbit  : %d\n", w->grbit);
					printf("itabCur: %d\n", w->itabCur);
					printf("itabFi : %d\n", w->itabFirst);
					printf("ctabSel: %d\n", w->ctabSel);
					printf("wTabRat: %d\n", w->wTabRatio);
				}
			}
			break;

        case XLS_RECORD_SST:
			//printf("ADD SST\n");
            xlsConvertSst((SST *)buf);
            if ((retval = xls_addSST(pWB,(SST*)buf,bof1.size)) != LIBXLS_OK) {
                goto cleanup;
            }
            break;

        case XLS_RECORD_EXTSST:
            break;

        case XLS_RECORD_BOUNDSHEET:
			{
				//printf("ADD SHEET\n");
				BOUNDSHEET *bs = (BOUNDSHEET *)buf;
                xlsConvertBoundsheet(bs);
				// different for BIFF5 and BIFF8
                if ((retval = xls_addSheet(pWB, bs, bof1.size)) != LIBXLS_OK) {
                    goto cleanup;
                }
			}
            break;

        case XLS_RECORD_XF:
			if(pWB->is5ver) {
				XF5 *xf;
				xf = (XF5 *)buf;
                xlsConvertXf5(xf);

				if ((retval = xls_addXF5(pWB,xf)) != LIBXLS_OK) {
                    goto cleanup;
                }
				if(xls_debug) {
					printf("   font: %d\n", xf->font);
					printf(" format: %d\n", xf->format);
					printf("   type: %.4x\n", xf->type);
					printf("  align: %.4x\n", xf->align);
					printf("rotatio: %.4x\n", xf->color);
					printf("  ident: %.4x\n", xf->fill);
					printf("usedatt: %.4x\n", xf->border);
					printf("linesty: %.4x\n", xf->linestyle);
				}
			} else {
				XF8 *xf;
				xf = (XF8 *)buf;
                xlsConvertXf8(xf);

				if ((retval = xls_addXF8(pWB,xf)) != LIBXLS_OK) {
                    goto cleanup;
                }

				if(xls_debug) {
					xls_showXF(xf);
				}
			}
            break;

        case XLS_RECORD_FONT:
        case XLS_RECORD_FONT_ALT:
			{
				char *s;
				FONT *f = (FONT*)buf;
                xlsConvertFont(f);
				s = xls_addFont(pWB,f, bof1.size);
				if(xls_debug) {
					printf(" height: %d\n", f->height);
					printf("   flag: 0x%x\n", f->flag);
					printf("  color: 0x%x\n", f->color);
					printf(" weight: %d\n", f->bold);
					printf("escapem: 0x%x\n", f->escapement);
					printf("underln: 0x%x\n", f->underline);
					printf(" family: 0x%x\n", f->family);
					printf("charset: 0x%x\n", f->charset);
					if(s) printf("   name: %s\n", s);
				}
			}
			break;

        case XLS_RECORD_FORMAT:
            xlsConvertFormat((FORMAT *)buf);
            if ((retval = xls_addFormat(pWB, (FORMAT*)buf, bof1.size)) != LIBXLS_OK) {
                goto cleanup;
            }
            break;

		case XLS_RECORD_STYLE:
			if(xls_debug) {
				struct { unsigned short idx; unsigned char ident; unsigned char lvl; } *styl;
				styl = (void *)buf;

				printf("    idx: 0x%x\n", styl->idx & 0x07FF);
				if(styl->idx & 0x8000) {
					printf("  ident: 0x%x\n", styl->ident);
					printf("  level: 0x%x\n", styl->lvl);
				} else {
					char *s = get_string((char *)&buf[2], bof1.size - 2, 1, pWB);
					printf("  name=%s\n", s);
                    free(s);
				}
			}
			break;

        case XLS_RECORD_PALETTE:
			if(xls_debug > 10) {
				unsigned char *p = buf + 2;
				int idx, len;

				len = buf[0] + (buf[1] << 8);
				for(idx=0; idx<len; ++idx) {
					printf("   Index=0x%2.2x %2.2x%2.2x%2.2x\n", idx+8, p[0], p[1], p[2] );
					p += 4;
				}
			}
			break;

		case XLS_RECORD_1904:
			pWB->is1904 = *(BYTE *)buf;	// the field is a short, but with little endian the first byte is 0 or 1
			if(xls_debug) {
				printf("   mode: 0x%x\n", pWB->is1904);
			}
			break;

		case XLS_RECORD_FILEPASS:
			retval = LIBXLS_ERROR_UNSUPPORTED_ENCRYPTION;
			goto cleanup;
		
    case XLS_RECORD_SUPBOOK:
      if ((retval = xls_addSupBook(pWB, (SUPBOOK*)buf, bof1.size)) != LIBXLS_OK) {
        goto cleanup;
      }
      break;

    case XLS_RECORD_EXTERNSHEET:
      if ((retval = xls_addExternSheet(pWB, (XTIARRAY*)buf, bof1.size)) != LIBXLS_OK) {
        goto cleanup;
      }
      break;

		case XLS_RECORD_DEFINEDNAME:
		  if ((retval = xls_addDefinedName(pWB, (LBL*)buf, bof1.size)) != LIBXLS_OK) {
		    goto cleanup;
		  }
		  break;
			if(xls_debug) {
				int i;
                printf("   DEFINEDNAME: ");
				for(i=0; i<bof1.size; ++i) printf("%2.2x ", buf[i]);
				printf("\n");
			}
			break;
			
        default:
			if(xls_debug)
			{
				//xls_showBOF(&bof1);
				printf("    Not Processed in parseWorkBook():  BOF=0x%4.4X size=%d\n", bof1.id, bof1.size);
			}
            break;
        }
        bof2=bof1;
		once=1;
    }
    while ((!pWB->olestr->eof)&&(bof1.id!=XLS_RECORD_EOF));

cleanup:
    if (buf)
        free(buf);

    return retval;
}


static xls_error_t xls_preparseWorkSheet(xlsWorkSheet* pWS)
{
    if(!pWS) return LIBXLS_ERROR_NULL_ARGUMENT;

    BOF tmp;
    BYTE* buf = NULL;
    xls_error_t retval = LIBXLS_OK;

    verbose ("xls_preparseWorkSheet");

    if (ole2_seek(pWS->workbook->olestr,pWS->filepos) == -1) {
        retval = LIBXLS_ERROR_SEEK;
        goto cleanup;
    }
    do
    {
		size_t read;
		if((read = ole2_read(&tmp, 1, 4, pWS->workbook->olestr)) != 4) {
            if (xls_debug) fprintf(stderr, "Error: failed to read OLE size\n");
            retval = LIBXLS_ERROR_READ;
            goto cleanup;
        }
        xlsConvertBof(&tmp);
        if (tmp.size) {
            if ((buf = realloc(buf, tmp.size)) == NULL) {
                if (xls_debug) fprintf(stderr, "Error: failed to allocate buffer of size %d\n", (int)tmp.size);
                retval = LIBXLS_ERROR_MALLOC;
                goto cleanup;
            }
            if((read = ole2_read(buf, 1, tmp.size, pWS->workbook->olestr)) != tmp.size) {
                if (xls_debug) fprintf(stderr, "Error: failed to read OLE block\n");
                retval = LIBXLS_ERROR_READ;
                goto cleanup;
            }
        }

        switch (tmp.id)
        {
        case XLS_RECORD_DEFCOLWIDTH:
            if (tmp.size < sizeof(WORD)) {
                retval = LIBXLS_ERROR_PARSE;
                goto cleanup;
            }
            pWS->defcolwidth = (buf[0] << 8) + (buf[1] << 16);
            break;
        case XLS_RECORD_COLINFO:
            if (tmp.size < sizeof(COLINFO)) {
                retval = LIBXLS_ERROR_PARSE;
                goto cleanup;
            }
            xlsConvertColinfo((COLINFO*)buf);
            if ((retval = xls_addColinfo(pWS,(COLINFO*)buf)) != LIBXLS_OK)
                goto cleanup;
            break;
        case XLS_RECORD_ROW:
            if (tmp.size < sizeof(ROW)) {
                retval = LIBXLS_ERROR_PARSE;
                goto cleanup;
            }
            xlsConvertRow((ROW*)buf);
            /* The lcell field is 1-indexed whereas lastcol is 0-indexed */
            if (pWS->rows.lastcol+1<((ROW*)buf)->lcell)
                pWS->rows.lastcol=((ROW*)buf)->lcell-1;
            if (pWS->rows.lastrow<((ROW*)buf)->index)
                pWS->rows.lastrow=((ROW*)buf)->index;
            break;
        /* If the ROW record is incorrect or missing, infer the information from
         * cell data. */
        case XLS_RECORD_MULRK:
            if (xls_isCellTooSmall(pWS->workbook, &tmp, buf)) {
                retval = LIBXLS_ERROR_PARSE;
                goto cleanup;
            }
            if (pWS->rows.lastcol<xlsShortVal(((MULRK*)buf)->col) + (tmp.size - 6)/6 - 1)
                pWS->rows.lastcol=xlsShortVal(((MULRK*)buf)->col) + (tmp.size - 6)/6 - 1;
            if (pWS->rows.lastrow<xlsShortVal(((MULRK*)buf)->row))
                pWS->rows.lastrow=xlsShortVal(((MULRK*)buf)->row);
            break;
        case XLS_RECORD_MULBLANK:
            if (xls_isCellTooSmall(pWS->workbook, &tmp, buf)) {
                retval = LIBXLS_ERROR_PARSE;
                goto cleanup;
            }
            if (pWS->rows.lastcol<xlsShortVal(((MULBLANK*)buf)->col) + (tmp.size - 6)/2 - 1)
                pWS->rows.lastcol=xlsShortVal(((MULBLANK*)buf)->col) + (tmp.size - 6)/2 - 1;
            if (pWS->rows.lastrow<xlsShortVal(((MULBLANK*)buf)->row))
                pWS->rows.lastrow=xlsShortVal(((MULBLANK*)buf)->row);
            break;
        case XLS_RECORD_NUMBER:
        case XLS_RECORD_RK:
        case XLS_RECORD_LABELSST:
        case XLS_RECORD_BLANK:
        case XLS_RECORD_LABEL:
        case XLS_RECORD_RSTRING:
        case XLS_RECORD_FORMULA:
        case XLS_RECORD_FORMULA_ALT:
        case XLS_RECORD_BOOLERR:
            if (xls_isCellTooSmall(pWS->workbook, &tmp, buf)) {
                retval = LIBXLS_ERROR_PARSE;
                goto cleanup;
            }
            if (pWS->rows.lastcol<xlsShortVal(((COL*)buf)->col))
                pWS->rows.lastcol=xlsShortVal(((COL*)buf)->col);
            if (pWS->rows.lastrow<xlsShortVal(((COL*)buf)->row))
                pWS->rows.lastrow=xlsShortVal(((COL*)buf)->row);
            break;
        }
        if (pWS->rows.lastcol > 255) {
            retval = LIBXLS_ERROR_PARSE;
            goto cleanup;
        }
    }
    while ((!pWS->workbook->olestr->eof)&&(tmp.id!=XLS_RECORD_EOF));

cleanup:
    if (buf)
        free(buf);
    return retval;
}

static xls_error_t xls_formatColumn(xlsWorkSheet* pWS)
{
    DWORD i,t,ii;
    DWORD fcol,lcol;
    WORD width;
    BYTE isHidden;

    for (i=0;i<pWS->colinfo.count;i++)
    {
        width = pWS->colinfo.col[i].width;
        isHidden = (pWS->colinfo.col[i].flags&1);
        if (pWS->colinfo.col[i].first<=pWS->rows.lastcol)
            fcol=pWS->colinfo.col[i].first;
        else
            fcol=pWS->rows.lastcol;

        if (pWS->colinfo.col[i].last<=pWS->rows.lastcol)
            lcol=pWS->colinfo.col[i].last;
        else
            lcol=pWS->rows.lastcol;

        for (ii=0;ii<=pWS->rows.lastrow;ii++) {
            for (t=fcol;t<=lcol;t++) {
                pWS->rows.row[ii].cells.cell[t].isHidden |= isHidden;
                pWS->rows.row[ii].cells.cell[t].width = width;
            }
        }
    }
    return LIBXLS_OK;
}

xls_error_t xls_parseWorkSheet(xlsWorkSheet* pWS)
{
    if(!pWS) return LIBXLS_ERROR_NULL_ARGUMENT;

    BOF tmp;
    BYTE* buf = NULL;
	long offset = pWS->filepos;
    size_t read;
    xls_error_t retval = 0;

	struct st_cell_data *cell = NULL;
	xlsWorkBook *pWB = pWS->workbook;

    verbose ("xls_parseWorkSheet");

    if ((retval = xls_preparseWorkSheet(pWS)) != LIBXLS_OK) {
        goto cleanup;
    }
	// printf("size=%d fatpos=%d)\n", pWS->workbook->olestr->size, pWS->workbook->olestr->fatpos);

    if ((retval = xls_makeTable(pWS)) != LIBXLS_OK) {
        goto cleanup;
    }

    if ((retval = xls_formatColumn(pWS)) != LIBXLS_OK) {
        goto cleanup;
    }

    if (ole2_seek(pWS->workbook->olestr,pWS->filepos) == -1) {
        retval = LIBXLS_ERROR_SEEK;
        goto cleanup;
    }
    do
    {
		long lastPos = offset;

		if(xls_debug > 10) {
			printf("LASTPOS=%ld pos=%d filePos=%d filePos=%d\n", lastPos, (int)pWB->olestr->pos, pWS->filepos, pWB->filepos);
		}
		if((read = ole2_read(&tmp, 1, 4, pWS->workbook->olestr)) != 4) {
            if (xls_debug) fprintf(stderr, "Error: failed to read OLE size\n");
            retval = LIBXLS_ERROR_READ;
            goto cleanup;
        }
        xlsConvertBof((BOF *)&tmp);
        if (tmp.size) {
            if ((buf = realloc(buf, tmp.size)) == NULL) {
                if (xls_debug) fprintf(stderr, "Error: failed to allocate buffer of size %d\n", (int)tmp.size);
                retval = LIBXLS_ERROR_MALLOC;
                goto cleanup;
            }
            if((read = ole2_read(buf, 1, tmp.size, pWS->workbook->olestr)) != tmp.size) {
                if (xls_debug) fprintf(stderr, "Error: failed to read OLE block\n");
                retval = LIBXLS_ERROR_READ;
                goto cleanup;
            }
        }
		offset += 4 + tmp.size;

		if(xls_debug)
			xls_showBOF(&tmp);

        switch (tmp.id)
        {
        case XLS_RECORD_EOF:
            break;
        case XLS_RECORD_MERGEDCELLS:
            if ((retval = xls_mergedCells(pWS,&tmp,buf)) != LIBXLS_OK) {
                goto cleanup;
            }
            break;
        case XLS_RECORD_ROW:
            if (tmp.size < sizeof(ROW)) {
                retval = LIBXLS_ERROR_PARSE;
                goto cleanup;
            }
			if(xls_debug > 10) printf("ROW: %x at pos=%ld\n", tmp.id, lastPos);
            xlsConvertRow((ROW *)buf);
            if ((retval = xls_addRow(pWS,(ROW*)buf)) != LIBXLS_OK) {
                goto cleanup;
            }
            break;
		case XLS_RECORD_DEFCOLWIDTH:
            if (tmp.size < sizeof(WORD)) {
                retval = LIBXLS_ERROR_PARSE;
                goto cleanup;
            }
			if(xls_debug > 10) printf("DEFAULT COL WIDTH: %d\n", ((WORD *)buf)[0]);
			break;
		case XLS_RECORD_DEFAULTROWHEIGHT:
            if (tmp.size < 2 * sizeof(WORD)) {
                retval = LIBXLS_ERROR_PARSE;
                goto cleanup;
            }
			if(xls_debug > 10) printf("DEFAULT ROW Height: 0x%x %d\n", ((WORD *)buf)[0], ((WORD *)buf)[1]);
			break;
		case XLS_RECORD_DBCELL:
			if(xls_debug > 10) {
				DWORD *foo = (DWORD *)buf;
                WORD *goo;
				int i;
                printf("DBCELL: size %d\n", tmp.size);
				printf("DBCELL OFFSET=%4.4u -> ROW %ld\n", foo[0], lastPos-foo[0]);
				++foo;
				goo = (WORD *)foo;
				for(i=0; i<5; ++i) printf("goo[%d]=%4.4x %u\n", i, goo[i], goo[i]);
			}
			break;
        case XLS_RECORD_INDEX:
			if(xls_debug > 10) {
				DWORD *foo = (DWORD *)buf;
                int i;
				printf("INDEX: size %d\n", tmp.size);
				for(i=0; i<5; ++i) printf("FOO[%d]=%4.4x %u\n", i, foo[i], foo[i]);
			}
#if 0
			0	4 4	4 8	4
			12	4 16	4∙nm
			Not used Index to first used row (rf, 0-based) Index to first row of unused tail of sheet (rl, last used row + 1, 0-based)
			Absolute stream position of the DEFCOLWIDTH record (➜5.32) of the current sheet. If this record does not exist, the offset points to the record at the position where the DEFCOLWIDTH record would occur.
			Array of nm absolute stream positions to the DBCELL record (➜5.29) of each Row Block
#endif
            break;
        case XLS_RECORD_MULRK:
        case XLS_RECORD_MULBLANK:
        case XLS_RECORD_NUMBER:
        case XLS_RECORD_BOOLERR:
        case XLS_RECORD_RK:
        case XLS_RECORD_LABELSST:
        case XLS_RECORD_BLANK:
        case XLS_RECORD_LABEL:
        case XLS_RECORD_RSTRING:
        case XLS_RECORD_FORMULA:
        case XLS_RECORD_FORMULA_ALT:
            if ((cell = xls_addCell(pWS, &tmp, buf)) == NULL) {
                retval = LIBXLS_ERROR_PARSE;
                goto cleanup;
            }
            break;
		case XLS_RECORD_ARRAY:
			if(formula_handler) formula_handler(tmp.id, tmp.size, buf);
			break;

		case XLS_RECORD_STRING:
			if(cell && (cell->id == XLS_RECORD_FORMULA || cell->id == XLS_RECORD_FORMULA_ALT)) {
                xls_cell_set_str(cell, get_string((char *)buf, tmp.size,
                            (BYTE)!pWB->is5ver, pWB));
				if (xls_debug) xls_showCell(cell);
			}
			break;

        default:
			if(xls_debug)
			{
				//xls_showBOF(&tmp);
                if (tmp.size >= sizeof(COL)) {
                    printf("   [%d:%d]: 0x%X at pos=%lu size=%u\n", xlsShortVal(((COL*)buf)->row), xlsShortVal(((COL*)buf)->col),
                            tmp.id, lastPos, tmp.size);
                } else {
                    printf("   0x%X at pos=%lu size=%u\n", tmp.id, lastPos, tmp.size);
                }
			}
            break;
        }
    }
    while ((!pWS->workbook->olestr->eof)&&(tmp.id!=XLS_RECORD_EOF));

cleanup:
    if (buf)
        free(buf);

    return retval;
}

xlsWorkSheet * xls_getWorkSheet(xlsWorkBook* pWB,int num)
{
    xlsWorkSheet * pWS = NULL;
    verbose ("xls_getWorkSheet");
    if (num >= 0 && num < (int)pWB->sheets.count) {
        pWS = calloc(1, sizeof(xlsWorkSheet));
        pWS->filepos=pWB->sheets.sheet[num].filepos;
        pWS->workbook=pWB;
        pWS->rows.lastcol=0;
        pWS->rows.lastrow=0;
        pWS->colinfo.count=0;
    }
    return pWS;
}

static xlsWorkBook *xls_open_ole(OLE2 *ole, const char *charset, xls_error_t *outError) {
    xlsWorkBook* pWB;
    xls_error_t retval = LIBXLS_OK;

    pWB = calloc(1, sizeof(xlsWorkBook));
    verbose ("xls_open_ole");

    if ((pWB->olestr=ole2_fopen(ole, "\005SummaryInformation")))
    {
        pWB->summary = calloc(1,4096);
		if (ole2_read(pWB->summary, 4096, 1, pWB->olestr) == -1) {
            if (xls_debug) fprintf(stderr, "SummaryInformation not found\n");
            retval = LIBXLS_ERROR_READ;
            goto cleanup;
        }
		ole2_fclose(pWB->olestr);
	}

    if ((pWB->olestr=ole2_fopen(ole, "\005DocumentSummaryInformation")))
    {
        pWB->docSummary = calloc(1, 4096);
		if (ole2_read(pWB->docSummary, 4096, 1, pWB->olestr) == -1) {
            if (xls_debug) fprintf(stderr, "DocumentSummaryInformation not found\n");
            retval = LIBXLS_ERROR_READ;
            goto cleanup;
        }
		ole2_fclose(pWB->olestr);
	}

#if 0
	if(xls_debug) {
		printf("summary=%d docsummary=%d\n", pWB->summary ? 1 : 0, pWB->docSummary ? 1 : 0);
		xlsSummaryInfo *si = xls_summaryInfo(pWB);
		printf("title=%s\n", si->title);
		printf("subject=%s\n", si->subject);
		printf("author=%s\n", si->author);
		printf("keywords=%s\n", si->keywords);
		printf("comment=%s\n", si->comment);
		printf("lastAuthor=%s\n", si->lastAuthor);
		printf("appName=%s\n", si->appName);
		printf("category=%s\n", si->category);
		printf("manager=%s\n", si->manager);
		printf("company=%s\n", si->company);
	}
#endif

    // open Workbook
    if (!(pWB->olestr=ole2_fopen(ole,"Workbook")) && !(pWB->olestr=ole2_fopen(ole,"Book")))
    {
        if(xls_debug) fprintf(stderr, "Workbook not found\n");
        retval = LIBXLS_ERROR_PARSE;
        goto cleanup;
    }

    pWB->sheets.count=0;
    pWB->xfs.count=0;
    pWB->fonts.count=0;
    pWB->supbooks.count=0;
    pWB->externsheets.count=0;
    pWB->definednames.count=0;
    pWB->charset = strdup(charset ? charset : "UTF-8");

    retval = xls_parseWorkBook(pWB);

cleanup:
    if (retval != LIBXLS_OK) {
        if (!pWB->olestr)
            ole2_close(ole);
        xls_close_WB(pWB);
        pWB = NULL;
    }
    if (outError)
        *outError = retval;

    return pWB;
}

xlsWorkBook* xls_open(const char *file, const char* charset)
{
    return xls_open_file(file, charset, NULL);
}

xlsWorkBook* xls_open_file(const char *file, const char* charset, xls_error_t *outError) {
    OLE2* ole = NULL;

    if (!(ole=ole2_open_file(file)))
    {
        if (xls_debug) fprintf(stderr, "File \"%s\" not found\n",file);
        if (outError) *outError = LIBXLS_ERROR_OPEN;
        return NULL;
    }

    return xls_open_ole(ole, charset, outError);
}

xlsWorkBook *xls_open_buffer(const unsigned char *buffer, size_t len,
        const char *charset, xls_error_t *outError) {
    OLE2* ole = NULL;

    if (!(ole=ole2_open_buffer(buffer, len)))
    {
        if (outError) *outError = LIBXLS_ERROR_OPEN;
        return NULL;
    }

    return xls_open_ole(ole, charset, outError);
}

xlsRow *xls_row(xlsWorkSheet* pWS, WORD cellRow)
{
    if(cellRow > pWS->rows.lastrow)
        return NULL;

    if (pWS->rows.row == NULL)
        return NULL;

    return &pWS->rows.row[cellRow];
}

xlsCell	*xls_cell(xlsWorkSheet* pWS, WORD cellRow, WORD cellCol)
{
    struct st_row_data	*row;

    if ((row = xls_row(pWS, cellRow)) == NULL)
        return NULL;

    if(cellCol >= row->cells.count)
        return NULL;

    return &row->cells.cell[cellCol];
}

void xls_close_WB(xlsWorkBook* pWB)
{
	OLE2*		ole;

	verbose ("xls_close");

	if(!pWB) return;

    // OLE first
    if (pWB->olestr) {
        ole=pWB->olestr->ole;
        ole2_fclose(pWB->olestr);
        ole2_close(ole);
    }

    // WorkBook
    free(pWB->charset);

    // Sheets
    {
        DWORD i;
        for(i=0; i<pWB->sheets.count; ++i) {
            free(pWB->sheets.sheet[i].name);
        }
        free(pWB->sheets.sheet);
    }

    // SST
    {
        DWORD i;
        for(i=0; i<pWB->sst.count; ++i) {
            free(pWB->sst.string[i].str);
        }
        free(pWB->sst.string);
    }

    // xfs
    {
        free(pWB->xfs.xf);
    }

    // fonts
    {
        DWORD i;
        for(i=0; i<pWB->fonts.count; ++i) {
            free(pWB->fonts.font[i].name);
        }
        free(pWB->fonts.font);
    }

    // formats
    {
        DWORD i;
        for(i=0; i<pWB->formats.count; ++i) {
            free(pWB->formats.format[i].value);
        }
        free(pWB->formats.format);
    }

    // supbooks
    {
      if(pWB->supbooks.count) {
        free(pWB->supbooks.supbook);
      }
    }

    // externsheets
    {
      if(pWB->externsheets.count) {
        free(pWB->externsheets.sheet);
      }
    }

    // definednames
    {
      DWORD i;
      for(i=0; i<pWB->definednames.count; ++i) {
        free(pWB->definednames.name[i].name);
        free(pWB->definednames.name[i].formulabytes);
      }
      free(pWB->definednames.name);
    }

    // buffers
	if(pWB->summary)  free(pWB->summary);
	if(pWB->docSummary) free(pWB->docSummary);

#ifdef HAVE_ICONV
    if (pWB->converter)
        iconv_close((iconv_t)pWB->converter);
    if (pWB->utf16_converter)
        iconv_close((iconv_t)pWB->utf16_converter);
#endif

    if (pWB->utf8_locale)
        xls_freelocale((xls_locale_t)pWB->utf8_locale);

    // TODO - free other dynamically allocated objects like string table??
    free(pWB);
}

void xls_close_WS(xlsWorkSheet* pWS)
{
	if(!pWS) return;

    if (pWS->rows.row) {
        DWORD i, j;
        for(j=0; j<=pWS->rows.lastrow; ++j) {
            struct st_row_data *row = &pWS->rows.row[j];
            for(i=0; i<row->cells.count; ++i) {
                free(row->cells.cell[i].str);
            }
            free(row->cells.cell);
        }
        free(pWS->rows.row);
    }

    // COLINFO
    {
        free(pWS->colinfo.col);
    }
    free(pWS);
}

const char* xls_getVersion(void)
{
    return PACKAGE_VERSION;
}

const char* xls_getError(xls_error_t code) {
    if (code == LIBXLS_OK)
        return "No error";
    if (code == LIBXLS_ERROR_READ)
        return "Unable to read from file";
    if (code == LIBXLS_ERROR_OPEN)
        return "Unable to open file";
    if (code == LIBXLS_ERROR_SEEK)
        return "Unable to seek within file";
    if (code == LIBXLS_ERROR_MALLOC)
        return "Unable to allocate memory";
    if (code == LIBXLS_ERROR_PARSE)
        return "Unable to parse file";
    if (code == LIBXLS_ERROR_UNSUPPORTED_ENCRYPTION)
        return "Unsupported encryption scheme";

    return "Unknown error";
}

//
// http://poi.apache.org/hpsf/internals.html
// or google "DocumentSummaryInformation and UserDefined Property Sets" and look for MSDN hits
//

xlsSummaryInfo *xls_summaryInfo(xlsWorkBook* pWB)
{
	xlsSummaryInfo	*pSI;

	pSI = (xlsSummaryInfo *)calloc(1, sizeof(xlsSummaryInfo));
	xls_dumpSummary(pWB->summary, 1, pSI);
	xls_dumpSummary(pWB->docSummary, 0, pSI);

	return pSI;
}

void xls_close_summaryInfo(xlsSummaryInfo *pSI)
{
	if(!pSI) return;

	if(pSI->title)		free(pSI->title);
	if(pSI->subject)	free(pSI->subject);
	if(pSI->author)		free(pSI->author);
	if(pSI->keywords)	free(pSI->keywords);
	if(pSI->comment)	free(pSI->comment);
	if(pSI->lastAuthor)	free(pSI->lastAuthor);
	if(pSI->appName)	free(pSI->appName);
	if(pSI->category)	free(pSI->category);
	if(pSI->manager)	free(pSI->manager);
	if(pSI->company)	free(pSI->company);

	free(pSI);
}

static void xls_dumpSummary(char *buf,int isSummary,xlsSummaryInfo *pSI) {
	header			*head;
	sectionList		*secList;
	propertyList	*plist;
	sectionHeader	*secHead;
	property		*prop;
	uint32_t i, j;

	if(!buf) return;	// perhaps the document was missing??

	head = (header *)buf;
	//printf("header: \n");
	//printf("  sig=%x\n", head->sig);
	//printf("  os=%x\n", head->os >> 16);
	//printf("  class=%8.8x%8.8x%8.8x%8.8x\n", head->format[0], head->format[1], head->format[2], head->format[3]);
	//printf("  count=%x\n", head->count);

	for(i=0; i<head->count; ++i) {
		secList = &head->secList[i];
		//printf("Section %d:\n", i);
		//printf("  class=%8.8x%8.8x%8.8x%8.8x\n", secList->format[0], secList->format[1], secList->format[2], secList->format[3]);
		//printf("  offset=%d (now at %ld\n", secList->offset, (char *)secList - (char *)buf + sizeof(sectionList));


		secHead = (sectionHeader *)((char *)head + secList->offset);
		//printf("  len=%d\n", secHead->length);
		//printf("  properties=%d\n", secHead->numProperties);
		for(j=0; j<secHead->numProperties; ++j) {
			BYTE **s;

			plist = &secHead->properties[j];
			//printf("      ---------\n");
			//printf("      propID=%d offset=%d\n", plist->propertyID, plist->sectionOffset);
			prop = (property *)((char *)secHead + plist->sectionOffset);
			//printf("      propType=%d\n", prop->propertyID);

			switch(prop->propertyID) {
			case 2:
				//printf("      xlsShortVal=%x\n", *(uint16_t *)prop->data);
				break;
			case 3:
				//printf("      wordVal=%x\n", *(uint32_t *)prop->data);
				break;
			case 30:
				//printf("      longVal=%llx\n", *(uint64_t *)prop->data);
				//printf("      s[%u]=%s\n", *(uint32_t  *)prop->data, (char *)prop->data + 4);
				if(isSummary) {
					switch(plist->propertyID) {
					case 2:		s = &pSI->title;		break;
					case 3:		s = &pSI->subject;		break;
					case 4:		s = &pSI->author;		break;
					case 5:		s = &pSI->keywords;		break;
					case 6:		s = &pSI->comment;		break;
					case 8:		s = &pSI->lastAuthor;	break;
					case 18:	s = &pSI->appName;		break;
					default:	s = NULL;				break;
					}
				} else {
					switch(plist->propertyID) {
					case 2:		s = &pSI->category;		break;
					case 14:	s = &pSI->manager;		break;
					case 15:	s = &pSI->company;		break;
					default:	s = NULL;				break;
					}
				}
				if(s) *s = (BYTE *)strdup((char *)prop->data + 4);
				break;
			case 64:
				//printf("      longVal=%llx\n", *(uint64_t *)prop->data);
				break;
			case 65:
#if 0
				{
				uint32_t k;
				for(k=0; k<*(uint32_t  *)prop->data; ++k) {
				unsigned char *t = (unsigned char *)prop->data + 4 + k;
				printf(" %2.2x(%c)", *t, *t);
				}
				printf("\n");
				}
#endif
				break;
			default:
				//printf("      UNKNOWN!\n");
				break;
			}
		}
	}
}

void xls_set_formula_hander(xls_formula_handler handler)
{
	formula_handler = handler;
}


char *xls_parseFormulaBytes(xlsWorkBook* pWB, BYTE *bytes, size_t len)
{
  int is_area;
  unsigned row1, row2, col1, col2, row;
  int col, is_relative_col, is_relative_row, is_all_rows, is_all_cols;
  char *sheet1, *sheet2;
  int escape;

  char cellref[50];
  int j = 0;
  int CELLREF_LEN = 50;

  size_t buflen = 50;
  char *result = calloc(buflen, sizeof(char));
  result[0] = '\0';

  char *LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

  RGCE_LOCREL *loc_rel;
  RGCE_AREAREL *area_rel;

  struct st_supbook_data *spbk = NULL;
  struct st_externsheet_data *ext = NULL;


  switch(bytes[0]) {
  default:
    goto return_null;

  case 0x3A: // PtgRef3D
  case 0x3B: // PtgArea3D
    is_area = (bytes[0] == 0x3B);
    if(is_area) {
      area_rel = (RGCE_AREAREL *)(bytes + 3);
    } else {
      loc_rel = (RGCE_LOCREL *)(bytes + 3);
    }

    // get related externsheet and supbook records
    uint16_t ext_i = *(uint16_t *)(bytes + 1);
    if(ext_i < pWB->externsheets.count) {
      ext = &(pWB->externsheets.sheet[ext_i]);
      if(ext->index < pWB->supbooks.count) {
        spbk = &(pWB->supbooks.supbook[ext->index]);
      }
    }
    if(!ext || !spbk)
      goto return_null;

    // return NULL if unsupported ref type
    if(spbk->reftype == ref_extern || spbk->reftype == ref_addin)
      goto return_null;

    // worksheet name
    sheet1 = pWB->sheets.sheet[ext->tabfirst].name;
    sheet2 = pWB->sheets.sheet[ext->tablast].name;
    escape = escape_sheetname(sheet1) || escape_sheetname(sheet2);
    if(escape)
      strcat_realloc(result, "'", &buflen);
    strcat_realloc(result, sheet1, &buflen);
    if(sheet2 != sheet1) {
      strcat_realloc(result, ":", &buflen);
      strcat_realloc(result, sheet2, &buflen);
    }
    if(escape)
      strcat_realloc(result, "'", &buflen);
    strcat_realloc(result, "!", &buflen);


    // determine cell reference type
    if(is_area) {
      is_all_rows = area_rel->row_first == 0 && area_rel->row_last == 0xFFFF;
      is_all_cols = area_rel->col_first == 0 && area_rel->col_last == 0xFF;
      row1 = area_rel->row_first;
      row2 = area_rel->row_last;
      col1 = area_rel->col_first;
      col2 = area_rel->col_last;
    } else {
      is_all_rows = 0;
      is_all_cols = 0;
      row1 = loc_rel->row;
      col1 = loc_rel->col;
      row2 = row1;
      col2 = col1;
    }


    // build cell reference
    row = row1;
    col = col1 & 0x3FFF;
    is_relative_col = (col1 & 0x4000);
    is_relative_row = (col1 & 0x8000);

    if(!is_all_cols) {
      if(!is_relative_col)
        cellref[j++] = '$';
      if(col >= 26)
        cellref[j++] = LETTERS[(col / 26)-1];
      cellref[j++] = LETTERS[col % 26];
    }
    if(!is_all_rows || is_all_cols) {
      if(!is_relative_row)
        cellref[j++] = '$';
      snprintf(cellref + j, CELLREF_LEN - j, "%d", row+1);
      j = strlen(cellref);
    }

    if(is_area) {
      cellref[j++] = ':';

      row = row2;
      col = col2 & 0x3FFF;
      is_relative_col = (col2 & 0x4000);
      is_relative_row = (col2 & 0x8000);

      if(!is_all_cols) {
        if(!is_relative_col)
          cellref[j++] = '$';
        if(col >= 26)
          cellref[j++] = LETTERS[(col / 26)-1];
        cellref[j++] = LETTERS[col % 26];
      }
      if(!is_all_rows || is_all_cols) {
        if(!is_relative_row)
          cellref[j++] = '$';
        snprintf(cellref + j, CELLREF_LEN - j, "%d", row+1);
        j = strlen(cellref);
      }
    }

    cellref[j++] = '\0';
    result = strcat_realloc(result, cellref, &buflen);

    break;
  }


  // shrink allocated memory to string size
  result = realloc(result, (strlen(result)+1) * sizeof(char));

  return result;

return_null:
  free(result);
  return NULL;
}


// returns true if worksheet name needs to be escaped with ''
static int escape_sheetname(char *sheetname)
{
  int n = strlen(sheetname);
  if(n == 0)
    return 0;

  char c = sheetname[0];
  if(!((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '_'))
    return 1;

  for(int i = 1; i < n; i++) {
    c = sheetname[i];
    if(!((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '_'))
      return 1;
  }
  return 0;
}
