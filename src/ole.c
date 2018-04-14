/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 *
 * This file is part of libxls -- A multiplatform, C/C++ library
 * for parsing Excel(TM) files.
 *
 * Redistribution and use in source and binary forms, with or without modification, are
 * permitted provided that the following conditions are met:
 *
 *    1. Redistributions of source code must retain the above copyright notice, this list of
 *       conditions and the following disclaimer.
 *
 *    2. Redistributions in binary form must reproduce the above copyright notice, this list
 *       of conditions and the following disclaimer in the documentation and/or other materials
 *       provided with the distribution.
 *
 * THIS SOFTWARE IS PROVIDED BY David Hoerl ''AS IS'' AND ANY EXPRESS OR IMPLIED
 * WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND
 * FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL David Hoerl OR
 * CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
 * CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
 * SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
 * ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
 * NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 *
 * Copyright 2004 Komarov Valery
 * Copyright 2006 Christophe Leitienne
 * Copyright 2013 Bob Colbert
 * Copyright 2008-2013 David Hoerl
 *
 */

#include "config.h" 

#include <memory.h>
#include <math.h>
#include <string.h>
#include <stdio.h>
#include <stdlib.h>

#include "libxls/ole.h"
#include "libxls/xlstool.h"
#include "libxls/endian.h"
/* Mask illegal functions for CMD check */
#include "cran.h"

extern int xls_debug;

//#define OLE_DEBUG

//static const DWORD MSATSECT		= 0xFFFFFFFC;	// -4
//static const DWORD FATSECT		= 0xFFFFFFFD;	// -3
static const DWORD ENDOFCHAIN	= 0xFFFFFFFE;	// -2
static const DWORD FREESECT		= 0xFFFFFFFF;	// -1

static size_t sector_pos(OLE2* ole2, size_t sid);
static ssize_t sector_read(OLE2* ole2, void *buffer, size_t sid);
static ssize_t read_MSAT(OLE2* ole2, OLE2Header *oleh);
static void *ole_malloc(size_t len);

static void *ole_malloc(size_t len) {
    if (len > (1<<24) || len == 0) {
        return NULL;
    }
    return malloc(len);
}

// Read next sector of stream
int ole2_bufread(OLE2Stream* olest) 
{
	BYTE *ptr;

    if (olest == NULL || olest->ole == NULL)
        return -1;

    if ((DWORD)olest->fatpos!=ENDOFCHAIN)
    {
		if(olest->sfat) {
            if (olest->ole->SSAT == NULL || olest->buf == NULL || olest->ole->SSecID == NULL)
                return -1;

            if (olest->fatpos*olest->ole->lssector + olest->bufsize > olest->ole->SSATCount) {
                if (xls_debug) fprintf3(stderr, "Error: fatpos %d out-of-bounds for SSAT\n", (int)olest->fatpos);
                return -1;
            }

			ptr = olest->ole->SSAT + olest->fatpos*olest->ole->lssector;
			memcpy(olest->buf, ptr, olest->bufsize); 

            if (olest->fatpos >= olest->ole->SSecIDCount) {
                if (xls_debug) fprintf4(stderr, "Error: fatpos %d out-of-bounds for SSecID[%d]\n",
                        (int)olest->fatpos, olest->ole->SSecIDCount);
                return -1;
            }

			olest->fatpos=xlsIntVal(olest->ole->SSecID[olest->fatpos]);
			olest->pos=0;
			olest->cfat++;
		} else {
			if ((int)olest->fatpos < 0 ||
                sector_read(olest->ole, olest->buf, olest->fatpos) == -1) {
                if (xls_debug) fprintf3(stderr, "Error: Unable to read sector #%d\n", (int)olest->fatpos);
                return -1;
            }

            if (olest->fatpos >= olest->ole->SecIDCount) {
                if (xls_debug) fprintf4(stderr, "Error: fatpos %d out-of-bounds for SecID[%d]\n",
                        (int)olest->fatpos, olest->ole->SecIDCount);
                return -1;
            }

			olest->fatpos=xlsIntVal(olest->ole->SecID[olest->fatpos]);
			olest->pos=0;
			olest->cfat++;
		}
    }
	// else printf("ENDOFCHAIN!!!\n");
    return 0;
}

// Read part of stream
ssize_t ole2_read(void* buf, size_t size, size_t count, OLE2Stream* olest)
{
    size_t didReadCount=0;
    size_t totalReadCount;
	size_t needToReadCount;

	totalReadCount=size*count;

	// olest->size inited to -1
	// printf("===== ole2_read(%ld bytes)\n", totalReadCount);

    if ((long)olest->size>=0 && !olest->sfat)	// directory is -1
    {
		size_t rem;
		rem = olest->size - (olest->cfat*olest->ole->lsector+olest->pos);		
        totalReadCount = rem<totalReadCount?rem:totalReadCount;
        if (rem<=0) olest->eof=1;

		// printf("  rem=%ld olest->size=%d - subfunc=%d\n", rem, olest->size, (olest->cfat*olest->ole->lsector+olest->pos) );
		//printf("  totalReadCount=%d (rem=%d size*count=%ld)\n", totalReadCount, rem, size*count);
	}

	while ((!olest->eof) && (didReadCount < totalReadCount))
	{
		unsigned long remainingBytes;

		needToReadCount	= totalReadCount - didReadCount;
		remainingBytes	= olest->bufsize - olest->pos;
		//printf("  test: (totalReadCount-didReadCount)=%d (olest->bufsize-olest->pos)=%d\n", (totalReadCount-didReadCount), (olest->bufsize-olest->pos) );

		if (needToReadCount < remainingBytes)	// does the current sector contain all the data I need?
		{
			// printf("  had %d bytes of memory, copy=%d\n", (olest->bufsize-olest->pos), needToReadCount);
			memcpy((BYTE*)buf + didReadCount, olest->buf + olest->pos, needToReadCount);
			olest->pos		+= needToReadCount;
			didReadCount	+= needToReadCount;
		} else {
			// printf("  had %d bytes of memory, copy=%d\n", remainingBytes, remainingBytes);
			memcpy((BYTE*)buf + didReadCount, olest->buf + olest->pos, remainingBytes);
			olest->pos		+= remainingBytes;
			didReadCount	+= remainingBytes;
			if (ole2_bufread(olest) == -1)
                return -1;
		}
		//printf("  if(fatpos=0x%X==EOC=0x%X) && (pos=%d >= bufsize=%d)\n", olest->fatpos, ENDOFCHAIN, olest->pos, olest->bufsize);
		if (((DWORD)olest->fatpos == ENDOFCHAIN) && (olest->pos >= olest->bufsize))
		{
			olest->eof=1;
		}

		//printf("  eof=%d (didReadCount=%ld != totalReadCount=%ld)\n", olest->eof, didReadCount, totalReadCount);
	}
    if (didReadCount > totalReadCount)
        return -1;

	// printf("  didReadCount=%ld EOF=%d\n", didReadCount, olest->eof);
	// printf("=====\n");

#ifdef OLE_DEBUG
    fprintf2(stderr, "----------------------------------------------\n");
    fprintf2(stderr, "ole2_read (end)\n");
    fprintf3(stderr, "start:		%d \n",olest->start);
    fprintf3(stderr, "pos:		%d \n",(int)olest->pos);
    fprintf3(stderr, "cfat:		%d \n",(int)olest->cfat);
    fprintf3(stderr, "size:		%d \n",(int)olest->size);
    fprintf3(stderr, "fatpos:		%d \n",(int)olest->fatpos);
    fprintf3(stderr, "bufsize:		%d \n",(int)olest->bufsize);
    fprintf3(stderr, "eof:		%d \n",olest->eof);
#endif

    return didReadCount;
}

// Open stream in logical ole file
OLE2Stream* ole2_sopen(OLE2* ole,DWORD start, size_t size)
{
    OLE2Stream* olest=NULL;

#ifdef OLE_DEBUG
    fprintf2(stderr, "----------------------------------------------\n");
    fprintf3(stderr, "ole2_sopen start=%Xh\n", start);
#endif

	olest=(OLE2Stream*)calloc(1, sizeof(OLE2Stream));
	olest->ole=ole;
	olest->size=size;
	olest->fatpos=start;
	olest->start=start;
	olest->pos=0;
	olest->eof=0;
	olest->cfat=-1;
	if((long)size > 0 && size < (size_t)ole->sectorcutoff) {
		olest->bufsize=ole->lssector;
		olest->sfat = 1;
	} else {
		olest->bufsize=ole->lsector;
	}
	olest->buf = ole_malloc(olest->bufsize);
	ole2_bufread(olest);

	// if(xls_debug) printf("sopen: sector=%d next=%d\n", start, olest->fatpos);
    return olest;
}

// Move in stream
int ole2_seek(OLE2Stream* olest,DWORD ofs)
{
#ifdef OLE_DEBUG
    fprintf3(stderr, "SEEK %x\n", ofs);
#endif
	if(olest->sfat) {
		ldiv_t div_rez=ldiv(ofs,olest->ole->lssector);
		int i;
		olest->fatpos=olest->start;

		if (div_rez.quot!=0)
		{
			for (i=0;i<div_rez.quot;i++) {
                if (olest->fatpos >= olest->ole->SSecIDCount)
                    return -1;
				olest->fatpos=xlsIntVal(olest->ole->SSecID[olest->fatpos]);
            }
		}

		ole2_bufread(olest);
		olest->pos=div_rez.rem;
		olest->eof=0;
		olest->cfat=div_rez.quot;
		//printf("%i=%i %i\n",ofs,div_rez.quot,div_rez.rem);
	} else {
		ldiv_t div_rez=ldiv(ofs,olest->ole->lsector);
		int i;
#ifdef OLE_DEBUG
        fprintf4(stderr, "seeking fatpos%lu start %u\n", olest->fatpos, olest->start);
#endif
		olest->fatpos=olest->start;

		if (div_rez.quot!=0)
		{
			for (i=0;i<div_rez.quot;i++) {
                if (olest->fatpos >= olest->ole->SecIDCount)
                    return -1;
                olest->fatpos=xlsIntVal(olest->ole->SecID[olest->fatpos]);
            }
		}

		ole2_bufread(olest);
		olest->pos=div_rez.rem;
		olest->eof=0;
		olest->cfat=div_rez.quot;
		//printf("%i=%i %i\n",ofs,div_rez.quot,div_rez.rem);
	}
    return 0;
}

// Open logical file contained in physical OLE file
OLE2Stream*  ole2_fopen(OLE2* ole, const char *file)
{
    int i;

#ifdef OLE_DEBUG
    fprintf2(stderr, "----------------------------------------------\n");
    fprintf3(stderr, "ole2_fopen %s\n", file);
#endif

    for (i=0;i<ole->files.count;i++) {
		char *str = ole->files.file[i].name;
#ifdef OLE_DEBUG
		fprintf2(stderr, "----------------------------------------------\n");
		fprintf3(stderr, "ole2_fopen found %s\n", str);
#endif
        if (str && strcmp(str,file)==0)	// newer versions of Excel don't write the "Root Entry" string for the first set of data
        {
            return ole2_sopen(ole,ole->files.file[i].start,ole->files.file[i].size);
        }
	}
    return NULL;
}

int ole2_fseek(OLE2 *ole2, size_t pos) {
    if (ole2->file)
        return fseek(ole2->file, pos, SEEK_SET);

    if (pos > ole2->buffer_len)
        return -1;

    ole2->buffer_pos = pos;
    return 0;
}

size_t ole2_fread(OLE2 *ole2, void *buffer, size_t size, size_t nitems) {
    if (ole2->file)
        return fread(buffer, size, nitems, ole2->file);

    size_t i = 0;
    for (i=0; i<nitems; i++) {
        if (ole2->buffer_pos + size > ole2->buffer_len)
            break;

        memcpy(buffer, (const char *)ole2->buffer + ole2->buffer_pos, size);
        ole2->buffer_pos += size;
    }
    return i;
}


// read header and check magic numbers
static ssize_t ole2_read_header(OLE2 *ole) {
    ssize_t bytes_read = 0, total_bytes_read = 0;
    OLE2Header *oleh = malloc(512);
    if (ole2_fread(ole, oleh, 512, 1) != 1) {
        total_bytes_read = -1;
        goto cleanup;
    }
    total_bytes_read += 512;
    xlsConvertHeader(oleh);

	// make sure the file looks good. Note: this code only works on Little Endian machines
	if(oleh->id[0] != 0xE011CFD0 || oleh->id[1] != 0xE11AB1A1 || oleh->byteorder != 0xFFFE) {
        if (xls_debug) fprintf2(stderr, "Not an excel file\n");
        total_bytes_read = -1;
        goto cleanup;
	}

    //ole->lsector=(WORD)pow(2,oleh->lsector);
    //ole->lssector=(WORD)pow(2,oleh->lssector);
	ole->lsector=512;
    ole->lssector=64;

	if (oleh->lsectorB != 9 || oleh->lssectorB != 6) {	// 2**9 == 512, 2**6 == 64
        if (xls_debug) fprintf2(stderr, "Unexpected sector size\n");
        total_bytes_read = -1;
        goto cleanup;
    }
	
    ole->cfat=oleh->cfat;
    ole->dirstart=oleh->dirstart;
    ole->sectorcutoff=oleh->sectorcutoff;
    ole->sfatstart=oleh->sfatstart;
    ole->csfat=oleh->csfat;
    ole->difstart=oleh->difstart;
    ole->cdif=oleh->cdif;
    ole->files.count=0;

#ifdef OLE_DEBUG
		fprintf2(stderr, "==== OLE HEADER ====\n");
		//printf ("Header Size:   %i \n", sizeof(OLE2Header));
		//printf ("id[0]-id[1]:   %X-%X \n", oleh->id[0], oleh->id[1]);
		fprintf3(stderr, "verminor:      %X \n",oleh->verminor);
		fprintf3(stderr, "verdll:        %X \n",oleh->verdll);
		//printf ("Byte order:    %X \n",oleh->byteorder);
		fprintf4(stderr, "sect len:      %X (%i)\n",ole->lsector,ole->lsector);		// ole
		fprintf4(stderr, "mini len:      %X (%i)\n",ole->lssector,ole->lssector);	// ole
		fprintf3(stderr, "Fat sect.:     %i \n",oleh->cfat);
		fprintf3(stderr, "Dir Start:     %i \n",oleh->dirstart);
		
		fprintf3(stderr, "Mini Cutoff:   %i \n",oleh->sectorcutoff);
		fprintf3(stderr, "MiniFat Start: %X \n",oleh->sfatstart);
		fprintf3(stderr, "Count MFat:    %i \n",oleh->csfat);
		fprintf3(stderr, "Dif start:     %X \n",oleh->difstart);
		fprintf3(stderr, "Count Dif:     %i \n",oleh->cdif);
		fprintf4(stderr, "Fat Size:      %u (0x%X) \n",oleh->cfat*ole->lsector,oleh->cfat*ole->lsector);
#endif
    // read directory entries
    if ((bytes_read = read_MSAT(ole, oleh)) == -1) {
        total_bytes_read = -1;
        goto cleanup;
    }
    total_bytes_read += bytes_read;

cleanup:
    free(oleh);

    return total_bytes_read;
}

static ssize_t ole2_read_body(OLE2 *ole) {
	// reuse this buffer
    PSS *pss = malloc(512);
    OLE2Stream *olest=ole2_sopen(ole,ole->dirstart, -1);
    char* name = NULL;
    ssize_t bytes_read = 0, total_bytes_read = 0;

    do {
        if ((bytes_read = ole2_read(pss,1,sizeof(PSS),olest)) == -1) {
            total_bytes_read = -1;
            goto cleanup;
        }
        total_bytes_read += bytes_read;
        xlsConvertPss(pss);
        if (pss->bsize > sizeof(pss->name)) {
            total_bytes_read = -1;
            goto cleanup;
        }
        name=unicode_decode(pss->name, pss->bsize, 0, "UTF-8");
#ifdef OLE_DEBUG	
		fprintf4(stderr, "OLE NAME: %s count=%d\n", name, (int)ole->files.count);
#endif
        if (pss->type == PS_USER_ROOT || pss->type == PS_USER_STREAM) // (name!=NULL) // 
        {

#ifdef OLE_DEBUG		
			fprintf3(stderr, "OLE TYPE: %s file=%d \n", pss->type == PS_USER_ROOT ? "root" : "user", (int)ole->files.count);
#endif		
            ole->files.file = realloc(ole->files.file,(ole->files.count+1)*sizeof(struct st_olefiles_data));
            ole->files.file[ole->files.count].name=name;
            ole->files.file[ole->files.count].start=pss->sstart;
            ole->files.file[ole->files.count].size=pss->size;
            ole->files.count++;
			
			if(pss->sstart == ENDOFCHAIN) {
				if (xls_debug) verbose("END OF CHAIN\n");
			} else if(pss->type == PS_USER_STREAM) {
#ifdef OLE_DEBUG
					fprintf2(stderr, "----------------------------------------------\n");
					fprintf5(stderr, "name: %s (size=%d [c=%c])\n", name, pss->bsize, name ? name[0]:' ');
					fprintf3(stderr, "bsize %i\n",pss->bsize);
					fprintf3(stderr, "type %i\n",pss->type);
					fprintf3(stderr, "flag %i\n",pss->flag);
					fprintf3(stderr, "left %X\n",pss->left);
					fprintf3(stderr, "right %X\n",pss->right);
					fprintf3(stderr, "child %X\n",pss->child);
					fprintf10(stderr, "guid %.4X-%.4X-%.4X-%.4X %.4X-%.4X-%.4X-%.4X\n",
                            pss->guid[0],pss->guid[1],pss->guid[2],pss->guid[3],
						pss->guid[4],pss->guid[5],pss->guid[6],pss->guid[7]);
					fprintf3(stderr, "user flag %.4X\n",pss->userflags);
					fprintf3(stderr, "sstart %.4d\n",pss->sstart);
					fprintf3(stderr, "size %.4d\n",pss->size);
#endif
			} else if(pss->type == PS_USER_ROOT) {
				DWORD sector, k, blocks;
				BYTE *wptr;
				
				blocks = (pss->size + (ole->lsector - 1)) / ole->lsector;	// count partial
				if ((ole->SSAT = ole_malloc(blocks*ole->lsector)) == NULL) {
                    total_bytes_read = -1;
                    goto cleanup;
                }
                ole->SSATCount = blocks*ole->lsector;
				// printf("blocks %d\n", blocks);

				sector = pss->sstart;
				wptr = (BYTE*)ole->SSAT;
				for(k=0; k<blocks; ++k) {
					// printf("block %d sector %d\n", k, sector);
                    if (sector == ENDOFCHAIN || sector_read(ole, wptr, sector) == -1) {
                        if (xls_debug) fprintf3(stderr, "Unable to read sector #%d\n", sector);
                        total_bytes_read = -1;
                        goto cleanup;
                    }
                    total_bytes_read += ole->lsector;
					wptr += ole->lsector;
					sector = xlsIntVal(ole->SecID[sector]);
				}
			}	
		} else {
			free(name);
		}
    }
    while (!olest->eof);

cleanup:
	ole2_fclose(olest);
    free(pss);

    return total_bytes_read;
}

// Open in-memory buffer
OLE2 *ole2_open_buffer(const void *buffer, size_t len) {
    OLE2 *ole=(OLE2*)calloc(1, sizeof(OLE2));

    ole->buffer = buffer;
    ole->buffer_len = len;

    if (ole2_read_header(ole) == -1) {
        free(ole);
        return NULL;
    }

    if (ole2_read_body(ole) == -1) {
        free(ole);
        return NULL;
    }

    return ole;
}

// Open physical file
OLE2* ole2_open_file(const char *file)
{
    OLE2* ole = NULL;

#ifdef OLE_DEBUG
    fprintf2(stderr, "----------------------------------------------\n");
    fprintf3(stderr, "ole2_open_file %s\n", file);
#endif

	if(xls_debug) printf("ole2_open: %s\n", file);
    ole=(OLE2*)calloc(1, sizeof(OLE2));

    if (!(ole->file=fopen(file, "rb"))) {
        if(xls_debug) fprintf2(stderr, "File not found\n");
        free(ole);
        return NULL;
    }

    if (ole2_read_header(ole) == -1) {
		fclose(ole->file);
        free(ole);
        return NULL;
    }

    if (ole2_read_body(ole) == -1) {
		fclose(ole->file);
        free(ole);
        return NULL;
    }

    return ole;
}

void ole2_close(OLE2* ole2)
{
    int i;
    if (ole2->file)
        fclose(ole2->file);

	for(i=0; i<ole2->files.count; ++i) {
		free(ole2->files.file[i].name);
	}
	free(ole2->files.file);
	free(ole2->SecID);
	free(ole2->SSecID);
	free(ole2->SSAT);
	free(ole2);
}

void ole2_fclose(OLE2Stream* ole2st)
{
	free(ole2st->buf);
	free(ole2st);
}

// Return offset in bytes of a sector from its sid
static size_t sector_pos(OLE2* ole2, size_t sid)
{
    return 512 + sid * ole2->lsector;
}
// Read one sector from its sid
static ssize_t sector_read(OLE2* ole2, void *buffer, size_t sid)
{
	size_t num;
	size_t seeked;

	//printf("sector_read: sid=%zu (0x%zx) lsector=%u sector_pos=%zu\n", sid, sid, ole2->lsector, sector_pos(ole2, sid) );
    seeked = ole2_fseek(ole2, sector_pos(ole2, sid));
	if(seeked != 0) {
		if (xls_debug) fprintf5(stderr, "Error: wanted to seek to sector %zu (0x%zx) loc=%zu\n", sid, sid, sector_pos(ole2, sid));
        return -1;
    }

	num = ole2_fread(ole2, buffer, ole2->lsector, 1);
    if(num != 1) {
        if (xls_debug) fprintf4(stderr, "Error: fread wanted 1 got %zu loc=%zu\n", num, sector_pos(ole2, sid));
        return -1;
    }

    return ole2->lsector;
}

// read first 109 sectors of MSAT from header
static ssize_t read_MSAT_header(OLE2* ole2, OLE2Header* oleh, int sectorCount) {
    BYTE *sector = (BYTE*)ole2->SecID;
    ssize_t bytes_read = 0, total_bytes_read = 0;
    int sectorNum;

    for (sectorNum = 0; sectorNum < sectorCount; sectorNum++)
    {
        if ((bytes_read = sector_read(ole2, sector, oleh->MSAT[sectorNum])) == -1) {
            if (xls_debug) fprintf3(stderr, "Error: Unable to read sector #%d\n", oleh->MSAT[sectorNum]);
            return -1;
        }
        sector += ole2->lsector;
        total_bytes_read += bytes_read;
    }
    return total_bytes_read;
}

// Add additional sectors of the MSAT
static ssize_t read_MSAT_body(OLE2 *ole2, int sectorOffset, int sectorCount) {
    DWORD sid = ole2->difstart;
    ssize_t bytes_read = 0, total_bytes_read = 0;
    int sectorNum = sectorOffset;

    BYTE *sector = ole_malloc(ole2->lsector);
    //printf("sid=%u (0x%x) sector=%u\n", sid, sid, ole2->lsector);
    while (sid != ENDOFCHAIN && sid != FREESECT) // FREESECT only here due to an actual file that requires it (old Apple Numbers bug)
    {
        int posInSector;
        // read MSAT sector
        if ((bytes_read = sector_read(ole2, sector, sid)) == -1) {
            total_bytes_read = -1;
            if (xls_debug) fprintf3(stderr, "Error: Unable to read sector #%d\n", sid);
            goto cleanup;
        }
        total_bytes_read += bytes_read;

        // read content
        for (posInSector = 0; posInSector < (ole2->lsector-4)/4; posInSector++)
        {
            DWORD s = *(DWORD_UA *)(sector + posInSector*4);
            //printf("   s[%d]=%d (0x%x)\n", posInSector, s, s);

            if (s != ENDOFCHAIN && s != FREESECT) // see patch in Bug 31. For very large files
            {
                if (sectorNum == sectorCount) {
                    if (xls_debug) fprintf3(stderr, "Error: Unable to seek to sector #%d\n", s);
                    total_bytes_read = -1;
                    goto cleanup;
                }
                if ((bytes_read = sector_read(ole2, (BYTE*)(ole2->SecID)+sectorNum*ole2->lsector, s)) == -1) {
                    if (xls_debug) fprintf3(stderr, "Error: Unable to read sector #%d\n", s);
                    total_bytes_read = -1;
                    goto cleanup;
                }
                total_bytes_read += bytes_read;
                sectorNum++;
            }
        }
        sid = *(DWORD_UA *)(sector + posInSector*4);
        //printf("   s[%d]=%d (0x%x)\n", posInSector, sid, sid);
    }
#ifdef OLE_DEBUG
    if(xls_debug) {
        //printf("==== READ IN SECTORS FOR MSAT TABLE====\n");
        int i;
        for(i=0; i<512/4; ++i) {	// just the first block
            if(ole2->SecID[i] != FREESECT) printf("SecID[%d]=%d\n", i, ole2->SecID[i]);
        }
    }
    //exit(0);
#endif

cleanup:
    free(sector);
    return total_bytes_read;
}

// read in short table
static ssize_t read_MSAT_trailer(OLE2 *ole2) {
    ssize_t total_bytes_read = 0;
    DWORD sector, k;
    BYTE *wptr;

	if(ole2->sfatstart != ENDOFCHAIN) {
		if ((ole2->SSecID = ole_malloc(ole2->csfat*(size_t)ole2->lsector)) == NULL) {
            return -1;
        }
        ole2->SSecIDCount = ole2->csfat*(size_t)ole2->lsector/4;
		sector = ole2->sfatstart;
		wptr=(BYTE*)ole2->SSecID;
		for(k=0; k<ole2->csfat; ++k) {
			if (sector == ENDOFCHAIN || sector_read(ole2, wptr, sector) == -1) {
                total_bytes_read = -1;
                goto cleanup;
            }
			wptr += ole2->lsector;
            total_bytes_read += ole2->lsector;
			sector = ole2->SecID[sector];
		}
#ifdef OLE_DEBUG
		if(xls_debug) {
			int i;
			for(i=0; i<ole2->csfat; ++i) {
				if(ole2->SSecID[i] != FREESECT) fprintf4(stderr, "SSecID[%d]=%d\n", i, ole2->SSecID[i]);
			}
		}
#endif
	}

cleanup:
    return total_bytes_read;
}


// Read MSAT
static ssize_t read_MSAT(OLE2* ole2, OLE2Header* oleh)
{
    // reconstitution of the MSAT
    int count = (ole2->cfat < 109) ? ole2->cfat : 109;
    if(count <= 0) {
        if (xls_debug) fprintf2(stderr, "Error: MSAT count out-of-bounds\n");
        return -1;
    }

    ssize_t total_bytes_read = 0;
    ssize_t bytes_read = 0;

    ole2->SecID = ole_malloc(count*ole2->lsector);
    ole2->SecIDCount = count*ole2->lsector/4;

    if ((bytes_read = read_MSAT_header(ole2, oleh, count)) == -1) {
        total_bytes_read = -1;
        goto cleanup;
    }
    total_bytes_read += bytes_read;

    if ((bytes_read = read_MSAT_body(ole2, total_bytes_read / ole2->lsector, count)) == -1) {
        total_bytes_read = -1;
        goto cleanup;
    }
    total_bytes_read += bytes_read;

    if ((bytes_read = read_MSAT_trailer(ole2)) == -1) {
        total_bytes_read = -1;
        goto cleanup;
    }
    total_bytes_read += bytes_read;

cleanup:
    if (total_bytes_read == -1) {
        if (ole2->SecID) {
            free(ole2->SecID);
            ole2->SecID = NULL;
        }
        if (ole2->SSecID) {
            free(ole2->SSecID);
            ole2->SSecID = NULL;
        }
    }

    return total_bytes_read;
}
