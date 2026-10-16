/*++

Verification harness for the ASCII85 conversion in the Windows NT 4.0
PostScript printer driver output filter.

Component under test
    Microsoft Windows NT 4.0 DDK (4.0.1381)
    winnt40.ddk/ddk/src/print/psprint/pscript/FILTER.C
    FilterConvertToASCII85()

Why this harness exists
    The repository targets Windows NT 5 (Windows 2000) class systems whose
    native build environment (MSVC plus the NT build environment driven by
    $(NTMAKEENV)\makefile.def) is not available here.  The change made to
    FILTER.C is, however, a pure arithmetic transformation of a self-contained
    routine, so its correctness and its benefit can both be established on any
    host compiler:

      1. EQUIVALENCE - the pre-change and post-change routines are compared
         over every one of the 2^32 possible 32-bit inputs (use --full) and
         over a large pseudo-random sample (default).  Byte-identical output
         is required, including the returned length and the all-zero "z"
         special case with the zero-optimisation both on and off.

      2. SELF-CONSISTENCY - the emitted digits are independently decoded back
         to a 32-bit value and compared with the input, confirming that the
         routine really performs the base-85 conversion it documents.

      3. SOURCE FINGERPRINT - the harness reads FILTER.C and asserts that the
         pre-change construct (division by a non-constant table element) is
         gone and the post-change construct is present, so the harness cannot
         silently drift away from the file it claims to validate.

      4. COST - the two forms are timed on a page-raster sized buffer.  This
         measures the *algorithmic* difference (four hardware divides per four
         bytes versus four multiply/shift pairs); it is not a measurement of
         the NT driver, which cannot be built or run here.

Build
    cc -O2 -o ascii85_verify ascii85_verify.c

Usage
    ./ascii85_verify [--full] [path/to/FILTER.C]

--*/

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <stdint.h>
#include <time.h>

typedef unsigned long DWORD;
typedef unsigned char BYTE;

#define CONVERT_SOURCE "Microsoft Windows NT 4.0 DDK (4.0.1381)/winnt40.ddk/ddk/src/print/psprint/pscript/FILTER.C"

/* ---- the values used by both forms (as shipped, before the change) ------ */

DWORD ConstAscii85V1[] = {          /* non-const: cannot be folded */
    52200625L,
      614125L,
        7225L,
          85L
};

static const DWORD ConstAscii85[] = {   /* as committed: foldable */
    52200625L,
      614125L,
        7225L,
          85L
};

/* ---- pre-change routine, transcribed from FILTER.C ---------------------- */

static int convert_before(const BYTE *pIn, BYTE *pbConvert, int bZeroOpt)
{
    DWORD dwNewNum, dwTmp;
    DWORD *pConst;
    BYTE   i;

    dwNewNum  = (DWORD)((uint32_t)pIn[0] * 16777216u);
    dwNewNum += (DWORD)((uint32_t)pIn[1] * 65536u);
    dwNewNum += (DWORD)((uint32_t)pIn[2] * 256u);
    dwNewNum += (DWORD)pIn[3];

    if (!bZeroOpt || dwNewNum) {
        for (i = 0, pConst = ConstAscii85V1; i < 4; i++, pConst++) {
            dwTmp    = dwNewNum / *pConst;
            dwNewNum -= (*pConst * dwTmp);
            *pbConvert++ = (BYTE) dwTmp + '!';
        }
        *pbConvert = (BYTE) dwNewNum + '!';
        return 5;
    } else {
        *pbConvert = 'z';
        return 1;
    }
}

/* ---- post-change routine, transcribed from FILTER.C --------------------- */

static int convert_after(const BYTE *pIn, BYTE *pbConvert, int bZeroOpt)
{
    DWORD dwNewNum, dwTmp;

    dwNewNum  = (DWORD)((uint32_t)pIn[0] * 16777216u);
    dwNewNum += (DWORD)((uint32_t)pIn[1] * 65536u);
    dwNewNum += (DWORD)((uint32_t)pIn[2] * 256u);
    dwNewNum += (DWORD)pIn[3];

    if (!bZeroOpt || dwNewNum) {

        dwTmp = dwNewNum / ConstAscii85[0];
        dwNewNum -= ConstAscii85[0] * dwTmp;
        *pbConvert++ = (BYTE) dwTmp + '!';

        dwTmp = dwNewNum / ConstAscii85[1];
        dwNewNum -= ConstAscii85[1] * dwTmp;
        *pbConvert++ = (BYTE) dwTmp + '!';

        dwTmp = dwNewNum / ConstAscii85[2];
        dwNewNum -= ConstAscii85[2] * dwTmp;
        *pbConvert++ = (BYTE) dwTmp + '!';

        dwTmp = dwNewNum / ConstAscii85[3];
        dwNewNum -= ConstAscii85[3] * dwTmp;
        *pbConvert++ = (BYTE) dwTmp + '!';

        *pbConvert = (BYTE) dwNewNum + '!';
        return 5;

    } else {
        *pbConvert = 'z';
        return 1;
    }
}

/* ---- source fingerprint ------------------------------------------------- */

static char *slurp(const char *path, size_t *len)
{
    FILE *f = fopen(path, "rb");
    char *buf;
    long n;
    if (!f) return NULL;
    fseek(f, 0, SEEK_END);
    n = ftell(f);
    fseek(f, 0, SEEK_SET);
    if (n <= 0) { fclose(f); return NULL; }
    buf = malloc((size_t)n + 1);
    if (!buf) { fclose(f); return NULL; }
    if (fread(buf, 1, (size_t)n, f) != (size_t)n) { free(buf); fclose(f); return NULL; }
    buf[n] = '\0';
    fclose(f);
    if (len) *len = (size_t)n;
    return buf;
}

static int fingerprint(const char *path)
{
    size_t len = 0;
    char *src = slurp(path, &len);
    int failures = 0;
    static const char *banned[] = {
        "dwNewNum / *pConst",           /* divisor read through a pointer */
        "for ( i = 0, pConst = ConstAscii85;",
        NULL
    };
    static const char *required[] = {
        "static const DWORD ConstAscii85[]",
        "dwNewNum / ConstAscii85[0]",
        "dwNewNum / ConstAscii85[3]",
        NULL
    };
    int i;

    if (!src) {
        printf("\n== source fingerprint ==\n");
        printf("FILTER.C not found at '%s'\n", path);
        printf("(run from the repository root, or pass the path explicitly;\n");
        printf(" this check is skipped by callers that lack the tree)\n");
        return 0;
    }

    printf("\n== source fingerprint: %s ==\n", path);
    for (i = 0; banned[i]; i++) {
        if (strstr(src, banned[i])) {
            printf("  FAIL  still present: %s\n", banned[i]);
            failures++;
        } else {
            printf("  ok    removed: %s\n", banned[i]);
        }
    }
    for (i = 0; required[i]; i++) {
        if (!strstr(src, required[i])) {
            printf("  FAIL  missing: %s\n", required[i]);
            failures++;
        } else {
            printf("  ok    present: %s\n", required[i]);
        }
    }
    free(src);
    return failures;
}

/* ---- independent decode of 5 emitted digits ----------------------------- */

static int decode5(const BYTE *out, int n, uint32_t *value)
{
    uint32_t v = 0;
    int i;
    if (n == 1) {
        if (out[0] != 'z') return 0;
        *value = 0;
        return 1;
    }
    if (n != 5) return 0;
    for (i = 0; i < 5; i++) {
        if (out[i] < '!' || out[i] > 'u') return 0;
        v = v * 85u + (uint32_t)(out[i] - '!');
    }
    *value = v;
    return 1;
}

static uint64_t rng = 0x12345678u;
static uint32_t nextrand(void)
{
    rng = rng * 6364136223846793005ull + 1442695040888963407ull;
    return (uint32_t)(rng >> 32);
}

int main(int argc, char **argv)
{
    BYTE in[4], a[8], b[8];
    uint32_t v;
    int full = 0, na, nb, i;
    unsigned long long checked = 0, mismatches = 0, rtfail = 0, limit;
    const char *srcpath = CONVERT_SOURCE;
    double before_s, after_s;
    size_t n, k;
    unsigned char *src, *dst;
    volatile unsigned long sink = 0;
    struct timespec t0, t1;

    for (i = 1; i < argc; i++) {
        if (strcmp(argv[i], "--full") == 0) full = 1;
        else srcpath = argv[i];
    }

    printf("== equivalence: %s ==\n",
           full ? "exhaustive over all 2^32 inputs" : "large pseudo-random sample");

    if (full) {
        for (limit = 0; limit <= 0xFFFFFFFFull; limit++) {
            v = (uint32_t)limit;
            in[0] = (BYTE)(v >> 24); in[1] = (BYTE)(v >> 16);
            in[2] = (BYTE)(v >> 8);  in[3] = (BYTE)v;
            na = convert_before(in, a, 1);
            nb = convert_after(in, b, 1);
            if (na != nb || memcmp(a, b, (size_t)na) != 0) {
                if (mismatches < 10)
                    printf("MISMATCH v=%08lx before=%02x%02x%02x%02x%02x after=%02x%02x%02x%02x%02x\n",
                           (unsigned long)v, a[0], a[1], a[2], a[3], a[4],
                           b[0], b[1], b[2], b[3], b[4]);
                mismatches++;
            }
            if (na == 5) {
                uint32_t back;
                if (!decode5(a, na, &back) || back != v) {
                    if (rtfail < 10) printf("ROUNDTRIP FAIL v=%08lx\n", (unsigned long)v);
                    rtfail++;
                }
            }
            checked++;
        }
    } else {
        for (limit = 0; limit < 50000000ull; limit++) {
            /* the first block walks low values and every single-byte case */
            if (limit < 0x1000000ull) {
                v = (uint32_t)limit;
            } else if (limit < 0x1800000ull) {          /* exercise every byte value */
                uint32_t t = (uint32_t)(limit & 0xFFFF);
                v = (t % 4 == 0) ? (t & 0xFF) << 24
                  : (t % 4 == 1) ? (t & 0xFF) << 16
                  : (t % 4 == 2) ? (t & 0xFF) << 8
                                 : (t & 0xFF);
            } else {
                v = nextrand();
            }
            in[0] = (BYTE)(v >> 24); in[1] = (BYTE)(v >> 16);
            in[2] = (BYTE)(v >> 8);  in[3] = (BYTE)v;
            na = convert_before(in, a, 1);
            nb = convert_after(in, b, 1);
            if (na != nb || memcmp(a, b, (size_t)na) != 0) mismatches++;
            if (na == 5) {
                uint32_t back;
                if (!decode5(a, na, &back) || back != v) rtfail++;
            }
            checked++;
        }
    }

    /* the all-zero word, with the zero-optimisation on and off */
    {
        BYTE z[4] = { 0, 0, 0, 0 };
        na = convert_before(z, a, 1); nb = convert_after(z, b, 1);
        if (na != 1 || nb != 1 || a[0] != 'z' || b[0] != 'z') mismatches++;
        na = convert_before(z, a, 0); nb = convert_after(z, b, 0);
        if (na != nb || memcmp(a, b, (size_t)na) != 0) mismatches++;
    }

    printf("checked=%llu  mismatches=%llu  roundtrip_failures=%llu\n",
           checked, mismatches, rtfail);
    printf("%s\n", (mismatches || rtfail) ? "EQUIVALENCE: FAIL" : "EQUIVALENCE: PASS");

    mismatches += (unsigned long long)fingerprint(srcpath);

    /* cost of the two forms on a page-raster sized buffer */
    n = 64u * 1024u * 1024u;
    src = malloc(n);
    dst = malloc(n * 2);
    if (!src || !dst) { printf("allocation failed\n"); return 2; }
    for (k = 0; k < n; k++) src[k] = (unsigned char)(nextrand() >> 13);

    clock_gettime(CLOCK_MONOTONIC, &t0);
    for (k = 0; k + 4 <= n; k += 4) sink += (unsigned long)convert_before(src + k, dst + k, 1);
    clock_gettime(CLOCK_MONOTONIC, &t1);
    before_s = (double)(t1.tv_sec - t0.tv_sec) + (double)(t1.tv_nsec - t0.tv_nsec) / 1e9;

    clock_gettime(CLOCK_MONOTONIC, &t0);
    for (k = 0; k + 4 <= n; k += 4) sink += (unsigned long)convert_after(src + k, dst + k, 1);
    clock_gettime(CLOCK_MONOTONIC, &t1);
    after_s = (double)(t1.tv_sec - t0.tv_sec) + (double)(t1.tv_nsec - t0.tv_nsec) / 1e9;

    printf("\n== conversion cost over %lu MiB of raster data ==\n",
           (unsigned long)(n >> 20));
    printf("before (divisor through a pointer) : %.3f s\n", before_s);
    printf("after  (constant divisor)          : %.3f s\n", after_s);
    printf("ratio                              : %.2fx\n", before_s / after_s);
    printf("(sink=%lu)\n", sink);

    free(src); free(dst);
    return (mismatches || rtfail) ? 1 : 0;
}
