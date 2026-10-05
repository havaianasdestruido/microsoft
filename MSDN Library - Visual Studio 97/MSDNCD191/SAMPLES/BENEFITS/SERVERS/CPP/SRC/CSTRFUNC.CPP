//cstrfunc.cpp  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Implementation file for CStrFnc

#pragma warning (disable : 4800)	// turns off bool warning

#ifndef _UNICODE
#define _UNICODE
#define TURNOFF
#endif

#include <stdio.h>
#include <tchar.h>
#include <stdarg.h>
#include <math.h>
#include "support.h"
#include "cstrfunc.h"

extern HINSTANCE g_hInst;

#define MIN_DATE                (-657434L)  // about year 100
#define MAX_DATE                2958465L    // about year 9999
#define MAX_TIME_BUFFER_SIZE    128         // matches that in timecore.cpp

// Half a second, expressed in days
#define HALF_SECOND  (1.0/172800.0)

// One-based array of days in year at month start
static int rgMonthDays[13] =
	{0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365};

extern "C"
{
static BOOL TmFromOleDate(DATE dtSrc, struct tm& tmDest);
static void TmConvertToStandardFormat(struct tm& tmSrc);
static BOOL OleDateFromTm(WORD wYear, WORD wMonth, WORD wDay,
					WORD wHour, WORD wMinute, WORD wSecond, DATE& dtDest);
}

CStrFnc::CStrFnc()
{
	m_bstrString = NULL;
}

CStrFnc::CStrFnc(long lSize)
{
	m_bstrString = ::SysAllocStringLen(NULL, lSize);
}

CStrFnc::CStrFnc(BSTR bstrString)
{
	BSTR bstrTemp;

	bstrTemp = ::SysAllocString(bstrString);
	m_bstrString = bstrTemp;
	::SysFreeString(bstrTemp);
}

CStrFnc::~CStrFnc()
{
}

//CStrFnc::LoadResString
//Purpose:Load a string resource for the resource id.
BSTR CStrFnc::LoadResString(UINT iResID)
{
	CHAR	Buffer[2048];
	HRESULT	hResult;
	LPSTR	pBuffer;
	_bstr_t bstrTemp1;
	BSTR	bstrTemp;
	DWORD dwrc;

	pBuffer = Buffer;
	hResult = ::LoadStringA(g_hInst, iResID, pBuffer, 2048);
	if (hResult == 0)
		dwrc = ::GetLastError();

	bstrTemp1 = pBuffer;
	bstrTemp = ::SysAllocString(bstrTemp1);

	return (bstrTemp);
}

void CStrFnc::Empty()
{
	m_bstrString = NULL;
}

BSTR CStrFnc::UCase(BSTR bstrString)
{
	BSTR			bstrTemp;
	unsigned int	uLen;
	unsigned int	n;

	bstrTemp = ::SysAllocString(bstrString);
	uLen = ::SysStringLen(bstrTemp);
	for (n = 0; n < uLen; n++)
		toupper(*((wchar_t *)bstrTemp + n));

	return (bstrTemp);
}

BSTR CStrFnc::LCase(BSTR bstrString)
{
	BSTR			bstrTemp;
	unsigned int	uLen;
	unsigned int	n;

	bstrTemp = ::SysAllocString(bstrString);
	uLen = ::SysStringLen(bstrTemp);
	for (n = 0; n < uLen; n++)
		tolower(*((wchar_t *)bstrTemp + n));

	return (bstrTemp);
}

BSTR CStrFnc::Trim(BSTR bstrString)
{
	return (CStrFnc::LTrim(CStrFnc::RTrim(bstrString)));
}

BSTR CStrFnc::LTrim(BSTR bstrString)
{
	unsigned int	uLen;
	_bstr_t			bstrTemp1;
	_bstr_t			bstrTemp2;
	_bstr_t			Temp;
	BSTR			Temp1;
	BSTR			Temp2;

	bstrTemp1 = bstrString;
	uLen = bstrTemp1.length();

	while (isspace(*((wchar_t *)bstrTemp1)))
	{
		Temp1 = bstrTemp1;
		Temp2 = ::SysAllocString(Temp1 + 1);
		bstrTemp2 = Temp2;
		
		bstrTemp1 = bstrTemp2;
		::SysFreeString(bstrTemp2);
		if (uLen == 0)
			break;
	}

	Temp = bstrTemp1;
	return (Temp);
}

BSTR CStrFnc::RTrim(BSTR bstrString)
{
	_bstr_t			bstrTemp1;
	_bstr_t			bstrTemp2;
	_bstr_t			Temp;
	unsigned int	uLen;

	bstrTemp1 = bstrString;
	uLen = ::SysStringLen(bstrTemp1);

	if (uLen)
	{
		while (isspace(*((wchar_t *)bstrTemp1 + uLen - 1)))
		{
			bstrTemp2 = ::SysAllocStringLen(bstrTemp1, --uLen);
			bstrTemp1 = bstrTemp2;
			::SysFreeString(bstrTemp2);
			if (uLen == 0)
				break;
		}
	}

	Temp = bstrTemp1;
	return (Temp);
}

BSTR CStrFnc::Left(BSTR String, int x)
{
	_bstr_t	bstrTemp1;
	int		iLen = ::SysStringLen(String);

	if (x == -1 || x > iLen)
		return (String);
	else
	{
		BSTR bstrTemp;
		bstrTemp = ::SysAllocStringLen(String, x);
		_bstr_t Temp = bstrTemp;
		::SysFreeString(bstrTemp);
		return (Temp);
	}
}

BSTR CStrFnc::Right(BSTR String, int x)
{
	BSTR	bstrTemp1;
	int		iLen;
	_bstr_t	Temp;

	iLen = ::SysStringLen(String);
	if (x == -1 || x > iLen)
		Temp = String;
	else
	{
		bstrTemp1 = ::SysAllocStringLen(String, x);
		Temp = bstrTemp1;
		::SysFreeString(bstrTemp1);
	}
	return (Temp);
}

BSTR CStrFnc::UCase()
{
	_bstr_t Temp;

	Temp = m_bstrString;
	this->UCase(Temp);
	return (Temp);
}

BSTR CStrFnc::LCase()
{
	_bstr_t Temp;

	Temp = m_bstrString;
	this->LCase(Temp);
	return (Temp);
}

BSTR CStrFnc::Trim()
{
	_bstr_t Temp;

	Temp = m_bstrString;
	this->Trim(Temp);
	return (Temp);
}

BSTR CStrFnc::LTrim()
{
	_bstr_t Temp;

	Temp = m_bstrString;
	this->LTrim(Temp);
	return (Temp);
}

BSTR CStrFnc::RTrim()
{
	_bstr_t Temp;

	Temp = m_bstrString;
	this->RTrim(Temp);
	return (Temp);
}

BSTR CStrFnc::Left(int x)
{
	_bstr_t Temp;
	
	Temp = m_bstrString;
	this->Left(Temp, x);
	return (Temp);
}

BSTR CStrFnc::Right(int x)
{
	_bstr_t Temp;

	Temp = m_bstrString;
	this->Right(Temp, x);
	return (Temp);
}

BSTR CStrFnc::GetBstr()
{
	return (m_bstrString);
}

int CStrFnc::Length()
{
	return (_tcslen(m_bstrString));
}

BSTR CStrFnc::Copy()
{
	return (::SysAllocString(m_bstrString));
}

void CStrFnc::Init()
{
	if (!bInit)
		bInit = TRUE;
	return;
}

void CStrFnc::Format(BSTR Format, ...)
{
	va_list argList;
	BSTR bstrResult;

	va_start(argList, Format);
	bstrResult = vaFormat(Format, argList);
	m_bstrString = ::SysAllocString(bstrResult);
}

BSTR CStrFnc::vaFormat(BSTR bstrFormat, va_list argList)
{
	BSTR bstrTemp;

	bstrTemp = ::SysAllocStringLen(NULL, 1024);
	vswprintf(bstrTemp, bstrFormat, argList);
	_bstr_t Temp(bstrTemp);
	::SysFreeString(bstrTemp);
	return (Temp);
}

CStrFnc CStrFnc::operator=(BSTR rhs)
{
	m_bstrString = rhs;
	return (*this);
}

long CStrFnc::Val()
{
	return (this->Val(m_bstrString));
}

long CStrFnc::Val(BSTR bstrString)
{
	BSTR bstr;
	VARIANT v1;
	VARIANT v2;

	::VariantInit(&v1);
	::VariantInit(&v2);
	
	bstr = ::SysAllocString(bstrString);

	v1.vt = VT_BSTR;
	V_BSTR(&v1) = bstr;
	::VariantChangeType(&v2, &v1, 0, VT_I4);

	::SysFreeString(bstr);

	return (V_I4(&v2));
}

BOOL CStrFnc::IsValidDate( DATE dateIn )
{
	BSTR		bstr;
	ULONG		dwFlags	= 0;
	HRESULT		hResult = NOERROR;
	LCID		lcid	= CP_ACP;

	if (dateIn == 0.0)
		return (FALSE);

	hResult = ::VarBstrFromDate(dateIn, lcid, dwFlags, &bstr);
	return (hResult == NOERROR);
}

BOOL CStrFnc::IsValidDate(BSTR DateValue)
{
	OLECHAR*	pStr;
	LCID		lcid	= CP_ACP;
	ULONG		dwFlags = 0;
	HRESULT		hResult = NOERROR;
	DATE		dateOut;
	_bstr_t		bstrIn( DateValue, FALSE );

	hResult = ::VarDateFromStr( (wchar_t *)pStr, lcid, dwFlags, &dateOut );
	return (hResult == NOERROR);
}

DATE CStrFnc::FormatDate(BSTR DateIn)
{
	DATE		dateOut	= 0;
	LCID		lcid	= CP_ACP;
	ULONG		dwFlags	= 0;
	WORD		wYear;
	WORD		wMonth;
	WORD		wDay;
	TCHAR		sChar[4];
	_bstr_t		sVal;

	sChar[0] = *((wchar_t *)DateIn + 0);
	sChar[1] = *((wchar_t *)DateIn + 1);
	sChar[2] = 0;
	sVal = sChar;
	CStrFnc sMonth= sVal;
	wMonth = (WORD) sMonth.Val();

	sChar[0] = *((wchar_t *)DateIn + 2);
	sChar[1] = *((wchar_t *)DateIn + 3);
	sChar[2] = 0;
	sVal = sChar;
	CStrFnc sDay = sVal;
	wDay = (WORD) sDay.Val();

	sChar[0] = *((wchar_t *)DateIn + 4);
	sChar[1] = *((wchar_t *)DateIn + 5);
	sChar[2] = *((wchar_t *)DateIn + 6);
	sChar[3] = *((wchar_t *)DateIn + 7);
	sChar[4] = 0;
	sVal = sChar;
	CStrFnc sYear = sVal;
	wYear = (WORD) sYear.Val();

	// convert date, use Jan 1, 1899 as date if failure
	if (!OleDateFromTm(wYear, wMonth, wDay, 0, 0, 0, dateOut))
			OleDateFromTm(WORD(1899), WORD(1), WORD(1), 0, 0, 0, dateOut);

	return (dateOut);
}

BSTR CStrFnc::FormatDate(DATE DateIn, BSTR strFormat)
{
	BSTR		bstrDate;
	_bstr_t		fmtTemp(strFormat);
	struct tm	tmTemp;

	if (fmtTemp.length() <= 0)
		return (bstrDate);

	if (!TmFromOleDate(DateIn, tmTemp))
		return (bstrDate);

	// Convert tm from internal format to standard format
	TmConvertToStandardFormat(tmTemp);

	TCHAR szTemp[MAX_TIME_BUFFER_SIZE];
	if (!IsValidDate(DateIn))
	{
		DATE dTemp;
		OleDateFromTm(WORD(0), WORD(1), WORD(1), 0, 0, 0, dTemp);
		TmFromOleDate(dTemp, tmTemp);
	}

	if (_tcsftime(szTemp, MAX_TIME_BUFFER_SIZE, strFormat, &tmTemp) > 0)
		bstrDate = szTemp;
	
	return (bstrDate);
}


int CStrFnc::BSTRCmpi(BSTR bstr1, BSTR bstr2)
{
	if (bstr1 == NULL && bstr2 == NULL)
		return (0);
	else
	if (bstr1 == NULL && ::SysStringLen(bstr2) == 0)
		return (0);
	else
	if (::SysStringLen(bstr1) == 0 && bstr2 == NULL)
		return (0);
	else
		return (wcsicmp(bstr1, bstr2));
}


BOOL CStrFnc::IsEmpty(BSTR bstrValue)
{
	if (bstrValue == NULL)
		return (TRUE);
	if (::SysStringLen(bstrValue) == 0)
		return (TRUE);
	return (FALSE);
}


BOOL CStrFnc::Boolean(BSTR bstrString)
{
	CStrFnc Temp(bstrString);
	return (Temp.Boolean());
}


BOOL CStrFnc::Boolean()
{
	BOOL brc = FALSE;
	if (!IsEmpty(m_bstrString))
	{
		if (BSTRCmpi(m_bstrString, TEXT("YES")))
			brc = TRUE;
		else
		if (BSTRCmpi(m_bstrString, TEXT("true")))
			brc = TRUE;
		else
		if (BSTRCmpi(m_bstrString, TEXT("ON")))
			brc = TRUE;
		else
		if (BSTRCmpi(m_bstrString, TEXT("1")))
			brc = TRUE;
		else
		if (BSTRCmpi(m_bstrString, TEXT("-1")))
			brc = TRUE;
	}

	return (brc);
}

CURRENCY CStrFnc::BstrToCy(BSTR bstrString)
{
	CStrFnc Temp(bstrString);
	return (Temp.BstrToCy());
}

CURRENCY CStrFnc::BstrToCy()
{
	VARIANT v1;
	VARIANT v2;

	::VariantInit(&v1);
	::VariantInit(&v2);

	v1.vt = VT_BSTR;
	V_BSTR(&v1) = m_bstrString;

	::VariantChangeType(&v2, &v1, 0, VT_CY);

	return (V_CY(&v2));
}

long CStrFnc::VariantBstrToLong(VARIANT vValue)
{
	VARIANT v2;

	::VariantInit(&v2);
	::VariantChangeType(&v2, &vValue, 0, VT_I4);
	return (V_I4(&v2));
}

void TmConvertToStandardFormat(struct tm& tmSrc)
{
	// Convert internal tm to format expected by runtimes (_tcsftime, etc)
	tmSrc.tm_year -= 1900;  // year is based on 1900
	tmSrc.tm_mon -= 1;      // month of year is 0-based
	tmSrc.tm_wday -= 1;     // day of week is 0-based
	tmSrc.tm_yday -= 1;     // day of year is 0-based
}

BOOL TmFromOleDate(DATE dtSrc, struct tm& tmDest)
{
	// The legal range does not actually span year 0 to 9999.
	if (dtSrc > MAX_DATE || dtSrc < MIN_DATE) // about year 100 to about 9999
		return (FALSE);

	long nDays;             // Number of days since Dec. 30, 1899
	long nDaysAbsolute;     // Number of days since 1/1/0
	long nSecsInDay;        // Time in seconds since midnight
	long nMinutesInDay;     // Minutes in day

	long n400Years;         // Number of 400 year increments since 1/1/0
	long n400Century;       // Century within 400 year block (0,1,2 or 3)
	long n4Years;           // Number of 4 year increments since 1/1/0
	long n4Day;             // Day within 4 year block
							//  (0 is 1/1/yr1, 1460 is 12/31/yr4)
	long n4Yr;              // Year within 4 year block (0,1,2 or 3)
	BOOL bLeap4 = TRUE;     // TRUE if 4 year block includes leap year

	double dblDate = dtSrc; // tempory serial date

	// If a valid date, then this conversion should not overflow
	nDays = (long)dblDate;

	// Round to the second
	dblDate += ((dtSrc > 0.0) ? HALF_SECOND : -HALF_SECOND);

	nDaysAbsolute = (long)dblDate + 693959L; // Add days from 1/1/0 to 12/30/1899

	dblDate = fabs(dblDate);
	nSecsInDay = (long)((dblDate - floor(dblDate)) * 86400.);

	// Calculate the day of week (sun=1, mon=2...)
	//   -1 because 1/1/0 is Sat.  +1 because we want 1-based
	tmDest.tm_wday = (int)((nDaysAbsolute - 1) % 7L) + 1;

	// Leap years every 4 yrs except centuries not multiples of 400.
	n400Years = (long)(nDaysAbsolute / 146097L);

	// Set nDaysAbsolute to day within 400-year block
	nDaysAbsolute %= 146097L;

	// -1 because first century has extra day
	n400Century = (long)((nDaysAbsolute - 1) / 36524L);

	// Non-leap century
	if (n400Century != 0)
	{
		// Set nDaysAbsolute to day within century
		nDaysAbsolute = (nDaysAbsolute - 1) % 36524L;

		// +1 because 1st 4 year increment has 1460 days
		n4Years = (long)((nDaysAbsolute + 1) / 1461L);

		if (n4Years != 0)
			n4Day = (long)((nDaysAbsolute + 1) % 1461L);
		else
		{
			bLeap4 = FALSE;
			n4Day = (long)nDaysAbsolute;
		}
	}
	else
	{
		// Leap century - not special case!
		n4Years = (long)(nDaysAbsolute / 1461L);
		n4Day = (long)(nDaysAbsolute % 1461L);
	}

	if (bLeap4)
	{
		// -1 because first year has 366 days
		n4Yr = (n4Day - 1) / 365;

		if (n4Yr != 0)
			n4Day = (n4Day - 1) % 365;
	}
	else
	{
		n4Yr = n4Day / 365;
		n4Day %= 365;
	}

	// n4Day is now 0-based day of year. Save 1-based day of year, year number
	tmDest.tm_yday = (int)n4Day + 1;
	tmDest.tm_year = n400Years * 400 + n400Century * 100 + n4Years * 4 + n4Yr;

	// Handle leap year: before, on, and after Feb. 29.
	if (n4Yr == 0 && bLeap4)
	{
		// Leap Year
		if (n4Day == 59)
		{
			// Feb. 29
			tmDest.tm_mon = 2;
			tmDest.tm_mday = 29;
			goto DoTime;
		}

		// Pretend it's not a leap year for month/day comp.
		if (n4Day >= 60)
			--n4Day;
	}

	// Make n4DaY a 1-based day of non-leap year and compute
	//  month/day for everything but Feb. 29.
	++n4Day;

	// Month number always >= n/32, so save some loop time */
	for (tmDest.tm_mon = (n4Day >> 5) + 1;
		n4Day > rgMonthDays[tmDest.tm_mon]; tmDest.tm_mon++);

	tmDest.tm_mday = (int)(n4Day - rgMonthDays[tmDest.tm_mon-1]);

DoTime:
	if (nSecsInDay == 0)
		tmDest.tm_hour = tmDest.tm_min = tmDest.tm_sec = 0;
	else
	{
		tmDest.tm_sec = (int)nSecsInDay % 60L;
		nMinutesInDay = nSecsInDay / 60L;
		tmDest.tm_min = (int)nMinutesInDay % 60;
		tmDest.tm_hour = (int)nMinutesInDay / 60;
	}

	return (TRUE);
}

BOOL OleDateFromTm(WORD wYear, WORD wMonth, WORD wDay,
	WORD wHour, WORD wMinute, WORD wSecond, DATE & dtDest)
{
	// Validate year and month (ignore day of week and milliseconds)
	if (wYear > 9999 || wMonth < 1 || wMonth > 12)
		return (FALSE);

	//  Check for leap year and set the number of days in the month
	BOOL bLeapYear = ((wYear & 3) == 0) &&
		((wYear % 100) != 0 || (wYear % 400) == 0);

	int nDaysInMonth =
		rgMonthDays[wMonth] - rgMonthDays[wMonth-1] +
		((bLeapYear && wDay == 29 && wMonth == 2) ? 1 : 0);

	// Finish validating the date
	if (wDay < 1 || wDay > nDaysInMonth || 	wHour > 23 || wMinute > 59 || wSecond > 59)
		return (FALSE);

	// Cache the date in days and time in fractional days
	long nDate;
	double dblTime;

	//It is a valid date; make Jan 1, 1AD be 1
	nDate = wYear*365L + wYear/4 - wYear/100 + wYear/400 +
		rgMonthDays[wMonth-1] + wDay;

	//  If leap year and it's before March, subtract 1:
	if (wMonth <= 2 && bLeapYear)
		--nDate;

	//  Offset so that 12/30/1899 is 0
	nDate -= 693959L;

	dblTime = (((long)wHour * 3600L) +  // hrs in seconds
		((long)wMinute * 60L) +  // mins in seconds
		((long)wSecond)) / 86400.;

	dtDest = (double) nDate + ((nDate >= 0) ? dblTime : -dblTime);

	return (TRUE);
}

#ifndef TURNOFF
#undef _UNICODE
#undef TURNOFF
#endif
