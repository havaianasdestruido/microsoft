//csaferay.cpp  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Implementation file for the CSafeArray object.

#pragma warning (disable : 4800)	// turns off bool warning

#include "csaferay.h"

CSafeArray::CSafeArray()
{
	m_paIndices = NULL;
	m_psa = NULL;
	m_psaBound = NULL;
	m_Initialized = FALSE;
}

CSafeArray::CSafeArray(SAFEARRAY * psa)
{
	int i;
	HRESULT hResult;

	hResult = NOERROR;
	m_Initialized = FALSE;

	hResult = ::SafeArrayCopy(psa, &m_psa);
	if ( hResult == NOERROR)
	{
		m_iDims = ::SafeArrayGetDim(psa);
		m_paIndices = new long [m_iDims];
		m_psaBound = new SAFEARRAYBOUND [m_iDims];
		for (i = 0; i < m_iDims; i++)
		{
			::SafeArrayGetLBound(m_psa, i + 1, &m_psaBound[i].lLbound);
			::SafeArrayGetUBound(m_psa, i + 1, (long *)&m_psaBound[i].cElements);
			m_psaBound[i].cElements++;
			m_paIndices[i] = m_psaBound[i].cElements - m_psaBound[i].lLbound + 1;
		}
		m_Initialized = TRUE;
	}
}


CSafeArray::CSafeArray(VARTYPE vtArrayType, int iDims, long * plMaxElements)
{
	Create(vtArrayType, iDims, plMaxElements);
}

//CSafeArray::IsInitialized
//Purpose: Returns the initialization state of the CSafeArray.

CSafeArray::IsInitialized()
{
	return (m_Initialized);
}


CSafeArray::~CSafeArray()
{
	Destroy();
}

//CSafeArray::Create
//Purpose:
//Creates a "iDims" dimensioned CSafeArray object with
//each dimension containing at most "lMaxElements" elements.
//Note: this method must be called first before the object
//is used.

BOOL CSafeArray::Create(VARTYPE vtArrayType, int iDims, long * plMaxElements)
{
	int i;

	m_Initialized = TRUE;
	m_iDims = iDims;

	m_psaBound = new SAFEARRAYBOUND [iDims];
	for (i = 0; i < iDims; i++)
	{
		m_psaBound[i].lLbound = 0;
		if (plMaxElements == NULL)
			m_psaBound[i].cElements = 0;
		else
			m_psaBound[i].cElements = *(plMaxElements + i);
	}

	m_paIndices = new long [iDims];

	m_psa = ::SafeArrayCreate(vtArrayType, iDims, m_psaBound);
	if (m_psa == NULL)
	{
		delete [] m_psaBound;
		m_Initialized = FALSE;
	}

	return (TRUE);
}

//CSafeArray::Destroy
//Purpose: Delete the array and restore any allocated memory.

void CSafeArray::Destroy()
{
	if (!m_Initialized)
		return;

	if (m_paIndices)
		delete [] m_paIndices;

	if (m_psaBound)
	{
		delete [] m_psaBound;
		m_psaBound = NULL;
	}

	if (m_psa)
	{
		::SafeArrayDestroy(m_psa);
		m_psa = NULL;
	}
}

//CSafeArray::LBound
//Purpose:Returns the Lower bound of CSafeArray object.

long CSafeArray::LBound(int iDim)
{
	long lrc;
	HRESULT  hResult;

	hResult = ::SafeArrayGetLBound(m_psa, iDim, &lrc);
	return (lrc);
}

//CSafeArray::UBound
//Purpose: Returns the Upper bound of CSafeArray object.

long CSafeArray::UBound(int iDim)
{
	long lrc;
	HRESULT  hResult;

	hResult = ::SafeArrayGetUBound(m_psa, iDim, &lrc);
	return (lrc + 1);
}

//CSafeArray::GetPSA
//Purpose:Returns a pointer to the wrappered SafeArray object.

SAFEARRAY * CSafeArray::GetPSA()
{
	return (m_psa);
}

//CSafeArray::iGet
//Purpose: Returns the integer Value at the specified dimension(s).

int CSafeArray::iGet(int iIndex1, ...)
{
	va_list argList;
	va_start(argList, iIndex1);

	int		i;
	int		iValue;
	HRESULT	hResult;

	m_paIndices[0] = (long)iIndex1;
	for (i = 1; i < m_iDims; i++)
		m_paIndices[i] = va_arg(argList, long);
	hResult = ::SafeArrayGetElement(m_psa, m_paIndices, &iValue);
	return (iValue);
}

//CSafeArray::iGet
//Purpose: Returns the integer Value at the specified dimension(s).

int CSafeArray::iGet(long lIndex1, ...)
{
	int i;
	int iValue;
	HRESULT hResult;

	va_list argList;
	va_start(argList, lIndex1);

	m_paIndices[0] = lIndex1;
	for (i = 1; i < m_iDims; i++)
		m_paIndices[i] = va_arg(argList, long);
	hResult = ::SafeArrayGetElement(m_psa, m_paIndices, &iValue);
	return (iValue);
}

//CSafeArray::iPut
//Purpose: Store the integer Value at the specified dimension(s).

int CSafeArray::iPut(int iValue, long lIndex1, ...)
{
	int i;
	HRESULT hResult;

	va_list argList;
	va_start(argList, lIndex1);

	m_paIndices[0] = lIndex1;
	for (i = 1; i < m_iDims; i++)
		m_paIndices[i] = va_arg(argList, long);
	hResult = ::SafeArrayPutElement(m_psa, m_paIndices, &iValue);
	return (iValue);
}

//CSafeArray::lGet
//Purpose: Returns the long Value at the specified dimension(s).

long CSafeArray::lGet(long lIndex1, ...)
{
	int i;
	long lValue;
	HRESULT hResult;

	va_list argList;
	va_start(argList, lIndex1);

	m_paIndices[0] = lIndex1;
	for (i = 1; i < m_iDims; i++)
		m_paIndices[i] = va_arg(argList, long);
	hResult = ::SafeArrayGetElement(m_psa, m_paIndices, &lValue);
	return (lValue);
}

//CSafeArray::lPut
//Purpose:Store the long Value at the specified dimension(s).

long CSafeArray::lPut(long lValue, long lIndex1, ...)
{
	int i;
	HRESULT hResult;

	va_list argList;
	va_start(argList, lIndex1);

	m_paIndices[0] = lIndex1;
	for (i = 1; i < m_iDims; i++)
		m_paIndices[i] = va_arg(argList, long);
	hResult = ::SafeArrayPutElement(m_psa, m_paIndices, &lValue);
	return (lValue);
}

//CSafeArray::bstrGet
//Purpose: Returns the BSTR Value at the specified dimension(s).

BSTR CSafeArray::bstrGet(int iIndex1, ...)
{
	int i;
	BSTR	bstrValue;
	HRESULT hResult;
	BSTR	bstrTemp;

	va_list argList;
	va_start(argList, iIndex1);

	m_paIndices[0] = (long)iIndex1;
	for (i = 1; i < m_iDims; i++)
		m_paIndices[i] = va_arg(argList, long);
	hResult = ::SafeArrayGetElement(m_psa, m_paIndices, &bstrTemp);
	if (hResult == NOERROR)
		bstrValue = bstrTemp;
	else
		bstrValue = NULL;

	return (bstrValue);
}

//CSafeArray::bstrGet
//Purpose: Returns the BSTR Value at the specified dimension(s).

BSTR CSafeArray::bstrGet(long lIndex1, ...)
{
	int		i;
	BSTR	bstrValue;
	HRESULT hResult;
	BSTR	bstrTemp;

	va_list argList;
	va_start(argList, lIndex1);

	m_paIndices[0] = lIndex1;
	for (i = 1; i < m_iDims; i++)
		m_paIndices[i] = va_arg(argList, long);
	hResult = ::SafeArrayGetElement(m_psa, m_paIndices, &bstrTemp);
	bstrValue = (hResult == NOERROR) ? bstrTemp : NULL;

	return (bstrValue);
}

//CSafeArray::bstrPut
//Purpose: Store the BSTR Value at the specified dimension(s).

BSTR CSafeArray::bstrPut(BSTR bstrValue, long lIndex1, ...)
{
	int		i;
	HRESULT hResult;
	BSTR	bstrTemp;

	va_list argList;
	va_start(argList, lIndex1);

	m_paIndices[0] = lIndex1;
	for (i = 1; i < m_iDims; i++)
		m_paIndices[i] = va_arg(argList, long);
	bstrTemp = ::SysAllocString(bstrValue);
	hResult = ::SafeArrayPutElement(m_psa, m_paIndices, bstrTemp);

	if (hResult != NOERROR)
		bstrValue = NULL;
	return (bstrValue);
}

//CSafeArray::iGet
//Purpose:
//Given a safearray pointer, return the integer Value
//at the specified dimension(s).

int CSafeArray::iGet(SAFEARRAY * psa, int iIndex1, ...)
{
	int		i;
	int		iDims;
	int		iValue;
	HRESULT hResult;
	long *	m_paIndices;

	va_list argList;
	va_start(argList, iIndex1);

	iDims = ::SafeArrayGetDim(psa);
	m_paIndices = new long [iDims];

	m_paIndices[0] = (long)iIndex1;
	for (i = 1; i < iDims; i++)
		m_paIndices[i] = va_arg(argList, long);

	hResult = ::SafeArrayGetElement(psa, m_paIndices, &iValue);

	va_end(argList);
	delete [] m_paIndices;
	return (iValue);
}

//CSafeArray::iGet
//Purpose:
//Given a safearray pointer, return the integer Value
//at the specified dimension(s).

int CSafeArray::iGet(SAFEARRAY * psa, long lIndex1, ...)
{
	int		i;
	int		iDims;
	int		iValue;
	HRESULT hResult;
	long *	m_paIndices;

	va_list argList;
	va_start(argList, lIndex1);

	iDims = ::SafeArrayGetDim(psa);
	m_paIndices = new long [iDims];

	m_paIndices[0] = lIndex1;
	for (i = 1; i < iDims; i++)
		m_paIndices[i] = va_arg(argList, long);

	hResult = ::SafeArrayGetElement(psa, m_paIndices, &iValue);

	va_end(argList);
	delete [] m_paIndices;

	return (iValue);
}

//CSafeArray::iPut
//Purpose:
//Given a safearray pointer, store the integer Value
//at the specified dimension(s).

int CSafeArray::iPut(SAFEARRAY * psa, int iValue, long lIndex1, ...)
{
	int i;
	int iDims;
	HRESULT hResult;
	long * m_paIndices;
	va_list argList;
	va_start(argList, lIndex1);

	iDims = ::SafeArrayGetDim(psa);
	m_paIndices = new long [iDims];

	m_paIndices[0] = lIndex1;
	for (i = 1; i < iDims; i++)
		m_paIndices[i] = va_arg(argList, long);

	hResult = ::SafeArrayPutElement(psa, m_paIndices, &iValue);

	va_end(argList);
	delete [] m_paIndices;

	return (iValue);
}

//CSafeArray::bstrGet
//Purpose:
//Given a safearray pointer, return a BSTR Value
//at the specified dimension(s).

BSTR CSafeArray::bstrGet(SAFEARRAY * psa, int iIndex1, ...)
{
	int		i;
	int		iDims;
	BSTR	bstrValue;
	HRESULT hResult;
	long *	m_paIndices;
	BSTR	bstrTemp;

	va_list argList;
	va_start(argList, iIndex1);

	iDims = ::SafeArrayGetDim(psa);
	m_paIndices = new long [iDims];

	m_paIndices[0] = (long)iIndex1;
	for (i = 1; i < iDims; i++)
		m_paIndices[i] = va_arg(argList, long);

	hResult = ::SafeArrayGetElement(psa, m_paIndices, &bstrTemp);
	bstrValue = (hResult == NOERROR) ? bstrTemp : NULL;

	va_end(argList);
	delete [] m_paIndices;
	return (bstrValue);
}

//CSafeArray::bstrGet
//Purpose:
//Given a safearray pointer, return a BSTR Value
//at the specified dimension(s).

BSTR CSafeArray::bstrGet(SAFEARRAY * psa, long lIndex1, ...)
{
	int		i;
	int		iDims;
	BSTR	bstrValue;
	HRESULT hResult = NOERROR;
	long *	m_paIndices;
	BSTR	bstrTemp;

	va_list argList;
	va_start(argList, lIndex1);

	hResult = NOERROR;
	iDims = ::SafeArrayGetDim(psa);
	m_paIndices = new long [iDims];

	m_paIndices[0] = lIndex1;
	for (i = 1; i < iDims; i++)
		m_paIndices[i] = va_arg(argList, long);

	hResult = ::SafeArrayGetElement(psa, m_paIndices, &bstrTemp);
	bstrValue = (hResult == NOERROR) ? bstrTemp : NULL;

	va_end(argList);
	delete [] m_paIndices;
	return (bstrValue);
}

//CSafeArray::bstrPut
//Purpose:
//Given a safearray pointer, store a BSTR Value
//at the specified dimension(s).

BSTR CSafeArray::bstrPut(SAFEARRAY * psa, BSTR bstrValue, long lIndex1, ...)
{
	int		i;
	int		iDims;
	HRESULT hResult;
	long *	m_paIndices;

	va_list argList;
	va_start(argList, lIndex1);

	hResult = NOERROR;
	iDims = ::SafeArrayGetDim(psa);
	m_paIndices = new long [iDims];

	m_paIndices[0] = lIndex1;
	for (i = 1; i < iDims; i++)
		m_paIndices[i] = va_arg(argList, long);

	hResult = ::SafeArrayPutElement(psa, m_paIndices, &bstrValue);
	if (hResult != NOERROR)
		bstrValue = NULL;

	va_end(argList);
	delete [] m_paIndices;

	return (bstrValue);
}
