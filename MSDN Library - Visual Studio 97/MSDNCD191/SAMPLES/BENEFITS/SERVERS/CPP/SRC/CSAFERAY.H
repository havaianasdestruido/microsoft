//CSaferay.h  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Header file for the CSafeArray object.

#ifndef __CSAFERAY_H__
#define __CSAFERAY_H__

#include <windows.h>
#include <stdarg.h>
#include <oleauto.h>
#include <comutil.h>

class CSafeArray
{
private:
	//Private member variables
	SAFEARRAY*		m_psa;
	SAFEARRAYBOUND*	m_psaBound;
	BOOL			m_Initialized;
	int				m_iDims;
	long *			m_paIndices;
	long			m_oid;

public:
	//Public static methods
	static int iGet(SAFEARRAY * psa, int iIndex1, ...);
	static int iGet(SAFEARRAY * psa, long lIndex1, ...);
	static int iPut(SAFEARRAY * psa, int iValue, long lIndex1, ...);
	static BSTR bstrGet(SAFEARRAY * psa, int iIndex1, ...);
	static BSTR bstrGet(SAFEARRAY * psa, long lIndex1, ...);
	static BSTR bstrPut(SAFEARRAY * psa, BSTR bstrValue, long lIndex1, ...);

	//Public methods
	CSafeArray();
	CSafeArray(SAFEARRAY * psa);
	CSafeArray(VARTYPE vtArrayType, int iDims = 1, long * plMaxElements = NULL);
	~CSafeArray();

	BOOL	Create(VARTYPE vtArrayType, int iDims = 1, long * plMaxElements = NULL);
	void	Destroy();
	SAFEARRAY * GetPSA();
	BOOL	IsInitialized();
	long	LBound(int iDim);
	long	UBound(int iDim);
		
	BSTR	bstrGet(int iIndex1, ...);
	BSTR	bstrGet(long lIndex1, ...);
	BSTR	bstrPut(BSTR bstrValue, long lIndex1, ...);
	int		iGet(int iIndex1, ...);
	int		iGet(long lIndex1, ...);
	int		iPut(int iValue, long lIndex1, ...);
	long	lGet(int iIndex1, ...);
	long	lGet(long lIndex1, ...);
	long	lPut(long lValue, long lIndex1, ...);
};

typedef CSafeArray *PCSafeArray;

#endif //#ifndef __CSAFERAY_H__