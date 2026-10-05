//benexe.h    - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Greg Shumate
//Purpose     - Header declarations for the Benefits Exe.

#ifndef _BenExe_H_
#define _BenExe_H_


//Get the object definitions
#include "benefitsr.h"


LRESULT APIENTRY BenExeWndProc(HWND, UINT, WPARAM, LPARAM);

class CApp
{
	friend LRESULT APIENTRY BenExeWndProc(HWND, UINT, WPARAM, LPARAM);

 private:
	enum	CLSFACT	{	eCAdmin = 0, eCBank, eCEmployee, eCBenefit, eCDependant, eCBenefitLink,
						eCList, eCListBank, eCListEmployee, eCListBenefit, eCListDependant, eCListBenefitLink, 
						eLAST };

	// Private Member Variables
	DWORD           m_dwRegCO[eLAST];   //Registration key
	BOOL            m_fInitialized;     //Did CoInitialize work?
	HINSTANCE       m_hInst;            //WinMain parameters
	HINSTANCE       m_hInstPrev;
	LPCLASSFACTORY  m_pIClassFactory[eLAST];   //Our class factory
	HWND            m_hWnd;             //Main window handle
	UINT            m_nCmdShow;
	LPSTR           m_pszCmdLine;
	
	// Private Member Methods
	BOOL	InitCF(void);

 public:
	// Public Member Methods
	CApp(HINSTANCE, HINSTANCE, LPSTR, UINT);
	~CApp(void);

	BOOL	Init(void);
	BOOL	Register(void);
};

typedef CApp *PAPP;

#define CBWNDEXTRA              sizeof(PAPP)
#define BENEXEWL_STRUCTURE       0

void ObjectDestroyed(void);

#endif
