// MfcXq.h : main header file for the MFCXQ application
//

#if !defined(AFX_MFCXQ_H__5B0098DA_EF79_4548_B820_5D9732B80C22__INCLUDED_)
#define AFX_MFCXQ_H__5B0098DA_EF79_4548_B820_5D9732B80C22__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CMfcXqApp:
// See MfcXq.cpp for the implementation of this class
//

class CMfcXqApp : public CWinApp
{
public:
	CMfcXqApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMfcXqApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CMfcXqApp)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft eMbedded Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MFCXQ_H__5B0098DA_EF79_4548_B820_5D9732B80C22__INCLUDED_)
