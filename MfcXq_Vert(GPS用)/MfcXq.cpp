// MfcXq.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "MfcXq.h"
#include "MfcXqDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CMfcXqApp

BEGIN_MESSAGE_MAP(CMfcXqApp, CWinApp)
	//{{AFX_MSG_MAP(CMfcXqApp)
	//}}AFX_MSG
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CMfcXqApp construction

CMfcXqApp::CMfcXqApp()
	: CWinApp()
{
}

/////////////////////////////////////////////////////////////////////////////
// The one and only CMfcXqApp object

CMfcXqApp theApp;

/////////////////////////////////////////////////////////////////////////////
// CMfcXqApp initialization

BOOL CMfcXqApp::InitInstance()
{
	// Standard initialization

	CMfcXqDlg dlg;
	m_pMainWnd = &dlg;

	int nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
	}
	else if (nResponse == IDCANCEL)
	{
	}

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}
