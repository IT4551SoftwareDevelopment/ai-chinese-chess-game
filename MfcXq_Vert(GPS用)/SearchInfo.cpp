// SearchInfo.cpp : implementation file
//

#include "stdafx.h"
#include "MfcXq.h"
#include "SearchInfo.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CSearchInfo dialog


CSearchInfo::CSearchInfo(CWnd* pParent /*=NULL*/)
	: CDialog(CSearchInfo::IDD, pParent)
{
	//{{AFX_DATA_INIT(CSearchInfo)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void CSearchInfo::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CSearchInfo)
	DDX_Control(pDX, lsvMoveList, m_lsvMoveList);
	DDX_Control(pDX, lsvDepthTimeCost, m_lsvDepthTimeCost);
	DDX_Control(pDX, lsvValue, m_lsvValue);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CSearchInfo, CDialog)
	//{{AFX_MSG_MAP(CSearchInfo)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CSearchInfo message handlers
