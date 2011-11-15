#if !defined(AFX_SEARCHINFO_H__E214C1A2_DA61_4F60_BDB5_31790946BC35__INCLUDED_)
#define AFX_SEARCHINFO_H__E214C1A2_DA61_4F60_BDB5_31790946BC35__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// SearchInfo.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CSearchInfo dialog

class CSearchInfo : public CDialog
{
// Construction
public:
	CSearchInfo(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CSearchInfo)
	enum { IDD = frmSearchInfo };
	CListCtrl	m_lsvMoveList;
	CListCtrl	m_lsvDepthTimeCost;
	CListCtrl	m_lsvValue;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSearchInfo)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CSearchInfo)
		// NOTE: the ClassWizard will add member functions here
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_SEARCHINFO_H__E214C1A2_DA61_4F60_BDB5_31790946BC35__INCLUDED_)
