// MfcXqDlg.h : header file
//

#if !defined(AFX_MFCXQDLG_H__2BCED10F_0AD6_4DBE_9794_40ED280FA80A__INCLUDED_)
#define AFX_MFCXQDLG_H__2BCED10F_0AD6_4DBE_9794_40ED280FA80A__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

/////////////////////////////////////////////////////////////////////////////
// CMfcXqDlg dialog

class CMfcXqDlg : public CDialog
{
// Construction
public:
	int mQSNodesALayer;
	int mVisitNodesALayer;
	byte m_MoveCount;

	CMfcXqDlg(CWnd* pParent = NULL);	// standard constructor
	void ClickSquare(int sq);
	void ResponseMove(void);
	void Startup(void);
	void SearchMain(void);
	int SearchQuiesc(int vlAlpha, int vlBeta);
	int SearchFull(int vlAlpha, int vlBeta, int nDepth, BOOL bNoNull);
	int SearchRoot(int nDepth);
	void DoEvents();
	void AddLsvValue(int vl, int i, int t, int VisitNodesTotal);

// Dialog Data
	//{{AFX_DATA(CMfcXqDlg)
	enum { IDD = IDD_MFCXQ_DIALOG };
	CListBox	m_lstMoveDesc;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMfcXqDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CMfcXqDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OncmdRedFirst();
	afx_msg void OnSelchangelstMoveDesc();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft eMbedded Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MFCXQDLG_H__2BCED10F_0AD6_4DBE_9794_40ED280FA80A__INCLUDED_)
