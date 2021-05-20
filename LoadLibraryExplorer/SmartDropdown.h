#ifndef __SMARTDROPDOWN_H__
#define __SMARTDROPDOWN_H__

// SmartDropdown.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CSmartDropdown window

class CSmartDropdown : public CComboBox
{
// Construction
public:
	CSmartDropdown();

// Attributes
public:
        int maxlen;  // maximum number of items displayed
                     // 0 - no limit
                     // -1 - limit to screen height (NYI)

// Operations
public:
   int Select(DWORD itemval);

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSmartDropdown)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CSmartDropdown();

	// Generated message map functions
protected:
	//{{AFX_MSG(CSmartDropdown)
	afx_msg void OnDropdown();
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

#endif //__SMARTDROPDOWN_H__