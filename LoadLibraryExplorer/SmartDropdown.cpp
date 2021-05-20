// SmartDropdown.cpp : implementation file
//

#include "stdafx.h"
#include "SmartDropdown.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CSmartDropdown

CSmartDropdown::CSmartDropdown()
{
    maxlen = -1; // default to screen limit
}

CSmartDropdown::~CSmartDropdown()
{
}


BEGIN_MESSAGE_MAP(CSmartDropdown, CComboBox)
	//{{AFX_MSG_MAP(CSmartDropdown)
	ON_CONTROL_REFLECT(CBN_DROPDOWN, OnDropdown)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CSmartDropdown message handlers


/****************************************************************************
*                              CSmartDropdown::Select
* Inputs:
*       DWORD itemval: ItemData value
* Result: int
*       The index of the selected value, or CB_ERR if no selection possible
* Effect: 
*       If the value exists in the list, selects it, otherwise selects -1
****************************************************************************/

int CSmartDropdown::Select(DWORD itemval)
   {
    for(int i = 0; i < GetCount(); i++)
       { /* select */
	DWORD_PTR data = GetItemData(i);
	if(itemval == data)
	   { /* found it */
	    SetCurSel(i);
	    return i;
	   } /* found it */
       } /* select */

    SetCurSel(-1);
    return CB_ERR;
   }

void CSmartDropdown::OnDropdown() 
{
 int n = GetCount();
 n = max(n, 2);

 int ht = GetItemHeight(0);
 CRect r;
 GetWindowRect(&r);

 if(maxlen > 0)
    n = max(maxlen, 2);

 CSize sz;
 sz.cx = r.Width();
 sz.cy = ht * (n + 2);

 if(maxlen < 0)
    { /* screen limit */
     if(r.top - sz.cy < 0 || r.bottom + sz.cy > ::GetSystemMetrics(SM_CYSCREEN))
        { /* invoke limit */
         // Compute the largest distance the dropdown can appear, 
         // relative to the screen (not the window!)

         int k = max( (r.top / ht), 
                      (::GetSystemMetrics(SM_CYSCREEN) - r.bottom) / ht);

         // compute new space. Note that we don't really fill the screen.
         // We only have to use this size if it is smaller than the max size
         // actually required
         int ht2 = ht * k;
         sz.cy = min(ht2, sz.cy);
        } /* invoke limit */
    } /* screen limit */

 SetWindowPos(NULL, 0, 0, sz.cx, sz.cy, SWP_NOMOVE | SWP_NOZORDER);
	
}
