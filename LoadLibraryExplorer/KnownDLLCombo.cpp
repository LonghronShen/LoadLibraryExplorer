// KnownDLLCombo.cpp : implementation file
//

#include "stdafx.h"
#include "KnownDLLCombo.h"
#include ".\knowndllcombo.h"


// CKnownDLLCombo

IMPLEMENT_DYNAMIC(CKnownDLLCombo, CSmartDropdown)

CKnownDLLCombo::CKnownDLLCombo()
{
}

CKnownDLLCombo::~CKnownDLLCombo()
{
}


BEGIN_MESSAGE_MAP(CKnownDLLCombo, CSmartDropdown)
END_MESSAGE_MAP()



// CKnownDLLCombo message handlers


/****************************************************************************
*                          CKnownDLLCombo::DrawItem
* Inputs:
*       LPDRAWITEMSTRUCT dis
* Result: void
*       
* Effect: 
*       Draws the information
* Notes:
*       Appearance:
*               Non-loaded DLL   |  whatever.dll        |
*               Loaded DLL       |O whatever.dll        |
* O is shown in WingDings, which means that it will be a checkmark
****************************************************************************/

void CKnownDLLCombo::DrawItem(LPDRAWITEMSTRUCT dis)
    {
     CRect r = dis->rcItem;
     CDC * dc = CDC::FromHandle(dis->hDC);
     COLORREF txcolor;
     COLORREF bkcolor;

     CSize sz = dc->GetTextExtent(_T("WWW"));

     int saved = dc->SaveDC();

     if(dis->itemState & ODS_SELECTED)
        { /* selected */
         if(::GetFocus() == m_hWnd)
            { /* has focus */
             bkcolor = ::GetSysColor(COLOR_HIGHLIGHT);
             txcolor = ::GetSysColor(COLOR_HIGHLIGHTTEXT);
            } /* has focus */       
         else
            { /* no focus */
             bkcolor = ::GetSysColor(COLOR_3DFACE);
             txcolor = ::GetSysColor(COLOR_WINDOWTEXT);
            } /* no focus */
        } /* selected */
     else
        { /* unselected */
         if(dis->itemState & (ODS_DISABLED | ODS_GRAYED))
            txcolor = ::GetSysColor(COLOR_GRAYTEXT);
         else
            txcolor = ::GetSysColor(COLOR_WINDOWTEXT);

         bkcolor = ::GetSysColor(COLOR_WINDOW);
        } /* unselected */

     dc->FillSolidRect(&r, bkcolor);

     dc->SetBkColor(bkcolor);
     dc->SetTextColor(txcolor);

     CRect focus;

     KnownDLLInfo * info = (KnownDLLInfo *)dis->itemData;

     if(dis->itemID != (UINT)-1 )
        { /* has item */
         r.left += ::GetSystemMetrics(SM_CXBORDER);

         if(info->preloaded)
            { /* preloaded */
             CFont * f = GetFont();
             ASSERT(f != NULL);
             LOGFONT lf;
             f->GetLogFont(&lf);
             _tcscpy(lf.lfFaceName, _T("Wingdings"));
             lf.lfCharSet = SYMBOL_CHARSET;
             lf.lfPitchAndFamily = FF_DECORATIVE;

             CFont wings;
             wings.CreateFontIndirect(&lf);

#ifdef _DEBUG //----------------------------------------------------------------+
             LOGFONT wf;                                                     // |
             wings.GetLogFont(&wf);                                          // |
#endif //-----------------------------------------------------------------------+

             int save2 = dc->SaveDC();
             dc->SelectObject(&wings);
             //dc->TextOut(r.left, r.top, CString(_T("ABC")));
             dc->TextOut(r.left, r.top, CString(_T("\xFC")));

#ifdef _DEBUG //----------------------------------------------------------------+
             CFont * fw = dc->GetCurrentFont();                              // |
             LOGFONT tf;                                                     // |
             fw->GetLogFont(&tf);                                            // |
             ::GdiFlush();                                                   // |
#endif //-----------------------------------------------------------------------+

             dc->RestoreDC(save2);
            } /* preloaded */

         r.left += sz.cx;
         
         dc->TextOut(r.left, r.top, info->name);
#ifdef _DEBUG //----------------------------------------------------------------+
         int len = info->name.GetLength();                                   // |
         ::GdiFlush();                                                       // |
#endif //-----------------------------------------------------------------------+
        } /* has item */

     if(dis->itemState & ODS_FOCUS && dis->itemAction != ODA_SELECT)
        dc->DrawFocusRect(&dis->rcItem);

     dc->RestoreDC(saved);
    }

/****************************************************************************
*                         CKnownDLLCombo::DeleteItem
* Inputs:
*       LPDELETEITEMSTRUCT dis:
* Result: void
*       
* Effect: 
*       Deletes the itemdata
****************************************************************************/

void CKnownDLLCombo::DeleteItem(LPDELETEITEMSTRUCT dis)
    {
     KnownDLLInfo * info = (KnownDLLInfo *)dis->itemData;

     delete info;

     CSmartDropdown::DeleteItem(dis);
    }

/****************************************************************************
*                         CKnownDLLCombo::CompareItem
* Inputs:
*       LPCOMPAREITEMSTRUCT cis:
* Result: int
*       -1 = item 1 sorts before item 2
*       0 = item 1 and item 2 sort the same
*       1 = item 1 sorts after item 2
****************************************************************************/

int CKnownDLLCombo::CompareItem(LPCOMPAREITEMSTRUCT cis)
    {
     KnownDLLInfo * item1 = (KnownDLLInfo *)cis->itemData1;
     KnownDLLInfo * item2 = (KnownDLLInfo *)cis->itemData2;

     return _tcsicmp(item1->name, item2->name);
    }
