// TraceList.cpp : implementation file
//

#include "stdafx.h"
#include "resource.h"
#include "Printer.h"
#include "TraceEvent.h"
#include "Trace.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

static UINT UWM_TRACE_CHANGE = ::RegisterWindowMessage(UWM_TRACE_CHANGE_MSG);// *** JMN
//    IDS_TRACELOG_UNKNOWN_ERROR "Unknown error %d (0x%08x)"
//    IDS_TRACELOG_FILTER     "Log files (*.log)|*.log|"
//    IDS_TRACELOG_FILE_SAVE  "Save Trace Log"
//    IDS_TRACELOG_FILE_SAVE_AS "Save Trace Log As"
//    IDS_TRACELOG_FILE_OPEN_FAILED "Trace Log file open failed"
//    IDS_TRACELOG_CAPTION    "Trace Log"
//    IDS_TRACELOG_DELETED_ENTRIES "Entries have been deleted from log (%d)"
/////////////////////////////////////////////////////////////////////////////
// CTraceList

CTraceList::CTraceList()
{
 fixed = NULL; // no initial font
 height = 0;   // initial height is unknown
 modified = FALSE;
 toDisk = FALSE;
 horizontalExtent = 0;
 indent1 = 0;
 indent2 = 0;
 limit = 0x7FFFFFFF; // actually, this is ignored if 'limited' is FALSE
 limited = FALSE;
 showThreadID = TRUE;
 boundary = 0;
 deleted = 0;
}

CTraceList::~CTraceList()
{
}

BEGIN_MESSAGE_MAP(CTraceList, CListBox)
	//{{AFX_MSG_MAP(CTraceList)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTraceList message handlers

/****************************************************************************
*                          CTraceList::ResetContent
* Result: void
*       
* Effect: 
*       Resets the listbox content
****************************************************************************/

void CTraceList::ResetContent()
   {
    modified = TRUE; 
    horizontalExtent = 0; 
    CListBox::ResetContent();
    GetParent()->SendMessage(UWM_TRACE_CHANGE);
   } // CTraceList::ResetContent

/****************************************************************************
*                          CTraceList::DeleteString
* Inputs:
*       int n: Item to delete
* Result: int
*       Result of CListBox::DeleteString
* Effect: 
*       Deletes the specified element
****************************************************************************/

int CTraceList::DeleteString(int n)
   {
    modified = TRUE;
    int result = CListBox::DeleteString(n);
    recomputeHorizontalExtent();
    GetParent()->SendMessage(UWM_TRACE_CHANGE);
    return result;
   } // CTraceList::DeleteString

/****************************************************************************
*                            CTraceList::setLimit
* Inputs:
*       int limit: Amount of limit, or < 0 to disable limiting
* Result: void
*       
* Effect: 
*       Sets the limit value, or disables limiting
****************************************************************************/

#define MIN_LOG_ENTRIES 20

void CTraceList::setLimit(int limit)
    {
     if(limit < 0)
        { /* disable */
	 limited = FALSE;
	 return;
	} /* disable */

     if(limit < MIN_LOG_ENTRIES) 
	limit = MIN_LOG_ENTRIES;
     if(limit < boundary)
	limit = boundary + MIN_LOG_ENTRIES;
     this->limit = limit; 
     limited = TRUE;     
    }

/****************************************************************************
*                           CTraceList::enableLimit
* Inputs:
*       BOOL enable: TRUE to enable limit
* Result: void
*       
* Effect: 
*       Enables the limit. Purges messages beyond limit
****************************************************************************/

void CTraceList::enableLimit(BOOL enable)
    {
     limited = enable;
     BOOL changed = FALSE;
     while(limited && GetCount() > limit - 1)
	{ /* trim */
	 changed = TRUE;
	 CListBox::DeleteString(0);
	} /* trim */
     if(changed)
	{ /* handle change */
	 recomputeHorizontalExtent();
	 GetParent()->SendMessage(UWM_TRACE_CHANGE);
	} /* handle change */
    } // CTraceList::EnableLimit

/****************************************************************************
*                           CTraceList::MeasureItem
* Inputs:
*       LPMEASUREITEMSTRUCT mis: measure struct
* Result: void
*       
* Effect: 
*       Asks for the display height, which an entry object returns as an
*	integer multiple of line height
****************************************************************************/

void CTraceList::MeasureItem(LPMEASUREITEMSTRUCT mis) 
{
 if(height == 0)
    { /* compute double-height */
     CClientDC dc(this);
     CFont * f = GetFont();
     dc.SelectObject(f);
     TEXTMETRIC tm;
     dc.GetTextMetrics(&tm);
     height = tm.tmHeight + tm.tmInternalLeading;
    } /* compute double-height */

 TraceEvent * e = (TraceEvent *)mis->itemData;
 mis->itemHeight = height * e->displayHeight();
}

/****************************************************************************
*                            CTraceList::isVisible
* Inputs:
*       int n: Position in list
* Result: BOOL
*       TRUE if visible
*	FALSE if hidden
* Notes: 
*       It is assumed that this is working on the last item in the list
*	so we are only concerned about it falling off the bottom
****************************************************************************/

BOOL CTraceList::isVisible(int n, BOOL fully)
    {
     CRect item;
     CRect r;
     GetItemRect(n, &item);
     GetClientRect(&r);

     if(!fully && item.top < r.bottom)
	return TRUE;
     if(fully && item.bottom < r.bottom)
	return TRUE;
     
     return FALSE;
    }

/****************************************************************************
*                            CTraceList::setToDisk
* Inputs:
*       BOOL v: Value to set
* Result: BOOL
*       TRUE if successful, FALSE if failed to open
* Effect: 
*       If not already set, requests a file to log to
* Notes:
*	Case	toDisk		v	Effect
*	I	FALSE		FALSE	ignored
*	II	TRUE		FALSE	toDisk = FALSE
*	III	TRUE		TRUE	ignored
*	IV	FALSE		TRUE	prompt for name
****************************************************************************/

BOOL CTraceList::setToDisk(BOOL v)
    {
     if(!toDisk && !v)
        { /* I. */
	 return FALSE; // ignore false-to-false
	} /* I. */

     if(toDisk)
        { /* II., III. */
	 toDisk = v;
	 return TRUE;
	} /* II., III. */

     // IV.
     toDisk = doSave(FALSE); // force prompt for file name and save current state
     return toDisk;
    }

/****************************************************************************
*                            CTraceList::AddString
* Inputs:
*       TraceEvent * event: Event to add
* Result: int
*       Place where event was added
* Effect: 
*       If the last item in the box was visible before the add, and is not
*	visible after the add, scroll the box up
****************************************************************************/

int CTraceList::AddString(TraceEvent * event)
    {
     if(toDisk)
	appendToFile(event);

     // During shutdown transient, the window may actually be gone
     if(!::IsWindow(m_hWnd))
        { /* gone */
	 delete event; // already logged to file if file logging on
	 return LB_ERR;
	} /* gone */

     // If we've exceeded the number of lines of trace that are set,
     // delete enough leading lines to get below the limit
     int count = GetCount();

     int top = GetTopIndex();
     BOOL changed = FALSE;
     BOOL first = TRUE;
     while(limited && count > limit - 1)
	{ /* delete */
	 if(first)
	    CListBox::DeleteString(boundary + 1);
	 first = FALSE;
	 CListBox::DeleteString(boundary + 1);
	 deleted++;
	 changed = TRUE;
	 count = GetCount();
	 if(boundary < top)
	    SetTopIndex(top);
	 else
	    SetTopIndex(top - 1);
	} /* delete */

     if(!first)
	{ /* deleted at least one */
         // "%d entries have been deleted from log"
	 InsertString(boundary + 1, new TraceFormatComment(TraceEvent::None, IDS_TRACELOG_DELETED_ENTRIES, deleted + 1));
	} /* deleted at least one */

     // We want to auto-scroll the list box if the last line is showing,
     // event if partially

     count = GetCount();

     BOOL visible = FALSE;
     CRect r;
     if(count > 0)
        { /* nonempty */
	 visible = isVisible(count - 1, FALSE); // any part showing?
	} /* nonempty */
     else
	visible = TRUE;

     // Add the element to the list box

     modified = TRUE; 
     int n = CListBox::AddString((LPCTSTR) event);      

     // If the last line had been visible, and the just-added line is
     // not visible, scroll the list box up to make it visible.

     // Note: This depends on the box being LBS_OWNERDRAWFIXED.  The
     // computation for variable is much more complex!
     if(visible)
        { /* was visible */
	 visible = isVisible(n, TRUE); // all of it showing?
	 if(!visible)
	    { /* dropped off */
	     int top = GetTopIndex();
	     if(top + 1 >= GetCount())
	        { /* can't scroll */
		} /* can't scroll */
	     else
	        { /* scroll it */
		 SetTopIndex(top + 1);
		} /* dropped off */
	    } /* dropped off */
	} /* was visible */

     // Get the width and compute the scroll width
     // Optimization: if the current width is less than the current
     // maximum, we don't need to do anything

     CClientDC dc(this);
     CString s = event->show();
     UINT w = indent1 + indent2 + dc.GetTextExtent(s, s.GetLength()).cx;
     if(w > horizontalExtent)
	{ /* recompute */
	 horizontalExtent = w;
	 SetHorizontalExtent(horizontalExtent);
	} /* recompute */
     else
     if(changed)
	recomputeHorizontalExtent();

     // If this is an error notification, notify the parent to update
     // the controls
     GetParent()->SendMessage(UWM_TRACE_CHANGE);
     return n;
    }

/****************************************************************************
*                            CTraceList::DrawItem
* Inputs:
*       LPDRAWITEMSTRUCT dis:
* Result: void
*       
* Effect: 
*       Draws the item
*
****************************************************************************/

void CTraceList::DrawItem(LPDRAWITEMSTRUCT dis) 
{
 CDC * dc = CDC::FromHandle(dis->hDC);
 CFont * textfont = NULL;
 
 if(height == 0)
    { /* compute line height */
     TEXTMETRIC tm;
     dc->GetTextMetrics(&tm);
     height = tm.tmHeight + tm.tmInternalLeading;
    } /* compute line height */

 if(indent1 == 0)
    { /* compute indent */
     static const CString prefix1 = _T("0000000 ");
     indent1 = dc->GetTextExtent(prefix1, prefix1.GetLength()).cx;
     static const CString prefix2 = _T("00000000  ");
     if(showThreadID)
	indent2 = dc->GetTextExtent(prefix2, prefix2.GetLength()).cx;
     else
	indent2 = 0;
    } /* compute indent */

 if(dis->itemID == -1)
    { /* no items */
     CBrush bg(::GetSysColor(COLOR_WINDOW));
     dc->FillRect(&dis->rcItem, &bg);
     if(dis->itemState & ODS_FOCUS)
        { /* selected */
	 dc->DrawFocusRect(&dis->rcItem);
	} /* selected */
     return;
    } /* no items */

 TraceEvent * e = (TraceEvent *)dis->itemData;

// if(dis->itemState & ODA_FOCUS)
//    { /* focus only */
//     dc->DrawFocusRect(&dis->rcItem);
//     return;
//    } /* focus only */

 int saved = dc->SaveDC();

 COLORREF txcolor;
 COLORREF bkcolor;

 if(dis->itemState & ODS_SELECTED)
    { /* selected */
     if(::GetFocus() == m_hWnd)
        { /* has focus */
	 bkcolor = ::GetSysColor(COLOR_HIGHLIGHT);
	 txcolor = ::GetSysColor(COLOR_HIGHLIGHTTEXT);
	} /* has focus */
     else
        { /* no focus */
	 bkcolor = ::GetSysColor(COLOR_BTNFACE);
	 txcolor = ::GetSysColor(COLOR_BTNTEXT);
	} /* no focus */
    } /* selected */
 else
    { /* not selected */
     txcolor = e->textcolor();
     bkcolor = ::GetSysColor(COLOR_WINDOW);
    } /* not selected */

 {
  CBrush bg(bkcolor);
  dc->FillRect(&dis->rcItem, &bg);
 }

 dc->SetTextColor(txcolor);
 dc->SetBkMode(TRANSPARENT);

 CString id = e->showID();
 int width = dc->GetTextExtent(id).cx;

 textfont = e->getFont();
 if(textfont != NULL)
    dc->SelectObject(textfont);

 // ID is output right-justified in its column
 dc->TextOut(dis->rcItem.left + indent1 - width, dis->rcItem.top, id);

 if(showThreadID)
    { /* show thread */
     CString t = e->showThread();

     dc->TextOut(dis->rcItem.left + indent1, dis->rcItem.top, t);
    } /* show thread */

 CString s = e->show();

 int x = dis->rcItem.left + indent1 - 3;
 dc->MoveTo(x, dis->rcItem.top);
 dc->LineTo(x, dis->rcItem.bottom);

 dc->TextOut(dis->rcItem.left + indent1 + indent2, dis->rcItem.top, s);

 if(dis->itemState & ODS_FOCUS)
    dc->DrawFocusRect(&dis->rcItem);

 dc->RestoreDC(saved);
} // CTraceList::DrawItem

/****************************************************************************
*                           CTraceList::DeleteItem
* Inputs:
*       LPDELETEITEMSTRUCT dis
* Result: void
*       
* Effect: 
*       Deletes the item attached to the element. This is the method called
*	whenever an owner-draw item is deleted.
****************************************************************************/

void CTraceList::DeleteItem(LPDELETEITEMSTRUCT dis) 
{
 TraceEvent * e = (TraceEvent *)dis->itemData;
 delete e;

 CListBox::DeleteItem(dis);
 // Note that we must use PostMessage here, because the item
 // would not be deleted until this method returns. This
 // means that if we use SendMessage, we might try to use
 // a pointer to deleted memory
 GetParent()->PostMessage(UWM_TRACE_CHANGE);
}

/****************************************************************************
*                              TraceList::doSave
* Inputs:
*       BOOL mode: TRUE for Save, FALSE for Save As
* Result: BOOL
*       TRUE if save occurred
*	FALSE if save was cancelled
* Effect: 
*       Does a save
****************************************************************************/

BOOL CTraceList::doSave(BOOL mode)
    {
     BOOL prompt = FALSE;

     if(mode)
        { /* Save */
	 if(SaveFileName.GetLength() == 0)
	    { /* no file name set */
	     prompt = TRUE;
	    } /* no file name set */
	} /* Save */
     else
        { /* SaveAs */
	 prompt = TRUE;
	} /* SaveAs */

     CString name;

     if(prompt)
        { /* get file name */
	 CString filter;
	 filter.LoadString(IDS_TRACELOG_FILTER);
	 CFileDialog dlg(FALSE, _T("log"), NULL,
	 			OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
				filter);

         CString s;
	 s.LoadString(mode ? IDS_TRACELOG_FILE_SAVE : IDS_TRACELOG_FILE_SAVE_AS);
	 dlg.m_ofn.lpstrTitle = s;

	 switch(dlg.DoModal())
	    { /* domodal */
	     case 0: // should never happen
	     case IDCANCEL:
		     return FALSE;
	    } /* domodal */

	 // We get here, it must be IDOK

	 name = dlg.GetPathName();
	} /* get file name */
     else
        { /* already has name */
	 name = SaveFileName;
	} /* already has name */

     TRY
	{
	 CStdioFile f(name, CFile::modeCreate | CFile::modeWrite | CFile::shareExclusive);
	 for(int i = 0; i < GetCount(); i++)
	    { /* write each */
	     TraceEvent * e = (TraceEvent *)GetItemDataPtr(i);
	     f.WriteString(e->showfile());

	     f.WriteString(_T("\r\n"));
	    } /* write each */

	 f.Close();

	 SaveFileName = name;

	 SetModified(FALSE);
	 return TRUE;
	}
     CATCH(CFileException, e)
	{ /* error */
	 DWORD err = ::GetLastError();
	 AfxMessageBox(IDS_TRACELOG_FILE_OPEN_FAILED, MB_ICONERROR | MB_OK);
         AddString(new TraceFormatMessage(TraceEvent::None, err));
         return FALSE; // return without doing anything further
	} /* error */
     END_CATCH
     ASSERT(FALSE); // should never get here
     return FALSE;
    }

/****************************************************************************
*                          CTraceList::appendToFile
* Inputs:
*       TraceEvent * event:
* Result: void
*       
* Effect: 
*       Logs the event to the current file
****************************************************************************/

void CTraceList::appendToFile(TraceEvent * event)
    {
     if(SaveFileName.GetLength() == 0)
	return; // shouldn't be here

     TRY {
	  CStdioFile f(SaveFileName, CFile::modeReadWrite | CFile::modeCreate |
     			CFile::modeNoTruncate | CFile::shareExclusive);


	  f.SeekToEnd();
	  f.WriteString(event->showfile());
	  f.WriteString(_T("\r\n"));
	  f.Close();
	 }
      CATCH(CFileException, e) 
         { 
	  // Do nothing.
	 }
      END_CATCH
    }

/****************************************************************************
*                              CTraceList::toTop
* Result: void
*       
* Effect: 
*       Positions the list box to the top
****************************************************************************/

void CTraceList::toTop()
    {
     if(GetCount() == 0)
	return;

     SetTopIndex(0);
     clearAllSelections();
     SetSel(0, TRUE);
    }

/****************************************************************************
*                              CTraceList::toEnd
* Result: void
*       
* Effect: 
*       Moves to the end of the box
****************************************************************************/

void CTraceList::toEnd()
    {
     int n = GetCount();
     if(n == 0)
	return;
     SetTopIndex(n - 1);
     clearAllSelections();
     SetSel(n - 1, TRUE);
     GetParent()->SendMessage(UWM_TRACE_CHANGE);
    }

/****************************************************************************
*                              CTraceList::canTop
* Result: BOOL
*       TRUE if toTop has meaning
****************************************************************************/

BOOL CTraceList::canTop()
    {
     if(GetCount() == 0)
	return FALSE;
     int n = GetTopIndex();
     if(!GetSel(n))
	return TRUE;
     return (n > 0);
    }

/****************************************************************************
*                              CTraceList::canEnd
* Result: BOOL
*       TRUE if toEnd has meaning
****************************************************************************/

BOOL CTraceList::canEnd()
    {
     int count = GetCount();

     if(count < 2)
	return FALSE;

     if(!isVisible(count - 1, TRUE))
	return TRUE;
     if(GetSelCount() == 0)
	return FALSE;
     if(GetSel(count - 1))
	return FALSE;
     return TRUE;     
    }

/****************************************************************************
*                            CTraceList::findNext
* Inputs:
*       int i: Starting position
* Result: int
*       Next error position, or LB_ERR if not found
****************************************************************************/

int CTraceList::findNext(int i)
    {
     int count = GetCount();

     for(++i ; i < count; i++)
        { /* scan forward */
	 TraceEvent * e = (TraceEvent *)GetItemDataPtr(i);
	 if(e->IsKindOf(RUNTIME_CLASS(TraceError)))
	    { /* found it */
	     return i;
	    } /* found it */
	} /* scan forward */
     return LB_ERR;
    }

/****************************************************************************
*                            CTraceList::findPrev
* Inputs:
*       int i: Starting position
* Result: int
*       Previous error position, or LB_ERR if not found
****************************************************************************/

int CTraceList::findPrev(int i)
    {
     int count = GetCount();
     if(count == 0)
	return LB_ERR;

     for(--i ; i >= 0; i--)
        { /* scan backward */
	 TraceEvent * e = (TraceEvent *)GetItemDataPtr(i);
	 if(e->IsKindOf(RUNTIME_CLASS(TraceError)))
	    { /* found it */
	     return i;
	    } /* found it */
	} /* scan backward */
     return LB_ERR;
    }

/****************************************************************************
*                             CTraceList::toNext
* Result: BOOL

*	FALSE if we didn't move
* Effect: 
*       Sets the selection to the next error
****************************************************************************/

BOOL CTraceList::toNext()
    {
     int start = GetCaretIndex(); 

     int i = findNext(start);
     if(i == LB_ERR)
	return FALSE;

     if(i > 3)
	SetTopIndex(i - 3);
     else
	SetTopIndex(i);

     clearAllSelections();
     SetSel(i, TRUE);
     GetParent()->SendMessage(UWM_TRACE_CHANGE);
     return TRUE;
    }

/****************************************************************************
*                             CTraceList::toPrev
* Result: BOOL
*       TRUE if we moved
*	FALSE if we didn't move
* Effect: 
*       Sets the selection to the previous error
****************************************************************************/

BOOL CTraceList::toPrev()
    {
     int start = GetCaretIndex();

     int i = findPrev(start);
     if(i == LB_ERR)
	return FALSE;

     if(i > 3)
	SetTopIndex(i - 3);
     else
	SetTopIndex(i);

     clearAllSelections();
     SetSel(i, TRUE);

     GetParent()->SendMessage(UWM_TRACE_CHANGE);
     return TRUE;
    }

/****************************************************************************
*                             CTraceList::canPrev
* Result: BOOL
*       TRUE if there is a previous error
****************************************************************************/

BOOL CTraceList::canPrev()
    {
     return findPrev(GetCaretIndex()) != LB_ERR;
    }

/****************************************************************************
*                             CTraceList::canNext
* Result: BOOL
*       TRUE if there is a next error
****************************************************************************/

BOOL CTraceList::canNext()
    {
     return findNext(GetCaretIndex()) != LB_ERR;
    }

/****************************************************************************
*                             CTraceList::canCopy
* Result: BOOL
*       TRUE if Copy operation can be done
*	FALSE if nothing to copy
****************************************************************************/

BOOL CTraceList::canCopy()
    {
     return GetSelCount() > 0;
    } // CTraceList::canCopy

/****************************************************************************
*                             CTraceList::canCut
* Result: BOOL
*       TRUE if Cut operation can be done
*	FALSE if nothing to cut
****************************************************************************/

BOOL CTraceList::canCut()
   {
    return GetSelCount() > 0;
   } // CTraceList::canCut

/****************************************************************************
*                           CTraceList::canClearAll
* Result: BOOL
*       TRUE if ClearAll makes sense
****************************************************************************/

BOOL CTraceList::canClearAll()
    {
     return GetCount() > 0;
    } // CTraceList::canClearAll

/****************************************************************************
*                            CTraceList::canDelete
* Result: BOOL
*       TRUE if something to delete
****************************************************************************/

BOOL CTraceList::canDelete()
    {
     return GetSelCount() > 0;
    } // CTraceList::canDelete

/****************************************************************************
*                              CTraceList::Copy
* Result: BOOL
*       TRUE if copy succeeded
*	FALSE if copy failed
* Effect: 
*       Copies selected items to the clipboard
****************************************************************************/

BOOL CTraceList::Copy()
    {
     CString s;
     int count = CListBox::GetSelCount();
     int * selections = new int[count];
     CListBox::GetSelItems(count, selections);
     for(int i = 0; i < count; i++)
	{ /* ¶ */
	 TraceEvent * e = (TraceEvent *)CListBox::GetItemDataPtr(selections[i]);
	 s += e->showfile();
	 s += "\r\n";
	} /* create string */
     delete [] selections;
     return sendToClipboard(s);
    } // CTraceList::Copy

/****************************************************************************
*                               CTraceList::Cut
* Result: BOOL
*       TRUE if cut succeeded
*	FALSE if cut failed
* Effect: 
*       Copies the elements to the clipboard and removes them
****************************************************************************/

BOOL CTraceList::Cut()
    {
     if(!Copy())
	return FALSE;
     return Delete();
    } // CTraceList::Cut

/****************************************************************************
*                             CTraceList::Delete
* Result: BOOL
*       TRUE if successful
*	FALSE if error
* Effect: 
*       Deletes the selected elements
****************************************************************************/

BOOL CTraceList::Delete()
    {
     int count = GetSelCount();
     if(count == 0)
	return FALSE;
     int  selection;
     BOOL result = FALSE;

     while(count > 0)
	{ /* delete one */
	 GetSelItems(1, &selection);
	 CListBox::DeleteString(selection);
	 count = GetSelCount();
	 result = TRUE;
	} /* delete one */
     if(result)
	{ /* notify */
	 recomputeHorizontalExtent();
	 GetParent()->SendMessage(UWM_TRACE_CHANGE);
	} /* notify */
     return result;
    } // CTraceList::Delete

BOOL CTraceList::PreCreateWindow(CREATESTRUCT& cs) 
{
 // Note that the styles must be as shown below (these are in the order of winuser.h)
 // LBS_NOTIFY            on
 // LBS_SORT              off
 // LBS_NOREDRAW          ---
 // LBS_MULTIPLESEL       on
 // LBS_ONWERDRAWFIXED    off
 // LBS_OWNERDRAWVARIABLE on
 // LBS_HASSTRINGS        off
 // LBS_USETABSTOPS       ---
 // LBS_NOINTEGRALHEIGHT  on
 // LBS_MULTICOLUMN       off
 // LBS_WANTKEYBOARDINPUT ---
 // LBS_EXTENDEDSEL       off
 // LBS_DISABLENOSCROLL	  ---
 // LBS_NODATA		  ---
 // LBS_NOSEL             off
	
 // Turn off disallowed styles
 cs.style &= ~(LBS_SORT |
	       LBS_HASSTRINGS |
	       LBS_OWNERDRAWFIXED |
	       LBS_EXTENDEDSEL |
	       LBS_NOSEL |
	       LBS_MULTICOLUMN);
 // Turn on mandatory styles
 cs.style |= (LBS_NOTIFY |
	      LBS_MULTIPLESEL |
	      LBS_OWNERDRAWVARIABLE |
	      LBS_NOINTEGRALHEIGHT);

 return CListBox::PreCreateWindow(cs);
}

/****************************************************************************
*                       CTraceList::clearAllSelections
* Result: void
*       
* Effect: 
*       Clears all the selections in the list box
****************************************************************************/

void CTraceList::clearAllSelections()
    {
     int count = CListBox::GetSelCount();
     while(count > 0)
	{ /* reset */
	 int selection;
	 CListBox::GetSelItems(1, &selection);
	 SetSel(selection, FALSE);
	 count = CListBox::GetSelCount();
	} /* reset */
    } // CTraceList::clearAllSelections

/****************************************************************************
*                    CTraceList::recomputeHorizontalExtent
* Result: void
*       
* Effect: 
*       Recomputes the horizontal extent
****************************************************************************/

void CTraceList::recomputeHorizontalExtent()
    {
     horizontalExtent = 0;
     CClientDC dc(this);
     for(int i = 0; i < CListBox::GetCount(); i++)
	{ /* scan */
	 TraceEvent * e = (TraceEvent *)GetItemDataPtr(i);
	 if(e == NULL)
	    continue;
	 CString s = e->show();
	 UINT w = dc.GetTextExtent(s).cx;
	 w += indent1 + indent2;
	 if(w > horizontalExtent)
	    horizontalExtent = w;
	} /* scan */
     SetHorizontalExtent(horizontalExtent);
    } // CTraceList::recomputeHorizontalExtent

/****************************************************************************
*                           CTraceList::setBoundary
* Result: void
*       
* Effect: 
*       Sets the high-water mark to be the current position
****************************************************************************/

void CTraceList::setBoundary()
    {
     boundary = AddString(new TraceComment(TraceEvent::None,
CString("================================================================")));
     if(limit > 0 && limit < boundary + MIN_LOG_ENTRIES)
	limit = boundary + MIN_LOG_ENTRIES;
    } // CTraceList::setBoundary

/****************************************************************************
*                             CTraceList::canSave
* Result: BOOL
*       TRUE if modified and nonempty
****************************************************************************/

BOOL CTraceList::canSave()
    {
     if(CListBox::GetCount() == 0)
	return FALSE;
     return modified;
    } // CTraceList::canSave

/****************************************************************************
*                            CTraceList::canSaveAs
* Result: BOOL
*       TRUE if a SaveAs is valid
****************************************************************************/

BOOL CTraceList::canSaveAs()
    {
     return CListBox::GetCount() > 0;
    } // CTraceList::canSaveAs

/****************************************************************************
*                         CTraceList::sendToClipboard
* Inputs:
*       const CString & s: String reference
* Result: BOOL
*       
* Effect: 
*       Puts the string on the clipboard
****************************************************************************/

BOOL CTraceList::sendToClipboard(const CString & s)
    {
     HGLOBAL mem = ::GlobalAlloc(GMEM_MOVEABLE, s.GetLength() + 1);
     if(mem == NULL)
	return FALSE;
     LPTSTR p = (LPTSTR)::GlobalLock(mem);
     lstrcpy(p, (LPCTSTR)s);
     ::GlobalUnlock(mem);

     if(!OpenClipboard())
	{ /* failed */
	 ::GlobalFree(mem);
	 return FALSE;
	} /* failed */
     EmptyClipboard();
     SetClipboardData(CF_TEXT, mem);
     CloseClipboard();
     return TRUE;
    } // CTraceList::sendToClipboard

/****************************************************************************
*                            CTraceList::canPrint
* Result: BOOL
*       TRUE if printing is possible
*	FALSE if nothing to print
****************************************************************************/

BOOL CTraceList::canPrint()
    {
     return GetCount() > 0;
    } // CTraceList::canPrint

/****************************************************************************
*                            CTraceList::printItem
* Inputs:
*       int i: Index of item to print
* Result: void
*       
* Effect: 
*       Prints the item. Handles page breaks, etc.
****************************************************************************/

void CTraceList::printItem(CDC & dc, int i)
    {
     int saved = dc.SaveDC();
     TraceEvent * e = (TraceEvent *)GetItemDataPtr(i);
     CString s = e->showfile();
     
     dc.SetTextColor(e->textcolor());
     printer->PrintLine(s);
     dc.RestoreDC(saved);
    } // CTraceList::printItem

/****************************************************************************
*                              CTraceList::Print
* Result: void
*       
* Effect: 
*       Prints the contents of the box. 
****************************************************************************/

void CTraceList::Print()
    {
     CPrintDialog dlg(FALSE,
		      PD_ALLPAGES |
		      PD_HIDEPRINTTOFILE |
		      PD_NOPAGENUMS |
		      PD_RETURNDC |
		      PD_USEDEVMODECOPIES);
     if(GetSelCount() < 2)
	{ /* enable selection */
	 dlg.m_pd.Flags |= PD_NOSELECTION;
	} /* enable selection */
     if(GetSelCount() > 10) // randomly-chosen number
	{ /* use selection */
	 dlg.m_pd.Flags |= PD_SELECTION;
	} /* use selection */
     switch(dlg.DoModal())
	{ /* DoModal */
	 case 0:
	 case IDCANCEL:
	    return;
	 case IDOK:
	    break;
	 default:
	    ASSERT(FALSE); // impossible condition
	    return;
	} /* DoModal */
     
     CDC dc;
     dc.Attach(dlg.m_pd.hDC);
     printer = new CPrinter(&dc);
     if(printer->StartPrinting())
	{ /* success */
	 int count = GetCount();

	 if(dlg.m_pd.Flags & PD_SELECTION)
	    { /* use selection */
	     int * selections = new int[count];
	     GetSelItems(count, selections);
	     for(int i = 0; i < count; i++)
		{ /* print each */
		 printItem(dc, selections[i]);
		} /* print each */
	     delete [] selections;
	    } /* use selection */
	 else
	    { /* print all */
	     for(int i = 0; i < count; i++)
		{ /* print line */
		 printItem(dc, i);
		} /* print line */

	     printer->EndPrinting();
	    } /* print all */
	} /* success */
     delete printer;
     ::DeleteDC(dc.Detach());
    } // CTraceList::Print

/****************************************************************************
*                          CTraceList::canSelectAll
* Result: BOOL
*       TRUE if anything to select
****************************************************************************/

BOOL CTraceList::canSelectAll()
    {
     return GetCount() > 0;
    } // CTraceList::canSelectAll

/****************************************************************************
*                            CTraceList::SelectAll
* Result: void
*       
* Effect: 
*       Selects all items
****************************************************************************/

void CTraceList::SelectAll()
    {
     SetRedraw(FALSE);
     for(int i = 0; i < GetCount(); i++)
	SetSel(i, TRUE);
     SetRedraw(TRUE);
    } // CTraceList::SelectAll
