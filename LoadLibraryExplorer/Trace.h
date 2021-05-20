#pragma once

// TraceList.h : header file
//

#include "TraceEvent.h"

/////////////////////////////////////////////////////////////////////////////
// CTraceList window

class CPrinter;

class CTraceList : public CListBox
{
// Construction
public:
        CTraceList();

// Attributes
public:
        int AddString(TraceEvent * event); 
        int DeleteString(int n);
        void ResetContent();
public:
        int InsertString(int index, TraceEvent * event)
              { modified = TRUE; return CListBox::InsertString(index, (LPCTSTR)event); }
// Operations
public:
        BOOL doSave(BOOL mode);
        BOOL toNext();
        BOOL toPrev();
        void toTop();
        void toEnd();
        BOOL canNext();
        BOOL canPrev();
        BOOL canTop();
        BOOL canEnd();
        BOOL Copy();
        BOOL Cut();
        BOOL Delete();
        void SelectAll();
        BOOL canCopy();
        BOOL canCut();
        BOOL canClearAll();
        BOOL canDelete();
        BOOL canSave();
        BOOL canSaveAs();
        BOOL canSelectAll();
        void setShowThreadID(BOOL mode) { showThreadID = mode; indent1 = 0; InvalidateRect(NULL); }
        void SetFont(CFont * font, BOOL redraw = TRUE)
           { indent1 = 0; CListBox::SetFont(font, redraw); }
        void setBoundary();
        BOOL canPrint();
        void Print();
// Overrides
        // ClassWizard generated virtual function overrides
        //{{AFX_VIRTUAL(CTraceList)
        public:
        virtual void MeasureItem(LPMEASUREITEMSTRUCT lpMeasureItemStruct);
        virtual void DrawItem(LPDRAWITEMSTRUCT lpDrawItemStruct);
        virtual void DeleteItem(LPDELETEITEMSTRUCT lpDeleteItemStruct);
        protected:
        virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
        //}}AFX_VIRTUAL

// Implementation
public:
        virtual ~CTraceList();
        CFont * fixed;
        int height;
        CString getFileName() { return SaveFileName; }
        void setLimit(int limit);
        void enableLimit(BOOL enable);
        BOOL setToDisk(BOOL v);
        // Generated message map functions
protected:
        CPrinter * printer;
        void printItem(CDC & dc, int i);
        BOOL sendToClipboard(const CString & s);
        void SetModified(BOOL mode = TRUE) { modified = mode; }
        BOOL GetModified() { return modified; }
        UINT deleted;
        BOOL showThreadID;
        void recomputeHorizontalExtent();
        int indent1;
        int indent2;
        void clearAllSelections();
        int findNext(int i);
        int findPrev(int i);
        UINT horizontalExtent;
        void appendToFile(TraceEvent * event);
        BOOL toDisk;
        BOOL isVisible(int n, BOOL fully);
        BOOL modified; 
        CString SaveFileName;
        int limit;
        BOOL limited;
        int boundary;

        //{{AFX_MSG(CTraceList)
        //}}AFX_MSG

        DECLARE_MESSAGE_MAP()
};

#define UWM_TRACE_CHANGE_MSG _T("UWM_TRACE_CHANGE-{44E678F1-DA79-11d3-9FE9-006067718D04}")

/////////////////////////////////////////////////////////////////////////////
