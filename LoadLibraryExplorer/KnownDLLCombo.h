#pragma once
#include "SmartDropdown.h"

// CKnownDLLCombo

class KnownDLLInfo {
    public:
       CString name;  // name of DLL
       BOOL preloaded;
       KnownDLLInfo(const CString & n, BOOL p) { name = n; preloaded = p; }
};

class CKnownDLLCombo : public CSmartDropdown
{
        DECLARE_DYNAMIC(CKnownDLLCombo)

public:
        CKnownDLLCombo();
        virtual ~CKnownDLLCombo();
        int AddString(KnownDLLInfo * info) { return CSmartDropdown::AddString((LPCTSTR)info); }
        KnownDLLInfo * GetItemDataPtr(int i) { return (KnownDLLInfo *)CSmartDropdown::GetItemDataPtr(i); }

protected:
        DECLARE_MESSAGE_MAP()
public:
    virtual void DrawItem(LPDRAWITEMSTRUCT /*lpDrawItemStruct*/);
    virtual void DeleteItem(LPDELETEITEMSTRUCT lpDeleteItemStruct);
    virtual int CompareItem(LPCOMPAREITEMSTRUCT /*lpCompareItemStruct*/);
};


