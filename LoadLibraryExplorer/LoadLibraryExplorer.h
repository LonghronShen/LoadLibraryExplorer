// LoadLibraryExplorer.h : main header file for the PROJECT_NAME application
//

#pragma once

#ifndef __AFXWIN_H__
        #error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"           // main symbols


// CLoadLibraryExplorerApp:
// See LoadLibraryExplorer.cpp for the implementation of this class
//

class CLoadLibraryExplorerApp : public CWinApp
{
public:
        CLoadLibraryExplorerApp();

// Overrides
        public:
        virtual BOOL InitInstance();

// Implementation

        DECLARE_MESSAGE_MAP()
};

extern CLoadLibraryExplorerApp theApp;
