// SafeDllTest.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "ErrorCodes.h"
#include "TheDll.h"
#include "CopyData.h"
#include "testguid.h"

/****************************************************************************
*                                   _tmain
* Inputs:
*       int argc: Count of args
*       _TCHAR * argv[]: Array of argument values
*               argv[0]: program name
*               argv[1]: handle of window to notify, as a hex number
* Result: int
*       0 if successful
*       error code if not
* Effect: 
*       Implicitly loads TheDll.Dll from the appropriate directory and
*       notifies the specified window about the actual path used
* Notes:
*       COPYDATASTRUCT(code, len, data)
*       code values determine length and data
*               SAFEDLL_NAME:   len is length of name in bytes (not chars)
*                               data is the actual name
*               SAFEDLL_ERROR:  len is 4
*                               data is ::GetLastError code
*       The STRINGTABLE resource maps the error numbers to strings
* 
*       Uses the method ::GetDLLName from the DLL to force the DLL to load, and
*       to retrieve its value
****************************************************************************/

int _tmain(int argc, _TCHAR* argv[])
   {
    if(argc < 2)
       { /* potential problem */
        TCHAR dll[MAX_PATH];
        ::GetDLLName(dll, MAX_PATH);
        _tprintf(dll);
        _tprintf(_T("\n"));
        return SAFEDLL_ERROR_INSUFFICIENT_PARAMETERS; 
       } /* potential problem */

    HWND h = (HWND)(DWORD_PTR)_tcstoui64(argv[1], NULL, 16); // convert to handle
    if(!::IsWindow(h))
       return SAFEDLL_ERROR_NOT_WINDOW;

    CopyData p(SAFEDLL_NAME, testguid);

    if(!::GetDLLName((LPTSTR)p.GetData(), MAX_PATH))
       { /* failed */
        DWORD err = ::GetLastError();
        CopyData e(SAFEDLL_ERROR, testguid);
        p.SetLength(sizeof(DWORD));
        p.SetData(&err, sizeof(DWORD));
        p.Send(h, NULL);
        return SAFEDLL_ERROR_BAD_GETDLLNAME;
       } /* failed */

    p.SetLength((DWORD)(DWORD_PTR) (_tcslen((LPTSTR)p.GetData()) + 1) * sizeof(TCHAR));
    p.Send(h, NULL);
    return 0;
   }

