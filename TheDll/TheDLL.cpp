// TheDLL.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"
#include "TheDLL.h"
static HINSTANCE hInstance;
BOOL APIENTRY DllMain( HINSTANCE hInst, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
            hInstance = hInst;
            break;
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
	case DLL_PROCESS_DETACH:
		break;
	}
    return TRUE;
}


// This is an example of an exported function.
extern "C" THEDLL_API BOOL GetDLLName(LPTSTR buffer, DWORD length)
{
return ::GetModuleFileName(hInstance, buffer, length);
}

