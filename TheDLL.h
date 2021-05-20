// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the THEDLL_EXPORTS
// symbol defined on the command line. this symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// THEDLL_API functions as being imported from a DLL, whereas this DLL sees symbols
// defined with this macro as being exported.
#ifdef THEDLL_EXPORTS
#define THEDLL_API __declspec(dllexport)
#else
#define THEDLL_API __declspec(dllimport)
#endif

// This class is exported from the TheDLL.dll

extern "C" THEDLL_API BOOL GetDLLName(LPTSTR buffer, DWORD length);


