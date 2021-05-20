/*****************************************************************************
  This software is deliverd "as is", and there is no warranty, express or
  implied, about its reliability, usablility, or fitness of purpose.  The author
  assumes no liability for any consequences of the use of this software, including
  lost time, lost data, financial impact, or anything else.  USE AT YOUR OWN RISK.

  This software makes changes in the state of the system.  YOU ARE RESPONSIBLE FOR
  SEEING THAT THE STATE OF YOUR SYSTEM IS IN THE STATE YOU DESIRE IT TO BE.  THE
  AUTHOR IS NOT RESPONSIBLE FOR CHANGES THAT ARE MADE WHICH MAY RESULT IN
  UNDESIRABLE CHANGES IN THE BEHAVIOR OF THE SYSTEM ON WHICH THIS SOFTWARE IS RUN.
*****************************************************************************/
  
// LoadLibraryExplorerDlg.cpp : implementation file
//

#include "stdafx.h"
#include "resource.h"
#include "LoadLibraryExplorerDlg.h"
#include "ErrorString.h"
#include "TraceEvent.h"
#include "Trace.h"
#include "RegVars.h"
#include "Registry.h"
#include "TheDLL.h"
#include "CopyData.h"
#include "ErrorCodes.h"
#include "testguid.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#define DLL_NAME _T("TheDLL.dll")
#define DLL_SOURCE _T("TheDLL.xxx")
#define TEST_PROGRAM_NAME _T("SafeDllTest.exe")

/****************************************************************************
*                                   UWM_LOG
* Inputs:
*       WPARAM: (WPARAM)(TraceEvent *) event to log
*       LPARAM: unused
* Result: LRESULT
*       Logically void, 0, always
* Effect: 
*       Logs the message
****************************************************************************/

static UINT UWM_LOG = ::RegisterWindowMessage(_T("UWM_LOG-{E4637467-B550-420d-BBF8-3C6F1E206226}"));


/****************************************************************************
*                           UWM_PROCESS_TERMINATED
* Inputs:
*       WPARAM: (WPARAM)(DWORD) exit code from process
*       LPARAM: (LPARAM)(HANDLE) process handle
* Result: LRESULT
*       Logically void, 0, always
* Effect: 
*       Notifies the recipient that the indicated process had terminated
****************************************************************************/

static UINT UWM_PROCESS_TERMINATED = ::RegisterWindowMessage(_T("UWM_PROCESS_TERMINATED-{12131336-39E9-427c-9017-5F5A4C76E6EF}"));


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
        CAboutDlg();

// Dialog Data
        enum { IDD = IDD_ABOUTBOX };

        protected:
        virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
        DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
        CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CLoadLibraryExplorerDlg dialog



/****************************************************************************
*              CLoadLibraryExplorerDlg::CLoadLibraryExplorerDlg
* Inputs:
*       CWNd * parent: Parent of dialog
* Effect: 
*       Constructor
****************************************************************************/

CLoadLibraryExplorerDlg::CLoadLibraryExplorerDlg(CWnd* pParent /*=NULL*/)
        : CDialog(CLoadLibraryExplorerDlg::IDD, pParent)
   {
    m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
    SetDllDirectory = NULL;
    launching = FALSE;
   }

/****************************************************************************
*                   CLoadLibraryExplorerDlg::DoDataExchange
* Inputs:
*       CDataExchange * pDX
* Result: void
*       
* Effect: 
*       Data exchange
****************************************************************************/

void CLoadLibraryExplorerDlg::DoDataExchange(CDataExchange* pDX)
{
CDialog::DoDataExchange(pDX);
DDX_Control(pDX, IDC_LOG, c_Log);
DDX_Control(pDX, IDC_SYSTEM32, c_SYSTEM32);
DDX_Control(pDX, IDC_WINDOWS, c_WINDOWS);
DDX_Control(pDX, IDC_WORKING, c_Working);
DDX_Control(pDX, IDC_APP_PATH, c_AppPath);
DDX_Control(pDX, IDC_PATH, c_Path);
DDX_Control(pDX, IDC_LIBRRARY, c_Library);
DDX_Control(pDX, IDC_LOADLIBRARY, c_LoadLibrary);
DDX_Control(pDX, IDC_LOAD_WITH_ALTERED_SEARCH_PATH, c_LOAD_WITH_ALTERED_SEARCH_PATH);
DDX_Control(pDX, IDC_USE_LOADLIBRARY, c_UseLoadLibrary);
DDX_Control(pDX, IDC_USE_LOADLIBRARYEX, c_UseLoadLibraryEx);
DDX_Control(pDX, IDC_USE_SHELLEXECUTE, c_UseShellExecute);
DDX_Control(pDX, IDC_EXECUTABLE, c_Executable);
DDX_Control(pDX, IDC_SYSTEM, c_SYSTEM);
DDX_Control(pDX, IDC_ALTERNATE, c_Alternate);
DDX_Control(pDX, IDC_FORCE_ALTERNATE, c_ForceAlternate);
DDX_Control(pDX, IDC_FORCE_WINDOWS, c_ForceWindows);
DDX_Control(pDX, IDC_FORCE_SYSTEM, c_ForceSystem);
DDX_Control(pDX, IDC_FORCE_WORKING, c_ForceWorking);
DDX_Control(pDX, IDC_FORCE_APP_PATH, c_ForceAppPath);
DDX_Control(pDX, IDC_FORCE_EXECUTABLE, c_ForceExecutable);
DDX_Control(pDX, IDC_FORCE, c_Force);
DDX_Control(pDX, IDC_FORCE_SYSTEM32, c_ForceSystem32);
DDX_Control(pDX, IDC_USE_SELECTED_SAFEDLLSEARCHMODE, c_UseSelectedSafeDllSearchMode);
DDX_Control(pDX, IDC_USE_DEFAULT_SEARCH_MODE, c_UseExistingSafeDllSearchMode);
DDX_Control(pDX, IDC_DESIRED_SAFEDLLSEARCHMODE, c_DesiredSafeDllSearchMode);
DDX_Control(pDX, IDC_DEFAULT_SAFE_DLL_SEARCH_MODE, c_DefaultSafeDllSearchMode);
DDX_Control(pDX, IDC_ALPHABETICAL, c_Alphabetical);
DDX_Control(pDX, IDC_RELAUNCH, c_Relaunch);
DDX_Control(pDX, IDC_MESSAGE, c_Message);
DDX_Control(pDX, IDC_STATE_SYSTEM32, c_StateSystem32);
DDX_Control(pDX, IDC_STATE_WINDOWS, c_StateWindows);
DDX_Control(pDX, IDC_STATE_SYSTEM, c_StateSystem);
DDX_Control(pDX, IDC_STATE_WORKING, c_StateWorking);
DDX_Control(pDX, IDC_STATE_APP_PATH, c_StateAppPath);
DDX_Control(pDX, IDC_STATE_EXECUTABLE, c_StateExecutable);
DDX_Control(pDX, IDC_STATE_ALTERNATE, c_StateAlternate);
DDX_Control(pDX, IDC_USE_THE_DLL, c_UseInternalDll);
DDX_Control(pDX, IDC_KNOWN_DLL, c_UseKnownDll);
DDX_Control(pDX, IDC_KNOWN_DLLS, c_KnownDlls);
DDX_Control(pDX, IDC_DLL_DIRECTORY, c_DllDirectory);
DDX_Control(pDX, IDC_REDIRECT_EXECUTABLE, c_RedirectExecutable);
DDX_Control(pDX, IDC_REDIRECT_LOCAL, c_RedirectDirectory);
DDX_Control(pDX, IDC_REDIRECT_NONE, c_RedirectNone);
DDX_Control(pDX, IDC_REDIRECT_CAPTION, c_RedirectCaption);
DDX_Control(pDX, IDC_STATE_LOCAL, c_StateLocal);
DDX_Control(pDX, IDC_LOCAL_REDIRECT, c_Local);
DDX_Control(pDX, IDC_DLL_LOCATIONS_CAPTION, c_DllLocations);
DDX_Control(pDX, IDC_MINFRAME, c_MinFrame);
DDX_Control(pDX, IDC_TEST_SAFE, c_TestSafeDllSearchMode);
DDX_Control(pDX, IDC_DIALOG_SELECTION_CAPTION, c_DialogSelectionCaption);
DDX_Control(pDX, IDC_LOAD_METHOD_CAPTION, c_LoadMethodCaption);
DDX_Control(pDX, IDC_LOAD, c_Load);
}

/****************************************************************************
*                                 Message Map
****************************************************************************/

BEGIN_MESSAGE_MAP(CLoadLibraryExplorerDlg, CDialog)
        ON_WM_SYSCOMMAND()
        ON_WM_PAINT()
        ON_WM_QUERYDRAGICON()
        ON_REGISTERED_MESSAGE(UWM_PROCESS_TERMINATED, OnProcessTerminated)
        ON_WM_CLOSE()
        ON_REGISTERED_MESSAGE(UWM_LOG, OnLog)
        ON_WM_SIZE()
        ON_WM_GETMINMAXINFO()
        ON_BN_CLICKED(IDC_SYSTEM32, OnBnClickedSystem32)
        ON_BN_CLICKED(IDC_WINDOWS, OnBnClickedWindows)
        ON_BN_CLICKED(IDC_WORKING, OnBnClickedWorking)
        ON_BN_CLICKED(IDC_APP_PATH, OnBnClickedAppPath)
        ON_BN_CLICKED(IDC_LOAD, OnBnClickedLoad)
        ON_BN_CLICKED(IDC_USE_LOADLIBRARY, OnBnClickedUseLoadlibrary)
        ON_BN_CLICKED(IDC_USE_LOADLIBRARYEX, OnBnClickedUseLoadlibraryex)
        ON_BN_CLICKED(IDC_LOAD_WITH_ALTERED_SEARCH_PATH, OnBnClickedLoadWithAlteredSearchPath)
        ON_BN_CLICKED(IDC_USE_SHELLEXECUTE, OnBnClickedUseShellexecute)
        ON_BN_CLICKED(IDC_EXECUTABLE, OnBnClickedExecutable)
        ON_BN_CLICKED(IDC_ALTERNATE, OnBnClickedAlternate)
        ON_BN_CLICKED(IDC_SYSTEM, OnBnClickedSystem)
        ON_BN_CLICKED(IDC_FORCE, OnBnClickedForce)
        ON_BN_CLICKED(IDC_FORCE_ALTERNATE, OnBnClickedForceAlternate)
        ON_BN_CLICKED(IDC_FORCE_APP_PATH, OnBnClickedForceAppPath)
        ON_BN_CLICKED(IDC_FORCE_EXECUTABLE, OnBnClickedForceExecutable)
        ON_BN_CLICKED(IDC_FORCE_SYSTEM, OnBnClickedForceSystem)
        ON_BN_CLICKED(IDC_FORCE_WINDOWS, OnBnClickedForceWindows)
        ON_BN_CLICKED(IDC_FORCE_WORKING, OnBnClickedForceWorking)
        ON_BN_CLICKED(IDC_FORCE_SYSTEM32, OnBnClickedForceSystem32)
        ON_BN_CLICKED(IDC_USE_SELECTED_SAFEDLLSEARCHMODE, OnBnClickedUseSelectedSafedllsearchmode)
        ON_BN_CLICKED(IDC_USE_DEFAULT_SEARCH_MODE, OnBnClickedUseDefaultSearchMode)
        ON_BN_CLICKED(IDC_DESIRED_SAFEDLLSEARCHMODE, OnBnClickedDesiredSafedllsearchmode)
        ON_BN_CLICKED(IDC_ENUM, OnBnClickedEnum)
        ON_BN_CLICKED(IDC_RELAUNCH, OnBnClickedRelaunch)
        ON_BN_CLICKED(IDC_USE_THE_DLL, OnBnClickedUseTheDll)
        ON_BN_CLICKED(IDC_KNOWN_DLL, OnBnClickedKnownDll)
        ON_CBN_SELENDOK(IDC_KNOWN_DLLS, OnCbnSelendokKnownDlls)
        ON_CBN_SELCHANGE(IDC_KNOWN_DLLS, OnCbnSelchangeKnownDlls)
        ON_BN_CLICKED(IDC_REDIRECT_NONE, OnBnClickedRedirectNone)
        ON_BN_CLICKED(IDC_REDIRECT_EXECUTABLE, OnBnClickedRedirectExecutable)
        ON_BN_CLICKED(IDC_REDIRECT_LOCAL, OnBnClickedRedirectLocal)
        ON_BN_CLICKED(IDC_TEST_SAFE, OnBnClickedTestSafeDllSearchMode)
        ON_WM_COPYDATA()
END_MESSAGE_MAP()


// CLoadLibraryExplorerDlg message handlers

/****************************************************************************
*                      CLoadLibraryExplorerDlg::MakeSubdirectory
* Inputs:
*       LPCTSTR name: Name of subdirectory to create
*       CString & dir: Name of directory created
* Result: BOOL
*       TRUE if successful
*       FALSE if error
* Effect: 
*       Creates a subdirectory of the executable image directory
* Notes:
*       dir is only defined if the return value is TRUE
****************************************************************************/

BOOL CLoadLibraryExplorerDlg::MakeSubdirectory(LPCTSTR name, CString & dir)
    {
     CString subdir = GetExeDir();
     AppendDir(subdir, name);

     if(!::CreateDirectory(subdir, NULL))
        { /* failed */
         DWORD err = ::GetLastError();
         if(err != ERROR_ALREADY_EXISTS)
            { /* some other failure */
             CString s = ::ErrorString(err);
             Alert();
             LogMessage(new TraceFormatError(TraceEvent::None, IDS_DIRECTORY_ERROR, subdir));
             LogMessage(new TraceFormatMessage(TraceEvent::None, err));
             dir = _T("");  // actually define it, to meaningless value
             return FALSE;
            } /* some other failure */
         else
            { /* already exists */
             LogMessage(new TraceFormatComment(TraceEvent::None, IDS_DIRECTORY_EXISTS, subdir));
             dir = subdir;
             return TRUE; // successful once
            } /* already exists */
        } /* failed */
     else
        { /* succeeded */
         LogMessage(new TraceFormatComment(TraceEvent::None, IDS_CREATED_DIRECTORY, subdir));
         dir = subdir;
         return TRUE;
        } /* succeeded */
    } // CLoadLibraryExplorerDlg::MakeSubdirectory

/****************************************************************************
*                            CLoadLibraryExplorerDlg::cwd
* Result: BOOL
*       TRUE if successful
*       FALSE if error
* Effect: 
*       Sets the current working directory to 'WorkingDir'
****************************************************************************/

BOOL CLoadLibraryExplorerDlg::cwd()
    {
     if(!SetCurrentDirectory(WorkingDir))
        { /* chdir failed */
         DWORD err = ::GetLastError();
         Alert();
         LogMessage(new TraceFormatError(TraceEvent::None, IDS_CD_FAILED, WorkingDir));
         LogMessage(new TraceFormatMessage(TraceEvent::None, err));
         return FALSE;
        } /* chdir failed */
     else
        { /* chdir worked */
         LogMessage(new TraceFormatComment(TraceEvent::None, IDS_CD_WORKED, WorkingDir));
         return TRUE;
        } /* chdir worked */
    } // CLoadLibraryExplorerDlg::cwd

/****************************************************************************
*                                  AppendDir
* Inputs:
*       CStrig & path: Path name
*       LPCTSTR name: Name to append to path
* Result: void
*       
* Effect: 
*       Modifies the path to have the name string appended
****************************************************************************/

void CLoadLibraryExplorerDlg::AppendDir(CString & path, LPCTSTR name)
    {
     if(path.Right(1) != _T("\\"))
        path += _T("\\");
     path += name;
    } // AppendDir

/****************************************************************************
*                          CLoadLibraryExplorerDlg::ShowPath
* Result: void
*       
* Effect: 
*       Shows the PATH variable
****************************************************************************/

void CLoadLibraryExplorerDlg::ShowPath()
    {
     DWORD len = GetEnvironmentVariable(_T("PATH"), NULL, 0);
     LPTSTR p = PathVariable.GetBuffer(len);
     GetEnvironmentVariable(_T("PATH"), p, len);
     PathVariable.ReleaseBuffer();

     c_Path.ResetContent();

     CString t = PathVariable;
     while(TRUE)
        { /* scan path */
         int n = t.Find(_T(';'));
         if(n < 0)
            break;
         CString s = t.Left(n);
         t = t.Mid(n+1);
         
         int item = c_Path.AddString(s);
         c_Path.SetCheck(item, BST_CHECKED);
        } /* scan path */

     if(!t.IsEmpty())
        { /* add last */
         int item = c_Path.AddString(t);
         c_Path.SetCheck(item, BST_CHECKED);
        } /* add last */
    } // CLoadLibraryExplorerDlg::ShowPath

/****************************************************************************
*                          CLoadLibraryExplorerDlg::SetPath
* Result: void
*       
* Effect: 
*       Sets the path based on the options selected by the user
****************************************************************************/

void CLoadLibraryExplorerDlg::SetPath()
    {
     CString path;
     CString sep = _T("");
     for(int i = 0; i < c_Path.GetCount(); i++)
        { /* build variable */
         if(c_Path.GetCheck(i))
            { /* active element */
             CString s;
             c_Path.GetText(i, s);
             path += sep + s;
             sep = _T(";");
            } /* active element */
        } /* build variable */
     SetEnvironmentVariable(_T("PATH"), path);
     LogMessage(new TraceFormatComment(TraceEvent::None, IDS_PATH_SET, path));
    } // CLoadLibraryExplorerDlg::SetPath

/****************************************************************************
*                   CLoadLibraryExplorerDlg::ClearLoadState
* Result: void
*       
* Effect: 
*       Sets the load state windows to empty
****************************************************************************/

void CLoadLibraryExplorerDlg::ClearLoadState()
    {
     c_StateWindows.SetWindowText(_T(""));
     c_StateSystem.SetWindowText(_T(""));
     c_StateSystem32.SetWindowText(_T(""));
     c_StateWorking.SetWindowText(_T(""));
     c_StateAppPath.SetWindowText(_T(""));
     c_StateExecutable.SetWindowText(_T(""));
     c_StateAlternate.SetWindowText(_T(""));
     c_StateLocal.SetWindowText(_T(""));
    } // CLoadLibraryExplorerDlg::ClearLoadState

/****************************************************************************
*                     CLoadLibraryExplorerDlg::ExtractDLL
* Inputs:
*       FILETIME ft: File time (may be <0,0> if error)
* Result: void
*       
* Effect: 
*       Extracts the DLL
****************************************************************************/

void CLoadLibraryExplorerDlg::ExtractDLL(FILETIME ft)
    {
     HRSRC hrsrc = ::FindResource(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_DLL), _T("DLL"));
     if(hrsrc == NULL)
        { /* we are in trouble */
         DWORD err = ::GetLastError();
         Alert();
         LogMessage(new TraceError(TraceEvent::None, IDS_DLL_RESOURCE_FAILURE));
         LogMessage(new TraceFormatMessage(TraceEvent::None, err));
         return;
        } /* we are in trouble */

     DWORD size = ::SizeofResource(AfxGetInstanceHandle(), hrsrc);
     if(size == 0)
        { /* really bad trouble */
         Alert();
         LogMessage(new TraceError(TraceEvent::None, IDS_DLL_RESOURCE_ZERO_SIZE));
         return;
        } /* really bad trouble */

     HGLOBAL res = ::LoadResource(AfxGetInstanceHandle(), hrsrc);
     if(res == NULL)
        { /* trouble so bad we cannot begin to imagine */
         DWORD err = ::GetLastError();
         Alert();
         LogMessage(new TraceError(TraceEvent::None, IDS_DLL_RESOURCE_LOAD_FAILED));
         return;
        } /* trouble so bad we cannot begin to imagine */

     LPVOID p = ::LockResource(res);
     if(p == NULL)
        { /* and we thought it was bad before this... */
         DWORD err = ::GetLastError();
         Alert();
         LogMessage(new TraceError(TraceEvent::None, IDS_DLL_LOCK_FAILED));
         return;
        } /* and we thought it was bad before this... */

     // Whew!!!  We made it, and we now have a bunch o' bits...

     CString target = GetExeDir();
     target += DLL_SOURCE; // create working copy
     HANDLE h = ::CreateFile(target,            // name
                             GENERIC_WRITE,     // writing it
                             0,                 // exclusive access
                             NULL,              // security irrelevant
                             CREATE_ALWAYS,     // create it new
                             FILE_ATTRIBUTE_NORMAL, // don't care
                             NULL);             // no template
     if(h == INVALID_HANDLE_VALUE)
        { /* after all that work, we're still hosed */
         DWORD err = ::GetLastError();
         Alert();
         LogMessage(new TraceFormatError(TraceEvent::None, IDS_DLL_CREATE_FAILED, target));
         LogMessage(new TraceFormatMessage(TraceEvent::None, err));
         return;
        } /* after all that work, we're still hosed */

     DWORD bytesWritten;
     BOOL ok = ::WriteFile(h, p, size, &bytesWritten, NULL);
     if(!ok)
        { /* we lose after all */
         DWORD err = ::GetLastError();
         Alert();
         LogMessage(new TraceFormatError(TraceEvent::None, IDS_ERROR_WRITING_DLL, target));
         LogMessage(new TraceFormatMessage(TraceEvent::None, err));
         ::CloseHandle(h);
         return;
        } /* we lose after all */

     // We have successfully written the file.  If the filetime is not 0, set it
     if(ft.dwLowDateTime != 0 && ft.dwHighDateTime != 0)
        { /* set timestamp */
         VERIFY(::SetFileTime(h, &ft, &ft, &ft));
        } /* set timestamp */
     ::CloseHandle(h);
     
     LogMessage(new TraceFormatComment(TraceEvent::None, IDS_DLL_EXTRACTED, target));
    } // CLoadLibraryExplorerDlg::ExtractDLL

/****************************************************************************
*                     CLoadLibraryExplorerDlg::MaybeExtractDLL
* Result: void
*       
* Effect: 
*       Attempts to extract the DLL from the resources, but only if
*       the existing file does not exist or has a date older than
*       the executable date
****************************************************************************/

void CLoadLibraryExplorerDlg::MaybeExtractDLL()
    {
     TCHAR self[MAX_PATH];
     ::GetModuleFileName(NULL, self, MAX_PATH);
     HANDLE h = ::CreateFile(self,
                             GENERIC_READ,
                             FILE_SHARE_READ, // allo reading
                             NULL, // security irrelevant
                             OPEN_EXISTING, // must exist
                             FILE_ATTRIBUTE_NORMAL, // don't care
                             NULL); // no template
     if(h == INVALID_HANDLE_VALUE)
        { /* force loading */
         DWORD err = ::GetLastError();
         LogMessage(new TraceFormatError(TraceEvent::None, IDS_ERROR_GETTING_EXE_DATE, self));
         LogMessage(new TraceFormatMessage(TraceEvent::None, err));
         FILETIME f={0};
         ExtractDLL(f);
         return;
        } /* force loading */

     FILETIME exe_created;
     FILETIME exe_accessed;
     FILETIME exe_modified;
     VERIFY(::GetFileTime(h, &exe_created, &exe_accessed, &exe_modified));
     ::CloseHandle(h);


     CString target = GetSourceFilename(UseInternalDll);
     h = ::CreateFile(target,
                      GENERIC_READ, // required for GetFileTime
                      0, // share exclusive
                      NULL, // security irrelevant
                      OPEN_EXISTING, // must exist
                      FILE_ATTRIBUTE_NORMAL, // don't care
                      NULL); // no template
     if(h == INVALID_HANDLE_VALUE)
        { /* does not exist */
         LogMessage(new TraceFormatComment(TraceEvent::None, IDS_NO_DLL_FOUND, target));
         ExtractDLL(exe_created);
         return;
        } /* does not exist */

     // It exists; check its "file time"
     FILETIME target_created;
     FILETIME target_accessed;
     FILETIME target_modified;

     VERIFY(::GetFileTime(h, &target_created, &target_accessed, &target_modified));
     ::CloseHandle(h);

     if(target_created.dwLowDateTime == exe_created.dwLowDateTime &&
        target_created.dwHighDateTime == exe_created.dwHighDateTime)
        { /* same */
         LogMessage(new TraceFormatComment(TraceEvent::None, IDS_DLL_CURRENT, target));
         return;
        } /* same */

     ExtractDLL(exe_created);
    } // CLoadLibraryExplorerDlg::MaybeExtractDLL

/****************************************************************************
*                        CLoadLibraryExplorerDlg::OnInitDialog
* Result: BOOL
*       TRUE, always
* Effect: 
*       Initialization
****************************************************************************/

BOOL CLoadLibraryExplorerDlg::OnInitDialog()
   {
    CDialog::OnInitDialog();

    // Add "About..." menu item to system menu.

    // IDM_ABOUTBOX must be in the system command range.
    ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
    ASSERT(IDM_ABOUTBOX < 0xF000);
    
    CMenu* pSysMenu = GetSystemMenu(FALSE);
    if (pSysMenu != NULL)
       {
        CString strAboutMenu;
        strAboutMenu.LoadString(IDS_ABOUTBOX);
        if (!strAboutMenu.IsEmpty())
           {
            pSysMenu->AppendMenu(MF_SEPARATOR);
            pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
           }
       }

    // Set the icon for this dialog.  The framework does this automatically
    //  when the application's main window is not a dialog
    SetIcon(m_hIcon, TRUE);                     // Set big icon
    SetIcon(m_hIcon, FALSE);            // Set small icon

    //****************************************************************
    // Extract the DLL from the .exe file if it is not present or
    // its data is older than the .exe file
    //****************************************************************

    MaybeExtractDLL();
    
    RegistryInt relaunch(IDS_REGISTRY_RELAUNCH);
    relaunch.load(FALSE);

#define RESTART_OPTION _T("/restart")

    for(int i = 0; i < __argc; i++)
       { /* see if explicit restart */
        if(__argv[i] == CString(RESTART_OPTION))
           { /* simulate restart */
            relaunch.value = TRUE;
            break;
           } /* simulate restart */
       } /* see if explicit restart */
#ifdef _DEBUG
    if(relaunch.value)
       DebugBreak();
#endif

    //****************
    // SetDllDirectory
    //****************
    HMODULE kernel32 = ::GetModuleHandle(_T("KERNEL32"));
    ASSERT(kernel32 != NULL);
#ifdef _UNICODE
    SetDllDirectory =(SetDllDirectoryProc)::GetProcAddress(kernel32, "SetDllDirectoryW");
#else
    SetDllDirectory = (SetDllDirectoryProc)::GetProcAddress(kernel32, "SetDllDirectoryA");
#endif
    // This returns NULL for any OS < XP SP1

    // Create the test DLL directory

    //****************
    // TestDir
    //****************
    MakeSubdirectory(_T("Test"), TestDir);

    //****************
    // WorkingDir
    //****************
    MakeSubdirectory(_T("Working"), WorkingDir);

    //****************
    // AppPathDir
    //****************
    MakeSubdirectory(_T("AppPath"), AppPathDir);

    //****************
    // AlternateDir
    //****************
    MakeSubdirectory(_T("Alternate"), AlternateDir);

    //****************
    // System32Dir
    //****************
    TCHAR path[MAX_PATH];
    
    ::GetSystemDirectory(path, MAX_PATH);
    System32Dir = path;
    c_SYSTEM32.SetWindowText(path);

    RegistryInt Sys32(IDS_REGISTRY_SYSTEM32);
    Sys32.load();
    c_SYSTEM32.SetCheck(Sys32.value ? BST_CHECKED : BST_UNCHECKED);

    //****************
    // WindowsDir
    //****************
    ::GetWindowsDirectory(path, MAX_PATH);
    WindowsDir = path;
    c_WINDOWS.SetWindowText(path);
    
    RegistryInt Windows(IDS_REGISTRY_WINDOWS);
    Windows.load();
    c_WINDOWS.SetCheck(Windows.value ? BST_CHECKED : BST_UNCHECKED);

    //****************
    // SystemDir
    //****************
    SystemDir = WindowsDir;
    AppendDir(SystemDir, _T("System"));
    c_SYSTEM.SetWindowText(SystemDir);
    
    RegistryInt System(IDS_REGISTRY_SYSTEM);
    System.load();
    c_SYSTEM.SetCheck(System.value ? BST_CHECKED : BST_UNCHECKED);

    //****************
    // WorkingDir
    //****************
    c_Working.SetWindowText(WorkingDir);
    RegistryInt Working(IDS_REGISTRY_WORKING);
    Working.load();
    c_Working.SetCheck(Working.value ? BST_CHECKED : BST_UNCHECKED);
    c_Working.EnableWindow(cwd());

    //****************
    // Executable
    //****************
    CString dir = GetExeDir();
    if(dir.Right(1) == _T("\\"))
       dir = dir.Left(dir.GetLength() - 1);
    c_Executable.SetWindowText(dir);
    RegistryInt Executable(IDS_REGISTRY_EXECUTABLE);
    Executable.load();
    c_Executable.SetCheck(Executable.value ? BST_CHECKED : BST_UNCHECKED);

    //****************
    // Alternate
    //****************
    RegistryInt Alternate(IDS_REGISTRY_ALTERNATE);
    Alternate.load();
    c_Alternate.SetCheck(Alternate.value ? BST_CHECKED : BST_UNCHECKED);
    c_Alternate.SetWindowText(AlternateDir);

    //****************
    // App Path
    //****************
    if(!AppPathDir.IsEmpty())
       { /* set AppPathDir */
        c_AppPath.SetWindowText(AppPathDir);
        RegistryInt AppPath(IDS_REGISTRY_APP_PATH);
        AppPath.load();
        c_AppPath.SetCheck(AppPath.value ? BST_CHECKED : BST_UNCHECKED);
       } /* set AppPathDir */
    else
       { /* no AppPathDir */
        c_AppPath.EnableWindow(FALSE);
       } /* no AppPathDir */
    
    //********************************
    // SafeDllSearchMode
    // Default SafeDllSearchMode
    //********************************
    //  (*) Force specific setting
    //             [x] Desired SafeDllSearchMode setting
    //  ( ) Use default setting  [ ]
    //----------------------------------------------------------------             
    //  (*) IDC_USE_SELECTED_SAFEDLLSEARCHMODE
    //             [x] IDC_DESIRED_SAFEDLLSEARCHMODE
    //  ( ) IDC_USE_DEFAULT_SEARCH_MODE [ ] IDC_DEFAULT_SAFE_DLL_SEARCH_MODE
    //----------------------------------------------------------------
    //  (*) c_UseSelectedSafeDllSearchMode
    //             [x] c_DesiredSafeDllSearchMode
    //  ( ) c_UseExistingSafeDllSearchMode); [ ] c_DefaultSafeDllSearchMode
    //----------------------------------------------------------------
    //  (*) IDS_REGISTRY_USE_SAFEDLLSEARCHMODE <-------------------------+
    //             [x] IDS_REGISTRY_USE_SAFEDLLSEARCHMODE_SETTING        +----- saves ID value
    //  ( ) IDS_REGISTRY_USE_SAFEDLLSEARCHMODE <-------------------------+
    //----------------------------------------------------------------
    RegistryInt TestSafe(IDS_REGISTRY_TEST_SAFE_DLL_SEARCH_MODE);
    TestSafe.load();
    c_TestSafeDllSearchMode.SetCheck(TestSafe.value ? BST_CHECKED : BST_UNCHECKED);

    RegistryInt SafeDllSearchMode(IDS_REGISTRY_SAFEDLLSEARCHMODE_HKLM, HKEY_LOCAL_MACHINE); // HKLM\System\CurrentControlSet\Control\Session Manager\SafeDllSearchMode
    SafeDllSearchMode.load();
    c_DefaultSafeDllSearchMode.SetCheck(SafeDllSearchMode.value ? BST_CHECKED : BST_UNCHECKED);

    ASSERT(IDC_USE_DEFAULT_SEARCH_MODE - IDC_USE_SELECTED_SAFEDLLSEARCHMODE == 1);

    RegistryInt SelectSafeDllSearchMode(IDS_REGISTRY_USE_SAFEDLLSEARCHMODE); 
    SelectSafeDllSearchMode.load(IDC_USE_DEFAULT_SEARCH_MODE);
    UINT mode;
    switch(SelectSafeDllSearchMode.value)
       { /* validate value */
        case SearchMode::UseSelectedSafeDllSearchMode:
              mode = IDC_USE_SELECTED_SAFEDLLSEARCHMODE;
              break;
        case SearchMode::UseDefaultSafeDllSearchMode:
              mode = IDC_USE_DEFAULT_SEARCH_MODE;
              break;
        default:
              //ASSERT(FALSE);
              mode = IDC_USE_DEFAULT_SEARCH_MODE;
              break;
       } /* validate value */
    CheckRadioButton(IDC_USE_SELECTED_SAFEDLLSEARCHMODE, IDC_USE_DEFAULT_SEARCH_MODE, mode);

    RegistryInt SafeDllSearchModeSetting(IDS_REGISTRY_USE_SAFEDLLSEARCHMODE_SETTING);
    SafeDllSearchModeSetting.load();
    c_DesiredSafeDllSearchMode.SetCheck(SafeDllSearchModeSetting.value ? BST_CHECKED : BST_UNCHECKED);

    //****************
    // Redirection
    //****************
    RegistryInt Redirect(IDS_REGISTRY_REDIRECTION);
    Redirect.load(Redirection::None);
    UINT redir;
    switch(Redirect.value)
       { /* redirect */
        case Redirection::None:
              redir = IDC_REDIRECT_NONE;
              break;
        case Redirection::File:
              redir = IDC_REDIRECT_EXECUTABLE;
              break;
        case Redirection::Directory:
              redir = IDC_REDIRECT_LOCAL;
              break;
        default:
              ASSERT(FALSE);
              break;
       } /* redirect */
    CheckRadioButton(IDC_REDIRECT_NONE, IDC_REDIRECT_LOCAL, redir);

    //****************
    // Dll type
    //****************
    LoadKnownDlls();

    RegistryInt DllSelection(IDS_REGISTRY_DLL_SELECTION);
    DllSelection.load(IDC_USE_THE_DLL);
    CheckRadioButton(IDC_USE_THE_DLL, IDC_KNOWN_DLL, DllSelection.value);
    SetCurrentMode();

    //****************
    // PATH=
    //****************
    ShowPath();

    //****************
    // Load type
    //****************
    UINT loadtype = IDC_USE_LOADLIBRARY;
    
    RegistryInt LoadType(IDS_REGISTRY_LOADTYPE);
    LoadType.load((int)Load::LoadLibrary);
    switch((Load::LoadType)LoadType.value)
       { /* Load Type */
        case Load::LoadLibrary:
        case 0:
              loadtype = IDC_USE_LOADLIBRARY;
              break;
        case Load::LoadLibraryEx:
              loadtype = IDC_USE_LOADLIBRARYEX;
              break;
       } /* Load Type */

    CheckRadioButton(IDC_USE_LOADLIBRARY, IDC_USE_LOADLIBRARYEX, loadtype);

    //********************************
    // LOAD_WITH_ALTERED_SEARCH_PATH
    //********************************
    RegistryInt AlteredPath(IDS_REGISTRY_LOAD_WITH_ALTERED_SEARCH_PATH);
    AlteredPath.load();
    c_LOAD_WITH_ALTERED_SEARCH_PATH.SetCheck(AlteredPath.value ? BST_CHECKED : BST_UNCHECKED);


    //********************************
    // Forced Selection Type
    //********************************

    RegistryInt Force(IDS_REGISTRY_FORCE);
    Force.load();
    c_Force.SetCheck(Force.value ? BST_CHECKED : BST_UNCHECKED);

    RegistryInt ForceSelection(IDS_REGISTRY_FORCE_SELECTION);
    ForceSelection.load();
    UINT selection;

    switch((Force::ForceType)ForceSelection.value)
       { /* Force */
        case Force::Alternate:
              selection = IDC_FORCE_ALTERNATE;
              break;
        case Force::AppPath:
              selection = IDC_FORCE_APP_PATH;
              break;
        case 0: // default if not set
        case Force::Executable:
        default:
              selection = IDC_FORCE_EXECUTABLE;
              break;
        case Force::System:
              selection = IDC_FORCE_SYSTEM;
              break;
        case Force::System32:
              selection = IDC_FORCE_SYSTEM32;
              break;
        case Force::Windows:
              selection = IDC_FORCE_WINDOWS;
              break;
        case Force::Working:
              selection = IDC_FORCE_WORKING;
              break;
       } /* Force */

    //************************************************
    // Check the initial Registry App Path state
    //************************************************

    TCHAR self[MAX_PATH];
    ::GetModuleFileName(NULL, self, MAX_PATH);
    CString AppPathKey = GetRegistryAppPathKey(self);
    CString apppath;
    
    if(!GetRegistryString(HKEY_LOCAL_MACHINE, AppPathKey, _T("Path"), apppath))
       { /* failed to get it */
        LogMessage(new TraceFormatComment(TraceEvent::None, IDS_NO_INITIAL_APP_PATH, AppPathKey));
        HasAppPath = FALSE;
       } /* failed to get it */
    else
       { /* got it */
        LogMessage(new TraceFormatComment(TraceEvent::None, IDS_INITIAL_APP_PATH, AppPathKey, apppath));
        HasAppPath = !apppath.IsEmpty();
       } /* got it */
    
    //********************************
    // Force the registry path to be
    // consistent with the option
    // setting
    //********************************
    SetAllRegistryAppPaths(); // set current values


    if(relaunch.value)
       { /* is relaunch */
        RegistryWindowPlacement wp(IDS_REGISTRY_WINDOWPLACEMENT);
   
        if(wp.load())
           { /* found old placement */
            SetWindowPlacement(&wp.value);
           } /* found old placement */
        relaunch.value = FALSE;
        relaunch.store();
       } /* is relaunch */

    CheckRadioButton(IDC_FORCE_SYSTEM32, IDC_FORCE_ALTERNATE, selection);
    
    ClearLoadState();

    //*****************************************************************************
    updateControls();

    return TRUE;  // return TRUE  unless you set the focus to a control
   }

/****************************************************************************
*                        CLoadLibraryExplorerDlg::SaveOptions
* Result: void
*       
* Effect: 
*       Saves the current options
****************************************************************************/

void CLoadLibraryExplorerDlg::SaveOptions()
   {
    //----------------
    // SYSTEM32
    //----------------
    RegistryInt Sys32(IDS_REGISTRY_SYSTEM32);
    Sys32.value = c_SYSTEM32.GetCheck() == BST_CHECKED;
    Sys32.store();

    //----------------
    // Windows
    //----------------
    RegistryInt Windows(IDS_REGISTRY_WINDOWS);
    Windows.value = c_WINDOWS.GetCheck() == BST_CHECKED;
    Windows.store();

    //----------------
    // System
    //----------------
    RegistryInt System(IDS_REGISTRY_SYSTEM);
    System.value = c_SYSTEM.GetCheck() == BST_CHECKED;
    System.store();

    //----------------
    // Working
    //----------------
    RegistryInt Working(IDS_REGISTRY_WORKING);
    Working.value = c_Working.GetCheck() == BST_CHECKED;
    Working.store();
    
    //----------------
    // AppPath
    //----------------
    
    RegistryInt AppPath(IDS_REGISTRY_APP_PATH);
    AppPath.value = c_AppPath.GetCheck() == BST_CHECKED;
    AppPath.store();
    
    //----------------
    // Executable
    //----------------

    RegistryInt Executable(IDS_REGISTRY_EXECUTABLE);
    Executable.value = c_Executable.GetCheck() == BST_CHECKED;
    Executable.store();

    //----------------
    // Alternate
    //----------------
    RegistryInt Alternate(IDS_REGISTRY_ALTERNATE);
    Alternate.value = c_Alternate.GetCheck() == BST_CHECKED;
    Alternate.store();
    //----------------
    // Force
    //----------------
    RegistryInt Force(IDS_REGISTRY_FORCE);
    Force.value = c_Force.GetCheck() == BST_CHECKED;
    Force.store();

    RegistryInt ForceSelection(IDS_REGISTRY_FORCE_SELECTION);

    //****************
    // Force System32
    //****************
    if(c_ForceSystem32.GetCheck() == BST_CHECKED)
       ForceSelection.value = (int)Force::System32;
    else
    //****************
    // Force Windows
    //****************
    if(c_ForceWindows.GetCheck() == BST_CHECKED)
       ForceSelection.value = (int)Force::Windows;
    else
    //****************
    // Force System
    //****************
    if(c_ForceSystem.GetCheck() == BST_CHECKED)
       ForceSelection.value = (int)Force::System;
    else
    //****************
    // Working
    //****************
    if(c_ForceWorking.GetCheck() == BST_CHECKED)
       ForceSelection.value = (int)Force::Working;
    else
    //****************
    // Force AppPath
    //****************
    if(c_ForceAppPath.GetCheck() == BST_CHECKED)
       ForceSelection.value = (int)Force::AppPath;
    else
    //****************
    // Force Executable
    //****************
    if(c_ForceExecutable.GetCheck() == BST_CHECKED)
       ForceSelection.value = (int)Force::Executable;
    else
    //****************
    // Force Alternate
    //****************
    if(c_ForceAlternate.GetCheck() == BST_CHECKED)
       ForceSelection.value = (int)Force::Alternate;
    else
       { /* unknown */
        ASSERT(FALSE);
        ForceSelection.value= Force::Executable;
       } /* unknown */
    
    ForceSelection.store();
    
    //********************************
    //LOAD_WITH_ALTERED_SEARCH_PATH
    //********************************
    RegistryInt AlteredPath(IDS_REGISTRY_LOAD_WITH_ALTERED_SEARCH_PATH);
    AlteredPath.value = c_LOAD_WITH_ALTERED_SEARCH_PATH.GetCheck();
    AlteredPath.store();


    //****************
    // LoadType
    //****************
    RegistryInt LoadType(IDS_REGISTRY_LOADTYPE);
    if(c_UseLoadLibrary.GetCheck() == BST_CHECKED)
       LoadType.value = (int)Load::LoadLibrary;
    else
    if(c_UseLoadLibraryEx.GetCheck() == BST_CHECKED)
       LoadType.value = (int)Load::LoadLibraryEx;

    LoadType.store();
    
    //****************
    // Redirect
    //****************
    RegistryInt Redirect(IDS_REGISTRY_REDIRECTION);
    if(c_RedirectNone.GetCheck() == BST_CHECKED)
       Redirect.value = Redirection::None;
    else
    if(c_RedirectExecutable.GetCheck() == BST_CHECKED)
       Redirect.value = Redirection::File;
    else
    if(c_RedirectDirectory.GetCheck() == BST_CHECKED)
       Redirect.value = Redirection::Directory;
    else
       { /* unknown */
        ASSERT(FALSE);
        Redirect.value = Redirection::None;
       } /* unknown */

    Redirect.store();

    //****************
    // SafeDllSearchMode
    //****************
    // [x] Test SafeDllSearchMode
    //
    //    (*) Force specific setting
    //             [x] Desired SafeDllSearchMode setting
    //    ( ) Use default setting  [ ]
    //----------------------------------------------------------------             
    // [x] IDC_TEST_SAFE
    //
    //    (*) IDC_USE_SELECTED_SAFEDLLSEARCHMODE
    //             [x] IDC_DESIRED_SAFEDLLSEARCHMODE
    //    ( ) IDC_USE_DEFAULT_SEARCH_MODE [ ] IDC_DEFAULT_SAFE_DLL_SEARCH_MODE
    //----------------------------------------------------------------
    // [x] c_TestSafeDllSearchMode
    //
    //    (*) c_UseSelectedSafeDllSearchMode
    //             [x] c_DesiredSafeDllSearchMode
    //    ( ) c_UseExistingSafeDllSearchMode); [ ] c_DefaultSafeDllSearchMode
    //----------------------------------------------------------------
    // [x] IDS_REGISTRY_REGISTRY_TEST_SAFE_DLL_SEARCH_MODE
    //c
    //    (*) IDS_REGISTRY_USE_SAFEDLLSEARCHMODE  <-------------------------+
    //             [x] IDS_REGISTRY_USE_SAFEDLLSEARCHMODE_SETTING         +------ saves ID value
    //    ( ) IDS_REGISTRY_USE_SAFEDLLSEARCHMODE  <-------------------------+
    //----------------------------------------------------------------
    RegistryInt SelectSafeDllSearchMode(IDS_REGISTRY_USE_SAFEDLLSEARCHMODE);
    SelectSafeDllSearchMode.value = c_UseSelectedSafeDllSearchMode.GetCheck() == BST_CHECKED ? SearchMode::UseSelectedSafeDllSearchMode : SearchMode::UseDefaultSafeDllSearchMode;
    SelectSafeDllSearchMode.store();

    RegistryInt SafeDllSearchModeSetting(IDS_REGISTRY_USE_SAFEDLLSEARCHMODE_SETTING);
    SafeDllSearchModeSetting.value = c_DesiredSafeDllSearchMode.GetCheck() == BST_CHECKED;
    SafeDllSearchModeSetting.store();

   } // CLoadLibraryExplorerDlg::SaveOptions

/****************************************************************************
*                    CLoadLibraryExplorerDlg::OnSysCommand
* Inputs:
*       UINT nID: ID of system menu item
*       LPARAM lParam:
* Result: void
*       
* Effect: 
*       If it is the About... item, launch the About box, otherwise pass
*       it on...
****************************************************************************/

void CLoadLibraryExplorerDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
        if ((nID & 0xFFF0) == IDM_ABOUTBOX)
        {
                CAboutDlg dlgAbout;
                dlgAbout.DoModal();
        }
        else
        {
                CDialog::OnSysCommand(nID, lParam);
        }
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CLoadLibraryExplorerDlg::OnPaint() 
{
        if (IsIconic())
        {
                CPaintDC dc(this); // device context for painting

                SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

                // Center icon in client rectangle
                int cxIcon = GetSystemMetrics(SM_CXICON);
                int cyIcon = GetSystemMetrics(SM_CYICON);
                CRect rect;
                GetClientRect(&rect);
                int x = (rect.Width() - cxIcon + 1) / 2;
                int y = (rect.Height() - cyIcon + 1) / 2;

                // Draw the icon
                dc.DrawIcon(x, y, m_hIcon);
        }
        else
        {
                CDialog::OnPaint();
        }
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CLoadLibraryExplorerDlg::OnQueryDragIcon()
{
        return static_cast<HCURSOR>(m_hIcon);
}

void CLoadLibraryExplorerDlg::OnClose()
    {
     EraseDLLs(CurrentMode);
     CDialog::OnOK();
    }

void CLoadLibraryExplorerDlg::OnOK()
    {
    }

void CLoadLibraryExplorerDlg::OnCancel()
    {
    }

/****************************************************************************
*                         CLoadLibraryExplorerDlg::LogMessage
* Inputs:
*       TraceEvent * e: Trace event to log
* Result: void
*       
* Effect: 
*       Logs the event asynchronously
****************************************************************************/

void CLoadLibraryExplorerDlg::LogMessage(TraceEvent * e)
    {
     PostMessage(UWM_LOG, (WPARAM)e);
    } // CLoadLibraryExplorerDlg::LogMessage

/****************************************************************************
*                           CLoadLibraryExplorerDlg::OnLog
* Inputs:
*       WPARAM: (WPARAM)(TraceEvent *): Event to log
*       LPARAM: Ignored
* Result: LRESULT
*       Logically void, 0, always
* Effect: 
*       Logs the message
****************************************************************************/

LRESULT CLoadLibraryExplorerDlg::OnLog(WPARAM wParam, LPARAM)
    {
     TraceEvent * e = (TraceEvent *)wParam;
     c_Log.AddString(e);
     return 0;
    } // CLoadLibraryExplorerDlg::OnLog

/****************************************************************************
*                           CLoadLibraryExplorerDlg::Resize
* Inputs:
*       CWnd & ctl: Control object to resize
*       BOOL toBottom: TRUE if it should extend vertically to the bottom
*       int adjust: Right-margin adjustment
* Result: void
*       
* Effect: 
*       Resizes a window to extend to the far right
****************************************************************************/

void CLoadLibraryExplorerDlg::Resize(CWnd & ctl, BOOL toBottom /*=FALSE*/, int adjust /*=0*/)
    {
     if(ctl.GetSafeHwnd() != NULL)
        { /* resize it */
         CRect client;
         GetClientRect(&client);

         CRect r;
         ctl.GetWindowRect(&r);
         ScreenToClient(&r);
         int height;
         if(toBottom)
            height = client.Height() - r.top;
         else
            height = r.Height();

         ctl.SetWindowPos(NULL, 0, 0, client.Width() - r.left - adjust , height, SWP_NOMOVE | SWP_NOZORDER);
        } /* resize it */
    } // CLoadLibraryExplorerDlg::Resize

/****************************************************************************
*                           CLoadLibraryExplorerDlg::OnSize
* Inputs:
*       UINT nType: type of resize
*       int cx: New width
*       int cy: New height
* Result: void
*       
* Effect: 
*       Handles resizing of controls within the dialog
****************************************************************************/

void CLoadLibraryExplorerDlg::OnSize(UINT nType, int cx, int cy)
    {
     CDialog::OnSize(nType, cx, cy);

     Resize(c_Log, TRUE);
     Resize(c_Path);
     Resize(c_Library);
     Resize(c_LoadLibrary);
     Resize(c_DllLocations);
     int gap = 4 * ::GetSystemMetrics(SM_CXEDGE);
     Resize(c_SYSTEM32, FALSE, gap);
     Resize(c_WINDOWS, FALSE, gap);
     Resize(c_SYSTEM, FALSE, gap);
     Resize(c_Working, FALSE, gap);
     Resize(c_AppPath, FALSE, gap);
     Resize(c_Executable, FALSE, gap);
     Resize(c_Alternate, FALSE, gap);
     Resize(c_Local, FALSE, gap);
    }

/****************************************************************************
*                      CLoadLibraryExplorerDlg::OnGetMinMaxInfo
* Inputs:
*       MINMAXINFO * lpMMI:
* Result: void
*       
* Effect: 
*       Limits the resize
****************************************************************************/

void CLoadLibraryExplorerDlg::OnGetMinMaxInfo(MINMAXINFO* lpMMI)
    {
     CDialog::OnGetMinMaxInfo(lpMMI);

     if(c_MinFrame.GetSafeHwnd() != NULL)
        { /* can limit size */
         CRect r;
         c_MinFrame.GetWindowRect(&r);
         ScreenToClient(&r);
         CalcWindowRect(&r);

         lpMMI->ptMinTrackSize.x = r.Width();
         lpMMI->ptMinTrackSize.y = r.Height();
        } /* can limit size */
    }

void CLoadLibraryExplorerDlg::OnBnClickedSystem32()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedWindows()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedWorking()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedAppPath()
    {
     SetAllRegistryAppPaths(); // set current values
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedExecutable()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedAlternate()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedSystem()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedForce()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedForceAlternate()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedForceAppPath()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedForceExecutable()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedForceSystem()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedForceWindows()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedForceWorking()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedForceSystem32()
   {
    SaveOptions();
    updateControls();
   }

/****************************************************************************
*                         CLoadLibraryExplorerDlg::GetExeDir
* Result: CString 
*       The path of the executable directory
****************************************************************************/

CString CLoadLibraryExplorerDlg::GetExeDir()
    {
     TCHAR exepath[MAX_PATH];
     ::GetModuleFileName(NULL, exepath, MAX_PATH);

    // Now create a new directory, unless it already exists

     TCHAR drive[MAX_PATH];
     TCHAR path[MAX_PATH];
     _tsplitpath(exepath, drive, path, NULL, NULL);
     _tmakepath(exepath, drive, path, NULL, NULL);
     return exepath;
    } // CLoadLibraryExplorerDlg::GetExeDir

/****************************************************************************
*                         CLoadLibraryExplorerDlg::GetAppPath
* Result: CString
*       The key for HKLM: \\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\file.exe
****************************************************************************/

CString CLoadLibraryExplorerDlg::GetAppPath()
    {
     CString AppPath;
     AppPath.LoadString(IDS_APP_PATH);
     // construct <AppPath>\file.exe

     TCHAR exepath[MAX_PATH];
     ::GetModuleFileName(NULL, exepath, MAX_PATH);
     TCHAR file[MAX_PATH];
     TCHAR ext[MAX_PATH];
     _tsplitpath(exepath, NULL, NULL, file, ext);
     TCHAR key[MAX_PATH];
     _tmakepath(key, NULL, NULL, file, ext);
     AppendDir(AppPath, key);

     // key is now <AppPath>\file.exe

     CString desired;
     GetRegistryString(HKEY_LOCAL_MACHINE, AppPath, _T("Path"), desired);
     return desired;
    } // CLoadLibraryExplorerDlg::GetAppPath

/****************************************************************************
*                         CLoadLibraryExplorerDlg::InstallDLL
* Inputs:
*       LPCTSTR target: Target directory
*       CState & state: Place to record state
*       DLLSourceMode mode: {UseInternalDll, UseKnownDll}
* Result: BOOL
*       TRUE if successful
*       FALSE if error
* Effect: 
*       Installs the prototype DLL in the target directory
****************************************************************************/

BOOL CLoadLibraryExplorerDlg::InstallDLL(LPCTSTR targetpath, CState & state, DLLSourceMode mode)
    {
     // First, see if the file already exists
     CString source = GetSourceFilename(mode);
     CString target = GetTargetFilename(targetpath, mode);

     if(_tcsicmp(source, target) == 0)
        { /* same, pretend success */
         return TRUE;
        } /* same, pretend success */

     HANDLE hSource = ::CreateFile(source,
                                   GENERIC_READ, // required for GetFileTime
                                   0, // share exclusive
                                   NULL, // security irrelevant
                                   OPEN_EXISTING, // must exist
                                   FILE_ATTRIBUTE_NORMAL, // don't care
                                   NULL); // no template
     // The source file must exist!
     if(hSource == INVALID_HANDLE_VALUE)
        { /* no source */
         DWORD err = ::GetLastError();
         LogMessage(new TraceError(TraceEvent::None, IDS_STARS));
         Alert();
         LogMessage(new TraceFormatError(TraceEvent::None, IDS_NO_SOURCE_DLL, source));
         LogMessage(new TraceFormatMessage(TraceEvent::None, err));
         LogMessage(new TraceError(TraceEvent::None, IDS_STARS));
         ::SetLastError(err);
         CString s;
         s.LoadString(IDS_SOURCE_ERROR);
         state.SetWindowText(s);
         return FALSE;
        } /* no source */

     FILETIME source_created;
     FILETIME source_accessed;
     FILETIME source_modified;

     VERIFY(::GetFileTime(hSource, &source_created, &source_accessed, &source_modified));
     
     HANDLE hTarget = ::CreateFile(target,
                                   GENERIC_READ, // required for GetFileTime
                                   0, // share exclusive
                                   NULL, // security irrelevant
                                   OPEN_EXISTING, // must already exist
                                   FILE_ATTRIBUTE_NORMAL, // don't care
                                   NULL); // no template
                             
     if(hTarget == INVALID_HANDLE_VALUE)
        { /* no target */
         DWORD err = ::GetLastError();
         if(err != ERROR_FILE_NOT_FOUND)
            { /* target failed */
             LogMessage(new TraceError(TraceEvent::None, IDS_STARS));
             Alert();
             LogMessage(new TraceFormatError(TraceEvent::None, IDS_QUERY_TARGET_FAILED, target));
             LogMessage(new TraceFormatMessage(TraceEvent::None, err));
             LogMessage(new TraceError(TraceEvent::None, IDS_STARS));
             ::CloseHandle(hSource);
             
             CString s;
             s.LoadString(IDS_TARGET_ERROR);
             state.SetWindowText(s);
             return FALSE;
            } /* target failed */
        } /* no target */
     else
        { /* has target */
         FILETIME target_created;
         VERIFY(GetFileTime(hTarget, &target_created, NULL, NULL));

         if(source_created.dwLowDateTime == target_created.dwLowDateTime &&
            source_created.dwHighDateTime == target_created.dwHighDateTime)
            { /* same file */
             LogMessage(new TraceFormatComment(TraceEvent::None, IDS_DLL_SAME, target));
             ::CloseHandle(hSource);
             ::CloseHandle(hTarget);
             CString s;
             s.LoadString(IDS_SOURCE_SAME);
             state.SetWindowText(s);
             return TRUE;
            } /* same file */
         ::CloseHandle(hTarget);
        } /* has target */

     // The source exists
     // The target does not exist
     //  OR
     // The target is inconsistent with the source

     ::CloseHandle(hSource);

     if(!::CopyFile(source, target, FALSE)) // allows overwrite
        { /* copy failed */
         DWORD err = ::GetLastError();
         LogMessage(new TraceError(TraceEvent::None, IDS_STARS));
         Alert();
         LogMessage(new TraceFormatError(TraceEvent::None, IDS_COPY_TARGET_FAILED, source, target));
         LogMessage(new TraceFormatMessage(TraceEvent::None, err));
         LogMessage(new TraceError(TraceEvent::None, IDS_STARS));
         CString s;
         s.LoadString(IDS_COPY_ERROR);
         state.SetWindowText(s);
         return FALSE;
        } /* copy failed */
     else
        { /* copy succeeded */
         LogMessage(new TraceFormatComment(TraceEvent::None, IDS_COPY_TARGET_WORKED, target));
         CString s;
         s.LoadString(IDS_COPY_OK);
         state.SetWindowText(s);
         // Note that if the original was read-only, the copy is read-only.  We want to remove
         // the read-only attribte
         DWORD attributes = ::GetFileAttributes(target);
         if(attributes == INVALID_FILE_ATTRIBUTES)
            { /* failed */
             DWORD err = ::GetLastError();
             Alert();
             LogMessage(new TraceFormatError(TraceEvent::None, IDS_GET_ATTRIBUTES_FAILED, target));
             LogMessage(new TraceFormatMessage(TraceEvent::None, err));
             // fall thru, we still have to fiddle times
            } /* failed */
         else
            { /* remove readonly if present */
             if(attributes & FILE_ATTRIBUTE_READONLY)
                { /* clear it */
                 attributes &= ~FILE_ATTRIBUTE_READONLY;
                 if(!::SetFileAttributes(target, attributes))
                    { /* set failed */
                     DWORD err = ::GetLastError();
                     Alert();
                     LogMessage(new TraceFormatError(TraceEvent::None, IDS_SET_ATTRIBUTES_FAILED, target));
                     LogMessage(new TraceFormatMessage(TraceEvent::None, err));
                     // fall thru, we still have to fiddle times
                    } /* set failed */
                } /* clear it */
            } /* remove readonly if present */
        } /* copy succeeded */

     // force the times to be consistent

     hTarget = ::CreateFile(target,
                            FILE_WRITE_ATTRIBUTES, // only attributes access required
                            0, // share exclusive
                            NULL, // security irrelevant
                            OPEN_EXISTING, // must already exist
                            FILE_ATTRIBUTE_NORMAL, // don't care
                            NULL); // no template

     if(hTarget == INVALID_HANDLE_VALUE)
        { /* no target */
         DWORD err = ::GetLastError();
         if(err != ERROR_FILE_NOT_FOUND)
            { /* target failed */
             Alert();
             LogMessage(new TraceFormatError(TraceEvent::None, IDS_SET_TARGET_FAILED, target));
             LogMessage(new TraceFormatMessage(TraceEvent::None, err));
             return TRUE; // copied, but times not updated.  Tough
            } /* target failed */
        } /* no target */
     
     VERIFY(::SetFileTime(hTarget, &source_created, &source_accessed, &source_modified));
     ::CloseHandle(hTarget);
     
     return TRUE;
    } // CLoadLibraryExplorerDlg::InstallDLL

/****************************************************************************
*                   CLoadLibraryExplorerDlg::SetCurrentMode
* Result: void
*       
* Effect: 
*       Sets CurrentMode based on the selections
****************************************************************************/

void CLoadLibraryExplorerDlg::SetCurrentMode()
    {
     if(c_TestSafeDllSearchMode.GetCheck() == BST_CHECKED)
        CurrentMode = UseInternalDll;
     else
     if(c_UseInternalDll.GetCheck() == BST_CHECKED)
        CurrentMode = UseInternalDll;
     else
     if(c_UseKnownDll.GetCheck() == BST_CHECKED)
        CurrentMode = UseKnownDll;
     else
        { /* impossible */
         ASSERT(FALSE);
         return;
        } /* impossible */
    } // CLoadLibraryExplorerDlg::SetCurrentMode

/****************************************************************************
*                          CLoadLibraryExplorerDlg::CopyDLLs
* Result: BOOL
*       TRUE if successful
*       FALSE if error
* Effect: 
*       Makes sure the DLL is on all the selected targets and on none
*       of the deslected targets
****************************************************************************/

BOOL CLoadLibraryExplorerDlg::CopyDLLs()
    {
     BOOL success = TRUE;
     DWORD err = ::GetLastError();
     SetCurrentMode();

     //****************
     // working
     //****************
     if(c_Working.GetCheck() == BST_CHECKED)
        { /* working */
         if(!InstallDLL(WorkingDir, c_StateWorking, CurrentMode))
            { /* failed */
             success = FALSE;
             err = ::GetLastError();
            } /* failed */
        } /* working */
     else
        { /* not working */
         RemoveDLL(WorkingDir, c_StateWorking, CurrentMode);
        } /* not working */

     //****************
     // System32
     //****************
     if(c_SYSTEM32.IsWindowEnabled())
        { /* handle system32 */
         if( c_SYSTEM32.GetCheck() == BST_CHECKED)
            { /* system32 */
             if(!InstallDLL(System32Dir, c_StateSystem32, CurrentMode))
                { /* system32 failed */
                 success = FALSE;
                 err = ::GetLastError();
                } /* system32 failed */
            } /* system32 */
         else
            { /* not system32 */
             RemoveDLL(System32Dir, c_StateSystem32, CurrentMode);
            } /* not system32 */
        } /* handle system32 */
     else
        { /* not used */
         CString s;
         s.LoadString(IDS_SYSTEM32_PRESENT);
         c_StateSystem32.SetWindowText(s);
        } /* not used */
     //****************
     // System
     //****************
     if(c_SYSTEM.GetCheck() == BST_CHECKED)
        { /* system32 */
         if(!InstallDLL(SystemDir, c_StateSystem, CurrentMode))
            { /* system failed */
             success = FALSE;
             err = ::GetLastError();
            } /* system failed */
        } /* system */
     else
        { /* not system */
         RemoveDLL(SystemDir, c_StateSystem, CurrentMode);
        } /* not system */

     //****************
     // Windows
     //****************
     if(c_WINDOWS.GetCheck() == BST_CHECKED)
        { /* windows */
         if(!InstallDLL(WindowsDir, c_StateWindows, CurrentMode))
            { /* windows failed */
             success = FALSE;
             err = ::GetLastError();
            } /* windows failed */
        } /* windows */
     else
        { /* not windows */
         RemoveDLL(WindowsDir, c_StateWindows, CurrentMode);
        } /* not windows */

     //****************
     // App Path
     //****************
     if(!AppPathDir.IsEmpty())
        { /* use App Path */
         if(c_AppPath.GetCheck() == BST_CHECKED)
            { /* App Path */
             if(!InstallDLL(AppPathDir, c_StateAppPath, CurrentMode))
                { /* App Path failed */
                 success = FALSE;
                 err = ::GetLastError();
                } /* App Path failed */
            } /* App Path */
         else
            { /* not App Path */
             RemoveDLL(AppPathDir, c_StateAppPath, CurrentMode);
            } /* not App Path */
        } /* use App Path */

     //****************
     // System
     //****************
     if(c_SYSTEM.GetCheck() == BST_CHECKED)
        { /* system32 */
         if(!InstallDLL(SystemDir, c_StateSystem, CurrentMode))
            { /* system failed */
             success = FALSE;
             err = ::GetLastError();
            } /* system failed */
        } /* system */
     else
        { /* not system */
         RemoveDLL(SystemDir, c_StateSystem, CurrentMode);
        } /* not system */

     //****************
     // Alternate
     //****************
     if(SetDllDirectory != NULL && c_Alternate.GetCheck() == BST_CHECKED)
        { /* Alternate */
         if(!InstallDLL(AlternateDir, c_StateAlternate, CurrentMode))
            { /* alternate failed */
             success = FALSE;
             err = ::GetLastError();
            } /* alternate failed */
        if(!SetDllDirectory(AlternateDir))
            { /* failed */
             DWORD err = ::GetLastError();
             Alert();
             LogMessage(new TraceFormatError(TraceEvent::None, IDS_SETDLLDIRECTORY_FAILED, AlternateDir));
             LogMessage(new TraceFormatMessage(TraceEvent::None, err));
            } /* failed */
         else
            { /* succeeded */
             LogMessage(new TraceFormatComment(TraceEvent::None, IDS_SETDLLDIRECTORY, AlternateDir));
            } /* succeeded */
        } /* Alternate */
     else
        { /* not Alternate */
         if(SetDllDirectory != NULL)
            { /* remove DLL directory from search */
             SetDllDirectory(NULL);
             LogMessage(new TraceFormatComment(TraceEvent::None, IDS_SETDLLDIRECTORY_NULL));
            } /* remove DLL directory from search */
         RemoveDLL(AlternateDir, c_StateAlternate, CurrentMode);
        } /* not Alternate */

     //****************
     // executable
     //****************
     if(c_Executable.GetCheck() == BST_CHECKED)
        { /* executable */
         if(!InstallDLL(GetExeDir(), c_StateExecutable, CurrentMode))
            { /* exe failed */
             success = FALSE;
             err = ::GetLastError();
            } /* exe failed */
        } /* executable */
     else
        { /* not executable */
         RemoveDLL(GetExeDir(), c_StateExecutable, CurrentMode);
        } /* not executable */

     SetRedirection();

     ::SetLastError(err);
     return success;
    } // CLoadLibraryExplorerDlg::CopyDLLs

/****************************************************************************
*               CLoadLibraryExplorerDlg::DeleteRedirectionFile
* Result: void
*       
* Effect: 
*       Deletes the redirection file if it exists.  If it does not exist
*       (or if it is a directory), does nothing
****************************************************************************/

void CLoadLibraryExplorerDlg::DeleteRedirectionFile()
    {
     CString filename = GetLocalName(TRUE); // get full local name
     WIN32_FIND_DATA fd;
     HANDLE h = ::FindFirstFile(filename, &fd);
     if(h == INVALID_HANDLE_VALUE)
        { /* did not exist */
         LogMessage(new TraceFormatComment(TraceEvent::None, IDS_NO_LOCAL_FOUND, filename));
         CString s;
         s.LoadString(IDS_ALREADY_DELETED);
         c_StateLocal.SetWindowText(s);
         return;
        } /* did not exist */

     if(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
        { /* it is directory */
         LogMessage(new TraceFormatComment(TraceEvent::None, IDS_LOCAL_IS_DIRECTORY, filename));
         ::FindClose(h);
         return;
        } /* it is directory */

     ::FindClose(h); // no longer need this

     if(!::DeleteFile(filename))
        { /* deletion failed */
         DWORD err = ::GetLastError();
         Alert();
         LogMessage(new TraceFormatError(TraceEvent::None, IDS_LOCAL_FILE_DELETE_FAILED, filename));
         LogMessage(new TraceFormatMessage(TraceEvent::None, err));
        } /* deletion failed */
     else
        { /* deletion succeeded */
         LogMessage(new TraceFormatComment(TraceEvent::None, IDS_LOCAL_FILE_DELETED, filename));
        } /* deletion succeeded */
     c_StateLocal.SetWindowText(_T(""));
    } // CLoadLibraryExplorerDlg::DeleteRedirectionFile

/****************************************************************************
*             CLoadLibraryExplorerDlg::DeleteRedirectionDirectory
* Result: void
*       
* Effect: 
*       Deletes the redirection directory if it exists.  If it does not
*       exist (it might be a redirection file), does nothing
****************************************************************************/

void CLoadLibraryExplorerDlg::DeleteRedirectionDirectory()
    {
     CString dirname = GetLocalName(TRUE);
     WIN32_FIND_DATA fd;
     HANDLE h = ::FindFirstFile(dirname, &fd);
     if(h == INVALID_HANDLE_VALUE)
        { /* did not exist */
         LogMessage(new TraceFormatComment(TraceEvent::None, IDS_NO_LOCAL_FOUND, dirname));
         CString s;
         s.LoadString(IDS_ALREADY_DELETED);
         c_StateLocal.SetWindowText(s);
         return; // did not exist
        } /* did not exist */

     if(!(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
        { /* not directory */
         LogMessage(new TraceFormatComment(TraceEvent::None, IDS_LOCAL_IS_FILE, dirname));
         ::FindClose(h);
         return;
        } /* not directory */

     // It is a directory.  Delete its contents and delete the directory
     ::FindClose(h);
     CString pattern = dirname;
     pattern += _T("\\*.*");
     h = ::FindFirstFile(pattern, &fd);
     if(h != INVALID_HANDLE_VALUE)
        { /* delete files */
         do
            {
             if(_tcscmp(fd.cFileName, _T(".")) == 0)
                 continue;
             if(_tcscmp(fd.cFileName, _T("..")) == 0)
                 continue;
             if(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
                 { /* bogus! directory present */
                  ASSERT(FALSE); // this should not happen
                  Alert();
                  LogMessage(new TraceFormatError(TraceEvent::None, IDS_DIRECTORY_IN_LOCAL, dirname, fd.cFileName));
                  continue;
                 } /* bogus! directory present */
             CString filename = dirname;
             filename += _T("\\");
             filename += fd.cFileName;
             if(!::DeleteFile(filename))
                { /* deletion failed */
                 DWORD err = ::GetLastError();
                 Alert();
                 LogMessage(new TraceFormatError(TraceEvent::None, IDS_LOCAL_FILENAME_DELETE_FAILED, filename));
                 LogMessage(new TraceFormatMessage(TraceEvent::None, err));
                } /* deletion failed */
             else
                { /* deletion succeeded */
                 LogMessage(new TraceFormatComment(TraceEvent::None, IDS_LOCAL_FILENAME_DELETE_SUCCEEDED, filename));
                } /* deletion succeeded */
            } while(::FindNextFile(h, &fd));
         ::FindClose(h);

         CString s;
         s.LoadString(IDS_DELETE_OK);
         c_StateLocal.SetWindowText(s);
        } /* delete files */
     else
        { /* no files */
         LogMessage(new TraceFormatComment(TraceEvent::None, IDS_LOCAL_DIRECTORY_EMPTY, dirname));
         CString s;
         s.LoadString(IDS_ALREADY_DELETED);
         c_StateLocal.SetWindowText(s);
        } /* no files */
     // The directory was empty or all files have been removed (we hope)

     if(!RemoveDirectory(dirname))
        { /* directory remove failed */
         DWORD err = ::GetLastError();
         Alert();
         LogMessage(new TraceFormatError(TraceEvent::None, IDS_LOCAL_DIRECTORY_DELETE_FAILED, dirname));
         LogMessage(new TraceFormatMessage(TraceEvent::None, err));
        } /* directory remove failed */
     else
        { /* directory remove succeeded */
         LogMessage(new TraceFormatComment(TraceEvent::None, IDS_LOCAL_DIRECTORY_DELETED, dirname));
        } /* directory remove succeeded */
    } // CLoadLibraryExplorerDlg::DeleteRedirectionDirectory

/****************************************************************************
*                     CLoadLibraryExplorerDlg::GetLocalName
* Inputs:
*       BOOL fullpath (= FALSE): TRUE to get full path to the local name
*                                FALSE to get just the local name
* Result: CString
*       <exe>.local
****************************************************************************/

CString CLoadLibraryExplorerDlg::GetLocalName(BOOL fullpath /* = FALSE */)
    {
     TCHAR name[MAX_PATH];
     ::GetModuleFileName(NULL, name, MAX_PATH);
     if(!fullpath)
        { /* name only */
         TCHAR file[MAX_PATH];
         TCHAR ext[MAX_PATH];
         _tsplitpath(name, NULL, NULL, file, ext);
         _tmakepath(name, NULL, NULL, file, ext);
        } /* name only */
     CString s = name;
     s += _T(".local");
     return s;
    } // CLoadLibraryExplorerDlg::GetLocalName

/****************************************************************************
*               CLoadLibraryExplorerDlg::CreateRedirectionFile
* Result: void
*       
* Effect: 
*       Creates a redirection .local file
****************************************************************************/

void CLoadLibraryExplorerDlg::CreateRedirectionFile()
    {
     CString filename = GetLocalName(TRUE);
     HANDLE h = CreateFile(filename,
                           GENERIC_WRITE,
                           0, // exclusive access
                           NULL, // no security
                           CREATE_ALWAYS,
                           FILE_ATTRIBUTE_NORMAL,
                           NULL);
     if(h == INVALID_HANDLE_VALUE)
        { /* create failed */
         DWORD err = ::GetLastError();
         Alert();
         LogMessage(new TraceFormatError(TraceEvent::None, IDS_CREATE_REDIRECT_FILE_FAILED, filename));
         LogMessage(new TraceFormatMessage(TraceEvent::None, err));
         return;
        } /* create failed */
     // contents don't matter, only existence of file
     LogMessage(new TraceFormatComment(TraceEvent::None, IDS_CREATE_REDIRECT_FILE, filename));
     ::CloseHandle(h);
    } // CLoadLibraryExplorerDlg::CreateRedirectionFile

/****************************************************************************
*                CLoadLibraryExplorerDlg::InstallInRedirectDir
* Result: void
*       
* Effect: 
*       Installs the DLL in the redirection directory
****************************************************************************/

void CLoadLibraryExplorerDlg::InstallInRedirectDir()
    {
     CString dirname;
     dirname = GetLocalName(TRUE);

     if(!InstallDLL(dirname, c_StateLocal, CurrentMode))
        { /* create local DLL failed */
         DWORD err = ::GetLastError();
         Alert();
         LogMessage(new TraceFormatError(TraceEvent::None, IDS_INSTALL_REDIRECTED_DLL_FAILED, dirname));
         LogMessage(new TraceFormatMessage(TraceEvent::None, err));
        } /* create local DLL failed */
    } // CLoadLibraryExplorerDlg::InstallInRedirectDir

/****************************************************************************
*             CLoadLibraryExplorerDlg::CreateRedirectionDirectory
* Result: void
*       
* Effect: 
*       Tries to create  a redirection directory and put the DLL in it
* Notes:
*       This assumes that the directory is "clean", that is, there is
*       no file already present
****************************************************************************/

void CLoadLibraryExplorerDlg::CreateRedirectionDirectory()
    {
     CString dirname;
     dirname = GetLocalName(TRUE);
     DWORD attributes = ::GetFileAttributes(dirname);
     if(attributes != INVALID_FILE_ATTRIBUTES)
        { /* got attributes */
         if(attributes & FILE_ATTRIBUTE_DIRECTORY)
            { /* already exists */
             LogMessage(new TraceFormatComment(TraceEvent::None, IDS_REDIRECTION_DIRECTORY_EXISTS, dirname));
             c_Local.SetWindowText(dirname);
             InstallInRedirectDir();
             return;
            } /* already exists */
         Alert();
         LogMessage(new TraceFormatError(TraceEvent::None, IDS_REDIRECTION_DIRECTORY_WRONG, dirname));
         return;
        } /* got attributes */

     if(!::CreateDirectory(dirname, NULL))
        { /* create directory failed */
         DWORD err = ::GetLastError();
         // Does it already exist?
         Alert();
         LogMessage(new TraceFormatError(TraceEvent::None, IDS_REDIRECTION_DIRECTORY_FAILED, dirname));
         LogMessage(new TraceFormatMessage(TraceEvent::None, err));
         return;
        } /* create directory failed */

     c_Local.SetWindowText(dirname);
     InstallInRedirectDir();
    } // CLoadLibraryExplorerDlg::CreateRedirectionDirectory

/****************************************************************************
*                   CLoadLibraryExplorerDlg::SetRedirection
* Result: void
*       
* Effect: 
*       Sets the redirection file/directory
*
* Notes:
*       None: see if <executable>.local directory exists.  If it exists,
*               delete all the files in it and then remove the directory
*       Executable: see if <executable>.local directory exists.  If it
*               exists, delete all the files in it and then remove the
*               directory.
*               Create an empty <executable>.local file.
*       .local directory: see if <executable>.local file exists.  If it exists,
*               delete it.
*               See if <executable>.local directory exists.  If
*               it does not exist, create it
*               Copy the DLL to the new directory
****************************************************************************/

void CLoadLibraryExplorerDlg::SetRedirection()
    {
     c_StateLocal.SetWindowText(_T(""));
     if(c_RedirectNone.GetCheck() == BST_CHECKED)
        { /* no redirection */
         DeleteRedirectionDirectory();
         DeleteRedirectionFile();
        } /* no redirection */
     else
     if(c_RedirectExecutable.GetCheck() == BST_CHECKED)
        { /* .local file */
         DeleteRedirectionDirectory();
         CreateRedirectionFile();
        } /* .local file */
     else
     if(c_RedirectDirectory.GetCheck() == BST_CHECKED)
        { /* .local directory */
         DeleteRedirectionFile();
         CreateRedirectionDirectory();
        } /* .local directory */
    } // CLoadLibraryExplorerDlg::SetRedirection

/****************************************************************************
*                         CLoadLibraryExplorerDlg::RemoveDLL
* Inputs:
*       LPCTSTR target: Path name
*       CState & state: Place to record state
*       DLLSource mode: {UseInternalDll, UseKnownDll}
* Result: BOOL
*       TRUE if successful
*       FALSE if error
* Effect: 
*       Deletes the DLL from the target
****************************************************************************/

BOOL CLoadLibraryExplorerDlg::RemoveDLL(LPCTSTR target, CState & state, DLLSourceMode mode)
    {
     CString path = GetTargetFilename(target, mode);
     CString source = GetSourceFilename(mode);
     if(_tcsicmp(path, source) == 0)
        { /* same one */
         // This is because we are using a known DLL.  We do not wish to delete
         // a known DLL
         return TRUE; // pretend we did it
        } /* same one */

     if(!DeleteFile(path))
        { /* failed */
         DWORD err = ::GetLastError();
         if(err == ERROR_FILE_NOT_FOUND)
            { /* didn't exist */
             LogMessage(new TraceFormatComment(TraceEvent::None, IDS_DELETE_NOT_NEEDED, path));
             CString s;
             s.LoadString(IDS_ALREADY_DELETED);
             state.SetWindowText(s);
             return TRUE; // silently succeed
            } /* didn't exist */
         
         Alert();
         LogMessage(new TraceFormatError(TraceEvent::None, IDS_DELETE_FAILED, path));
         LogMessage(new TraceFormatMessage(TraceEvent::None, err));
         CString s;
         s.LoadString(IDS_DELETE_FAILED);
         state.SetWindowText(s);

         ::SetLastError(err);
         return FALSE;
        } /* failed */
     else
        { /* deleted OK */
         LogMessage(new TraceFormatComment(TraceEvent::None, IDS_DELETED_DLL, path));
         CString s;
         s.LoadString(IDS_DELETE_OK);
         state.SetWindowText(s);
         return TRUE;
        } /* deleted OK */
    } // CLoadLibraryExplorerDlg::RemoveDLL

/****************************************************************************
*                          CLoadLibraryExplorerDlg::EraseDLLs
* Result: BOOL
*       TRUE if successful
*       FALSE if error
* Effect: 
*       Deletes the DLL from all potential targets
****************************************************************************/

BOOL CLoadLibraryExplorerDlg::EraseDLLs(DLLSourceMode mode)
    {
     CString path;
     BOOL success = TRUE;
     DWORD err = ::GetLastError();

     //****************
     // Working directory
     //****************
     if(!RemoveDLL(WorkingDir, c_StateWorking, mode))
        { /* failed */
         success = FALSE;
         err = ::GetLastError();
        } /* failed */
     
     //****************
     // SYSTEM32
     //****************
     if(!RemoveDLL(System32Dir, c_StateSystem32, mode))
        { /* system32 failed */
         success = FALSE;
         err = ::GetLastError();
        } /* system32 failed */

     //****************
     // WINDOWS
     //****************
     if(!RemoveDLL(WindowsDir, c_StateWindows, mode))
        { /* windows failed */
         success = FALSE;
         err = ::GetLastError();
        } /* windows failed */

     //****************
     // App Path directory
     //****************
     if(!AppPathDir.IsEmpty())
        { /* remove AppPathDir */
         if(!RemoveDLL(AppPathDir, c_StateAppPath, mode))
            { /* failed AppPathDir */
             success = FALSE;
             err = ::GetLastError();
            } /* failed AppPathDir */
        } /* remove AppPathDir */

     //****************
     // executable
     //****************
     if(!RemoveDLL(GetExeDir(), c_StateExecutable, mode))
        { /* failed exe */
         success = FALSE;
         err = ::GetLastError();
        } /* failed exe */
     
     ::SetLastError(err);

     //****************
     // redirection
     //****************
     DeleteRedirectionDirectory();
     DeleteRedirectionFile();
     return success;
    } // CLoadLibraryExplorerDlg::EraseDLLs


/****************************************************************************
*                 CLoadLibraryExplorerDlg::GetRegistryAppPathKey
* Inputs:
*       LPCTSTR module: Executable image we are attaching key to
* Result: CString
*       HKLM\Software\Microsofot\Windows\CurrentVersion\App Paths\<exe>
****************************************************************************/

CString CLoadLibraryExplorerDlg::GetRegistryAppPathKey(LPCTSTR module)
    {
     CString AppPathKey;
     AppPathKey.LoadString(IDS_APP_PATH); // HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths

     if(AppPathKey.Right(1) != _T("\\"))
        AppPathKey += _T("\\");

     CString executable = module;

     TCHAR file[MAX_PATH];
     TCHAR ext[MAX_PATH];
     _tsplitpath(module, NULL, NULL, file, ext);

     TCHAR result[MAX_PATH];
     _tmakepath(result, NULL, AppPathKey, file, ext);

     return CString(result);
    } // CLoadLibraryExplorerDlg::GetRegistryAppPathKey

/****************************************************************************
*                 CLoadLibraryExplorerDlg::SetRegistryAppPath
* Inputs:
*       LPCTSTR executable: Name of executable to set
* Result: BOOL
*       TRUE if successful
*       FALSE if error
* Effect: 
*       Sets the App Path for the specified executable
****************************************************************************/

BOOL CLoadLibraryExplorerDlg::SetRegistryAppPath(LPCTSTR executable)
    {
     BOOL result = TRUE;
     CString module = GetRegistryAppPathKey(executable);

     if(c_AppPath.GetCheck() == BST_CHECKED)
        { /* force App Path */
         //----------------------------------------------------------------
         // Set the value
         // HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\whatever.exe = (default) AppPathDir\file.ext\(default)
         //----------------------------------------------------------------

         if(!SetRegistryString(HKEY_LOCAL_MACHINE, module, _T(""), executable))
            { /* failed to set path (default) */
             DWORD err = ::GetLastError();
             Alert();
             LogMessage(new TraceFormatError(TraceEvent::None, IDS_APPPATH_FAILED, module, executable));
             LogMessage(new TraceFormatMessage(TraceEvent::None, err));
             result =  FALSE;
            } /* failed to set path (default) */
         else
            { /* set AppPath */
             LogMessage(new TraceFormatComment(TraceEvent::None, IDS_APPPATH_SET, module, executable));
            } /* set AppPath */

         //----------------------------------------------------------------
         // Set the value
         // HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\whatever.exe = (default) AppPathDir\file.ext\Path
         //----------------------------------------------------------------

         if(!SetRegistryString(HKEY_LOCAL_MACHINE, module, _T("Path"), AppPathDir))
            { /* failed to set Path */
             DWORD err = ::GetLastError();
             Alert();
             LogMessage(new TraceFormatError(TraceEvent::None, IDS_APPPATH_PATH_FAILED, module, AppPathDir));
             LogMessage(new TraceFormatMessage(TraceEvent::None, err));
             result = FALSE;
            } /* failed to set Path */
         else
            { /* set path */
             LogMessage(new TraceFormatComment(TraceEvent::None, IDS_APPPATH_PATH_SET, module, AppPathDir));
            } /* set path */
        } /* force App Path */
     else
        { /* Remove App Path */
         //----------------------------------------------------------------
         // Remove the key
         // HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\whatever.exe = (default) AppPathDir\file.ext
         //----------------------------------------------------------------
         if(!DeleteRegistryKey(HKEY_LOCAL_MACHINE, module))
            { /* delete failed */
             DWORD err = ::GetLastError();
             if(err != ERROR_FILE_NOT_FOUND)
                { /* existing key delete failed */
                 Alert();
                 LogMessage(new TraceFormatError(TraceEvent::None, IDS_APPPATH_NOT_DELETED, module));
                 LogMessage(new TraceFormatMessage(TraceEvent::None, err));
                 result = FALSE;
                } /* existing key delete failed */
             else
                { /* key did not exist */
                 // do nothing right now...just assume failure to delete nonexistent key valid
                } /* key did not exist */
            } /* delete failed */
         else
            { /* delete succeeded */
             // Delete of existing key succeeded
             LogMessage(new TraceFormatComment(TraceEvent::None, IDS_APPPATH_DELETED, module));
            } /* delete succeeded */
        } /* Remove App Path */
     return result;
    } // CLoadLibraryExplorerDlg::SetRegistryAppPath

/****************************************************************************
*               CLoadLibraryExplorerDlg::SetAllRegistryAppPaths
* Result: BOOL
*       TRUE if successful
*       FALSE if failed
* Effect: 
*       Sets the AppPath variable
****************************************************************************/

BOOL CLoadLibraryExplorerDlg::SetAllRegistryAppPaths()
    {
     BOOL result = TRUE;
     //********************************
     // Set HKLM\...\App Path value
     //********************************

     TCHAR executable[MAX_PATH];
     ::GetModuleFileName(NULL, executable, MAX_PATH);

     result = SetRegistryAppPath(executable);

     CString tester;
     tester = GetExeDir();
     tester += TEST_PROGRAM_NAME;

     result &= SetRegistryAppPath(tester);
     return result;
    } // CLoadLibraryExplorerDlg::SetAllRegistryAppPaths

/****************************************************************************
*                    CLoadLibraryExplorerDlg::SetupContext
* Result: void
*       
* Effect: 
*       Establishes the context for the execution
****************************************************************************/

void CLoadLibraryExplorerDlg::SetupContext(BOOL & oldsearch)
    {
     RegistryInt SafeDllSearchMode(IDS_REGISTRY_SAFEDLLSEARCHMODE, HKEY_LOCAL_MACHINE);
     SafeDllSearchMode.load();

     oldsearch = SafeDllSearchMode.value;

     CString appPath;
     appPath = GetAppPath();

     if(c_UseSelectedSafeDllSearchMode.GetCheck() == BST_CHECKED)
        { /* change mode */
         if(SafeDllSearchMode.value != c_DesiredSafeDllSearchMode.GetCheck())
            { /* values are different */
             SafeDllSearchMode.value = c_DesiredSafeDllSearchMode.GetCheck();

             if(!SafeDllSearchMode.store())
                { /* error */
                 DWORD err = ::GetLastError();
                 LogMessage(new TraceError(TraceEvent::None, IDS_CANT_STORE_SAFEDLLSEARCHMODE));

                } /* error */
             else
                { /* success */
                 LogMessage(new TraceFormatComment(TraceEvent::None, IDS_SET_SAFEDLLSEARCHMODE, SafeDllSearchMode.value));
                } /* success */
            } /* values are different */
        } /* change mode */

     CopyDLLs();

     SetPath();
    } // CLoadLibraryExplorerDlg::SetupContext

/****************************************************************************
*                    CLoadLibraryExplorerDlg::UndoContext
* Inputs:
*       BOOL OldSafeDllMode: 
* Result: void
*       
* Effect: 
*       Restores the state
****************************************************************************/

void CLoadLibraryExplorerDlg::UndoContext(BOOL oldSafeDllMode)
    {
     RegistryInt SafeDllSearchMode(IDS_REGISTRY_SAFEDLLSEARCHMODE, HKEY_LOCAL_MACHINE);
     SafeDllSearchMode.value = oldSafeDllMode;
     SafeDllSearchMode.store();
    } // CLoadLibraryExplorerDlg::UndoContext

/****************************************************************************
*                          CLoadLibraryExplorerDlg::LoadDLL
* Result: BOOL
*       TRUE if successful
*       FALSE if error: use GetLastError to determine reason
* Effect: 
*       Sets some parameters
****************************************************************************/

BOOL CLoadLibraryExplorerDlg::LoadDLL()
    {
     LogMessage(new TraceComment(TraceEvent::None, IDS_LINE));

     BOOL oldSearchMode;
     SetupContext(oldSearchMode);

     //********************************
     // HKLM path has been set
     //********************************

     BOOL result = DoLoad();
     DWORD err = ::GetLastError();

     UndoContext(oldSearchMode);
     
     SetEnvironmentVariable(_T("PATH"), PathVariable);

     ::SetLastError(err);
     return result;
    } // CLoadLibraryExplorerDlg::LoadDLL

/****************************************************************************
*                           CLoadLibraryExplorerDlg::DoLoad
* Result: BOOL
*       TRUE if successful
*       FALSE if error
* Effect: 
*       Makes sure the DLL is in all the selected locations and NOT in
*       any of the unselected locations
****************************************************************************/

BOOL CLoadLibraryExplorerDlg::DoLoad()
    {
     RegistryInt SafeDllSearchMode(IDS_REGISTRY_SAFEDLLSEARCHMODE, HKEY_LOCAL_MACHINE);
     SafeDllSearchMode.load();


     DWORD flags = 0;
     // Set flags here
     if(c_LOAD_WITH_ALTERED_SEARCH_PATH.GetCheck() == BST_CHECKED)
        flags |= LOAD_WITH_ALTERED_SEARCH_PATH;

     CString path = GetLoadPath(CurrentMode);
     // Decide if full path or partial path should be used!


     HMODULE lib;
     if(c_UseLoadLibrary.GetCheck() == BST_CHECKED)
        { /* LoadLibrary */
         lib = ::LoadLibrary(path);
        } /* LoadLibrary */
     else
     if(c_UseLoadLibraryEx.GetCheck() == BST_CHECKED)
        { /* Use LoadLibraryEx */
         lib = ::LoadLibraryEx(path, NULL, flags);
        } /* Use LoadLibraryEx */
     else
     if(c_UseShellExecute.GetCheck() == BST_CHECKED)
        { /* Use ShellExecute */
         LogMessage(new TraceError(TraceEvent::None, IDS_SHELLEXECUTE_NOT_IMPLEMENTED));
         SetState(_T(""));
         return FALSE;
        } /* Use ShellExecute */
     if(lib == NULL)
        { /* Load failed */
         DWORD err = ::GetLastError();
         Alert();
         LogMessage(new TraceFormatError(TraceEvent::None, IDS_LOAD_FAILED, path));
         LogMessage(new TraceFormatMessage(TraceEvent::None, err));
         c_Library.SetWindowText(ErrorString(err));
         SetState(_T(""));
         return FALSE;
        } /* Load failed */

#if 0
     typedef BOOL (*GetDLLNameProc)(LPTSTR buffer, DWORD len);
     GetDLLNameProc GetDLLName = (GetDLLNameProc)::GetProcAddress(lib, "GetDLLName");

     if(GetDLLName == NULL)
        { /* GetProcAddress failed */
         DWORD err = ::GetLastError();
         Alert();
         LogMessage(new TraceFormatError(TraceEvent::None, IDS_GETPROCADDRESS_FAILED));
         LogMessage(new TraceFormatMessage(TraceEvent::None, err));
         c_Library.SetWindowText(ErrorString(err));
         SetState(_T(""));
         ::FreeLibrary(lib);
         return FALSE;
        } /* GetProcAddress failed */

     TCHAR buffer[MAX_PATH];
     BOOL result = GetDLLName(buffer, MAX_PATH);
     if(!result)
        { /* failed */
         DWORD err = ::GetLastError();
         Alert();
         LogMessage(new TraceFormatError(TraceEvent::None, IDS_GETDLLNAME_FAILED));
         LogMessage(new TraceFormatMessage(TraceEvent::None, err));
         c_Library.SetWindowText(ErrorString(err));
         SetState(_T(""));
         ::FreeLibrary(lib);
         return FALSE;
        } /* failed */
#else
     TCHAR buffer[MAX_PATH];
     if(!::GetModuleFileName(lib, buffer, MAX_PATH))
        { /* failed */
         DWORD err = ::GetLastError();
         Alert();
         LogMessage(new TraceFormatError(TraceEvent::None, IDS_GETDLLNAME_FAILED));
         LogMessage(new TraceFormatMessage(TraceEvent::None, err));
         c_Library.SetWindowText(ErrorString(err));
         SetState(_T(""));
         ::FreeLibrary(lib);
         return FALSE;
        } /* failed */
#endif

     c_Library.SetWindowText(buffer);
     SetState(buffer);
     LogMessage(new TraceFormatComment(TraceEvent::None, IDS_GETDLLNAME_WORKED, buffer));

     ::FreeLibrary(lib);
    return TRUE;
    } // CLoadLibraryExplorerDlg::DoLoad

void CLoadLibraryExplorerDlg::OnBnClickedLoad()
    {
     if(c_TestSafeDllSearchMode.GetCheck() == BST_CHECKED)
        { /* SafeDllSearchMode */
         LaunchTestProgram();
        } /* SafeDllSearchMode */
     else
        { /* LoadLibrary/LoadLibraryEx */
         LoadDLL();
        } /* LoadLibrary/LoadLibraryEx */
    }

void CLoadLibraryExplorerDlg::OnBnClickedUseLoadlibrary()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedUseLoadlibraryex()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedLoadWithAlteredSearchPath()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedUseShellexecute()
    {
     SaveOptions();
     updateControls();
    }


/****************************************************************************
*                 CLoadLibraryExplorerDlg::ConditionalEnable
* Inputs:
*       CButton & ctl: Control to enable/disable
*       LPCTSTR text: Text to compare
* Result: void
*       
* Effect: 
*       The control whose content matches the text is disabled
****************************************************************************/

void CLoadLibraryExplorerDlg::ConditionalEnable(CButton & ctl, LPCTSTR text)
    {
     CString s;
     ctl.GetWindowText(s);
     ctl.EnableWindow(_tcsicmp(s, text) != 0);
    } // CLoadLibraryExplorerDlg::ConditionalEnable

/****************************************************************************
*                    CLoadLibraryExplorerDlg::updateControls
* Result: void
*       
* Effect: 
*       Updates the controls
****************************************************************************/

void CLoadLibraryExplorerDlg::updateControls()
    {
     CString LL;
     BOOL enable = FALSE;

     CString path = GetLoadPath(CurrentMode);

     BOOL SafeModeTest = c_TestSafeDllSearchMode.GetCheck() == BST_CHECKED;
     
     //***************************
     // Show LoadLibrary(...) call
     //***************************
     c_UseLoadLibrary.EnableWindow(!SafeModeTest);
     c_UseLoadLibraryEx.EnableWindow(!SafeModeTest);
     c_LoadMethodCaption.EnableWindow(!SafeModeTest);
     
     if(SafeModeTest)
        { /* no LoadLibrary */
         LL.Format(_T("ShellExecute(NULL, _T(\"open\"), _T(\"%s\"), ...)"),  TEST_PROGRAM_NAME);
        } /* no LoadLibrary */
     else
     if(c_UseLoadLibrary.GetCheck() == BST_CHECKED)
        { /* Use LoadLibrary */
         LL.Format(_T("LoadLibrary(_T(\"%s\"))"), path);
        } /* Use LoadLibrary */
     else
     if(c_UseLoadLibraryEx.GetCheck() == BST_CHECKED)
        { /* Use LoadLibraryEx */
         CString flagstring = _T("0");
         if(c_LOAD_WITH_ALTERED_SEARCH_PATH.GetCheck() == BST_CHECKED)
            flagstring = _T("LOAD_WITH_ALTERED_SEARCH_PATH");
         LL.Format(_T("LoadLibraryEx(_T(\"%s\"), NULL, %s);"), path, flagstring);
         enable = TRUE;
        } /* Use LoadLibraryEx */
     else
     if(c_UseShellExecute.GetCheck() == BST_CHECKED)
        { /* use ShellExecute */
         LL.Format(_T("ShellExecute(...); // Not implemented"));
        } /* use ShellExecute */

     c_LoadLibrary.SetWindowText(LL);

     //********************************
     // LOAD_WITH_ALTERED_SEARCH_PATH
     //********************************
     
     c_LOAD_WITH_ALTERED_SEARCH_PATH.EnableWindow(enable & !SafeModeTest);

     //********************************
     // SafeDllSearchMode options
     //********************************
     c_DesiredSafeDllSearchMode.EnableWindow(SafeModeTest);
     c_DesiredSafeDllSearchMode.EnableWindow(SafeModeTest && c_UseSelectedSafeDllSearchMode.GetCheck() == BST_CHECKED);
     c_DefaultSafeDllSearchMode.EnableWindow(SafeModeTest && c_UseExistingSafeDllSearchMode.GetCheck() == BST_CHECKED);
     c_UseSelectedSafeDllSearchMode.EnableWindow(SafeModeTest);
     c_UseExistingSafeDllSearchMode.EnableWindow(SafeModeTest);

     //********************************
     // Load button
     //********************************
     c_Load.EnableWindow(!launching);

     //********************************
     // Enable force-path controls
     //********************************
     enable = !SafeModeTest && c_Force.GetCheck() == BST_CHECKED;
     c_ForceSystem32.EnableWindow(enable);
     c_ForceWindows.EnableWindow(enable);
     c_ForceSystem.EnableWindow(enable);
     c_ForceWorking.EnableWindow(enable);
     c_ForceAppPath.EnableWindow(enable);
     c_ForceExecutable.EnableWindow(enable);
     c_ForceAlternate.EnableWindow(enable);
     c_Force.EnableWindow(!SafeModeTest);

     BOOL NeedRestart = FALSE;
     //================================
     // AppPath reporting
     //================================
     //      | Registry        In
     // Case | at startup      PATH=    c_AppPath       Effect
     //------+-----------------------------------------------------------
     // I.   | Absent          No       [ ]             All good
     // II.  | Absent          No       [X]             Must restart
     // III. | Absent          Yes      [ ]             Must restart
     // IV.  | Absent          Yes      [X]             All good
     // V.   | Present         No       [ ]             All good
     // VI.  | Present         No       [X]             Must restart
     // VII. | Present         Yes      [ ]             All good
     // VIII.| Present         Yes      [X]             All good
     //       (HasAppPath)    (InPath)  c_AppPath.GetCheck()

     BOOL InPath = IsInPath(GetAppPath());                                                         

     if(!HasAppPath && !InPath && c_AppPath.GetCheck() == BST_UNCHECKED)
        { /* I. */
         NeedRestart = FALSE;
#ifdef _DEBUG
         LogMessage(new TraceComment(TraceEvent::None, _T("I.")));
#endif
        } /* I. */
     else
     if(!HasAppPath && !InPath && c_AppPath.GetCheck() == BST_CHECKED)
        { /* II. */
         NeedRestart = TRUE;
#ifdef _DEBUG
         LogMessage(new TraceComment(TraceEvent::None, _T("II.")));
#endif
        } /* II. */
     else
     if(!HasAppPath && InPath && c_AppPath.GetCheck() == BST_UNCHECKED)
        { /* III. */
         NeedRestart = TRUE;
#ifdef _DEBUG
         LogMessage(new TraceComment(TraceEvent::None, _T("III.")));
#endif
        } /* III. */
     else
     if(!HasAppPath && InPath && c_AppPath.GetCheck() == BST_CHECKED)
        { /* IV. */
         NeedRestart = FALSE;
#ifdef _DEBUG
         LogMessage(new TraceComment(TraceEvent::None, _T("IV.")));
#endif
        } /* IV. */
     else
     if(HasAppPath && !InPath && c_AppPath.GetCheck() == BST_UNCHECKED)
        { /* V. */
         NeedRestart = FALSE;
#ifdef _DEBUG
         LogMessage(new TraceComment(TraceEvent::None, _T("V.")));
#endif
        } /* V. */
     else
     if(HasAppPath && !InPath && c_AppPath.GetCheck() == BST_CHECKED)
        { /* VI. */
         NeedRestart = TRUE;
#ifdef _DEBUG
         LogMessage(new TraceComment(TraceEvent::None, _T("VI.")));
#endif
        } /* VI. */
     else
     if(HasAppPath && InPath && c_AppPath.GetCheck() == BST_UNCHECKED)
        { /* VII. */
         NeedRestart = FALSE;
#ifdef _DEBUG
         LogMessage(new TraceComment(TraceEvent::None, _T("VII.")));
#endif
        } /* VII. */
     else
     if(HasAppPath && InPath && c_AppPath.GetCheck() == BST_CHECKED)
        { /* VIII. */
         NeedRestart = FALSE;
#ifdef _DEBUG
         LogMessage(new TraceComment(TraceEvent::None, _T("VIII.")));
#endif
        } /* VIII. */

     c_Relaunch.EnableWindow(!SafeModeTest && NeedRestart);

     BOOL hasMessage = FALSE;
     
     if(NeedRestart)
        { /* relaunch required */
         CString s;
         s.LoadString(IDS_RELAUNCH_REQUIRED);
         c_Message.SetWindowText(s);
         hasMessage = TRUE;
        } /* relaunch required */

     if(hasMessage && !SafeModeTest)
        { /* show message */
         c_Message.ShowWindow(SW_SHOW);
        } /* show message */
     else
        { /* nothing to show */
         c_Message.ShowWindow(SW_HIDE);
        } /* nothing to show */     

     //****************
     // Dll selection
     //****************
     c_DialogSelectionCaption.EnableWindow(!SafeModeTest);
     c_UseKnownDll.EnableWindow(!SafeModeTest);
     c_UseInternalDll.EnableWindow(!SafeModeTest);

     enable = c_UseKnownDll.GetCheck() == BST_CHECKED;
     c_KnownDlls.EnableWindow(!SafeModeTest && enable);
     c_DllDirectory.EnableWindow(!SafeModeTest && enable);

     CString dlldir;
     c_DllDirectory.GetWindowText(dlldir);
     if(c_UseKnownDll.GetCheck() == BST_CHECKED)
        { /* use known */
         CString expanded = ExpandString(dlldir);
         ConditionalEnable(c_SYSTEM32, expanded);
         ConditionalEnable(c_WINDOWS, expanded);
         ConditionalEnable(c_Working, expanded);
         ConditionalEnable(c_AppPath, expanded);
         ConditionalEnable(c_Executable, expanded);
         ConditionalEnable(c_SYSTEM, expanded);
         if(SetDllDirectory != NULL)
            ConditionalEnable(c_Alternate, expanded);
        } /* use known */
     else
        { /* all possible */
         c_SYSTEM32.EnableWindow(TRUE);
         c_WINDOWS.EnableWindow(TRUE);
         c_Working.EnableWindow(TRUE);
         c_AppPath.EnableWindow(TRUE);
         c_Executable.EnableWindow(TRUE);
         c_SYSTEM.EnableWindow(TRUE);
         c_Alternate.EnableWindow(TRUE && SetDllDirectory != NULL);
        } /* all possible */

     c_RedirectExecutable.EnableWindow(SetDllDirectory != NULL);
     c_RedirectDirectory.EnableWindow(SetDllDirectory != NULL);
     c_RedirectNone.EnableWindow(SetDllDirectory != NULL);
     c_RedirectCaption.EnableWindow(SetDllDirectory != NULL);
     if(c_RedirectDirectory.GetCheck() == BST_CHECKED)
        { /* show state */
         CString dirname;
         dirname = GetLocalName(TRUE);
         c_Local.SetWindowText(dirname);
         c_StateLocal.ShowWindow(SW_SHOW);
         c_Local.ShowWindow(SW_SHOW);
         c_Local.SetCheck(BST_CHECKED);
        } /* show state */
     else
        { /* hide state */
         c_StateLocal.ShowWindow(SW_HIDE);
         c_Local.ShowWindow(SW_HIDE);
        } /* hide state */
    } // CLoadLibraryExplorerDlg::updateControls

/****************************************************************************
*                    CLoadLibraryExplorerDlg::GetLoadName
* Inputs:
*       DLLSource mode: {UseInternalDll, UseKnownDll}
* Result: CString
*       The filename part of the load path
****************************************************************************/

CString CLoadLibraryExplorerDlg::GetLoadName(DLLSourceMode mode)
    {
     if(mode == UseInternalDll)
        return DLL_NAME;
     else
     if(mode == UseKnownDll)
        { /* known dll */
         int n = c_KnownDlls.GetCurSel();
         ASSERT(n != CB_ERR);
         if(n == CB_ERR)
            return _T("");
         KnownDLLInfo * info =  c_KnownDlls.GetItemDataPtr(n);
         return info->name;
        } /* known dll */
     else
        { /* failed */
         ASSERT(FALSE);
         return _T("");
        } /* failed */
    } // CLoadLibraryExplorerDlg::GetLoadName

/****************************************************************************
*                        CLoadLibraryExplorerDlg::GetLoadPath
* Inputs:
*       DLLSource mode: {UseInternalDll, UseKnownDll}
* Result: CString
*       The path to use to load the library
****************************************************************************/

CString CLoadLibraryExplorerDlg::GetLoadPath(DLLSourceMode mode)
    {
     CString basename = GetLoadName(mode);
     if(c_Force.GetCheck() != BST_CHECKED)
        { /* just plain name */
         return basename;
        } /* just plain name */

     // We have something other than a simple base name
     
     CString path;
     //****************
     // Alternae
     //****************
     if(c_ForceAlternate.GetCheck() == BST_CHECKED)
        path = AlternateDir;
     else
     //****************
     // AppPath
     //****************
     if(c_ForceAppPath.GetCheck() == BST_CHECKED)
        path = AppPathDir;
     else
     //****************
     // Executable
     //****************
     if(c_ForceExecutable.GetCheck() == BST_CHECKED)
        path = GetExeDir();
     else
     //****************
     // System
     //****************
     if(c_ForceSystem.GetCheck() == BST_CHECKED)
        path = SystemDir;
     else
     //****************
     // System32
     //****************
     if(c_ForceSystem32.GetCheck() == BST_CHECKED)
        path = System32Dir;
     else
     //****************
     // Windows
     //****************
     if(c_ForceWindows.GetCheck() == BST_CHECKED)
        path = WindowsDir;
     else
     //****************
     // Working
     //****************
     if(c_ForceWorking.GetCheck() == BST_CHECKED)
        path = WorkingDir;
     else
        { /* unknown */
         ASSERT(FALSE);
         path = GetExeDir();
        } /* unknown */
 
     AppendDir(path, basename);
     return path;
    } // CLoadLibraryExplorerDlg::GetLoadPath


/****************************************************************************
*       CLoadLibraryExplorerDlg::OnBnClickedUseSlectedSafedllsearchmode
*           CLoadLibraryExplorerDlg::OnBnClickedUseDefaultSearchMode
*         CLoadLibraryExplorerDlg::OnBnClickedDesiredSafedllsearchmode
*
* Result: void
*       
* Effect: 
*       selects the "Safe DLL Search Mode" variant to be used
****************************************************************************/


void CLoadLibraryExplorerDlg::OnBnClickedUseSelectedSafedllsearchmode()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedUseDefaultSearchMode()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedDesiredSafedllsearchmode()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedRedirectNone()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedRedirectExecutable()
    {
     SaveOptions();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedRedirectLocal()
    {
     SaveOptions();
     updateControls();
    }


/****************************************************************************
*                         CLoadLibraryExplorerDlg::stringsort
* Inputs:
*       const void * e1: (const void *)(const CString *)
*       const void * e2: (const void *)(const CString *)
* Result: int
*       <0 if e1 < e2
*       == 0 if e1 == e2
*       > 0 if e1 > e2
****************************************************************************/

/* static */ int CDECL CLoadLibraryExplorerDlg::stringsort(const void * e1, const void * e2)
    {
     const CString * s1 = (const CString *)e1;
     const CString * s2 = (const CString *)e2;

     return lstrcmp(*s1, *s2);
    } // CLoadLibraryExplorerDlg::stringsort

/****************************************************************************
*                      CLoadLibraryExplorerDlg::OnBnClickedEnum
* Result: void
*       
* Effect: 
*       Produces a list of all the currently loaded modules
****************************************************************************/

void CLoadLibraryExplorerDlg::OnBnClickedEnum()
    {
     CArray<HMODULE, HMODULE> modules;
     modules.SetSize(1); // start off assuming 1 (why not...)

     LogMessage(new TraceComment(TraceEvent::None, IDS_LINE));

     while(TRUE)
        { /* enumerate all */
         size_t desired = modules.GetSize() * sizeof(HMODULE);
         DWORD gotten;
         BOOL result = EnumProcessModules(GetCurrentProcess(), modules.GetData(), (DWORD)(UINT_PTR)desired, &gotten);
         if(!result)
            { /* failed */
             DWORD err = ::GetLastError();
             LogMessage(new TraceError(TraceEvent::None, IDS_ENUM_FAILED));
             LogMessage(new TraceFormatMessage(TraceEvent::None, err));
             return;
            } /* failed */

         if(gotten < desired)
            { /* all done */
             modules.SetSize(gotten / sizeof(HMODULE));
             break; // got them all
            } /* all done */
         modules.SetSize(max(1, 2 * modules.GetSize()));
        } /* enumerate all */


     CStringArray ModuleList;
     
     LogMessage(new TraceComment(TraceEvent::None, IDS_MODULE_LIST));
     for(int i = 0; i < modules.GetSize(); i++)
        { /* print each */
         TCHAR name[MAX_PATH];
         if(GetModuleFileName(modules[i], name, MAX_PATH))
            { /* got name */
             ModuleList.Add(name);
            } /* got name */
         else
            { /* failed name */
             DWORD err = ::GetLastError();
             Alert();
             LogMessage(new TraceFormatError(TraceEvent::None, IDS_BAD_MODULE_HANDLE, modules[i]));
             LogMessage(new TraceFormatMessage(TraceEvent::None, err));
            } /* failed name */
        } /* print each */
                                

     if(c_Alphabetical.GetCheck() == BST_CHECKED)
        { /* sort it */
         qsort(ModuleList.GetData(), ModuleList.GetSize(), sizeof(CString), stringsort);
        } /* sort it */

     for(int i = 0; i < ModuleList.GetSize(); i++)
        { /* log each */
         LogMessage(new TraceComment(TraceEvent::None, ModuleList[i]));
        } /* log each */
    }

/****************************************************************************
*                   CLoadLibraryExplorerDlg::LoadKnownDlls
* Result: void
*       
* Effect: 
*       Lists the known DLLs
****************************************************************************/

void CLoadLibraryExplorerDlg::LoadKnownDlls()
    {
     CStringArray dllkeys;

     CString KnownDlls = _T("\\System\\CurrentControlSet\\Control\\Session Manager\\KnownDLLs");
     
     EnumRegistryValues(HKEY_LOCAL_MACHINE, KnownDlls, dllkeys);
     
     for(int i = 0; i < dllkeys.GetSize(); i++)
        { /* scan each */
         CString value;
         if(GetRegistryString(HKEY_LOCAL_MACHINE, KnownDlls, dllkeys[i], value))
            { /* got it */
             CString key = dllkeys[i];
             if(_tcsicmp(key, _T("dlldirectory")) == 0)
                { /* DllDirectory */
                 c_DllDirectory.SetWindowText(value);
                } /* DllDirectory */
             else\
                { /* Known DLL */
                 HMODULE module = ::GetModuleHandle(value);
                 BOOL preloaded = module != NULL;
                 c_KnownDlls.AddString(new KnownDLLInfo(value, preloaded));
                } /* Known DLL */
            } /* got it */
         else
            { /* failed */
             DWORD err = ::GetLastError();
             Alert();
             LogMessage(new TraceFormatError(TraceEvent::None, IDS_KNOWN_DLL_FAILURE, dllkeys[i]));
             LogMessage(new TraceFormatMessage(TraceEvent::None, err));
            } /* failed */
        } /* scan each */

     CString dlldir;
     c_DllDirectory.GetWindowText(dlldir);
     dlldir.TrimLeft();
     dlldir.TrimRight();
     BOOL cantuseknown = dlldir.IsEmpty();

     // We want to make sure the dllDirectory is also the SYSTEM32 directory
     CString system32;
     dlldir = ExpandString(dlldir);
     c_SYSTEM32.GetWindowText(system32);
     // They had better be equal or we don't allow the option
     if(_tcsicmp(system32, dlldir) != 0)
        cantuseknown = TRUE;
     
     if(cantuseknown)
        { /* really bad trouble */
         CheckRadioButton(IDC_USE_THE_DLL, IDC_KNOWN_DLL, IDC_USE_THE_DLL); // force "internal DLL"
         c_UseKnownDll.EnableWindow(FALSE); // and don't allow anything else
         // This depends on updateControls being called later.
        } /* really bad trouble */
     
     c_KnownDlls.SetCurSel(0);
     // If we are restarting, we need to re-select the originally-selected DLL
     RegistryString knowndll(IDS_REGISTRY_KNOWN_DLL);
     knowndll.load();
     if(!knowndll.value.IsEmpty())
        { /* has selection */
         for(int i = 0; i < c_KnownDlls.GetCount(); i++)
            { /* find it */
             KnownDLLInfo * info =  c_KnownDlls.GetItemDataPtr(i);
             if(_tcsicmp(info->name, knowndll.value) == 0)
                { /* found it */
                 c_KnownDlls.SetCurSel(i);
                 break;
                } /* found it */
            } /* find it */
        } /* has selection */
    }

/****************************************************************************
*                CLoadLibraryExplorerDlg::OnBnClickedRelaunch
* Result: void
*       
* Effect: 
*       Re-launches this app, and if successful, kills this running version
****************************************************************************/

void CLoadLibraryExplorerDlg::OnBnClickedRelaunch()
    {
     TCHAR executable[MAX_PATH];
     ::GetModuleFileName(NULL, executable, MAX_PATH);
     
     RegistryInt relaunch(IDS_REGISTRY_RELAUNCH);
     relaunch.value = TRUE;
     relaunch.store();

     RegistryWindowPlacement wp(IDS_REGISTRY_WINDOWPLACEMENT);
     GetWindowPlacement(&wp.value);
     wp.store();
     
     HINSTANCE inst = ShellExecute(m_hWnd, _T("open"), executable, NULL, NULL, SW_SHOW);
     if((UINT_PTR)inst > 32)
        { /* successful launch */
         PostMessage(WM_CLOSE); // kill current executable
        } /* successful launch */
     else
        { /* launch failure */
         DWORD err = (DWORD)(UINT_PTR)inst;  // error code
         Alert();
         LogMessage(new TraceFormatError(TraceEvent::None, IDS_RELAUNCH_FAILED, executable));
         LogMessage(new TraceFormatMessage(TraceEvent::None, err));
         return;
        } /* launch failure */
    }

/****************************************************************************
*                      CLoadLibraryExplorerDlg::SetState
* Inputs:
*       CString name: The path to the DLL
* Result: void
*       
* Effect: 
*       Matches the path to the options
****************************************************************************/

void CLoadLibraryExplorerDlg::SetState(CString name)
    {
     TCHAR drive[MAX_PATH];
     TCHAR path[MAX_PATH];
     _tsplitpath(name, drive, path, NULL, NULL);

     TCHAR target[MAX_PATH];
     _tmakepath(target, drive, path, NULL, NULL);
     if(target[lstrlen(target) - 1] == _T('\\'))
         target[lstrlen(target) - 1] = _T('\0');

     CheckPath(target, c_SYSTEM32, c_StateSystem32);
     CheckPath(target, c_SYSTEM, c_StateSystem);
     CheckPath(target, c_WINDOWS, c_StateWindows);
     CheckPath(target, c_Working, c_StateWorking);
     CheckPath(target, c_AppPath, c_StateAppPath);
     CheckPath(target, c_Executable, c_StateExecutable);
     CheckPath(target, c_Alternate, c_StateAlternate);
     CheckPath(target, c_Local, c_StateLocal);
    } // CLoadLibraryExplorerDlg::SetState

/****************************************************************************
*                     CLoadLibraryExplorerDlg::CheckPath
* Inputs:
*       CString target: Target name
*       CButton & button: Button (caption text should be path)
*       CState & state: State to set
* Result: void
*       
* Effect: 
*       Sets the "State" property for the button that matches the
*       found path
****************************************************************************/

void CLoadLibraryExplorerDlg::CheckPath(CString target, CButton & button, CState & state)
    {
     CString caption;
     button.GetWindowText(caption);
     state.SetHighlight(_tcsicmp(caption, target) == 0);
    } // CLoadLibraryExplorerDlg::CheckPath

/****************************************************************************
*                      CLoadLibraryExplorerDlg::IsInPath
* Inputs:
*       CString name: Path name
* Result: BOOL
*       TRUE if the path name appears in the PATH=
****************************************************************************/

BOOL CLoadLibraryExplorerDlg::IsInPath(CString name)
    {
     for(int i = 0; i < c_Path.GetCount(); i++)
        { /* check each */
         CString s;
         c_Path.GetText(i, s);
         if(_tcsicmp(name, s) == 0)
            return TRUE;
        } /* check each */
     return FALSE;
    } // CLoadLibraryExplorerDlg::IsInPath

/****************************************************************************
*                CLoadLibraryExplorerDlg::OnBnClickedUseTheDll
*                 CLoadLibraryExplorerDlg::OnBnClickedKnownDll
* Result: void
*       
* Effect: 
*       Records which button was selected and updates the controls
****************************************************************************/

void CLoadLibraryExplorerDlg::OnBnClickedUseTheDll()
    {
     if(CurrentMode == UseInternalDll)
         return; // nothing to do, already in this mode
     EraseDLLs(CurrentMode);
     RegistryInt DllSelection(IDS_REGISTRY_DLL_SELECTION);
     DllSelection.value = IDC_USE_THE_DLL;
     DllSelection.store();
     ClearLoadState();
     SetCurrentMode();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnBnClickedKnownDll()
    {
     if(CurrentMode == UseKnownDll)
         return; // nothing to do, already in this mode
     EraseDLLs(CurrentMode);
     RegistryInt DllSelection(IDS_REGISTRY_DLL_SELECTION);
     DllSelection.value = IDC_KNOWN_DLL;
     DllSelection.store();
     ClearLoadState();
     SetCurrentMode();
     updateControls();
    }

/****************************************************************************
*                 CLoadLibraryExplorerDlg::GetSourceFilename
* Inputs:
*       DLLSource mode: {UseInternalDll, UseKnownDll}
* Result: CString
*       The source file name to use for the DLL
****************************************************************************/

CString CLoadLibraryExplorerDlg::GetSourceFilename(DLLSourceMode mode)
    {
     if(mode == UseKnownDll)
        { /* known DLL */
         CString dlldir;
         c_DllDirectory.GetWindowText(dlldir);
         dlldir = ExpandString(dlldir);
         CString name = GetLoadName(mode);
         AppendDir(dlldir, name);
         return dlldir;
        } /* known DLL */
     else
     if(mode == UseInternalDll)
        { /* internal DLL */
         CString source = GetExeDir();
         AppendDir(source, DLL_SOURCE);
         return source;
        } /* internal DLL */
     else
        { /* should not happen */
         ASSERT(FALSE);
         return _T("");
        } /* should not happen */
    } // CLoadLibraryExplorerDlg::GetSourceFilename

/****************************************************************************
*                 CLoadLibraryExplorerDlg::GetTargetFilename
* Inputs:
*       LPCTSTR targetpath: Desired target path
*       DLLSource mode: Desired mode of source
* Result: CString
*       The desired target path
****************************************************************************/

CString CLoadLibraryExplorerDlg::GetTargetFilename(LPCTSTR targetpath, DLLSourceMode mode)
    {
     if(mode == UseKnownDll)
        { /* known DLL */
         CString name = GetLoadName(mode);
         CString target = targetpath;
         AppendDir(target, name);
         return target;
        } /* known DLL */
     else
     if(mode == UseInternalDll)
        { /* internal DLL */
         CString target = targetpath;
         AppendDir(target, DLL_NAME);
         return target;
        } /* internal DLL */
     else
        { /* should not happen */
         ASSERT(FALSE);
         return _T("");
        } /* should not happen */
    } // CLoadLibraryExplorerDlg::GetTargetFilename

/****************************************************************************
*                    CLoadLibraryExplorerDlg::ExpandString
* Inputs:
*       const CString & name: Name using perhaps %env% names
* Result: CString
*       The expanded name
****************************************************************************/

CString CLoadLibraryExplorerDlg::ExpandString(const CString & name)
    {
     TCHAR expanded[MAX_PATH];
     ExpandEnvironmentStrings(name, expanded, MAX_PATH);
     return expanded;
    } // CLoadLibraryExplorerDlg::ExpandString

/****************************************************************************
*                    CLoadLibraryExplorerDlg:SaveKnownDll
* Result: void
*       
* Effect: 
*       Saves the current selection so it can be restored on a restart
****************************************************************************/

void CLoadLibraryExplorerDlg::SaveKnownDll()
    {
     RegistryString knowndll(IDS_REGISTRY_KNOWN_DLL);
     int n = c_KnownDlls.GetCurSel();
     ASSERT(n != CB_ERR);
     if(n == CB_ERR)
        { /* error */
         knowndll.Delete();
         return;
        } /* error */
     KnownDLLInfo * info =  c_KnownDlls.GetItemDataPtr(n);
     knowndll.value = info->name;
     VERIFY(knowndll.store());
    } // CLoadLibraryExplorerDlg:SaveKnownDll

void CLoadLibraryExplorerDlg::OnCbnSelendokKnownDlls()
    {
     SaveKnownDll();
     updateControls();
    }

void CLoadLibraryExplorerDlg::OnCbnSelchangeKnownDlls()
    {
     SaveKnownDll();
     updateControls();
    }


/****************************************************************************
*                       CLoadLibraryExplorerDlg::Alert
* Result: void
*       
* Effect: 
*       Indicates visually/aurally that there is an error
****************************************************************************/

void CLoadLibraryExplorerDlg::Alert()
    {
     MessageBeep(MB_ICONSTOP);
    } // CLoadLibraryExplorerDlg::Alert

/****************************************************************************
*          CLoadLibraryExplorerDlg::OnBnClickedTestSafeDllSearchMode
*
* Result: void
*       
* Effect: 
*       Sets up to test SafeDllSearchMode
****************************************************************************/

void CLoadLibraryExplorerDlg::OnBnClickedTestSafeDllSearchMode()
    {
     RegistryInt TestSafe(IDS_REGISTRY_TEST_SAFE_DLL_SEARCH_MODE);
     TestSafe.value = c_TestSafeDllSearchMode.GetCheck() == BST_CHECKED;
     TestSafe.store();
     updateControls();
    }

/****************************************************************************
*                 CLoadLibraryExplorerDlg::LaunchTestProgram
* Result: void
*       
* Effect: 
*       Launches the test program
****************************************************************************/

void CLoadLibraryExplorerDlg::LaunchTestProgram()
    {
     LogMessage(new TraceComment(TraceEvent::None, IDS_LINE));

     BOOL oldSearchMode;
     SetupContext(oldSearchMode);
     
     CString cmdline;
     cmdline.Format(_T("%p"), m_hWnd); // pass notification window

     CString exe = GetExeDir();
     exe += TEST_PROGRAM_NAME;

     SHELLEXECUTEINFO info = {sizeof(SHELLEXECUTEINFO)};
     //info.fMask
     info.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_FLAG_NO_UI;
     info.hwnd = m_hWnd;
     info.lpVerb = _T("open");
     info.lpFile = (LPCTSTR)exe;
     info.lpParameters = (LPCTSTR)cmdline;
     // info.lpDirecotry: NULL defaults to current working directory
     info.nShow = SW_SHOW;
     // info.hInstApp will be set to indicate failure, but GetLastError is better
     // info.lpIDList is NULL and is ignored
     // info.hKeyClass is NULL and is ignored
     // info.dwHotKey is 0 and is ignored
     // info.hProcess will be set to the process handle of the newly-created process
     

     BOOL lannched = ShellExecuteEx(&info);
     DWORD err = ::GetLastError(); // in case it failed

     UndoContext(oldSearchMode);
     
     if(!lannched)
        { /* failed */
         Alert();
         LogMessage(new TraceFormatError(TraceEvent::None, IDS_LAUNCH_FAILED, exe));
         LogMessage(new TraceFormatMessage(TraceEvent::None, err));
         c_Library.SetWindowText(ErrorString(err));
         return;
        } /* failed */

     c_Library.SetWindowText(_T(""));
     // We have a running process.  Wait for it to terminate
     
     WaiterParameters * parameters = new WaiterParameters(info.hProcess, m_hWnd);
     launching = TRUE;
     updateControls();
     
     AfxBeginThread(waiter, parameters);
    } // CLoadLibraryExplorerDlg::LaunchTestProgram

/****************************************************************************
*                       CLoadLibraryExplorerDlg::waiter
* Inputs:
*       LPVOID p: (LPVOID)(WaiterParameters *)
* Result: UINT
*       logically void, 0, always
* Effect: 
*       Waits for the process to finish
****************************************************************************/

/* static */ UINT CLoadLibraryExplorerDlg::waiter(LPVOID p)
    {
     WaiterParameters * wp = (WaiterParameters *)p;

     switch(::WaitForSingleObject(wp->process, INFINITE))
        { /* waitfor */
         case WAIT_OBJECT_0:
            break;
         default:
            ASSERT(FALSE);
            return 0;
        } /* waitfor */
     DWORD exitcode;
     ::GetExitCodeProcess(wp->process, &exitcode);
     if(exitcode != 0)
        { /* process bailed out */
         ::PostMessage(wp->hWnd, UWM_PROCESS_TERMINATED, exitcode, (LPARAM)wp->process);
         delete wp;
         return 0;
        } /* process bailed out */

     ::PostMessage(wp->hWnd, UWM_PROCESS_TERMINATED, 0, (LPARAM)wp->process);
     delete wp;
     return 0;
    } // CLoadLibraryExplorerDlg::waiter

/****************************************************************************
*                CLoadLibraryExplorerDlg::OnProcessTerminated
* Inputs:
*       WPARAM: (WPARAM) exit code of process
*       LPARAM: (LPARAM)(HANDLE) process handle
* Result: LRESULT
*       Logically void, 0, always
* Effect: 
*       Logs the fact that the process terminated
****************************************************************************/

LRESULT CLoadLibraryExplorerDlg::OnProcessTerminated(WPARAM wParam, LPARAM lParam)
    {
     launching = FALSE;
     if(wParam == 0)
        { /* successful termination */
         LogMessage(new TraceFormatComment(TraceEvent::None, IDS_PROCESS_TERMINATED_OK, lParam));
        } /* successful termination */
     else
        { /* failure of process */
         LogMessage(new TraceFormatError(TraceEvent::None, IDS_PROCESS_TERMINATED_BADLY, lParam, wParam));
        } /* failure of process */
     updateControls();
     return 0;
    } // CLoadLibraryExplorerDlg::OnProcessTerminated

/****************************************************************************
*                     CLoadLibraryExplorerDlg::OnCopyData
* Inputs:
*       WPARAM: ignored
*       LPARAM: (LPARAM)(COPYDATASTRUCT *) data passed
* Result: LRESULT
*       (LRESULT)(BOOL)TRUE, if message is recognized
* Effect: 
*       Handles the WM_COPYDATA
****************************************************************************/

BOOL CLoadLibraryExplorerDlg::OnCopyData(CWnd* pWnd, COPYDATASTRUCT* cds)
    {
     CopyPacket * cp = (CopyPacket *)cds->lpData;

     if(!cp->Same(testguid))
        { /* error */
         Alert();
         LogMessage(new TraceError(TraceEvent::None, IDS_BAD_SIGNATURE));
         return FALSE;
        } /* error */

     BOOL result = TRUE;
     switch(cds->dwData)
        { /* data type */
         case SAFEDLL_NAME:
            LogMessage(new TraceFormatComment(TraceEvent::None, IDS_PROCESS_LOADED, (LPCTSTR)cp->data));
            c_Library.SetWindowText((LPCTSTR)cp->data);
            SetState(CString((LPCTSTR)cp->data));
            break;
         case SAFEDLL_ERROR:
            { /* failed */
             Alert();
             LogMessage(new TraceFormatError(TraceEvent::None, IDS_PROCESS_ERROR));
             DWORD err = *(DWORD *)cp->data;
             LogMessage(new TraceFormatMessage(TraceEvent::None, err));
             c_Library.SetWindowText(ErrorString(err));
            } /* failed */
            break;
         default:
             result = FALSE;
             break;
        } /* data type */
     return result;
    }
