// LoadLibraryExplorerDlg.h : header file
//

#pragma once
#include "atltrace.h"
#include "TraceEvent.h"
#include "Trace.h"
#include "afxwin.h"
#include "state.h"
#include "SmartDropdown.h"
#include "KnownDLLCombo.h"

// CLoadLibraryExplorerDlg dialog
class CLoadLibraryExplorerDlg : public CDialog
{
// Construction
public:
        CLoadLibraryExplorerDlg(CWnd* pParent = NULL);  // standard constructor

// Dialog Data
        enum { IDD = IDD_LOADLIBRARYEXPLORER_DIALOG };

        protected:
        virtual void DoDataExchange(CDataExchange* pDX);        // DDX/DDV support


// Implementation
protected: // variables
   HICON m_hIcon;
   BOOL launching;
   CString TestDir;  // random test directory
   CString DLLName;
   CString DLLSource;

   //****************************************************************
   // Directory names
   //****************************************************************
   CString AppPathDir;   // App Path directory
   CString WorkingDir;  // current working directory
   CString System32Dir;
   CString SystemDir;
   CString WindowsDir;
   CString AlternateDir;

   //****************************************************************
   // The PATH string
   //****************************************************************
   CString PathVariable;


protected: // message handlers
        
   virtual BOOL OnInitDialog();
   afx_msg HCURSOR OnQueryDragIcon();
   afx_msg LRESULT OnLog(WPARAM, LPARAM);
   afx_msg LRESULT OnProcessTerminated(WPARAM, LPARAM);
   afx_msg void OnBnClickedAlternate();
   afx_msg void OnBnClickedAppPath();
   afx_msg void OnBnClickedExecutable();
   afx_msg void OnBnClickedForce();
   afx_msg void OnBnClickedForceAlternate();
   afx_msg void OnBnClickedForceAppPath();
   afx_msg void OnBnClickedForceExecutable();
   afx_msg void OnBnClickedForceSystem();
   afx_msg void OnBnClickedForceSystem32();
   afx_msg void OnBnClickedForceWindows();
   afx_msg void OnBnClickedForceWorking();
   afx_msg void OnBnClickedLoad();
   afx_msg void OnBnClickedLoadWithAlteredSearchPath();
   afx_msg void OnBnClickedSystem();
   afx_msg void OnBnClickedSystem32();
   afx_msg void OnBnClickedUseLoadlibrary();
   afx_msg void OnBnClickedUseLoadlibraryex();
   afx_msg void OnBnClickedUseShellexecute();
   afx_msg void OnBnClickedWindows();
   afx_msg void OnBnClickedWorking();
   afx_msg void OnClose();
   afx_msg void OnGetMinMaxInfo(MINMAXINFO* lpMMI);
   afx_msg void OnPaint();
   afx_msg void OnSize(UINT nType, int cx, int cy);
   afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
   afx_msg void OnBnClickedUseSelectedSafedllsearchmode();
   afx_msg void OnBnClickedUseDefaultSearchMode();
   afx_msg void OnBnClickedDesiredSafedllsearchmode();
   afx_msg void OnBnClickedEnum();
   afx_msg void OnBnClickedRelaunch();
   afx_msg void OnBnClickedUseTheDll();
   afx_msg void OnBnClickedKnownDll();
   afx_msg void OnCbnSelendokKnownDlls();
   afx_msg void OnCbnSelchangeKnownDlls();
   afx_msg void OnBnClickedRedirectNone();
   afx_msg void OnBnClickedRedirectExecutable();
   afx_msg void OnBnClickedRedirectLocal();
   afx_msg void OnBnClickedTestSafeDllSearchMode();
   afx_msg BOOL OnCopyData(CWnd* pWnd, COPYDATASTRUCT* pCopyDataStruct);
   DECLARE_MESSAGE_MAP()


protected: // virtual methods
   virtual void OnOK();
   virtual void OnCancel();

protected: // methods
   static int CDECL stringsort(const void * e1, const void * e2);
   typedef enum { UseInternalDll, UseKnownDll } DLLSourceMode;

   void Alert();
   void InstallInRedirectDir();
   void LogMessage(TraceEvent * e);
   void LoadKnownDlls();
   BOOL MakeSubdirectory(LPCTSTR name, CString & dir);
   void ConditionalEnable(CButton & ctl, LPCTSTR expanded);
   BOOL SetRegistryAppPath(LPCTSTR module);
   BOOL SetAllRegistryAppPaths();
   BOOL IsInPath(CString name);
   CString GetRegistryAppPathKey(LPCTSTR module);
   void SaveOptions();
   BOOL cwd();
   BOOL LoadDLL();
   BOOL DoLoad();
   CString GetAppPath();
   CString FormAppPathKey();
   void updateControls();
   CString GetLoadPath(DLLSourceMode mode);
   BOOL CopyDLLs();
   BOOL EraseDLLs(DLLSourceMode mode);
   BOOL RemoveDLL(LPCTSTR target, CState & state, DLLSourceMode mode);
   CString GetExeDir();
   BOOL InstallDLL(LPCTSTR targetpath, CState & state, DLLSourceMode mode);
   void Resize(CWnd & wnd, BOOL toBottom = FALSE, int adjust = 0);
   void ShowPath();
   void AppendDir(CString & path, LPCTSTR name);
   void SetPath();
   void SetState(CString name);
   void CheckPath(CString target, CButton & button, CState & state);
   CString ExpandString(const CString & s);
   CString GetTargetFilename(LPCTSTR targetpath, DLLSourceMode mode);
   CString GetSourceFilename(DLLSourceMode mode);
   void SaveKnownDll();
   CString GetLoadName(DLLSourceMode mode);
   void ClearLoadState();
   DLLSourceMode CurrentMode;
   void SetCurrentMode();
   typedef BOOL (WINAPI * SetDllDirectoryProc)(LPCTSTR path);
   SetDllDirectoryProc SetDllDirectory;
   void SetRedirection();
   void DeleteRedirectionDirectory();
   void CreateRedirectionDirectory();
   void DeleteRedirectionFile();
   void CreateRedirectionFile();
   CString GetLocalName(BOOL fullpath = FALSE);
   void LaunchTestProgram();
   void MaybeExtractDLL();
   void ExtractDLL(FILETIME t);
   void SetupContext(BOOL & oldSafeMode);
   void UndoContext(BOOL oldSafeMode);

   static UINT waiter(LPVOID parms);

   class WaiterParameters {
       public:
          WaiterParameters(HANDLE p, HWND h) {process = p; hWnd = h; }
          HANDLE process;
          HWND hWnd;
   };

// protected: // enum classes   
   class Force {
       public:
          typedef enum ForceType { Alternate = 1, AppPath, Executable, System32, System, Windows, Working };
   };

   class Load {
       public:
          typedef enum LoadType { LoadLibrary = 1, LoadLibraryEx };
   };

   class SearchMode {
       public:
          typedef enum Search { UseDefaultSafeDllSearchMode = 0, UseSelectedSafeDllSearchMode };
   };

   class Redirection {
       public:
          typedef enum RedirectionMode { None, File, Directory };
   };

protected: // control variables
   BOOL HasAppPath;
   CTraceList c_Log;

   CStatic c_DllLocations;
   CButton c_SYSTEM32;
   CButton c_WINDOWS;
   CButton c_SYSTEM;
   CButton c_Working;
   CButton c_AppPath;
   CButton c_Executable;
   CButton c_Alternate;
   CButton c_Local;

   CCheckListBox c_Path;
   CEdit c_Library;
   CEdit c_LoadLibrary;
   CButton c_LOAD_WITH_ALTERED_SEARCH_PATH;
   CButton c_UseLoadLibrary;
   CButton c_UseLoadLibraryEx;
   CButton c_UseShellExecute;

   CButton c_Force;
   CButton c_ForceSystem32;
   CButton c_ForceWindows;
   CButton c_ForceSystem;
   CButton c_ForceWorking;
   CButton c_ForceAppPath;
   CButton c_ForceExecutable;
   CButton c_ForceAlternate;

   CButton c_UseSelectedSafeDllSearchMode;
   CButton c_UseExistingSafeDllSearchMode;
   CButton c_DesiredSafeDllSearchMode;
   CButton c_DefaultSafeDllSearchMode;
   CButton c_Alphabetical;

   CButton c_Relaunch;
   CStatic c_Message;

   CState c_StateSystem32;
   CState c_StateWindows;
   CState c_StateSystem;
   CState c_StateWorking;
   CState c_StateAppPath;
   CState c_StateExecutable;
   CState c_StateAlternate;

   CButton c_UseInternalDll;
   CButton c_UseKnownDll;
   CKnownDLLCombo c_KnownDlls;
   CStatic c_DllDirectory;

   CStatic c_RedirectCaption;
   CButton c_RedirectExecutable;
   CButton c_RedirectDirectory;
   CButton c_RedirectNone;
   CState c_StateLocal;
   CStatic c_MinFrame;
   CButton c_TestSafeDllSearchMode;
   CStatic c_DialogSelectionCaption;
   CStatic c_LoadMethodCaption;
   CButton c_Load;
};
