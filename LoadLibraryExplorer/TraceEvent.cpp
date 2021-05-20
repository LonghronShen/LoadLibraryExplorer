/*****************************************************************************
*           Change Log
*  Date     | Change
*-----------+-----------------------------------------------------------------
* 24-Jan-00 | Created
*****************************************************************************/
#include "stdafx.h"
#include "TraceEvent.h"
#include "resource.h"

IMPLEMENT_DYNAMIC(TraceEvent, CObject)
IMPLEMENT_DYNAMIC(TraceComment, TraceEvent)
IMPLEMENT_DYNAMIC(TraceFormatComment, TraceComment)
IMPLEMENT_DYNAMIC(TraceError, TraceEvent)
IMPLEMENT_DYNAMIC(TraceFormatError, TraceError)
IMPLEMENT_DYNAMIC(TraceFormatMessage, TraceError)
IMPLEMENT_DYNAMIC(TraceData, TraceEvent)

CFont * TraceComment::font = NULL;
CFont * TraceError::font = NULL;
CFont * TraceData::font = NULL;

UINT TraceEvent::counter = 0;

/****************************************************************************
*                            TraceEvent::showfile
* Result: CString
*       Nicely formatted string for writing to file
* Effect: 
*       Prepares a line for file output; uses fixed-width fields for the prefix
* Notes:
*       Cleverly uses the virtual show() method to get the rest of the string
****************************************************************************/

CString TraceEvent::showfile()
    {
     CString t;
     if(filterID != None)
        t.Format(_T("%04d  "), filterID);
     else
        t = _T("      ");

     CString s;
     s.Format(_T("%06d| "), seq);
     return s + t + showThread() + show();
     
    }

/****************************************************************************
*                              TraceEvent::show
* Result: CString
*       
* Effect: 
*       Formats the base string for the trace event
****************************************************************************/

CString TraceEvent::show()
    {
     CString s; // = time.Format(_T("%H:%M:%S  "));
     return s;
    }

/****************************************************************************
*                             TraceEvent::showID
* Result: CString
*       The trace ID, or an empty string
* Effect: 
*       Formats the trace ID.  Connectionless events have no id
****************************************************************************/

CString TraceEvent::showID()
    {
     CString t;
     if(filterID != None)
        t.Format(_T("%d  "), filterID);
     else
        t = _T("");

     return t;
    }

/****************************************************************************
*                       CString TraceEvent::showThread
* Result: CString
*       Thread ID, formatted
****************************************************************************/

CString TraceEvent::showThread()
    {
     CString s;
     s.Format(_T("%08x "), threadID);
     return s;
    }

/****************************************************************************
*                             TraceEvent::cvtext
* Inputs:
*       CString s: String to convert
* Result: CString
*       String formatted for printout
****************************************************************************/

CString TraceEvent::cvtext(CString s)
    {
     CString str;
     for(int i = 0; i < s.GetLength(); i++)
        { /* convert to printable */
         str += cvtext(s[i]);
        } /* convert to printable */
     return str;     
    }

/****************************************************************************
*                            TraceEvent::cvtext
* Inputs:
*       TCHAR ch: Character to display (8-bit or 16-bit)
* Result: CString
*       Displayable string
****************************************************************************/

CString TraceEvent::cvtext(TCHAR ch)
    {
     BYTE ch8 = (BYTE)ch;
     CString s;

#ifdef _UNICODE
     BYTE ucp = (BYTE)(ch >> 8);
     if(ucp != 0)
        s.Format(_T(".\x%02x."), ucp);
#endif     
     if(ch8 >= ' ' && ch8 < 0177)
        { /* simple char */
         return CString((TCHAR)ch8); // NYI: handle Unicode
        } /* simple char */

     switch(ch8)
        { /* ch */
         case '\a':
                 return CString(_T("\\a"));
         case '\b':
                 return CString(_T("\\b"));
         case '\f':
                 return CString(_T("\\f"));
         case '\n':
                 return CString(_T("\\n"));
         case '\r':
                 return CString(_T("\\r"));
         case '\t':
                 return CString(_T("\\t"));
         case '\v':
                 return CString(_T("\\v"));
         case '\\':
                 return CString(_T("\\\\"));
         case '\"':
                 return CString(_T("\\\""));
         default:
                 s.Format(_T("\\x%02x"), ch);
                 return s;
        } /* ch */
    }

/****************************************************************************
*                    TraceFormatComment::TraceFormatComment
* Inputs:
*       UINT id: ID for line
*       CString fmt: Formatting string
*       ...: Arguments for formatting
* Effect: 
*       Constructor
****************************************************************************/

TraceFormatComment::TraceFormatComment(UINT id, CString fmt, ...) : TraceComment(id, CString(""))
    {
     va_list args;
     va_start(args, fmt);
     LPTSTR p = comment.GetBuffer(2048);
     _vstprintf(p, fmt, args);
     comment.ReleaseBuffer();
     va_end(args);
    } // TraceFormatComment::TraceFormatComment

/****************************************************************************
*                   TraceFormatComment::TraceFormatComment
* Inputs:
*       UINT id: Line number
*       UINT fmt: ID of formatting string
*       ...: Arguments
* Effect: 
*       Creates a trace comment entry, formatted
****************************************************************************/

TraceFormatComment::TraceFormatComment(UINT id, UINT fmt, ...) : TraceComment(id, CString(""))
    {
     va_list args;
     va_start(args, fmt);
     CString fmtstr;
     fmtstr.LoadString(fmt);
     LPTSTR p = comment.GetBuffer(2048);
     _vstprintf(p, fmtstr, args);
     comment.ReleaseBuffer();
     va_end(args);
     
    } // TraceFormatComment::TraceFormatComment

/****************************************************************************
*                             TraceComment::show
* Result: CString
*       Display string
****************************************************************************/

CString TraceComment::show()
    {
     return TraceEvent::show() + comment;
    }

/****************************************************************************
*                      TraceFormatError::TraceFormatError
* Inputs:
*       UINT id: ID for line
*       CString fmt: Formatting string
*       ...: Arguments for formatting
* Effect: 
*       Constructor
****************************************************************************/

TraceFormatError::TraceFormatError(UINT id, CString fmt, ...) : TraceError(id, CString(""))
   {
    va_list args;
    va_start(args, fmt);
    LPTSTR p = Error.GetBuffer(2048);
    _vstprintf(p, fmt, args);
    Error.ReleaseBuffer();
    va_end(args);
   } // TraceFormatError::TraceFormatError

/****************************************************************************
*                   TraceFormatError::TraceFormatError
* Inputs:
*       UINT id: Line number
*       UINT fmt: ID of formatting string
*       ...: Arguments
* Effect: 
*       Creates a trace Error entry, formatted
****************************************************************************/

TraceFormatError::TraceFormatError(UINT id, UINT fmt, ...) : TraceError(id, CString(""))
   {
    va_list args;
    va_start(args, fmt);
    CString fmtstr;
    fmtstr.LoadString(fmt);
    LPTSTR p = Error.GetBuffer(2048);
    _vstprintf(p, fmtstr, args);
    Error.ReleaseBuffer();
    va_end(args);
   } // TraceFormatError::TraceFormatError

/****************************************************************************
*                             TraceError::show
* Result: CString
*       Display string
****************************************************************************/

CString TraceError::show()
    {
     return TraceEvent::show() + Error;
    }

/****************************************************************************
*                     TraceFormatMessage::TraceFormatMessage
* Inputs:
*       UINT id:
*       DWORD error:
* Effect: 
*       Constructor
****************************************************************************/

TraceFormatMessage::TraceFormatMessage(UINT id, DWORD err) : TraceError(id)
    {
     LPTSTR s;
     if(::FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER |
                        FORMAT_MESSAGE_FROM_SYSTEM,
                        NULL,
                        err,
                        0,
                        (LPTSTR)&s,
                        0,
                        NULL) == 0)
        { /* failed */
         // See if it is a known error code
         CString t; 
         t.LoadString(err);
         if(t.GetLength() == 0 || t[0] != _T('®'))     
            { /* other error */                        
             CString fmt;                               
             fmt.LoadString(IDS_TRACELOG_UNKNOWN_ERROR);
             t.Format(fmt, err, err);                   
            } /* other error */                         
         else                                           
         if(t.GetLength() > 0 && t[0] == _T('®'))       
            { /* drop prefix */                         
             t = t.Mid(1);                              
            } /* drop prefix */                         
         Error = t;
        } /* failed */
     else
        { /* success */
         LPTSTR p = _tcschr(s, _T('\r'));
         if(p != NULL)
            { /* lose CRLF */
             *p = _T('\0');
            } /* lose CRLF */
         Error = s;
         ::LocalFree(s);
        } /* success */
    }

/****************************************************************************
*                       TraceData::TraceData
* Inputs:
*       UINT id: numeric ID to display, or TraceEvent::None
*       LPBYTE data: Pointer to binary data
*       UINT length: Length of string currently buffered
* Effect: 
*       Constructor
****************************************************************************/

TraceData::TraceData(UINT id, LPBYTE indata, UINT n) : TraceEvent(id)
   {
    length = n;
    binary = TRUE;
    data = new BYTE[length];
    if(data != NULL)
       memcpy(data, indata, n);
   }

/****************************************************************************
*                       TraceData::TraceData
* Inputs:
*       UINT id: Numeric ID to display, or TraceEvent::None
*       const CString & s: String data
* Effect: 
*       Constructor
****************************************************************************/

TraceData::TraceData(UINT id, const CString & s) : TraceEvent(id)
   {
    length = s.GetLength() * sizeof(TCHAR);
    binary = FALSE;
    this->data = new BYTE[length + sizeof(TCHAR)];
    if(this->data != NULL)
       memcpy(this->data, (LPCTSTR)s, length + sizeof(TCHAR));
   }

/****************************************************************************
*                            TraceData::TraceData
* Inputs:
*       UINT id: ID of connection
*       LPCTSTR s: String data
* Effect: 
*       Constructs a text data element
****************************************************************************/

TraceData::TraceData(UINT id, LPCTSTR s) : TraceEvent(id)
    {
     length = lstrlen(s) * sizeof(TCHAR);
     binary = FALSE;
     data = new BYTE[length + sizeof(TCHAR)];
     if(data != NULL)
        lstrcpy((LPTSTR)data, s);
    } // TraceData::TraceData

/****************************************************************************
*                               TraceData::show
* Result: CString
*       The data, formatted according to the type of data stored
****************************************************************************/

CString TraceData::show()
    {
     if(binary)
        { /* binary data */
         CString t;
         CString str;
         for(UINT i = 0; i < length; i++)
            { /* convert */
             t.Format(_T("%02x "), data[i]);
             str += t;
            } /* convert */
         return str;
        } /* binary data */
     else
        { /* text data */
         CString str = cvtext((LPCTSTR)data);
         return TraceEvent::show() + str;
        } /* text data */
    } // TraceData::show
