
#pragma once


/****************************************************************************
*                           class TraceEvent
* Inputs:
*       UINT id: numeric ID, or TraceEvent::None
* Notes:
*       This constructor is used only by its subclasses.  This is a virtual
*       class and is never instantiated on its own
****************************************************************************/

class TraceEvent : public CObject {
    DECLARE_DYNAMIC(TraceEvent)
protected:
    static UINT counter;
    TraceEvent(UINT id) { time = CTime::GetCurrentTime(); 
                          filterID = id; 
                          threadID = ::GetCurrentThreadId(); 
                          seq = counter++;
                        }
public:
    enum { None = -1 };
    virtual ~TraceEvent() { }
    virtual COLORREF textcolor() = 0;
    virtual CString show();
    virtual CString showID();
    virtual CString showfile();
    virtual CString showThread();
    virtual int displayHeight() { return 1; }
    virtual UINT getID() {return filterID; }
    virtual CFont * getFont() { return NULL; }
    static CString cvtext(TCHAR ch); // useful for various subclasses
    static CString cvtext(CString s); // useful for various subclasses
protected:
    CTime time;
    UINT  filterID;
    DWORD threadID;
    UINT  seq;
};

/****************************************************************************
*                           class TraceComment
* Inputs:
*       UINT id: numeric ID, or TraceEvent::None
*       CString s: String to display
* Inputs:
*       UINT id: numeric ID, or TraceEvent::None
*       UINT u: IDS_ index of string to display
****************************************************************************/

class TraceComment: public TraceEvent {
    DECLARE_DYNAMIC(TraceComment)
public:
    TraceComment(UINT id, const CString & s) : TraceEvent(id) {comment = s; }
    TraceComment(UINT id, UINT u) : TraceEvent(id){ CString s; s.LoadString(u); comment = s; }
    virtual ~TraceComment() { }
    virtual COLORREF textcolor() { return RGB(0, 128, 0); }
    virtual CString show();
    virtual CFont * getFont() { return font; }
    static CFont * setFont(CFont * newfont) { CFont * old = font; font = newfont; return old; }
protected:
    static CFont *  font;
    CString comment;
   };

/****************************************************************************
*                          class TraceFormatComment
* Inputs:
*       UINT id: Line number
*       CString fmt: Formatting string
*       ...
****************************************************************************/

class TraceFormatComment : public TraceComment
   {
    DECLARE_DYNAMIC(TraceFormatComment)
    public:
       TraceFormatComment(UINT id, CString fmt, ...);
       TraceFormatComment(UINT id, UINT fmt, ...);
   }; // class TraceFormatComment

/****************************************************************************
*                           class TraceError
* Inputs:
*       UINT id: numeric ID, or TraceEvent::None
* Inputs:
*       UINT id: numeric ID, or TraceEvent::None
*       CString s: String to display
* Inputs:
*       UINT id: numeric ID, or TraceEvent::None
*       UINT u: IDS_ token index of string to display
* Notes:
*       The first form is not used directly, only by subclasses
****************************************************************************/

class TraceError: public TraceEvent {
    DECLARE_DYNAMIC(TraceError)
public:
    TraceError(UINT id, const CString & s) : TraceEvent(id) {Error = s; }
    TraceError(UINT id, UINT u) : TraceEvent(id){ CString s; s.LoadString(u); Error = s; }
    virtual ~TraceError() { }
    virtual COLORREF textcolor() { return RGB(128, 0, 0); }
    virtual CString show();
    virtual CFont * getFont() { return font; }
    static CFont * setFont(CFont * newfont) { CFont * old = font; font = newfont; return old; }
protected:
    TraceError(UINT id) : TraceEvent(id) {}
    CString Error;
    static CFont * font;
   };


/****************************************************************************
*                           class TraceFormatMessage
* Inputs:
*       UINT id: numeric ID, or TraceEvent::None
*       DWORD error: Error code from GetLastError()
****************************************************************************/

class TraceFormatMessage: public TraceError {
  DECLARE_DYNAMIC(TraceFormatMessage)
  public:
    TraceFormatMessage(UINT id, DWORD error);
    virtual ~TraceFormatMessage() { }
                                          };

/****************************************************************************
*                            class TraceFormatError
* Inputs:
*       UINT id: Line number
*       CString fmt: Formatting string
*       ...
****************************************************************************/

class TraceFormatError : public TraceError
   {
    DECLARE_DYNAMIC(TraceFormatError)
    public:
       TraceFormatError(UINT id, CString fmt, ...);
       TraceFormatError(UINT id, UINT fmt, ...);
   }; // class TraceFormatError

/****************************************************************************
*                           class TraceData
* Inputs:
*       UINT id: ID of message
*       LPBYTE data: 8-bit data string, possibly containing NUL bytes
*       UINT length: Length of data string
****************************************************************************/

class TraceData: public TraceEvent {
   DECLARE_DYNAMIC(TraceData)
    public:
      TraceData(UINT id, LPBYTE data, UINT length);
      TraceData(UINT id, LPCSTR data);
      TraceData(UINT id, const CString & s);
      virtual ~TraceData() {if(data != NULL) delete [] data; }
      virtual COLORREF textcolor() { return RGB(0, 0, 128); }
      virtual CString show();
      virtual CFont * getFont() { return font; }
      static CFont * setFont(CFont * newfont) { CFont * old = font; font = newfont; return old; }
    protected:
       BOOL binary;
       UINT length;
       LPBYTE data;
       static CFont * font;
};
