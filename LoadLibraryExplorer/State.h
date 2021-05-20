#pragma once


// CState

class CState : public CStatic
{
        DECLARE_DYNAMIC(CState)

public:
        CState();
        virtual ~CState();
        void SetHighlight(BOOL mode);

protected:
        DECLARE_MESSAGE_MAP()
        BOOL state;
        CBrush br;
public:
    afx_msg HBRUSH CtlColor(CDC* /*pDC*/, UINT /*nCtlColor*/);
};


