// State.cpp : implementation file
//

#include "stdafx.h"
#include "State.h"


// CState

IMPLEMENT_DYNAMIC(CState, CStatic)
CState::CState()
   {
    state = FALSE;
    br.CreateSolidBrush(RGB(255, 255, 0));
   }

CState::~CState()
   {
   }


BEGIN_MESSAGE_MAP(CState, CStatic)
    ON_WM_CTLCOLOR_REFLECT()
END_MESSAGE_MAP()



// CState message handlers


HBRUSH CState::CtlColor(CDC* pDC, UINT /*nCtlColor*/)
    {
     pDC->SetBkMode(TRANSPARENT);
     if(state)
        return (HBRUSH)br;
     else
        return NULL;
    }

/****************************************************************************
*                            CState::SetHighlight
* Inputs:
*       BOOL mode: Highlight mode
* Result: void
*       
* Effect: 
*       Sets the highlight mode
****************************************************************************/

void CState::SetHighlight(BOOL mode)
    {
     if(state == mode)
        return;
     state = mode;

     if(GetSafeHwnd() != NULL)
        Invalidate();
    } // CState::SetHighlight
