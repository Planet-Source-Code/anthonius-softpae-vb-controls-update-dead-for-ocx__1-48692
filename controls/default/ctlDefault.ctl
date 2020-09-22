VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   ScaleHeight     =   3375
   ScaleWidth      =   5040
   ToolboxBitmap   =   "ctlDefault.ctx":0000
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Const OFFSET_P1   As Long = 9
Private Const OFFSET_P2   As Long = 22
Private Const OFFSET_P3   As Long = 37
Private Const OFFSET_P4   As Long = 51
Private Const OFFSET_P5   As Long = 69
Private Const OFFSET_P6   As Long = 141
Private Const OFFSET_P7   As Long = 146
Private Const OFFSET_P8   As Long = 154
Private Const OFFSET_P9   As Long = 169
Private Const OFFSET_PA   As Long = 183
Private Const OFFSET_PB   As Long = 201
Private Const OFFSET_PC   As Long = 250
Private Const OFFSET_PD   As Long = 260
Private Const ARRAY_LB    As Long = 1

Private Type tCode
    Buf(ARRAY_LB To 272)    As Byte
End Type

Private Type tCodeBuf
    Code                    As tCode
End Type

Private CodeBuf           As tCodeBuf
Private nBreakGate        As Long
Private nMsgCntB          As Long
Private nMsgCntA          As Long
Private aMsgTblB()        As WinSubHook.eMsg
Private aMsgTblA()        As WinSubHook.eMsg
Private hWndSubclass      As Long
Private nWndProcSubclass  As Long
Private nWndProcOriginal  As Long

Implements WinSubHook.iSubclass

'***********************************************************************************
'add your declaration here !!


Private Sub subInitialize()

    Const OPS As String = "558BEC83C4F85756BE_patch1_33C08945FC8945F8B90000000083F900746183F9FF740CBF000000008B450CF2AF755033C03D_patch4_740B833E007542C70601000000BA_patch5_8B0283F8000F84A50000008D4514508D4510508D450C508D4508508D45FC508D45F8508B0252FF5020C706000000008B45F883F8007570FF7514FF7510FF750CFF750868_patch6_E8_patch7_8945FCB90000000083F900744D83F9FF740CBF000000008B450CF2AF753C33C03D_patchA_740B833E00752EC70601000000BA_patchB_8B0283F8007425FF7514FF7510FF750CFF75088D45FC508B0252FF501CC706000000005E5F8B45FCC9C2100068_patchC_6AFCFF7508E8_patchD_33C08945FCEBE190"
    
    Dim i As Long, j As Long, nIDE As Long

    With CodeBuf.Code

        j = 1

        For i = ARRAY_LB To UBound(.Buf)

            .Buf(i) = Val("&H" & Mid$(OPS, j, 2))
            j = j + 2

        Next i

        nWndProcSubclass = VarPtr(.Buf(ARRAY_LB))

    End With
  
    nIDE = InIDE

    Call PatchVal(OFFSET_P1, VarPtr(nBreakGate))
    Call PatchVal(OFFSET_P4, nIDE)
    Call PatchRel(OFFSET_P7, AddrFunc("CallWindowProcA"))
    Call PatchVal(OFFSET_PA, nIDE)
    Call PatchRel(OFFSET_PD, AddrFunc("SetWindowLongA"))

End Sub

Public Sub AddMsg(uMsg As WinSubHook.eMsg, When As WinSubHook.eMsgWhen)

    If When = WinSubHook.MSG_BEFORE Then
        Call AddMsgSub(uMsg, aMsgTblB, nMsgCntB, When)
    Else
        Call AddMsgSub(uMsg, aMsgTblA, nMsgCntA, When)
    End If

End Sub

Public Function CallOrigWndProc(ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long) As Long

    If hWndSubclass <> 0 Then
        CallOrigWndProc = WinSubHook.CallWindowProc(nWndProcOriginal, hWndSubclass, uMsg, wParam, lParam)
    Else
        Debug.Assert False
    End If

End Function

Public Sub DelMsg(uMsg As WinSubHook.eMsg, When As WinSubHook.eMsgWhen)

    If When = WinSubHook.MSG_BEFORE Then
        Call DelMsgSub(uMsg, aMsgTblB, nMsgCntB, When)
    Else
        Call DelMsgSub(uMsg, aMsgTblA, nMsgCntA, When)
    End If

End Sub

Public Sub Subclass(hwnd As Long, Owner As WinSubHook.iSubclass)

    Debug.Assert (hWndSubclass = 0)
    Debug.Assert IsWindow(hwnd)
  
    hWndSubclass = hwnd
    nWndProcOriginal = WinSubHook.SetWindowLong(hwnd, WinSubHook.GWL_WNDPROC, nWndProcSubclass)
    Debug.Assert nWndProcOriginal
  
    Call PatchVal(OFFSET_P5, ObjPtr(Owner))
    Call PatchVal(OFFSET_P6, nWndProcOriginal)
    Call PatchVal(OFFSET_PB, ObjPtr(Owner))
    Call PatchVal(OFFSET_PC, nWndProcOriginal)

End Sub

Public Sub UnSubclass()

    If hWndSubclass <> 0 Then

        Call PatchVal(OFFSET_P2, 0)
        Call PatchVal(OFFSET_P8, 0)
        Call WinSubHook.SetWindowLong(hWndSubclass, WinSubHook.GWL_WNDPROC, nWndProcOriginal)
        hWndSubclass = 0
        nMsgCntB = 0
        nMsgCntA = 0

    End If

End Sub

Private Sub AddMsgSub(uMsg As WinSubHook.eMsg, aMsgTbl() As WinSubHook.eMsg, nMsgCnt As Long, When As WinSubHook.eMsgWhen)

    Dim nEntry  As Long, nOff1   As Long, nOff2   As Long
  
    If uMsg = WinSubHook.ALL_MESSAGES Then

        nMsgCnt = -1

    Else

        For nEntry = ARRAY_LB To nMsgCnt

            Select Case aMsgTbl(nEntry)

                Case -1
                    aMsgTbl(nEntry) = uMsg
                    Exit Sub

                Case uMsg
                    Exit Sub

            End Select

        Next nEntry

        ReDim Preserve aMsgTbl(ARRAY_LB To nEntry)
        nMsgCnt = nEntry
        aMsgTbl(nEntry) = uMsg

    End If
  
    If When = WinSubHook.MSG_BEFORE Then
        nOff1 = OFFSET_P2
        nOff2 = OFFSET_P3
    Else
        nOff1 = OFFSET_P8
        nOff2 = OFFSET_P9
    End If

    Call PatchVal(nOff1, nMsgCnt)
    Call PatchVal(nOff2, AddrMsgTbl(aMsgTbl))

End Sub

Private Sub DelMsgSub(uMsg As WinSubHook.eMsg, aMsgTbl() As WinSubHook.eMsg, nMsgCnt As Long, When As WinSubHook.eMsgWhen)

    Dim nEntry As Long
  
    If uMsg = WinSubHook.ALL_MESSAGES Then

        nMsgCnt = 0

        If When = WinSubHook.MSG_BEFORE Then
            nEntry = OFFSET_P2
        Else
            nEntry = OFFSET_P8
        End If

        Call PatchVal(nEntry, 0)

    Else

        For nEntry = ARRAY_LB To nMsgCnt

            If aMsgTbl(nEntry) = uMsg Then
                aMsgTbl(nEntry) = -1
                Exit For
            End If

        Next nEntry

    End If

End Sub

Private Function AddrFunc(sProc As String) As Long
    AddrFunc = WinSubHook.GetProcAddress(WinSubHook.GetModuleHandle("user32"), sProc)
End Function

Private Function AddrMsgTbl(aMsgTbl() As WinSubHook.eMsg) As Long
    On Error Resume Next
    AddrMsgTbl = VarPtr(aMsgTbl(ARRAY_LB))
    On Error GoTo 0
End Function

Private Sub PatchVal(nOffset As Long, nValue As Long)
    Call WinSubHook.CopyMemory(ByVal (nWndProcSubclass + nOffset), nValue, 4)
End Sub

Private Sub PatchRel(nOffset As Long, nTargetAddr As Long)
    Call WinSubHook.CopyMemory(ByVal (nWndProcSubclass + nOffset), nTargetAddr - nWndProcSubclass - nOffset - 4, 4)
End Sub

Private Function InIDE() As Long

    Static Value As Long
  
    If Value = 0 Then

        Value = 1
        Debug.Assert True Or InIDE()
        InIDE = Value - 1

    End If
  
    Value = 0

End Function

Private Sub iSubclass_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long)

    'funkcia volaná po hlavnej (systémovej) wnd funkcii
    'not used in this

End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam As Long)

    On Error Resume Next
    
    'funkcia volaná pred hlavnou (systémovou) wnd funkcii
    
    Select Case uMsg

        Case WM_SHOWWINDOW
            
            If wParam = 1 Then
                'If InIDE = False Then Call SetTimer(hwnd, 101, 100, ByVal 0&)
            ElseIf wParam = 0 Then
                'Call KillTimer(hwnd, 101)
            End If
        
    End Select
    bHandled = True

End Sub

Private Sub UserControl_Initialize()

    On Error Resume Next

    Call subInitialize
    Call Subclass(UserControl.hwnd, Me)
    
    'add messages to subclass
    Call AddMsg(WM_SHOWWINDOW, MSG_BEFORE)

End Sub

Private Sub UserControl_Terminate()

    On Error Resume Next

    If hWndSubclass <> 0 Then
        Call UnSubclass
    End If
    
End Sub

'***********************************************************************************
'add your properties and metthots here !!

