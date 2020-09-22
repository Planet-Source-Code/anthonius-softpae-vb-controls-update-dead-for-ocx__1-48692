Attribute VB_Name = "mdlRichEdit"
Option Explicit

Private Const OFS_MAXPATHNAME = 128

Public Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long 'lpOverlapped As OVERLAPPED) As Long
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long 'lpOverlapped As OVERLAPPED) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private m_sText As String
Private m_lPos As Long
Private m_lLen As Long
Private m_bFileMode As Boolean
Private m_lObj As Long

Public Property Let RichEdit(ByVal edtThis As RichEdit)

    On Error Resume Next

    m_lObj = ObjPtr(edtThis)

End Property

Public Property Get RichEdit() As RichEdit

    On Error Resume Next

    Dim rT As RichEdit
    
    If (m_lObj <> 0) Then
        CopyMemory rT, m_lObj, 4
        Set RichEdit = rT
        CopyMemory rT, 0&, 4
    End If

End Property

Public Sub ClearRichEdit()

    On Error Resume Next

    m_lObj = 0

End Sub

Public Property Let FileMode(ByVal bMode As Boolean)

    On Error Resume Next

    m_bFileMode = bMode

End Property

Public Property Get FileMode() As Boolean

    On Error Resume Next

    FileMode = m_bFileMode

End Property

Public Sub ClearStreamText()

    On Error Resume Next

    m_sText = ""

End Sub

Public Property Get StreamText() As String

    On Error Resume Next

    StreamText = m_sText

End Property

Public Property Let StreamText(ByRef sText As String)

    On Error Resume Next

    m_sText = sText
    m_lPos = 1
    m_lLen = Len(m_sText)

End Property

Public Function LoadCallBack(ByVal dwCookie As Long, ByVal lPtrPbBuff As Long, ByVal cb As Long, ByVal pcb As Long) As Long

    On Error Resume Next

    Dim sBuf As String
    Dim b() As Byte
    Dim lLen As Long
    Dim lRead As Long

    If (m_bFileMode) Then
        ReadFile dwCookie, ByVal lPtrPbBuff, cb, ByVal pcb, ByVal 0&
        CopyMemory lRead, ByVal pcb, 4
        If (lRead < cb) Then
            LoadCallBack = 0
        Else
            LoadCallBack = 0
        End If
    Else
        CopyMemory lRead, ByVal pcb, 4
        Debug.Print lRead, cb
        
        If (m_lLen - m_lPos > 0) Then
            If (m_lLen - m_lPos < cb) Then
                ReDim b(0 To (m_lLen - m_lPos)) As Byte
                b = StrConv(Mid$(m_sText, m_lPos), vbFromUnicode)
                lRead = m_lLen - m_lPos + 1
                CopyMemory ByVal lPtrPbBuff, b(0), lRead
                m_lPos = m_lLen + 1
            Else
                ReDim b(0 To cb - 1) As Byte
                b = StrConv(Mid$(m_sText, m_lPos, cb), vbFromUnicode)
                CopyMemory ByVal lPtrPbBuff, b(0), cb
                m_lPos = m_lPos + cb
                lRead = cb
            End If

            CopyMemory ByVal pcb, lRead, 4
            LoadCallBack = 0
        Else
            lRead = 0
            CopyMemory ByVal pcb, lRead, 4
            LoadCallBack = 0
        End If

    End If

End Function

Public Function SaveCallBack(ByVal dwCookie As Long, ByVal lPtrPbBuff As Long, ByVal cb As Long, ByVal pcb As Long) As Long

    On Error Resume Next

    Dim sBuf As String
    Dim b() As Byte
    Dim lLen As Long

    lLen = cb

    If (lLen > 0) Then
        If (m_bFileMode) Then
            WriteFile dwCookie, ByVal lPtrPbBuff, cb, ByVal pcb, ByVal 0&
        Else
            ReDim b(0 To lLen - 1) As Byte
            CopyMemory b(0), ByVal lPtrPbBuff, lLen
            sBuf = StrConv(b, vbUnicode)
            CopyMemory ByVal pcb, lLen, 4
            m_sText = m_sText & sBuf
            m_lPos = 1
            m_lLen = Len(m_sText)

        End If
    End If
    SaveCallBack = 0

End Function

