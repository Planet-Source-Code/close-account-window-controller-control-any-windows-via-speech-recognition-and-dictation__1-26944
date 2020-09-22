Attribute VB_Name = "modEnumChildren"
'  ___________________    ________________________________
' /                   \  /                                \
' | Window controller |--| By David Fiala djf1010@aol.com |
' \___________________/  \________________________________/
'
' Version 1.00   Released date: Sept. 03 2001

Option Explicit

Private Type udtChildEnum
    lngCHWND As Long
    strCText As String
End Type

Private aChildren() As udtChildEnum

Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const BM_CLICK = &HF5

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hwndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Any) As Long

Public Sub DoChildEnum(ByVal lngParentHandle As Long)
    Call EnumChildWindows(lngParentHandle, AddressOf AddChildEnum, "Enum Children")
End Sub

Private Function AddChildEnum(ByVal lngHWND As Long, ByVal lParam As String) As Long
    On Error Resume Next
    
    ReDim Preserve aChildren(UBound(aChildren) + 1)
    
    If Err.Number <> 0 Then
        ReDim aChildren(0)
        Err.Clear
    End If
    
    On Error GoTo 0
    
    If aChildren(0).lngCHWND = 0 And aChildren(0).strCText = "" Then
        ReDim aChildren(0 To 0)
    End If
    
    With aChildren(UBound(aChildren))
        .lngCHWND = lngHWND
        .strCText = LCase(Replace(GetText(lngHWND), "&", ""))
        
        Debug.Print "hwnd: " & .lngCHWND & "    text: " & .strCText
    End With
    
    AddChildEnum = 1
    
End Function

Private Function GetText(ByVal lngHWND As Long) As String
    Dim lngTextLen As Long
    Dim strText As String

    lngTextLen = SendMessage(lngHWND, WM_GETTEXTLENGTH, 0, 0)
    If lngTextLen = 0 Then
        GetText = ""
        Exit Function
    End If
    lngTextLen = lngTextLen + 1
    strText = Space(lngTextLen)
    lngTextLen = SendMessage(lngHWND, WM_GETTEXT, lngTextLen, ByVal strText)
    GetText = Left(strText, lngTextLen)
End Function

Public Function ClickIt(ByVal strObjectText As String) As Byte
    On Error GoTo ErrExit
    Dim i As Long
    strObjectText = LCase(strObjectText)
    For i = 0 To UBound(aChildren)
        If strObjectText = aChildren(i).strCText Then
            SendMessage aChildren(i).lngCHWND, BM_CLICK, 0, 0
            ClickIt = 1
            Exit Function
        End If
    Next
ErrExit:
End Function
