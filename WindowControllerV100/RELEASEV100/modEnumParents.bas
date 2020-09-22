Attribute VB_Name = "modEnumParents"
'  ___________________    ________________________________
' /                   \  /                                \
' | Window controller |--| By David Fiala djf1010@aol.com |
' \___________________/  \________________________________/
'
' Version 1.00   Released date: Sept. 03 2001

Option Explicit

Private Type udtParentEnum
    lngCHWND As Long
    strCText As String
End Type

Private aParent() As udtParentEnum

Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const BM_CLICK = &HF5

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As String) As Long

Public Sub DoParentEnum()
    Call EnumWindows(AddressOf AddParentEnum, "Enum Parents")
End Sub

Private Function AddParentEnum(ByVal lngHWND As Long, ByVal lParam As String) As Long
    If GetText(lngHWND) = "" Then
        AddParentEnum = 1
        Exit Function
    End If
    
    On Error Resume Next
    
    ReDim Preserve aParent(UBound(aParent) + 1)
    
    If Err.Number <> 0 Then
        ReDim aParent(0)
        Err.Clear
    End If
    
    On Error GoTo 0
    
    If aParent(0).lngCHWND = 0 And aParent(0).strCText = "" Then
        ReDim aParent(0 To 0)
    End If
    
    With aParent(UBound(aParent))
        .lngCHWND = lngHWND
        .strCText = LCase(Replace(GetText(lngHWND), "&", ""))
        
        Debug.Print "hwnd: " & .lngCHWND & "    text: " & .strCText
    End With
    
    AddParentEnum = 1
    
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

Public Sub GiveMeParents(ByRef astrParents() As String)
    Dim i As Long
    ReDim astrParents(0 To UBound(aParent))
    For i = 0 To UBound(aParent)
        astrParents(i) = aParent(i).strCText
    Next
End Sub

Public Function ParentHWNDFromText(ByVal strParentText As String) As Long
    Dim i As Long
    For i = 0 To UBound(aParent)
        If strParentText = aParent(i).strCText Then
            ParentHWNDFromText = aParent(i).lngCHWND
            Exit Function
        End If
    Next
End Function
