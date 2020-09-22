Attribute VB_Name = "modEnumMenus"
'  ___________________    ________________________________
' /                   \  /                                \
' | Window controller |--| By David Fiala djf1010@aol.com |
' \___________________/  \________________________________/
'
' Version 1.00   Released date: Sept. 03 2001
'
'*******MENU SECTION NOT COMPLETED AND IS NOT BEING USED.

Option Explicit

Private Type MENUITEMINFO 'I didn't make this type. I think microsoft did.
    cbSize As Long        'Its listed in API Viewer this way.
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private Type udtMenuEnum 'This is mine
    lngMenuHWND As Long
    strMenuText As String
End Type

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private aMenu() As udtMenuEnum

Private Const WM_COMMAND = &H111

Public Sub EnumMenus(lngFormHWND As Long)
    ReDim aMenu(0 To GetMenuItemCount(GetMenu(lngFormHWND)))
    Call DoEnum(GetMenu(lngFormHWND))
End Sub

Private Sub DoEnum(lngMenuHWND As Long)
    ReDim Preserve aMenu(0 To (UBound(aMenu) + GetMenuItemCount(lngMenuHWND)))
    Dim lngTempPos As Long
    Dim lngTempID As Long
    Dim miiTemp As MENUITEMINFO
    
    For lngTempPos = 0 To UBound(aMenu) Step 1
        lngTempID = GetMenuItemID(lngMenuHWND, lngTempPos)
    Next
    
End Sub
