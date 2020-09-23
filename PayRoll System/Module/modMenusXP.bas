Attribute VB_Name = "modMenus"
Option Explicit

'                              IMPORTANT
'======================================================================
' Set the following constant to TRUE if you need to debug your code]
' When set to False, stopping your code will crash VB
'======================================================================
Public Const bAmDebugging As Boolean = False
' =====================================================================
' Go to end of module (ReadMe) for details on how to use this module
' =====================================================================

' Types used to retrieve current menu item information
Public Type MenuDataInformation    ' information to store menu data
    ItemHeight As Integer       ' submenu item height
    ItemWidth As Long           ' pixel width of caption and hotkey
    Icon As Long                ' icon index
    HotKeyPos As Integer        ' instr position for hotkey
    Status As Byte              ' 2=Separator, 4=ForceTransparency 8=ForceNoTransparency
    Caption As String           ' Caption
    OriginalCaption As String   ' used to check for updated menu captions
    Parent As Long              ' submenu ID
    ID As Long                  ' menu item ID
End Type
Public Type PanelDataInformation
    Height As Long          ' height of the menu panel
    Width As Long           ' width of the menu panel
    HKeyPos As Long         ' left edge for all hot keys
    SideBar As Long         ' width of SideBar (default is 32)
    SideBarXY As Long       ' X,Y coords of image/text within sidebar
    PanelIcon As Long       ' does 1 or more menu items have an icon
    Status As Byte          ' icon or bitmap, 0 for text
    Caption As String       ' Text, unless image is used instead
    FColor As Long          ' Sidebar text fore color
    BColor As Long          ' Sidebar back color
    SBarIcon As Long        ' icon/bitmap ID for sidebar, Font ID for text
    ID As Long
End Type
Private Type MENUITEMINFO
     cbSize As Long
     fMask As Long
     fType As Long
     fState As Long
     wID As Long
     hSubMenu As Long
     hbmpChecked As Long
     hbmpUnchecked As Long
     dwItemData As Long
     dwTypeData As Long 'String
     cch As Long
End Type
Private Type MEASUREITEMSTRUCT
     CtlType As Long
     CtlID As Long
     ItemId As Long
     ItemWidth As Long
     ItemHeight As Long
     ItemData As Long
End Type
Private Type DRAWITEMSTRUCT
     CtlType As Long
     CtlID As Long
     ItemId As Long
     itemAction As Long
     itemState As Long
     hwndItem As Long
     hDC As Long
     rcItem As RECT
     ItemData As Long
End Type
Private Type OSVERSIONINFO          ' used to help identify operating system
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type ICONINFO
    fIcon As Long
    xHotSpot As Long
    yHotSpot As Long
    hbmMask As Long
    hbmColor As Long
End Type

' APIs needed to retrieve menu information
Private Declare Function WindowFromDC Lib "user32" (ByVal hDC As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias _
     "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, _
     ByVal byPosition As Long, lpMenuItemInfo As MENUITEMINFO) As Boolean
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias _
     "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
' Subclassing APIs & stuff
Public Declare Function CallWindowProc Lib "user32" Alias _
     "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
     ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, _
     ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
     (ByVal hwnd As Long, ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long
Private Declare Function SetGraphicsMode Lib "gdi32" (ByVal hDC As Long, ByVal iMode As Long) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
' Subclassing & Windows Message Constants
Public Const GWL_WNDPROC = (-4)
Private Const WM_DRAWITEM = &H2B
Private Const WM_MEASUREITEM = &H2C
Private Const WM_INITMENU = &H116
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_ENTERIDLE = &H121
Private Const WM_MDICREATE = &H220
Private Const WM_MDIACTIVATE = &H222
Private Const WM_ENTERMENULOOP = &H211
Private Const WM_EXITMENULOOP = &H212

Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

' Menu Constants
Private Const MF_BYCOMMAND = &H0
Private Const MF_BYPOSITION = &H400
Private Const MF_OWNERDRAW = &H100
Private Const MF_SEPARATOR = &H800
Private Const MFT_SEPARATOR = MF_SEPARATOR
Private Const ODS_SELECTED = &H1
Private Const ODT_MENU = 1
Private Const MIIM_TYPE = &H10
Private Const MIIM_DATA = &H20
Private Const MIIM_SUBMENU = &H4

Private MenuData As Collection  ' Collection of clsMyMenu objects
Private ActiveHwnd As String    ' Index to focused form
Private iTabOffset As Integer   ' See DetermineOS function
Private lSubMenu As Long
Private lMDIchildClosed As Long
Private VisibleMenus As Collection

Public Sub SetMenus(Form_hWnd As Long, Optional MenuImageList As Control)
' =====================================================================
' This is the routine that will subclass form's menu & gather initial
' menu data
' =====================================================================
If bAmDebugging Then Exit Sub
' here we set the collection index & see if it's already been subclassed
Dim lMenus As Long, Looper As Integer
On Error Resume Next
If GetFormHandle(Form_hWnd) = -1 Then Exit Sub

lMenus = MenuData(CStr(Form_hWnd)).MainMenuID
If Err Then ' then new form to subclass
   ' Initialize a collection of classes if needed
   If MenuData Is Nothing Then Set MenuData = New Collection
   Dim NewMenuData As New clsMyMenu
   ' save the ImageList & Handle to the form's menu
   With NewMenuData
        .SetImageViewer MenuImageList
        .MainMenuID = GetMenu(Form_hWnd)
        ' used to redirect MDI children to parent for submenu info (see MsgProc:MDIactivate)
        .ParentForm = Form_hWnd
    End With
    ' Add the class to the class collection & remove the instance of the new class
    MenuData.Add NewMenuData, CStr(Form_hWnd)
    Set NewMenuData = Nothing
Else
    ' form is already subclassed, do nothing!
    Exit Sub
End If
Err.Clear
ActiveHwnd = CStr(Form_hWnd)    ' set collection index to current form
CleanMDIchildMenus
lMenus = GetMenuItemCount(MenuData(ActiveHwnd).MainMenuID)
For Looper = 0 To lMenus - 1
    'GetMenuMetrics GetSubMenu(MenuData(ActiveHwnd).MainMenuID, Looper)
Next
SetFreeWindow True              ' hook the window so we can intercept windows messages
End Sub

Public Sub ReleaseMenus(hwnd As Long)
' =====================================================================
' Sub prepares for Forms unloading
' This must be placed in the forms Unload event in order to
' release memory & prevent crash of program
' =====================================================================

If MenuData Is Nothing Then Exit Sub
On Error GoTo ByPassRelease
ActiveHwnd = CStr(hwnd)     ' set current index
SetFreeWindow False         ' unhook the window
On Error Resume Next
If MenuData(ActiveHwnd).ChildStatus = 1 Then
    lMDIchildClosed = MenuData(ActiveHwnd).ParentForm
End If
' remove references to that form's class & ultimately unload the class
MenuData.Remove ActiveHwnd
If MenuData.Count = 0 Then
    ' here we clean up a little when all subclassed forms have been unloaded
    Set MenuData = Nothing      ' erase the collection of classes which will unload the class
    DestroyMenuFont             ' get rid of memory font
    modDrawing.TargethDC = 0    ' get rid of refrence in that module
End If
ByPassRelease:
End Sub

Private Sub CleanMDIchildMenus()
' reset parent's menu items (see that routine for remarks)
If lMDIchildClosed = 0 Then Exit Sub
Dim Looper As Long, mMenu As Long, mII As MENUITEMINFO
mII.cbSize = Len(mII)
mII.fMask = &H1 Or &H2
mII.fType = 0
On Error Resume Next
With MenuData(CStr(lMDIchildClosed))
    For Looper = .PanelIDcount To 1 Step -1
        mMenu = .GetPanelID(Looper)
        If GetMenuItemCount(mMenu) < 0 Then .PurgeObsoleteMenus mMenu
    Next
End With
lMDIchildClosed = 0
End Sub

Public Function MsgProc(ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

' =====================================================================
' Here we determine which messages will be processed, relayed or
' skipped. Basically, we send anything thru unless we are measuring
' or drawing an item.
' =====================================================================

On Error GoTo SendMessageAsIs
' the following is a tell-tale sign of a system menu
If lParam = &H10000 Then Err.Raise 5
ActiveHwnd = CStr(hwnd) ' ensure index to current form is set
Select Case wMsg
    Case WM_ENTERMENULOOP
        'Debug.Print "entering loop"
        ' When a menu is activated, no changes can be made to the captions, enabled status, etc
        ' So we will save each submenu as it is opened and read the info only once,
        ' this will prevent unnecessary reads each time the submenu is displayed
        Set VisibleMenus = New Collection
    Case WM_MDIACTIVATE
        'Debug.Print "MDI child created"
        ' MDI children get their menus subclassed to the parent by Windows
        ' We set the class's parentform value to the MDI's parent & when
        ' submenus are processed, they are redirected to the parent
        ' The ChildStatus is set to clean out the parent's class when the
        ' child window is closed
        ' The GetSetMDIchildSysMenu command is run to store the system menu
        ' with the parent form. When the child is maximized its system menu
        ' shows up on the parent form & needs to be compared so the class
        ' doesn't draw for the system menu which it can't do!
        MenuData(ActiveHwnd).ParentForm = GetParent(GetParent(hwnd))
        MenuData(CStr(MenuData(ActiveHwnd).ParentForm)).GetSetMDIchildSysMenu GetSystemMenu(hwnd, False), True
        MenuData(ActiveHwnd).ChildStatus = 1
    Case WM_MEASUREITEM
        'Debug.Print "measuring"
        ' occurs after menu initialized & before drawing takes place
        ' send to drawing routine to measure the height/width of the menu panel
        ' If we measured it, don't let windows measure it again
        If CustomDrawMenu(wMsg, lParam, wParam) = True Then Exit Function
    Case WM_INITMENUPOPUP   ', WM_INITMENU
        If wParam = 0 Then Err.Raise 5  ' ignore these messages & pass them thru
        'Debug.Print "Popup starts"
        ' Occurs each time a menu is about to be displayed, wMsg is the handle
        ' Send flag to drawing routine to allow icons to be redrawn
        CustomDrawMenu wMsg, 0, 0
        GetMenuMetrics wParam    ' get measurements for menu items
        ' allow message to pass to the destintation
    Case WM_DRAWITEM
        'Debug.Print "drawing"
        ' sent numerous times, just about every time the mouse moves
        ' over the menu. Send flag to redraw menu if needed
        ' If we drew it, don't let windows redraw it
        If CustomDrawMenu(wMsg, lParam, wParam) = True Then Exit Function
    Case WM_EXITMENULOOP
        'Debug.Print "exiting loop"
        ' When a menu is clicked on or closed, we remove the collection of submenus
        ' so they can be redrawn again as needed
        Set VisibleMenus = Nothing
    Case WM_ENTERIDLE
        'Debug.Print "Popup ends"
        ' occurs after the entire menu has been measured & displayed
        ' at least once. Send flag to not redraw icons
        CustomDrawMenu wMsg, 0, 0
End Select
SendMessageAsIs:
MsgProc = CallWindowProc(MenuData(ActiveHwnd).OldWinProc, hwnd, wMsg, wParam, lParam)
End Function

Public Function GetMenuIconID(Menu_Caption As String) As Long
' =====================================================================
'   Returns the icon assigned in the menu caption as a long value
'   Example: {IMG:9}&Open would return 9
'   Note: Not used in any modules here, but provided for programmer use
'         if needed in their applications
' =====================================================================
Dim i As Integer
On Error GoTo NoIcon
i = InStr(Menu_Caption, "{IMG:")
If i Then
    GetMenuIconID = Val(Mid$(Menu_Caption, InStr(Menu_Caption, ":") + 1))
End If
Exit Function
NoIcon:
GetMenuIconID = 0
End Function

Private Sub GetMenuMetrics(hSubMenu As Long)
' =====================================================================
' Routine gets the meaurements of the submenus & their submenus,
'   their checked status, enabled status,
'   control keys, icon index, etc
' =====================================================================

Dim lMenus As Long, hWndRedirect As String
Dim Looper As Long, meDC As Long, lmnuID As Long, sysMenuLoc As Long
Dim mII As MENUITEMINFO, mI() As Byte
Dim tRect As RECT, lMetrics(0 To 10) As Long
Dim sCaption As String, sBarCaption As String
Dim sHotKey As String, bTabOffset As Boolean
Dim IconID As Integer, iTransparency As Integer
Dim bSetHotKeyOffset As Boolean, bNewItem As Boolean
Dim bHasIcon As Boolean, bRecalcSideBar As Long
Dim iSeparator As Integer, bSpecialSeparator As Boolean

On Error Resume Next
If MenuData(ActiveHwnd).GetSetMDIchildSysMenu(hSubMenu, False) = True Then Exit Sub
If Not VisibleMenus Is Nothing Then
    ' here we track which submenus are currently visible so we don't
    ' re-process data which isn't needed until after the submenu is closed
    lMenus = VisibleMenus(CStr(hSubMenu))
    If lMenus Then Exit Sub
End If
On Error GoTo 0
meDC = GetDC(CLng(ActiveHwnd))
hWndRedirect = MenuData(ActiveHwnd).ParentForm
' Get the ID for the next submenu item
lMenus = GetMenuItemCount(hSubMenu)
lSubMenu = hSubMenu
modDrawing.TargethDC = meDC
DetermineOS
With MenuData(hWndRedirect)         ' class for this form
    For Looper = 0 To lMenus - 1    ' loop thru each subitem
        ' get the submenu item
        bSpecialSeparator = False
        iSeparator = 0: iTransparency = 0
        sHotKey = ""
        ' now set some flags & stuff to return the caption,  checked & enabled status
        ' by referencing the dwTypeData as a byte array vs long or string,
        ' we bypass the VB crash that happens on Win98 & XP & probably others
        ReDim mI(0 To 255)
        mII.cbSize = Len(mII)
        mII.fMask = &H10 Or &H1 Or &H2
        mII.fType = 0
        mII.dwTypeData = VarPtr(mI(0))
        mII.cch = UBound(mI)
        ' get the submenu item information
        GetMenuItemInfo hSubMenu, Looper, True, mII
        'Debug.Print lmnuID; "has submenus"; mII.hSubMenu
        If Abs(mII.wID) = 4096 Or mII.wID = -1 Then Exit Sub
        lmnuID = mII.wID
        bNewItem = .SetMenuID(lmnuID, hSubMenu, False, True)
        sCaption = Left$(StrConv(mI, vbUnicode), mII.cch)
        If Len(Replace$(sCaption, Chr$(0), "")) = 0 Then sCaption = .OriginalCaption
        If Left(UCase(sCaption), 9) = "{SIDEBAR:" Then sBarCaption = sCaption
        'Debug.Print hWndRedirect; hSubMenu; lmnuID; " Caption: "; sCaption
        If .OriginalCaption = sCaption And bNewItem = False Then
            ' here we can get cached info vs reprocessing it again
            lMetrics(1) = lMetrics(1) + .ItemHeight
            lMetrics(10) = .ItemWidth
            If LoWord(lMetrics(10)) > lMetrics(0) Then lMetrics(0) = LoWord(lMetrics(10))
            If HiWord(lMetrics(10)) > lMetrics(9) Then lMetrics(9) = HiWord(lMetrics(10))
            lMetrics(4) = .SideBarWidth
            If .Icon <> 0 Then bHasIcon = True
            If InStr(sCaption, Chr$(9)) Then bTabOffset = True
            'Debug.Print "reading existing " & Looper + 1, sCaption
        Else
            bNewItem = True
            If Len(sBarCaption) > 0 And bRecalcSideBar = 0 Then bRecalcSideBar = lmnuID
            .OriginalCaption = sCaption
            .Status = 0
            ' new item or change in caption, let's get some measurements
            ' first extract the caption, controlkeys & icon
            If InStr(sCaption, Chr$(9)) Then bTabOffset = True
            ' when Win98 encounters a hotkey above, it automatically
            ' increases the menu panel width. We need to note that
            ' so we can decrease the panel widh appropriately and
            ' offset the automatic increase. This helps prevent extra
            ' wide menu panels
            If Left(UCase(sCaption), 9) = "{SIDEBAR:" Then
                iSeparator = 1
                .Status = .Status Or 16
                .ItemHeight = 0
                .ItemWidth = 0
                .Icon = 0
            Else
                'Debug.Print "Caption "; sCaption
                FindImageAndHotKey hWndRedirect, sCaption, iTransparency, sHotKey, IconID
                Debug.Print "iconid="; IconID
                ' identify whether or not this is a separator
                iSeparator = Abs(CInt(Len(sCaption) = 0 Or Left$(sCaption, 1) = "-"))
                If iSeparator = 0 Then iSeparator = Abs(CInt(mII.fType And MF_SEPARATOR) = MF_SEPARATOR)
                If iSeparator Then IconID = 0   ' no pictures on separator bars!
                If Len(sCaption) > 0 And iSeparator = 1 Then
                    ' separator bar with text
                    ' calculate entire caption & set a few flags
                    sCaption = Mid$(sCaption, 2) & "  " & sHotKey
                    bSpecialSeparator = True
                    sHotKey = ""                ' not used for separators
                End If
                ' start saving the information
                .Caption = Trim$(sCaption & " " & sHotKey)
                .Icon = IconID
                .Status = .Status Or iTransparency * 4
                .Status = .Status Or iSeparator * 2
                If IconID Then bHasIcon = True
                SetMenuFont True, , bSpecialSeparator    ' add smaller menu font
                ' measure the caption width to help identify how wide
                ' the menu panel should be (greatest width of all submenu items)
                DrawText meDC, sCaption, Len(sCaption), tRect, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE Or DT_NOCLIP
                ' keep track of the largest width, this will be used to
                ' left align control keys for the entire panel
                If tRect.Right > lMetrics(0) Then lMetrics(0) = tRect.Right
                lMetrics(10) = tRect.Right
                If iSeparator = 0 Or bSpecialSeparator = True Then
                    ' set min height text menu items to match 16x16 icon height
                    If tRect.Bottom < 10 And bSpecialSeparator = False Then tRect.Bottom = 10
                    tRect.Bottom = tRect.Bottom + 6
                Else
                    tRect.Bottom = 5    ' make default separators 0 height
                End If
                ' store the height of the caption text
                .ItemHeight = tRect.Bottom
                lMetrics(1) = lMetrics(1) + tRect.Bottom
                SetMenuFont False
                If Len(sHotKey) Then
                    .HotKeyPos = Len(sCaption) + 1
                    ' now do the same for the hotkey
                    DrawText meDC, Trim(sHotKey), Len(Trim(sHotKey)), tRect, DT_CALCRECT Or DT_LEFT Or DT_NOCLIP Or DT_SINGLELINE
                    ' keep track of the widest control key text
                    ' this is used w/widest caption to determine overall
                    ' panel width including icons & checkmarks. Add 12 pixels for
                    ' buffer between end of caption & beginning of control key
                    If tRect.Right > lMetrics(9) Then lMetrics(9) = tRect.Right
                    .ItemWidth = MakeLong(CInt(lMetrics(10)), CInt(tRect.Right))
                Else
                    .ItemWidth = MakeLong(CInt(lMetrics(10)), 0)
                End If
            End If
        End If
        ' we ensure the item is drawn by us
        ' force a separator status if appropriate
        mII.fMask = 0
        If mII.fType = MF_SEPARATOR Or iSeparator = 1 Then
           mII.fType = MF_SEPARATOR Or MF_OWNERDRAW
        Else    ' otherwise it's normal
           mII.fType = mII.fType Or MF_OWNERDRAW
        End If
        mII.fMask = mII.fMask Or MIIM_TYPE Or MIIM_DATA   ' reset mask
        ' save updates to allow us to draw the menu item
        SetMenuItemInfo hSubMenu, Looper, True, mII
    Next
    If Looper > 0 Then  ' menu items processed
        If bRecalcSideBar = 0 Then  ' sidebar menu id
            ' if no sidebar was processed, then check the overall panel height
            ' if it changed, we need to reprocess the sidebar again since
            ' the graphics & text are centered in the panel
            If .PanelHeight <> lMetrics(1) And .SideBarItem <> 0 Then bRecalcSideBar = lmnuID
        End If
        lMetrics(3) = 5 + Abs(CInt(bHasIcon)) * 18
        lMetrics(2) = lMetrics(0) + 12
        lMetrics(0) = lMetrics(2) + lMetrics(9) + lMetrics(3) + lMetrics(4) + CInt(bTabOffset) * iTabOffset
        If bRecalcSideBar Then
            .SetMenuID bRecalcSideBar, hSubMenu, False, False
            ReturnSideBarInfo hWndRedirect, sBarCaption, lMetrics(), meDC
        End If
        .UpdatePanelID lMetrics(), sBarCaption, (bRecalcSideBar = 0)
    End If
End With
If Not VisibleMenus Is Nothing Then VisibleMenus.Add 1, (CStr(hSubMenu))
' now we replace the default font & release the form's DC
SetMenuFont False, meDC
ReleaseDC CLng(ActiveHwnd), meDC
Erase lMetrics
Erase mI
End Sub

Private Sub FindImageAndHotKey(hWndRedirect As String, sKey As String, imgTransparency As Integer, sAccel As String, imgIndex As Integer)
' =====================================================================
' This routine extracts the imagelist refrence and resets it if the
' image doesn't exist or not imagelist was provided
' =====================================================================
On Error Resume Next
Dim i As Integer, sSpecial As String, sHeader As String
imgIndex = 0
imgTransparency = 0
If Left$(UCase(sKey), 5) = "{IMG:" Then
    i = InStr(sKey, "}")
    If i Then
        sHeader = UCase(Left$(sKey, i))
        sKey = Mid$(sKey, i + 1)
        ' extract the image index
        imgIndex = Val(Mid$(sHeader, 6))
        ' if the value<1 or >nr of images, then reset it to zero
        Debug.Print "icon count="; MenuData(hWndRedirect).TotalIcons
        If imgIndex < 1 Or imgIndex > MenuData(hWndRedirect).TotalIcons Then
            imgIndex = 0
        Else    ' optional transparency flag
                ' Y=always use transparency
                ' N=never user transparency
                ' default: Icons never use transparency, Bitmaps always
            If InStr(sHeader, "|Y}") Then imgTransparency = 1
            If InStr(sHeader, "|N}") Then imgTransparency = 2
        End If
    End If
End If
' Parse the Caption & the Control Key
sAccel = ""
' First let's see if it's a menu builder supplied control key
' if so, it's easy to identify 'cause it is preceeded by a vbTab
i = InStr(sKey, Chr$(9))
If i Then       ' yep, menu builder supplied control key
    sAccel = Trim$(Mid$(sKey, i + 1))
    sKey = Trim$(Left$(sKey, i - 1))
Else
    ' user supplied control key, a little more difficult to find
    For i = 1 To 3  ' look for Ctrl, Alt & Shift combinations 1st
        If InStr(UCase(sKey), Choose(i, "CTRL+", "SHIFT+", "ALT+")) Then
            ' if found, then exit routine
            sAccel = Trim$(Mid$(sKey, InStr(UCase(sKey), Choose(i, "CTRL+", "SHIFT+", "ALT+"))))
            sKey = Trim$(Left$(sKey, InStr(UCase(sKey), Choose(i, "CTRL+", "SHIFT+", "ALT+")) - 1))
            Exit Sub
        End If
    Next
    For i = 1 To 15 ' look for F keys next
        If Right$(UCase(sKey), Len("F" & i)) = "F" & i Then
            ' if found, then exit routine
            sAccel = Trim$(Mid$(sKey, InStrRev(UCase(sKey), "F" & i)))
            sKey = Trim$(Left$(sKey, InStrRev(UCase(sKey), UCase(sAccel)) - 1))
            Exit Sub
        End If
    Next
    ' here we look for other types of hot keys, these can be customized
    ' as needed by following the logic below
    For i = 1 To 6
        ' hot key looking for, it will be preceded by a space and must
        ' be at end of caption, otherwise we ignore it
        sSpecial = Choose(i, " DEL", " INS", " HOME", " END", " PGUP", " PGDN")
        If Right$(UCase(sKey), Len(sSpecial)) = sSpecial Then
            sAccel = Trim$(Mid$(sKey, InStrRev(UCase(sKey), sSpecial)))
            sKey = Trim$(Left$(sKey, InStrRev(UCase(sKey), sSpecial) - 1))
            Exit For
        End If
    Next
End If
End Sub

Private Sub ReturnSideBarInfo(hWndRedirect As String, sBarInfo As String, vBarInfo() As Long, tDC As Long)
' =======================================================================
' This routine returns the sidebar information for the current submenu
' Basically we are parsing out the SIDEBAR caption
' =======================================================================

Dim i As Integer, sImgID As String
Dim lRatio As Single, sText As String
Dim bMetrics As Boolean, sTmp As String
Dim lFont As Long, lFontM As LOGFONT, hPrevFont As Long
Dim tRect As RECT
Dim imgInfo As BITMAP, picInfo As ICONINFO
Dim TempBMP As Long, ImageDC As Long, sbarType As Integer

' here we are just adding a delimeter at end of string to make parsing easier
If Right$(sBarInfo, 1) = "}" Then sBarInfo = Left$(sBarInfo, Len(sBarInfo) - 1)
sBarInfo = sBarInfo & "|"
' stripoff the SIDEBAR header
i = InStr(UCase(sBarInfo), "{SIDEBAR:")
sBarInfo = Mid$(sBarInfo, InStr(sBarInfo, ":") + 1)
' return the type of sidebar Image or Text
i = InStr(sBarInfo, "|")
' if the next line <> TEXT then we have an image handle or image control
sImgID = Left$(sBarInfo, i - 1)

On Error Resume Next
' can't leave memory fonts running around loose -- wasted memory
If MenuData(hWndRedirect).SideBarIsText = True And MenuData(hWndRedirect).SideBarItem <> 0 Then
    ' kill the previous font for this item, if any
    DeleteObject MenuData(hWndRedirect).SideBarItem
End If
vBarInfo(10) = 0                  ' reset to force no sidebar
' use with caution. Making width too small or too large
' may prevent menu from displaying or crash on memory
' suggest using between 32 & 64
If InStr(UCase(sBarInfo), "|WIDTH:") Then      ' width of the sidebar (user-provided)
    ' undocumented! this allows the sidebar width to be modified
    vBarInfo(4) = Val(Mid$(sBarInfo, InStr(UCase(sBarInfo), "|WIDTH:") + 7))
Else
    ' however, 32 pixels wide seems to look the best
    vBarInfo(4) = 32                            ' default width of sidebars
End If
If IsNumeric(sImgID) Then         ' user is providing image handle vs a form picture object
    vBarInfo(10) = Val(sImgID)    ' ref to picture if it exists
    sbarType = 2                  ' status: image sidebar
    vBarInfo(9) = 8               ' type default as bmp
Else
    If sImgID = "TEXT" Then
        sbarType = 4              ' status: text sidebar
        vBarInfo(9) = 0
        If InStr(UCase(sBarInfo), "|CAPTION:") Then
            sText = Mid$(sBarInfo, InStr(UCase(sBarInfo), "|CAPTION:") + 9)
            i = InStr(sText, "|")
            sText = Left$(sText, i - 1)
        End If
        sBarInfo = UCase(sBarInfo)  ' make it easier to parse
        If InStr(sBarInfo, "|FONT:") Then
            ' parse out the font
            sTmp = Mid$(sBarInfo, InStr(sBarInfo, "|FONT:") + 6)
            i = InStr(sTmp, "|")
            sTmp = Left$(sTmp, i - 1)
        Else
            sTmp = "Arial"     ' default if not provided
        End If
        lFontM.lfCharSet = 0   ' scalable only
        lFontM.lfFaceName = sTmp
        ' if user wants other font attributes, then make it so
        If InStr(sBarInfo, "|BOLD") Then sTmp = sTmp & " Bold"
        If InStr(sBarInfo, "|ITALIC") Then sTmp = sTmp & " Italic"
        lFontM.lfFaceName = sTmp & Chr$(0)
        If InStr(sBarInfo, "|UNDERLINE") Then lFontM.lfUnderline = 1
        ' if user wants a different fontsize then make it so
        If InStr(sBarInfo, "|FSIZE:") Then
            i = Val(Mid$(sBarInfo, InStr(sBarInfo, "|FSIZE:") + 7))
            If i < 4 Then i = 12        ' min & max fonts
            If i > 24 Then i = 24
        Else
            i = 12  ' default font size
        End If
        Do
            ' here we are going to create fonts to see if it will
            ' fit in the sidebar, unfortunately we need to do this
            ' each time the menubar is initially displayed or resized because
            ' the sidebar height may have changed with adding/removing
            ' or making menu items invisible
            lFontM.lfHeight = (i * -20) / Screen.TwipsPerPixelY
            ' can't rotate the font before measuring it - per MSDN drawtext won't measure rotated fonts
            lFont = CreateFontIndirect(lFontM)    ' create the font without rotation
            hPrevFont = SelectObject(tDC, lFont)  ' load it into the DC
            ' see if it will fit in the sidebar
            DrawText tDC, sText, Len(sText), tRect, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE Or DT_NOCLIP Or &H800
            ' regardless we delete the font, cause we'll need to rotate it
            SelectObject tDC, hPrevFont
            DeleteObject lFont
            If tRect.Right > vBarInfo(1) Or tRect.Bottom > vBarInfo(4) Then
                ' font is too big, reduce it by 1 and try again
                i = i - 1
                If i < 4 Then Exit Do
            Else    ' font is ok, now we rotate it & save it
                lFontM.lfEscapement = 900
                lFont = CreateFontIndirect(lFontM)  ' create the font
                vBarInfo(10) = lFont                 ' save it
                vBarInfo(8) = tRect.Right           ' measurements
                vBarInfo(5) = tRect.Bottom
                Exit Do
            End If
        Loop
    Else
        ' here we have an image/picturebox control containing an image
        ' we need to extract the image handle
        Dim formID As Long, vControl As Control, bIsMDI As Boolean
        ' loop thru each open form to determine which is the active
        formID = GetFormHandle(CLng(hWndRedirect), bIsMDI)
        If formID > -1 Then
            sbarType = 2     'status: image sidebar
            ' let's see if the control passed is indexed
            If Right$(sImgID, 1) = ")" Then  ' indexed image
                i = InStrRev(sImgID, "(")
                sTmp = Left$(sImgID, i - 1)
                i = Val(Mid$(sImgID, i + 1))
                If bIsMDI Then
                    If Forms(formID).ActiveForm Is Nothing Then
                        Set vControl = Forms(formID).Controls(sTmp).Item(i)
                    Else
                        ' when control is in an MDIs active form, we reference it this way
                        Set vControl = Forms(formID).ActiveForm.Controls(sTmp).Item(i)
                    End If
                Else
                    Set vControl = Forms(formID).Controls(sTmp).Item(i)
                End If
            Else
                If bIsMDI Then
                    If Forms(formID).ActiveForm Is Nothing Then
                        Set vControl = Forms(formID).Controls(sImgID)
                    Else
                        ' when control is in an MDIs active form, we reference it this way
                        Set vControl = Forms(formID).ActiveForm.Controls(sImgID)
                    End If
                Else
                    Set vControl = Forms(formID).Controls(sImgID)
                End If
            End If
            ' cache the picture handle & type
            vBarInfo(10) = vControl.Picture.Handle
            If vControl.Picture.Type = 3 Then vBarInfo(9) = 16 Else vBarInfo(9) = 8
            Set vControl = Nothing
        End If
    End If
End If
If vBarInfo(10) = 0 Then
    'failed retrieving sidebar information
    Debug.Print "Sidebar failed"
    vBarInfo(4) = 0
    Exit Sub
End If
sBarInfo = UCase(sBarInfo)  ' make it easier to parse
'ok, let's get the rest of the attributes
If InStr(sBarInfo, "|BCOLOR:") Then
    ' Background color for the sidebar
    Select Case Left$(Mid$(sBarInfo, InStr(sBarInfo, "|BCOLOR:") + 8), 4)
    Case "NONE": vBarInfo(6) = -1
    Case "BACK":    ' short for background
        ' if a text sidebar & background was provided we change to default
        If sbarType = 2 Then vBarInfo(6) = -2 Else vBarInfo(6) = -1
    Case Else   ' numeric background color -- use it
        vBarInfo(6) = Val(Mid$(sBarInfo, InStr(sBarInfo, "|BCOLOR:") + 8))
    End Select
Else
    vBarInfo(6) = -1    ' default: use the menubar background color
End If
If vBarInfo(6) = -1 Then vBarInfo(6) = GetSysColor(COLOR_MENU)
If vBarInfo(10) Then
    If sbarType = 2 Then
        ' now if an image sidebar, we call subroutine for more attributes
        GoSub DrawPicture
        ' let's get the size of the image vs the size of the menu panel &
        ' either center or shrink the image to fit
        ' we will return the left offset, top offset & new image width, height
        If vBarInfo(5) > vBarInfo(4) Or vBarInfo(8) > vBarInfo(1) Then      ' image is larger than menu panel
            If vBarInfo(5) / vBarInfo(4) > vBarInfo(8) / vBarInfo(1) Then
                lRatio = vBarInfo(4) / vBarInfo(5)
            Else
                lRatio = vBarInfo(1) / vBarInfo(8)
            End If
            vBarInfo(5) = CInt(vBarInfo(5) * lRatio)
            vBarInfo(8) = CInt(vBarInfo(8) * lRatio)
        End If
        vBarInfo(7) = MakeLong(CInt(vBarInfo(5)), CInt(vBarInfo(8)))
        ' save the left & top offsets for the image, this way we don't have
        ' to remeasure when the menu is being displayed.
        vBarInfo(5) = MakeLong((vBarInfo(4) - vBarInfo(5)) \ 2, (vBarInfo(1) - vBarInfo(8)) \ 2)
    Else
        ' if user want's gradient background for text sidebar then
        If InStr(sBarInfo, "|GRADIENT") > 0 And sbarType = 4 Then vBarInfo(9) = vBarInfo(9) Or 32
        ' text sidebar, let's get the forecolor of the text & black is default
        If InStr(sBarInfo, "|FCOLOR:") Then
            vBarInfo(7) = Val(Mid$(sBarInfo, InStr(sBarInfo, "|FCOLOR:") + 8))
            If vBarInfo(7) < 0 Then vBarInfo(7) = 0
        Else
            vBarInfo(7) = 0
        End If
        vBarInfo(5) = MakeLong(CInt(vBarInfo(5)), CInt(vBarInfo(8)))
    End If
End If
vBarInfo(9) = sbarType Or vBarInfo(9)
vBarInfo(0) = vBarInfo(0) + vBarInfo(4)
'Debug.Print "font?"; (vBarInfo(9) And 4) = 4; vBarInfo(10)
sBarInfo = sText
Exit Sub

DrawPicture:
' this routine is used when....
' 1. When we need the background color for a mask
' 2. Image passed is a control to get height/width values
'Get the info about our image
If GetObject(vBarInfo(10), Len(imgInfo), imgInfo) = 0 Then 'And vControl Is Nothing Then
    GetIconInfo vBarInfo(10), picInfo
    If picInfo.xHotSpot = 0 Or picInfo.yHotSpot = 0 Then
        'if the image passed was a handle vs control and not a bitmap
        ' sidebar fails
        Debug.Print "Sidebar failed image is not a bitmap or icon type"
        vBarInfo(10) = 0
        vBarInfo(4) = 0
        Return
    End If
    vBarInfo(9) = 16
    vBarInfo(5) = picInfo.xHotSpot
    vBarInfo(8) = picInfo.yHotSpot
Else
    vBarInfo(9) = 8
    vBarInfo(5) = imgInfo.bmWidth
    vBarInfo(8) = imgInfo.bmHeight
End If
Err.Clear
If vBarInfo(6) = -2 Then
    Dim picIcon As PictureBox
    Forms(formID).Controls.Add "VB.PictureBox", "pic___Ic_on_s", Forms(formID)
    With Forms(formID).Controls("pic___Ic_on_s")
        .Visible = False
        .AutoRedraw = True
        If vBarInfo(6) = -2 Then
            If vBarInfo(9) = 8 Then i = 4 Else i = 3
            ' draw the image to the picturebox
            If DrawState(.hDC, 0, 0, vBarInfo(10), 0, 0, 0, 0, 0, CLng(i)) = 0 Then
                ' drawing failed, try again with differnt picture type
                If i = 4 Then i = 3 Else i = 4
                DrawState .hDC, 0, 0, vBarInfo(10), 0, 0, 0, 0, 0, CLng(i)
            End If
            ' get the mask color
            vBarInfo(6) = GetPixel(.hDC, 0, 0)
        End If
    End With
    Forms(formID).Controls.Remove "pic___Ic_on_s"
End If
Return
End Sub

Private Sub SetFreeWindow(bSet As Boolean)
' =====================================================================
' This routine hooks or unhooks a window & is used when
' menus are first set and when a form closes
' =====================================================================

If MenuData(ActiveHwnd).OldWinProc = 0 And bSet = True Then
    ' hook only if window not already hooked
    MenuData(ActiveHwnd).OldWinProc = SetWindowLong(CLng(ActiveHwnd), GWL_WNDPROC, AddressOf MsgProc)
Else
    If MenuData(ActiveHwnd).OldWinProc <> 0 And bSet = False Then
        ' hook only if window was already hooked
         SetWindowLong CLng(ActiveHwnd), GWL_WNDPROC, MenuData(ActiveHwnd).OldWinProc
         MenuData(ActiveHwnd).OldWinProc = 0
    End If
End If
End Sub

Private Function CustomDrawMenu(wMsg As Long, lParam As Long, wParam As Long) As Boolean
' =====================================================================
' Here we simply measure & draw menu items based on settings saved
' in the form's related class
' =====================================================================

Dim IsSep As Boolean, hWndRedirect As String
Static bDrawIcon As Boolean, bDrawPanel As Boolean, bGetPanelData As Boolean
Static lOffsets(0 To 2) As Long, lLastSubMenu As Long
' MDI children menus are subclassed to parent by Windows
' However, if the child isn't maximized in the MDI parent, then the menus are
' not subclassed (pain in the neck until this was figured out & re-thought)
' To work around this & prevent the submenus from being stored in both the parent
' and child classes, I redirect the actions to the parent via the GetMenuMetrics sub
' regardless whether or not the child is maximized
' Since each menu drawn is now stored the parent class, we redirect to the routine to
' get the info from the parent. If the form is the MDI parent or is a non-MDI form,
' then the ParentForm property is the same as the form's actual handle
hWndRedirect = MenuData(ActiveHwnd).ParentForm ' here we set this flag.

Select Case wMsg
Case WM_INITMENUPOPUP
    ' menu is about to be displayed, set flag to allow drawing of icons
    bDrawIcon = True: bDrawPanel = True: bGetPanelData = True
    lLastSubMenu = 0
Case WM_DRAWITEM
    Dim DrawInfo As DRAWITEMSTRUCT
    Dim IsSideBar As Boolean
    Dim hBR As Long, hOldBr As Long, hChkBr As Long
    Dim hPen As Long, hOldPen As Long, lTextColor As Long
    Dim tRect As RECT
    Dim iRectOffset As Integer, iSBoffset As Integer
    Dim sAccelKey As String, sCaption As String
    Dim bMenuItemDisabled As Boolean, bMenuItemChecked As Boolean
    Dim bSelected As Boolean, bHasIcon As Boolean
    
    'Get DRAWINFOSTRUCT which gives us sizes & indexes
    Call CopyMemory(DrawInfo, ByVal lParam, LenB(DrawInfo))
    ' only process menu items, other windows items send above message
    ' and we don't want to interfere with those. Also if we didn't
    ' process it, we don't touch it
    lSubMenu = DrawInfo.hwndItem
    If MenuData(hWndRedirect).SetMenuID(DrawInfo.ItemId, DrawInfo.hwndItem, False, False) = False Then Exit Function
    If DrawInfo.CtlType <> ODT_MENU Then Exit Function
    CustomDrawMenu = True
    IsSideBar = CBool((MenuData(hWndRedirect).Status And 16) = 16)
    If (IsSideBar = True And bDrawPanel = False) Then Exit Function
    IsSep = (MenuData(hWndRedirect).Status And 2) = 2 And IsSideBar = False
    ' get the checked & enabled status
    bMenuItemDisabled = CBool((DrawInfo.itemState And 6) = 6 Or (DrawInfo.itemState And 7) = 7)
    ' don't continue the process if the disabled item or separator
    ' was already drawn, no need to redraw it again - it doesn't change
    If bDrawIcon = False And (bMenuItemDisabled = True Or IsSep = True) Then Exit Function
    bMenuItemChecked = CBool((DrawInfo.itemState And 8) = 8 Or (DrawInfo.itemState And 9) = 9)
    ' set a reference in the drawing module to this menu's DC & set the font
    modDrawing.TargethDC = DrawInfo.hDC
    If bDrawPanel = True Or lLastSubMenu <> DrawInfo.hwndItem Then
        Dim pData(0 To 10) As Long
        MenuData(hWndRedirect).GetPanelInformation pData(), sCaption
        lOffsets(2) = pData(3)
        If lOffsets(2) Then lOffsets(2) = lOffsets(2) + 5
        lOffsets(1) = pData(4)
        If pData(4) Then lOffsets(1) = lOffsets(1) + 3
        lOffsets(0) = lOffsets(1) + lOffsets(2)
        If bDrawPanel = True Then
            If pData(10) <> 0 Then
                Debug.Print "panel xy:"; pData(4), pData(1)
                tRect.Bottom = pData(1)
                tRect.Right = pData(4)
                hBR = CreateSolidBrush(pData(6))
                hPen = GetPen(1, pData(6))
                hOldPen = SelectObject(DrawInfo.hDC, hPen)
                hOldBr = SelectObject(DrawInfo.hDC, hBR)
                DrawRect 0, 0, tRect.Right, tRect.Bottom
                SelectObject DrawInfo.hDC, hOldBr
                DeleteObject hBR
                SelectObject DrawInfo.hDC, hOldPen
                DeleteObject hPen
                pData(8) = CLng(HiWord(pData(5)))
                pData(5) = CLng(LoWord(pData(5)))
                If (pData(9) And 2) = 2 Then
                    modDrawing.TargethDC = DrawInfo.hDC
                    DrawMenuIcon pData(10), Abs(CInt((pData(9) Or 16) = 16) * 2) + 1, _
                        tRect, False, , 2, CInt(pData(5)), CInt(pData(8)), LoWord(pData(7)), HiWord(pData(7)), pData(6)
                Else
                    SetBkMode DrawInfo.hDC, NEWTRANSPARENT
                    DetermineOS DrawInfo.hDC
                    If (pData(9) And 32) = 32 Then DoGradientBkg pData(6), tRect, CLng(hWndRedirect)
                    SetMenuFont True, , , pData(10)
                    tRect.Left = (pData(4) - pData(5)) \ 2
                    tRect.Top = (pData(1) - pData(8)) \ 2 + pData(8)
                    SetTextColor DrawInfo.hDC, pData(7)
                    DrawText DrawInfo.hDC, sCaption, Len(sCaption), tRect, DT_LEFT Or DT_NOCLIP Or DT_SINGLELINE Or &H800
                    SetMenuFont False
                End If
            End If
        End If
        bDrawPanel = False
        lLastSubMenu = DrawInfo.hwndItem
        Erase pData
    End If
    If IsSideBar Then
        CustomDrawMenu = True
        Exit Function
    End If
    SetMenuFont True
    ' determine if this item is focused or not which also determines
    ' what colors we use when we are drawing
    bSelected = (DrawInfo.itemState And ODS_SELECTED) = ODS_SELECTED
    ' Now let's set some colors to draw with
    With DrawInfo
        If bSelected = True And bMenuItemDisabled = False And IsSep = False Then
             hBR = CreateSolidBrush(GetSysColor(COLOR_HIGHLIGHT))
             hPen = GetPen(1, GetSysColor(COLOR_HIGHLIGHT))
             lTextColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
        Else
             hBR = CreateSolidBrush(GetSysColor(COLOR_MENU))
             hPen = GetPen(1, GetSysColor(COLOR_MENU))
             If bMenuItemDisabled Or IsSep = True Then
                  lTextColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
             Else
                  lTextColor = GetSysColor(COLOR_MENUTEXT)
             End If
        End If
        If bMenuItemDisabled = True Then
             ' for checked & disabled items, we use default back color
             hChkBr = CreateSolidBrush(GetSysColor(COLOR_MENU))
        Else
            ' here we set the back color of a depressed button
            hChkBr = CreateSolidBrush(GetSysColor(COLOR_BTNLIGHT))
        End If
        'Select our new, correctly colored objects:
        hOldBr = SelectObject(.hDC, hBR)
        hOldPen = SelectObject(.hDC, hPen)
        'Do we have a separator bar?
        bHasIcon = False
        sCaption = MenuData(hWndRedirect).Caption
        If Not IsSep Then
        ' Ok, does this item have an icon?
        ' Here we do one more extra check in case the ImageViewer
        ' is no longer available or has no images (then handle is 0)
        ' we also set the offset for highlighting rectangle's left
        ' edge so it doesn't highlight icons
            If MenuData(hWndRedirect).ImageViewer > 0 And _
              MenuData(hWndRedirect).Icon > 0 Then
                bHasIcon = True
                iRectOffset = lOffsets(0) - 2
            Else
                'If bMenuItemChecked Then
                '    iRectOffset = lOffsets(0) - 2
                'Else
                    iRectOffset = lOffsets(1)
                'End If
            End If
            'Draw the highlighting rectangle
            DrawRect .rcItem.Left + iRectOffset, .rcItem.Top, .rcItem.Right, .rcItem.Bottom
            'Print the menu item's text
            If MenuData(hWndRedirect).HotKeyPos Then
                ' we have a control key, so identify it & its left edge
                sAccelKey = Mid$(sCaption, MenuData(hWndRedirect).HotKeyPos)
                sCaption = Left$(sCaption, InStr(sCaption, sAccelKey))
            End If
            ' send the caption, control key, icon offset, etc to be printed
            tRect = .rcItem
            DrawCaption .rcItem.Left + lOffsets(0), .rcItem.Top + 3, _
                tRect, sCaption, sAccelKey, MenuData(hWndRedirect).HotKeyEdge, lTextColor
            If bMenuItemDisabled Then   ' add the engraved affect
                tRect = .rcItem         ' get starting rectangle &
                OffsetRect tRect, -1, -1 ' offset by 1 top & left
                ' print text again with offsets
                DrawCaption .rcItem.Left + lOffsets(0) - 1, .rcItem.Top + 2, _
                    tRect, sCaption, sAccelKey, MenuData(hWndRedirect).HotKeyEdge, _
                    GetSysColor(COLOR_GRAYTEXT)
            End If
            If bMenuItemChecked Then
                ' for checked items, since they can have icons, we do a few
                ' things different. We make the checked item appear in a sunken
                ' box and make the backcolor of the box lighter than normal
                SelectObject .hDC, hChkBr
                DrawRect lOffsets(1), .rcItem.Top, lOffsets(0) - 5, .rcItem.Bottom - 1
                ThreeDbox lOffsets(1) - 2, .rcItem.Top, lOffsets(0) - 3, .rcItem.Bottom - 2, True, True
                If bHasIcon = False Then
                    ' now if the checked item doesn't have an icon we draw a checkmark in the icons' place
                    DrawCheckMark .rcItem, IIf(bMenuItemDisabled, lTextColor, GetSysColor(COLOR_MENUTEXT)), False, lOffsets(1)
                    If bMenuItemDisabled Then DrawCheckMark .rcItem, GetSysColor(COLOR_GRAYTEXT), bMenuItemDisabled, lOffsets(1)
                End If
            End If
        End If
        'If the item has an icon, selected or not, disabled or not
        If bHasIcon = True Then
            If bDrawIcon = True Or bMenuItemChecked = True Then ' we are redrawing icons
                ' extract icon handle, type & transparency option
                Dim vIconDat() As Long
                MenuData(hWndRedirect).GetIconData vIconDat(), MenuData(hWndRedirect).Icon
                'set up the location to be drawn
                tRect.Left = 4 + lOffsets(1)
                tRect.Top = ((.rcItem.Bottom - .rcItem.Top) - 16) \ 2 + .rcItem.Top
                tRect.Right = tRect.Left + 16
                tRect.Bottom = tRect.Top + 16
                'send the icon information to be drawn
                DrawMenuIcon vIconDat(0), vIconDat(1), tRect, bMenuItemDisabled, True, vIconDat(2)
            End If
            SelectObject .hDC, hBR
            If bMenuItemDisabled = False And bMenuItemChecked = False Then
                ' here we draw or remove the 3D box around the icon
                ThreeDbox lOffsets(1), .rcItem.Top, lOffsets(0) - 5, .rcItem.Bottom - 1, bSelected
             End If
        End If
        If IsSep Then
             'Finally, draw the special separator bar if needed
             ' however, if the separator has text, then we need to do
             '    some additional calculations
             If Len(sCaption) Then
                  ' separator bars with text
                  SetMenuFont True, , True    ' use smaller font
                  tRect = .rcItem             ' copy the menuitem coords
                  ' send caption to be printed in menu-select color
                  ' of course any color can be used & if you want to use the
                  ' standard 3D gray disabled color then Rem out the next line
                  ' and un-rem the next 3 lines & the second DrawCapton line
                  DrawCaption .rcItem.Left, .rcItem.Top + 3, tRect, sCaption, "", 0, GetSysColor(COLOR_HIGHLIGHT), True, CInt(lOffsets(1))
                  'DrawCaption .rcItem.Left, .rcItem.Top + 3, tRect, sCaption, "", 0, lTextColor, True
                  'tRect = .rcItem             ' recopy menuitem coords
                  'OffsetRect tRect, -1, -1    ' move coords up & left by 1
                  ' send caption again in gray
                  'DrawCaption .rcItem.Left - 1, .rcItem.Top + 2, tRect, sCaption, "", 0, GetSysColor(COLOR_GRAYTEXT), True
                  If bMenuItemChecked = False Then
                      ' here we add the lines on both sides of the separator caption
                      ThreeDbox 4 + lOffsets(1), _
                          (.rcItem.Bottom - .rcItem.Top) \ 2 + .rcItem.Top, _
                          tRect.Left - 4, _
                          (.rcItem.Bottom - .rcItem.Top) \ 2 + 1 + .rcItem.Top, True
                      ThreeDbox tRect.Right + 4, _
                          (.rcItem.Bottom - .rcItem.Top) \ 2 + .rcItem.Top, _
                          .rcItem.Right - 4, _
                          (.rcItem.Bottom - .rcItem.Top) \ 2 + 1 + .rcItem.Top, True
                  End If
             Else
              ' This will remove or add a 3D raised box for checked/non-checked items
              If bMenuItemChecked = False Then ThreeDbox lOffsets(1) + .rcItem.Left, .rcItem.Top + 2, .rcItem.Right - 4 + lOffsets(1), .rcItem.Bottom - 2, True
             End If
        End If
        'Select the old objects into the menu's DC
        Call SelectObject(.hDC, hOldBr)
        Call SelectObject(.hDC, hOldPen)
        'Delete the ones we created
        Call DeleteObject(hBR)
        Call DeleteObject(hPen)
        Call DeleteObject(hChkBr)
        SetMenuFont False
    End With
    CustomDrawMenu = True   ' set flag to prevent resending to form
Case WM_MEASUREITEM
    Dim MeasureInfo As MEASUREITEMSTRUCT
    'Get the MEASUREITEM info, basically submenu item height/width
    Call CopyMemory(MeasureInfo, ByVal lParam, Len(MeasureInfo))
    ' only process menu items, other windows items send above message
    ' and we don't want to interfere with those. Also if we didn't
    ' process it, we don't touch it
    If MenuData(hWndRedirect).SetMenuID(MeasureInfo.ItemId, lSubMenu, False, False) = False Then Exit Function
    If MeasureInfo.CtlType <> ODT_MENU Then Exit Function
    IsSep = (((MenuData(hWndRedirect).Status And 2) = 2) And (Not MenuData(hWndRedirect).Status And 16) = 16)
    'Tell Windows how big our items are.
    ' add height of each item, add a buffer of 3 pixels top/bottom for text
    MeasureInfo.ItemHeight = MenuData(hWndRedirect).ItemHeight
    MeasureInfo.ItemWidth = MenuData(hWndRedirect).PanelWidth
    'Return the information back to Windows
    Call CopyMemory(ByVal lParam, MeasureInfo, Len(MeasureInfo))
    CustomDrawMenu = True
Case WM_ENTERIDLE ' done displaying panel, let's stop drawing icons
    bDrawIcon = False
End Select
End Function

Public Function HiWord(LongIn As Long) As Integer
' =====================================================================
'   Returns the high integer of a long variable
' =====================================================================
  Call CopyMemory(HiWord, ByVal VarPtr(LongIn) + 2, 2)
End Function

Public Function LoWord(LongIn As Long) As Integer
' =====================================================================
'   Returns low integer of a long variable
' =====================================================================
  Call CopyMemory(LoWord, LongIn, 2)
End Function

Private Function MakeLong(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
' =====================================================================
'   Converts 2 integers to a long variable
' =====================================================================
  MakeLong = CLng(LoWord)
  Call CopyMemory(ByVal VarPtr(MakeLong) + 2, HiWord, 2)
End Function

Private Function DetermineOS(Optional SetGraphicsModeDC As Long = 0) As Integer
' Determine OS. Win98, for sure, seems to adjust the menu panel width
'   to accomodate for the accelerator key within the menu. If the opposite
'   adjustment isn't made, the panels wind up being wider than desired.
'   Win98: adjustment needed
'   Win2K: adjustment not needed
'   WinNT: adjustment not needed
'   WinXP: adjustment not needed
'   Other O/S: ?

' The following are the platform, major version & minor version of OS to date (acquired from MSDN)
Const os_Win95 = "1.4.0"
Const os_Win98 = "1.4.10"
Const os_WinNT4 = "2.4.0"
Const os_WinNT351 = "2.3.51"
Const os_Win2K = "2.5.0"
Const os_WinME = "1.4.90"
Const os_WinXP = "2.5.1"

  Dim verinfo As OSVERSIONINFO, sVersion As String
  verinfo.dwOSVersionInfoSize = Len(verinfo)
  If (GetVersionEx(verinfo)) = 0 Then Exit Function         ' use default 0
  With verinfo
    sVersion = .dwPlatformId & "." & .dwMajorVersion & "." & .dwMinorVersion
  End With
  ' those where the iTabOffset is set are systems that I have seen the
  ' results on; otherwise, assume no adjustment is necessary
  Select Case sVersion
  Case os_Win98: iTabOffset = 32
  Case os_Win2K: iTabOffset = 0
  Case os_WinNT4: iTabOffset = 0
  Case os_WinNT351
    ' Problems when printing rotated text
    'According to MSDN, NT 3.51 only works on a setting of 2. Don't have the opportunity to test this.
    SetGraphicsMode SetGraphicsModeDC, 2
  Case os_Win95
  Case os_WinXP: iTabOffset = 0
  Case os_WinME
  End Select
End Function

Public Function GetFormHandle(hwnd As Long, Optional bIsMDI As Boolean) As Long
Dim i As Long
For i = Forms.Count - 1 To 0 Step -1
    If Forms(i).hwnd = hwnd Then Exit For
Next
If i > -1 Then
    If TypeOf Forms(i) Is MDIForm Then bIsMDI = True
    GetFormHandle = i
End If
End Function

Public Sub ReadMe()
'==============================FOR FORM TRANSPARENCY============================
Dim Msg As Long
On Error Resume Next
If Perc < 0 Or 100 > 255 Then
  MakeTransparent = 1
Else
  Msg = GetWindowLong(hwnd, -20)
  Msg = Msg Or &H80000
  SetWindowLong hwnd, -20, Msg
  SetLayeredWindowAttributes hwnd, 0, 100, &H2
  MakeTransparent = 0
End If
If Err Then
  MakeTransparent = 2
End If
'==============================================================================
End Sub
