Attribute VB_Name = "VisualStyles"
Option Explicit
#If (VBA7 = 0) Then
Private Enum LongPtr
[_]
End Enum
#End If
#If Win64 Then
Private Const cbPtr = 8
#Else
Private Const cbPtr = 4
#End If
Public Declare PtrSafe Function ActivateVisualStyles Lib "uxtheme" Alias "SetWindowTheme" (ByVal hWnd As LongPtr, Optional ByVal pszSubAppName As LongPtr = 0, Optional ByVal pszSubIdList As LongPtr = 0) As Long
Public Declare PtrSafe Function RemoveVisualStyles Lib "uxtheme" Alias "SetWindowTheme" (ByVal hWnd As LongPtr, Optional ByRef pszSubAppName As String = " ", Optional ByRef pszSubIdList As String = " ") As Long
Public Declare PtrSafe Function GetVisualStyles Lib "uxtheme" Alias "GetWindowTheme" (ByVal hWnd As LongPtr) As Long
Private Type TINITCOMMONCONTROLSEX
dwSize As Long
dwICC As Long
End Type
Private Type TRELEASE
IUnk As IUnknown
VTable(0 To 2) As LongPtr
VTableHeaderPointer As LongPtr
End Type
Private Type TRACKMOUSEEVENTSTRUCT
cbSize As Long
dwFlags As Long
hWndTrack As LongPtr
dwHoverTime As Long
End Type
Private Enum UxThemeButtonParts
BP_PUSHBUTTON = 1
BP_RADIOBUTTON = 2
BP_CHECKBOX = 3
BP_GROUPBOX = 4
BP_USERBUTTON = 5
End Enum
Private Enum UxThemeButtonStates
PBS_NORMAL = 1
PBS_HOT = 2
PBS_PRESSED = 3
PBS_DISABLED = 4
PBS_DEFAULTED = 5
End Enum
Private Type POINTAPI
X As Long
Y As Long
End Type
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type PAINTSTRUCT
hDC As LongPtr
fErase As Long
RCPaint As RECT
fRestore As Long
fIncUpdate As Long
RGBReserved(0 To 31) As Byte
End Type
Private Type DLLVERSIONINFO
cbSize As Long
dwMajor As Long
dwMinor As Long
dwBuildNumber As Long
dwPlatformID As Long
End Type
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function InitCommonControlsEx Lib "comctl32" (ByRef ICCEX As TINITCOMMONCONTROLSEX) As Long
Private Declare PtrSafe Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Declare PtrSafe Function DllGetVersion Lib "comctl32" (ByRef pdvi As DLLVERSIONINFO) As Long
Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function ExtSelectClipRgn Lib "gdi32" (ByVal hDC As LongPtr, ByVal hRgn As LongPtr, ByVal fnMode As Long) As Long
Private Declare PtrSafe Function DrawState Lib "user32" Alias "DrawStateW" (ByVal hDC As LongPtr, ByVal hBrush As LongPtr, ByVal lpDrawStateProc As LongPtr, ByVal lData As LongPtr, ByVal wData As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal fFlags As Long) As Long
Private Declare PtrSafe Function GetProp Lib "user32" Alias "GetPropW" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Function SetProp Lib "user32" Alias "SetPropW" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr, ByVal hData As LongPtr) As Long
Private Declare PtrSafe Function RemoveProp Lib "user32" Alias "RemovePropW" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr) As LongPtr
Private Declare PtrSafe Function BeginPaint Lib "user32" (ByVal hWnd As LongPtr, lpPaint As PAINTSTRUCT) As LongPtr
Private Declare PtrSafe Function EndPaint Lib "user32" (ByVal hWnd As LongPtr, lpPaint As PAINTSTRUCT) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function InvalidateRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As Any, ByVal bErase As Long) As Long
Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hDC As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DrawFocusRect Lib "user32" (ByVal hdc As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As LongPtr, ByVal lpchText As LongPtr, ByVal cchText As Long, ByRef lprc As RECT, ByVal format As Long) As Long
Private Declare PtrSafe Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENTSTRUCT) As Long
Private Declare PtrSafe Function TransparentBlt Lib "Msimg32.dll" (ByVal hdcDest As LongPtr, ByVal xoriginDest As Long, ByVal yoriginDest As Long, ByVal wDest As Long, ByVal hDest As Long, ByVal hdcSrc As LongPtr, ByVal xoriginSrc As Long, ByVal yoriginSrc As Long, ByVal wSrc As Long, ByVal hSrc As Long, ByVal crTransparent As Long) As Long
Private Declare PtrSafe Function IsThemeBackgroundPartiallyTransparent Lib "uxtheme" (ByVal hTheme As LongPtr, ByVal iPartId As Long, ByVal iStateId As Long) As Long
Private Declare PtrSafe Function DrawThemeParentBackground Lib "uxtheme" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr, prc As RECT) As Long
Private Declare PtrSafe Function DrawThemeBackground Lib "uxtheme" (ByVal hTheme As LongPtr, ByVal hDC As LongPtr, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
Private Declare PtrSafe Function DrawThemeText Lib "uxtheme" (ByVal hTheme As LongPtr, ByVal hDC As LongPtr, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As LongPtr, ByVal cchText As Long, ByVal dwTextFlags As Long, ByVal dwTextFlags2 As Long, pRect As RECT) As Long
Private Declare PtrSafe Function GetThemeBackgroundRegion Lib "uxtheme" (ByVal hTheme As LongPtr, ByVal hDC As LongPtr, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pRegion As LongPtr) As Long
Private Declare PtrSafe Function GetThemeBackgroundContentRect Lib "uxtheme" (ByVal hTheme As LongPtr, ByVal hDC As LongPtr, ByVal iPartId As Long, ByVal iStateId As Long, pBoundingRect As RECT, pContentRect As RECT) As Long
Private Declare PtrSafe Function OpenThemeData Lib "uxtheme" (ByVal hWnd As LongPtr, ByVal pszClassList As LongPtr) As LongPtr
Private Declare PtrSafe Function CloseThemeData Lib "uxtheme" (ByVal hTheme As LongPtr) As Long
Private Declare PtrSafe Function IsAppThemed Lib "uxtheme" () As Long
Private Declare PtrSafe Function IsThemeActive Lib "uxtheme" () As Long
Private Declare PtrSafe Function GetThemeAppProperties Lib "uxtheme" () As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Private Declare PtrSafe Function SetWindowSubclass Lib "comctl32.dll" Alias "#410" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, Optional ByVal dwRefData As LongPtr) As Long
Private Declare PtrSafe Function DefSubclassProc Lib "comctl32.dll" Alias "#413" (ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32.dll" Alias "#412" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Long
Private Declare PtrSafe Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Const ICC_STANDARD_CLASSES As Long = &H4000
Private Const STAP_ALLOW_CONTROLS As Long = (1 * (2 ^ 1))
Private Const S_OK As Long = &H0
Private Const UIS_CLEAR As Long = 2
Private Const UISF_HIDEFOCUS As Long = &H1
Private Const UISF_HIDEACCEL As Long = &H2
Private Const WM_UPDATEUISTATE As Long = &H128
Private Const WM_QUERYUISTATE As Long = &H129
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_ENABLE As Long = &HA
Private Const WM_SETREDRAW As Long = &HB
Private Const WM_PAINT As Long = &HF
Private Const WM_NCPAINT As Long = &H85
Private Const WM_NCDESTROY As Long = &H82
Private Const BM_GETSTATE As Long = &HF2
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_PRINTCLIENT As Long = &H318
Private Const WM_THEMECHANGED As Long = &H31A
Private Const BST_PUSHED As Long = &H4
Private Const BST_FOCUS As Long = &H8
Private Const DT_CENTER As Long = &H1
Private Const DT_WORDBREAK As Long = &H10
Private Const DT_CALCRECT As Long = &H400
Private Const DT_HIDEPREFIX As Long = &H100000
Private Const TME_LEAVE As Long = 2
Private Const RGN_DIFF As Long = 4
Private Const RGN_COPY As Long = 5
Private Const DST_ICON As Long = &H3
Private Const DST_BITMAP As Long = &H4
Private Const DSS_DISABLED As Long = &H20

Public Sub InitVisualStylesFixes()
If App.LogMode <> 0 Then Call InitReleaseVisualStylesFixes(AddressOf ReleaseVisualStylesFixes)
Dim ICCEX As TINITCOMMONCONTROLSEX
With ICCEX
.dwSize = LenB(ICCEX)
.dwICC = ICC_STANDARD_CLASSES
End With
InitCommonControlsEx ICCEX
End Sub

Private Sub InitReleaseVisualStylesFixes(ByVal Address As LongPtr)
Static Release As TRELEASE
If Release.VTableHeaderPointer <> 0 Then Exit Sub
If GetComCtlVersion >= 6 Then
    Release.VTable(2) = Address
    Release.VTableHeaderPointer = VarPtr(Release.VTable(0))
    CopyMemory Release.IUnk, VarPtr(Release.VTableHeaderPointer), cbPtr
End If
End Sub

Private Function ReleaseVisualStylesFixes() As Long
Const SEM_NOGPFAULTERRORBOX As Long = &H2
SetErrorMode SEM_NOGPFAULTERRORBOX
End Function

Public Sub SetupVisualStylesFixes(ByVal Form As VB.Form)
If GetComCtlVersion() >= 6 Then SendMessage Form.hWnd, WM_UPDATEUISTATE, MakeDWord(UIS_CLEAR, UISF_HIDEFOCUS Or UISF_HIDEACCEL), ByVal 0&
If EnabledVisualStyles() = False Then Exit Sub
Dim CurrControl As VB.Control
For Each CurrControl In Form.Controls
    Select Case TypeName(CurrControl)
        Case "Frame"
            SetWindowSubclass CurrControl.hWnd, AddressOf RedirectFrame, ObjPtr(CurrControl), 0
        Case "CommandButton", "OptionButton", "CheckBox"
            If CurrControl.Style = vbButtonGraphical Then
                SetProp CurrControl.hWnd, StrPtr("VisualStyles"), GetVisualStyles(CurrControl.hWnd)
                If CurrControl.Enabled = True Then SetProp CurrControl.hWnd, StrPtr("Enabled"), 1
                SetWindowSubclass CurrControl.hWnd, AddressOf RedirectButton, ObjPtr(CurrControl), ObjPtr(CurrControl)
            End If
    End Select
Next CurrControl
End Sub

Public Sub RemoveVisualStylesFixes(ByVal Form As VB.Form)
If EnabledVisualStyles() = False Then Exit Sub
Dim CurrControl As VB.Control
For Each CurrControl In Form.Controls
    Select Case TypeName(CurrControl)
        Case "Frame"
            RemoveWindowSubclass CurrControl.hWnd, AddressOf RedirectFrame, ObjPtr(CurrControl)
        Case "CommandButton", "OptionButton", "CheckBox"
            If CurrControl.Style = vbButtonGraphical Then
                RemoveProp CurrControl.hWnd, StrPtr("VisualStyles")
                RemoveProp CurrControl.hWnd, StrPtr("Enabled")
                RemoveProp CurrControl.hWnd, StrPtr("Hot")
                RemoveProp CurrControl.hWnd, StrPtr("Painted")
                RemoveProp CurrControl.hWnd, StrPtr("ButtonPart")
                RemoveWindowSubclass CurrControl.hWnd, AddressOf RedirectButton, ObjPtr(CurrControl)
            End If
    End Select
Next CurrControl
End Sub

Public Function EnabledVisualStyles() As Boolean
If GetComCtlVersion() >= 6 Then
    If IsThemeActive() <> 0 Then
        If IsAppThemed() <> 0 Then
            EnabledVisualStyles = True
        ElseIf (GetThemeAppProperties() And STAP_ALLOW_CONTROLS) <> 0 Then
            EnabledVisualStyles = True
        End If
    End If
End If
End Function

Public Function GetComCtlVersion() As Long
Static Done As Boolean, Value As Long
If Done = False Then
    Dim Version As DLLVERSIONINFO
    On Error Resume Next
    Version.cbSize = LenB(Version)
    If DllGetVersion(Version) = S_OK Then Value = Version.dwMajor
    Done = True
End If
GetComCtlVersion = Value
End Function

Private Function RedirectFrame(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
Select Case wMsg
    Case WM_PRINTCLIENT, WM_MOUSELEAVE
        RedirectFrame = DefWindowProc(hWnd, wMsg, wParam, lParam)
        Exit Function
End Select
RedirectFrame = DefSubclassProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_NCDESTROY Then Call RemoveRedirectFrame(hWnd, uIdSubclass)
End Function

Private Sub RemoveRedirectFrame(ByVal hWnd As LongPtr, ByVal uIdSubclass As LongPtr)
RemoveWindowSubclass hWnd, AddressOf RedirectFrame, uIdSubclass
End Sub

Private Function RedirectButton(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal Button As Object) As LongPtr
Dim SetRedraw As Boolean
Select Case wMsg
    Case WM_NCPAINT
        Exit Function
    Case WM_PAINT
        If IsWindowVisible(hWnd) <> 0 And GetProp(hWnd, StrPtr("VisualStyles")) <> 0 Then
            Dim PS As PAINTSTRUCT
            SetProp hWnd, StrPtr("Painted"), 1
            Call DrawButton(hWnd, BeginPaint(hWnd, PS), Button)
            EndPaint hWnd, PS
            Exit Function
        End If
    Case WM_SETFOCUS, WM_ENABLE
        If IsWindowVisible(hWnd) <> 0 Then
            SetRedraw = True
            SendMessage hWnd, WM_SETREDRAW, 0, ByVal 0&
        End If
End Select
RedirectButton = DefSubclassProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_NCDESTROY Then
    Call RemoveRedirectButton(hWnd, uIdSubclass)
    RemoveProp hWnd, StrPtr("VisualStyles")
    RemoveProp hWnd, StrPtr("Enabled")
    RemoveProp hWnd, StrPtr("Hot")
    RemoveProp hWnd, StrPtr("Painted")
    RemoveProp hWnd, StrPtr("ButtonPart")
ElseIf IsWindow(hWnd) <> 0 Then
    Select Case wMsg
        Case WM_THEMECHANGED
            SetProp hWnd, StrPtr("VisualStyles"), GetVisualStyles(hWnd)
            Button.Refresh
        Case WM_MOUSELEAVE
            SetProp hWnd, StrPtr("Hot"), 0
            Button.Refresh
        Case WM_MOUSEMOVE
            If GetProp(hWnd, StrPtr("Hot")) = 0 Then
                SetProp hWnd, StrPtr("Hot"), 1
                InvalidateRect hWnd, ByVal 0&, 0
                Dim TME As TRACKMOUSEEVENTSTRUCT
                With TME
                .cbSize = LenB(TME)
                .hWndTrack = hWnd
                .dwFlags = TME_LEAVE
                End With
                TrackMouseEvent TME
            ElseIf GetProp(hWnd, StrPtr("Painted")) = 0 Then
                Button.Refresh
            End If
        Case WM_SETFOCUS, WM_ENABLE
            If SetRedraw = True Then
                SendMessage hWnd, WM_SETREDRAW, 1, ByVal 0&
                If wMsg = WM_ENABLE Then
                    SetProp hWnd, StrPtr("Enabled"), 0
                    InvalidateRect hWnd, ByVal 0&, 0
                Else
                    SetProp hWnd, StrPtr("Enabled"), 1
                    Button.Refresh
                End If
            End If
        Case WM_LBUTTONDOWN, WM_LBUTTONUP, WM_RBUTTONUP
            Button.Refresh
    End Select
End If
End Function

Private Sub RemoveRedirectButton(ByVal hWnd As LongPtr, ByVal uIdSubclass As LongPtr)
RemoveWindowSubclass hWnd, AddressOf RedirectButton, uIdSubclass
End Sub

Private Sub DrawButton(ByVal hWnd As LongPtr, ByVal hDC As LongPtr, ByVal Button As Object)
Dim Theme As LongPtr, ButtonPart As Long, ButtonState As Long, UIState As Long
Dim Enabled As Boolean, Checked As Boolean, Default As Boolean, Hot As Boolean, Pushed As Boolean, Focused As Boolean
Dim hFontOld As LongPtr, ButtonFont As IFont
Dim ButtonPicture As IPictureDisp, DisabledPictureAvailable As Boolean
Dim ClientRect As RECT, TextRect As RECT, RgnClip As LongPtr
Dim CX As Long, CY As Long, X As Long, Y As Long
ButtonPart = GetProp(hWnd, StrPtr("ButtonPart"))
If ButtonPart = 0 Then
    Select Case TypeName(Button)
        Case "CommandButton"
            ButtonPart = BP_PUSHBUTTON
        Case "OptionButton"
            ButtonPart = BP_RADIOBUTTON
        Case "CheckBox"
            ButtonPart = BP_CHECKBOX
    End Select
    If ButtonPart <> 0 Then SetProp hWnd, StrPtr("ButtonPart"), ButtonPart
End If
Select Case ButtonPart
    Case BP_PUSHBUTTON
        Default = Button.Default
        If GetFocus() <> hWnd Then
            On Error Resume Next
            If CLng(Button.Parent.ActiveControl.Default) > 0 Then Else Default = False
            On Error GoTo 0
        End If
    Case BP_RADIOBUTTON
        Checked = Button.Value
        Default = False
    Case BP_CHECKBOX
        Checked = IIf(Button.Value = vbChecked, True, False)
        Default = False
End Select
ButtonPart = BP_PUSHBUTTON
ButtonState = CLng(SendMessage(hWnd, BM_GETSTATE, 0, ByVal 0&))
UIState = CLng(SendMessage(hWnd, WM_QUERYUISTATE, 0, ByVal 0&))
Enabled = IIf(GetProp(hWnd, StrPtr("Enabled")) = 1, True, Button.Enabled)
Hot = IIf(GetProp(hWnd, StrPtr("Hot")) = 0, False, True)
If Checked = True Then Hot = False
Pushed = IIf((ButtonState And BST_PUSHED) = 0, False, True)
Focused = IIf((ButtonState And BST_FOCUS) = 0, False, True)
If Enabled = False Then
    ButtonState = PBS_DISABLED
    Set ButtonPicture = CoalescePicture(Button.DisabledPicture, Button.Picture)
    If Not Button.DisabledPicture Is Nothing Then
        If Button.DisabledPicture.Handle <> 0 Then DisabledPictureAvailable = True
    End If
ElseIf Hot = True And Pushed = False Then
    ButtonState = PBS_HOT
    If Checked = True Then
        Set ButtonPicture = CoalescePicture(Button.DownPicture, Button.Picture)
    Else
        Set ButtonPicture = Button.Picture
    End If
ElseIf Checked = True Or Pushed = True Then
    ButtonState = PBS_PRESSED
    Set ButtonPicture = CoalescePicture(Button.DownPicture, Button.Picture)
ElseIf Focused = True Or Default = True Then
    ButtonState = PBS_DEFAULTED
    Set ButtonPicture = Button.Picture
Else
    ButtonState = PBS_NORMAL
    Set ButtonPicture = Button.Picture
End If
If Not ButtonPicture Is Nothing Then
    If ButtonPicture.Handle = 0 Then Set ButtonPicture = Nothing
End If
GetClientRect hWnd, ClientRect
Theme = OpenThemeData(hWnd, StrPtr("Button"))
If Theme <> 0 Then
    GetThemeBackgroundRegion Theme, hDC, ButtonPart, ButtonState, ClientRect, RgnClip
    ExtSelectClipRgn hDC, RgnClip, RGN_DIFF
    Dim Brush As LongPtr
    Brush = CreateSolidBrush(WinColor(Button.BackColor))
    FillRect hDC, ClientRect, Brush
    DeleteObject Brush
    If IsThemeBackgroundPartiallyTransparent(Theme, ButtonPart, ButtonState) <> 0 Then DrawThemeParentBackground hWnd, hDC, ClientRect
    ExtSelectClipRgn hDC, 0, RGN_COPY
    DeleteObject RgnClip
    DrawThemeBackground Theme, hDC, ButtonPart, ButtonState, ClientRect, ClientRect
    GetThemeBackgroundContentRect Theme, hDC, ButtonPart, ButtonState, ClientRect, ClientRect
    If Focused = True Then
        If Not (UIState And UISF_HIDEFOCUS) = UISF_HIDEFOCUS Then DrawFocusRect hDC, ClientRect
    End If
    If Not Button.Caption = vbNullString Then
        Set ButtonFont = Button.Font
        hFontOld = SelectObject(hDC, ButtonFont.hFont)
        LSet TextRect = ClientRect
        DrawText hDC, StrPtr(Button.Caption), -1, TextRect, DT_CALCRECT Or DT_WORDBREAK Or CLng(IIf((UIState And UISF_HIDEACCEL) = UISF_HIDEACCEL, DT_HIDEPREFIX, 0))
        TextRect.Left = ClientRect.Left
        TextRect.Right = ClientRect.Right
        If ButtonPicture Is Nothing Then
            TextRect.Top = ((ClientRect.Bottom - TextRect.Bottom) / 2) + (3 * PixelsPerDIP_Y())
            TextRect.Bottom = TextRect.Top + TextRect.Bottom
        Else
            TextRect.Top = (ClientRect.Bottom - TextRect.Bottom) + (1 * PixelsPerDIP_Y())
            TextRect.Bottom = ClientRect.Bottom
        End If
        DrawThemeText Theme, hDC, ButtonPart, ButtonState, StrPtr(Button.Caption), -1, DT_CENTER Or DT_WORDBREAK Or CLng(IIf((UIState And UISF_HIDEACCEL) = UISF_HIDEACCEL, DT_HIDEPREFIX, 0)), 0, TextRect
        SelectObject hDC, hFontOld
        ClientRect.Bottom = TextRect.Top
        ClientRect.Left = TextRect.Left
    End If
    CloseThemeData Theme
End If
If Not ButtonPicture Is Nothing Then
    CX = CHimetricToPixel_X(ButtonPicture.Width)
    CY = CHimetricToPixel_Y(ButtonPicture.Height)
    X = ClientRect.Left + ((ClientRect.Right - ClientRect.Left - CX) / 2)
    Y = ClientRect.Top + ((ClientRect.Bottom - ClientRect.Top - CY) / 2)
    If Enabled = True Or DisabledPictureAvailable = True Then
        If ButtonPicture.Type = vbPicTypeBitmap And Button.UseMaskColor = True Then
            Dim hDCScreen As LongPtr
            Dim hDC1 As LongPtr, hBmpOld1 As LongPtr
            hDCScreen = GetDC(0)
            If hDCScreen <> 0 Then
                hDC1 = CreateCompatibleDC(hDCScreen)
                If hDC1 <> 0 Then
                    hBmpOld1 = SelectObject(hDC1, ButtonPicture.Handle)
                    TransparentBlt hDC, X, Y, CX, CY, hDC1, 0, 0, CX, CY, WinColor(Button.MaskColor)
                    SelectObject hDC1, hBmpOld1
                    DeleteDC hDC1
                End If
                ReleaseDC 0, hDCScreen
            End If
        Else
            With ButtonPicture
            #If Win64 Then
            Dim hDCl As Long
            CopyMemory hDCl, hDC, 4
            .Render hDCl Or 0&, X Or 0&, Y Or 0&, CX Or 0&, CY Or 0&, 0&, .Height, .Width, -.Height, ByVal 0&
            #Else
            .Render hDC Or 0&, X Or 0&, Y Or 0&, CX Or 0&, CY Or 0&, 0&, .Height, .Width, -.Height, ByVal 0&
            #End If
            End With
        End If
    Else
        If ButtonPicture.Type = vbPicTypeIcon Then
            DrawState hDC, 0, 0, ButtonPicture.Handle, 0, X, Y, CX, CY, DST_ICON Or DSS_DISABLED
        Else
            Dim hImage As LongPtr
            hImage = BitmapHandleFromPicture(ButtonPicture, vbWhite)
            ' The DrawState API with DSS_DISABLED will draw white as transparent.
            ' This will ensure GIF bitmaps or metafiles are better drawn.
            DrawState hDC, 0, 0, hImage, 0, X, Y, CX, CY, DST_BITMAP Or DSS_DISABLED
            DeleteObject hImage
        End If
    End If
End If
End Sub

Private Function CoalescePicture(ByVal Picture As IPictureDisp, ByVal DefaultPicture As IPictureDisp) As IPictureDisp
If Picture Is Nothing Then
    Set CoalescePicture = DefaultPicture
ElseIf Picture.Handle = 0 Then
    Set CoalescePicture = DefaultPicture
Else
    Set CoalescePicture = Picture
End If
End Function
