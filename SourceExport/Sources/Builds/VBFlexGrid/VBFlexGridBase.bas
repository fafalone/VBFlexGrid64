Attribute VB_Name = "VBFlexGridBase"
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

#Const ImplementPreTranslateMsg = (VBFLXGRD_OCX <> 0)

Private Type TINITCOMMONCONTROLSEX
dwSize As Long
dwICC As Long
End Type
Private Type WNDCLASSEX
cbSize As Long
dwStyle As Long
lpfnWndProc As LongPtr
cbClsExtra As Long
cbWndExtra As Long
hInstance As LongPtr
hIcon As LongPtr
hCursor As LongPtr
hbrBackground As LongPtr
lpszMenuName As LongPtr
lpszClassName As LongPtr
hIconSm As LongPtr
End Type
Private Type BITMAP
BMType As Long
BMWidth As Long
BMHeight As Long
BMWidthBytes As Long
BMPlanes As Integer
BMBitsPixel As Integer
BMBits As LongPtr
End Type

#If Win64 Then
Public Declare PtrSafe Function FlexPtrToShadowObj Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (ByRef Destination As Any, ByVal lpObject As LongPtr) As LongPtr
Public Declare PtrSafe Function FlexShadowObjPtrAddRef Lib "msvbvm60.dll" Alias "__vbaObjAddref" (ByVal lpObject As LongPtr) As LongPtr
Public Declare PtrSafe Function FlexShadowObjPtr Lib "msvbvm60.dll" Alias "__vbaObjSet" (ByRef Destination As Any, ByVal lpObject As LongPtr) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function InitCommonControlsEx Lib "comctl32" (ByRef ICCEX As TINITCOMMONCONTROLSEX) As Long
Private Declare PtrSafe Function GetClassInfoEx Lib "user32" Alias "GetClassInfoExW" (ByVal hInstance As LongPtr, ByVal lpClassName As LongPtr, lpWndClass As WNDCLASSEX) As Long
Private Declare PtrSafe Function RegisterClassEx Lib "user32" Alias "RegisterClassExW" (ByRef lpWndClassEx As WNDCLASSEX) As Integer
Private Declare PtrSafe Function UnregisterClass Lib "user32" Alias "UnregisterClassW" (ByVal lpClassName As LongPtr, ByVal hInstance As LongPtr) As Long
Private Declare PtrSafe Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As LongPtr, ByVal lpCursorName As Any) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As Any) As LongPtr
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As LongPtr) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function CreateBitmapIndirect Lib "gdi32" (ByRef lpBitmap As BITMAP) As LongPtr
Private Declare PtrSafe Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function GetProp Lib "user32" Alias "GetPropW" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Function SetProp Lib "user32" Alias "SetPropW" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr, ByVal hData As LongPtr) As Long
Private Declare PtrSafe Function RemoveProp Lib "user32" Alias "RemovePropW" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowSubclass Lib "comctl32.dll" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, Optional ByVal dwRefData As LongPtr) As Long
Private Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32.dll" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Long
Private Declare PtrSafe Function DefSubclassProc Lib "comctl32.dll" (ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowSubclassW2K Lib "comctl32.dll" Alias "#410" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, Optional ByVal dwRefData As LongPtr) As Long
Private Declare PtrSafe Function RemoveWindowSubclassW2K Lib "comctl32.dll" Alias "#412" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Long
Private Declare PtrSafe Function DefSubclassProcW2K Lib "comctl32.dll" Alias "#413" (ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Declare Function FlexPtrToShadowObj Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (ByRef Destination As Any, ByVal lpObject As Long) As Long
Public Declare Function FlexShadowObjPtrAddRef Lib "msvbvm60.dll" Alias "__vbaObjAddref" (ByVal lpObject As Long) As Long
Public Declare Function FlexShadowObjPtr Lib "msvbvm60.dll" Alias "__vbaObjSet" (ByRef Destination As Any, ByVal lpObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function InitCommonControlsEx Lib "comctl32" (ByRef ICCEX As TINITCOMMONCONTROLSEX) As Long
Private Declare Function GetClassInfoEx Lib "user32" Alias "GetClassInfoExW" (ByVal hInstance As Long, ByVal lpClassName As Long, ByRef lpWndClassEx As WNDCLASSEX) As Long
Private Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExW" (ByRef lpWndClassEx As WNDCLASSEX) As Integer
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassW" (ByVal lpClassName As Long, ByVal hInstance As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As Any) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function CreateBitmapIndirect Lib "gdi32" (ByRef lpBitmap As BITMAP) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropW" (ByVal hWnd As Long, ByVal lpString As Long, ByVal hData As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowSubclassW2K Lib "comctl32" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclassW2K Lib "comctl32" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProcW2K Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
	
#End If

Private Const WM_CREATE As Long = &H1
Private Const WM_DESTROY As Long = &H2
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_UAHDESTROYWINDOW As Long = &H90
Private Const CS_DBLCLKS As Long = &H8
Private Const CS_DROPSHADOW As Long = &H20000
Private Const IDC_ARROW As Long = 32512
Private ShellModHandle As LongPtr, ShellModCount As Long
Private FlexSubclassProcPtr As LongPtr
Private FlexSubclassW2K As Integer
Private FlexClassAtom As Integer, FlexRefCount As Long
Private FlexComboCalendarClassAtom As Integer, FlexComboCalendarRefCount As Long
Private FlexSplitterBrush As LongPtr

#If ImplementPreTranslateMsg = True Then

Private Type POINTAPI
X As Long
Y As Long
End Type
Private Type TMSG
hWnd As LongPtr
Message As Long
wParam As LongPtr
lParam As LongPtr
Time As Long
PT As POINTAPI
End Type
#If Win64 Then
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByRef lParam As Any) As LongPtr
Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExW" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hhk As LongPtr) As Long
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal ncode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExW" (ByVal IDHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadID As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Private Const WM_USER As Long = &H400
Private Const UM_PRETRANSLATEMSG As Long = (WM_USER + 333)
Private FlexPreTranslateMsgHookHandle As LongPTr
Private FlexPreTranslateMsgHwnd() As LongPtr, FlexPreTranslateMsgCount As Long

#End If

Public Sub FlexLoadShellMod()
If (ShellModHandle Or ShellModCount) = 0 Then ShellModHandle = LoadLibrary(StrPtr("Shell32.dll"))
ShellModCount = ShellModCount + 1
End Sub

Public Sub FlexReleaseShellMod()
ShellModCount = ShellModCount - 1
If ShellModCount = 0 And ShellModHandle <> 0 Then
    FreeLibrary ShellModHandle
    ShellModHandle = 0
End If
End Sub

Public Sub FlexInitCC(ByVal ICC As Long)
Dim ICCEX As TINITCOMMONCONTROLSEX
With ICCEX
.dwSize = LenB(ICCEX)
.dwICC = ICC
End With
InitCommonControlsEx ICCEX
End Sub

Public Sub FlexSetSubclass(ByVal hWnd As LongPtr, ByVal This As VBFlexGrid, ByVal dwRefData As LongPtr, Optional ByVal Name As String)
If hWnd = 0 Then Exit Sub
If Name = vbNullString Then Name = "Flex"
If GetProp(hWnd, StrPtr(Name & "SubclassInit")) = 0 Then
    If FlexSubclassProcPtr = 0 Then FlexSubclassProcPtr = ProcPtr(AddressOf FlexSubclassProc)
    If FlexSubclassW2K = 0 Then
        Dim hLib As LongPtr
        hLib = LoadLibrary(StrPtr("comctl32.dll"))
        If hLib <> 0 Then
            If GetProcAddress(hLib, "SetWindowSubclass") <> 0 Then
                FlexSubclassW2K = 1
            ElseIf GetProcAddress(hLib, 410&) <> 0 Then
                FlexSubclassW2K = -1
            End If
            FreeLibrary hLib
        End If
    End If
    If FlexSubclassW2K > -1 Then
        SetWindowSubclass hWnd, FlexSubclassProcPtr, ObjPtr(This), dwRefData
    Else
        SetWindowSubclassW2K hWnd, FlexSubclassProcPtr, ObjPtr(This), dwRefData
    End If
    SetProp hWnd, StrPtr(Name & "SubclassID"), ObjPtr(This)
    SetProp hWnd, StrPtr(Name & "SubclassInit"), 1
End If
End Sub

Public Function FlexDefaultProc(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
If FlexSubclassW2K > -1 Then
    FlexDefaultProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
Else
    FlexDefaultProc = DefSubclassProcW2K(hWnd, wMsg, wParam, lParam)
End If
End Function

Public Sub FlexRemoveSubclass(ByVal hWnd As LongPtr, Optional ByVal Name As String)
If hWnd = 0 Then Exit Sub
If Name = vbNullString Then Name = "Flex"
If GetProp(hWnd, StrPtr(Name & "SubclassInit")) = 1 Then
    If FlexSubclassW2K > -1 Then
        RemoveWindowSubclass hWnd, FlexSubclassProcPtr, GetProp(hWnd, StrPtr(Name & "SubclassID"))
    Else
        RemoveWindowSubclassW2K hWnd, FlexSubclassProcPtr, GetProp(hWnd, StrPtr(Name & "SubclassID"))
    End If
    RemoveProp hWnd, StrPtr(Name & "SubclassID")
    RemoveProp hWnd, StrPtr(Name & "SubclassInit")
End If
End Sub

Public Function FlexSubclassProc(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
Select Case wMsg
    Case WM_DESTROY
        FlexSubclassProc = FlexDefaultProc(hWnd, wMsg, wParam, lParam)
        Exit Function
    Case WM_NCDESTROY, WM_UAHDESTROYWINDOW
        FlexSubclassProc = FlexDefaultProc(hWnd, wMsg, wParam, lParam)
        If FlexSubclassW2K > -1 Then
            RemoveWindowSubclass hWnd, FlexSubclassProcPtr, uIdSubclass
        Else
            RemoveWindowSubclassW2K hWnd, FlexSubclassProcPtr, uIdSubclass
        End If
        Exit Function
End Select
On Error Resume Next
Dim This As VBFlexGrid
FlexPtrToShadowObj This, uIdSubclass
If Err.Number = 0 Then
    FlexSubclassProc = This.FSubclass_Message(hWnd, wMsg, wParam, lParam, dwRefData)
Else
    FlexSubclassProc = FlexDefaultProc(hWnd, wMsg, wParam, lParam)
End If
End Function

Public Sub FlexWndRegisterClass()
If (FlexClassAtom Or FlexRefCount) = 0 Then
    Dim WCEX As WNDCLASSEX, ClassName As String
    ClassName = "VBFlexGridWndClass"
    With WCEX
    .cbSize = LenB(WCEX)
    ' CS_VREDRAW and CS_HREDRAW will not be specified as entire redraw upon resize is not necessary.
    .dwStyle = CS_DBLCLKS
    .lpfnWndProc = ProcPtr(AddressOf FlexWindowProc)
    .cbWndExtra = cbPtr
    .hInstance = App.hInstance
    .hCursor = LoadCursor(0, IDC_ARROW)
    .hbrBackground = 0
    .lpszClassName = StrPtr(ClassName)
    End With
    FlexClassAtom = RegisterClassEx(WCEX)
    If FlexSplitterBrush = 0 Then
        Dim Bmp As BITMAP, Pattern(0 To 3) As Long, hBmp As LongPtr
        Pattern(0) = &HAAAA5555
        Pattern(1) = &HAAAA5555
        Pattern(2) = &HAAAA5555
        Pattern(3) = &HAAAA5555
        With Bmp
        .BMType = 0
        .BMWidth = 16
        .BMHeight = 8
        .BMWidthBytes = 2
        .BMPlanes = 1
        .BMBitsPixel = 1
        .BMBits = VarPtr(Pattern(0))
        End With
        hBmp = CreateBitmapIndirect(Bmp)
        If hBmp <> 0 Then
            FlexSplitterBrush = CreatePatternBrush(hBmp)
            DeleteObject hBmp
        End If
    End If
End If
FlexRefCount = FlexRefCount + 1
End Sub

Public Sub FlexWndReleaseClass()
FlexRefCount = FlexRefCount - 1
If FlexRefCount = 0 And FlexClassAtom <> 0 Then
    UnregisterClass MakeDWord(FlexClassAtom, 0), App.hInstance
    FlexClassAtom = 0
    If FlexSplitterBrush <> 0 Then
        DeleteObject FlexSplitterBrush
        FlexSplitterBrush = 0
    End If
End If
End Sub

Public Sub FlexComboCalendarRegisterClass()
If (FlexComboCalendarClassAtom Or FlexComboCalendarRefCount) = 0 Then
    Dim WCEX As WNDCLASSEX, ClassName As String
    GetClassInfoEx App.hInstance, StrPtr("SysMonthCal32"), WCEX
    ClassName = "VBFlexGridComboCalendarClass"
    With WCEX
    .cbSize = LenB(WCEX)
    If Not (.dwStyle And CS_DROPSHADOW) = CS_DROPSHADOW Then .dwStyle = .dwStyle Or CS_DROPSHADOW
    .hInstance = App.hInstance
    .lpszClassName = StrPtr(ClassName)
    End With
    FlexComboCalendarClassAtom = RegisterClassEx(WCEX)
End If
FlexComboCalendarRefCount = FlexComboCalendarRefCount + 1
End Sub

Public Sub FlexComboCalendarReleaseClass()
FlexComboCalendarRefCount = FlexComboCalendarRefCount - 1
If FlexComboCalendarRefCount = 0 And FlexComboCalendarClassAtom <> 0 Then
    UnregisterClass MakeDWord(FlexComboCalendarClassAtom, 0), App.hInstance
    FlexComboCalendarClassAtom = 0
End If
End Sub

Public Function FlexGetSplitterBrush() As LongPtr
FlexGetSplitterBrush = FlexSplitterBrush
End Function

Public Function FlexWindowProc(ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Dim lCustData As LongPtr
Select Case wMsg
    Case WM_CREATE
        CopyMemory ByVal VarPtr(lCustData), ByVal lParam, cbPtr
        SetWindowLong hWnd, 0, lCustData
        FlexWindowProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
        Exit Function
    Case WM_DESTROY, WM_NCDESTROY, WM_UAHDESTROYWINDOW
        FlexWindowProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
        Exit Function
End Select
lCustData = GetWindowLong(hWnd, 0)
If lCustData <> 0 Then
    On Error Resume Next
    Dim This As VBFlexGrid
    FlexPtrToShadowObj This, lCustData
    If Err.Number = 0 Then
        FlexWindowProc = This.FSubclass_Message(hWnd, wMsg, wParam, lParam, 1)
    Else
        FlexWindowProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
    End If
Else
    FlexWindowProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
End If
End Function

#If ImplementPreTranslateMsg = True Then

Public Sub FlexPreTranslateMsgAddHook(ByVal hWnd As LongPtr)
If (FlexPreTranslateMsgHookHandle Or FlexPreTranslateMsgCount) = 0 Then
    Const WH_GETMESSAGE As Long = 3
    FlexPreTranslateMsgHookHandle = SetWindowsHookEx(WH_GETMESSAGE, AddressOf FlexPreTranslateMsgHookProc, 0, App.ThreadID)
    ReDim FlexPreTranslateMsgHwnd(0) As LongPtr
    FlexPreTranslateMsgHwnd(0) = hWnd
Else
    ReDim Preserve FlexPreTranslateMsgHwnd(0 To FlexPreTranslateMsgCount) As LongPtr
    FlexPreTranslateMsgHwnd(FlexPreTranslateMsgCount) = hWnd
End If
FlexPreTranslateMsgCount = FlexPreTranslateMsgCount + 1
End Sub

Public Sub FlexPreTranslateMsgReleaseHook(ByVal hWnd As LongPtr)
FlexPreTranslateMsgCount = FlexPreTranslateMsgCount - 1
If FlexPreTranslateMsgCount = 0 And FlexPreTranslateMsgHookHandle <> 0 Then
    UnhookWindowsHookEx FlexPreTranslateMsgHookHandle
    FlexPreTranslateMsgHookHandle = 0
    Erase FlexPreTranslateMsgHwnd()
Else
    If FlexPreTranslateMsgCount > 0 Then
        Dim i As Long
        For i = 0 To FlexPreTranslateMsgCount
            If FlexPreTranslateMsgHwnd(i) = hWnd And i < FlexPreTranslateMsgCount Then
                FlexPreTranslateMsgHwnd(i) = FlexPreTranslateMsgHwnd(i + 1)
            End If
        Next i
        ReDim Preserve FlexPreTranslateMsgHwnd(0 To FlexPreTranslateMsgCount - 1) As Long
    End If
End If
End Sub

Private Function FlexPreTranslateMsgHookProc(ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Const HC_ACTION As Long = 0, PM_REMOVE As Long = &H1
Const WM_KEYFIRST As Long = &H100, WM_KEYLAST As Long = &H108, WM_NULL As Long = &H0
If nCode >= HC_ACTION And wParam = PM_REMOVE Then
    Dim Msg As TMSG
    CopyMemory Msg, ByVal lParam, LenB(Msg)
    If Msg.Message >= WM_KEYFIRST And Msg.Message <= WM_KEYLAST Then
        If FlexPreTranslateMsgCount > 0 Then
            Dim i As Long
            For i = 0 To FlexPreTranslateMsgCount - 1
                If Msg.hWnd = FlexPreTranslateMsgHwnd(i) Then
                    If SendMessage(Msg.hWnd, UM_PRETRANSLATEMSG, 0, ByVal lParam) <> 0 Then
                        Msg.Message = WM_NULL
                        Msg.wParam = 0
                        Msg.lParam = 0
                        CopyMemory ByVal lParam, Msg, LenB(Msg)
                        Exit For
                    End If
                End If
            Next i
        End If
    End If
End If
FlexPreTranslateMsgHookProc = CallNextHookEx(FlexPreTranslateMsgHookHandle, nCode, wParam, lParam)
End Function

#End If
