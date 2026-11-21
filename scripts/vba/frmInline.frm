Option Explicit

' ---------------------------------------------------------------------------
'  frmInline : chromeless host for the interactive pages.
'  It is never shown as a dialog. Attach strips its title bar and border and
'  makes it a child window inside the slide show window; Place puts it exactly
'  over the "htmlArea" shape of the current slide. It holds one Microsoft Web
'  Browser control (Shell.Explorer.2) per page, created at run time; the page
'  of the current slide sits at 0,0 and the others wait outside the visible
'  area, so changing slides is instant and every page keeps its state.
'
'  How it gets inside the slide show window (first that works):
'    1  direct child of the show window (same DPI awareness)
'    2  child of a small container window that is itself a child of the show
'       window and allows mixed DPI hosting (PowerPoint's show window is
'       per-monitor DPI aware, VBA forms are system DPI aware)
'    3  borderless window owned by the show window, kept over the same spot
' ---------------------------------------------------------------------------

Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const GWLP_HWNDPARENT As Long = -8
Private Const WS_CHILD As Long = &H40000000
Private Const WS_POPUP As Long = &H80000000
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_CLIPCHILDREN As Long = &H2000000
Private Const WS_CLIPSIBLINGS As Long = &H4000000
Private Const WS_EX_DLGMODALFRAME As Long = &H1&
Private Const WS_EX_TOOLWINDOW As Long = &H80&
Private Const WS_EX_WINDOWEDGE As Long = &H100&
Private Const WS_EX_CLIENTEDGE As Long = &H200&
Private Const WS_EX_STATICEDGE As Long = &H20000
Private Const WS_EX_APPWINDOW As Long = &H40000
Private Const WS_EX_NOACTIVATE As Long = &H8000000
Private Const SS_OWNERDRAW As Long = &HD&
Private Const SWP_NOSIZE As Long = &H1&
Private Const SWP_NOMOVE As Long = &H2&
Private Const SWP_NOZORDER As Long = &H4&
Private Const SWP_NOACTIVATE As Long = &H10&
Private Const SWP_FRAMECHANGED As Long = &H20&
Private Const SWP_SHOWWINDOW As Long = &H40&

Private mWb() As Object                  ' browser controls (late bound)
Private mBr() As clsDomBridge            ' page -> host bridges
Private mHooked() As Boolean
Private mCanvas() As Long                ' canvas width of each page in CSS px (0 = leave the zoom alone)
Private mCount As Long
Private mCur As Long
Private mHwnd As LongPtr                 ' the form's own window
Private mBox As LongPtr                  ' mixed-DPI container (mode 2)
Private mHost As LongPtr                 ' window of the slide show the form is attached to
Private mMode As Long                    ' 0 not attached, 1 child, 2 child of container, 3 owned popup
Private mShown As Boolean
Private mPaused As Boolean               ' the visible page was told that it left the screen
Private mInfo As String
Private mX As Long, mY As Long, mW As Long, mH As Long
Private mSnapWhy As String               ' what the last Snap did (troubleshooting)

Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.Caption = INLINE_TOKEN
    Me.StartUpPosition = 0
    mHwnd = FindWindowA("ThunderDFrame", INLINE_TOKEN)
    mW = -1
End Sub

Public Property Get FormHwnd() As LongPtr
    FormHwnd = mHwnd
End Property

Public Property Get Mode() As Long
    Mode = mMode
End Property

Public Property Get Shown() As Boolean
    Shown = mShown
End Property

Public Property Get Info() As String
    Info = mInfo
End Property

' True while the form sits in (or over) the given slide show window.
Public Function AttachedTo(ByVal hHost As LongPtr) As Boolean
    If mMode = 0 Or mHwnd = 0 Or hHost = 0 Or mHost <> hHost Then Exit Function
    If IsWindow(mHwnd) = 0 Then Exit Function
    Select Case mMode
        Case 1: AttachedTo = (GetParent(mHwnd) = hHost)
        Case 2: AttachedTo = (IsWindow(mBox) <> 0 And GetParent(mHwnd) = mBox)
        Case 3: AttachedTo = True
    End Select
End Function

' Creates the browser controls. False when the control is not available on this PC.
Public Function Setup(ByVal pageCount As Long) As Boolean
    Dim i As Long
    If pageCount < 1 Then Exit Function
    ReDim mWb(1 To pageCount)
    ReDim mBr(1 To pageCount)
    ReDim mHooked(1 To pageCount)
    ReDim mCanvas(1 To pageCount)
    On Error Resume Next
    mCount = pageCount
    For i = 1 To pageCount
        Err.Clear
        Set mWb(i) = Me.Controls.Add("Shell.Explorer.2", "wb" & CStr(i), True)
        If Err.Number <> 0 Then
            mInfo = mInfo & "add-err " & Err.Number & ";"
            Set mWb(i) = Nothing
            Err.Clear
            Exit For
        End If
        mWb(i).Silent = True
    Next i
    Err.Clear
    Setup = (mHwnd <> 0) And Not (mWb(pageCount) Is Nothing)
End Function

Public Sub SetCanvas(ByVal i As Long, ByVal canvasW As Long)
    If i >= 1 And i <= mCount Then mCanvas(i) = canvasW
End Sub

' Zooms page i (the browser's own zoom, like Ctrl + mouse wheel) so that its viewport is the page's canvas
' width in CSS px, whatever the size of the slide show on this screen. The zoom is a whole percentage, so the
' viewport cannot always be exact: it is never narrower than the canvas and at most about 1 % wider.
Private Sub FitCanvas(ByVal i As Long)
    Dim doc As Object, scr As Object, cw As Double, cv As Double, r As Long, r2 As Long, k As Long
    On Error Resume Next
    If i < 1 Or i > mCount Then Exit Sub
    If mCanvas(i) <= 0 Then Exit Sub
    Set doc = mWb(i).Document
    If doc Is Nothing Then Exit Sub
    cv = mCanvas(i)
    cw = doc.documentElement.clientWidth
    If Err.Number <> 0 Or cw < 8 Then Exit Sub
    If cw >= cv And cw <= cv * 1.015 Then Exit Sub
    ' The zoom command takes a percentage of 96 dpi; deviceXDPI tells which percentage is in effect now.
    Set scr = doc.parentWindow.screen
    r = CLng(Int(scr.deviceXDPI / 96# * 100# * cw / cv + 0.0001))
    If Err.Number <> 0 Or r < 1 Then Exit Sub
    For k = 1 To 5
        If r < 10 Then r = 10
        If r > 1000 Then r = 1000
        Err.Clear
        ' 63 = OLECMDID_OPTICAL_ZOOM, 2 = do not prompt; the percentage has to be passed by value
        mWb(i).Object.ExecWB 63, 2, CLng(r), Null
        If Err.Number <> 0 Then
            If InStr(mInfo, "zoom-err") = 0 Then mInfo = mInfo & "zoom-err " & Err.Number & ";"
            Exit For
        End If
        cw = doc.documentElement.clientWidth
        If Err.Number <> 0 Or cw < 8 Then Exit For
        If cw < cv Then
            r2 = r - 1                              ' a little too large: one step back
        ElseIf cw > cv * 1.015 Then
            r2 = CLng(Int(r * cw / cv + 0.0001))    ' far too small: scale up
            If r2 <= r Then Exit For
        Else
            Exit For
        End If
        If r2 = r Or r2 < 10 Or r2 > 1000 Then Exit For
        r = r2
    Next k
    Err.Clear
End Sub

' Checks the zoom of the visible page again (after the host became visible or changed size).
Public Sub Refit()
    If mCur > 0 Then FitCanvas mCur
End Sub

' Container that may host a child with another DPI awareness: created as a child of the show window,
' in that window's DPI context, while the thread allows mixed DPI hosting.
Private Function MakeBox(ByVal hHost As LongPtr) As LongPtr
    Dim oldCtx As LongPtr, oldHb As Long, gotCtx As Boolean, gotHb As Boolean
    On Error Resume Next
    oldCtx = SetThreadDpiAwarenessContext(GetWindowDpiAwarenessContext(hHost))
    gotCtx = (Err.Number = 0 And oldCtx <> 0)
    Err.Clear
    oldHb = SetThreadDpiHostingBehavior(1)
    gotHb = (Err.Number = 0)
    Err.Clear
    MakeBox = CreateWindowExA(0, "Static", "", WS_CHILD Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or SS_OWNERDRAW, 0, 0, 16, 16, hHost, 0, 0, 0)
    If gotHb Then SetThreadDpiHostingBehavior oldHb
    If gotCtx Then SetThreadDpiAwarenessContext oldCtx
    Err.Clear
End Function

' Puts the form inside the slide show window hHost (hTop = its top-level window), without title bar or border.
Public Sub Attach(ByVal hHost As LongPtr, ByVal hTop As LongPtr, Optional ByVal forceMode As Long = 0)
    Dim st As LongPtr, ex As LongPtr, e1 As Long, e2 As Long
    If mHwnd = 0 Then Exit Sub
    If mMode <> 0 Then Detach
    st = GetWindowLongPtrA(mHwnd, GWL_STYLE)
    st = st And Not (WS_POPUP Or WS_CAPTION Or WS_THICKFRAME Or WS_SYSMENU Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
    st = st Or WS_CHILD Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS
    SetWindowLongPtrA mHwnd, GWL_STYLE, st
    ex = GetWindowLongPtrA(mHwnd, GWL_EXSTYLE)
    ex = ex And Not (WS_EX_DLGMODALFRAME Or WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE Or WS_EX_APPWINDOW)
    SetWindowLongPtrA mHwnd, GWL_EXSTYLE, ex
    If forceMode = 3 Then
        e1 = -1
    ElseIf SetParent(mHwnd, hHost) <> 0 Then
        mMode = 1
    Else
        e1 = Err.LastDllError
        mBox = MakeBox(hHost)
        If mBox <> 0 Then
            If SetParent(mHwnd, mBox) <> 0 Then
                mMode = 2
            Else
                e2 = Err.LastDllError
                DestroyWindow mBox
                mBox = 0
            End If
        End If
    End If
    If mMode = 0 Then
        ' last resort: borderless top-level window that the show window owns (always directly above it)
        st = (st And Not WS_CHILD) Or WS_POPUP
        SetWindowLongPtrA mHwnd, GWL_STYLE, st
        SetWindowLongPtrA mHwnd, GWL_EXSTYLE, ex Or WS_EX_NOACTIVATE Or WS_EX_TOOLWINDOW
        SetWindowLongPtrA mHwnd, GWLP_HWNDPARENT, hTop
        mMode = 3
    End If
    mHost = hHost
    mW = -1
    mInfo = mInfo & "mode" & mMode & "(" & e1 & "," & e2 & ");"
    SetWindowPos mHwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED
End Sub

' Takes the form out of the slide show window again (hidden top-level window, as loaded).
Public Sub Detach()
    Dim st As LongPtr, was As Long
    If mHwnd = 0 Or mMode = 0 Then Exit Sub
    was = mMode
    mMode = 0
    mHost = 0
    mShown = False
    mW = -1
    If IsWindow(mHwnd) <> 0 Then
        ShowWindow mHwnd, 0
        If was = 3 Then
            SetWindowLongPtrA mHwnd, GWLP_HWNDPARENT, 0
        Else
            SetParent mHwnd, 0
            st = GetWindowLongPtrA(mHwnd, GWL_STYLE)
            st = (st And Not WS_CHILD) Or WS_POPUP
            SetWindowLongPtrA mHwnd, GWL_STYLE, st
        End If
    End If
    If mBox <> 0 Then
        If IsWindow(mBox) <> 0 Then DestroyWindow mBox
        mBox = 0
    End If
End Sub

' Position and size in pixels, relative to the client area of the slide show window.
Public Sub Place(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    Dim pt As POINT_T
    If mHwnd = 0 Or mMode = 0 Then Exit Sub
    If mMode = 3 Then
        pt.x = x
        pt.y = y
        ClientToScreen mHost, pt
        x = pt.x
        y = pt.y
    End If
    If x = mX And y = mY And w = mW And h = mH Then Exit Sub
    mX = x: mY = y: mW = w: mH = h
    If mMode = 2 Then
        SetWindowPos mBox, 0, x, y, w, h, SWP_NOZORDER Or SWP_NOACTIVATE
        SetWindowPos mHwnd, 0, 0, 0, w, h, SWP_NOZORDER Or SWP_NOACTIVATE
    Else
        SetWindowPos mHwnd, 0, x, y, w, h, SWP_NOZORDER Or SWP_NOACTIVATE
    End If
    Layout
End Sub

Private Sub Layout()
    Dim i As Long, iw As Single, ih As Single
    On Error Resume Next
    iw = Me.InsideWidth
    ih = Me.InsideHeight
    If iw < 8 Or ih < 8 Then Exit Sub
    For i = 1 To mCount
        If Not mWb(i) Is Nothing Then
            If i = mCur Then
                mWb(i).Move 0, 0, iw, ih
            Else
                mWb(i).Move -(iw + 80), 0, iw, ih
            End If
        End If
    Next i
    If mCur > 0 Then FitCanvas mCur
End Sub

Private Sub UserForm_Resize()
    Layout
End Sub

' Brings page n to the visible area.
Public Sub ShowPage(ByVal n As Long, ByVal bg As Long)
    On Error Resume Next
    If n < 1 Or n > mCount Then Exit Sub
    If n = mCur Then
        FitCanvas n
        Exit Sub
    End If
    Me.BackColor = bg
    FitCanvas n          ' while it still waits outside the visible area (same size as the visible one)
    If Not mPaused Then Notify mCur, "ihHide"     ' the page that leaves the screen may pause itself
    mCur = n
    mPaused = False
    Layout
    Notify n, "ihShow"
End Sub

' Tells a page that it came on screen / left the screen (the build tool adds the hidden #ihShow / #ihHide).
Private Sub Notify(ByVal i As Long, ByVal id As String)
    On Error Resume Next
    If i < 1 Or i > mCount Then Exit Sub
    mWb(i).Document.getElementById(id).click
End Sub

Public Sub ShowNow()
    If mHwnd = 0 Or mMode = 0 Then Exit Sub
    If mShown Then
        If IsWindowVisible(mHwnd) <> 0 Then Exit Sub
    End If
    SetWindowPos mHwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
    If mMode = 2 Then SetWindowPos mBox, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
    mShown = True
    If mPaused Then
        ' the same page comes back after the host was hidden (black screen, a slide without a page ...)
        mPaused = False
        Notify mCur, "ihShow"
    End If
End Sub

Public Sub HideNow()
    If mShown And Not mPaused Then
        Notify mCur, "ihHide"
        mPaused = True
    End If
    If mMode = 2 And mBox <> 0 Then ShowWindow mBox, 0
    If mHwnd <> 0 Then ShowWindow mHwnd, 0
    mShown = False
End Sub

' ------------------------------------------------------------------ pages

Public Sub Go(ByVal i As Long, ByVal path As String)
    On Error Resume Next
    mHooked(i) = False
    If Len(path) = 0 Then
        mWb(i).Navigate "about:blank"
    Else
        mWb(i).Navigate path
    End If
End Sub

' True once the page script has run (the added script sets data-ih-ready="1" on <body>).
Public Function PageAlive(ByVal i As Long) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = mWb(i).Document.body.getAttribute("data-ih-ready")
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    PageAlive = (CStr(v) = "1")
End Function

' Script errors the page reported (the added script collects them in data-ih-err on <html>).
Public Function PageErr(ByVal i As Long) As String
    Dim v As Variant
    On Error Resume Next
    v = mWb(i).Document.documentElement.getAttribute("data-ih-err")
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    PageErr = CStr(v)
End Function

Public Function PageComplete(ByVal i As Long) As Boolean
    Dim rs As Long, isBusy As Boolean
    On Error Resume Next
    rs = -1
    isBusy = True
    rs = mWb(i).ReadyState
    isBusy = mWb(i).Busy
    PageComplete = (rs = 4 And Not isBusy)
End Function

Public Sub WriteDirect(ByVal i As Long, ByVal html As String)
    Dim doc As Object
    On Error Resume Next
    mHooked(i) = False
    Set doc = mWb(i).Document
    doc.Open
    doc.Write html
    doc.Close
End Sub

' Attaches the bridge to the page's hidden #ihCmd element (the page "clicks" it to talk to the host).
Public Sub Hook(ByVal i As Long)
    Dim el As Object
    On Error Resume Next
    Set el = mWb(i).Document.getElementById("ihCmd")
    If el Is Nothing Then
        mInfo = mInfo & "no-ihCmd" & i & ";"
        Exit Sub
    End If
    If mBr(i) Is Nothing Then
        Set mBr(i) = New clsDomBridge
        mBr(i).Page = i
    End If
    Err.Clear
    Set el.onclick = mBr(i)
    If Err.Number <> 0 Then
        mInfo = mInfo & "hook-err" & i & " " & Err.Number & ";"
        Err.Clear
    Else
        mHooked(i) = True
        mInfo = mInfo & "ok" & i & ";"
    End If
End Sub

Public Function PageReady(ByVal i As Long) As Boolean
    If i >= 1 And i <= mCount Then PageReady = mHooked(i)
End Function

Public Function PageCmd(ByVal i As Long) As String
    Dim v As Variant
    On Error Resume Next
    v = mWb(i).Document.getElementById("ihCmd").getAttribute("data-cmd")
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    PageCmd = CStr(v)
End Function

Public Function PageArg(ByVal i As Long) As String
    Dim v As Variant
    On Error Resume Next
    v = mWb(i).Document.getElementById("ihCmd").getAttribute("data-arg")
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    PageArg = CStr(v)
End Function

Public Function Browser(ByVal i As Long) As Object
    Set Browser = mWb(i)
End Function

' The zoom of page i is brought in line with its canvas (also while the page is not the visible one).
Public Sub FitPage(ByVal i As Long)
    FitCanvas i
End Sub

' Before page i is loaded again: its old document no longer counts as running.
Public Sub MarkReload(ByVal i As Long)
    On Error Resume Next
    mHooked(i) = False
    mWb(i).Document.body.setAttribute "data-ih-ready", "0"
End Sub

' Draws page i, as it is now, into a PNG file (at the size the page has on this screen). False when nothing was drawn.
Public Property Get SnapWhy() As String
    SnapWhy = mSnapWhy
End Property

Public Function Snap(ByVal i As Long, ByVal file As String) As Boolean
    Dim rc As RECT_T, w As Long, h As Long, hScr As LongPtr, hDC As LongPtr, hBmp As LongPtr, hOld As LongPtr, hBr As LongPtr
    Dim tok As LongPtr, si As GDIP_START_T, img As LongPtr, enc As GUID_T, hr As Long, x As Long, y As Long, drawn As Boolean
    Dim pass As Long, g1 As Long, g2 As Long, g3 As Long, obj As Object
    Const SENTINEL As Long = &HFE00FF
    On Error GoTo Done
    mSnapWhy = "start"
    If i < 1 Or i > mCount Or mHwnd = 0 Then Exit Function
    If IsWindow(mHwnd) = 0 Then Exit Function
    GetClientRect mHwnd, rc
    w = rc.Right
    h = rc.Bottom
    mSnapWhy = "size " & w & "x" & h
    If w < 16 Or h < 16 Then Exit Function
    hScr = GetDC(0)
    hDC = CreateCompatibleDC(hScr)
    hBmp = CreateCompatibleBitmap(hScr, w, h)
    ReleaseDC 0, hScr
    If hDC = 0 Or hBmp = 0 Then GoTo Done
    hOld = SelectObject(hDC, hBmp)
    rc.Left = 0: rc.Top = 0: rc.Right = w: rc.Bottom = h
    ' 1 = DVASPECT_CONTENT: the object draws its page, whether or not its window is on screen.
    ' The control itself is asked first, its document second.
    For pass = 1 To 2
        hBr = CreateSolidBrush(SENTINEL)
        FillRect hDC, rc, hBr
        DeleteObject hBr
        Set obj = Nothing
        If pass = 1 Then Set obj = mWb(i).Object Else Set obj = mWb(i).Document
        hr = -1
        If Not obj Is Nothing Then hr = OleDraw(ObjPtr(obj), 1, hDC, rc)
        drawn = False
        If hr = 0 Then
            ' something must have been drawn: not every sample point may still have the fill colour
            For y = 1 To 5
                For x = 1 To 5
                    If GetPixel(hDC, (w * x) \ 6, (h * y) \ 6) <> SENTINEL Then drawn = True
                Next x
            Next y
        End If
        mSnapWhy = mSnapWhy & " pass" & pass & " hr=" & Hex$(hr) & " drawn=" & drawn
        If drawn Then Exit For
    Next pass
    SelectObject hDC, hOld
    If drawn Then
        si.GdiplusVersion = 1
        g1 = GdiplusStartup(tok, si, 0)
        If g1 = 0 Then
            g2 = GdipCreateBitmapFromHBITMAP(hBmp, 0, img)
            If g2 = 0 Then
                CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), enc
                If Len(Dir$(file)) > 0 Then Kill file
                g3 = GdipSaveImageToFile(img, StrPtr(file), enc, 0)
                Snap = (g3 = 0)
                GdipDisposeImage img
            End If
            GdiplusShutdown tok
        End If
        mSnapWhy = mSnapWhy & " gdip=" & g1 & "," & g2 & "," & g3
    End If
Done:
    If Err.Number <> 0 Then mSnapWhy = mSnapWhy & " err=" & Err.Number & " " & Err.Description
    On Error Resume Next
    If hBmp <> 0 Then DeleteObject hBmp
    If hDC <> 0 Then DeleteDC hDC
End Function

Public Function Diag() As String
    Dim r As String, rc As RECT_T, i As Long
    On Error Resume Next
    r = "hwnd=" & CStr(mHwnd) & " mode=" & mMode & " box=" & CStr(mBox) & " shown=" & mShown & " cur=" & mCur
    If mHwnd <> 0 Then
        r = r & " alive=" & (IsWindow(mHwnd) <> 0) & " visible=" & (IsWindowVisible(mHwnd) <> 0) & " parent=" & CStr(GetParent(mHwnd))
        GetWindowRect mHwnd, rc
        r = r & " rect=" & rc.Left & "," & rc.Top & " " & (rc.Right - rc.Left) & "x" & (rc.Bottom - rc.Top)
        GetClientRect mHwnd, rc
        r = r & " client=" & rc.Right & "x" & rc.Bottom
        r = r & " style=" & Hex$(CLng(GetWindowLongPtrA(mHwnd, GWL_STYLE) And &H7FFFFFFF))
    End If
    r = r & " place=" & mX & "," & mY & " " & mW & "x" & mH & " inside=" & Me.InsideWidth & "x" & Me.InsideHeight
    For i = 1 To mCount
        r = r & " p" & i & "[" & IIf(PageAlive(i), "alive", "-") & IIf(mHooked(i), ",hooked", "") & " " & mWb(i).Left & "," & mWb(i).Top & " " & mWb(i).Width & "x" & mWb(i).Height
        r = r & " view=" & mWb(i).Document.documentElement.clientWidth & "x" & mWb(i).Document.documentElement.clientHeight & " mode=" & mWb(i).Document.documentMode
        If Len(PageErr(i)) > 0 Then r = r & " ERR{" & PageErr(i) & "}"
        r = r & "]"
    Next i
    Diag = r & " info=" & mInfo
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' never closed by the user (there is no title bar); only Unload from code
    If CloseMode = 0 Then Cancel = True
End Sub

Private Sub UserForm_Terminate()
    Dim i As Long
    On Error Resume Next
    For i = 1 To mCount
        Set mBr(i) = Nothing
        Set mWb(i) = Nothing
    Next i
    If mBox <> 0 Then
        If IsWindow(mBox) <> 0 Then DestroyWindow mBox
        mBox = 0
    End If
End Sub
