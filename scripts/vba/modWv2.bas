Option Explicit

' ---------------------------------------------------------------------------
'  Modern page engine for the slide show: Microsoft Edge WebView2.
'
'  WebView2 is the Chromium engine of Microsoft Edge as a component. Its runtime is part of Windows 11 and of
'  current Windows 10 / Office installations; Office itself brings the small loader (WebView2Loader.dll) next
'  to POWERPNT.EXE. This module talks to it directly: COM calls through the method tables (DispCallFunc) and
'  small hand-made callback objects, so nothing has to be installed or registered and no reference is needed.
'
'  One browser view per page, each a child window of a container window that the host puts over the slide.
'  A page talks to the host with window.chrome.webview.postMessage("command" + "\n" + argument); the host
'  answers with PostWebMessageAsString("show" | "hide").
'
'  Method numbers ("slots") are positions in the COM method tables of WebView2 (IUnknown takes 0 to 2). They
'  were read from Microsoft's own interop assembly and never change for a published interface.
' ---------------------------------------------------------------------------

#If Win64 Then
Private Const PTR As Long = 8
#Else
Private Const PTR As Long = 4
#End If

Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" (ByVal browserFolder As LongPtr, ByVal userDataFolder As LongPtr, ByVal options As LongPtr, ByVal handler As LongPtr) As Long
Private Declare PtrSafe Function LoadLibraryW Lib "kernel32" (ByVal lpFileName As LongPtr) As LongPtr
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByRef prgvt As Integer, ByRef prgpvarg As LongPtr, ByRef pvargResult As Variant) As Long
Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32" (ByVal cb As LongPtr) As LongPtr
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal pv As LongPtr)
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As LongPtr, ByVal src As LongPtr, ByVal cb As LongPtr)
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal psz As LongPtr) As Long
Private Declare PtrSafe Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As LongPtr, ByVal fDeleteOnRelease As Long, ByRef ppstm As LongPtr) As Long
Private Declare PtrSafe Function GetHGlobalFromStream Lib "ole32" (ByVal pstm As LongPtr, ByRef phglobal As LongPtr) As Long
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function SetEnvironmentVariableW Lib "kernel32" (ByVal lpName As LongPtr, ByVal lpValue As LongPtr) As Long
Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef iid As GUID_T) As Long
Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long

' slots: environment
Private Const S_ENV_CREATECONTROLLER As Long = 3
' slots: controller
Private Const S_CTL_PUT_VISIBLE As Long = 4
Private Const S_CTL_PUT_BOUNDS As Long = 6
Private Const S_CTL_PUT_ZOOM As Long = 8
Private Const S_CTL_MOVEFOCUS As Long = 12
Private Const S_CTL_NOTIFYMOVED As Long = 23
Private Const S_CTL_CLOSE As Long = 24
Private Const S_CTL_GET_WEBVIEW As Long = 25
Private Const S_CTL2_PUT_BACKCOLOR As Long = 27
' slots: view
Private Const S_WV_GET_SETTINGS As Long = 3
Private Const S_WV_NAVIGATE As Long = 5
Private Const S_WV_EXECUTESCRIPT As Long = 29
Private Const S_WV_CAPTUREPREVIEW As Long = 30
Private Const S_WV_POSTSTRING As Long = 33
Private Const S_WV_ADD_MESSAGE As Long = 34
Private Const S_WV_CALLCDP As Long = 36
Private Const S_WV3_MAPHOST As Long = 71
' slots: settings (put_...)
Private Const S_SET_STATUSBAR As Long = 10
Private Const S_SET_DEVTOOLS As Long = 12
Private Const S_SET_CONTEXTMENUS As Long = 14
Private Const S_SET_ZOOMCONTROL As Long = 18
' slots: message arguments, stream
Private Const S_ARGS_TRYGETSTRING As Long = 5
Private Const S_STM_SEEK As Long = 5

Private mLoaderTried As Boolean, mLoaderOk As Boolean
Private mEnv As LongPtr                 ' ICoreWebView2Environment
Private mEnvDone As Boolean, mEnvHr As Long, mEnvFailed As Boolean
Private mWhy As String                  ' why the engine is not available (troubleshooting)
Private mVtEnv As LongPtr, mVtCtl As LongPtr, mVtMsg As LongPtr, mVtText As LongPtr, mVtSnap As LongPtr
Private mObjEnv As LongPtr, mObjText As LongPtr, mObjSnap As LongPtr
Private mN As Long                      ' pages
Private mCtl() As LongPtr               ' ICoreWebView2Controller per page
Private mWv() As LongPtr                ' ICoreWebView2 per page
Private mObjCtl() As LongPtr, mObjMsg() As LongPtr
Private mCtlDone() As Boolean, mCtlHr() As Long
Private mTextDone As Boolean, mTextHr As Long, mText As String
Private mSnapDone As Boolean, mSnapHr As Long
' Messages of the pages wait here and are handed to the host from a timer. WebView2 calls its event handlers
' from inside its own message handling and does not allow a message loop to run in there; the host, however,
' may have to wait for a new view while it answers a message. So the handler only stores the message.
Private mQPage(0 To 255) As Long, mQCmd(0 To 255) As String, mQArg(0 To 255) As String
Private mQHead As Long, mQTail As Long  ' the queue holds the entries from mQHead up to, not including, mQTail
Private mTimer As LongPtr, mDraining As Boolean

' ------------------------------------------------------------------ small tools

Private Function FnPtr(ByVal p As LongPtr) As LongPtr
    FnPtr = p
End Function

' Calls method number "slot" of the COM object at pObj. Arguments must be typed: Long (CLng, 0&), LongPtr
' (pointers, StrPtr, VarPtr) or Double. Returns the HRESULT.
Private Function VCall(ByVal pObj As LongPtr, ByVal slot As Long, ParamArray a() As Variant) As Long
    Dim vt(0 To 9) As Integer, pv(0 To 9) As LongPtr, v(0 To 9) As Variant, r As Variant
    Dim i As Long, n As Long, hr As Long
    If pObj = 0 Then
        VCall = &H80004003
        Exit Function
    End If
    n = UBound(a) - LBound(a) + 1
    For i = 0 To n - 1
        v(i) = a(LBound(a) + i)
        vt(i) = VarType(v(i))
        pv(i) = VarPtr(v(i))
    Next i
    hr = DispCallFunc(pObj, CLngPtr(slot) * PTR, 4, vbLong, n, vt(0), pv(0), r)
    If hr <> 0 Then
        VCall = hr
    Else
        VCall = CLng(r)
    End If
End Function

Private Function PtrToStr(ByVal p As LongPtr) As String
    Dim n As Long
    If p = 0 Then Exit Function
    n = lstrlenW(p)
    If n <= 0 Then Exit Function
    PtrToStr = String$(n, vbNullChar)
    CopyMemory StrPtr(PtrToStr), p, CLngPtr(n) * 2
End Function

Private Function ElapsedSince(ByVal t0 As Single) As Single
    Dim el As Single
    el = Timer - t0
    If el < 0 Then el = el + 86400!
    ElapsedSince = el
End Function

' ------------------------------------------------------------------ callback objects
' A callback object is a block of two pointers: [method table][tag]. The method table holds IUnknown's three
' methods and Invoke. The objects live as long as the process (a few bytes each).

Private Function MakeVtbl(ByVal pInvoke As LongPtr) As LongPtr
    Dim p As LongPtr, f(0 To 3) As LongPtr
    p = CoTaskMemAlloc(4 * PTR)
    f(0) = FnPtr(AddressOf WvQI)
    f(1) = FnPtr(AddressOf WvAddRef)
    f(2) = FnPtr(AddressOf WvRelease)
    f(3) = pInvoke
    CopyMemory p, VarPtr(f(0)), 4 * PTR
    MakeVtbl = p
End Function

Private Function MakeObj(ByVal vtbl As LongPtr, ByVal tag As Long) As LongPtr
    Dim p As LongPtr, d(0 To 1) As LongPtr
    p = CoTaskMemAlloc(2 * PTR)
    d(0) = vtbl
    d(1) = tag
    CopyMemory p, VarPtr(d(0)), 2 * PTR
    MakeObj = p
End Function

Private Function TagOf(ByVal this As LongPtr) As Long
    Dim t As LongPtr
    CopyMemory VarPtr(t), this + PTR, PTR
    TagOf = CLng(t)
End Function

' IUnknown of the callback objects. Only IUnknown and the five handler interfaces are answered.
Private Function WvQI(ByVal this As LongPtr, ByVal riid As LongPtr, ByVal ppv As LongPtr) As Long
    Dim d(0 To 3) As Long, ok As Boolean, z As LongPtr
    On Error Resume Next
    If ppv = 0 Or riid = 0 Then
        WvQI = &H80004003
        Exit Function
    End If
    CopyMemory VarPtr(d(0)), riid, 16
    If d(0) = 0 And d(1) = 0 And d(2) = &HC0 And d(3) = &H46000000 Then ok = True
    Select Case d(0)
        Case &H4E8A3389, &H6C4819F3, &H57213F19, &H49511172, &H697E05E9, &H5C4889F0: ok = True
    End Select
    If ok Then
        CopyMemory ppv, VarPtr(this), PTR
        WvQI = 0
    Else
        CopyMemory ppv, VarPtr(z), PTR
        WvQI = &H80004002
    End If
End Function

Private Function WvAddRef(ByVal this As LongPtr) As Long
    WvAddRef = 2
End Function

Private Function WvRelease(ByVal this As LongPtr) As Long
    WvRelease = 1
End Function

Private Function EnvDone(ByVal this As LongPtr, ByVal hr As Long, ByVal pEnv As LongPtr) As Long
    On Error Resume Next
    mEnvHr = hr
    If hr = 0 And pEnv <> 0 Then
        mEnv = pEnv
        VCall mEnv, 1
    End If
    mEnvDone = True
End Function

Private Function CtlDone(ByVal this As LongPtr, ByVal hr As Long, ByVal pCtl As LongPtr) As Long
    Dim i As Long, pw As LongPtr
    On Error Resume Next
    i = TagOf(this)
    If i < 1 Or i > mN Then Exit Function
    mCtlHr(i) = hr
    If hr = 0 And pCtl <> 0 Then
        mCtl(i) = pCtl
        VCall pCtl, 1
        VCall pCtl, S_CTL_GET_WEBVIEW, VarPtr(pw)
        mWv(i) = pw
    End If
    mCtlDone(i) = True
End Function

Private Function MsgIn(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    Dim i As Long, p As LongPtr, s As String, hr As Long, k As Long
    On Error Resume Next
    i = TagOf(this)
    hr = VCall(args, S_ARGS_TRYGETSTRING, VarPtr(p))
    If hr <> 0 Or p = 0 Then Exit Function
    s = PtrToStr(p)
    CoTaskMemFree p
    k = InStr(1, s, vbLf, vbBinaryCompare)
    If k > 0 Then
        Enqueue i, Left$(s, k - 1), Mid$(s, k + 1)
    Else
        Enqueue i, s, ""
    End If
End Function

Private Sub Enqueue(ByVal page As Long, ByVal c As String, ByVal a As String)
    Dim nxt As Long, last As Long
    On Error Resume Next
    If c = "tick" And mQHead <> mQTail Then
        ' one waiting "tick" is enough
        last = (mQTail + 255) And 255
        If mQCmd(last) = "tick" Then Exit Sub
    End If
    nxt = (mQTail + 1) And 255
    If nxt = mQHead Then mQHead = (mQHead + 1) And 255
    mQPage(mQTail) = page
    mQCmd(mQTail) = c
    mQArg(mQTail) = a
    mQTail = nxt
    If mTimer = 0 And Not mDraining Then mTimer = SetTimer(0, 0, 1, FnPtr(AddressOf WvPumpProc))
End Sub

' Timer callback: hands the waiting messages to the host (modInline.InlineWvMsg), outside any WebView2 handler.
Private Sub WvPumpProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)
    Dim page As Long, c As String, a As String
    On Error Resume Next
    KillTimer 0, idEvent
    If idEvent = mTimer Then mTimer = 0
    If mDraining Then Exit Sub
    mDraining = True
    Do While mQHead <> mQTail
        page = mQPage(mQHead)
        c = mQCmd(mQHead)
        a = mQArg(mQHead)
        mQArg(mQHead) = ""
        mQHead = (mQHead + 1) And 255
        InlineWvMsg page, c, a
    Loop
    mDraining = False
End Sub

Private Function TextDone(ByVal this As LongPtr, ByVal hr As Long, ByVal pText As LongPtr) As Long
    On Error Resume Next
    mTextHr = hr
    mText = ""
    If pText <> 0 Then mText = PtrToStr(pText)
    mTextDone = True
End Function

Private Function SnapDone(ByVal this As LongPtr, ByVal hr As Long) As Long
    On Error Resume Next
    mSnapHr = hr
    mSnapDone = True
End Function

' ------------------------------------------------------------------ engine

Public Property Get WvWhy() As String
    WvWhy = mWhy
End Property

Private Function LoadLoader() As Boolean
    Dim c(0 To 3) As String, i As Long, pf As String
    If mLoaderTried Then
        LoadLoader = mLoaderOk
        Exit Function
    End If
    mLoaderTried = True
    On Error Resume Next
    ' next to POWERPNT.EXE (Office brings it), then the usual Office folders
    c(0) = Application.Path & "\WebView2Loader.dll"
    pf = Environ$("ProgramFiles")
    c(1) = pf & "\Microsoft Office\root\Office16\WebView2Loader.dll"
    c(2) = Environ$("ProgramFiles(x86)") & "\Microsoft Office\root\Office16\WebView2Loader.dll"
    c(3) = "WebView2Loader.dll"
    For i = 0 To 3
        If Len(c(i)) > 0 Then
            If LoadLibraryW(StrPtr(c(i))) <> 0 Then
                mLoaderOk = True
                Exit For
            End If
        End If
    Next i
    LoadLoader = mLoaderOk
End Function

Private Sub MakeFolders(ByVal path As String)
    Dim parts() As String, i As Long, p As String
    On Error Resume Next
    parts = Split(path, "\")
    p = parts(0)
    For i = 1 To UBound(parts)
        p = p & "\" & parts(i)
        If Len(Dir$(p, vbDirectory)) = 0 Then MkDir p
    Next i
End Sub

' Starts the engine (once). False when WebView2 cannot be used on this PC: WvWhy says why.
Public Function WvStart() As Boolean
    Dim udf As String, hr As Long, t0 As Single, extra As String
    If mEnv <> 0 Then
        WvStart = True
        Exit Function
    End If
    If mEnvFailed Then Exit Function
    On Error GoTo Fail
    If Not LoadLoader() Then
        mWhy = "the WebView2 loader was not found"
        mEnvFailed = True
        Exit Function
    End If
    If mVtEnv = 0 Then
        mVtEnv = MakeVtbl(FnPtr(AddressOf EnvDone))
        mVtCtl = MakeVtbl(FnPtr(AddressOf CtlDone))
        mVtMsg = MakeVtbl(FnPtr(AddressOf MsgIn))
        mVtText = MakeVtbl(FnPtr(AddressOf TextDone))
        mVtSnap = MakeVtbl(FnPtr(AddressOf SnapDone))
        mObjEnv = MakeObj(mVtEnv, 0)
        mObjText = MakeObj(mVtText, 0)
        mObjSnap = MakeObj(mVtSnap, 0)
    End If
    ' A browser profile of its own, never shared with the views Office itself opens (views that share a profile
    ' must be started the same way, or creating a view fails with 8007139F). PPTHTML_WV2_DATA names another
    ' folder: the build tool's hidden PowerPoint uses it, so that it never shares a profile with a visible one.
    udf = Environ$("PPTHTML_WV2_DATA")
    If Len(udf) = 0 Then
        udf = Environ$("LOCALAPPDATA")
        If Len(udf) = 0 Then udf = Environ$("TEMP")
        udf = udf & "\ppt-html\wv2"
    End If
    MakeFolders udf
    ' sound and video may start without a click, as in a presentation one expects
    extra = "--autoplay-policy=no-user-gesture-required"
    SetEnvironmentVariableW StrPtr("WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"), StrPtr(extra)
    mEnvDone = False
    mEnvHr = 0
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(udf), 0, mObjEnv)
    SetEnvironmentVariableW StrPtr("WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"), 0
    If hr <> 0 Then
        mWhy = "the engine did not start (" & Hex$(hr) & ")"
        mEnvFailed = True
        Exit Function
    End If
    t0 = Timer
    Do While Not mEnvDone And ElapsedSince(t0) < 20
        DoEvents
        Sleep 2
    Loop
    If mEnv = 0 Then
        mWhy = "the engine did not answer (" & Hex$(mEnvHr) & ", done=" & mEnvDone & ")"
        mEnvFailed = True
        Exit Function
    End If
    WvStart = True
    Exit Function
Fail:
    mWhy = "the engine could not be called (" & Err.Number & " " & Err.Description & ")"
    mEnvFailed = True
End Function

' Room for n pages.
Public Sub WvSetup(ByVal n As Long)
    Dim i As Long
    mN = n
    ReDim mCtl(1 To n)
    ReDim mWv(1 To n)
    ReDim mObjCtl(1 To n)
    ReDim mObjMsg(1 To n)
    ReDim mCtlDone(1 To n)
    ReDim mCtlHr(1 To n)
    For i = 1 To n
        mObjCtl(i) = MakeObj(mVtCtl, i)
        mObjMsg(i) = MakeObj(mVtMsg, i)
    Next i
End Sub

Public Function WvHas(ByVal i As Long) As Boolean
    If i >= 1 And i <= mN Then WvHas = (mWv(i) <> 0)
End Function

' Creates the browser view of page i as a child of hParent (invisible until WvVisible). bg = RGB of the page.
Public Function WvCreate(ByVal i As Long, ByVal hParent As LongPtr, ByVal bg As Long) As Boolean
    Dim hr As Long, t0 As Single, pSettings As LongPtr, tok(0 To 1) As Long, iid As GUID_T, p2 As LongPtr, argb As Long
    On Error GoTo Fail
    If i < 1 Or i > mN Or mEnv = 0 Then Exit Function
    If mWv(i) <> 0 Then
        WvCreate = True
        Exit Function
    End If
    mCtlDone(i) = False
    mCtlHr(i) = 0
    hr = VCall(mEnv, S_ENV_CREATECONTROLLER, hParent, mObjCtl(i))
    If hr <> 0 Then
        mWhy = "no view for page " & i & " (" & Hex$(hr) & ")"
        Exit Function
    End If
    t0 = Timer
    Do While Not mCtlDone(i) And ElapsedSince(t0) < 20
        DoEvents
        Sleep 2
    Loop
    If mWv(i) = 0 Then
        mWhy = "no view for page " & i & " (" & Hex$(mCtlHr(i)) & ", done=" & mCtlDone(i) & ")"
        Exit Function
    End If
    VCall mCtl(i), S_CTL_PUT_VISIBLE, 0&
    If VCall(mWv(i), S_WV_GET_SETTINGS, VarPtr(pSettings)) = 0 And pSettings <> 0 Then
        VCall pSettings, S_SET_STATUSBAR, 0&
        VCall pSettings, S_SET_CONTEXTMENUS, 0&
        VCall pSettings, S_SET_ZOOMCONTROL, 0&
        VCall pSettings, S_SET_DEVTOOLS, 0&
        VCall pSettings, 2
    End If
    ' the colour behind the page until it has drawn itself: the page's own background (alpha, red, green, blue)
    IIDFromString StrPtr("{C979903E-D4CA-4228-92EB-47EE3FA96EAB}"), iid
    If VCall(mCtl(i), 0, VarPtr(iid), VarPtr(p2)) = 0 And p2 <> 0 Then
        argb = &HFF Or ((bg And &HFF) * &H100&) Or (((bg \ &H100&) And &HFF) * &H10000) Or BlueHigh(bg)
        VCall p2, S_CTL2_PUT_BACKCOLOR, argb
        VCall p2, 2
    End If
    VCall mWv(i), S_WV_ADD_MESSAGE, mObjMsg(i), VarPtr(tok(0))
    WvCreate = True
    Exit Function
Fail:
    mWhy = "view " & i & " failed (" & Err.Number & " " & Err.Description & ")"
End Function

' The blue byte of an RGB value moved to the highest byte of a Long (without overflow).
Private Function BlueHigh(ByVal bg As Long) As Long
    Dim b As Long
    b = (bg \ &H10000) And &HFF
    If b >= &H80 Then
        BlueHigh = ((b - &H80) * &H1000000) Or &H80000000
    Else
        BlueHigh = b * &H1000000
    End If
End Function

' The folder whose files the pages may load ("https://<host>/file"). Call before WvGo. False = use file addresses.
Public Function WvMapFolder(ByVal i As Long, ByVal host As String, ByVal folder As String) As Boolean
    Dim iid As GUID_T, p3 As LongPtr, hr As Long
    On Error Resume Next
    If i < 1 Or i > mN Then Exit Function
    If mWv(i) = 0 Then Exit Function
    IIDFromString StrPtr("{A0D6DF20-3B92-416D-AA0C-437A9C727857}"), iid
    If VCall(mWv(i), 0, VarPtr(iid), VarPtr(p3)) <> 0 Or p3 = 0 Then Exit Function
    hr = VCall(p3, S_WV3_MAPHOST, StrPtr(host), StrPtr(folder), 1&)
    VCall p3, 2
    WvMapFolder = (hr = 0)
End Function

Public Sub WvGo(ByVal i As Long, ByVal url As String)
    On Error Resume Next
    If i < 1 Or i > mN Then Exit Sub
    VCall mWv(i), S_WV_NAVIGATE, StrPtr(url)
End Sub

' Size of the view inside its parent, in pixels.
Public Sub WvSize(ByVal i As Long, ByVal w As Long, ByVal h As Long)
    Dim rc As RECT_T
    On Error Resume Next
    If i < 1 Or i > mN Then Exit Sub
    If mCtl(i) = 0 Then Exit Sub
    rc.Left = 0: rc.Top = 0: rc.Right = w: rc.Bottom = h
#If Win64 Then
    VCall mCtl(i), S_CTL_PUT_BOUNDS, VarPtr(rc)
#Else
    VCall mCtl(i), S_CTL_PUT_BOUNDS, 0&, 0&, CLng(w), CLng(h)
#End If
End Sub

Public Sub WvMoved(ByVal i As Long)
    On Error Resume Next
    If i < 1 Or i > mN Then Exit Sub
    If mCtl(i) <> 0 Then VCall mCtl(i), S_CTL_NOTIFYMOVED
End Sub

Public Sub WvZoom(ByVal i As Long, ByVal factor As Double)
    On Error Resume Next
    If i < 1 Or i > mN Then Exit Sub
    If mCtl(i) <> 0 And factor > 0.05 And factor < 20 Then VCall mCtl(i), S_CTL_PUT_ZOOM, CDbl(factor)
End Sub

Public Sub WvVisible(ByVal i As Long, ByVal yes As Boolean)
    On Error Resume Next
    If i < 1 Or i > mN Then Exit Sub
    If mCtl(i) <> 0 Then VCall mCtl(i), S_CTL_PUT_VISIBLE, CLng(IIf(yes, 1, 0))
End Sub

Public Sub WvFocus(ByVal i As Long)
    On Error Resume Next
    If i < 1 Or i > mN Then Exit Sub
    If mCtl(i) <> 0 Then VCall mCtl(i), S_CTL_MOVEFOCUS, 0&
End Sub

Public Sub WvPost(ByVal i As Long, ByVal msg As String)
    On Error Resume Next
    If i < 1 Or i > mN Then Exit Sub
    If mWv(i) <> 0 Then VCall mWv(i), S_WV_POSTSTRING, StrPtr(msg)
End Sub

' Runs JavaScript in page i and returns the value as JSON text ("" on failure). Waits up to secs.
Public Function WvEval(ByVal i As Long, ByVal js As String, Optional ByVal secs As Single = 8) As String
    Dim t0 As Single, hr As Long
    On Error GoTo Fail
    If i < 1 Or i > mN Then Exit Function
    If mWv(i) = 0 Then Exit Function
    mTextDone = False
    mText = ""
    hr = VCall(mWv(i), S_WV_EXECUTESCRIPT, StrPtr(js), mObjText)
    If hr <> 0 Then Exit Function
    t0 = Timer
    Do While Not mTextDone And ElapsedSince(t0) < secs
        DoEvents
        Sleep 2
    Loop
    If mTextDone And mTextHr = 0 Then WvEval = mText
    Exit Function
Fail:
End Function

' A command of the browser's developer protocol (the show test uses it for real mouse and key input).
Public Function WvCdp(ByVal i As Long, ByVal method As String, ByVal json As String, Optional ByVal secs As Single = 8) As String
    Dim t0 As Single, hr As Long
    On Error GoTo Fail
    If i < 1 Or i > mN Then Exit Function
    If mWv(i) = 0 Then Exit Function
    mTextDone = False
    mText = ""
    hr = VCall(mWv(i), S_WV_CALLCDP, StrPtr(method), StrPtr(json), mObjText)
    If hr <> 0 Then
        WvCdp = "CDP-ERR " & Hex$(hr)
        Exit Function
    End If
    t0 = Timer
    Do While Not mTextDone And ElapsedSince(t0) < secs
        DoEvents
        Sleep 2
    Loop
    If mTextDone And mTextHr = 0 Then
        WvCdp = mText
    Else
        WvCdp = "CDP-ERR " & Hex$(mTextHr)
    End If
    Exit Function
Fail:
End Function

' Draws page i, as it is now, into a PNG file. False when that did not work.
Public Function WvSnap(ByVal i As Long, ByVal file As String) As Boolean
    Dim pStm As LongPtr, hG As LongPtr, p As LongPtr, n As Long, b() As Byte, f As Integer
    Dim t0 As Single, hr As Long, pos As Currency
    On Error GoTo Fail
    If i < 1 Or i > mN Then Exit Function
    If mWv(i) = 0 Then Exit Function
    If CreateStreamOnHGlobal(0, 1, pStm) <> 0 Or pStm = 0 Then Exit Function
    mSnapDone = False
    mSnapHr = -1
    hr = VCall(mWv(i), S_WV_CAPTUREPREVIEW, 0&, pStm, mObjSnap)
    If hr = 0 Then
        t0 = Timer
        Do While Not mSnapDone And ElapsedSince(t0) < 12
            DoEvents
            Sleep 2
        Loop
    End If
    If hr = 0 And mSnapDone And mSnapHr = 0 Then
        ' the end of the stream is how many bytes were written (2 = seek from the end)
#If Win64 Then
        VCall pStm, S_STM_SEEK, CLngLng(0), 2&, VarPtr(pos)
#Else
        VCall pStm, S_STM_SEEK, 0&, 0&, 2&, VarPtr(pos)
#End If
        n = CLng(pos * 10000@)
        If n > 100 And GetHGlobalFromStream(pStm, hG) = 0 Then
            If CLngPtr(n) <= GlobalSize(hG) Then
                p = GlobalLock(hG)
                If p <> 0 Then
                    ReDim b(0 To n - 1)
                    CopyMemory VarPtr(b(0)), p, CLngPtr(n)
                    GlobalUnlock hG
                    If Len(Dir$(file)) > 0 Then Kill file
                    f = FreeFile
                    Open file For Binary Access Write As #f
                    Put #f, , b
                    Close #f
                    WvSnap = True
                End If
            End If
        End If
    End If
    VCall pStm, 2
    Exit Function
Fail:
    On Error Resume Next
    Close #f
    If pStm <> 0 Then VCall pStm, 2
End Function

' Closes the view of page i.
Public Sub WvDrop(ByVal i As Long)
    On Error Resume Next
    If i < 1 Or i > mN Then Exit Sub
    If mCtl(i) <> 0 Then
        VCall mCtl(i), S_CTL_CLOSE
        VCall mCtl(i), 2
    End If
    If mWv(i) <> 0 Then VCall mWv(i), 2
    mCtl(i) = 0
    mWv(i) = 0
End Sub

' Closes every view and lets the engine go (it is started again by the next show).
Public Sub WvStop()
    Dim i As Long
    On Error Resume Next
    For i = 1 To mN
        WvDrop i
    Next i
    If mEnv <> 0 Then VCall mEnv, 2
    mEnv = 0
    mN = 0
    If mTimer <> 0 Then KillTimer 0, mTimer
    mTimer = 0
    mQHead = 0
    mQTail = 0
End Sub
