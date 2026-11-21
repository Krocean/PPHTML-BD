Option Explicit

' ---------------------------------------------------------------------------
'  Inline HTML for the slide show.
'
'  Each slide that carries the tag HtmlPage=N and a shape named "htmlArea"
'  gets the interactive page N drawn exactly over that shape while the slide
'  show runs. The page lives in a chromeless host (frmInline, no title bar,
'  no border) that is made a child window of the slide show window, so it is
'  part of the slide surface: it moves, clips and closes with the show.
'
'  Self-contained: the pages travel inside the presentation, in a custom XML
'  part (namespace DATA_NS), and are written to %TEMP% at run time. This code
'  is the same for every deck. No server, no add-in, no registry entry.
'
'  A page talks to this host through two hidden elements that the build tool
'  adds to it: #ihCmd (the page "clicks" it; data-cmd holds the command) and
'  #ihShow (the host clicks it when the page's slide comes on screen).
'
'  Started automatically by any of:
'    - the ribbon onLoad callback (InlineHtml_OnLoad) -> application events
'    - PowerPoint's OnSlideShowPageChange auto macro
'    - the action on the htmlArea picture (InlineHtml_Activate)
' ---------------------------------------------------------------------------

Public Type RECT_T
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINT_T
    x As Long
    y As Long
End Type

Public Type GDIP_START_T
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Public Type GUID_T
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Public Declare PtrSafe Function FindWindowExA Lib "user32" (ByVal hParent As LongPtr, ByVal hAfter As LongPtr, ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Public Declare PtrSafe Function SetParent Lib "user32" (ByVal hChild As LongPtr, ByVal hNewParent As LongPtr) As LongPtr
Public Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Public Declare PtrSafe Function IsWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Public Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As LongPtr) As Long
Public Declare PtrSafe Function IsChild Lib "user32" (ByVal hParent As LongPtr, ByVal hWnd As LongPtr) As Long
Public Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, ByRef rc As RECT_T) As Long
Public Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, ByRef rc As RECT_T) As Long
Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Public Declare PtrSafe Function SetFocusApi Lib "user32" Alias "SetFocus" (ByVal hWnd As LongPtr) As LongPtr
Public Declare PtrSafe Function GetFocusApi Lib "user32" Alias "GetFocus" () As LongPtr
Public Declare PtrSafe Function PostMessageA Lib "user32" (ByVal hWnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Public Declare PtrSafe Function MapVirtualKeyA Lib "user32" (ByVal uCode As Long, ByVal uMapType As Long) As Long
Public Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As LongPtr, ByRef procId As Long) As Long
Public Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare PtrSafe Function GetWindowTextW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr, ByVal cch As Long) As Long
Public Declare PtrSafe Function GetThreadDpiAwarenessContext Lib "user32" () As LongPtr
Public Declare PtrSafe Function SetThreadDpiAwarenessContext Lib "user32" (ByVal ctx As LongPtr) As LongPtr
Public Declare PtrSafe Function GetWindowDpiAwarenessContext Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Public Declare PtrSafe Function GetAwarenessFromDpiAwarenessContext Lib "user32" (ByVal ctx As LongPtr) As Long
Public Declare PtrSafe Function AreDpiAwarenessContextsEqual Lib "user32" (ByVal ctxA As LongPtr, ByVal ctxB As LongPtr) As Long
Public Declare PtrSafe Function GetDpiForWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Public Declare PtrSafe Function SetThreadDpiHostingBehavior Lib "user32" (ByVal value As Long) As Long
Public Declare PtrSafe Function CreateWindowExA Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByVal lpParam As LongPtr) As LongPtr
Public Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Public Declare PtrSafe Function ClientToScreen Lib "user32" (ByVal hWnd As LongPtr, ByRef pt As POINT_T) As Long
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
' drawing a page into a picture file (the still that follows what was edited during the show)
Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Public Declare PtrSafe Function FillRect Lib "user32" (ByVal hDC As LongPtr, ByRef rc As RECT_T, ByVal hBrush As LongPtr) As Long
Public Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Public Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As LongPtr, ByVal w As Long, ByVal h As Long) As LongPtr
Public Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal color As Long) As LongPtr
Public Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObj As LongPtr) As LongPtr
Public Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObj As LongPtr) As Long
Public Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Public Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hDC As LongPtr, ByVal x As Long, ByVal y As Long) As Long
Public Declare PtrSafe Function OleDraw Lib "ole32" (ByVal pUnk As LongPtr, ByVal dwAspect As Long, ByVal hdcDraw As LongPtr, ByRef rc As RECT_T) As Long
Public Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef clsid As GUID_T) As Long
Public Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (ByRef token As LongPtr, ByRef inp As GDIP_START_T, ByVal outp As LongPtr) As Long
Public Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr)
Public Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As LongPtr, ByVal hpal As LongPtr, ByRef bitmap As LongPtr) As Long
Public Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus" (ByVal image As LongPtr, ByVal filename As LongPtr, ByRef clsidEncoder As GUID_T, ByVal encoderParams As LongPtr) As Long
Public Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal image As LongPtr) As Long
Public Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
Public Declare PtrSafe Function UrlCreateFromPathW Lib "shlwapi" (ByVal pszPath As LongPtr, ByVal pszUrl As LongPtr, ByRef pcchUrl As Long, ByVal dwFlags As Long) As Long
#If Win64 Then
Public Declare PtrSafe Function GetWindowLongPtrA Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Public Declare PtrSafe Function SetWindowLongPtrA Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Public Declare PtrSafe Function GetWindowLongPtrA Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Public Declare PtrSafe Function SetWindowLongPtrA Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If

Public Const INLINE_TOKEN As String = "InlineHtmlHost_7C41"
Public Const TAG_DECK As String = "InlineHtml"
Public Const TAG_PAGE As String = "HtmlPage"
Public Const AREA_NAME As String = "htmlArea"
Public Const DATA_NS As String = "urn:ppt-html:inline:v1"
Private Const TAG_STATE As String = "HtmlState"
Private Const HEAD_MARK As String = "<!--/ih:head-->"

Private Const MOTW As String = "<!-- saved from url=(0014)about:internet -->"

Private mApp As clsAppEvents
Private mLog As String
Private mLogFile As String      ' "" not looked up yet, "-" none
Private mBusy As Boolean        ' a sync is running (possibly waiting for pages)
Private mDirty As Boolean       ' something changed while busy: sync once more
Private mUp As Boolean          ' frmInline is loaded and holds the browser controls
Private mStale As Boolean       ' the pages belong to a finished show: reload before reuse
Private mFailed As Boolean      ' no browser control on this PC: keep the static pictures
Private mAbort As Boolean       ' the show ended while pages were loading
Private mTop As LongPtr         ' slide show window (top level)
Private mHost As LongPtr        ' window the host form is a child of
Private mCur As Long            ' page on screen (0 = none)
Private mLastTick As Single
Private mForceMode As Long      ' 0 = automatic; 3 = always use the owned borderless window

Private mCount As Long          ' pages of the deck whose data is loaded (0 = none)
Private mDeckKey As String      ' full name of that deck
Private mHtml() As String       ' page documents
Private mBg() As Long           ' page background colours (RGB value)
Private mCanvasW() As Long      ' page canvas widths in CSS px (0 = do not zoom)
Private mPaths() As String      ' temp files of the running show
Private mSettle() As Long       ' time the build gave each page before it took the still picture (ms)
Private mHow() As Long          ' how each page was loaded: 0 temp file with the Mark of the Web, 1 without it, 2 written into the control
Private mTouched() As Boolean   ' pages that stored something during this show: their still is drawn again when the show ends
Private mShowPres As Presentation   ' the deck of the running show

' Two engines can show the pages. The modern one (Microsoft Edge WebView2, module modWv2) runs pages as they
' were written for a current browser; it is used when the file carries such pages and WebView2 is on the PC.
' The classic one (the Internet Explorer 11 browser control in frmInline) is the fallback.
Private mEngine As Long         ' engine of the running show: 0 not decided yet, 1 modern, 2 classic
Private mForceEngine As Long    ' 0 = automatic, 2 = always the classic engine (troubleshooting, tests)
Private mHtmlM() As String      ' page documents for the modern engine ("" = the page has none)
Private mCanvasH() As Long      ' page canvas heights in CSS px
Private mBoxWv As LongPtr       ' container window of the modern views, a child of the slide show window
Private mHoldWv As LongPtr      ' window that is never shown: a page is drawn in it for its still picture after the show
Private mWvShown As Boolean     ' the container is visible
Private mWvX As Long, mWvY As Long, mWvW As Long, mWvH As Long   ' where the container is, in pixels of the show window
Private mWvState() As Long      ' per page: 0 no view, 1 loading, 2 running, -1 failed
Private mWvView() As String     ' viewport the page reported ("1280x720")
Private mWvErr() As String      ' script errors the page reported
Private mWvSeq() As Long        ' how often the page was loaded in this show (part of its file name)
Private mWvUsed() As Long       ' when the page was last on screen (for dropping the least recently used view)
Private mWvFit() As Long        ' zoom corrections made since the view was last sized
Private mWvZoomF() As Double    ' zoom factor of the view
Private mWvClock As Long
Private mWvRun As String        ' part of the file names of this show
Private mWvFolder As String     ' temp folder of the page files of this show
Private mWvAny As Boolean       ' a view was made in this show
Private mEndLater As Boolean    ' the show ended while a sync was running: finish afterwards
Private Const WV_HOST As String = "ppt-html.example"
Private Const WV_MAXVIEWS As Long = 10

' ------------------------------------------------------------------ entry points

' Ribbon onLoad callback (customUI part): runs when the file is opened with macros enabled.
Public Sub InlineHtml_OnLoad(ByVal ribbon As Object)
    On Error Resume Next
    LogIt "ribbon-load"
    HookApp
End Sub

' PowerPoint calls this by name when a slide show starts and on every slide change.
Public Sub OnSlideShowPageChange(ByVal SSW As SlideShowWindow)
    On Error Resume Next
    HookApp
    InlineSync "auto"
End Sub

' PowerPoint calls this by name when the slide show ends.
Public Sub OnSlideShowTerminate(ByVal SSW As SlideShowWindow)
    On Error Resume Next
    InlineShowEnded "auto"
End Sub

' Action of the htmlArea picture (click / mouse over): starts the inline page if it is not up yet.
Public Sub InlineHtml_Activate()
    On Error Resume Next
    HookApp
    InlineSync "action"
End Sub

Public Sub HookApp(Optional ByVal unused As Long)
    On Error Resume Next
    If mApp Is Nothing Then
        Set mApp = New clsAppEvents
        Set mApp.App = Application
        LogIt "hook"
    End If
End Sub

Public Sub InlineShowBegan(Optional ByVal unused As Long)
    On Error Resume Next
    LogIt "begin"
    If mEngine = 1 Then
        If mEndLater And Not mBusy Then EndModern "begin"
    ElseIf mUp And mStale And Not mBusy Then
        Unload frmInline
        mUp = False
        mStale = False
        mEngine = 0
    End If
    mFailed = False
    InlineSync "begin"
End Sub

' ------------------------------------------------------------------ log / status

' Troubleshooting: 0 = automatic (child window inside the slide show), 3 = borderless window over the slide show.
Public Sub InlineHtml_SetMode(ByVal m As Long)
    mForceMode = m
End Sub

' Troubleshooting / tests: 0 = automatic (modern engine when possible), 2 = always the classic engine.
' Takes effect with the next slide show.
Public Sub InlineHtml_SetEngine(ByVal m As Long)
    mForceEngine = m
End Sub

' Which engine shows the pages: "modern", "classic" or "" (no show running yet).
Public Function InlineHtml_Engine() As String
    If mEngine = 1 Then InlineHtml_Engine = "modern"
    If mEngine = 2 Then InlineHtml_Engine = "classic"
End Function

' The browser's own picture of page n, as it is now, written to a PNG file (modern engine). "ok" or why not.
Public Function InlineHtml_Snap(ByVal n As Long, ByVal file As String) As String
    On Error Resume Next
    If mEngine <> 1 Or Not mUp Or n < 1 Or n > mCount Then
        InlineHtml_Snap = "SNAP-NOPAGE"
        Exit Function
    End If
    If WvSnap(n, file) Then
        InlineHtml_Snap = "ok"
    Else
        InlineHtml_Snap = "SNAP-FAILED"
    End If
End Function

' A command of the browser's developer protocol for page n (modern engine; the show test uses it for input).
Public Function InlineHtml_Cdp(ByVal n As Long, ByVal method As String, ByVal json As String) As String
    On Error Resume Next
    If mEngine <> 1 Or Not mUp Or n < 1 Or n > mCount Then
        InlineHtml_Cdp = "CDP-NOPAGE"
        Exit Function
    End If
    InlineHtml_Cdp = WvCdp(n, method, json)
End Function

Public Sub LogIt(ByVal s As String)
    Dim f As Integer
    If Len(mLog) > 8000 Then mLog = Right$(mLog, 4000)
    mLog = mLog & Format$(Timer, "0.00") & " " & s & vbCrLf
    ' troubleshooting: the environment variable PPTHTML_LOG names a file that gets every line as it happens
    On Error Resume Next
    If Len(mLogFile) = 0 Then
        mLogFile = Environ$("PPTHTML_LOG")
        If Len(mLogFile) = 0 Then mLogFile = "-"
    End If
    If mLogFile <> "-" Then
        f = FreeFile
        Open mLogFile For Append As #f
        Print #f, Format$(Now, "hh:nn:ss") & " " & Format$(Timer, "0.00") & " " & s
        Close #f
    End If
End Sub

Public Function InlineHtml_Status() As String
    Dim r As String
    On Error Resume Next
    r = "pages=" & mCount & " up=" & mUp & " busy=" & mBusy & " stale=" & mStale & " failed=" & mFailed & " cur=" & mCur
    r = r & " top=" & CStr(mTop) & " host=" & CStr(mHost) & " hooked=" & Not (mApp Is Nothing)
    If mUp Then
        If mEngine = 1 Then
            r = r & " | " & ModernDiag()
        Else
            r = r & " | " & frmInline.Diag()
        End If
    End If
    InlineHtml_Status = r & vbCrLf & mLog
End Function

' Troubleshooting: what page n looks like from the inside (data, temp file, browser state).
Public Function InlineHtml_Probe(ByVal n As Long) As String
    Dim r As String, wb As Object, doc As Object, s As String
    On Error Resume Next
    r = "page=" & n & " of " & mCount
    If n < 1 Or n > mCount Then
        InlineHtml_Probe = r
        Exit Function
    End If
    r = r & " dataLen=" & Len(mHtml(n)) & " modernLen=" & Len(mHtmlM(n)) & " canvas=" & mCanvasW(n) & " bg=" & mBg(n)
    r = r & " head=[" & Left$(mHtml(n), 60) & "]"
    r = r & " path=[" & mPaths(n) & "]"
    If Len(mPaths(n)) > 0 Then r = r & " fileLen=" & FileLen(mPaths(n))
    If mUp And mEngine = 1 Then
        r = r & " engine=modern state=" & mWvState(n) & " view=" & mWvView(n) & " folder=[" & mWvFolder & "] url=" & WvEval(n, "location.href") & " ua=" & WvEval(n, "navigator.userAgent")
    ElseIf mUp Then
        Set wb = frmInline.Browser(n)
        r = r & " url=[" & wb.LocationURL & "] readyState=" & wb.ReadyState & " busy=" & wb.Busy
        Set doc = wb.Document
        r = r & " docMode=" & doc.documentMode & " title=[" & doc.Title & "] charset=" & doc.Charset
        r = r & " ready=[" & doc.body.getAttribute("data-ih-ready") & "] err=[" & doc.documentElement.getAttribute("data-ih-err") & "]"
        r = r & " scripts=" & doc.scripts.Length & " bodyLen=" & Len(doc.body.innerHTML)
        s = ""
        s = doc.parentWindow.eval("typeof pptHtml")
        r = r & " eval=[" & s & "] evalErr=" & Err.Number
        Err.Clear
        r = r & " zone=[" & doc.security & "] domain=[" & doc.domain & "] urlDoc=[" & doc.URL & "]"
    End If
    InlineHtml_Probe = r
End Function

' Runs a JavaScript expression in page n and returns its value as text (show test, troubleshooting).
Public Function InlineHtml_Eval(ByVal n As Long, ByVal code As String) As String
    Dim v As Variant
    On Error Resume Next
    If Not mUp Or n < 1 Or n > mCount Then
        InlineHtml_Eval = "EVAL-NOPAGE"
        Exit Function
    End If
    If mEngine = 1 Then
        InlineHtml_Eval = JsonText(WvEval(n, code))
        Exit Function
    End If
    v = frmInline.Browser(n).Document.parentWindow.eval(code)
    If Err.Number <> 0 Then
        InlineHtml_Eval = "EVAL-ERR " & Err.Description
        Err.Clear
        Exit Function
    End If
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    InlineHtml_Eval = CStr(v)
End Function

' ------------------------------------------------------------------ what pages keep (pptHtml.saveState)
' A page may ask the host to keep a text for it (edits made during the show). It is stored in tags of the
' page's slide, in pieces, and handed back to the page the next time it is loaded. The text is URI-encoded by
' the page, so it is plain ASCII without quotes.

Private Function PageSlide(ByVal pres As Presentation, ByVal n As Long) As Slide
    Dim s As Slide
    On Error Resume Next
    For Each s In pres.Slides
        If Val(s.Tags(TAG_PAGE)) = n Then
            Set PageSlide = s
            Exit Function
        End If
    Next s
End Function

Private Function ReadState(ByVal pres As Presentation, ByVal n As Long) As String
    Dim s As Slide, k As Long, i As Long, r As String
    On Error Resume Next
    Set s = PageSlide(pres, n)
    If s Is Nothing Then Exit Function
    k = CLng(Val(s.Tags(TAG_STATE & "0")))
    If k > 200 Then Exit Function
    For i = 1 To k
        r = r & s.Tags(TAG_STATE & CStr(i))
    Next i
    ReadState = r
End Function

Private Sub WriteState(ByVal pres As Presentation, ByVal n As Long, ByVal text As String)
    Dim s As Slide, k As Long, i As Long, old As Long
    On Error Resume Next
    Set s = PageSlide(pres, n)
    If s Is Nothing Then Exit Sub
    If Len(text) > 400000 Then Exit Sub
    If InStr(text, """") > 0 Or InStr(text, "<") > 0 Or InStr(text, "\") > 0 Or InStr(text, vbCr) > 0 Or InStr(text, vbLf) > 0 Then Exit Sub
    old = CLng(Val(s.Tags(TAG_STATE & "0")))
    If old > 200 Then old = 200
    For i = 0 To old
        s.Tags.Delete TAG_STATE & CStr(i)
    Next i
    If Len(text) > 0 Then
        k = (Len(text) + 3999) \ 4000
        s.Tags.Add TAG_STATE & "0", CStr(k)
        For i = 1 To k
            s.Tags.Add TAG_STATE & CStr(i), Mid$(text, (i - 1) * 4000 + 1, 4000)
        Next i
    End If
    pres.Saved = msoFalse
    LogIt "state " & n & " len=" & Len(text)
End Sub

' The document of page n as it is loaded: the page plus what it stored earlier.
Private Function PageDoc(ByVal n As Long, Optional ByVal modern As Boolean = False) As String
    Dim st As String, ssw As SlideShowWindow, p As Long, pres As Presentation, doc As String
    On Error Resume Next
    If modern Then doc = mHtmlM(n) Else doc = mHtml(n)
    PageDoc = doc
    ' the deck the host remembers comes first: while a show ends its window may still be listed, but dead
    Set pres = mShowPres
    If pres Is Nothing Then
        Set ssw = FindShow()
        If Not ssw Is Nothing Then Set pres = ssw.Presentation
    End If
    If pres Is Nothing Then Exit Function
    st = ReadState(pres, n)
    If Len(st) = 0 Then Exit Function
    p = InStr(1, doc, HEAD_MARK, vbBinaryCompare)
    If p = 0 Then Exit Function
    PageDoc = Left$(doc, p + Len(HEAD_MARK) - 1) & "<script>window.__ihState=""" & st & """;</script>" & Mid$(doc, p + Len(HEAD_MARK))
End Function

' ------------------------------------------------------------------ core

Private Function FindShow() As SlideShowWindow
    Dim i As Long, w As SlideShowWindow, mark As String
    On Error Resume Next
    For i = 1 To SlideShowWindows.Count
        Set w = Nothing
        Set w = SlideShowWindows(i)
        If Not w Is Nothing Then
            mark = ""
            mark = w.Presentation.Tags(TAG_DECK)
            If mark = "1" Then
                Set FindShow = w
                Exit Function
            End If
        End If
    Next i
End Function

' Reads the pages of a deck from its custom XML part:
'   <inlineHtml xmlns="urn:ppt-html:inline:v1"><page><bg>RGB value</bg><w>canvas px</w><h>canvas px</h><settle>ms</settle>
'     <html>document for the classic engine</html><htmlm>document for the modern engine</htmlm></page>...</inlineHtml>
' Either document may be missing or empty.
Private Function LoadDeck(ByVal pres As Presentation) As Boolean
    Dim parts As Object, root As Object, pg As Object, c As Object
    Dim i As Long, n As Long, key As String
    On Error GoTo Fail
    key = pres.FullName
    If mCount > 0 And key = mDeckKey Then
        LoadDeck = True
        Exit Function
    End If
    If mUp Then Exit Function
    mCount = 0
    mDeckKey = ""
    Set parts = pres.CustomXMLParts.SelectByNamespace(DATA_NS)
    If parts.Count < 1 Then
        LogIt "no-data-part"
        Exit Function
    End If
    Set root = parts.Item(1).DocumentElement
    For Each pg In root.ChildNodes
        If pg.NodeType = 1 Then
            If pg.BaseName = "page" Then n = n + 1
        End If
    Next pg
    If n < 1 Then
        LogIt "no-pages"
        Exit Function
    End If
    ReDim mHtml(1 To n)
    ReDim mHtmlM(1 To n)
    ReDim mBg(1 To n)
    ReDim mCanvasW(1 To n)
    ReDim mCanvasH(1 To n)
    ReDim mPaths(1 To n)
    ReDim mSettle(1 To n)
    ReDim mHow(1 To n)
    ReDim mTouched(1 To n)
    For Each pg In root.ChildNodes
        If pg.NodeType = 1 Then
            If pg.BaseName = "page" Then
                i = i + 1
                mBg(i) = RGB(255, 255, 255)
                For Each c In pg.ChildNodes
                    If c.NodeType = 1 Then
                        Select Case c.BaseName
                            Case "bg": mBg(i) = CLng(Val(c.Text))
                            Case "w": mCanvasW(i) = CLng(Val(c.Text))
                            Case "h": mCanvasH(i) = CLng(Val(c.Text))
                            Case "settle": mSettle(i) = CLng(Val(c.Text))
                            Case "html": mHtml(i) = c.Text
                            Case "htmlm": mHtmlM(i) = c.Text
                        End Select
                    End If
                Next c
            End If
        End If
    Next pg
    mCount = n
    mDeckKey = key
    LogIt "data " & n & " pages"
    LoadDeck = True
    Exit Function
Fail:
    LogIt "data-err " & Err.Number & " " & Err.Description
    mCount = 0
    mDeckKey = ""
End Function

' Brings the overlay in line with the slide show: right page, right place, or hidden.
Public Sub InlineSync(ByVal why As String)
    Dim ssw As SlideShowWindow, sld As Slide, shp As Shape
    Dim n As Long, st As Long, guard As Long
    Dim hTop As LongPtr, hHost As LongPtr
    Dim x As Long, y As Long, w As Long, h As Long
    If mBusy Then
        mDirty = True
        Exit Sub
    End If
    mBusy = True
    On Error GoTo Fail
    Do
        mDirty = False
        guard = guard + 1
        Set ssw = FindShow()
        If ssw Is Nothing Then
            HideOverlay True
        Else
            n = 0
            st = 0
            hTop = 0
            Set sld = Nothing
            Set shp = Nothing
            On Error Resume Next
            Set mShowPres = ssw.Presentation
            st = ssw.View.State
            If st = 1 Or st = 2 Then
                Set sld = ssw.View.Slide
                If Not sld Is Nothing Then
                    n = Val(sld.Tags(TAG_PAGE))
                    If n >= 1 Then
                        If Not LoadDeck(ssw.Presentation) Then n = 0
                    End If
                    If n >= 1 And n <= mCount Then Set shp = sld.Shapes(AREA_NAME)
                    If shp Is Nothing Then n = 0
                End If
            End If
            hTop = ShowHwnd(ssw)
            Err.Clear
            On Error GoTo Fail
            If n < 1 Or n > mCount Or mFailed Or hTop = 0 Then
                HideOverlay False
            Else
                hHost = HostWindow(hTop)
                If SlideRectPx(ssw, hHost, shp, x, y, w, h) Then
                    ShowOverlay hTop, hHost, n, x, y, w, h
                Else
                    HideOverlay False
                End If
            End If
        End If
    Loop While mDirty And guard < 4
    mBusy = False
    If mEndLater Then EndModern "later"
    Exit Sub
Fail:
    LogIt "sync-err(" & why & ") " & Err.Number & " " & Err.Description
    mBusy = False
    If mEndLater Then EndModern "later"
End Sub

' Window handle of the slide show. (SlideShowWindow.HWND cannot be used from 64-bit VBA, so the window
' is looked up: a top-level "screenClass" (full screen) or "PPTFrameClass" (window) of this process whose
' caption carries "[" + presentation name, e.g. "PowerPoint Slide Show - [deck.pptm]" or "... - [deck.pptm [Read-Only]]".)
Private Function ShowHwnd(ByVal ssw As SlideShowWindow) As LongPtr
    Dim nm As String, h As LongPtr, firstScreen As LongPtr, pass As Long, cls As String
    Dim o As Object, v As Variant
    On Error Resume Next
    ' where the object model offers the handle (late bound: the member is not usable early bound in 64-bit VBA)
    Set o = ssw
    v = o.HWND
    If Err.Number = 0 And Not IsEmpty(v) Then
        h = CLngPtr(v)
        If h <> 0 Then
            If IsWindow(h) <> 0 And OwnWindow(h) Then
                ShowHwnd = h
                Exit Function
            End If
        End If
    End If
    Err.Clear
    h = 0
    nm = "[" & ssw.Presentation.Name
    For pass = 1 To 2
        If pass = 1 Then cls = "screenClass" Else cls = "PPTFrameClass"
        h = 0
        Do
            h = FindWindowExA(0, h, cls, vbNullString)
            If h = 0 Then Exit Do
            If OwnWindow(h) And IsWindowVisible(h) <> 0 Then
                If InStr(1, WindowTitle(h), nm, vbBinaryCompare) > 0 Then
                    ShowHwnd = h
                    Exit Function
                End If
                If pass = 1 And firstScreen = 0 Then firstScreen = h
            End If
        Loop
    Next pass
    If SlideShowWindows.Count = 1 Then ShowHwnd = firstScreen
End Function

Private Function OwnWindow(ByVal h As LongPtr) As Boolean
    Dim p As Long
    GetWindowThreadProcessId h, p
    OwnWindow = (p = GetCurrentProcessId())
End Function

Private Function WindowTitle(ByVal h As LongPtr) As String
    Dim buf As String, n As Long
    buf = String$(600, vbNullChar)
    n = GetWindowTextW(h, StrPtr(buf), 599)
    If n > 0 Then WindowTitle = Left$(buf, n)
End Function

' Full screen show: the slide is drawn in the show window itself. Window show: in its MDI child.
Private Function HostWindow(ByVal hTop As LongPtr) As LongPtr
    Dim c As LongPtr
    HostWindow = hTop
    c = FindWindowExA(hTop, 0, "MDIClient", vbNullString)
    If c <> 0 Then c = FindWindowExA(c, 0, "mdiClass", vbNullString)
    If c <> 0 Then HostWindow = c
End Function

' Pixel rectangle of the shape inside the host window (the slide is fitted and centred there).
Private Function SlideRectPx(ByVal ssw As SlideShowWindow, ByVal hHost As LongPtr, ByVal shp As Shape, ByRef x As Long, ByRef y As Long, ByRef w As Long, ByRef h As Long) As Boolean
    Dim rc As RECT_T, rs As RECT_T, hBar As LongPtr
    Dim cw As Double, ch As Double, sw As Double, sh As Double, k As Double, ox As Double, oy As Double
    Dim x2 As Long, y2 As Long
    On Error GoTo Fail
    If GetClientRect(hHost, rc) = 0 Then Exit Function
    cw = rc.Right - rc.Left
    ch = rc.Bottom - rc.Top
    hBar = FindWindowExA(hHost, 0, "NUIScrollbar", vbNullString)
    If hBar <> 0 Then
        If IsWindowVisible(hBar) <> 0 Then
            If GetClientRect(hBar, rs) <> 0 Then cw = cw - (rs.Right - rs.Left)
        End If
    End If
    sw = ssw.Presentation.PageSetup.SlideWidth
    sh = ssw.Presentation.PageSetup.SlideHeight
    If cw < 80 Or ch < 60 Or sw <= 0 Or sh <= 0 Then Exit Function
    k = cw / sw
    If ch / sh < k Then k = ch / sh
    ox = (cw - sw * k) / 2
    oy = (ch - sh * k) / 2
    x = Int(ox + shp.Left * k + 0.01)
    y = Int(oy + shp.Top * k + 0.01)
    x2 = -Int(-(ox + (shp.Left + shp.Width) * k - 0.01))
    y2 = -Int(-(oy + (shp.Top + shp.Height) * k - 0.01))
    w = x2 - x
    h = y2 - y
    SlideRectPx = (w > 20 And h > 20)
    Exit Function
Fail:
    SlideRectPx = False
End Function

Private Sub ShowOverlay(ByVal hTop As LongPtr, ByVal hHost As LongPtr, ByVal n As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    Dim fresh As Boolean, i As Long
    If mEngine = 0 Then DecideEngine
    If mEngine = 1 Then
        ' (False = the modern engine could not make its first view: the classic engine takes over below)
        If ShowModern(hTop, hHost, n, x, y, w, h) Then Exit Sub
    End If
    If Len(mHtml(n)) = 0 Then
        ' this page has no document for the classic engine: its still picture stays
        HideOverlay False
        Exit Sub
    End If
    If Not mUp Then
        ' another open deck with this code may already serve this show
        If OtherHostPresent(hHost) Then Exit Sub
        Load frmInline
        If Not frmInline.Setup(mCount) Then
            mFailed = True
            LogIt "no-browser " & frmInline.Info
            Unload frmInline
            Exit Sub
        End If
        For i = 1 To mCount
            frmInline.SetCanvas i, mCanvasW(i)
        Next i
        mUp = True
        fresh = True
        LogIt "form " & CStr(frmInline.FormHwnd)
    End If
    If mStale Then
        fresh = True
        mStale = False
    End If
    If mHost <> hHost Or Not frmInline.AttachedTo(hHost) Then
        frmInline.Attach hHost, hTop, mForceMode
        mHost = hHost
        LogIt "attach " & CStr(hHost) & " mode=" & frmInline.Mode
    End If
    mTop = hTop
    frmInline.Place x, y, w, h
    If fresh Then
        mAbort = False
        frmInline.HideNow
        mCur = 0
        LoadPages
        If mAbort Then Exit Sub
    End If
    If frmInline.PageReady(n) Then
        If n <> mCur Then LogIt "page " & n & " @" & x & "," & y & " " & w & "x" & h
        frmInline.ShowPage n, mBg(n)
        frmInline.ShowNow
        frmInline.Refit
        mCur = n
    Else
        frmInline.HideNow
        mCur = 0
    End If
End Sub

Private Function OtherHostPresent(ByVal hHost As LongPtr) As Boolean
    Dim b As LongPtr
    If FindWindowExA(hHost, 0, "ThunderDFrame", INLINE_TOKEN) <> 0 Then
        OtherHostPresent = True
        Exit Function
    End If
    ' (the container of the modern engine carries the same token as its window text)
    If FindWindowExA(hHost, 0, "Static", INLINE_TOKEN) <> 0 Then
        OtherHostPresent = True
        Exit Function
    End If
    Do
        b = FindWindowExA(hHost, b, "Static", vbNullString)
        If b = 0 Then Exit Do
        If FindWindowExA(b, 0, "ThunderDFrame", INLINE_TOKEN) <> 0 Then
            OtherHostPresent = True
            Exit Function
        End If
    Loop
End Function

Private Sub HideOverlay(ByVal detach As Boolean)
    If Not mUp Then Exit Sub
    If mEngine = 1 Then
        HideModern
        ' (no slide show any more: the modern engine is closed once the running sync is over)
        If detach Then mEndLater = True
        Exit Sub
    End If
    If frmInline.Shown Then
        ReturnFocus
        frmInline.HideNow
        LogIt "hide"
    End If
    If detach And mHost <> 0 Then
        frmInline.Detach
        mHost = 0
        mTop = 0
        mStale = True
        LogIt "detach"
    End If
    mCur = 0
End Sub

' The slide show is over: take the host out of the show window and drop it.
Public Sub InlineShowEnded(ByVal why As String)
    On Error Resume Next
    If Not mUp Then Exit Sub
    If mEngine = 1 Then
        EndModern why
        Exit Sub
    End If
    LogIt "end(" & why & ") busy=" & mBusy & " formAlive=" & (IsWindow(frmInline.FormHwnd) <> 0)
    mAbort = True
    frmInline.HideNow
    frmInline.Detach
    mHost = 0
    mTop = 0
    mCur = 0
    If mBusy Then
        ' called from inside a running sync (page callback or wait loop): unload later
        mStale = True
    Else
        ' pages that stored something are drawn again for their slides before the host goes (other calls wait meanwhile)
        mBusy = True
        RefreshStills
        mBusy = False
        Unload frmInline
        mUp = False
        mStale = False
        mEngine = 0
        CleanTemp
        Set mShowPres = Nothing
        If Not FindShow() Is Nothing Then InlineSync "after-end"
    End If
End Sub

' ------------------------------------------------------------------ pages

Private Sub LoadPages()
    Dim i As Long, again As Boolean
    ' (a page without a document for this engine is left alone: its slide keeps the still picture)
    For i = 1 To mCount
        mHow(i) = 0
        mTouched(i) = False
        mPaths(i) = ""
        If Len(mHtml(i)) > 0 Then
            mPaths(i) = WriteTempHtml(i, PageDoc(i), True)
            frmInline.Go i, mPaths(i)
        End If
    Next i
    WaitPages 6 + mCount * 0.5, 0
    If mAbort Then Exit Sub
    ' zone policy differs per PC: if the script did not start, retry without the Mark of the Web ...
    For i = 1 To mCount
        If Len(mHtml(i)) > 0 And Not frmInline.PageAlive(i) Then
            LogIt "retry-nomotw " & i
            mHow(i) = 1
            mPaths(i) = WriteTempHtml(i, PageDoc(i), False)
            frmInline.Go i, mPaths(i)
            again = True
        End If
    Next i
    If again Then WaitPages 4 + mCount * 0.3, 0.3
    If mAbort Then Exit Sub
    ' ... and finally by writing the document straight into an empty page of the control
    again = False
    For i = 1 To mCount
        If Len(mHtml(i)) > 0 And Not frmInline.PageAlive(i) Then
            LogIt "retry-write " & i
            frmInline.Go i, ""
            again = True
        End If
    Next i
    If again Then
        WaitPages 3 + mCount * 0.2, 0.2
        If mAbort Then Exit Sub
        For i = 1 To mCount
            If Len(mHtml(i)) > 0 And Not frmInline.PageAlive(i) Then
                mHow(i) = 2
                frmInline.WriteDirect i, PageDoc(i)
            End If
        Next i
        WaitPages 2 + mCount * 0.2, 0.2
    End If
    If mAbort Then Exit Sub
    For i = 1 To mCount
        If frmInline.PageAlive(i) Then frmInline.Hook i
    Next i
    LogIt "loaded " & frmInline.Info
End Sub

' Pumps messages until every page is loaded (script running) or finished without script, or time is up.
Private Sub WaitPages(ByVal secs As Single, ByVal minWait As Single)
    Dim t0 As Single, el As Single, i As Long, done As Boolean
    t0 = Timer
    Do
        DoEvents
        Sleep 5
        If mAbort Then Exit Do
        el = Timer - t0
        If el < 0 Then el = el + 86400!
        done = (el >= minWait)
        If done Then
            For i = 1 To mCount
                If Len(mHtml(i)) > 0 And Not frmInline.PageAlive(i) Then
                    If Not frmInline.PageComplete(i) Then
                        done = False
                        Exit For
                    End If
                End If
            Next i
        End If
    Loop While Not done And el < secs
End Sub

' Text -> UTF-8 bytes.
Private Function Utf8Bytes(ByVal s As String, ByRef b() As Byte) As Boolean
    Dim n As Long
    If Len(s) = 0 Then Exit Function
    n = WideCharToMultiByte(65001, 0, StrPtr(s), Len(s), 0, 0, 0, 0)
    If n <= 0 Then Exit Function
    ReDim b(0 To n - 1)
    n = WideCharToMultiByte(65001, 0, StrPtr(s), Len(s), VarPtr(b(0)), n, 0, 0)
    Utf8Bytes = (n > 0)
End Function

' Writes page N to %TEMP%\ppthtml_<process>_N<tag>.html and returns the path ("" if that is not possible).
' The file is UTF-8 without a byte order mark (the page declares the charset itself), so that it starts with
' plain "<!DOCTYPE html>" + the Mark of the Web, which the browser only recognises in single-byte text.
Private Function WriteTempHtml(ByVal n As Long, ByVal html As String, ByVal withMotw As Boolean, Optional ByVal tag As String = "") As String
    Dim p As String, f As Integer, b() As Byte
    On Error GoTo Fail
    If Len(html) = 0 Then Exit Function
    p = Environ$("TEMP")
    If Len(p) = 0 Then p = Environ$("TMP")
    If Len(p) = 0 Then p = CurDir$
    If Right$(p, 1) <> "\" Then p = p & "\"
    p = p & "ppthtml_" & CStr(GetCurrentProcessId()) & "_" & CStr(n) & tag & ".html"
    ' CR LF line ends throughout: the browser only honours the Mark of the Web when its line ends that way
    html = Replace(Replace(html, vbCrLf, vbLf), vbLf, vbCrLf)
    If withMotw Then html = Replace(html, "<!DOCTYPE html>", "<!DOCTYPE html>" & vbCrLf & MOTW, 1, 1)
    If Not Utf8Bytes(html, b) Then Exit Function
    If Len(Dir$(p)) > 0 Then Kill p
    f = FreeFile
    Open p For Binary Access Write As #f
    Put #f, , b
    Close #f
    WriteTempHtml = p
    Exit Function
Fail:
    On Error Resume Next
    Close #f
    WriteTempHtml = ""
End Function

Private Sub CleanTemp()
    Dim i As Long
    On Error Resume Next
    For i = 1 To mCount
        If Len(mPaths(i)) > 0 Then Kill mPaths(i)
        mPaths(i) = ""
    Next i
End Sub

' ------------------------------------------------------------------ page -> host

' Called by the bridge object when a page clicks its hidden #ihCmd element.
' Commands: tick | focus | key<code> | next | prev | goto<slide number> | save (text in data-arg)
Public Sub InlineBridge(ByVal page As Long)
    Dim c As String, t As Single, k As Long, ssw As SlideShowWindow
    On Error GoTo Fail
    If Not mUp Then Exit Sub
    c = frmInline.PageCmd(page)
    If c = "tick" Then
        t = Timer
        If t >= mLastTick And t - mLastTick < 0.3 Then Exit Sub
        mLastTick = t
        InlineSync "tick"
    ElseIf page <> mCur Then
        ' only the page on screen may steer the show
    ElseIf c = "focus" Then
        ReturnFocus
    ElseIf Left$(c, 3) = "key" Then
        ForwardKey CLng(Val(Mid$(c, 4)))
    ElseIf c = "save" Then
        ' (a page may still store something while the show window is already going away: the remembered deck is used)
        If mShowPres Is Nothing Then
            Set ssw = FindShow()
            If Not ssw Is Nothing Then Set mShowPres = ssw.Presentation
        End If
        If Not mShowPres Is Nothing Then
            WriteState mShowPres, page, frmInline.PageArg(page)
            If page >= 1 And page <= mCount Then mTouched(page) = True
        End If
    ElseIf c = "next" Then
        ForwardKey 39
    ElseIf c = "prev" Then
        ForwardKey 37
    ElseIf Left$(c, 4) = "goto" Then
        k = CLng(Val(Mid$(c, 5)))
        Set ssw = FindShow()
        If Not ssw Is Nothing Then
            If k >= 1 And k <= ssw.Presentation.Slides.Count Then
                ReturnFocus
                ssw.View.GotoSlide k
            End If
        End If
    End If
    Exit Sub
Fail:
    LogIt "bridge-err " & Err.Number & " " & Err.Description
End Sub

' Gives the keyboard back to the slide show (arrow keys, Esc, B ... keep working after a click on the page).
Private Sub ReturnFocus()
    Dim f As LongPtr
    On Error Resume Next
    If mTop = 0 Or mHost = 0 Then Exit Sub
    If IsWindow(mHost) = 0 Then Exit Sub
    If mEngine = 1 Then
        ' (the focused window of a modern view belongs to the browser: the show window simply takes the keyboard)
        SetFocusApi mHost
        Exit Sub
    End If
    f = GetFocusApi()
    If f = 0 Then Exit Sub
    If f = frmInline.FormHwnd Or IsChild(frmInline.FormHwnd, f) <> 0 Then SetFocusApi mHost
End Sub

' A key pressed while the page had the focus: pass it on to the slide show window.
Private Sub ForwardKey(ByVal vk As Long)
    Dim sc As Long
    On Error Resume Next
    If mHost = 0 Then Exit Sub
    If IsWindow(mHost) = 0 Then Exit Sub
    If vk <= 0 Or vk > 254 Then Exit Sub
    If vk = 16 Or vk = 17 Or vk = 18 Or vk = 91 Or vk = 92 Then Exit Sub
    ReturnFocus
    sc = MapVirtualKeyA(vk, 0)
    PostMessageA mHost, &H100, vk, 1 Or (sc * &H10000)
    PostMessageA mHost, &H101, vk, 1 Or (sc * &H10000) Or &HC0000000
End Sub

' ------------------------------------------------------------------ the still pictures follow the edits
' A page that stored something during the show (edited text, moved cards, votes ...) is loaded once more with
' what it stored and drawn; the drawing replaces the still picture on its slide. So the slide looks the same
' outside the slide show: in the editing view, in thumbnails, in print and on PCs without macros.
Private Sub RefreshStills()
    Dim i As Long, pres As Presentation, f As String, tmp As String, some As Boolean, w As Single, p As String
    On Error Resume Next
    Set pres = mShowPres
    If pres Is Nothing Then Exit Sub
    For i = 1 To mCount
        If mTouched(i) Then some = True
    Next i
    If Not some Then Exit Sub
    If IsWindow(frmInline.FormHwnd) = 0 Then
        LogIt "still-skip host window gone"
        Exit Sub
    End If
    tmp = Environ$("TEMP")
    If Len(tmp) = 0 Then tmp = Environ$("TMP")
    If Len(tmp) = 0 Then tmp = CurDir$
    If Right$(tmp, 1) <> "\" Then tmp = tmp & "\"
    For i = 1 To mCount
        If mTouched(i) Then
            mTouched(i) = False
            If mHow(i) = 2 Then
                LogIt "still-skip " & i & " written page"
            Else
                frmInline.MarkReload i
                ' a file of its own: navigating to the file the control already shows would give the cached page
                p = WriteTempHtml(i, PageDoc(i), (mHow(i) = 0), "s")
                If Len(p) > 0 Then
                    frmInline.Go i, p
                    PumpUntilAlive i, 8
                    If frmInline.PageAlive(i) Then
                        ' the page gets its canvas size, then the time the build gave it to draw itself
                        frmInline.FitPage i
                        w = mSettle(i) / 1000!
                        If w < 0.9 Then w = 0.9
                        If w > 5 Then w = 5
                        PumpFor w
                        frmInline.FitPage i
                        f = tmp & "ppthtml_" & CStr(GetCurrentProcessId()) & "_still" & CStr(i) & ".png"
                        If frmInline.Snap(i, f) Then
                            ReplaceStill pres, i, f, frmInline.SnapWhy
                        Else
                            LogIt "still-fail " & i & " " & frmInline.SnapWhy
                        End If
                        If Len(Dir$(f)) > 0 Then Kill f
                    Else
                        LogIt "still-timeout " & i
                    End If
                    Kill p
                End If
            End If
        End If
    Next i
End Sub

Private Sub PumpFor(ByVal secs As Single)
    Dim t0 As Single, el As Single
    t0 = Timer
    Do
        DoEvents
        Sleep 5
        el = Timer - t0
        If el < 0 Then el = el + 86400!
    Loop While el < secs
End Sub

Private Sub PumpUntilAlive(ByVal i As Long, ByVal secs As Single)
    Dim t0 As Single, el As Single
    t0 = Timer
    Do
        DoEvents
        Sleep 5
        el = Timer - t0
        If el < 0 Then el = el + 86400!
    Loop While Not frmInline.PageAlive(i) And el < secs
End Sub

' Puts a new picture in the place of the still of page n: same position, size, name, order and action.
Private Sub ReplaceStill(ByVal pres As Presentation, ByVal n As Long, ByVal file As String, ByVal how As String)
    Dim s As Slide, old As Shape, pic As Shape, z As Long, guard As Long, k As Long
    Dim l As Single, t As Single, w As Single, h As Single, alt As String
    On Error GoTo Fail
    Set s = PageSlide(pres, n)
    If s Is Nothing Then Exit Sub
    Set old = s.Shapes(AREA_NAME)
    z = old.ZOrderPosition
    l = old.Left: t = old.Top: w = old.Width: h = old.Height
    alt = old.AlternativeText
    Set pic = s.Shapes.AddPicture(file, msoFalse, msoTrue, l, t, w, h)
    pic.LockAspectRatio = msoFalse
    pic.Left = l: pic.Top = t: pic.Width = w: pic.Height = h
    pic.AlternativeText = alt
    For k = 1 To 2
        pic.ActionSettings(k).Action = ppActionRunMacro
        pic.ActionSettings(k).Run = "InlineHtml_Activate"
    Next k
    old.Delete
    Set old = Nothing
    pic.Name = AREA_NAME
    On Error Resume Next
    Do While pic.ZOrderPosition > z And guard < 500
        pic.ZOrder msoSendBackward
        guard = guard + 1
    Loop
    LogIt "still " & n & " " & how
    Exit Sub
Fail:
    LogIt "still-err " & n & " " & Err.Number & " " & Err.Description
    On Error Resume Next
    If Not pic Is Nothing Then
        If Not old Is Nothing Then pic.Delete
    End If
End Sub


' ------------------------------------------------------------------ modern engine (WebView2, see modWv2)
' One browser view per page, made when its slide (or the slide before it) comes on screen. The views are
' children of a container window that sits over the slide area inside the slide show window. Until a view
' has loaded its page, the still picture of the slide is what one sees.

Private Sub DecideEngine()
    Dim i As Long, some As Boolean, want As String
    mEngine = 2
    ' the presentation tag HtmlEngine = classic asks for the classic engine (tests, troubleshooting)
    On Error Resume Next
    want = ""
    If Not mShowPres Is Nothing Then want = LCase$(mShowPres.Tags("HtmlEngine"))
    Err.Clear
    On Error GoTo 0
    If mForceEngine = 2 Or want = "classic" Then
        LogIt "engine classic (asked)"
        Exit Sub
    End If
    For i = 1 To mCount
        If Len(mHtmlM(i)) > 0 Then some = True
    Next i
    If Not some Then
        LogIt "engine classic (the file has no pages for the modern engine)"
        Exit Sub
    End If
    If WvStart() Then
        mEngine = 1
        LogIt "engine modern"
    Else
        LogIt "engine classic: " & WvWhy
    End If
End Sub

' Container of the views: a child of the slide show window, made in that window's DPI context.
Private Function MakeWvBox(ByVal hHost As LongPtr) As LongPtr
    Dim oldCtx As LongPtr, gotCtx As Boolean
    On Error Resume Next
    oldCtx = SetThreadDpiAwarenessContext(GetWindowDpiAwarenessContext(hHost))
    gotCtx = (Err.Number = 0 And oldCtx <> 0)
    Err.Clear
    ' child window that clips its children and siblings; an owner-drawn static paints nothing itself
    MakeWvBox = CreateWindowExA(0, "Static", INLINE_TOKEN, &H40000000 Or &H2000000 Or &H4000000 Or &HD&, 0, 0, 16, 16, hHost, 0, 0, 0)
    If gotCtx Then SetThreadDpiAwarenessContext oldCtx
    Err.Clear
End Function

' Window for drawing a page after the show: top level, never shown, far outside the screen. It is made in the
' DPI context of the slide show window, like the container: views that share a browser must all sit in windows
' of the same DPI awareness (otherwise making a view fails with 8007139F).
Private Function MakeWvHold(ByVal hHost As LongPtr) As LongPtr
    Dim oldCtx As LongPtr, gotCtx As Boolean
    On Error Resume Next
    oldCtx = SetThreadDpiAwarenessContext(GetWindowDpiAwarenessContext(hHost))
    gotCtx = (Err.Number = 0 And oldCtx <> 0)
    Err.Clear
    ' tool window (not in the task bar), never activated, pop-up style, not visible
    MakeWvHold = CreateWindowExA(&H80& Or &H8000000, "Static", "", &H80000000, -30000, -30000, 64, 64, 0, 0, 0, 0)
    If gotCtx Then SetThreadDpiAwarenessContext oldCtx
    Err.Clear
End Function

Private Sub DropViews()
    Dim i As Long
    On Error Resume Next
    For i = 1 To mCount
        WvDrop i
        mWvState(i) = 0
    Next i
End Sub

' The modern engine cannot be used after all: everything of it goes, the classic engine takes over.
Private Function LeaveModern() As Boolean
    On Error Resume Next
    DropViews
    If mBoxWv <> 0 Then
        If IsWindow(mBoxWv) <> 0 Then DestroyWindow mBoxWv
    End If
    mBoxWv = 0
    If mHoldWv <> 0 Then
        If IsWindow(mHoldWv) <> 0 Then DestroyWindow mHoldWv
    End If
    mHoldWv = 0
    WvStop
    mUp = False
    mEngine = 2
    mCur = 0
    mWvShown = False
    LogIt "engine classic: " & WvWhy
    LeaveModern = False
End Function

' Shows page n over the rectangle x, y, w, h of the slide show window. False = use the classic engine instead.
Private Function ShowModern(ByVal hTop As LongPtr, ByVal hHost As LongPtr, ByVal n As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long) As Boolean
    Dim i As Long, moved As Boolean
    ShowModern = True
    If Not mUp Then
        ' another open deck with this code may already serve this show
        If OtherHostPresent(hHost) Then Exit Function
        WvSetup mCount
        ReDim mWvState(1 To mCount)
        ReDim mWvView(1 To mCount)
        ReDim mWvErr(1 To mCount)
        ReDim mWvSeq(1 To mCount)
        ReDim mWvUsed(1 To mCount)
        ReDim mWvFit(1 To mCount)
        ReDim mWvZoomF(1 To mCount)
        For i = 1 To mCount
            mTouched(i) = False
        Next i
        mWvRun = Format$(Now, "hhnnss")
        mWvFolder = ""
        mWvAny = False
        mWvShown = False
        mEndLater = False
        mAbort = False
        mCur = 0
        mUp = True
        LogIt "modern up"
    End If
    If mBoxWv <> 0 Then
        If IsWindow(mBoxWv) = 0 Then
            mBoxWv = 0
        ElseIf GetParent(mBoxWv) <> hHost Then
            DestroyWindow mBoxWv
            mBoxWv = 0
        End If
        ' (the views went with their parent window)
        If mBoxWv = 0 Then DropViews
    End If
    If mBoxWv = 0 Then
        mBoxWv = MakeWvBox(hHost)
        If mBoxWv = 0 Then
            LogIt "no-box"
            ShowModern = LeaveModern()
            Exit Function
        End If
        mHost = hHost
        mWvW = -1
        mWvShown = False
        mCur = 0
        LogIt "box " & CStr(mBoxWv)
        If mHoldWv = 0 Then mHoldWv = MakeWvHold(hHost)
    End If
    mTop = hTop
    If x <> mWvX Or y <> mWvY Or w <> mWvW Or h <> mWvH Then
        mWvX = x: mWvY = y: mWvW = w: mWvH = h
        ' 4 = keep the z order, 16 = do not activate
        SetWindowPos mBoxWv, 0, x, y, w, h, &H4& Or &H10&
        moved = True
    End If
    If mWvState(n) = 0 Then
        StartView n
        If mAbort Then Exit Function
        If mEngine <> 1 Then
            ShowModern = False
            Exit Function
        End If
    End If
    If mWvState(n) = 2 Then
        If moved Or n <> mCur Then FitView n, mBoxWv
        If n <> mCur Then
            WvVisible n, True
            If mCur >= 1 And mCur <= mCount Then
                WvPost mCur, "hide"
                WvVisible mCur, False
            End If
            mCur = n
            WvPost n, "show"
            LogIt "page " & n & " @" & x & "," & y & " " & w & "x" & h
        End If
        mWvClock = mWvClock + 1
        mWvUsed(n) = mWvClock
        If Not mWvShown Then
            ' 8 = show without activating
            ShowWindow mBoxWv, 8
            mWvShown = True
        End If
        If moved Then WvMoved n
        Preload
    Else
        ' not loaded yet (the page says "ready" when it is) or failed: the still picture of the slide stays
        HideModern
    End If
End Function

Private Sub HideModern()
    On Error Resume Next
    If mWvShown Then
        If mCur >= 1 And mCur <= mCount Then
            WvPost mCur, "hide"
            WvVisible mCur, False
        End If
        If mBoxWv <> 0 Then ShowWindow mBoxWv, 0
        mWvShown = False
        LogIt "hide"
    End If
    mCur = 0
End Sub

' Makes the view of page n and loads the page into it (the view stays invisible until the page is shown).
Private Sub StartView(ByVal n As Long)
    Dim p As String, nm As String
    If Len(mHtmlM(n)) = 0 Then
        mWvState(n) = -1
        Exit Sub
    End If
    EvictViews n
    If Not WvCreate(n, mBoxWv, mBg(n)) Then
        LogIt "wv-view-fail " & n & " " & WvWhy
        mWvState(n) = -1
        If Not mWvAny Then LeaveModern
        Exit Sub
    End If
    mWvAny = True
    If mAbort Then Exit Sub
    mWvSeq(n) = mWvSeq(n) + 1
    nm = "p" & n & "_" & mWvRun & "_" & mWvSeq(n) & ".html"
    p = WriteModernFile(nm, PageDoc(n, True))
    If Len(p) = 0 Then
        LogIt "wv-file-fail " & n
        mWvState(n) = -1
        Exit Sub
    End If
    FitView n, mBoxWv
    mWvState(n) = 1
    mWvView(n) = ""
    mWvErr(n) = ""
    If WvMapFolder(n, WV_HOST, mWvFolder) Then
        WvGo n, "https://" & WV_HOST & "/" & nm
    Else
        WvGo n, FileUrl(p)
    End If
    LogIt "wv-go " & n
End Sub

' Size and zoom of the view of page n: the page gets a viewport of exactly its canvas width in CSS px, whatever
' the size of the slide show on this screen. (The browser measures its window in whole device-independent
' pixels, rounded up; the zoom is taken from that number.) hDpi: window whose DPI applies.
Private Sub FitView(ByVal n As Long, ByVal hDpi As LongPtr)
    Dim dpi As Long, dip As Double
    On Error Resume Next
    If mWvW < 8 Or mWvH < 8 Then Exit Sub
    WvSize n, mWvW, mWvH
    mWvFit(n) = 0
    If mCanvasW(n) > 0 Then
        dpi = GetDpiForWindow(hDpi)
        If Err.Number <> 0 Or dpi < 48 Then dpi = 96
        Err.Clear
        dip = -Int(-(CDbl(mWvW) * 96# / dpi))
        mWvZoomF(n) = dip / mCanvasW(n)
        WvZoom n, mWvZoomF(n)
    End If
End Sub

' The page reported its viewport ("1281x720|1280.6"): remember it and correct the zoom when the width is off.
Private Sub CheckFit(ByVal n As Long, ByVal a As String)
    Dim k As Long, vw As Double
    On Error Resume Next
    k = InStr(1, a, "|", vbBinaryCompare)
    If k > 0 Then
        mWvView(n) = Left$(a, k - 1)
        vw = Val(Mid$(a, k + 1))
    Else
        mWvView(n) = a
        vw = Val(a)
    End If
    If mCanvasW(n) <= 0 Or vw < 8 Or mWvZoomF(n) <= 0 Then Exit Sub
    If Abs(vw - mCanvasW(n)) < 0.6 Then Exit Sub
    If mWvFit(n) >= 3 Then Exit Sub
    mWvFit(n) = mWvFit(n) + 1
    mWvZoomF(n) = mWvZoomF(n) * vw / mCanvasW(n)
    WvZoom n, mWvZoomF(n)
End Sub

' The page of the next slide is loaded ahead, so that changing slides is instant.
Private Sub Preload()
    Dim ssw As SlideShowWindow, idx As Long, total As Long, m As Long
    On Error Resume Next
    If mShowPres Is Nothing Then Exit Sub
    Set ssw = FindShow()
    If ssw Is Nothing Then Exit Sub
    idx = 0
    idx = ssw.View.Slide.SlideIndex
    total = 0
    total = mShowPres.Slides.Count
    If idx < 1 Or idx >= total Then Exit Sub
    m = 0
    m = CLng(Val(mShowPres.Slides(idx + 1).Tags(TAG_PAGE)))
    If m < 1 Or m > mCount Then Exit Sub
    If mWvState(m) = 0 Then StartView m
End Sub

' Not more than WV_MAXVIEWS views at a time (each is a browser process): the one that was on screen longest
' ago goes. What its page stored is in the presentation; the page is loaded again when its slide comes back.
Private Sub EvictViews(ByVal keep As Long)
    Dim i As Long, cnt As Long, best As Long, bestUsed As Long
    For i = 1 To mCount
        If mWvState(i) > 0 Then cnt = cnt + 1
    Next i
    If cnt < WV_MAXVIEWS Then Exit Sub
    bestUsed = &H7FFFFFFF
    For i = 1 To mCount
        If mWvState(i) > 0 And i <> mCur And i <> keep Then
            If mWvUsed(i) < bestUsed Then
                bestUsed = mWvUsed(i)
                best = i
            End If
        End If
    Next i
    If best > 0 Then
        WvDrop best
        mWvState(best) = 0
        LogIt "wv-evict " & best
    End If
End Sub

' Called by modWv2 when a page sends a message (window.chrome.webview.postMessage("command\nargument")).
' Commands: ready | vp (viewport) | err (script errors) | tick | key<code> | next | prev | goto<slide number> | save
Public Sub InlineWvMsg(ByVal page As Long, ByVal c As String, ByVal a As String)
    Dim t As Single, k As Long, ssw As SlideShowWindow
    On Error GoTo Fail
    If mEngine <> 1 Or Not mUp Then Exit Sub
    If page < 1 Or page > mCount Then Exit Sub
    If c = "tick" Then
        t = Timer
        If t >= mLastTick And t - mLastTick < 0.3 Then Exit Sub
        mLastTick = t
        If Not mAbort Then InlineSync "tick"
    ElseIf c = "ready" Then
        If mWvState(page) = 1 Then
            mWvState(page) = 2
            LogIt "wv-ready " & page
            If Not mAbort Then InlineSync "ready"
        End If
    ElseIf c = "vp" Then
        CheckFit page, a
    ElseIf c = "err" Then
        mWvErr(page) = Left$(Replace(Replace(Replace(Replace(a, "[", "("), "]", ")"), "{", "("), "}", ")"), 500)
    ElseIf page <> mCur Then
        ' only the page on screen may steer the show
    ElseIf Left$(c, 3) = "key" Then
        ForwardKey CLng(Val(Mid$(c, 4)))
    ElseIf c = "save" Then
        If mShowPres Is Nothing Then
            Set ssw = FindShow()
            If Not ssw Is Nothing Then Set mShowPres = ssw.Presentation
        End If
        If Not mShowPres Is Nothing Then
            WriteState mShowPres, page, a
            mTouched(page) = True
        End If
    ElseIf c = "next" Then
        ForwardKey 39
    ElseIf c = "prev" Then
        ForwardKey 37
    ElseIf Left$(c, 4) = "goto" Then
        k = CLng(Val(Mid$(c, 5)))
        Set ssw = FindShow()
        If Not ssw Is Nothing Then
            If k >= 1 And k <= ssw.Presentation.Slides.Count Then
                ReturnFocus
                ssw.View.GotoSlide k
            End If
        End If
    End If
    Exit Sub
Fail:
    LogIt "wvmsg-err " & Err.Number & " " & Err.Description
End Sub

' Writes a page file of this show into its own temp folder (%TEMP%\ppthtml_<process>_m), UTF-8.
Private Function WriteModernFile(ByVal nm As String, ByVal html As String) As String
    Dim b() As Byte, f As Integer, p As String
    On Error GoTo Fail
    If Len(html) = 0 Then Exit Function
    If Len(mWvFolder) = 0 Then
        p = Environ$("TEMP")
        If Len(p) = 0 Then p = Environ$("TMP")
        If Len(p) = 0 Then p = CurDir$
        If Right$(p, 1) <> "\" Then p = p & "\"
        p = p & "ppthtml_" & CStr(GetCurrentProcessId()) & "_m"
        If Len(Dir$(p, vbDirectory)) = 0 Then MkDir p
        mWvFolder = p
    End If
    p = mWvFolder & "\" & nm
    If Not Utf8Bytes(html, b) Then Exit Function
    If Len(Dir$(p)) > 0 Then Kill p
    f = FreeFile
    Open p For Binary Access Write As #f
    Put #f, , b
    Close #f
    WriteModernFile = p
    Exit Function
Fail:
    On Error Resume Next
    Close #f
    WriteModernFile = ""
End Function

Private Sub CleanModernFiles()
    On Error Resume Next
    If Len(mWvFolder) = 0 Then Exit Sub
    Kill mWvFolder & "\*.*"
    RmDir mWvFolder
    mWvFolder = ""
End Sub

' file:///... address of a path (used when the folder cannot be served under its own host name).
Private Function FileUrl(ByVal p As String) As String
    Dim buf As String, n As Long
    On Error Resume Next
    buf = String$(4200, vbNullChar)
    n = 4190
    If UrlCreateFromPathW(StrPtr(p), StrPtr(buf), n, 0) = 0 And n > 0 Then
        FileUrl = Left$(buf, n)
    Else
        FileUrl = "file:///" & Replace(p, "\", "/")
    End If
End Function

' The text of a JSON value as the browser returns it ("abc" with escapes -> abc; numbers and true / false as they are).
Private Function JsonText(ByVal s As String) As String
    Dim i As Long, c As String, r As String, n As Long
    On Error Resume Next
    If Len(s) < 2 Or Left$(s, 1) <> """" Then
        If s = "null" Or s = "undefined" Then Exit Function
        JsonText = s
        Exit Function
    End If
    n = Len(s) - 1
    i = 2
    Do While i <= n
        c = Mid$(s, i, 1)
        If c = "\" And i < n Then
            i = i + 1
            c = Mid$(s, i, 1)
            Select Case c
                Case "n": r = r & vbLf
                Case "r": r = r & vbCr
                Case "t": r = r & vbTab
                Case "b", "f"
                Case "u"
                    r = r & ChrW$(CLng("&H" & Mid$(s, i + 1, 4)))
                    i = i + 4
                Case Else: r = r & c
            End Select
        Else
            r = r & c
        End If
        i = i + 1
    Loop
    JsonText = r
End Function

Private Function ModernDiag() As String
    Dim r As String, i As Long, st As String
    On Error Resume Next
    r = "engine=modern box=" & CStr(mBoxWv) & " mode=4 shown=" & mWvShown & " cur=" & mCur & " place=" & mWvX & "," & mWvY & " " & mWvW & "x" & mWvH
    For i = 1 To mCount
        Select Case mWvState(i)
            Case 2: st = "alive,hooked"
            Case 1: st = "loading"
            Case -1: st = "failed"
            Case Else: st = "-"
        End Select
        r = r & " p" & i & "[" & st & " view=" & mWvView(i) & " mode=0 zoom=" & Format$(mWvZoomF(i), "0.0000")
        If Len(mWvErr(i)) > 0 Then r = r & " ERR{" & mWvErr(i) & "}"
        r = r & "]"
    Next i
    ModernDiag = r
End Function

' The slide show is over (modern engine): stills of edited pages, then everything of the engine goes.
Private Sub EndModern(ByVal why As String)
    On Error Resume Next
    LogIt "end(" & why & ") modern busy=" & mBusy
    mAbort = True
    HideModern
    mHost = 0
    mTop = 0
    If mBusy Then
        ' called from inside a running sync (a page message or a wait loop): finish when that is over
        mEndLater = True
        Exit Sub
    End If
    mEndLater = False
    mBusy = True
    RefreshStillsModern
    DropViews
    If mBoxWv <> 0 Then
        If IsWindow(mBoxWv) <> 0 Then DestroyWindow mBoxWv
    End If
    mBoxWv = 0
    If mHoldWv <> 0 Then
        If IsWindow(mHoldWv) <> 0 Then DestroyWindow mHoldWv
    End If
    mHoldWv = 0
    WvStop
    PumpFor 0.3
    CleanModernFiles
    mUp = False
    mStale = False
    mEngine = 0
    mBusy = False
    Set mShowPres = Nothing
    LogIt "modern down"
    If Not FindShow() Is Nothing Then InlineSync "after-end"
End Sub

' The still pictures follow the edits (modern engine). A page that stored something during the show is loaded
' once more, with what it stored, in a view inside a window that is never shown; the browser's own picture of
' it replaces the still on its slide.
Private Sub RefreshStillsModern()
    Dim i As Long, pres As Presentation, some As Boolean, hPop As LongPtr, nm As String, p As String, f As String
    Dim t0 As Single, el As Single, w As Single
    On Error Resume Next
    Set pres = mShowPres
    If pres Is Nothing Then Exit Sub
    For i = 1 To mCount
        If mTouched(i) Then some = True
    Next i
    If Not some Then Exit Sub
    If mWvW < 16 Or mWvH < 16 Then Exit Sub
    hPop = mHoldWv
    If hPop <> 0 Then
        If IsWindow(hPop) = 0 Then hPop = 0
    End If
    If hPop = 0 Then
        LogIt "still-skip no window"
        Exit Sub
    End If
    SetWindowPos hPop, 0, -30000, -30000, mWvW, mWvH, &H4& Or &H10&
    For i = 1 To mCount
        If mTouched(i) Then
            mTouched(i) = False
            WvDrop i
            mWvState(i) = 0
            If WvCreate(i, hPop, mBg(i)) Then
                mWvSeq(i) = mWvSeq(i) + 1
                nm = "p" & i & "_" & mWvRun & "_" & mWvSeq(i) & ".html"
                p = WriteModernFile(nm, PageDoc(i, True))
                FitView i, hPop
                WvVisible i, True
                mWvState(i) = 1
                If WvMapFolder(i, WV_HOST, mWvFolder) Then
                    WvGo i, "https://" & WV_HOST & "/" & nm
                Else
                    WvGo i, FileUrl(p)
                End If
                t0 = Timer
                Do
                    DoEvents
                    Sleep 5
                    el = Timer - t0
                    If el < 0 Then el = el + 86400!
                Loop While mWvState(i) = 1 And el < 10
                If mWvState(i) = 2 Then
                    WvPost i, "show"
                    ' the time the build gave the page before it took the still picture
                    w = mSettle(i) / 1000!
                    If w < 0.9 Then w = 0.9
                    If w > 5 Then w = 5
                    PumpFor w
                    f = mWvFolder & "\still" & CStr(i) & ".png"
                    If WvSnap(i, f) Then
                        ReplaceStill pres, i, f, "modern " & mWvView(i)
                    Else
                        LogIt "still-fail " & i
                    End If
                    If Len(Dir$(f)) > 0 Then Kill f
                Else
                    LogIt "still-timeout " & i
                End If
            Else
                LogIt "still-noview " & i & " " & WvWhy
            End If
            WvDrop i
            mWvState(i) = 0
        End If
    Next i
End Sub
