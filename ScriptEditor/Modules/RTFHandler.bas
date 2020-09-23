Attribute VB_Name = "RTFHandler"
Option Explicit
'// Win32 API Declarations
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessageRef Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, wParam As Any, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageByNum Lib "user32" _
    Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

'// Window Messages
Private Const WM_CLEAR = &H303
Private Const WM_COPY = &H301
Private Const WM_CUT = &H300
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_PASTE = &H302
Private Const WM_USER = &H400
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115

'// Edit Messages
Private Const EM_AUTOURLDETECT = (WM_USER + 91)
Private Const EM_CANPASTE = (WM_USER + 50)
Private Const EM_CANREDO = (WM_USER + 85)
Private Const EM_CANUNDO = &HC6
Private Const EM_CHARFROMPOS = &HD7
Private Const EM_CONVPOSITION = (WM_USER + 108)
Private Const EM_DISPLAYBAND = (WM_USER + 51)
Private Const EM_EMPTYUNDOBUFFER = &HCD
Private Const EM_EXGETSEL = (WM_USER + 52)
Private Const EM_EXLIMITTEXT = (WM_USER + 53)
Private Const EM_EXLINEFROMCHAR = (WM_USER + 54)
Private Const EM_EXSETSEL = (WM_USER + 55)
Private Const EM_FINDTEXT = (WM_USER + 56)
Private Const EM_FINDTEXTEX = (WM_USER + 79)
Private Const EM_FINDTEXTEXW = (WM_USER + 124)
Private Const EM_FINDTEXTW = (WM_USER + 123)
Private Const EM_FINDWORDBREAK = (WM_USER + 76)
Private Const EM_FMTLINES = &HC8
Private Const EM_FORMATRANGE = (WM_USER + 57)
Private Const EM_GETAUTOURLDETECT = (WM_USER + 92)
Private Const EM_GETBIDIOPTIONS = (WM_USER + 201)
Private Const EM_GETCHARFORMAT = (WM_USER + 58)
Private Const EM_GETEDITSTYLE = (WM_USER + 205)
Private Const EM_GETEVENTMASK = (WM_USER + 59)
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_GETHANDLE = &HBD
Private Const EM_GETIMECOLOR = (WM_USER + 105)
Private Const EM_GETIMECOMPMODE = (WM_USER + 122)
Private Const EM_GETIMEMODEBIAS = (WM_USER + 127)
Private Const EM_GETIMEOPTIONS = (WM_USER + 107)
Private Const EM_GETIMESTATUS = &HD9
Private Const EM_GETLANGOPTIONS = (WM_USER + 121)
Private Const EM_GETLIMITTEXT = (WM_USER + 37)
Private Const EM_GETLINE = &HC4
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_GETMARGINS = &HD4
Private Const EM_GETMODIFY = &HB8
Private Const EM_GETOLEINTERFACE = (WM_USER + 60)
Private Const EM_GETOPTIONS = (WM_USER + 78)
Private Const EM_GETPARAFORMAT = (WM_USER + 61)
Private Const EM_GETPASSWORDCHAR = &HD2
Private Const EM_GETPUNCTUATION = (WM_USER + 101)
Private Const EM_GETRECT = &HB2
Private Const EM_GETREDONAME = (WM_USER + 87)
Private Const EM_GETSCROLLPOS = (WM_USER + 221)
Private Const EM_GETSEL = &HB0
Private Const EM_GETSELTEXT = (WM_USER + 62)
Private Const EM_GETTEXTEX = (WM_USER + 94)
Private Const EM_GETTEXTLENGTHEX = (WM_USER + 95)
Private Const EM_GETTEXTMODE = (WM_USER + 90)
Private Const EM_GETTEXTRANGE = (WM_USER + 75)
Private Const EM_GETTHUMB = &HBE
Private Const EM_GETTYPOGRAPHYOPTIONS = (WM_USER + 203)
Private Const EM_GETUNDONAME = (WM_USER + 86)
Private Const EM_GETWORDBREAKPROC = &HD1
Private Const EM_GETWORDBREAKPROCEX = (WM_USER + 80)
Private Const EM_GETWORDWRAPMODE = (WM_USER + 103)
Private Const EM_GETZOOM = (WM_USER + 224)
Private Const EM_HIDESELECTION = (WM_USER + 63)
Private Const EM_LIMITTEXT = &HC5
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_LINESCROLL = &HB6
Private Const EM_OUTLINE = (WM_USER + 220)
Private Const EM_PASTESPECIAL = (WM_USER + 64)
Private Const EM_POSFROMCHAR = (WM_USER + 38)
Private Const EM_RECONVERSION = (WM_USER + 125)
Private Const EM_REDO = (WM_USER + 84)
Private Const EM_REPLACESEL = &HC2
Private Const EM_REQUESTRESIZE = (WM_USER + 65)
Private Const EM_SCROLL = &HB5
Private Const EM_SCROLLCARET = &HB7
Private Const EM_SELECTIONTYPE = (WM_USER + 66)
Private Const EM_SETBIDIOPTIONS = (WM_USER + 200)
Private Const EM_SETBKGNDCOLOR = (WM_USER + 67)
Private Const EM_SETCHARFORMAT = (WM_USER + 68)
Private Const EM_SETEDITSTYLE = (WM_USER + 204)
Private Const EM_SETEVENTMASK = (WM_USER + 69)
Private Const EM_SETFONTSIZE = (WM_USER + 223)
Private Const EM_SETHANDLE = &HBC
Private Const EM_SETIMECOLOR = (WM_USER + 104)
Private Const EM_SETIMEMODEBIAS = (WM_USER + 126)
Private Const EM_SETIMEOPTIONS = (WM_USER + 106)
Private Const EM_SETIMESTATUS = &HD8
Private Const EM_SETLANGOPTIONS = (WM_USER + 120)
Private Const EM_SETLIMITTEXT = EM_LIMITTEXT
Private Const EM_SETMARGINS = &HD3
Private Const EM_SETMODIFY = &HB9
Private Const EM_SETOLECALLBACK = (WM_USER + 70)
Private Const EM_SETOPTIONS = (WM_USER + 77)
Private Const EM_SETPALETTE = (WM_USER + 93)
Private Const EM_SETPARAFORMAT = (WM_USER + 71)
Private Const EM_SETPASSWORDCHAR = &HCC
Private Const EM_SETPUNCTUATION = (WM_USER + 100)
Private Const EM_SETREADONLY = &HCF
Private Const EM_SETRECT = &HB3
Private Const EM_SETRECTNP = &HB4
Private Const EM_SETSCROLLPOS = (WM_USER + 222)
Private Const EM_SETSEL = &HB1
Private Const EM_SETTABSTOPS = &HCB
Private Const EM_SETTARGETDEVICE = (WM_USER + 72)
Private Const EM_SETTEXTEX = (WM_USER + 97)
Private Const EM_SETTEXTMODE = (WM_USER + 89)
Private Const EM_SETTYPOGRAPHYOPTIONS = (WM_USER + 202)
Private Const EM_SETUNDOLIMIT = (WM_USER + 82)
Private Const EM_SETWORDBREAKPROC = &HD0
Private Const EM_SETWORDBREAKPROCEX = (WM_USER + 81)
Private Const EM_SETWORDWRAPMODE = (WM_USER + 102)
Private Const EM_SETZOOM = (WM_USER + 225)
Private Const EM_SHOWSCROLLBAR = (WM_USER + 96)
Private Const EM_STOPGROUPTYPING = (WM_USER + 88)
Private Const EM_STREAMIN = (WM_USER + 73)
Private Const EM_STREAMOUT = (WM_USER + 74)
Private Const EM_UNDO = &HC7
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetCaretPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Type POINTAPI
    X As Long
    Y As Long
End Type


Public Function GetCurrentpos(ByVal pParentHwnd As Long, rtfText As RichTextBox) As POINTAPI
    Dim pPos As POINTAPI
    Dim lRetVal As Long
    lRetVal = SendMessageRef(rtfText.hwnd, EM_POSFROMCHAR, pPos, 0&)
    'Call GetCursorPos(pPos)
    'If WindowFromPoint(pPos.x, pPos.y) = rtfText.hwnd Then
        Call ClientToScreen(rtfText.hwnd, pPos)
        Call ScreenToClient(pParentHwnd, pPos)
        GetCurrentpos = pPos
        
    'Else
    '    pPos.x = 0
    '    pPos.y = 0
    'End If
End Function


Public Function hiByte(ByVal w As Integer) As Byte
    Dim hi As Integer
    If w And &H8000 Then hi = &H4000
    
    hiByte = (w And &H7FFE) \ 256
    hiByte = (hiByte Or (hi \ 128))
    
End Function

Public Function HiWord(dw As Long) As Integer
 If dw And &H80000000 Then
      HiWord = (dw \ 65535) - 1
 Else
    HiWord = dw \ 65535
 End If
End Function

Public Function LoByte(w As Integer) As Byte
 LoByte = w And &HFF
End Function

Public Function LoWord(dw As Long) As Integer
  If dw And &H8000& Then
      LoWord = &H8000 Or (dw And &H7FFF&)
   Else
      LoWord = dw And &HFFFF&
   End If
End Function

Public Function MakeInt(ByVal LoByte As Byte, _
   ByVal hiByte As Byte) As Integer

MakeInt = ((hiByte * &H100) + LoByte)

End Function

Public Function MakeLong(ByVal LoWord As Integer, _
  ByVal HiWord As Integer) As Long

MakeLong = ((HiWord * &H10000) + LoWord)

End Function

Public Function GetCurrentColumn(rtfText As RichTextBox) As Long
    Dim pPos As POINTAPI
    Dim pRet As Long
    pRet = GetCaretPos(pPos)
    GetCurrentColumn = (pPos.X - 1) / 8 + 1
    'GetCurrentColumn = rtfText.SelStart / GetCurrentLine(rtfText)
    'GetCurrentColumn = rtfText.SelStart - SendMessageByNum(rtfText.hwnd, _
EM_LINEINDEX, -1&, 0&) + 1
    
End Function

Public Function GetCurrentLine(rtfText As RichTextBox) As Long
    '// Get current line
    GetCurrentLine = SendMessage(rtfText.hwnd, EM_LINEFROMCHAR, -1, 0&) + 1
End Function


Public Function GetFirstVisibleLine(rtfText As RichTextBox) As Long
    GetFirstVisibleLine = SendMessage(rtfText.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
End Function


Public Function GetNormalisedColumn(pEditorCtrl As Editor, rtfText As RichTextBox) As Long
    Dim pPos As POINTAPI
    Dim pRet As Long
    pRet = GetCaretPos(pPos)
    GetNormalisedColumn = (pPos.X - 1) / 8 + 1
End Function

Public Function GetNormalisedLine(pEditorCtrl As Editor, rtfText As RichTextBox) As Long
    GetNormalisedLine = GetCurrentLine(rtfText) - GetFirstVisibleLine(rtfText)
End Function


Public Sub GoToLineNumber(rtfText As RichTextBox, ByVal lLine As Long)
    Dim lRetVal As Long
    '// Go to specified line number
    lRetVal = SendMessage(rtfText.hwnd, EM_LINEINDEX, lLine - 1, 0&)
    If lRetVal = -1 Then '// If invalid number then
        MsgBox "Invalid line number!", vbCritical '// show error
        Exit Sub '// and Exit Sub
    End If
    rtfText.SelStart = lRetVal '// Go to selected line
End Sub

Public Sub SelectLine(rtfText As RichTextBox)
    GetSelectedLine rtfText '// Select current line
End Sub
Public Function GetSelectedLine(rtfText As RichTextBox) As String
    Dim lStart As Long
    Dim lLen As Long
    Dim lSStart As Long
    Dim lSEnd As Long
    Dim intCurrLine As Integer
    Dim intStartPos As Integer
    Dim intEndLen As Integer
    Dim intSelLen As Integer
    Dim strSearch As String
    
    '// Save selected start and length
    lStart = rtfText.SelStart
    lLen = rtfText.SelLength
    '// Get current line
    intCurrLine = SendMessage(rtfText.hwnd, EM_LINEFROMCHAR, lStart, 0&)
    '// Set the start pos at the beginning of the line
    lSStart = SendMessage(rtfText.hwnd, EM_LINEINDEX, intCurrLine, 0&)
    If Err Then
        '// Line does not exist
        GetSelectedLine = ""
        Exit Function
    End If
    '// Get the length of the line
    intSelLen = SendMessage(rtfText.hwnd, EM_LINELENGTH, lStart, 0&)
    '// Select the line
    rtfText.SelStart = lSStart
    rtfText.SelLength = intSelLen
    '// Return selected text
    GetSelectedLine = rtfText.SelText
End Function


