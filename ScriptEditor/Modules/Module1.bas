Attribute VB_Name = "modGlobal"
Option Explicit

Public Const SM_BUILTINOBJECTS = " SMDEBUG SMEVENT PARENTFORM "
Public Const SM_RESERVEDWORDS = " Declare = And ByRef ByVal Call Case Class Const Dim Do Each Else ElseIf Empty End Eqv Erase Error Exit Explicit False For Function Get If Imp In Is Let Loop Mod Next Not Nothing Null On Option Or Private Property Public Randomize ReDim Resume Select Set Step Sub Then To True Until Wend While Xor As Long Integer Boolean New With "
'Public Const SM_FUNCTION = " Anchor Array Asc Atn CBool CByte CCur CDate CDbl Chr CInt CLng Cos CreateObject CSng CStr Date DateAdd DateDiff DatePart DateSerial DateValue Day Dictionary Document Element Err Exp FileSystemObject  Filter Fix Int Form FormatCurrency FormatDateTime FormatNumber FormatPercent GetObject Hex History Hour InputBox InStr InstrRev IsArray IsDate IsEmpty IsNull IsNumeric IsObject Join LBound LCase Left Len Link LoadPicture Location Log LTrim RTrim Trim Mid Minute Month MonthName MsgBox Navigator Now Oct Replace Right Rnd Round ScriptEngine ScriptEngineBuildVersion ScriptEngineMajorVersion ScriptEngineMinorVersion Second Sgn Sin Space Split Sqr StrComp String StrReverse Tan Time TextStream TimeSerial TimeValue TypeName UBound UCase VarType Weekday WeekDayName Window Year "
'Public Const SM_FUNCTIONWORDS = " Anchor Array Asc Atn CBool CByte CCur CDate CDbl Chr CInt CLng Cos CreateObject CSng CStr Date DateAdd DateDiff DatePart DateSerial DateValue Day Dictionary Document Element Err Exp FileSystemObject  Filter Fix Int Form FormatCurrency FormatDateTime FormatNumber FormatPercent GetObject Hex History Hour InputBox InStr InstrRev IsArray IsDate IsEmpty IsNull IsNumeric IsObject Join LBound LCase Left Len Link LoadPicture Location Log LTrim RTrim Trim Mid Minute Month MonthName MsgBox Navigator Now Oct Replace Right Rnd Round ScriptEngine ScriptEngineBuildVersion ScriptEngineMajorVersion ScriptEngineMinorVersion Second Sgn Sin Space Split Sqr StrComp String StrReverse Tan Time TextStream TimeSerial TimeValue TypeName UBound UCase VarType Weekday WeekDayName Window Year "
Public SM_FUNCTIONWORDS As String
Public Const CALPHA = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public Const CNUMBERS = "0123456789"
Private Const LB_GETHORIZONTALEXTENT = &H193
Private Const LB_SETHORIZONTALEXTENT = &H194
Private Const LB_SETTABSTOPS = &H192
Public g_Project As Project


Public Property Get FileExists(ByVal sFile As String) As Boolean
   On Error GoTo Error
   If FileLen(sFile) > 0 Then
        FileExists = True
   Else
        FileExists = True
   End If
   sFile = Dir(sFile)
Error:
   FileExists = ((Err.Number = 0) And sFile <> "")
   On Error GoTo 0
End Property

