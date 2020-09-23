VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIntellisense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

Private m_TLIApplication As TLI.TLIApplication
Private m_CreatableClasses As Collection
Private m_AllClasses As Collection
Private m_VBAFunctions As Collection
Public Unload As Boolean
Public IsBusyLoading As Boolean


Public Enum OperatorType
    otUnknown = 0 '00
    otCreateObject = 1 '01
    otDimAs = 2 '10
    otDot = 4 '100
    otEqualTo = 8 '1000
    otShowTip = 16
End Enum

Public Function GetRefrenceFromObject(ByVal pObject As Object) As Refrence
    Dim pRefrence As New Refrence
    Dim pTypeLibInfo As TLI.TypeLibInfo
    Dim pInterfaceInfo As TLI.InterfaceInfo
    Set pInterfaceInfo = m_TLIApplication.InterfaceInfoFromObject(pObject)
    Set pTypeLibInfo = pInterfaceInfo.Parent
    With pRefrence
        .BinaryPath = pTypeLibInfo.ContainingFile
        .GUID = pTypeLibInfo.GUID
        .Name = pTypeLibInfo.HelpString
        .Verion = pTypeLibInfo.MajorVersion & "." & pTypeLibInfo.MinorVersion
    End With
    
    Set GetRefrenceFromObject = pRefrence
    Set pTypeLibInfo = Nothing
    Set pInterfaceInfo = Nothing
    Set pRefrence = Nothing
End Function


Public Property Get CreatableClasses() As Collection
    If m_CreatableClasses Is Nothing Then
        Set m_CreatableClasses = New Collection
    End If
    Set CreatableClasses = m_CreatableClasses
End Property

Public Property Get AllClasses() As Collection
    If m_AllClasses Is Nothing Then
         Set m_AllClasses = New Collection
    End If
    Set AllClasses = m_AllClasses
End Property

Public Property Get VBAFunctions() As Collection
    If m_VBAFunctions Is Nothing Then
        Set m_VBAFunctions = New Collection
    End If
    Set VBAFunctions = m_VBAFunctions
End Property

Public Sub AddRefToIntellsense(ByVal pRefrence As Refrence)
    Dim pTypeLibInfo As New cTypeLibInfo
    With pTypeLibInfo
        .CLSID = pRefrence.GUID
        .Exists = True
        .Name = pRefrence.Name
        .Ver = pRefrence.Verion
        .Path = pRefrence.BinaryPath
    End With
    Call GetProgID(pTypeLibInfo)
End Sub

Public Function GetVBARefrence() As Refrence
    Dim pRefrence As Refrence
    Dim pTypeLibInfo As TLI.TypeLibInfo
    Dim pTypeInfo As TLI.TypeInfo
    
    Set pTypeInfo = m_TLIApplication.ClassInfoFromObject(New Collection)
    Set pTypeLibInfo = m_TLIApplication.TypeLibInfoFromFile(pTypeInfo.Parent.ContainingFile)
    
    Set pRefrence = New Refrence
    With pRefrence
        .BinaryPath = pTypeLibInfo.ContainingFile
        .GUID = pTypeLibInfo.GUID
        .Name = pTypeLibInfo.HelpString
        .Verion = pTypeLibInfo.MajorVersion & "." & pTypeLibInfo.MinorVersion
    End With
    Set GetVBARefrence = pRefrence
    Set pTypeLibInfo = Nothing
    Set pTypeInfo = Nothing
End Function



Private Sub GetProgID(pExTypeLib As cTypeLibInfo)
    Dim pTypeLibInfo As TLI.TypeLibInfo
    Dim pExTypeLibInfo As cTypeLibInfo
    Dim pData As String
    Dim pTypeInfo As TLI.TypeInfo
    Dim pConstantInfo As ConstantInfo
    Dim pDeclarationInfo As DeclarationInfo
    Dim pMemberInfo As MemberInfo
    Dim pIsVBALib As Boolean
    
    
   On Error Resume Next
    
    Set pTypeLibInfo = m_TLIApplication.TypeLibInfoFromFile(pExTypeLib.Path)
    If Not pTypeLibInfo Is Nothing Then
        Set pExTypeLibInfo = New cTypeLibInfo
        pData = pTypeLibInfo.Name
        
        With pExTypeLibInfo
            .ProgID = pData
            .ClassName = pData
            .TypeKind = iLibrary
            .CLSID = pExTypeLib.CLSID
            .Exists = pExTypeLib.Exists
            .Name = pExTypeLib.Name
            .Path = pExTypeLib.Path
            .Ver = pExTypeLib.Ver
        End With
        Me.AllClasses.Add pExTypeLibInfo, pExTypeLibInfo.ProgID
        Me.CreatableClasses.Add pExTypeLibInfo, pExTypeLibInfo.ProgID
        Set pExTypeLibInfo = Nothing
        If Err.Number = 457 Then    'This key is already associated with an element of this collection
            GoTo exitNormal
            
        End If
        pIsVBALib = False
        If UCase(pData) = "VBA" Then
            pIsVBALib = True
        End If
                        
        For Each pConstantInfo In pTypeLibInfo.Constants
            If pConstantInfo.TypeKind = TKIND_ENUM Then
                If Not Left(pConstantInfo.Name, 1) = "_" Then
                    Set pExTypeLibInfo = New cTypeLibInfo
                    pData = pTypeLibInfo.Name & "." & pConstantInfo.Name
                    With pExTypeLibInfo
                        .ProgID = pData
                        .ClassName = pData
                        .TypeKind = iEnum
                        .CLSID = pExTypeLib.CLSID
                        .Exists = pExTypeLib.Exists
                        .Name = pExTypeLib.Name
                        .Path = pExTypeLib.Path
                        .Ver = pExTypeLib.Ver
                    End With
                    Me.AllClasses.Add pExTypeLibInfo, pExTypeLibInfo.ProgID
                    Debug.Print pConstantInfo.Name
                End If
            End If
        Next
        
        If pIsVBALib Then
            For Each pDeclarationInfo In pTypeLibInfo.Declarations
                If Left(pDeclarationInfo.Name, 1) <> "_" Then
                    For Each pMemberInfo In pDeclarationInfo.Members
                        If Left(pMemberInfo.Name, 1) <> "_" Then
                            Set pExTypeLibInfo = New cTypeLibInfo
                             pData = pDeclarationInfo.Name & "." & pMemberInfo.Name
                             
                             With pExTypeLibInfo
                                 .ProgID = pData
                                 .ClassName = pMemberInfo.Name
                                 .CLSID = pExTypeLib.CLSID
                                 .Exists = pExTypeLib.Exists
                                 .Name = pExTypeLib.Name
                                 .Path = pExTypeLib.Path
                                 .Ver = pExTypeLib.Ver
                                 .Tag = GetParameterTip(pMemberInfo)
                             End With
                             
                             Select Case pMemberInfo.InvokeKind
                                Case InvokeKinds.INVOKE_CONST:
                                    pExTypeLibInfo.TypeKind = iConstant
                                Case InvokeKinds.INVOKE_EVENTFUNC:
                                    pExTypeLibInfo.TypeKind = iEvent
                                Case InvokeKinds.INVOKE_FUNC:
                                    pExTypeLibInfo.TypeKind = iFunction
                                Case InvokeKinds.INVOKE_PROPERTYGET:
                                    pExTypeLibInfo.TypeKind = iProperty
                                Case InvokeKinds.INVOKE_PROPERTYPUT:
                                    pExTypeLibInfo.TypeKind = iProperty
                                Case InvokeKinds.INVOKE_PROPERTYPUTREF:
                                    pExTypeLibInfo.TypeKind = iProperty
                                Case InvokeKinds.INVOKE_UNKNOWN:
                                    GoTo lblVBAContinue
                             End Select
                             If pExTypeLibInfo.TypeKind = iConstant Then
                                pExTypeLibInfo.Tag = "CONST:" & CLng(pMemberInfo.Value)
                             End If
                             Me.VBAFunctions.Add pExTypeLibInfo, pExTypeLibInfo.ProgID
                        End If
                    Next

                End If
lblVBAContinue:
            Next
        End If
        
        For Each pTypeInfo In pTypeLibInfo.TypeInfos
            If (CBool(CLng(pTypeInfo.AttributeMask And TYPEFLAG_FCANCREATE) = TYPEFLAG_FCANCREATE) Or _
            CBool(CLng(pTypeInfo.AttributeMask And TYPEFLAG_FLICENSED) = TYPEFLAG_FLICENSED) Or _
            CBool(CLng(pTypeInfo.AttributeMask And TYPEFLAG_FOLEAUTOMATION) = TYPEFLAG_FOLEAUTOMATION) _
            Or CBool(CLng(pTypeInfo.AttributeMask And TYPEFLAG_FDISPATCHABLE) = TypeFlags.TYPEFLAG_FDISPATCHABLE) Or pTypeInfo.AttributeMask = 0) Then
            
                If pTypeInfo.AttributeMask <> 0 Then
                    If CBool(CLng(pTypeInfo.AttributeMask And TYPEFLAG_FHIDDEN) = TypeFlags.TYPEFLAG_FHIDDEN) Then
                        If Not CBool(CLng(pTypeInfo.AttributeMask And TYPEFLAG_FCANCREATE) = TypeFlags.TYPEFLAG_FCANCREATE) Then
                            GoTo lblContinue
                        End If
                    End If
                    
                    If Not CBool(CLng(pTypeInfo.AttributeMask And TYPEFLAG_FNONEXTENSIBLE) = TypeFlags.TYPEFLAG_FNONEXTENSIBLE) Then
                        If Not CBool(CLng(pTypeInfo.AttributeMask And TYPEFLAG_FCANCREATE) = TYPEFLAG_FCANCREATE) Then
                            GoTo lblContinue
                        End If
                    End If
                End If
                
                If Left(pTypeInfo.Name, 1) = "_" Then
                    GoTo lblContinue
                End If
                
                Set pExTypeLibInfo = New cTypeLibInfo
                
              
                
                pData = pTypeLibInfo.Name & "." & pTypeInfo.Name
                
                With pExTypeLibInfo
                    .ProgID = pData
                    .ClassName = pData
                    .CLSID = pExTypeLib.CLSID
                    .Exists = pExTypeLib.Exists
                    .Name = pExTypeLib.Name
                    .Path = pExTypeLib.Path
                    .Ver = pExTypeLib.Ver
                End With
                
                Select Case pTypeInfo.TypeKind
                    Case TypeKinds.TKIND_ALIAS:
                            pExTypeLibInfo.TypeKind = iDataType
                            'GoTo lblContinue
                    Case TypeKinds.TKIND_COCLASS:
                            pExTypeLibInfo.TypeKind = iClass
                    Case TypeKinds.TKIND_DISPATCH:
                            pExTypeLibInfo.TypeKind = iClass
                    Case TypeKinds.TKIND_ENUM:
                            pExTypeLibInfo.TypeKind = iEnum
                    Case TypeKinds.TKIND_INTERFACE:
                        GoTo lblContinue
                    Case TypeKinds.TKIND_MAX:
                        GoTo lblContinue
                    Case TypeKinds.TKIND_MODULE:
                        GoTo lblContinue
                    Case TypeKinds.TKIND_RECORD:
                        pExTypeLibInfo.TypeKind = iType
                    Case TypeKinds.TKIND_UNION:
                        GoTo lblContinue
                End Select
                
                Me.AllClasses.Add pExTypeLibInfo, pExTypeLibInfo.ProgID
                
                If (pTypeInfo.AttributeMask And TYPEFLAG_FCANCREATE) Then
                    ' the class is creatable or we don't care
                    ' about the creatable attribute
                   Me.CreatableClasses.Add pExTypeLibInfo, pExTypeLibInfo.ProgID
                End If
                
            End If
lblContinue:
            Set pExTypeLibInfo = Nothing
            DoEvents
        Next
    End If
exitNormal:
    Set pTypeLibInfo = Nothing
End Sub

Private Sub Class_Initialize()
     Set m_TLIApplication = New TLI.TLIApplication
     Unload = False
End Sub

Private Sub Class_Terminate()
    Set m_TLIApplication = Nothing
    Set m_AllClasses = Nothing
    Set m_VBAFunctions = Nothing
    Set m_CreatableClasses = Nothing
End Sub

Public Function ParseOperator(ByVal prtEditor As RichTextBox, ByVal pLine As String, ByVal pPartialText As String) As OperatorType
    Dim pText As String
    Dim pPos As Long
    Dim pOperator As OperatorType
    Dim pFormatedText As String
    
    pText = Replace(pLine, vbTab, " ")
    pText = Replace(pText, vbCr, "")
    pText = Replace(pText, vbLf, "")
    pFormatedText = Trim(pText)
    pOperator = otUnknown
    If Len(pText) > 13 Then
        If Mid(pText, Len(pText) - 13, 12) = "CreateObject" Then
            If InStr(1, pText, "Set ", vbTextCompare) > 0 Then
                pOperator = otCreateObject
            End If
        End If
    End If
    If InStr(1, pText, "'As ") > 0 Then
        If InStr(1, pText, "Dim ", vbTextCompare) > 0 Then
            pOperator = otDimAs
        End If
    End If
    
    'Parse for =
    pPos = InStr(1, StrReverse(pText), "=")
    
    If pPos > 0 And pOperator = otUnknown And InStr(1, pText, "Set ") <= 0 Then
        If Mid(Trim(pText), pPos + 1, 1) <> " " Then
            pPos = Len(pText) - pPos
            pText = Left(pText, pPos)
            pPos = InStr(1, StrReverse(pText), " ")
            If pPos > 0 Then
                pPos = Len(pText) - pPos
                pText = Trim(Mid(pText, pPos + 1))
            End If
            pPos = InStr(1, pText, ".")
            If pPos > 0 Then '
                pText = Left(pText, pPos - 1)
            End If
            If IsVariable(pText) Then
                pOperator = otEqualTo
            End If
        End If
    End If
    
    'If No = then parse for .
    pText = Replace(pLine, vbTab, " ")
    pText = Replace(pText, vbCr, "")
    pText = Replace(pText, vbLf, "")
    pText = Trim(pText)
    pPos = InStr(1, StrReverse(pText), ".")
    
    If pPos > 0 And pOperator = otUnknown Then
        If Mid(pText, pPos + 1, 1) <> " " Then
            pPos = Len(pText) - pPos
            pText = Left(pText, pPos)
            pPos = InStr(1, StrReverse(pText), " ")
            If pPos > 0 Then
                pPos = Len(pText) - pPos
                pText = Trim(Mid(pText, pPos + 1))
            End If
            pPos = InStr(1, pText, ".")
            If pPos > 0 Then '
                pText = Left(pText, pPos - 1)
            End If
            If IsVariable(pText) Then
                pPos = InStr(1, StrReverse(pFormatedText), ".")
                If pPos > 0 Then
                    pPos = Len(pFormatedText) - pPos
                    If Not (InStr(pPos + 1, pFormatedText, "(") > 0 Or InStr(pPos + 1, pFormatedText, ")") > 0) Then
                        pOperator = otDot
                    Else
                        pOperator = otShowTip
                    End If
                End If
                If Right(pFormatedText, 1) = "." Then
                    pOperator = otDot
                End If
            End If
                
            
        End If
    End If
    
    
    ParseOperator = pOperator
End Function

Private Function GetParameterTip(ByVal pMemberInfo As MemberInfo) As String
    Dim pParameterInfo As ParameterInfo
    Dim pReturnString As String
    Dim pVarType As String
    
    If pMemberInfo.Parameters.Count > 0 Then
        pReturnString = pMemberInfo.Name & "("
        For Each pParameterInfo In pMemberInfo.Parameters
            
            
            If Not pParameterInfo.VarTypeInfo.TypeInfo Is Nothing Then
                If pParameterInfo.VarTypeInfo.TypeInfo.TypeKind = TKIND_ENUM Then
                    'pExTypeKind = iEnum
                End If
                
                pVarType = " As " & pParameterInfo.VarTypeInfo.TypeInfo.Name
            Else
                Select Case pParameterInfo.VarTypeInfo.VarType
                    Case TliVarType.VT_ARRAY:
                        
                        Debug.Assert False
                    Case TliVarType.VT_BLOB:
                        Debug.Assert False
                    Case TliVarType.VT_BLOB_OBJECT:
                        Debug.Assert False
                    Case TliVarType.VT_BOOL:
                        pVarType = " As Boolean"
                    Case TliVarType.VT_BSTR:
                        pVarType = " As String"
                        'Debug.Assert False
                    Case TliVarType.VT_BYREF:
                        Debug.Assert False
                    Case TliVarType.VT_CARRAY:
                        Debug.Assert False
                    Case TliVarType.VT_CF:
                        Debug.Assert False
                    Case TliVarType.VT_CLSID:
                        Debug.Assert False
                    Case TliVarType.VT_CY:
                        pVarType = " As Currency"
                    Case TliVarType.VT_DATE:
                        pVarType = " As Date"
                    Case TliVarType.VT_DECIMAL:
                        pVarType = " As Double"
                    Case TliVarType.VT_DISPATCH:
                        pVarType = " As Object"
'                        Debug.Assert False
                    Case TliVarType.VT_EMPTY:
                        Debug.Assert False
                    Case TliVarType.VT_ERROR:
                        Debug.Assert False
                    Case TliVarType.VT_FILETIME:
                        Debug.Assert False
                    Case TliVarType.VT_HRESULT:
                        pVarType = " As Long"
                    Case TliVarType.VT_I1:
                        pVarType = " As Byte"
                    Case TliVarType.VT_I2:
                        pVarType = "As Boolean"
                    Case TliVarType.VT_I4:
                        pVarType = " As Long"
                        'Debug.Assert False
                    Case TliVarType.VT_I8:
                        pVarType = " As Long"
                    Case TliVarType.VT_INT:
                        pVarType = " As Long"
                    Case TliVarType.VT_LPSTR:
                        pVarType = " As Long"
                    Case TliVarType.VT_LPWSTR:
                        pVarType = " As Long"
                    Case TliVarType.VT_NULL:
                        Debug.Assert False
                    Case TliVarType.VT_PTR:
                        pVarType = " As Long"
                    Case TliVarType.VT_R4:
                        pVarType = " As Single"
                    Case TliVarType.VT_R8:
                        pVarType = " As Double"
                    Case TliVarType.VT_RECORD:
                        Debug.Assert False
                    Case TliVarType.VT_RESERVED:
                        Debug.Assert False
                    Case TliVarType.VT_SAFEARRAY:
                        Debug.Assert False
                    Case TliVarType.VT_STORAGE:
                        Debug.Assert False
                    Case TliVarType.VT_STORED_OBJECT:
                        pVarType = " As Object"
                    Case TliVarType.VT_STREAM:
                        Debug.Assert False
                    Case TliVarType.VT_STREAMED_OBJECT:
                        pVarType = " As Object"
                    Case TliVarType.VT_UI1:
                        pVarType = " As Byte"
                    Case TliVarType.VT_UI2:
                        pVarType = " As Single"
                    Case TliVarType.VT_UI4:
                        pVarType = " As Long"
                    Case TliVarType.VT_UI8:
                        pVarType = " As Long"
                    Case TliVarType.VT_UINT:
                        pVarType = " As Integer"
                    Case TliVarType.VT_UNKNOWN:
                        pVarType = ""
                    Case TliVarType.VT_USERDEFINED:
                        Debug.Assert False
                    Case TliVarType.VT_VARIANT:
                        pVarType = ""
                        'Debug.Assert False
                    Case TliVarType.VT_VECTOR:
                        
                        Debug.Assert False
                    Case TliVarType.VT_VOID:
                        pVarType = ""
                End Select
                'pExTypeKind = iProperty
            End If
            
            If CInt(pParameterInfo.Flags And PARAMFLAG_FOPT) = PARAMFLAG_FOPT Then
                If pParameterInfo.Default Then
                    If Len(pParameterInfo.DefaultValue) > 0 Then
                        pReturnString = pReturnString & "[" & pParameterInfo.Name & pVarType & " = " & pParameterInfo.DefaultValue & "],"
                    Else
                        pReturnString = pReturnString & "[" & pParameterInfo.Name & pVarType & "],"
                    End If
                Else
                
                    pReturnString = pReturnString & "[" & pParameterInfo.Name & pVarType & "],"
                End If
            Else
                pReturnString = pReturnString & pParameterInfo.Name & pVarType & ","
            End If
            
        Next
        If Right(pReturnString, 1) = "," Then
            pReturnString = Left(pReturnString, Len(pReturnString) - 1)
        End If
        pReturnString = pReturnString & ")"
        pVarType = ""
        If Not pMemberInfo.ReturnType.TypeInfo Is Nothing Then
            pVarType = " As " & Replace(pMemberInfo.ReturnType.TypeInfo.Name, "_", "")
            'Debug.Assert False
        Else
            
            Select Case pMemberInfo.ReturnType
                    Case TliVarType.VT_ARRAY:
                        
                        Debug.Assert False
                    Case TliVarType.VT_BLOB:
                        Debug.Assert False
                    Case TliVarType.VT_BLOB_OBJECT:
                        Debug.Assert False
                    Case TliVarType.VT_BOOL:
                        pVarType = " As Boolean"
                    Case TliVarType.VT_BSTR:
                        pVarType = " As String"
                        'Debug.Assert False
                    Case TliVarType.VT_BYREF:
                        Debug.Assert False
                    Case TliVarType.VT_CARRAY:
                        Debug.Assert False
                    Case TliVarType.VT_CF:
                        Debug.Assert False
                    Case TliVarType.VT_CLSID:
                        Debug.Assert False
                    Case TliVarType.VT_CY:
                        pVarType = " As Currency"
                    Case TliVarType.VT_DATE:
                        pVarType = " As Date"
                    Case TliVarType.VT_DECIMAL:
                        pVarType = " As Double"
                    Case TliVarType.VT_DISPATCH:
                        pVarType = " As Object"
'                        Debug.Assert False
                    Case TliVarType.VT_EMPTY:
                        Debug.Assert False
                    Case TliVarType.VT_ERROR:
                        Debug.Assert False
                    Case TliVarType.VT_FILETIME:
                        Debug.Assert False
                    Case TliVarType.VT_HRESULT:
                        pVarType = " As Long"
                    Case TliVarType.VT_I1:
                        pVarType = " As Byte"
                    Case TliVarType.VT_I2:
                        pVarType = "As Boolean"
                    Case TliVarType.VT_I4:
                        pVarType = " As Long"
                        'Debug.Assert False
                    Case TliVarType.VT_I8:
                        pVarType = " As Long"
                    Case TliVarType.VT_INT:
                        pVarType = " As Long"
                    Case TliVarType.VT_LPSTR:
                        pVarType = " As Long"
                    Case TliVarType.VT_LPWSTR:
                        pVarType = " As Long"
                    Case TliVarType.VT_NULL:
                        Debug.Assert False
                    Case TliVarType.VT_PTR:
                        pVarType = " As Long"
                    Case TliVarType.VT_R4:
                        pVarType = " As Single"
                    Case TliVarType.VT_R8:
                        pVarType = " As Double"
                    Case TliVarType.VT_RECORD:
                        Debug.Assert False
                    Case TliVarType.VT_RESERVED:
                        Debug.Assert False
                    Case TliVarType.VT_SAFEARRAY:
                        Debug.Assert False
                    Case TliVarType.VT_STORAGE:
                        Debug.Assert False
                    Case TliVarType.VT_STORED_OBJECT:
                        pVarType = " As Object"
                    Case TliVarType.VT_STREAM:
                        Debug.Assert False
                    Case TliVarType.VT_STREAMED_OBJECT:
                        pVarType = " As Object"
                    Case TliVarType.VT_UI1:
                        pVarType = " As Byte"
                    Case TliVarType.VT_UI2:
                        pVarType = " As Single"
                    Case TliVarType.VT_UI4:
                        pVarType = " As Long"
                    Case TliVarType.VT_UI8:
                        pVarType = " As Long"
                    Case TliVarType.VT_UINT:
                        pVarType = " As Integer"
                    Case TliVarType.VT_UNKNOWN:
                        pVarType = ""
                    Case TliVarType.VT_USERDEFINED:
                        Debug.Assert False
                    Case TliVarType.VT_VARIANT:
                        pVarType = ""
                        'Debug.Assert False
                    Case TliVarType.VT_VECTOR:
                        
                        Debug.Assert False
                    Case TliVarType.VT_VOID:
                        pVarType = ""
                End Select
        End If
        If Len(pVarType) > 0 Then
            pReturnString = pReturnString & pVarType
        End If
    End If
    GetParameterTip = pReturnString
End Function
Private Function GetClassName(ByVal pShowAfterEqual As Boolean, ByVal pExtVars As String, ByVal pProgID As String, ByVal pRefrence As Refrence, ByRef pColl As Collection) As String
    Dim pobjExtends() As String
    Dim pCounter As Long
    Dim pCurrentItem As String
    Dim pTypeLibInfo As TLI.TypeLibInfo
    Dim pCoClassInfo As TLI.CoClassInfo
    Dim pInterfaceInfo As TLI.InterfaceInfo
    Dim pConstantInfo As TLI.ConstantInfo
    Dim pMemberInfo As MemberInfo
    Dim pLib As String
    Dim pClass As String
    Dim pPos As Long
    
    Dim pTempText As String
    pPos = InStr(1, pProgID, ".")
    If pPos > 0 Then
        pLib = Left(pProgID, pPos - 1)
        pClass = Mid(pProgID, pPos + 1)
    Else
        Exit Function
         Debug.Assert False
    End If
    pPos = InStr(1, pExtVars, "(", vbTextCompare)
    If pPos > 0 Then
        pTempText = Left(pExtVars, pPos - 1)
        pPos = InStr(1, pExtVars, ")", vbTextCompare)
        If pPos > 0 Then
            pTempText = pTempText & Mid(pExtVars, pPos + 1)
        End If
        pExtVars = pTempText
    End If
    
    pobjExtends = Split(pExtVars, ".")
        
    Set pTypeLibInfo = m_TLIApplication.TypeLibInfoFromFile(pRefrence.BinaryPath)
    Dim pSearchItem As SearchItem
    Dim pSearchResults As SearchResults
    Dim pSearchLong As Long
    Dim pTypeInfo As TLI.TypeInfo
    Dim pMemberID As Long
    Dim pcTypeInfo As cTypeLibInfo
    Dim pReturnString As String
    
    Set pColl = New Collection
    'For pCounter = LBound(pobjExtends) To UBound(pobjExtends)
        pCurrentItem = pClass 'pobjExtends(pCounter)
        pSearchLong = pTypeLibInfo.MakeSearchData(pCurrentItem, tliStAll)
        Set pSearchResults = pTypeLibInfo.GetMembers(pSearchLong)
        For Each pSearchItem In pSearchResults
            Set pMemberInfo = pTypeLibInfo.GetMemberInfo(pSearchLong, pSearchItem.InvokeKinds, pSearchItem.MemberId, pSearchItem.Name)
            Set pcTypeInfo = New cTypeLibInfo
            If Len(pExtVars) > 0 Then
                If pMemberInfo.Name = pobjExtends(0) Then
                    If pMemberInfo.ReturnType.TypeInfo Is Nothing Then
                        pReturnString = GetParameterTip(pMemberInfo)
                    Else
                        pReturnString = GetParameterTip(pMemberInfo)
                        If Len(pReturnString) > 0 Then
                            Call EnumFindMemberID(pTypeLibInfo, pShowAfterEqual, pExtVars, pMemberInfo.ReturnType.TypeInfo, pColl)
                        Else
                            pReturnString = EnumFindMemberID(pTypeLibInfo, pShowAfterEqual, pExtVars, pMemberInfo.ReturnType.TypeInfo, pColl)
                        End If
                        
                    End If
                End If
            Else
                Set pcTypeInfo = New cTypeLibInfo
                pcTypeInfo.ClassName = pMemberInfo.Name
                Select Case pSearchItem.InvokeKinds
                    Case InvokeKinds.INVOKE_CONST:
                        pcTypeInfo.TypeKind = iConstant
                    Case InvokeKinds.INVOKE_EVENTFUNC:
                        pcTypeInfo.TypeKind = iEvent
                    Case InvokeKinds.INVOKE_FUNC:
                        pcTypeInfo.TypeKind = iFunction
                    Case InvokeKinds.INVOKE_PROPERTYGET:
                        pcTypeInfo.TypeKind = iProperty
                    Case InvokeKinds.INVOKE_PROPERTYPUT:
                        pcTypeInfo.TypeKind = iProperty
                    Case InvokeKinds.INVOKE_PROPERTYPUTREF:
                        pcTypeInfo.TypeKind = iProperty
                    Case InvokeKinds.INVOKE_UNKNOWN:
                        pcTypeInfo.TypeKind = iFunction
                    Case Else
                        If CInt(pSearchItem.InvokeKinds And INVOKE_CONST) = InvokeKinds.INVOKE_CONST Then
                            pcTypeInfo.TypeKind = iConstant
                        ElseIf CInt(pSearchItem.InvokeKinds And INVOKE_EVENTFUNC) = InvokeKinds.INVOKE_EVENTFUNC Then
                            pcTypeInfo.TypeKind = iEvent
                        ElseIf CInt(pSearchItem.InvokeKinds And INVOKE_FUNC) = InvokeKinds.INVOKE_FUNC Then
                            pcTypeInfo.TypeKind = iFunction
                        ElseIf CInt(pSearchItem.InvokeKinds And INVOKE_PROPERTYGET) = InvokeKinds.INVOKE_PROPERTYGET Then
                            pcTypeInfo.TypeKind = iProperty
                        ElseIf CInt(pSearchItem.InvokeKinds And INVOKE_PROPERTYPUT) = InvokeKinds.INVOKE_PROPERTYPUT Then
                            pcTypeInfo.TypeKind = iProperty
                        ElseIf CInt(pSearchItem.InvokeKinds And INVOKE_PROPERTYPUTREF) = InvokeKinds.INVOKE_PROPERTYPUTREF Then
                            pcTypeInfo.TypeKind = iProperty
                        ElseIf CInt(pSearchItem.InvokeKinds And INVOKE_UNKNOWN) = InvokeKinds.INVOKE_UNKNOWN Then
                            pcTypeInfo.TypeKind = IUnknown
                        End If
                        
                End Select
                If pcTypeInfo.TypeKind = iConstant Then
                    pcTypeInfo.Tag = "CONST:" & CLng(pMemberInfo.Value)
                End If
                If (pShowAfterEqual) Then
                    If pcTypeInfo.TypeKind = iConstant Then
                        pColl.Add pcTypeInfo, pcTypeInfo.ClassName
                    End If
                Else
                    pColl.Add pcTypeInfo, pcTypeInfo.ClassName
                End If
                
                'If pMemberInfo.ReturnType.TypeInfo Is Nothing Then
                '    Debug.Print pMemberInfo.Name, pMemberInfo.InvokeKind, pMemberInfo.ReturnType
                'Else
                '    Debug.Print pMemberInfo.Name, pMemberInfo.InvokeKind, pMemberInfo.ReturnType.TypeInfo.Name
                'End If
            End If
            Set pcTypeInfo = Nothing
        Next
    'Next
    GetClassName = pReturnString
End Function

Private Function EnumFindMemberID(ByVal pTypeLibInfo As TLI.TypeLibInfo, ByVal pShowAfterEqual As Boolean, ByVal pMemberName As String, ByVal pTypeInfo As TypeInfo, ByRef pColl As Collection) As String
    Dim pSearchData As Long
    Dim pSearchItem As SearchItem
    Dim pSearchResults As SearchResults
    Dim pMemberInfo As MemberInfo
    Dim pName As String
    Dim pNextName As String
    Dim pcTypeInfo As cTypeLibInfo
    Dim pRestName As String
    Dim pPos As Long
    Dim pReturnString As String
    Dim pFound As Boolean
    
    pPos = InStr(1, pMemberName, ".")
    If pPos > 0 Then
        pName = Left(pMemberName, pPos - 1)
        pNextName = Mid(pMemberName, pPos + 1)
    Else
        pName = pMemberName
        pNextName = ""
        
    End If
    
    pPos = InStr(1, pNextName, ".", vbTextCompare)
    If pPos > 0 Then
        pRestName = Mid(pNextName, pPos + 1)
        pNextName = Left(pNextName, pPos - 1)
    End If
    
lblPopulateAgain:
    
    For Each pMemberInfo In pTypeInfo.Members
        If pNextName <> "" Then
            If UCase(pMemberInfo.Name) = UCase(pNextName) Then
                    pFound = True
                    If pMemberInfo.ReturnType.TypeInfo Is Nothing Then
                        pReturnString = GetParameterTip(pMemberInfo)
                    Else
                        If Len(pRestName) = 0 Then
                            pReturnString = GetParameterTip(pMemberInfo)
                            Call EnumFindMemberID(pTypeLibInfo, pShowAfterEqual, pNextName & "." & pRestName, pMemberInfo.ReturnType.TypeInfo, pColl)
                        Else
                            pReturnString = EnumFindMemberID(pTypeLibInfo, pShowAfterEqual, pNextName & "." & pRestName, pMemberInfo.ReturnType.TypeInfo, pColl)
                        End If
                    End If
                    Exit For
            End If
        Else
            pFound = True
            If CInt(pMemberInfo.AttributeMask) = TYPEFLAG_NONE Then
            pReturnString = GetParameterTip(pMemberInfo)
              '  pSearchData = pTypelibInfo.MakeSearchData(pMemberInfo.Name, tliStAll)
                Set pSearchResults = pTypeLibInfo.GetTypesWithMember(pMemberInfo.Name, , tliStAll, False, False)
                If Not pSearchResults Is Nothing Then
                    If pSearchResults.Count = 1 Then
                        Set pSearchItem = pSearchResults.Item(1)
                    End If
                End If
                Set pcTypeInfo = New cTypeLibInfo
                pcTypeInfo.ClassName = pMemberInfo.Name
                Select Case pMemberInfo.InvokeKind
                    Case InvokeKinds.INVOKE_CONST:
                        pcTypeInfo.TypeKind = iConstant
                    Case InvokeKinds.INVOKE_EVENTFUNC:
                        pcTypeInfo.TypeKind = iEvent
                    Case InvokeKinds.INVOKE_FUNC:
                        pcTypeInfo.TypeKind = iFunction
                    Case InvokeKinds.INVOKE_PROPERTYGET:
                        pcTypeInfo.TypeKind = iProperty
                    Case InvokeKinds.INVOKE_PROPERTYPUT:
                        pcTypeInfo.TypeKind = iProperty
                    Case InvokeKinds.INVOKE_PROPERTYPUTREF:
                        pcTypeInfo.TypeKind = iProperty
                    Case InvokeKinds.INVOKE_UNKNOWN:
                        
                        
                            pcTypeInfo.TypeKind = iConstant
                        
                    Case Else
                        If CInt(pSearchItem.InvokeKinds And INVOKE_CONST) = InvokeKinds.INVOKE_CONST Then
                            pcTypeInfo.TypeKind = iConstant
                        ElseIf CInt(pSearchItem.InvokeKinds And INVOKE_EVENTFUNC) = InvokeKinds.INVOKE_EVENTFUNC Then
                            pcTypeInfo.TypeKind = iEvent
                        ElseIf CInt(pSearchItem.InvokeKinds And INVOKE_FUNC) = InvokeKinds.INVOKE_FUNC Then
                            pcTypeInfo.TypeKind = iFunction
                        ElseIf CInt(pSearchItem.InvokeKinds And INVOKE_PROPERTYGET) = InvokeKinds.INVOKE_PROPERTYGET Then
                            pcTypeInfo.TypeKind = iProperty
                        ElseIf CInt(pSearchItem.InvokeKinds And INVOKE_PROPERTYPUT) = InvokeKinds.INVOKE_PROPERTYPUT Then
                            pcTypeInfo.TypeKind = iProperty
                        ElseIf CInt(pSearchItem.InvokeKinds And INVOKE_PROPERTYPUTREF) = InvokeKinds.INVOKE_PROPERTYPUTREF Then
                            pcTypeInfo.TypeKind = iProperty
                        ElseIf CInt(pSearchItem.InvokeKinds And INVOKE_UNKNOWN) = InvokeKinds.INVOKE_UNKNOWN Then
                            pcTypeInfo.TypeKind = IUnknown
                        End If
                End Select
                If pcTypeInfo.TypeKind = iConstant Then
                    pcTypeInfo.Tag = "CONST:" & CLng(pMemberInfo.Value)
                End If
                If Not (Not pShowAfterEqual And pSearchResults.Item(1).SearchType = tliStConstants) Then
                    On Error Resume Next
                    pColl.Add pcTypeInfo, pcTypeInfo.ClassName
                End If
            End If
        End If
lblSkip:
    Next
    If Not pFound Then
        pNextName = ""
        GoTo lblPopulateAgain
    End If
    
    EnumFindMemberID = pReturnString
End Function

Public Function GetVarProperties(ByVal pShowAfterEqual As Boolean, ByVal pExtVars As String, ByVal pProgID As String, ByVal pRefrence As Refrence, ByRef pColl As Collection) As String
    Dim pReturnString As String
    pReturnString = GetClassName(pShowAfterEqual, pExtVars, pProgID, pRefrence, pColl)
    GetVarProperties = pReturnString
End Function

Private Function IsVariable(ByVal pVarName As String) As Boolean
    Dim pModule As Module
    Dim pVariable As Variable
    On Error Resume Next
    IsVariable = False
    pVarName = Trim(pVarName)
    If Len(pVarName) > 0 Then
    For Each pModule In g_Project.Modules
        Set pVariable = pModule.Variables.Item(pVarName)
        If Not pVariable Is Nothing Then
            IsVariable = True
            Exit For
        End If
    Next
    End If
End Function
