VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPropertyBag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_StartNode As Object 'MSXML.IXMLDOMNode 'Object
Private m_XMLProps As Object 'MSXML.DOMDocument  'Object

Private Enum epbtTypeConstants
   epbtImplementsPropBag
   epbtEmpty
   epbtNull
   epbtNothing
   epbtByteArray
   epbtDate
   epbtNumber
   epbtString
   epbtStringCData
   epbtError
End Enum

Public Enum epbtErrorConstants
   eecPropertyBagErrorBase = 29450
   eeInvalidVariantType
   eeInvalidXMLFile
End Enum

Private m_sTypeName As String

Private Sub Class_Terminate()

   Set m_StartNode = Nothing
   Set m_XMLProps = Nothing

End Sub

Sub WriteProperty(ByVal PropertyName As String, ByVal PropertyValue As Variant, Optional DefaultValue As Variant)

Dim oNode As Object 'IXMLDOMNode 'Object
Dim oClassElement As Object 'IXMLDOMNode  'Object
Dim oTransactionData As VbScriptEditor.cPropertyBag
Dim eType As epbtTypeConstants

   
   ' Evaluate the type of variant we are adding to the XML:
   eType = peValidateValueType(PropertyValue)
   
   ' Write out accordingly:
   If eType = epbtError Then
      ' Can't write!
      Err.Raise (vbObjectError + 1048 + eeInvalidVariantType), App.EXEName & ".vbalXMLPBag", "Invalid PropertyValue type."
   
   ElseIf eType = epbtImplementsPropBag Then
      ' Must recurse into that property bag implementation:
      Set oTransactionData = New cPropertyBag
      ' Prepare the property bag:
      'Ask the object to serialize itself onto a new FCBcPropertyBag object
      Set oNode = m_StartNode.appendChild(m_StartNode.ownerDocument.createElement(TypeName(PropertyValue)))
      IXMLDOMElement(oNode).setAttribute "name", PropertyName
      Set oTransactionData.StartNode = oNode
      PersistPropertyBag(PropertyValue).Save oTransactionData
      
   Else
      ' We can write:
      Set oNode = m_StartNode.appendChild(m_StartNode.ownerDocument.createElement(PropertyName))
      
      Select Case eType
      Case epbtDate
         IXMLDOMElement(oNode).Text = PropertyValue
      Case epbtNumber
         IXMLDOMElement(oNode).Text = PropertyValue
      Case epbtString
         IXMLDOMElement(oNode).Text = PropertyValue
      Case epbtStringCData
         Dim cD As Object 'IXMLDOMCDATASection
         Set cD = m_StartNode.ownerDocument.createCDATASection(PropertyValue)
         oNode.appendChild cD
      Case epbtByteArray
         With IXMLDOMElement(oNode)
            .dataType = "bin.base64"
            .nodeTypedValue = PropertyValue
         End With
      Case epbtEmpty
         ' Special values:
         IXMLDOMElement(oNode).setAttribute "vartype", "empty"
      Case epbtNothing
         ' Special values:
         IXMLDOMElement(oNode).setAttribute "vartype", "nothing"
      Case epbtNull
         ' Special values:
         IXMLDOMElement(oNode).setAttribute "vartype", "null"
      End Select
   End If

End Sub
Private Function peValidateValueType(Value As Variant) As epbtTypeConstants
   
   ' Here we evaluate what type of variant we're
   ' trying to write out to the XML:
   peValidateValueType = epbtError
   If IsEmpty(Value) Then
      peValidateValueType = epbtEmpty
   ElseIf IsNull(Value) Then
      peValidateValueType = epbtNull
   ElseIf IsObject(Value) Then
      If Value Is Nothing Then
         peValidateValueType = epbtNothing
      ElseIf pbImplementsXMLPropBag(Value) Then
         peValidateValueType = epbtImplementsPropBag
      Else
         ' Evalute for default value.
         peValidateValueType = peGetOtherType(Value)
      End If
   ElseIf IsError(Value) Then
      peValidateValueType = epbtError
   ElseIf IsNumeric(Value) Then
      peValidateValueType = epbtNumber
   ElseIf IsDate(Value) Then
      peValidateValueType = epbtDate
   ElseIf IsArray(Value) Then
      ' only byte arrays are supported:
      If (VarType(Value) And vbByte) = vbByte Then
         peValidateValueType = epbtByteArray
      End If
   Else
      peValidateValueType = peGetOtherType(Value)
   End If
   
End Function

Private Function peGetOtherType(Value As Variant) As epbtTypeConstants
Dim sThis As String
   
   On Error Resume Next
   sThis = Value
   If Err.Number = 0 Then
      peGetOtherType = epbtString
      
      ' Does sThis contain XML delimiters?
      ' NB: rough check.  If you're worried, pass the
      ' information into WriteProperties as a byte
      ' array & the bin.base64 will do its work, then
      ' you'll never have difficulties.
      If Not (InStr(sThis, "<") = 0) Then
         peGetOtherType = epbtStringCData
      ElseIf Not (InStr(sThis, ">") = 0) Then
         peGetOtherType = epbtStringCData
      End If
      
   Else
      peGetOtherType = epbtError
   End If
   Err.Clear
   
End Function

Private Function pbImplementsXMLPropBag(Value As Variant) As epbtTypeConstants
Dim ixpb As PersistPropertyBag

   On Error Resume Next
   Set ixpb = Value
   pbImplementsXMLPropBag = (Err.Number = 0)
   Err.Clear
   
End Function

Function ReadProperty(PropertyName As String, Optional DefaultValue As Variant) As Variant

Dim oTransactionData As VbScriptEditor.cPropertyBag
Dim Value As Variant
Dim oNode As Object 'IXMLDOMNode  'Object
Dim oAtt  As Object 'IXMLDOMAttribute  'Object
Dim sQuery As String
Dim bFound As Boolean
Dim oRes As Object 'IXMLDOMNodeList  'Object
Dim sTypeName As String

   If pbImplementsXMLPropBag(DefaultValue) Then
      'sTypeName = m_sTypeName & "/" & TypeName(DefaultValue) & "[@name = """ & PropertyName & """]"
      sTypeName = TypeName(DefaultValue) & "[@name = """ & PropertyName & """]"
      'sQuery = "//" & sTypeName
      sQuery = sTypeName
      Set oRes = m_StartNode.selectNodes(sQuery)
      If oRes.length > 0 Then
         Set oTransactionData = New cPropertyBag
         Set oNode = oRes(0)
         Set oTransactionData.StartNode = oNode
         oTransactionData.TransTypeName = sTypeName
         PersistPropertyBag(DefaultValue).Load oTransactionData
      End If
   Else
      ReadProperty = DefaultValue
         
      'sQuery = "//" & m_sTypeName & "/" & PropertyName
      sQuery = PropertyName
      Set oRes = m_StartNode.selectNodes(sQuery)
      If oRes.length > 0 Then
         Set oNode = oRes(0)
         If oNode.dataType = "bin.base64" Then
            ReadProperty = oNode.nodeTypedValue
         Else
            ReadProperty = oNode.Text
         End If
      End If
   End If
   
End Function


Public Sub SaveState(ByVal RootObject As PersistPropertyBag, Optional ByVal sName As String)
Dim sQuery As String
Dim oRes As Object 'IXMLDOMNodeList  'Object
Dim oPBagElement As Object 'IXMLDOMNode  'Object
Dim oElement As Object 'IXMLDOMNode  'Object
Dim oClassElement As Object 'IXMLDOMNode  'Object
Dim pBaseEleQuery As String
Static pStackDepth As Long

    If Not RootObject Is Nothing Then
        pStackDepth = pStackDepth + 1
        ' Ensure the document is available:
        If m_XMLProps Is Nothing Then
            'Set m_XMLProps = New DOMDocument
            Set m_XMLProps = CreateObject("MSXML.DOMDocument")
            Set m_StartNode = m_XMLProps
        End If
        ' Make sure the root tag is present:
        If pStackDepth = 1 Then
            sQuery = "//cPropertyBag"
            pBaseEleQuery = "cPropertyBag"
        Else
            sQuery = TypeName(RootObject)
            pBaseEleQuery = sQuery
        End If
        
        Set oRes = m_StartNode.selectNodes(sQuery)
        
        If oRes.length = 0 Then
            If pStackDepth = 1 Then
                Set oPBagElement = m_StartNode.appendChild(m_XMLProps.createElement(pBaseEleQuery))
            Else
                'Debug.Assert False
            End If
            
        Else
            Set oPBagElement = oRes.Item(0)
        End If
        
        If pStackDepth > 1 Then
            m_sTypeName = m_sTypeName & "/" & TypeName(RootObject)
        Else
            m_sTypeName = TypeName(RootObject)
        End If
        
        'Set oElement = m_XMLProps.createElement(m_sTypeName)
        'Set oElement = m_XMLProps.createElement(TypeName(RootObject))
        If pStackDepth > 1 Then
            'Set oElement = m_StartNode.appendChild(m_StartNode.ownerDocument.createElement(TypeName(RootObject)))
            Set oElement = m_StartNode.appendChild(m_XMLProps.createElement(TypeName(RootObject)))
        Else
            Set oElement = m_XMLProps.createElement(TypeName(RootObject))
        End If
        
        If Len(sName) > 0 Then
            IXMLDOMElement(oElement).setAttribute "name", sName
        End If
        
        sQuery = TypeName(RootObject)
        
        If Len(sName) > 0 Then
            sQuery = sQuery & "[@name=""" & sName & """]"
        End If
        
        Set oRes = m_StartNode.selectNodes(sQuery)
        If oRes.length = 0 Then
            ' Need a new element:
            Set oClassElement = IXMLDOMElement(oPBagElement).appendChild(oElement)
        Else
            ' replace
            'Call m_StartNode.replaceChild(oElement, IXMLDOMElement(oRes.Item(0)))
            Set oClassElement = oElement
        End If
        
        Set m_StartNode = oClassElement
        RootObject.Save Me
        pStackDepth = pStackDepth - 1
        If Len(m_sTypeName) - Len(TypeName(RootObject)) > 0 Then
            m_sTypeName = Left(m_sTypeName, Len(m_sTypeName) - Len(TypeName(RootObject)) - 1)
            Set m_StartNode = m_StartNode.parentNode
        Else
            m_sTypeName = ""
            Set m_StartNode = Nothing
        End If
    End If
   
End Sub

Public Sub RestoreState(ByVal RootObject As PersistPropertyBag, Optional ByVal sName As String = "")
    '
    'Starts the process of restoring the object hierarchy
    '
    Dim sQuery As String
    Dim oRes As Object 'IXMLDOMNodeList  'Object
    Dim pBaseEleQuery As String
    Static pStackDepth As Long
    
    pStackDepth = pStackDepth + 1
    
    ' Make sure the root tag is present:
    If pStackDepth = 1 Then
        sQuery = "//cPropertyBag/" & TypeName(RootObject)
        pBaseEleQuery = "cPropertyBag"
        Set m_StartNode = m_XMLProps
    Else
        sQuery = TypeName(RootObject)
        pBaseEleQuery = sQuery
    End If
    
    If Len(sName) > 0 Then
        sQuery = sQuery & "[@name=""" & sName & """]"
    End If
    
    Set oRes = m_StartNode.selectNodes(sQuery)
    
    If pStackDepth > 1 Then
        m_sTypeName = m_sTypeName & "/" & TypeName(RootObject)
    Else
        m_sTypeName = TypeName(RootObject)
    End If
    
    
    If oRes.length > 0 Then
        Set m_StartNode = oRes.Item(0)
        RootObject.Load Me
        pStackDepth = pStackDepth - 1
        'If Len(sName) > 0 Then
        '    m_sTypeName = Left(m_sTypeName, Len(m_sTypeName) - Len("[@name=""" & sName & """]"))
        'End If
        
        If Len(m_sTypeName) - Len(TypeName(RootObject)) > 0 Then
            m_sTypeName = Left(m_sTypeName, Len(m_sTypeName) - Len(TypeName(RootObject)) - 1)
            Set m_StartNode = m_StartNode.parentNode
        Else
            m_sTypeName = ""
            Set m_StartNode = Nothing
        End If
    End If
    
End Sub


Public Sub DeleteObject(ByVal sTypeName As String, Optional ByVal sName As String = "")
Dim sQuery As String
Dim oRes As Object 'IXMLDOMNodeList  'Object
Dim oClassElement As Object 'IXMLDOMNode  'Object
Dim oRemove As Object 'IXMLDOMNode  'Object

   m_sTypeName = "cPropertyBag/" & sTypeName
   If Len(sName) > 0 Then
      m_sTypeName = m_sTypeName & "[@name=""" & sName & """]"
   End If
   
   sQuery = "//" & m_sTypeName
   Set oRes = m_XMLProps.selectNodes(sQuery)
   If oRes.length > 0 Then
      Set oRemove = oRes.Item(0)
      sQuery = "//cPropertyBag"
      Set oRes = m_XMLProps.selectNodes(sQuery)
      IXMLDOMElement(oRes.Item(0)).removeChild oRemove
   End If
End Sub

Friend Property Let TransTypeName(ByVal sTypeName As String)
   m_sTypeName = sTypeName
End Property

Friend Property Set StartNode(ByVal Node As Object) 'IXMLDOMNode)  'Object)
   Set m_StartNode = Node
   Set m_XMLProps = m_StartNode.ownerDocument
End Property

Friend Property Get StartNode() As Object 'IXMLDOMNode  'Object
   Set StartNode = m_StartNode
End Property

Public Property Let Contents(ByVal sXML As String)
   'Set m_XMLProps = New DOMDocument
   Set m_XMLProps = CreateObject("MSXML.DOMDocument")
   
   If Not m_XMLProps.loadXML(sXML) Then
      ' need to raise an error!
      Err.Raise (vbObjectError + 1048 + eeInvalidXMLFile), App.EXEName & ".vbalXMLPBag", m_XMLProps.parseError.reason
   End If
End Property

Public Property Get Contents() As String
   If Not m_XMLProps Is Nothing Then
      Contents = m_XMLProps.XML
   End If
End Property

Public Property Get Document() As Object 'DOMDocument  'Object
   Set Document = m_XMLProps
End Property
Public Property Let Document(oDoc As Object) 'DOMDocument) 'Object)
   Set m_XMLProps = oDoc
End Property

'Private Function IXMLDOMElement(ByVal Node As IXMLDOMNode) As IXMLDOMElement 'Object) As Object
Private Function IXMLDOMElement(ByVal Node As Object) As Object
   Set IXMLDOMElement = Node
End Function

Private Function PersistPropertyBag(ByVal Source As PersistPropertyBag) As PersistPropertyBag
   Set PersistPropertyBag = Source
End Function

'Private Function AttributeValue(ByVal Node As IXMLDOMNode, ByVal sAttributeName As String, Optional ByVal vDefault As Variant = Empty) As Variant
Private Function AttributeValue(ByVal Node As Object, ByVal sAttributeName As String, Optional ByVal vDefault As Variant = Empty) As Variant
   If Not IXMLDOMElement(m_StartNode).getAttributeNode(sAttributeName) Is Nothing Then
      AttributeValue = IXMLDOMElement(m_StartNode).getAttribute(sAttributeName)
   Else
      AttributeValue = vDefault
   End If
End Function



