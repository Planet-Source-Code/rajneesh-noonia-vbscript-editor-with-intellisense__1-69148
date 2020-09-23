VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QSXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' DECLARATIONS - START
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Option Explicit

' The DOMDocument object which holds the XML Document
' will be redimmed later, when exact class will be known
Dim XMLDoc As Object
Dim xmlTmpDoc As Object
' Constant which holds the generic object name for the DOMDocument
Private Const CONST_DOMDOCUMENT As String = "DOMDocument"

' Constants used to test-invoke the different MS XML parser version
Private Const CONST_INVOKE_DOMDOCUMENT_40 As String = "Msxml2.DOMDocument.4.0"  ' this can be invoked by MSXML40
Private Const CONST_INVOKE_DOMDOCUMENT_30 As String = "Msxml2.DOMDocument.3.0"  ' this can be invoked by MSXML30
Private Const CONST_INVOKE_DOMDOCUMENT_26 As String = "Msxml2.DOMDocument.2.6"  ' this can be invoked by MSXML30 & MSXML26
Private Const CONST_INVOKE_DOMDOCUMENT_20 As String = "Msxml.DOMDocument"       ' this can be invoked by MSXML30 & MSXML26 & MSXML20

' Module level variable which holds the .ParserVersion property
Dim mpavCurrentParserVersion As ParserVersionConstant

' Enumeration of currently supported parser versions the user may select
' with the .Initialize method
' NOTE: This is also used as a return value with the
' .Initialize method & .XMLParserVersion property (both public) and the
' .InvokeParser & .AutoDetectParser methods (both private)
Public Enum ParserVersionConstant
    ' 0 = return value only, do not assign!
    pavV20 = 1
    pavV26 = 2
    pavV30 = 3
    pavV40 = 4
    pavAUTO = 99
End Enum

' Enumeration of possible node relationships the user may select
' with the .InsertNode method
Public Enum NodeRelationshipConstant
    norSIBLING = 0
    norCHILD = 1
End Enum

' Enumeration of possible XML document information the user may select
' with the .XMLDocumentInfo property
Public Enum XMLDocumentInfoConstant
    xdiVERSION = 0
    xdiENCODING = 1
    xdiSTANDALONE = 2
End Enum

' Enumeration of possible Constants used
' with the .OpenXML method
Private Enum OpenXMLConstant
    oxFILE = 0
    oxSTRING = 1
End Enum

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' DECLARATIONS - END
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'-------------------------------------------------------------------------------

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' PROPERTIES (PUBLIC) - START
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Property Get XML() As String
    XML = XMLDoc.XML
End Property



'.------------------------------------------------------------------------------
'.  Function  : Public Property Get XMLDocumentInfo
'.
'.  Parameters: xdiDesiredInfo As XMLDocumentInfoConstant
'.
'.  Returns   : Variant (NULL = Error)
'.
'.  Comments  : Returns information on the XML document's "version", "encoding" or
'.              "standalone" attributes
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Property Get XMLDocumentInfo(xdiDesiredInfo As XMLDocumentInfoConstant) As Variant
    On Error GoTo ErrorHandler
    
    Dim strDesiredInfo As String
    If xdiDesiredInfo = xdiVERSION Then
        strDesiredInfo = "version"
    ElseIf xdiDesiredInfo = xdiENCODING Then
        strDesiredInfo = "encoding"
    ElseIf xdiDesiredInfo = xdiSTANDALONE Then
        strDesiredInfo = "standalone"
    End If
    
    Dim XMLNode As Object
    Set XMLNode = XMLDoc.childNodes(0).Attributes.getNamedItem(strDesiredInfo)
    XMLDocumentInfo = XMLNode.Text
    Set XMLNode = Nothing
    
    Exit Property
ErrorHandler:
    XMLDocumentInfo = Null
    Set XMLNode = Nothing
End Property

'.------------------------------------------------------------------------------
'.  Function  : Public Property Get XMLParserVersion
'.
'.  Parameters:
'.
'.  Returns   : ParserVersionConstant (0 = Error / none selected)
'.
'.  Comments  : Returns the currently invoked MS XML Parser version
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Property Get XMLParserVersion() As ParserVersionConstant
    XMLParserVersion = mpavCurrentParserVersion
End Property

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' PROPERTIES (PUBLIC) - END
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'-------------------------------------------------------------------------------

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' METHODS (PUBLIC) - START
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'.------------------------------------------------------------------------------
'.  Function  : Public Function ReadNodeXML
'.
'.  Parameters: ByVal strQueryString As String
'.
'.  Returns   : String (Error = "")
'.
'.  Comments  : Return the supplied node's XML value
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function ReadNodeXML(ByVal strQueryString As String) As String
    On Error GoTo ErrorHandler
    
    Dim XMLNode As Object
    Set XMLNode = XMLDoc.documentElement.selectSingleNode(strQueryString)
  
    ReadNodeXML = XMLNode.XML
    
    Set XMLNode = Nothing
    Exit Function
ErrorHandler:
    ReadNodeXML = ""
    Set XMLNode = Nothing
   
End Function

'.------------------------------------------------------------------------------
'.  Function  : Public Function Initialize
'.
'.  Parameters: pavDesiredVersion As ParserVersionConstant
'.
'.  Returns   : ParserVersionConstant (0 = Error)
'.
'.  Comments  : Tries to initialize the desired MS XML parser version
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function Initialize(Optional pavDesiredVersion As ParserVersionConstant = pavAUTO) As ParserVersionConstant
    Select Case pavDesiredVersion ' Which parser does the user want ot invoke?
        Case pavAUTO: ' Autodetect parser, use latest version
            mpavCurrentParserVersion = AutoDetectParser()
            If mpavCurrentParserVersion = 0 Then ' Oops, no parser installed
                Initialize = 0
                Exit Function
            Else
                 Initialize = mpavCurrentParserVersion
            End If
        Case Else: ' User selects parser himself. Better check, if really available
            ' Test invoke
            If InvokeParser(pavDesiredVersion) = True Then
                mpavCurrentParserVersion = pavDesiredVersion
                Initialize = mpavCurrentParserVersion
            Else
                mpavCurrentParserVersion = 0
                Initialize = 0
                Exit Function
            End If
    End Select
    Set XMLDoc = InitializeObject(CONST_DOMDOCUMENT)
End Function
'.------------------------------------------------------------------------------
'.  Function  : Private Function OpenXML
'.
'.  Parameters: ByVal oxFileOrString As OpenXMLConstant
'.              ByVal strSource As String
'.              Optional ByVal bolVerboseError As Boolean = False
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Tries to open and parse strSource,
'.              either from file or from string.
'.              Option: Display MS XML Parser Error in MsgBox?
'.
'.              !!!WHEN OPENING FROM STRING!!!
'.              MSXML 3: No Problems
'.
'.              MSXML 2.0 & MSXML 2.6:
'.              Will originally work with UTF-16 or UCS-2 encodings only
'.              (limitation of LoadXML method)
'.              If you use this and save, your XML Document
'.              *will be changed into UTF-16* !
'.
'.              ^^^
'.              Added a workaround to solve this problem:
'.              Write the supplied string into a text file,
'.              use OpenFromFile method,
'.              delete the text file

'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Private Function OpenXML(ByVal oxFileOrString As OpenXMLConstant, ByVal strSource As String, Optional ByVal bolVerboseError As Boolean = False) As Boolean
    ' Create the DOMDocument
    Set XMLDoc = Nothing
    Set XMLDoc = InitializeObject(CONST_DOMDOCUMENT)

    XMLDoc.resolveExternals = False
    XMLDoc.validateOnParse = True
    XMLDoc.async = False
    
    Dim bolLoadSuccess As Boolean
    Select Case oxFileOrString
        Case oxFILE:
            bolLoadSuccess = XMLDoc.Load(strSource) ' .Load = Load from File
        Case oxSTRING:
            bolLoadSuccess = XMLDoc.loadXML(strSource) ' .LoadXML = Load from String
    End Select
        
    Select Case bolLoadSuccess ' Loading successful?
        Case True:
            OpenXML = True
        Case False:
            On Error Resume Next
            Dim xmlParseError As Object
            Set xmlParseError = XMLDoc.parseError
            
            ' Can we apply the workaround? Is that the problem?
            If mpavCurrentParserVersion = pavV20 And oxFileOrString = oxSTRING And xmlParseError.errorCode = -1072896657 Then ' Guess so
                
                Dim strFile As String
                Dim intFileNumber As Integer

                strFile = App.Path & "\" & "tempgoxml.xml"
                intFileNumber = FreeFile()

                Open strFile For Output As intFileNumber
                Print #intFileNumber, strSource
                Close intFileNumber
                
                ' Try again, from file this time
                OpenXML = OpenFromFile(strFile, bolVerboseError)
            
                Kill (strFile)
            Else
                If bolVerboseError = True Then Call DisplayParseError
                OpenXML = False
                Set XMLDoc = Nothing
            End If
            
            Set xmlParseError = Nothing
        End Select
End Function

'.------------------------------------------------------------------------------
'.  Function  : Public Function OpenFromFile
'.
'.  Parameters: ByVal strFileName As String
'.              Optional ByVal bolVerboseError As Boolean = False
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Open XML document from file
'.              Option: Display MS XML Parser Error in MsgBox
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function OpenFromFile(ByVal strFilename As String, Optional ByVal bolVerboseError As Boolean = False) As Boolean
    OpenFromFile = OpenXML(oxFILE, strFilename, bolVerboseError)
End Function

'.------------------------------------------------------------------------------
'.  Function  : Public Function OpenFromString
'.
'.  Parameters: ByVal strSource As String
'.              Optional ByVal bolVerboseError As Boolean = False
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Open XML document from file
'.              Option: Display MS XML Parser Error in MsgBox
'.
'.              !!!READ THE .OPENXML METHODS COMMENT BLOCK!!!
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   :
'.------------------------------------------------------------------------------

Public Function OpenFromString(ByVal strSource As String, Optional ByVal bolVerboseError As Boolean = False) As Boolean
    OpenFromString = OpenXML(oxSTRING, strSource, bolVerboseError)
End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function NodeCount
'.
'.  Parameters: ByVal strQueryString As String
'.
'.  Returns   : Long (Error = -1)
'.
'.  Comments  : Return the number of nodes found
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function NodeCount(ByVal strQueryString As String) As Long
    On Error GoTo ErrorHandler

    Dim xmlNodeList As Object
    Set xmlNodeList = XMLDoc.documentElement.selectNodes(strQueryString)
    NodeCount = xmlNodeList.Length
    Set xmlNodeList = Nothing

    Exit Function
ErrorHandler:
    NodeCount = -1
    Set xmlNodeList = Nothing
End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function ReadNode
'.
'.  Parameters: ByVal strQueryString As String
'.
'.  Returns   : Variant (Error = NULL)
'.
'.  Comments  : Return the node text found
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function ReadNode(ByVal strQueryString As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim XMLNode As Object
    Set XMLNode = XMLDoc.documentElement.selectSingleNode(strQueryString)
  
    ReadNode = XMLNode.Text
    
    Set XMLNode = Nothing
    Exit Function
ErrorHandler:
    ReadNode = Null
    Set XMLNode = Nothing
   
End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function ReadAttribute
'.
'.  Parameters: ByVal strQueryString As String
'.              ByVal strAttributeName As String
'.
'.  Returns   : Variant (Error = NULL)
'.
'.  Comments  : Return supplied attribute's value from supplied node
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function ReadAttribute(ByVal strQueryString As String, ByVal strAttributeName As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim xmlAttributeNode As Object
    Set xmlAttributeNode = XMLDoc.documentElement.selectSingleNode(strQueryString)

    ReadAttribute = xmlAttributeNode.getAttribute(strAttributeName)

    Set xmlAttributeNode = Nothing
    Exit Function
ErrorHandler:
    ReadAttribute = Null
    Set xmlAttributeNode = Nothing
End Function
Public Function RemoveNode(ByVal XMLNode As Object) As Boolean
    On Error GoTo ErrorHandler
    XMLNode.ParentNode.removeChild XMLNode ' This one was tricky ;-)
    
    Set XMLNode = Nothing
    RemoveNode = True
    Exit Function
ErrorHandler:
    RemoveNode = False
    Set XMLNode = Nothing
End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function DeleteNode
'.
'.  Parameters: ByVal strQueryString As String
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Delete node
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function DeleteNode(ByVal strQueryString As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim XMLNode As Object
    Set XMLNode = XMLDoc.documentElement.selectSingleNode(strQueryString)
    
    XMLNode.ParentNode.removeChild XMLNode ' This one was tricky ;-)
    
    Set XMLNode = Nothing
    DeleteNode = True
    Exit Function
ErrorHandler:
    DeleteNode = False
    Set XMLNode = Nothing
End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function WriteNode
'.
'.  Parameters: ByVal strQueryString As String
'.              ByVal vntNewNodeValue As Variant
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Write a supplied node's text value (as variant)
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function WriteNode(ByVal strQueryString As String, ByVal vntNewNodeValue As Variant) As Boolean
    On Error GoTo ErrorHandler
   
    Dim XMLNode As Object
    Set XMLNode = XMLDoc.documentElement.selectSingleNode(strQueryString)
    
    XMLNode.Text = vntNewNodeValue
    
    Set XMLNode = Nothing
    WriteNode = True
    Exit Function
ErrorHandler:
    WriteNode = False
    Set XMLNode = Nothing
End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function WriteAttribute
'.
'.  Parameters: ByVal strQueryString As String
'.              ByVal strAttributeName As String
'.              ByVal vntNewAttributeValue As Variant
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Write supplied node's supplied attribute's value
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function WriteAttribute(ByVal strQueryString As String, ByVal strAttributeName As String, ByVal vntNewAttributeValue As Variant) As Boolean
    On Error GoTo ErrorHandler

    Dim xmlAttributeNode As Object
    Set xmlAttributeNode = XMLDoc.documentElement.selectSingleNode(strQueryString)
    
    xmlAttributeNode.SetAttribute strAttributeName, vntNewAttributeValue
    
    Set xmlAttributeNode = Nothing
    WriteAttribute = True
    Exit Function
ErrorHandler:
    WriteAttribute = False
    Set xmlAttributeNode = Nothing
End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function WriteCDataSection
'.
'.  Parameters: ByVal strQueryString As String
'.              ByVal vntNewCDataValue As Variant
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Write supplied node's CDATA section
'.              WILL TRY TO OVERWRITE AN EXISTING CDATA SECTION,
'.              INSTEAD OF APPENDING TO IT.
'.              Personally, I can't think of any good reason to have
'.              more than one CDATA section, but I may just as well be wrong...
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function WriteCDataSection(ByVal strQueryString As String, ByVal vntNewCDataValue As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    Dim XMLNode As Object
    Dim xmlCDATASection As Object
    Set XMLNode = XMLDoc.documentElement.selectSingleNode(strQueryString)

    Select Case XMLNode.childNodes.Length
        Case 0: ' No Problem, go ahead
            Set xmlCDATASection = XMLDoc.createCDATASection(vntNewCDataValue)
            XMLNode.appendChild xmlCDATASection
        Case 1: ' UHOH, better check, if there is already a CData section in form of a CHILD Node present
            Select Case XMLNode.childNodes(0).nodeTypeString
                Case "cdatasection": ' Yes, CDATA section already exists, change Text-Property only!
                    XMLNode.childNodes.Item(0).Text = vntNewCDataValue
                Case Else: ' Probably a Text section, delete Text and replace with CData section
                    XMLNode.Text = ""
                    Set xmlCDATASection = XMLDoc.createCDATASection(vntNewCDataValue)
                    XMLNode.appendChild xmlCDATASection
            End Select
        Case Else: ' Raise Error, if more than 1 child node
            GoTo ErrorHandler
    End Select

    Set xmlCDATASection = Nothing
    Set XMLNode = Nothing
    WriteCDataSection = True
    Exit Function
ErrorHandler:
    WriteCDataSection = False
    Set xmlCDATASection = Nothing
    Set XMLNode = Nothing
End Function
Public Function GetRootElement() As Object
    Set GetRootElement = XMLDoc.documentElement()
End Function
Public Function GetRootChildren() As Object
Dim ret As Object
    Set ret = XMLDoc.documentElement()
    If ret.hasChildNodes Then
        Set GetRootChildren = ret.childNodes
        Set ret = Nothing
    Else
        Set ret = Nothing
        Set GetFirstChild = Nothing
    End If
End Function
Public Function GetChildNodeList(nd As Object) As Object
    Set GetChildNodeList = nd.childNodes
End Function
Public Function GetFirstChild() As Object
Dim ret As Object
    Set ret = XMLDoc.documentElement()
    If ret.hasChildNodes Then
        Set GetFirstChild = ret.firstChild
        Set ret = Nothing
    Else
        Set ret = Nothing
        Set GetFirstChild = Nothing
    End If
End Function
Public Function CreateRootElement(ByVal strID As String, ByVal strNewNodeName As String, Optional ByVal vntNewNodeValue As Variant, Optional ByVal strNewAttributeName As String, Optional ByVal vntNewAttributeValue As Variant) As Boolean
    On Error GoTo ErrorHandler
    Dim XMLNode As Object
    Dim xmlChildNode As Object
    Dim xmlAttributeNode As Object
    Set xmlChildNode = XMLDoc.createElement(strNewNodeName)
    If Not IsMissing(vntNewNodeValue) Then
        xmlChildNode.Text = vntNewNodeValue
    End If

Dim NamedNodeMap As Object
Dim xAtt As Object
Set xAtt = XMLDoc.createAttribute("id")
Set NamedNodeMap = xmlChildNode.Attributes
NamedNodeMap.setNamedItem xAtt
    Set xmlAttributeNode = XMLDoc.appendChild(xmlChildNode)
    xmlAttributeNode.SetAttribute "id", strID
    
    If Not IsMissing(strNewAttributeName) And Not IsMissing(vntNewAttributeValue) Then
        xmlAttributeNode.SetAttribute strNewAttributeName, vntNewAttributeValue
    End If
    Set NamedNodeMap = Nothing
    Set xAtt = Nothing
    Set xmlAttributeNode = Nothing
    Set xmlChildNode = Nothing
    Set XMLNode = Nothing
    CreateRootElement = True
    Exit Function
ErrorHandler:
    CreateRootElement = False
    Set NamedNodeMap = Nothing
    Set xAtt = Nothing
    Set xmlAttributeNode = Nothing
    Set xmlChildNode = Nothing
    Set XMLNode = Nothing
End Function
Public Function GetNodeList(strNodeName As String) As Object
Dim l As Object
Dim i As Integer, j As Integer

Set GetNodeList = XMLDoc.getElementsByTagName(strNodeName)


End Function
Public Function GetNodeFromAttribute(strTag As String, strAttName As String, strAttValue As String) As Object
Dim l As Object
Dim i As Integer, j As Integer

Set l = XMLDoc.getElementsByTagName(strTag)
  For i = 0 To (l.Length - 1)
    For j = 0 To l.Item(i).Attributes.Length - 1
        If (UCase$(l.Item(i).Attributes.Item(j).nodename) = UCase$(strAttName)) _
        And (UCase$(l.Item(i).Attributes.Item(j).nodeValue) = UCase$(strAttValue)) Then
            Set GetNodeFromAttribute = l.Item(i)
            Set l = Nothing
            Exit Function
        End If
    Next j
  Next i
Set GetNodeFromAttribute = Nothing
End Function
Public Function GetAttributeValue(objNode As Object, strAttName As String) As String
Dim i As Integer, j As Integer
    For j = 0 To objNode.Attributes.Length - 1
        If (UCase$(objNode.Attributes.Item(j).nodename) = UCase$(strAttName)) Then
            GetAttributeValue = objNode.Attributes.Item(j).nodeValue
            Exit Function
        End If
    Next j
GetAttributeValue = ""
End Function
Public Sub SetAttribute(ByVal xmlAttributeNode As Object, strAtt As String, strValue As String)
    xmlAttributeNode.SetAttribute strAtt, strValue
End Sub
Public Function IsChildNode(ndNode As Object, strNodeName As String) As Boolean
    Dim i As Long
    Dim ndName As String
    Dim ndList As Object
    If Not ndNode.hasChildNodes Then
        IsChildNode = False
        Exit Function
    End If
    ndName = UCase$(Replace(strNodeName, " ", "_"))
    Set ndList = ndNode.childNodes
    For i = 0 To ndList.Length - 1
        If UCase$(ndList(i).nodename) = ndName Then
            IsChildNode = True
            Set ndList = Nothing
            Exit Function
        End If
    Next
    Set ndList = Nothing
    IsChildNode = False
End Function

Public Function GetChildNode(ndList As Object, strNodeName As String) As Object
    Dim i As Long
    Dim ndName As String
    ndName = UCase$(Replace(strNodeName, " ", "_"))
        
    For i = 0 To ndList.Length - 1
        If UCase$(ndList(i).nodename) = ndName Then
            Set GetChildNode = ndList(i)
            Exit Function
        End If
    Next
    Set GetChildNode = Nothing
End Function
Public Function CountChildNodes(ndNode As Object, strNodeName As String) As Long
    Dim i As Long, ret As Long
    Dim ndName As String
    Dim x As Object
    Set x = ndNode.childNodes
    On Error GoTo ERRCN545454
    If x.Length = 0 Then
        CountChildNodes = 0
        Exit Function
    End If
    ndName = UCase$(Replace(strNodeName, " ", "_"))
    For i = 0 To x.Length - 1
        If UCase$(x(i).nodename) = ndName Then
            ret = ret + 1
        End If
    Next
    CountChildNodes = ret
Exit Function
ERRCN545454:
Err.Clear
        CountChildNodes = 0
        Exit Function

End Function
Public Function GetChildNodeText(ndNode As Object, strNodeName As String) As String
    Dim i As Long, ret As Long
    Dim ndName As String
    Dim x As Object
    Set x = ndNode.childNodes
    On Error GoTo ERRCN545454
    
    If x.Length = 0 Then
        GetChildNodeText = ""
        Exit Function
    End If
    ndName = UCase$(Replace(strNodeName, " ", "_"))
    For i = 0 To x.Length - 1
        Debug.Print x(i).nodename
        If UCase$(x(i).nodename) = ndName Then
            GetChildNodeText = x(i).Text
            Exit Function
        End If
    Next
    GetChildNodeText = ""
Exit Function
ERRCN545454:
Err.Clear
        GetChildNodeText = ""
        Exit Function

End Function
Public Sub SetChildNodeText(ndNode As Object, strNodeName As String, strNodeValue)
    Dim i As Long, ret As Long
    Dim ndName As String
    Dim x As Object
    Set x = ndNode.childNodes
    On Error GoTo ERRCN545454
    If x.Length = 0 Then
        Exit Sub
    End If
    ndName = Replace(strNodeName, " ", "_")
    For i = 0 To x.Length - 1
        If UCase$(x(i).nodename) = UCase$(ndName) Then
            x(i).Text = strNodeValue
            Exit Sub
        End If
    Next
    AddNode ndNode, "", ndName, strNodeValue
Exit Sub
ERRCN545454:
Err.Clear
        Exit Sub

End Sub
Public Function GetCNI(ndList As Object, strNodeName As String) As Long
    'Get Child Node Index
    Dim i As Long
    Dim ndName As String
    ndName = UCase$(Replace(strNodeName, " ", "_"))
    For i = 0 To ndList.Length - 1
        If UCase$(ndList(i).nodename) = ndName Then
            GetCNI = i
            Exit Function
        End If
    Next
    GetCNI = -1
End Function

Public Function GetNodeListByTag(strTag As String) As Object
    On Error GoTo ErrorHandler
    Set GetNodeListByTag = XMLDoc.getElementsByTagName(strTag)
    Exit Function
ErrorHandler:
    Set GetNodeListByTag = Nothing
End Function
Public Function QSCloneNode(ByVal ndNodeToClone As Object, Optional bDeep As Boolean = True) As Object
    Set QSCloneNode = ndNodeToClone.cloneNode(bDeep)
End Function
Public Function AppendChildNode(ByVal ndParentNode As Object, ByVal ndChildNode As Object) As Boolean
On Error GoTo ERRAPPCHILD54545
ndParentNode.appendChild ndChildNode
AppendChildNode = True
Exit Function
ERRAPPCHILD54545:
MsgBox Err.Description
Err.Clear
AppendChildNode = False
End Function
Public Function XMLAddNode(ByVal xmlRelative As Object, strXML As String) As Object
On Error GoTo ERRHDL
Dim nDoc As Object
Dim nd As Object
Set nDoc = InitializeObject(CONST_DOMDOCUMENT)
If Not nDoc.loadXML(strXML) Then
    Set XMLAddNode = Nothing
    Set nDoc = Nothing
    Exit Function
End If
Set nd = nDoc.documentElement()
Set XMLAddNode = xmlRelative.appendChild(nd)
Exit Function
ERRHDL:
Set XMLAddNode = Nothing
End Function
Public Function AddNode(ByVal xmlRelative As Object, ByVal strID As String, ByVal ndName As String, Optional ByVal vntNewNodeValue As Variant, Optional ByVal strNewAttributeName As String, Optional ByVal vntNewAttributeValue As Variant, Optional ByVal norRelationship As NodeRelationshipConstant = norCHILD) As Object
    On Error GoTo ErrorHandler
    Dim XMLNode As Object
    Dim xmlChildNode As Object
    Dim xmlAttributeNode As Object
    Dim strNewNodeName
    If InStr(ndName, " ") > 0 Then
        strNewNodeName = Replace(ndName, " ", "_")
    Else
        strNewNodeName = ndName
    End If
    Select Case norRelationship
        Case norCHILD:
            Set XMLNode = xmlRelative
        Case norSIBLING:
            Set XMLNode = xmlRelative.ParentNode
    End Select
   
    Set xmlChildNode = XMLDoc.createElement(strNewNodeName)
            
    If Not IsMissing(vntNewNodeValue) Then
        xmlChildNode.Text = vntNewNodeValue
    End If

    Set xmlAttributeNode = XMLNode.appendChild(xmlChildNode)
    If Len(strID) > 0 Then
        xmlAttributeNode.SetAttribute "id", strID
    End If
    If Not IsMissing(strNewAttributeName) And Not IsMissing(vntNewAttributeValue) Then
        xmlAttributeNode.SetAttribute strNewAttributeName, vntNewAttributeValue
    End If
            
    Set xmlAttributeNode = Nothing
    'Set xmlChildNode = Nothing
    Set XMLNode = Nothing
    Set AddNode = xmlChildNode
    Exit Function
ErrorHandler:
    Set xmlAttributeNode = Nothing
    Set xmlChildNode = Nothing
    Set XMLNode = Nothing
    Set AddNode = xmlChildNode
End Function

'.------------------------------------------------------------------------------
'.  Function  : Public Function InsertNode
'.
'.  Parameters: ByVal strQueryString As String
'.              ByVal strNewNodeName As String
'.              Optional ByVal vntNewNodeValue As Variant
'.              Optional ByVal strNewAttributeName As String
'.              Optional ByVal vntNewAttributeValue As Variant
'.              Optional ByVal norRelationship As NodeRelationshipConstant = norCHILD
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Lets you insert a new Child OR SIBLING node, along
'.              with ONE attribute / attribute value pair
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function InsertNode(ByVal strQueryString As String, ByVal strNewNodeName As String, Optional ByVal vntNewNodeValue As Variant, Optional ByVal strNewAttributeName As String, Optional ByVal vntNewAttributeValue As Variant, Optional ByVal norRelationship As NodeRelationshipConstant = norCHILD) As Boolean
    On Error GoTo ErrorHandler
    
    Dim XMLNode As Object
    Dim xmlChildNode As Object
    Dim xmlAttributeNode As Object
    
    Select Case norRelationship
        Case norCHILD:
            Set XMLNode = XMLDoc.documentElement.selectSingleNode(strQueryString)
        Case norSIBLING:
            Set XMLNode = XMLDoc.documentElement.selectSingleNode(strQueryString).ParentNode
    End Select
   
    Set xmlChildNode = XMLDoc.createElement(strNewNodeName)
            
    If Not IsMissing(vntNewNodeValue) Then
        xmlChildNode.Text = vntNewNodeValue
    End If

    Set xmlAttributeNode = XMLNode.appendChild(xmlChildNode)
            
    If Not IsMissing(strNewAttributeName) And Not IsMissing(vntNewAttributeValue) Then
        xmlAttributeNode.SetAttribute strNewAttributeName, vntNewAttributeValue
    End If
            
    Set xmlAttributeNode = Nothing
    Set xmlChildNode = Nothing
    Set XMLNode = Nothing
    InsertNode = True
    Exit Function
ErrorHandler:
    InsertNode = False
    Set xmlAttributeNode = Nothing
    Set xmlChildNode = Nothing
    Set XMLNode = Nothing
End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function DeleteAttribute
'.
'.  Parameters: ByVal strQueryString As String
'.              ByVal strAttributeName As String
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Delete attribute / attribute value pair form supplied node
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function DeleteAttribute(ByVal strQueryString As String, ByVal strAttributeName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim xmlAttributeNode As Object
    Set xmlAttributeNode = XMLDoc.documentElement.selectSingleNode(strQueryString).Attributes.removeNamedItem(strAttributeName)
    
    Set xmlAttributeNode = Nothing
    DeleteAttribute = True
    Exit Function
ErrorHandler:
    DeleteAttribute = False
    Set xmlAttributeNode = Nothing
End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function Reparse
'.
'.  Parameters: -
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Recheck if current XML document is still valid XML
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function Reparse() As Boolean
    On Error GoTo ErrorHandler
    
    ' Create the DOMDocument
    Dim xmlDocReparse As Object
    Set xmlDocReparse = InitializeObject(CONST_DOMDOCUMENT)
    
    If xmlDocReparse.loadXML(XMLDoc.XML) = True Then
        Reparse = True
    Else
        Reparse = False
    End If
    
    Set xmlDocReparse = Nothing
    
    Exit Function
ErrorHandler:
    Reparse = False
    Set xmlDocReparse = Nothing
End Function
'.------------------------------------------------------------------------------
'.  Function  : Public Function Save
'.
'.  Parameters: ByVal strFileName As String
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Save current XML document to file
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Public Function Save(ByVal strFilename As String) As Boolean
    On Error GoTo ErrorHandler
    
    XMLDoc.Save (strFilename)
    Save = True
    
    Exit Function
ErrorHandler:
    Save = False
End Function

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' METHODS (PUBLIC) - END
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'-------------------------------------------------------------------------------

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' METHODS (PRIVATE) - START
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'.------------------------------------------------------------------------------
'.  Function  : Private Function InitializeObject
'.
'.  Parameters: strObjectName As String
'.
'.  Returns   : Object (NOTHING = Error)
'.
'.  Comments  : Try to initialize an instance of supplied MS XML Parser class /
'.              object name combination
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Private Function InitializeObject(strObjectName As String) As Object
    On Error GoTo ErrorHandler
    Dim objDummy As Object
    Dim strClassName As String

    Select Case mpavCurrentParserVersion
        Case pavV20:
             strClassName = "Msxml"
        Case pavV26, pavV30, pavV40:
            strClassName = "Msxml2"
        Case Else: ' currently not allowed = Error
            GoTo ErrorHandler
    End Select
    
    Dim strObjectToCreate As String
    strObjectToCreate = strClassName & "." & strObjectName
    
    Set objDummy = CreateObject(strObjectToCreate)
    
    Set InitializeObject = objDummy

    Set objDummy = Nothing
    Exit Function
ErrorHandler:
    If Err.Number = 429 Then
        Debug.Print "Can't CreateObject(" & strObjectToCreate & ")"
    End If
   
    Set InitializeObject = Nothing
    Set objDummy = Nothing
End Function
'.------------------------------------------------------------------------------
'.  Function  : Private Function InvokeParser
'.
'.  Parameters: pavDesiredVersion As ParserVersionConstant
'.
'.  Returns   : Boolean (Success = True / Error = False)
'.
'.  Comments  : Tries to initialize the supplied MS XML Parser version
'.              Used by the .AutoDetectParser function
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Private Function InvokeParser(pavDesiredVersion As ParserVersionConstant) As Boolean
    Dim objDummy As Object
    Dim strInvokation As String
    
    Select Case pavDesiredVersion
        Case pavV40:
            strInvokation = CONST_INVOKE_DOMDOCUMENT_40
        Case pavV30:
            strInvokation = CONST_INVOKE_DOMDOCUMENT_30
        Case pavV26:
            strInvokation = CONST_INVOKE_DOMDOCUMENT_26
        Case pavV20:
            strInvokation = CONST_INVOKE_DOMDOCUMENT_20
    End Select
    
    On Error Resume Next
    Set objDummy = CreateObject(strInvokation)
    If Err.Number = 0 Then
        InvokeParser = True
    Else
        InvokeParser = False
    End If
    Set objDummy = Nothing
End Function
'.------------------------------------------------------------------------------
'.  Function  : Private Function AutoDetectParser
'.
'.  Parameters: -
'.
'.  Returns   : As ParserVersionConstant (0 = Error)
'.
'.  Comments  : Check which versions of MSXML are available
'.              and return the latest version
'.
'.              VERSION         CAN INITIALIZE THE FOLLOWING OBJECTS
'.              -------         ------------------------------------
'.              Microsoft XML
'.              v 2.0
'.              MSXML.dll  =    Msxml.DOMDocument
'.
'.              v 2.6
'.              MSXML2.dll =    Msxml2.DOMDocument
'.                              Msxml2.DOMDocument.2.6
'.
'.              v3.0
'.              MSXML3.dll =    Msxml2.DOMDocument
'.                              Msxml2.DOMDocument.2.6
'.                              Msxml2.DOMDocument.3.0
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Private Function AutoDetectParser() As ParserVersionConstant
    ' Check latest versions first
    ' check for v40
    If InvokeParser(pavV40) = True Then
        AutoDetectParser = pavV40
        Exit Function
    End If
    ' check for v30
    If InvokeParser(pavV30) = True Then
        AutoDetectParser = pavV30
        Exit Function
    End If
    
    ' check for v26
    If InvokeParser(pavV26) = True Then
        AutoDetectParser = pavV26
        Exit Function
    End If
       
    ' check for v20
    If InvokeParser(pavV20) = True Then
        AutoDetectParser = pavV20
        Exit Function
    End If
    
    ' No XML Parser detected
    AutoDetectParser = 0
End Function
Public Function GetPageRootNode(Optional MyPageNum As Long = 1) As Object
Dim ndList As Object
Dim rootnode As Object

    'First Get The Prompts
    Set rootnode = GetRootElement()
    Set ndList = rootnode.childNodes
    If ndList.Length >= MyPageNum Then
        Set GetPageRootNode = rootnode.childNodes(MyPageNum - 1)
        Set ndList = Nothing
        Set rootnode = Nothing
    Exit Function
    End If
    Set GetPageRootNode = Nothing

End Function

'.------------------------------------------------------------------------------
'.  Function  : Private Sub DisplayParseError
'.
'.  Parameters: -
'.
'.  Returns   : -
'.
'.  Comments  : Display the latest ParseError from MS XML Parser
'.
'.  Author    : Rajneesh Noonia - 09 Jan 2007
'.  Changed   : -
'.------------------------------------------------------------------------------
Private Sub DisplayParseError()
    Dim xmlParseError As Object
    Set xmlParseError = XMLDoc.parseError
    
    MsgBox "Error loading XML" & vbCrLf & vbCrLf & _
        xmlParseError.errorCode & vbCrLf & _
        xmlParseError.reason & "Line " & xmlParseError.Line & _
        ", Position " & xmlParseError.linepos & vbCrLf & vbCrLf & _
        xmlParseError.srcText & vbCrLf _
        , vbExclamation, "XML Parse Error"
    
    Set xmlParseError = Nothing
End Sub

Private Sub Class_Terminate()
    Set XMLDoc = Nothing
End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' METHODS (PRIVATE) - END
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<










