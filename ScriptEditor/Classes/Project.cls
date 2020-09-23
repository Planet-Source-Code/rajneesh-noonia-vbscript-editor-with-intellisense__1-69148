VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Project"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Modules"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"Refrences"
Option Explicit
Implements VbScriptEditor.PersistPropertyBag


Private mvarRefrences As Refrences
Private mvarModules As Modules

Public Name As String

Public Property Get XML() As String
    Dim pcPropertyBag As VbScriptEditor.cPropertyBag
    Dim pVarContents As Variant
    Set pcPropertyBag = New VbScriptEditor.cPropertyBag
    Call pcPropertyBag.SaveState(Me, "Project")
    pVarContents = pcPropertyBag.Contents
    XML = pVarContents
    Set pcPropertyBag = Nothing
End Property

Public Property Let XML(ByVal pXMl As String)
    Dim pcPropertyBag As VbScriptEditor.cPropertyBag
    Dim pBytesContents() As Byte
    Dim pVarContents As Variant
    On Error GoTo ErrorTrap
    Set pcPropertyBag = New VbScriptEditor.cPropertyBag
    
    pVarContents = pXMl
    pBytesContents = pVarContents
    pcPropertyBag.Contents = pBytesContents
    Call pcPropertyBag.RestoreState(Me, "Project")
    Set pcPropertyBag = Nothing
    Exit Property
ErrorTrap:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property

Public Property Get Modules() As Modules
    If mvarModules Is Nothing Then
        Set mvarModules = New Modules
    End If
    Set Modules = mvarModules
End Property


Public Property Set Modules(vData As Modules)
    Set mvarModules = vData
End Property


Public Property Get Refrences() As Refrences
    If mvarRefrences Is Nothing Then
        Set mvarRefrences = New Refrences
    End If
    Set Refrences = mvarRefrences
End Property


Public Property Set Refrences(vData As Refrences)
    Set mvarRefrences = vData
End Property

Private Sub Class_Terminate()
  Set mvarModules = Nothing
    Set mvarRefrences = Nothing
End Sub


Private Sub PersistPropertyBag_Load(ByVal PropBag As VbScriptEditor.cPropertyBag)
    Name = PropBag.ReadProperty("Name", "")
    Call PropBag.RestoreState(Me.Modules, "Modules")
    Call PropBag.RestoreState(Me.Refrences, "Refrences")
End Sub

Private Sub PersistPropertyBag_Save(ByVal PropBag As VbScriptEditor.cPropertyBag)
    Call PropBag.WriteProperty("Name", Name, "")
    Call PropBag.SaveState(Me.Modules, "Modules")
    Call PropBag.SaveState(Me.Refrences, "Refrences")
End Sub
