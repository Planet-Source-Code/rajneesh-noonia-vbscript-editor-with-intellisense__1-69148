VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Module"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Variables"
Option Explicit
Implements VbScriptEditor.PersistPropertyBag

Public Name As String
Public Key As String
Private mvarVariables As Variables

Public Property Get Variables() As Variables
    If mvarVariables Is Nothing Then
        Set mvarVariables = New Variables
    End If
    Set Variables = mvarVariables
End Property


Public Property Set Variables(vData As Variables)
    Set mvarVariables = vData
End Property

Private Sub Class_Terminate()
    Set mvarVariables = Nothing
End Sub

Private Sub PersistPropertyBag_Load(ByVal PropBag As VbScriptEditor.cPropertyBag)
    Name = PropBag.ReadProperty("Name", "")
    Key = PropBag.ReadProperty("Key", "")
    Call PropBag.RestoreState(Me.Variables, "Variables")
End Sub

Private Sub PersistPropertyBag_Save(ByVal PropBag As VbScriptEditor.cPropertyBag)
    Call PropBag.WriteProperty("Name", Name, "")
    Call PropBag.WriteProperty("Key", Key, "")
    Call PropBag.SaveState(Me.Variables, "Variables")
End Sub
