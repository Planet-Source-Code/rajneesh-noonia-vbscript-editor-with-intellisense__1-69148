VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PersistPropertyBag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Interface to be implemented by custom task and all persistable Objects
'Desc AMS definition of Automation enabled persistence into a property bag. This implementation Also support persistence of generic blobs or objects.
'Member Reads a named property from the property bag. This implementation does not support generic blobs or objects as property values.

Public Sub Load(ByVal PropBag As VbScriptEditor.cPropertyBag)
'
End Sub

Public Sub Save(ByVal PropBag As VbScriptEditor.cPropertyBag)
'
End Sub

