VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTypeLibInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum iTypeKind
    IUnknown
    iLibrary
    iModule
    iClass
    iFunction
    iType
    iEnum
    iConstant
    iProperty
    iEvent
    iDataType
End Enum

Public Name As String
Public ClassName As String
Public CLSID As String
Public Ver As String
Public Path As String
Public Exists As Boolean
Public ProgID As String
Public TypeKind As iTypeKind
Public Tag As String

