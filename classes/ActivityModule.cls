VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActivityModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Interface to activity objects

Public Function populate(inst_file As InstFile)
Public Function attach(act As Activity)
Public Function search(query As String) As Variant
Public Function searchcoll(query As String) As Collection
Public Function str() As String
