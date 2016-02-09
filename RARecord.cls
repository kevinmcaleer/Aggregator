VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RARecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public tableName As String ' the Friendly name for the assessment
Public fileName As String ' the actual file name (without any path)
Public path As String ' the full path and filename

Private Sub Class_Initialize()
' Ensure the variables are wiped clean, ready for use
tableName = ""
fileName = ""
path = ""
End Sub
