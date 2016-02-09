VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QuestionResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Question As String
Public RAG As String
Public Status As String
Public ExpectedCompletionDate As String
Public EvidenceExpected As String
Public EvidenceDate As String
Public MilestoneRef As String
Public RiskIssue As String
Public Rational As String
Public impact As String
Public Probability As String
Public Driver As String
Public ReturnToGreenDate As String
Private RAGNo As Integer ' the number value of the RAG status

Private Sub Class_Initialize()
Question = ""
RAG = ""
Status = ""
ExpectedCompletionDate = ""
EvidenceExpected = ""
EvidenceDate = ""
MilestoneRef = ""
RiskIssue = ""
Rational = ""
impact = ""
Probabilty = ""
Driver = ""
ReturnToGreenDate = ""
RAGNo = 0
End Sub

Public Sub show()
MsgBox ("Question - " & Question & vbLf & "RAG - " & RAG & vbLf & "Summary: " & Summary)
End Sub

Public Property Get RAGValue() As Integer
Select Case UCase(RAG)
    Case "RED"
        RAGValue = 6
    Case "AMBER"
        RAGValue = 5
    Case "GREEN"
        RAGValue = 4
    Case "COMPLETE"
        RAGValue = 3
    Case "NOT STARTED YET" ' need to check this value
        RAGValue = 2
    Case "NOT IN SCOPE"
        RAGValue = 1
End Select
RAGNo = RAGValue
End Property
