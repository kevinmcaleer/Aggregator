VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Assessment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Responses As Collection ' Collection of Summaries
Private Summaries As Collection ' Collection of QuestionResponses
Private NoAssessments As Integer
Private assList As ListObject

Public Property Let List(ByRef myList As ListObject)
    Set assList = myList
End Property

Public Property Get List() As ListObject
Set List = assList
End Property

Public Property Let Response(ByRef myResponse As QuestionResponse)
Set Responses = myResponse
End Property

Public Property Get Response() As QuestionResponse
Response = Responses
End Property

Public Property Let NoReponses(ByRef myNo As Integer)
NoAssessments = myNo
End Property

Public Property Get NoResponse() As Integer
NoResponse = NoAssessment
End Property

Private Sub Class_Initialize()
Set Responses = New Collection
Set Summaries = New Collection
NoAssessments = 0
End Sub

Public Sub UpdatesAssessment()
NoAssessments = NoAssessments + 1
Application.StatusBar = "Number of Assessments: " & NoAssessments
End Sub

Public Sub ReadAssessment(assName As String)
Const SubQuestion = 2 ' Column 1 = Subquestion
Const Evidence = 3 ' Column 2 = Evidence
Const RAGStatus = 4 ' Column 3 = RAG Status
Const CurrentStatus = 5 ' Column 4 = Current Status

End Sub
