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
Const ExpectedCompletionDate = 6 ' Column 5 = Expected Completion Date
Const EvidenceExpected = 7 ' Column 6 = Excidence Expected
Const EvidenceDate = 8 ' Column 7 = Evidence Date
Const MilestoneRef = 9 ' Column 8 = Milestone Reference
Const RiskIssue = 10 ' Column 9 = Risk / Issue
Const Rational = 11 ' Column 10 = Rational
Const impact = 12 ' Column 11 = Impact
Const Probabililty = 13 ' Column 12 = Probability
Const Driver = 14 ' Column 13 = Driver
Const RtG = 15 ' column 14 = Expected Return to Green Date

Dim QResponse As QuestionResponse ' Stores the Answers (1 row = 1 Qresponse)
Dim ResponseList As QuestionResponse ' stores a list of all the questions
Dim myRow As ListRow ' user to go through each row in the Question table
Dim count As Integer ' used to count the current row number

count = 2 ' row counter
' read each line in and add to appropriate questionresponse
' need to add an array of 168 questions

For Each myRow In assList.ListRows
    
    Set QResponse = New QuestionResponse ' create a new instance of Question Responses
    ' read a line in
    QResponse.Question = assList.Range(count, SubQuestion)
    QResponse.RAG = assList.Range(count, RAGStatus)
    QResponse.Driver = assList.Range(count, Driver)
    QResponse.EvidenceDate = assList.Range(count, EvidenceDate)
    QResponse.EvidenceExpected = assList.Range(count, EvidenceExpected)
    QResponse.ExpectedCompletionDate = assList.Range(count, ExpectedCompletionDate)
    QResponse.impact = assList.Range(count, impact)
    QResponse.MilestoneRef = assList.Range(count, MilestoneRef)
    QResponse.Probability = assList.Range(count, Probability)
    QResponse.Rational = assList.Range(count, Rational)
    QResponse.ReturnToGreenDate = assList.Range(count, RtG)
    QResponse.RiskIssue = assList.Range(count, RiskIssue)
    QResponse.Status = assList.Range(count, CurrentStatus)
    'add to combined response
    
    'check if its not the first line
    'if count <> 1 then
    
    'compare RAG
    If QResponse.RAGValue > Combined(count).RAGValue Then
    Combined(count).RAG = QResponse.RAG
    End If

    ' Check if is a subquestion
    ' Row 2 Q1
    ' Row 14 Q2
    ' Row 36 Q3
    ' Row 59 Q4
    ' Row 63 Q5
    ' Row 75 Q6
    ' Row 85 Q7
    ' Row 94 Q8
    ' Row 106 Q9
    ' Row 121 Q10
    ' Row 134 Q11
    ' Row 147 Q12
    ' Row 169 Q13
    If count <> 2 Or count <> 14 Or count <> 36 Or count <> 59 _
    Or count <> 63 Or count <> 75 Or count <> 85 Or count <> 94 _
    Or count <> 106 Or count <> 121 Or count <> 134 Or count <> 147 _
    Or count <> 169 Then
    Combined(count).Question = QResponse.Question
    Combined(count).Status = Combined(count).Status & assName & ": " & QResponse.Status & vbLf
    'concatenate statements
    Combined(count).EvidenceDate = Combined(count).EvidenceDate & assName & ": " & QResponse.EvidenceDate & vbLf
    Combined(count).Driver = Combined(count).Driver & assName & ": " & QResponse.Driver & vbLf
    Combined(count).EvidenceExpected = Combined(count).EvidenceExpected & assName & ": " & QResponse.EvidenceExpected & vbLf
    Combined(count).impact = Combined(count).impact & assName & ": " & QResponse.impact & vbLf
    Combined(count).MilestoneRef = Combined(count).MilestoneRef & assName & ": " & QResponse.MilestoneRef & vbLf
    Combined(count).Probability = Combined(count).Probability & assName & ": " & QResponse.Probability & vbLf
    Combined(count).Rational = Combined(count).Rational & assName & ": " & QResponse.Rational & vbLf
    Combined(count).ReturnToGreenDate = Combined(count).ReturnToGreenDate & assName & ": " & QResponse.ReturnToGreenDate & vbLf
    Combined(count).RiskIssue = Combined(count).RiskIssue & assName & ": " & QResponse.RiskIssue & vbLf
    End If
    
    'end if
    count = count + 1
    
    Next myRow
    'combined(2).show ' show results durind debug
End Sub

Public Sub InitializeDefaultValues()
Set Responses = New Collection
Set Summaries = New Collection
NoAssessments = 0
End Sub
