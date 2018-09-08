VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBAUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pDescription As String

Private LatestDescription As String

Private OutputMessage As New Collection

Private CountFailures As Integer

Private CountRuns As Integer

Public Function AssertEquals(Expected As Variant, Actual As Variant) As VBAUnit
    If Expected <> Actual Then
        PrintMessage Expected, Actual
    End If

    AddRun
    Set AssertEquals = Me
End Function

Public Function AssertNotEquals(Unexpected As Variant, Actual As Variant) As VBAUnit
    If Unexpected = Actual Then
        PrintMessageUnexpected Unexpected, Actual
    End If
    
    AddRun
    Set AssertNotEquals = Me
End Function

Public Function AssertTrue(Condition As Boolean) As VBAUnit
    If Condition = False Then
        PrintMessage "True", Condition
    End If
    
    AddRun
    Set AssertTrue = Me
End Function

Public Function AssertFalse(Condition As Boolean) As VBAUnit
    If Condition = True Then
        PrintMessage "False", Condition
    End If
    
    AddRun
    Set AssertFalse = Me
End Function

Public Function AssertSame(Expected As Variant, Actual As Variant) As VBAUnit
    If Not Expected Is Actual Then
        If TypeName(Expected) = TypeName(Actual) Then
            PrintMessage "object " & TypeName(Expected), "other object " & TypeName(Actual)
        Else
            PrintMessage "object " & TypeName(Expected), "object " & TypeName(Actual)
        End If
    End If
    
    AddRun
    Set AssertSame = Me
End Function

Public Function AssertNotSame(Unexpected As Variant, Actual As Variant) As VBAUnit
    If Unexpected Is Actual Then
        PrintMessageUnexpected "object " & TypeName(Unexpected), "same object " & TypeName(Actual)
    End If
    
    AddRun
    Set AssertNotSame = Me
End Function

Public Function AssertNull(Value As Variant) As VBAUnit
    If IsObject(Value) Then
        If TypeName(Value) <> "Nothing" Then
            PrintMessage "Nothing", TypeName(Value)
        End If
    Else
        If TypeName(Value) <> "Empty" Then
            PrintMessage "Empty", TypeName(Value)
        End If
    End If
    
    AddRun
    Set AssertNull = Me
End Function

Public Function AssertNotNull(Value As Variant) As VBAUnit
    If IsObject(Value) Then
        If TypeName(Value) = "Nothing" Then
            PrintMessage "Not Nothing", TypeName(Value)
        End If
    Else
        If TypeName(Value) = "Empty" Then
            PrintMessage "Not Empty", TypeName(Value)
        End If
    End If
    
    AddRun
    Set AssertNotNull = Me
End Function

Public Function It(Description As String) As VBAUnit
    pDescription = Description
    
    Set It = Me
End Function

Private Sub PrintMessage(Expected As Variant, But As Variant)
    PrintFailDescription
    PrintExpected Expected
    PrintBut But
End Sub

Private Sub PrintMessageUnexpected(Unexpected As Variant, But As Variant)
    PrintFailDescription
    PrintUnexpected Unexpected
    PrintButUnexpected But
End Sub

Private Sub PrintFailDescription()
    AddFailure

    If LatestDescription <> pDescription Then
        OutputMessage.Add "--------------------------------------------------------------"
        OutputMessage.Add "FAIL: " & pDescription
        OutputMessage.Add "--------------------------------------------------------------"
        LatestDescription = pDescription
    Else
        OutputMessage.Add Chr(9) & "----------------------------------------------------------"
    End If
End Sub

Private Sub PrintExpected(Value As Variant)
    OutputMessage.Add Chr(9) & "Expected: " & Value
End Sub

Private Sub PrintBut(Value As Variant)
    OutputMessage.Add Chr(9) & "     But: " & Value
End Sub

Private Sub PrintUnexpected(Value As Variant)
    OutputMessage.Add Chr(9) & "Unexpected: " & Value
End Sub

Private Sub PrintButUnexpected(Value As Variant)
    OutputMessage.Add Chr(9) & "       But: " & Value
End Sub

Private Sub AddFailure()
    CountFailures = CountFailures + 1
End Sub

Private Sub AddRun()
    CountRuns = CountRuns + 1
End Sub

Private Sub Class_Terminate()
    PrintSummary
    PrintOutputMessage
End Sub

Private Sub PrintSummary()
    If CountRuns > 0 Then
        Debug.Print "Ran " & CountRuns & " test"
    End If
    
    If CountFailures > 0 Then
        Debug.Print "FAILED (failures=" & CountFailures & ")" & Chr(13)
    End If
    
    If CountFailures = 0 And CountRuns > 0 Then
        Debug.Print "OK" & Chr(13)
    End If
End Sub

Private Sub PrintOutputMessage()
    Dim Line As Variant
    
    For Each Line In OutputMessage
        Debug.Print Line
    Next Line
End Sub
