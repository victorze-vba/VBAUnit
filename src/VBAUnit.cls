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
        PrintFailDescription
        PrintExpected Expected
        PrintBut Actual
    End If

    AddRun
    Set AssertEquals = Me
End Function

Public Function AssertNotEquals(Unexpected As Variant, Actual As Variant) As VBAUnit
    If Unexpected = Actual Then
        PrintFailDescription
        OutputMessage.Add Chr(9) & "Unexpected: " & Unexpected
        OutputMessage.Add Chr(9) & "       But: " & Actual
    End If
    
    AddRun
    Set AssertNotEquals = Me
End Function

Public Function AssertTrue(Condition As Boolean) As VBAUnit
    If Condition = False Then
        PrintFailDescription
        PrintExpected "True"
        PrintBut Condition
    End If
    
    AddRun
    Set AssertTrue = Me
End Function

Public Function AssertFalse(Condition As Boolean) As VBAUnit
    If Condition = True Then
        PrintFailDescription
        PrintExpected "False"
        PrintBut Condition
    End If
    
    AddRun
    Set AssertFalse = Me
End Function

Public Function AssertSame(Expected As Variant, Actual As Variant) As VBAUnit
    If Not Expected Is Actual Then
        PrintFailDescription
        
        If TypeName(Expected) = TypeName(Actual) Then
            PrintExpected "object " & TypeName(Expected)
            PrintBut "other object " & TypeName(Actual)
        Else
            PrintExpected "object " & TypeName(Expected)
            PrintBut "object " & TypeName(Actual)
        End If
    End If
    
    AddRun
    Set AssertSame = Me
End Function

Public Function AssertNotSame(Unexpected As Variant, Actual As Variant) As VBAUnit
    If Unexpected Is Actual Then
        PrintFailDescription
        OutputMessage.Add Chr(9) & "Unexpected: " & "object " & TypeName(Unexpected)
        OutputMessage.Add Chr(9) & "       But: " & "same object " & TypeName(Actual)
    End If
    
    AddRun
    Set AssertNotSame = Me
End Function

Public Function AssertNull(Value As Variant) As VBAUnit
    If IsObject(Value) Then
        If TypeName(Value) <> "Nothing" Then
            PrintFailDescription
            PrintExpected "Nothing"
            PrintBut TypeName(Value)
        End If
    Else
        If TypeName(Value) <> "Empty" Then
            PrintFailDescription
            PrintExpected "Empty"
            PrintBut TypeName(Value)
        End If
    End If
    
    AddRun
    Set AssertNull = Me
End Function

Public Function AssertNotNull(Value As Variant) As VBAUnit
    If IsObject(Value) Then
        If TypeName(Value) = "Nothing" Then
            PrintFailDescription
            PrintExpected "Not Nothing"
            PrintBut TypeName(Value)
        End If
    Else
        If TypeName(Value) = "Empty" Then
            PrintFailDescription
            PrintExpected "Not Empty"
            PrintBut TypeName(Value)
        End If
    End If
    
    AddRun
    Set AssertNotNull = Me
End Function

Sub AddFailure()
    CountFailures = CountFailures + 1
End Sub

Sub AddRun()
    CountRuns = CountRuns + 1
End Sub

Public Function It(Description As String) As VBAUnit
    pDescription = Description
    
    Set It = Me
End Function

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
