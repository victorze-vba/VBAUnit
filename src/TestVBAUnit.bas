Attribute VB_Name = "TestVBAUnit"
Option Explicit

' AssertEquals
Sub OkAssertEquals()
    Dim Test As New VBAUnit
    
    With Test.It("Test Ok AssertEquals")
        .AssertEquals 5, 5
        .AssertEquals "VBA", "VBA"
        .AssertEquals Now, Now
    End With
End Sub

Sub FailAssertEquals()
    Dim Test As New VBAUnit
    
    With Test.It("Test Fail AssertEquals")
        .AssertEquals 5, 6
        .AssertEquals "VBA", "VBA Excel"
        .AssertEquals Now + 1, Now
    End With
End Sub

' AssertNotEquals
Sub OkAssertNotEquals()
    Dim Test As New VBAUnit
    
    With Test.It("Test Ok AssertNotEquals")
        .AssertNotEquals 5, 6
        .AssertNotEquals "VBA", "VBA Excel"
        .AssertNotEquals Now + 1, Now
    End With
End Sub

Sub FailAssertNotEquals()
    Dim Test As New VBAUnit
    
    With Test.It("Test Fail AssertNotEquals")
        .AssertNotEquals 5, 5
        .AssertNotEquals "VBA", "VBA"
        .AssertNotEquals Now, Now
    End With
End Sub

' AssertTrue
Sub OkAssertTrue()
    Dim Test As New VBAUnit
    
    With Test.It("Test Ok AssertTrue")
        .AssertTrue True
        .AssertTrue 5 = 5
    End With
End Sub

Sub FailAssertTrue()
    Dim Test As New VBAUnit
    
     With Test.It("Test Fail AssertEquals")
        .AssertEquals 5, 6
        .AssertEquals "VBA", "VBA Excel"
        .AssertEquals Now + 1, Now
    End With
End Sub

' AssertFalse
Sub OkAssertFalse()
    Dim Test As New VBAUnit
    
    With Test.It("Test Ok AssertFalse")
        .AssertFalse False
        .AssertFalse 5 = 6
    End With
End Sub

Sub FailAssertFalse()
    Dim Test As New VBAUnit
    
    With Test.It("Test Fail AssertFalse")
        .AssertFalse True
        .AssertFalse 5 = 5
    End With
End Sub

' AssertSame
Sub OkAssertSame()
    Dim Test As New VBAUnit
    Dim Obj1 As New VBAUnit
    Dim Obj2 As VBAUnit
    
    Set Obj2 = Obj1
    
    With Test.It("Test Ok AssertSame")
        .AssertSame Obj1, Obj2
        .AssertSame ActiveWorkbook, ActiveWorkbook
    End With
End Sub

Sub FailAssertSame()
    Dim Test As New VBAUnit
    Dim Obj1 As New VBAUnit
    Dim Obj2 As New VBAUnit
    Dim Obj3 As Range
    
    Set Obj3 = Cells(1, 1)
    
    With Test.It("Test Fail AssertSame")
        .AssertSame Obj1, Obj2
        .AssertSame Obj2, Obj3
    End With
End Sub

' AssertNotSame
Sub OkAssertNotSame()
    Dim Test As New VBAUnit
    Dim Obj1 As New VBAUnit
    Dim Obj2 As New VBAUnit
    Dim Obj3 As Range
    
    Set Obj3 = Cells(1, 1)
    
    With Test.It("Test Ok AssertNotSame")
        .AssertNotSame Obj1, Obj2
        .AssertNotSame Obj2, Obj3
    End With
End Sub

Sub FailAssertNotSame()
    Dim Test As New VBAUnit
    Dim Obj1 As New VBAUnit
    Dim Obj2 As VBAUnit
    
    Set Obj2 = Obj1
    
    With Test.It("Test Fail AssertNotSame")
        .AssertNotSame Obj1, Obj2
        .AssertNotSame ActiveWorkbook, ActiveWorkbook
    End With
End Sub

' AssertNull
Sub OkAssertNull()
    Dim Test As New VBAUnit
    Dim Value As Variant
    Dim Obj As VBAUnit
    
    With Test.It("Test Ok AssertNull")
        .AssertNull Value
        .AssertNull Obj
    End With
End Sub

Sub FailAssertNull()
    Dim Test As New VBAUnit
    Dim Value As Variant
    Dim Obj As New VBAUnit
    
    Value = 23
    
    With Test.It("Test Fail AssertNull")
        .AssertNull Value
        .AssertNull Obj
    End With
End Sub

' AssertNotNull
Sub OkAssertNotNull()
    Dim Test As New VBAUnit
    Dim Value As Variant
    Dim Obj As New VBAUnit

    Value = 23

    With Test.It("Test Ok AssertNotNull")
        .AssertNotNull Value
        .AssertNotNull Obj
    End With
End Sub

Sub FailAssertNotNull()
    Dim Test As New VBAUnit
    Dim Value As Variant
    Dim Obj As VBAUnit

    With Test.It("Test Fail AssertNotNull")
        .AssertNotNull Value
        .AssertNotNull Obj
    End With
End Sub

'MultipleAssert
Sub OKMultipleAssert()
    Dim Test As New VBAUnit

    With Test.It("Test Fail AssertNull")
        .AssertEquals 5, 5
        .AssertTrue True
    End With
End Sub

Sub FailMultipleAssert()
    Dim Test As New VBAUnit
    
    With Test.It("Test Fail AssertNull")
        .AssertEquals 5, 6
        .AssertTrue False
    End With
End Sub

' Multiple fail
Sub TestVBAUnit()
    Dim Test As New VBAUnit
    
    With Test.It("Test Fail AssertEquals")
        .AssertEquals 5, 6
        .AssertEquals "VBA", "VBA Excel"
        .AssertEquals Now + 1, Now
    End With
    
    With Test.It("Test Fail AssertNotEquals")
        .AssertNotEquals 5, 5
        .AssertNotEquals "VBA", "VBA"
        .AssertNotEquals Now, Now
    End With

    Dim Obj1 As New VBAUnit
    Dim Obj2 As New VBAUnit
    Dim Obj3 As Range
    
    Set Obj3 = Cells(1, 1)
    
    With Test.It("Test Fail AssertSame")
        .AssertSame Obj1, Obj2
        .AssertSame Obj2, Obj3
    End With
End Sub
