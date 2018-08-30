# VBAUnit
## VBAUnit Unit Testing Library for VBA

### Example
```vb
Sub TestApp()
    Dim Test As New VBAUnit

    With Test.It("Description")
        .AssertEquals 5, 5
        .AssertEquals "VBA", "VBA"
    End With
    ' output:
    ' Ran 2 test
    ' OK


    With Test.It("Description")
        .AssertEquals 5, 6
        .AssertEquals "VBA", "VBA Excel"
    End With
    ' output:
    ' Ran 2 test
    ' FAILED (failures=2)
    '
    ' ----------------------------------------------------
    ' FAIL: Description
    ' ----------------------------------------------------
    '     Expected: 5
    '          But: 6
    '     ------------------------------------------------
    '     Expected: VBA
    '          But: VBA Excel
End Sub
```

