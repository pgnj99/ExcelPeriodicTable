VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPercent 
   Caption         =   "Percent Composition"
   ClientHeight    =   9720.001
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   4740
   OleObjectBlob   =   "frmPercent.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPercent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnEle1_Click()
Dim E1, E2, E3, E4, Q1, Q2, Q3, Q4, S 'units to be entered; each E is an element, each Q is a quantity, and S is a symbol
Dim arr1(1 To 119), arr2(1 To 119) 'these arrays will be used for holding words
Dim arr3(1 To 119) 'this array will be used for holding and adding whole numbers
Dim i 'dim variable to add onto in arrays
Dim msgError 'error message box in case of a misinput

i = 1 'set i to first value

'Note: For arrays, the row value is "I + 1" because the first row in the spreadsheet is used for labels
Do 'setting up arrays
    arr1(i) = Cells(i + 1, 1).Text 'array 1 is the names of each element
    arr2(i) = Cells(i + 1, 2).Text 'array 2 is the symbols of each element
    arr3(i) = Cells(i + 1, 4).Text 'array 3 is the atomic mass of each element
    i = i + 1 'Add to i to repeat process on new row
Loop Until i = 119 'There are 119 rows, so stop at the end

If txtPercEle1.Text <> "" Then
'Now set it to its atomic mass value using arrays
i = 1 'reset i value
Do Until E1 = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtPercEle1.Text = arr1(i) Or txtPercEle1.Text = arr2(i) Then 'when the value entered is equal to the name of an element or symbol on the spreadsheet
        E1 = arr3(i)
        S = arr2(i)
    Else: i = i + 1 'Add to i to repeat process on new row
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into the first text box.", vbCritical, "Entry Error") 'therefore, an error message box will appear\
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied
Else: msgError = MsgBox("Sorry, please enter an element into the first text box.", vbCritical, "Entry Error")
End
End If

If txtPercEle2.Text <> "" Then
'Now set it to its atomic mass value using arrays
i = 1 'reset i value
Do Until E2 = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtPercEle2.Text = arr1(i) Or txtPercEle2.Text = arr2(i) Then 'when the value entered is equal to the name of an element or symbol on the spreadsheet
        E2 = arr3(i)
    Else: i = i + 1 'Add to i to repeat process on new row
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into the second text box.", vbCritical, "Entry Error") 'therefore, an error message box will appear\
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied
End If

If txtPercEle3.Text <> "" Then
'Now set it to its atomic mass value using arrays
i = 1 'reset i value
Do Until E3 = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtPercEle3.Text = arr1(i) Or txtPercEle3.Text = arr2(i) Then 'when the value entered is equal to the name of an element or symbol on the spreadsheet
        E3 = arr3(i)
    Else: i = i + 1 'Add to i to repeat process on new row
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into the third text box.", vbCritical, "Entry Error") 'therefore, an error message box will appear\
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied
End If

If txtPercEle4.Text <> "" Then
'Now set it to its atomic mass value using arrays
i = 1 'reset i value
Do Until E4 = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtPercEle4.Text = arr1(i) Or txtPercEle4.Text = arr2(i) Then 'when the value entered is equal to the name of an element or symbol on the spreadsheet
        E4 = arr3(i)
    Else: i = i + 1 'Add to i to repeat process on new row
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into the fourth text box.", vbCritical, "Entry Error") 'therefore, an error message box will appear\
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied
End If

'Now each molar mass must be multiplied by the quantity of each element
If IsNumeric(txtPercQuan1.Text) Then 'If the text entered is purely numbers...
Q1 = txtPercQuan1.Text 'Make it equal to its respective quantity variable
E1 = E1 * Q1 'And then multiply it by its respective element
End If

If IsNumeric(txtPercQuan2.Text) Then
Q2 = txtPercQuan2.Text
E2 = E2 * Q2
End If

If IsNumeric(txtPercQuan3.Text) Then
Q3 = txtPercQuan3.Text
E3 = E3 * Q3
End If

If IsNumeric(txtPercQuan4.Text) Then
Q4 = txtPercQuan4.Text
E4 = E4 * Q4
End If

txtPercResult.Text = (E1 / (E1 + E2 + E3 + E4)) * 100 & "% " & S

End Sub

Private Sub btnEle2_Click()
Dim E1, E2, E3, E4, Q1, Q2, Q3, Q4, S 'units to be entered; each E is an element, each Q is a quantity, and S is a symbol
Dim arr1(1 To 119), arr2(1 To 119) 'these arrays will be used for holding words
Dim arr3(1 To 119) 'this array will be used for holding and adding whole numbers
Dim i 'dim variable to add onto in arrays
Dim msgError 'error message box in case of a misinput

i = 1 'set i to first value

'Note: For arrays, the row value is "I + 1" because the first row in the spreadsheet is used for labels
Do 'setting up arrays
    arr1(i) = Cells(i + 1, 1).Text 'array 1 is the names of each element
    arr2(i) = Cells(i + 1, 2).Text 'array 2 is the symbols of each element
    arr3(i) = Cells(i + 1, 4).Text 'array 3 is the atomic mass of each element
    i = i + 1 'Add to i to repeat process on new row
Loop Until i = 119 'There are 119 rows, so stop at the end

If txtPercEle1.Text <> "" Then
'Now set it to its atomic mass value using arrays
i = 1 'reset i value
Do Until E1 = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtPercEle1.Text = arr1(i) Or txtPercEle1.Text = arr2(i) Then 'when the value entered is equal to the name of an element or symbol on the spreadsheet
        E1 = arr3(i)
    Else: i = i + 1 'Add to i to repeat process on new row
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into the first text box.", vbCritical, "Entry Error") 'therefore, an error message box will appear\
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied
End If

If txtPercEle2.Text <> "" Then
'Now set it to its atomic mass value using arrays
i = 1 'reset i value
Do Until E2 = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtPercEle2.Text = arr1(i) Or txtPercEle2.Text = arr2(i) Then 'when the value entered is equal to the name of an element or symbol on the spreadsheet
        E2 = arr3(i)
        S = arr2(i)
    Else: i = i + 1 'Add to i to repeat process on new row
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into the second text box.", vbCritical, "Entry Error") 'therefore, an error message box will appear\
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied
Else: msgError = MsgBox("Sorry, please enter an element into the second text box.", vbCritical, "Entry Error")
End
End If

If txtPercEle3.Text <> "" Then
'Now set it to its atomic mass value using arrays
i = 1 'reset i value
Do Until E3 = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtPercEle3.Text = arr1(i) Or txtPercEle3.Text = arr2(i) Then 'when the value entered is equal to the name of an element or symbol on the spreadsheet
        E3 = arr3(i)
    Else: i = i + 1 'Add to i to repeat process on new row
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into the third text box.", vbCritical, "Entry Error") 'therefore, an error message box will appear\
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied
End If

If txtPercEle4.Text <> "" Then
'Now set it to its atomic mass value using arrays
i = 1 'reset i value
Do Until E4 = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtPercEle4.Text = arr1(i) Or txtPercEle4.Text = arr2(i) Then 'when the value entered is equal to the name of an element or symbol on the spreadsheet
        E4 = arr3(i)
    Else: i = i + 1 'Add to i to repeat process on new row
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into the fourth text box.", vbCritical, "Entry Error") 'therefore, an error message box will appear\
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied
End If

If IsNumeric(txtPercQuan1.Text) Then
Q1 = txtPercQuan1.Text
E1 = E1 * Q1
End If

If IsNumeric(txtPercQuan2.Text) Then
Q2 = txtPercQuan2.Text
E2 = E2 * Q2
End If

If IsNumeric(txtPercQuan3.Text) Then
Q3 = txtPercQuan3.Text
E3 = E3 * Q3
End If

If IsNumeric(txtPercQuan4.Text) Then
Q4 = txtPercQuan4.Text
E4 = E4 * Q4
End If

txtPercResult.Text = (E2 / (E1 + E2 + E3 + E4)) * 100 & "% " & S
End Sub

Private Sub btnEle3_Click()
Dim E1, E2, E3, E4, Q1, Q2, Q3, Q4, S 'units to be entered; each E is an element, each Q is a quantity, and S is a symbol
Dim arr1(1 To 119), arr2(1 To 119) 'these arrays will be used for holding words
Dim arr3(1 To 119) 'this array will be used for holding and adding whole numbers
Dim i 'dim variable to add onto in arrays
Dim msgError 'error message box in case of a misinput

i = 1 'set i to first value

'Note: For arrays, the row value is "I + 1" because the first row in the spreadsheet is used for labels
Do 'setting up arrays
    arr1(i) = Cells(i + 1, 1).Text 'array 1 is the names of each element
    arr2(i) = Cells(i + 1, 2).Text 'array 2 is the symbols of each element
    arr3(i) = Cells(i + 1, 4).Text 'array 3 is the atomic mass of each element
    i = i + 1 'Add to i to repeat process on new row
Loop Until i = 119 'There are 119 rows, so stop at the end

If txtPercEle1.Text <> "" Then
'Now set it to its atomic mass value using arrays
i = 1 'reset i value
Do Until E1 = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtPercEle1.Text = arr1(i) Or txtPercEle1.Text = arr2(i) Then 'when the value entered is equal to the name of an element or symbol on the spreadsheet
        E1 = arr3(i)
    Else: i = i + 1 'Add to i to repeat process on new row
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into the first text box.", vbCritical, "Entry Error") 'therefore, an error message box will appear\
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied
End If

If txtPercEle2.Text <> "" Then
'Now set it to its atomic mass value using arrays
i = 1 'reset i value
Do Until E2 = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtPercEle2.Text = arr1(i) Or txtPercEle2.Text = arr2(i) Then 'when the value entered is equal to the name of an element or symbol on the spreadsheet
        E2 = arr3(i)
    Else: i = i + 1 'Add to i to repeat process on new row
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into the second text box.", vbCritical, "Entry Error") 'therefore, an error message box will appear\
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied
End If

If txtPercEle3.Text <> "" Then
'Now set it to its atomic mass value using arrays
i = 1 'reset i value
Do Until E3 = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtPercEle3.Text = arr1(i) Or txtPercEle3.Text = arr2(i) Then 'when the value entered is equal to the name of an element or symbol on the spreadsheet
        E3 = arr3(i)
        S = arr2(i)
    Else: i = i + 1 'Add to i to repeat process on new row
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into the third text box.", vbCritical, "Entry Error") 'therefore, an error message box will appear\
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied
Else: msgError = MsgBox("Sorry, please enter an element into the third text box.", vbCritical, "Entry Error")
End
End If

If txtPercEle4.Text <> "" Then
'Now set it to its atomic mass value using arrays
i = 1 'reset i value
Do Until E4 = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtPercEle4.Text = arr1(i) Or txtPercEle4.Text = arr2(i) Then 'when the value entered is equal to the name of an element or symbol on the spreadsheet
        E4 = arr3(i)
    Else: i = i + 1 'Add to i to repeat process on new row
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into the fourth text box.", vbCritical, "Entry Error") 'therefore, an error message box will appear\
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied
End If

If IsNumeric(txtPercQuan1.Text) Then
Q1 = txtPercQuan1.Text
E1 = E1 * Q1
End If

If IsNumeric(txtPercQuan2.Text) Then
Q2 = txtPercQuan2.Text
E2 = E2 * Q2
End If

If IsNumeric(txtPercQuan3.Text) Then
Q3 = txtPercQuan3.Text
E3 = E3 * Q3
End If

If IsNumeric(txtPercQuan4.Text) Then
Q4 = txtPercQuan4.Text
E4 = E4 * Q4
End If

txtPercResult.Text = (E3 / (E1 + E2 + E3 + E4)) * 100 & "% " & S
End Sub

Private Sub btnEle4_Click()
Dim E1, E2, E3, E4, Q1, Q2, Q3, Q4, S 'units to be entered; each E is an element, each Q is a quantity, and S is a symbol
Dim arr1(1 To 119), arr2(1 To 119) 'these arrays will be used for holding words
Dim arr3(1 To 119) 'this array will be used for holding and adding whole numbers
Dim i 'dim variable to add onto in arrays
Dim msgError 'error message box in case of a misinput

i = 1 'set i to first value

'Note: For arrays, the row value is "I + 1" because the first row in the spreadsheet is used for labels
Do 'setting up arrays
    arr1(i) = Cells(i + 1, 1).Text 'array 1 is the names of each element
    arr2(i) = Cells(i + 1, 2).Text 'array 2 is the symbols of each element
    arr3(i) = Cells(i + 1, 4).Text 'array 3 is the atomic mass of each element
    i = i + 1 'Add to i to repeat process on new row
Loop Until i = 119 'There are 119 rows, so stop at the end

If txtPercEle1.Text <> "" Then
'Now set it to its atomic mass value using arrays
i = 1 'reset i value
Do Until E1 = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtPercEle1.Text = arr1(i) Or txtPercEle1.Text = arr2(i) Then 'when the value entered is equal to the name of an element or symbol on the spreadsheet
        E1 = arr3(i)
    Else: i = i + 1 'Add to i to repeat process on new row
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into the first text box.", vbCritical, "Entry Error") 'therefore, an error message box will appear\
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied
End If

If txtPercEle2.Text <> "" Then
'Now set it to its atomic mass value using arrays
i = 1 'reset i value
Do Until E2 = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtPercEle2.Text = arr1(i) Or txtPercEle2.Text = arr2(i) Then 'when the value entered is equal to the name of an element or symbol on the spreadsheet
        E2 = arr3(i)
    Else: i = i + 1 'Add to i to repeat process on new row
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into the second text box.", vbCritical, "Entry Error") 'therefore, an error message box will appear\
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied
End If

If txtPercEle3.Text <> "" Then
'Now set it to its atomic mass value using arrays
i = 1 'reset i value
Do Until E3 = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtPercEle3.Text = arr1(i) Or txtPercEle3.Text = arr2(i) Then 'when the value entered is equal to the name of an element or symbol on the spreadsheet
        E3 = arr3(i)
    Else: i = i + 1 'Add to i to repeat process on new row
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into the third text box.", vbCritical, "Entry Error") 'therefore, an error message box will appear\
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied
End If

If txtPercEle4.Text <> "" Then
'Now set it to its atomic mass value using arrays
i = 1 'reset i value
Do Until E4 = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtPercEle4.Text = arr1(i) Or txtPercEle4.Text = arr2(i) Then 'when the value entered is equal to the name of an element or symbol on the spreadsheet
        E4 = arr3(i)
        S = arr2(i)
    Else: i = i + 1 'Add to i to repeat process on new row
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into the fourth text box.", vbCritical, "Entry Error") 'therefore, an error message box will appear\
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied
Else: msgError = MsgBox("Sorry, please enter an element into the fourth text box.", vbCritical, "Entry Error")
End
End If

If IsNumeric(txtPercQuan1.Text) Then
Q1 = txtPercQuan1.Text
E1 = E1 * Q1
End If

If IsNumeric(txtPercQuan2.Text) Then
Q2 = txtPercQuan2.Text
E2 = E2 * Q2
End If

If IsNumeric(txtPercQuan3.Text) Then
Q3 = txtPercQuan3.Text
E3 = E3 * Q3
End If

If IsNumeric(txtPercQuan4.Text) Then
Q4 = txtPercQuan4.Text
E4 = E4 * Q4
End If

txtPercResult.Text = (E4 / (E1 + E2 + E3 + E4)) * 100 & "% " & S
End Sub

Private Sub btnFindFormula_Click()
Dim E1 As String, E2 As String, E3 As String, E4 As String, M1, M2, M3, M4, A1, A2, A3, A4 'units to be entered
Dim arr1(1 To 119), arr2(1 To 119) 'these arrays will be used for holding words
Dim arr3(1 To 119) 'this array will be used for holding and adding whole numbers
Dim i 'dim variable to add onto in arrays
Dim msgError 'error message box in case of a misinput

i = 1 'set i to first value

'Note: For arrays, the row value is "I + 1" because the first row in the spreadsheet is used for labels
Do 'setting up arrays
    arr1(i) = Cells(i + 1, 1).Text 'array 1 is the names of each element
    arr2(i) = Cells(i + 1, 2).Text 'array 2 is the symbols of each element
    arr3(i) = Cells(i + 1, 4).Text 'array 3 is the atomic masses of each element
    i = i + 1 'Add to i to repeat process on new row
Loop Until i = 119 'There are 119 rows, so stop at the end

'This program relies on that if a text box is entered into, every prior text box has something entered as well
If txtFormEle2.Text <> "" And txtFormEle1.Text = "" Then 'If the second is not empty but the first is...
    msgError = MsgBox("Sorry, you must enter the elements consecutively in the text boxes; you left the first one empty.", vbCritical, "Entry Error") 'An error message will appear
    ElseIf txtFormEle3.Text <> "" And txtFormEle2.Text = "" Then 'If the third is not empty but the second is...
        msgError = MsgBox("Sorry, you must enter the elements consecutively in the text boxes; you left the second one empty.", vbCritical, "Entry Error") 'An error message will appear
            ElseIf txtFormEle4.Text <> "" And txtFormEle3.Text = "" Then 'If the fourth is not empty but the third is...
                msgError = MsgBox("Sorry, you must enter the elements in the consecutive text boxes; you left the third one empty.", vbCritical, "Entry Error") 'An error message will appear
End If

    i = 1 'reset i value
Do Until A1 = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtFormEle1.Text = arr1(i) Or txtFormEle1.Text = arr2(i) Then 'when the value entered is equal to the name of an element or symbol on the spreadsheet
        E1 = arr2(i)
        A1 = arr3(i)
    Else: i = i + 1 'Add to i to repeat process on new row
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into the first text box.", vbCritical, "Entry Error") 'therefore, an error message box will appear\
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied

    i = 1 'reset i value
Do Until A2 = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtFormEle2.Text = arr1(i) Or txtFormEle2.Text = arr2(i) Then 'when the value entered is equal to the name of an element or symbol on the spreadsheet
        E2 = arr2(i)
        A2 = arr3(i)
    Else: i = i + 1 'Add to i to repeat process on new row
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into the second text box.", vbCritical, "Entry Error") 'therefore, an error message box will appear\
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied

    i = 1 'reset i value
Do Until A3 = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtFormEle3.Text = arr1(i) Or txtFormEle3.Text = arr2(i) Then 'when the value entered is equal to the name of an element or symbol on the spreadsheet
        E3 = arr2(i)
        A3 = arr3(i)
    Else: i = i + 1 'Add to i to repeat process on new row
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into the second text box.", vbCritical, "Entry Error") 'therefore, an error message box will appear\
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied

    i = 1 'reset i value
Do Until A4 = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtFormEle4.Text = arr1(i) Or txtFormEle4.Text = arr2(i) Then 'when the value entered is equal to the name of an element or symbol on the spreadsheet
        E4 = arr2(i)
        A4 = arr3(i)
    Else: i = i + 1 'Add to i to repeat process on new row
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into the second text box.", vbCritical, "Entry Error") 'therefore, an error message box will appear\
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied

'Each A variable is to divide their masses in order to create their respective M variables
If txtFormEle1.Text <> "" And IsNumeric(txtFormMass1.Text) Then 'If the text entered for the element is not blank and the mass entered is numeric...
    M1 = txtFormMass1.Text / A1 'Make M the quotient of the mass and A variable
    ElseIf txtFormEle1.Text <> "" Then 'If the text entered for the mass is not numeric...
        msgError = MsgBox("Sorry, you must enter the first element's mass or percent composition into its corresponding text box.", vbCritical, "Entry Error") 'Create an error message box
End If

If txtFormEle2.Text <> "" And IsNumeric(txtFormMass2.Text) Then
    M2 = txtFormMass2.Text / A2
    ElseIf txtFormEle2.Text <> "" Then
        msgError = MsgBox("Sorry, you must enter the second element's mass or percent composition into its corresponding text box.", vbCritical, "Entry Error")
End If

If txtFormEle3.Text <> "" And IsNumeric(txtFormMass3.Text) Then
    M3 = txtFormMass3.Text / A3
    ElseIf txtFormEle3.Text <> "" Then
        msgError = MsgBox("Sorry, you must enter the third element's mass or percent composition into its corresponding text box.", vbCritical, "Entry Error")
End If

If txtFormEle4.Text <> "" And IsNumeric(txtFormMass4.Text) Then
    M4 = txtFormMass4.Text / A4
    ElseIf txtFormEle4.Text <> "" Then
        msgError = MsgBox("Sorry, you must enter the fourth element's mass or percent composition into its corresponding text box.", vbCritical, "Entry Error")
End If

'Now, all of the M variables must be divided by the smallest of the four. Each If statement will factor out any empty M variables.
If txtFormEle4.Text <> "" And IsNumeric(txtFormMass4.Text) Then 'If all four elements are present...
If M1 <= M2 And M1 <= M3 And M1 <= M4 Then 'If M1 is the smallest...
    M1 = M1 / M1
    M2 = M2 / M1
    M3 = M3 / M1
    M4 = M4 / M1
    ElseIf M2 <= M1 And M2 <= M3 And M2 <= M4 Then 'If M2 is the smallest...
        M1 = M1 / M2
        M2 = M2 / M2
        M3 = M3 / M2
        M4 = M4 / M2
        ElseIf M3 <= M1 And M3 <= M2 And M3 <= M4 Then 'If M3 is the smallest...
            M1 = M1 / M3
            M2 = M2 / M3
            M3 = M3 / M3
            M4 = M4 / M3
                ElseIf M4 <= M1 And M4 <= M2 And M4 <= M3 Then 'If M4 is the smallest...
                    M1 = M1 / M2
                    M2 = M2 / M2
                    M3 = M3 / M2
                    M4 = M4 / M2
                    End If
ElseIf txtFormEle3.Text <> "" And IsNumeric(txtFormMass3.Text) Then 'If only three elements are present...
If M1 <= M2 And M1 <= M3 Then 'If M1 is the smallest...
    M1 = M1 / M1
    M2 = M2 / M1
    M3 = M3 / M1
    ElseIf M2 <= M1 And M2 <= M3 Then 'If M2 is the smallest...
        M1 = M1 / M2
        M2 = M2 / M2
        M3 = M3 / M2
        ElseIf M3 <= M1 And M3 <= M2 Then 'If M3 is the smallest...
            M1 = M1 / M3
            M2 = M2 / M3
            M3 = M3 / M3
            End If
ElseIf txtFormEle2.Text <> "" And IsNumeric(txtFormMass2.Text) Then 'If only two elements are present...
If M1 <= M2 Then 'If M1 is the smallest...
    M1 = M1 / M1
    M2 = M2 / M1
    ElseIf M2 <= M1 Then 'If M2 is the smallest...
        M1 = M1 / M2
        M2 = M2 / M2
        End If
Else 'If nothing past the first element's text box has text entered...
    msgError = MsgBox("Sorry, you must enter more than one element.", vbCritical, "Entry Error") 'Create an error message box
End If

M1 = Math.Round(M1, 1) 'Round M1 to the nearest tenth
M2 = Math.Round(M2, 1) 'Round M2 to the nearest tenth
M3 = Math.Round(M3, 1) 'Round M3 to the nearest tenth
M4 = Math.Round(M4, 1) 'Round M4 to the nearest tenth

Do
If Right(M1, 1) = 5 Or Right(M2, 1) = 5 Or Right(M3, 1) = 5 Or Right(M4, 1) = 5 Then
M1 = M1 * 2
M2 = M2 * 2
M3 = M3 * 2
M4 = M4 * 2
End If
Loop Until Right(M1, 1) <> 5 And Right(M2, 1) <> 5 And Right(M3, 1) <> 5 And Right(M4, 1) <> 5

M1 = Math.Round(M1, 0) 'Round M1 to the decimal point
M2 = Math.Round(M2, 0) 'Round M2 to the decimal point
M3 = Math.Round(M3, 0) 'Round M3 to the decimal point
M4 = Math.Round(M4, 0) 'Round M4 to the decimal point

'Once again, the answer depends on how many elements were entered
If txtFormEle4.Text <> "" And IsNumeric(txtFormMass4.Text) Then 'If all four are present...
    txtFormResult.Text = E1 & M1 & E2 & M2 & E3 & M3 & E4 & M4
ElseIf txtFormEle3.Text <> "" And IsNumeric(txtFormMass3.Text) Then 'If only three are present...
    txtFormResult.Text = E1 & M1 & E2 & M2 & E3 & M3
ElseIf txtFormEle2.Text <> "" And IsNumeric(txtFormMass2.Text) Then 'If only two are present...
    txtFormResult.Text = E1 & M1 & E2 & M2
End If

End Sub
