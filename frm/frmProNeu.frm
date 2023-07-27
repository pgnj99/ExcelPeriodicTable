VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProNeu 
   Caption         =   "Protons, Neutrons, and Electrons"
   ClientHeight    =   5400
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   4416
   OleObjectBlob   =   "frmProNeu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProNeu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnProtons_Click()
Dim arr1(1 To 119), arr2(1 To 119) 'these arrays will be used for holding words
Dim arr3(1 To 119) As Integer 'these arrays will be used for holding and adding whole numbers
Dim I1, I2, I3, I4, T1, T2, T3, T4, R 'dim the eight variables; I will be added onto in arrays, T will be for text boxes, R will be for result
Dim msgError 'error message box in case of a misinput

I1 = 1 'set I1 to first value
I2 = 1 'set I2 to first value
I3 = 1 'set I3 to first value
I4 = 1 'set I4 to first value

'Note: For arrays, the row value is "I + 1" because the first row in the spreadsheet is used for labels
Do 'setting up arrays
    arr1(I1) = Cells(I1 + 1, 1).Text 'array 1 is the names of each element
    arr1(I2) = Cells(I2 + 1, 1).Text
    arr1(I3) = Cells(I3 + 1, 1).Text
    arr1(I4) = Cells(I4 + 1, 1).Text
    arr2(I1) = Cells(I1 + 1, 2).Text 'array 2 is the symbols of each element
    arr2(I2) = Cells(I2 + 1, 2).Text
    arr2(I3) = Cells(I3 + 1, 2).Text
    arr2(I4) = Cells(I4 + 1, 2).Text
    arr3(I1) = Cells(I1 + 1, 3).Text 'array 3 is the atomic numbers of each element (text box 1)
    arr3(I2) = Cells(I2 + 1, 3).Text 'array 4 is the atomic numbers of each element (text box 2)
    arr3(I3) = Cells(I3 + 1, 3).Text 'array 5 is the atomic numbers of each element (text box 3)
    arr3(I4) = Cells(I4 + 1, 3).Text 'array 6 is the atomic numbers of each element (text box 4)
    I1 = I1 + 1 'Add to I1 to repeat process on new row
    I2 = I2 + 1 'Add to I2 to repeat process on new row
    I3 = I3 + 1 'Add to I3 to repeat process on new row
    I4 = I4 + 1 'Add to I4 to repeat process on new row
Loop Until I1 = 119 And I2 = 119 And I3 = 119 And I4 = 119 'There are 119 rows, so stop at the end

T1 = txtProNeuElement1.Text 'T1 is what is entered into the first element's text box
T2 = txtProNeuElement2.Text 'T2 is what is entered into the second element's text box
T3 = txtProNeuElement3.Text 'T3 is what is entered into the third element's text box
T4 = txtProNeuElement4.Text 'T4 is what is entered into the fourth element's text box

'For the first text box:
I1 = 1 'reset I1 value
Do Until T1 = Cells(I1 + 1, 1).Text Or T1 = Cells(I1 + 1, 2).Text 'the goal is to find the equivalent on the spreadsheet to the value entered
If T1 = Cells(I1 + 1, 1).Text Then 'when the value entered is equal to the name of an element on the spreadsheet
    arr1(I1) = arr3(I1) 'the value entered equals its atomic number in the third column on that row
    ElseIf T1 = Cells(I1 + 1, 2).Text Then 'when the value entered is equal to the name of a symbol on the spreadsheet
        arr2(I1) = arr3(I1) 'the value entered equals its atomic number in the third column on that row
    Else: I1 = I1 + 1 'Add to I1 to repeat process on new row
If I1 > 119 Then 'I1 will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, your first answer is not an element.", vbCritical, "Entry Error") 'therefore, an error message box will appear
    End 'end program when message box is closed
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied

'For the second text box:
I2 = 1 'reset I2 value
Do Until T2 = Cells(I2 + 1, 1).Text Or T2 = Cells(I2 + 1, 2).Text 'the goal is to find the equivalent on the spreadsheet to the value entered
If T2 = Cells(I2 + 1, 1).Text Then 'when the value entered is equal to the name of an element on the spreadsheet
    arr1(I2) = arr3(I2) 'the value entered equals its atomic number in the third column on that row
    ElseIf T2 = Cells(I2 + 1, 2).Text Then 'when the value entered is equal to the name of a symbol on the spreadsheet
        arr2(I2) = arr3(I2) 'the value entered equals its atomic number in the third column on that row
    Else: I2 = I2 + 1 'Add to I2 to repeat process on new row
If I2 > 119 Then 'I2 will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, your second answer is not an element.", vbCritical, "Entry Error") 'therefore, an error message box will appear
    End 'end program when message box is closed
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied

'For the third text box:
I3 = 1 'reset I3 value
Do Until T3 = Cells(I3 + 1, 1).Text Or T3 = Cells(I3 + 1, 2).Text 'the goal is to find the equivalent on the spreadsheet to the value entered
If T3 = Cells(I3 + 1, 1).Text Then 'when the value entered is equal to the name of an element on the spreadsheet
    arr1(I3) = arr3(I3) 'the value entered equals its atomic number in the third column on that row
    ElseIf T3 = Cells(I3 + 1, 2).Text Then 'when the value entered is equal to the name of a symbol on the spreadsheet
        arr2(I3) = arr3(I3) 'the value entered equals its atomic number in the third column on that row
    Else: I3 = I3 + 1 'Add to I3 to repeat process on new row
If I3 > 119 Then 'I3 will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, your third answer is not an element.", vbCritical, "Entry Error") 'therefore, an error message box will appear
    End 'end program when message box is closed
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied

'For the fourth text box:
I4 = 1 'reset I4 value
Do Until T4 = Cells(I4 + 1, 1).Text Or T4 = Cells(I4 + 1, 2).Text 'the goal is to find the equivalent on the spreadsheet to the value entered
If T4 = Cells(I4 + 1, 1).Text Then 'when the value entered is equal to the name of an element on the spreadsheet
    arr1(I4) = arr3(I4) 'the value entered equals its atomic number in the third column on that row
    ElseIf T4 = Cells(I4 + 1, 2).Text Then 'when the value entered is equal to the name of a symbol on the spreadsheet
        arr2(I4) = arr3(I4) 'the value entered equals its atomic number in the third column on that row
    Else: I4 = I4 + 1 'Add to I4 to repeat process on new row
If I4 > 119 Then 'I4 will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, your fourth answer is not an element.", vbCritical, "Entry Error") 'therefore, an error message box will appear
    End 'end program when message box is closed
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied

R = arr3(I1) + arr3(I2) + arr3(I3) + arr3(I4) 'the sum of each array will be used for the result

If R = 1 Then 'if the sum of each array happens to be 1
    txtProNeuResult = R & " Proton" 'display result in bottom text box with singular proton label
    Else: txtProNeuResult = R & " Protons" 'otherwise, display result in bottom text box with plural protons label
End If 'end process

End Sub

Private Sub btnNeutrons_Click()
'Note: This sub is almost exactly the same as the one for protons, with the differences being the different column used for values

Dim arr1(1 To 119), arr2(1 To 119) 'these arrays will be used for holding words
Dim arr3(1 To 119) As Integer 'these arrays will be used for holding and adding whole numbers
Dim I1, I2, I3, I4, T1, T2, T3, T4, R 'dim the eight variables; I will be added onto in arrays, T will be for text boxes, R will be for result
Dim msgError 'error message box in case of a misinput

I1 = 1 'set I1 to first value
I2 = 1 'set I2 to first value
I3 = 1 'set I3 to first value
I4 = 1 'set I4 to first value

'Note: For arrays, the row value is "I + 1" because the first row in the spreadsheet is used for labels
Do 'setting up arrays
    arr1(I1) = Cells(I1 + 1, 1).Text 'array 1 is the names of each element
    arr1(I2) = Cells(I2 + 1, 1).Text
    arr1(I3) = Cells(I3 + 1, 1).Text
    arr1(I4) = Cells(I4 + 1, 1).Text
    arr2(I1) = Cells(I1 + 1, 2).Text 'array 2 is the symbols of each element
    arr2(I2) = Cells(I2 + 1, 2).Text
    arr2(I3) = Cells(I3 + 1, 2).Text
    arr2(I4) = Cells(I4 + 1, 2).Text
    arr3(I1) = Cells(I1 + 1, 5).Text 'array 3 is the neutron count of each element (text box 1)
    arr3(I2) = Cells(I2 + 1, 5).Text 'array 4 is the neutron count of each element (text box 2)
    arr3(I3) = Cells(I3 + 1, 5).Text 'array 5 is the neutron count of each element (text box 3)
    arr3(I4) = Cells(I4 + 1, 5).Text 'array 6 is the neutron count of each element (text box 4)
    I1 = I1 + 1 'Add to I1 to repeat process on new row
    I2 = I2 + 1 'Add to I2 to repeat process on new row
    I3 = I3 + 1 'Add to I3 to repeat process on new row
    I4 = I4 + 1 'Add to I4 to repeat process on new row
Loop Until I1 = 119 And I2 = 119 And I3 = 119 And I4 = 119 'There are 119 rows, so stop at the end

T1 = txtProNeuElement1.Text 'T1 is what is entered into the first element's text box
T2 = txtProNeuElement2.Text 'T2 is what is entered into the second element's text box
T3 = txtProNeuElement3.Text 'T3 is what is entered into the third element's text box
T4 = txtProNeuElement4.Text 'T4 is what is entered into the fourth element's text box

'For the first text box:
I1 = 1 'reset I1 value
Do Until T1 = Cells(I1 + 1, 1).Text Or T1 = Cells(I1 + 1, 2).Text 'the goal is to find the equivalent on the spreadsheet to the value entered
If T1 = Cells(I1 + 1, 1).Text Then 'when the value entered is equal to the name of an element on the spreadsheet
    arr1(I1) = arr3(I1) 'the value entered equals its neutron count in the third column on that row
    ElseIf T1 = Cells(I1 + 1, 2).Text Then 'when the value entered is equal to the name of a symbol on the spreadsheet
        arr2(I1) = arr3(I1) 'the value entered equals its neutron count in the third column on that row
    Else: I1 = I1 + 1 'Add to I1 to repeat process on new row
If I1 > 119 Then 'I1 will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, your first answer is not an element.", vbCritical, "Entry Error") 'therefore, an error message box will appear
    End 'end program when message box is closed
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied

'For the second text box:
I2 = 1 'reset I2 value
Do Until T2 = Cells(I2 + 1, 1).Text Or T2 = Cells(I2 + 1, 2).Text 'the goal is to find the equivalent on the spreadsheet to the value entered
If T2 = Cells(I2 + 1, 1).Text Then 'when the value entered is equal to the name of an element on the spreadsheet
    arr1(I2) = arr3(I2) 'the value entered equals its neutron count in the third column on that row
    ElseIf T2 = Cells(I2 + 1, 2).Text Then 'when the value entered is equal to the name of a symbol on the spreadsheet
        arr2(I2) = arr3(I2) 'the value entered equals its neutron count in the third column on that row
    Else: I2 = I2 + 1 'Add to I2 to repeat process on new row
If I2 > 119 Then 'I2 will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, your second answer is not an element.", vbCritical, "Entry Error") 'therefore, an error message box will appear
    End 'end program when message box is closed
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied

'For the third text box:
I3 = 1 'reset I3 value
Do Until T3 = Cells(I3 + 1, 1).Text Or T3 = Cells(I3 + 1, 2).Text 'the goal is to find the equivalent on the spreadsheet to the value entered
If T3 = Cells(I3 + 1, 1).Text Then 'when the value entered is equal to the name of an element on the spreadsheet
    arr1(I3) = arr3(I3) 'the value entered equals its neutron count in the third column on that row
    ElseIf T3 = Cells(I3 + 1, 2).Text Then 'when the value entered is equal to the name of a symbol on the spreadsheet
        arr2(I3) = arr3(I3) 'the value entered equals its neutron count in the third column on that row
    Else: I3 = I3 + 1 'Add to I3 to repeat process on new row
If I3 > 119 Then 'I3 will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, your third answer is not an element.", vbCritical, "Entry Error") 'therefore, an error message box will appear
    End 'end program when message box is closed
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied

'For the fourth text box:
I4 = 1 'reset I4 value
Do Until T4 = Cells(I4 + 1, 1).Text Or T4 = Cells(I4 + 1, 2).Text 'the goal is to find the equivalent on the spreadsheet to the value entered
If T4 = Cells(I4 + 1, 1).Text Then 'when the value entered is equal to the name of an element on the spreadsheet
    arr1(I4) = arr3(I4) 'the value entered equals its neutron count in the third column on that row
    ElseIf T4 = Cells(I4 + 1, 2).Text Then 'when the value entered is equal to the name of a symbol on the spreadsheet
        arr2(I4) = arr3(I4) 'the value entered equals its neutron count in the third column on that row
    Else: I4 = I4 + 1 'Add to I4 to repeat process on new row
If I4 > 119 Then 'I4 will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, your fourth answer is not an element.", vbCritical, "Entry Error") 'therefore, an error message box will appear
    End 'end program when message box is closed
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied

R = arr3(I1) + arr3(I2) + arr3(I3) + arr3(I4) 'the sum of each array will be used for the result

If R = 1 Then 'if the sum of each array happens to be 1
    txtProNeuResult = R & " Neutron" 'display result in bottom text box with singular neutron label
    Else: txtProNeuResult = R & " Neutrons" 'otherwise, display result in bottom text box with plural neutrons label
End If 'end process

End Sub

Private Sub btnElectrons_Click()
'Note: This sub is almost exactly the same as the one for protons, with the differences being the addition of charges

Dim arr1(1 To 119), arr2(1 To 119) 'these arrays will be used for holding words
Dim arr3(1 To 119) As Integer 'these arrays will be used for holding and adding whole numbers
Dim I1, I2, I3, I4, T1, T2, T3, T4, R 'dim the eight variables; I will be added onto in arrays, T will be for text boxes, R will be for result
Dim C1 As Integer, C2 As Integer, C3 As Integer, C4 As Integer 'these will act as the charges for the right text boxes
Dim msgError 'error message box in case of a misinput

I1 = 1 'set I1 to first value
I2 = 1 'set I2 to first value
I3 = 1 'set I3 to first value
I4 = 1 'set I4 to first value

'Note: For arrays, the row value is "I + 1" because the first row in the spreadsheet is used for labels
Do 'setting up arrays
    arr1(I1) = Cells(I1 + 1, 1).Text 'array 1 is the names of each element
    arr1(I2) = Cells(I2 + 1, 1).Text
    arr1(I3) = Cells(I3 + 1, 1).Text
    arr1(I4) = Cells(I4 + 1, 1).Text
    arr2(I1) = Cells(I1 + 1, 2).Text 'array 2 is the symbols of each element
    arr2(I2) = Cells(I2 + 1, 2).Text
    arr2(I3) = Cells(I3 + 1, 2).Text
    arr2(I4) = Cells(I4 + 1, 2).Text
    arr3(I1) = Cells(I1 + 1, 3).Text 'array 3 is the atomic numbers of each element (text box 1)
    arr3(I2) = Cells(I2 + 1, 3).Text 'array 4 is the atomic numbers of each element (text box 2)
    arr3(I3) = Cells(I3 + 1, 3).Text 'array 5 is the atomic numbers of each element (text box 3)
    arr3(I4) = Cells(I4 + 1, 3).Text 'array 6 is the atomic numbers of each element (text box 4)
    I1 = I1 + 1 'Add to I1 to repeat process on new row
    I2 = I2 + 1 'Add to I2 to repeat process on new row
    I3 = I3 + 1 'Add to I3 to repeat process on new row
    I4 = I4 + 1 'Add to I4 to repeat process on new row
Loop Until I1 = 119 And I2 = 119 And I3 = 119 And I4 = 119 'There are 119 rows, so stop at the end

T1 = txtProNeuElement1.Text 'T1 is what is entered into the first element's text box
T2 = txtProNeuElement2.Text 'T2 is what is entered into the second element's text box
T3 = txtProNeuElement3.Text 'T3 is what is entered into the third element's text box
T4 = txtProNeuElement4.Text 'T4 is what is entered into the fourth element's text box

'For the first text box:
I1 = 1 'reset I1 value
Do Until T1 = Cells(I1 + 1, 1).Text Or T1 = Cells(I1 + 1, 2).Text 'the goal is to find the equivalent on the spreadsheet to the value entered
If T1 = Cells(I1 + 1, 1).Text Then 'when the value entered is equal to the name of an element on the spreadsheet
    arr1(I1) = arr3(I1) 'the value entered equals its atomic number in the third column on that row
    ElseIf T1 = Cells(I1 + 1, 2).Text Then 'when the value entered is equal to the name of a symbol on the spreadsheet
        arr2(I1) = arr3(I1) 'the value entered equals its atomic number in the third column on that row
    Else: I1 = I1 + 1 'Add to I1 to repeat process on new row
If I1 > 119 Then 'I1 will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, your first answer is not an element.", vbCritical, "Entry Error") 'therefore, an error message box will appear
    End 'end program when message box is closed
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied

'For the second text box:
I2 = 1 'reset I2 value
Do Until T2 = Cells(I2 + 1, 1).Text Or T2 = Cells(I2 + 1, 2).Text 'the goal is to find the equivalent on the spreadsheet to the value entered
If T2 = Cells(I2 + 1, 1).Text Then 'when the value entered is equal to the name of an element on the spreadsheet
    arr1(I2) = arr3(I2) 'the value entered equals its atomic number in the third column on that row
    ElseIf T2 = Cells(I2 + 1, 2).Text Then 'when the value entered is equal to the name of a symbol on the spreadsheet
        arr2(I2) = arr3(I2) 'the value entered equals its atomic number in the third column on that row
    Else: I2 = I2 + 1 'Add to I2 to repeat process on new row
If I2 > 119 Then 'I2 will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, your second answer is not an element.", vbCritical, "Entry Error") 'therefore, an error message box will appear
    End 'end program when message box is closed
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied

'For the third text box:
I3 = 1 'reset I3 value
Do Until T3 = Cells(I3 + 1, 1).Text Or T3 = Cells(I3 + 1, 2).Text 'the goal is to find the equivalent on the spreadsheet to the value entered
If T3 = Cells(I3 + 1, 1).Text Then 'when the value entered is equal to the name of an element on the spreadsheet
    arr1(I3) = arr3(I3) 'the value entered equals its atomic number in the third column on that row
    ElseIf T3 = Cells(I3 + 1, 2).Text Then 'when the value entered is equal to the name of a symbol on the spreadsheet
        arr2(I3) = arr3(I3) 'the value entered equals its atomic number in the third column on that row
    Else: I3 = I3 + 1 'Add to I3 to repeat process on new row
If I3 > 119 Then 'I3 will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, your third answer is not an element.", vbCritical, "Entry Error") 'therefore, an error message box will appear
    End 'end program when message box is closed
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied

'For the fourth text box:
I4 = 1 'reset I4 value
Do Until T4 = Cells(I4 + 1, 1).Text Or T4 = Cells(I4 + 1, 2).Text 'the goal is to find the equivalent on the spreadsheet to the value entered
If T4 = Cells(I4 + 1, 1).Text Then 'when the value entered is equal to the name of an element on the spreadsheet
    arr1(I4) = arr3(I4) 'the value entered equals its atomic number in the third column on that row
    ElseIf T4 = Cells(I4 + 1, 2).Text Then 'when the value entered is equal to the name of a symbol on the spreadsheet
        arr2(I4) = arr3(I4) 'the value entered equals its atomic number in the third column on that row
    Else: I4 = I4 + 1 'Add to I4 to repeat process on new row
If I4 > 119 Then 'I4 will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, your fourth answer is not an element.", vbCritical, "Entry Error") 'therefore, an error message box will appear
    End 'end program when message box is closed
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied


'The charges are integers, so blank values will crash the program; therefore, they must be numeric
If IsNumeric(txtProNeuCharge1.Text) Then 'If the charge value is numeric...
    C1 = txtProNeuCharge1.Text 'The charge will equal its equivalent text box
    Else
        C1 = 0 'If what is entered is not a number, then the charge is 0
End If

'The program will crash if nothing is entered into a charge's text box, so that must be accounted for
If IsNumeric(txtProNeuCharge2.Text) Then 'Only if the text box is numeric...
    C2 = txtProNeuCharge2.Text 'make the value equal to the charge variable
    Else 'If anything other than numbers are entered...
        C2 = 0 'make it equal 0, like the user would want
End If

If IsNumeric(txtProNeuCharge3.Text) Then
    C3 = txtProNeuCharge3.Text
    Else
        C3 = 0
End If

If IsNumeric(txtProNeuCharge4.Text) Then
    C4 = txtProNeuCharge4.Text
    Else
        C4 = 0
End If

R = arr3(I1) + arr3(I2) + arr3(I3) + arr3(I4) - C1 - C2 - C3 - C4 'the sum of each array will be subtracted from each charge for the result

If R = 1 Then 'if the sum of each array happens to be 1
    txtProNeuResult = R & " Electron" 'display result in bottom text box with singular electron label
    Else: txtProNeuResult = R & " Electrons" 'otherwise, display result in bottom text box with plural electrons label
End If 'end process

End Sub
