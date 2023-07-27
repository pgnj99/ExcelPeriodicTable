VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMolarity 
   Caption         =   "Molarity"
   ClientHeight    =   9084.001
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   3840
   OleObjectBlob   =   "frmMolarity.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMolarity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnMola1_Click()
'mol to M
Dim mole As Double, liter As Double, molarity As Double, mass As Double, grams As Double, result As Double 'units to be entered
Dim msgError

If IsNumeric(txtMolaMol.Text) Then 'If a numeric value is entered into mol's text box then
    mole = txtMolaMol.Text 'Proceed as normal by assigning the variable to it
Else 'Otherwise, a default must be set to prevent crashing
    mole = 1 'The default is 1, so set it to this instead
End If

If IsNumeric(txtMolaLit.Text) Then 'If the text entered for liters is numeric...
    liter = txtMolaLit.Text 'text entered into liters = second text box
    mole = txtMolaMol.Text 'text entered into moles = third text box
    result = (mole) / (liter)
Else 'If not numeric...
msgError = MsgBox("Enter numerical values into L and mol.", vbCritical, "Entry Error") 'Create an error message box
End If

txtMolaResult1.Text = result & " M" 'display result in the result text box

End Sub

Private Sub btnMola2_Click()
'grams to M
Dim mole As Double, liter As Double, molarity As Double, mass As Double, grams As Double, result As Double 'units to be entered
Dim arr1(1 To 119), arr2(1 To 119) 'these arrays will be used for holding words
Dim arr3(1 To 119) 'this array will be used for holding and adding whole numbers
Dim i 'dim variable to add onto in arrays
Dim msgError 'error message box in case of a misinput

i = 1 'set I1 to first value

'Note: For arrays, the row value is "I + 1" because the first row in the spreadsheet is used for labels
Do 'setting up arrays
    arr1(i) = Cells(i + 1, 1).Text 'array 1 is the names of each element
    arr2(i) = Cells(i + 1, 2).Text 'array 2 is the symbols of each element
    arr3(i) = Cells(i + 1, 4).Text 'array 3 is the atomic mass of each element (text box 1)
    i = i + 1 'Add to I1 to repeat process on new row
Loop Until i = 119 'There are 119 rows, so stop at the end

If IsNumeric(txtMolaMass.Text) Then 'If the value entered into the molar mass text box is purely numbers...
    mass = txtMolaMass.Text 'simply make it equal to the mass variable
Else
i = 1 'reset i value
Do Until mass = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtMolaMass.Text = arr1(i) Or txtMolaMass.Text = arr2(i) Then 'when the value entered is equal to the name or symbol of an element on the spreadsheet
    mass = arr3(i) 'the value entered equals its atomic mass in the fourth column on that row
    Else: i = i + 1 'Add to i to repeat process on new row
End If
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a number or valid element into the mol. mass box.", vbCritical, "Entry Error") 'therefore, an error message box will appear
    End 'end program when message box is closed
    Else
End If 'end the process
Loop 'loop until the do requirement is satisfied
End If

If IsNumeric(txtMolaMol.Text) Then 'If a numeric value is entered into mol's text box then
    mole = txtMolaMol.Text 'Proceed as normal by assigning the variable to it
Else 'Otherwise, a default must be set to prevent crashing
    mole = 1 'The default is 1, so set it to this instead
End If

If IsNumeric(txtMolaGram.Text) And IsNumeric(txtMolaLit.Text) Then
    liter = txtMolaLit.Text 'text entered into liters = second text box
    grams = txtMolaGram.Text 'text entered into moles = third text box
    result = (grams * (mole / mass)) / liter
Else 'If not numeric...
msgError = MsgBox("Enter numerical values into L and mol.", vbCritical, "Entry Error") 'Create an error message box
End If


txtMolaResult1.Text = result & " M" 'display result

End Sub

Private Sub btnMola3_Click()
'M to mol
Dim mole As Double, liter As Double, molarity As Double, mass As Double, grams As Double, result As Double 'units to be entered
Dim msgError

If IsNumeric(txtMolaMola.Text) And IsNumeric(txtMolaLit.Text) Then 'If the text entered for molarity and liters are both numeric...
    liter = txtMolaLit.Text 'text entered into liters = second text box
    molarity = txtMolaMola.Text 'text entered into molarity = third text box
    result = (molarity) * (liter)
Else 'If not numeric...
msgError = MsgBox("Enter numerical values into L and mol.", vbCritical, "Entry Error") 'Create an error message box
End If

txtMolaResult1.Text = result & " moles" 'display result
End Sub

Private Sub btnMola4_Click()
'M to grams
Dim mole As Double, liter As Double, molarity As Double, mass As Double, grams As Double, result As Double 'units to be entered
Dim arr1(1 To 119), arr2(1 To 119) 'these arrays will be used for holding words
Dim arr3(1 To 119) 'this array will be used for holding and adding whole numbers
Dim i 'dim variable to add onto in arrays
Dim msgError 'error message box in case of a misinput

i = 1 'set I1 to first value

'Note: For arrays, the row value is "I + 1" because the first row in the spreadsheet is used for labels
Do 'setting up arrays
    arr1(i) = Cells(i + 1, 1).Text 'array 1 is the names of each element
    arr2(i) = Cells(i + 1, 2).Text 'array 2 is the symbols of each element
    arr3(i) = Cells(i + 1, 4).Text 'array 3 is the atomic mass of each element (text box 1)
    i = i + 1 'Add to I1 to repeat process on new row
Loop Until i = 119 'There are 119 rows, so stop at the end

If IsNumeric(txtMolaMass.Text) Then 'If the value entered into the molar mass text box is purely numbers...
    mass = txtMolaMass.Text 'simply make it equal to the mass variable
Else
i = 1 'reset i value
Do Until mass = arr3(i) 'the goal is to find the equivalent on the spreadsheet to the value entered
If txtMolaMass.Text = arr1(i) Or txtMolaMass.Text = arr2(i) Then 'when the value entered is equal to the name or symbol of an element on the spreadsheet
    mass = arr3(i) 'the value entered equals its atomic mass in the fourth column on that row
    Else: i = i + 1 'Add to i to repeat process on new row
End If
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a number or valid element into the mol. mass box.", vbCritical, "Entry Error") 'therefore, an error message box will appear
    End 'end program when message box is closed
    Else
End If 'end the process
Loop 'loop until the do requirement is satisfied
End If

If IsNumeric(txtMolaMol.Text) Then 'If a numeric value is entered into mol's text box then
    mole = txtMolaMol.Text 'Proceed as normal by assigning the variable to it
Else 'Otherwise, a default must be set to prevent crashing
    mole = 1 'The default is 1, so set it to this instead
End If

If IsNumeric(txtMolaMola.Text) And IsNumeric(txtMolaLit.Text) Then
    liter = txtMolaLit.Text 'text entered into liters = second text box
    molarity = txtMolaMola.Text 'text entered into moles = third text box
    result = (molarity * liter) * (mass / mole)
Else 'If not numeric...
msgError = MsgBox("Enter numerical values into L and mol.", vbCritical, "Entry Error") 'Create an error message box
End If


txtMolaResult1.Text = result & " grams" 'display result
End Sub

Private Sub btnMolaFindM_Click()
Dim M1 As Double, V1 As Double, M2 As Double, V2 As Double 'units to be entered
Dim msgError

If IsNumeric(txtMolaM1.Text) And IsNumeric(txtMolaV1.Text) And IsNumeric(txtMolaV2.Text) Then
    M1 = txtMolaM1.Text 'text entered into Substance 1 Molarity
    V1 = txtMolaV1.Text 'text entered into Substance 1 Liters
    V2 = txtMolaV2.Text 'text entered into Substance 2 Liters
    M2 = (M1 * V1) / V2
Else
msgError = MsgBox("Enter numerical values into both Sub. 1 values and Sub. 2 Liters.", vbCritical, "Entry Error") 'error message if numbers are not used
End If

txtMolaResult2.Text = M2 & " M" 'display result
End Sub

Private Sub btnMolaFindV_Click()
Dim M1 As Double, V1 As Double, M2 As Double, V2 As Double 'units to be entered
Dim msgError

If IsNumeric(txtMolaM1.Text) And IsNumeric(txtMolaV1.Text) And IsNumeric(txtMolaM2.Text) Then
    M1 = txtMolaM1.Text 'text entered into Substance 1 Molarity
    V1 = txtMolaV1.Text 'text entered into Substance 1 Liters
    M2 = txtMolaM2.Text 'text entered into Substance 2 Molarity
    V2 = (M1 * V1) / M2
Else
msgError = MsgBox("Enter numerical values into both Sub. 1 values and Sub. 2 Molarity.", vbCritical, "Entry Error") 'error message if numbers are not used
End If

txtMolaResult2.Text = V2 & " L" 'display result
End Sub
