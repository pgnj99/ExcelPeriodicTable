VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCompound 
   Caption         =   "Ionic Compounds"
   ClientHeight    =   3828
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4428
   OleObjectBlob   =   "frmCompound.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCompound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCompCalc_Click()
Dim E1 As String, E2 As String, C1 As Integer, C2 As Integer, N1 As Integer, N2 As Integer 'units to be entered: the E variables are for element symbols, the C variables are for initial charges, and the N variables are for edited C variables
Dim arr1(1 To 119), arr2(1 To 119) 'these arrays will be used for holding words
Dim arr3(1 To 119) 'this array will be used for holding and adding whole numbers
Dim i 'dim variable to add onto in arrays
Dim msgError 'error message box in case of a misinput

i = 1 'set i to first value

'The first step is to set the three arrays up
'Note: For arrays, the row value is "i + 1" because the first row in the spreadsheet is used for labels
Do 'setting up arrays
    arr1(i) = Cells(i + 1, 1).Text 'array 1 is the names of each element
    arr2(i) = Cells(i + 1, 2).Text 'array 2 is the symbols of each element
    arr3(i) = Cells(i + 1, 7).Text 'array 3 is the charges of each element
    i = i + 1 'Add to i to repeat process on new row
Loop Until i = 119 'There are 119 rows, so stop at the end

'Now the elements' text boxes will be set up for entry
'The program cannot commence unless both elements are entered, so an error will be created if otherwise
If txtCompEle1.Text = "" Or txtCompEle2.Text = "" Then 'If an element is not entered into a text box...
    msgError = MsgBox("Please enter an element into both text boxes.", vbCritical, "Entry Error") 'Create an error message because it cannot continue
    Else 'Assuming both elements have been entered, continue
    
    i = 1 'reset i value
Do Until E1 = arr2(i) 'the goal is to find the equivalent on the spreadsheet to the first element entered
If txtCompEle1.Text = arr1(i) Or txtCompEle1.Text = arr2(i) Then 'when the value entered into the first element's text box is equal to the name of an element or symbol on the spreadsheet
        E1 = arr2(i) '... make the E1 variable equal to the row's element's symbol
    Else: i = i + 1 'Add to i to repeat process on new row if it is not equal to a matching row number
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into text boxes.", vbCritical, "Entry Error") 'therefore, an error message box will appear
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied


If txtCompChar1.Text = "" Then 'Assuming there is nothing entered into the first charge text box, a default must be chosen
If E1 = arr2(i) Then 'when the value of E1 is equal to the name of an element on the spreadsheet, assuming no error occurred
    If arr3(i) <> "No" Then 'if the value of column 7 on that row is not "No", then...
        C1 = arr3(i) '... make C1 equal to the number listed
        Else: msgError = MsgBox("Sorry, the first element does not have a default ionic charge. Please enter one.", vbCritical, "Entry Error") 'otherwise, an error message will appear because there is no charge
    End If
    Else 'assuming E1 is not equal to anything in the array
End If
ElseIf IsNumeric(txtCompChar1.Text) Then 'if the user does enter a number into the charge's text box...
    C1 = txtCompChar1.Text '... simply make the variable equal to it
    End If
    

    i = 1 'reset i value
Do Until E2 = arr2(i) 'the goal is to find the equivalent on the spreadsheet to the second element entered
If txtCompEle2.Text = arr1(i) Or txtCompEle2.Text = arr2(i) Then 'when the value entered into the second element's text box is equal to the name of an element or symbol on the spreadsheet
        E2 = arr2(i) '... make the E2 variable equal to the row's element's symbol
    Else: i = i + 1 'Add to i to repeat process on new row if it is not equal to a matching row number
If i > 119 Then 'i will pass 119 if the value entered is nowhere in the two rows
    msgError = MsgBox("Sorry, you must enter a valid element into text boxes.", vbCritical, "Entry Error") 'therefore, an error message box will appear
    Else: End If 'if the process runs without error, end it normally
End If 'end the process
Loop 'loop until the do requirement is satisfied


If txtCompChar2.Text = "" Then 'Assuming there is nothing entered into the second charge text box, a default must be chosen
If E2 = arr2(i) Then 'when the value of E2 is equal to the name of an element on the spreadsheet, assuming no error occurred
    If arr3(i) <> "No" Then 'if the value of column 7 on that row is not "No", then...
        C2 = arr3(i) '... make C2 equal to the number listed
        Else: msgError = MsgBox("Sorry, the first element does not have a default ionic charge. Please enter one.", vbCritical, "Entry Error") 'otherwise, an error message will appear because there is no charge
    End If
    Else 'assuming E2 is not equal to anything in the array
End If
ElseIf IsNumeric(txtCompChar2.Text) Then 'if the user does enter a number into the charge's text box...
    C2 = txtCompChar2.Text '... simply make the variable equal to it
    End If

End If

'Now the charge values must be adjusted to be shown in the results text box
'If we're to assume these are molecular compounds, the sum of their charges must equal 0, so they cannot both have positive or negative changes
If C1 >= 1 And C2 >= 1 Then 'If both charges are positive...
    msgError = MsgBox("Sorry, but atoms don't normally form molecular compounds unless the sum of all charges is 0. You have two positives.", vbCritical, "Entry Error") '... an error message box will appear
    ElseIf C1 < 0 And C2 < 0 Then 'If both charges are negative...
        msgError = MsgBox("Sorry, but atoms don't normally form molecular compounds unless the sum of all charges is 0. You have two negatives.", vbCritical, "Entry Error") '... an error message box will appear
        Else
End If

'The charges also cannot equal 0, as elements with this charge generally don't bond very easily, so create an error message box if either charge is 0
If C1 = 0 Then
    msgError = MsgBox("Sorry, you usually won't have much luck with an element that has a charge of 0.", vbCritical, "Entry Error")
    End If

If C2 = 0 Then
    msgError = MsgBox("Sorry, you usually won't have much luck with an element that has a charge of 0.", vbCritical, "Entry Error")
    End If

'If both elements have an equal charge, then in order to display this properly, they can both be simplified to 1
If C1 = -C2 Then 'If one charge is equal to the negative value of the other charge, make both equal to 1
    C1 = 1
    C2 = 1
End If

'Now that all the equations are done and all that's left is the display, we need to make sure the charge numbers are positive
If C1 < 0 Then 'If the first charge is negative...
    N1 = -C1 'Make this new variable N1 equal to its negative variant
    Else 'Otherwise, if it is positive...
        N1 = C1 'Make N1 equal to the positive variant
End If

If C2 < 0 Then 'If the second charge is negative...
    N2 = -C2 'Make this new variable N2 equal to its negative variant
    Else 'Otherwise, if it is positive...
        N2 = C2 'Make N2 equal to the positive variant
End If

'It's time for the result to be displayed, but one thing to note is that if an element in a compound has a ratio of 1, the 1 is not shown
'Therefore, the result must be tailored to not display a 1
If N1 = 1 And N2 = 1 Then 'If both charges equal 1...
    txtCompResult = E1 & E2 'Display the two elements as a result but do not display any charges
ElseIf N1 = 1 And N2 <> 1 Then 'If only the first charge equals 1...
    txtCompResult = E1 & N2 & E2 'Display the two elements as a result but only display the second charge
ElseIf N1 <> 1 And N2 = 1 Then 'If only the second charge equals 1...
    txtCompResult = E1 & E2 & N1 'Display the two elements as a result but only display the first charge
Else 'If there is no 1 to be found...
    txtCompResult = E1 & N2 & E2 & N1 'Display everything normally as a result
End If

End Sub
