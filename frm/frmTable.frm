VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTable 
   Caption         =   "Periodic Table"
   ClientHeight    =   10368
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   17556
   OleObjectBlob   =   "frmTable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnActin_Click()
Dim i, ele
Dim arr1(1 To 119)

i = 1
Do Until i = 119 'setting up arrays
    arr1(i) = Cells(i + 1, 2).Text 'array 1 is the symbols of each element, the second half of each button name
    i = i + 1 'Add to i to repeat process on new row
Loop 'There are 119 rows, so stop at the end

i = 1 'reset i value

Do Until i = 119
    ele = arr1(i) 'ele will be used to hold each element's symbol, the end of each button and label
    If frmTable.Controls("btn" & ele).BackColor = &HFFC0FF Then 'buttons with this color are for actinoids
        frmTable.Controls("btn" & ele).Visible = True 'make only them visible
        frmTable.Controls("lblAto" & ele).Visible = True 'also make sure their atomic masses are visible
        Else 'if the button is not of this color, it is a different class...
            frmTable.Controls("btn" & ele).Visible = False 'make them invisible
            frmTable.Controls("lblAto" & ele).Visible = False 'also make sure their atomic masses are invisible
End If
i = i + 1 'add 1 onto i to advance the array
Loop 'loop until the do requirement is satisfied
End Sub

Private Sub btnAlkali_Click()
Dim i, ele 'i will be used to add onto arrays, ele will store each element's symbol
Dim arr1(1 To 119) 'one array will be used for the element symbols

i = 1 'reset the value of i
Do Until i = 119 'setting up arrays until i reaches the end of the 118 rows
    arr1(i) = Cells(i + 1, 2).Text 'array 1 is the symbols of each element, the second half of each button name
    i = i + 1 'Add to i to repeat process on new row
Loop 'There are 119 rows, so stop at the end

i = 1 'reset i value

Do Until i = 119 'stop when i reaches the final row on the spreadsheet
    ele = arr1(i) 'ele will be used to hold each element's symbol, the end of each button and label
    If frmTable.Controls("btn" & ele).BackColor = &HC0C0& Then 'buttons with this color are for alkali metals
        frmTable.Controls("btn" & ele).Visible = True 'make only them visible
        frmTable.Controls("lblAto" & ele).Visible = True 'also make sure their atomic masses are visible
        Else 'if the button is not of this color, it is a different class...
            frmTable.Controls("btn" & ele).Visible = False 'make them invisible
            frmTable.Controls("lblAto" & ele).Visible = False 'also make sure their atomic masses are invisible
End If
i = i + 1 'add 1 onto i to advance the array
Loop 'loop until the do requirement is satisfied
End Sub

Private Sub btnAlkaline_Click()
Dim i, ele 'i will be used to add onto arrays, ele will store each element's symbol
Dim arr1(1 To 119) 'one array will be used for the element symbols

i = 1 'reset the value of i
Do Until i = 119 'setting up arrays until i reaches the end of the 118 rows
    arr1(i) = Cells(i + 1, 2).Text 'array 1 is the symbols of each element, the second half of each button name
    i = i + 1 'Add to i to repeat process on new row
Loop 'There are 119 rows, so stop at the end

i = 1 'reset i value

Do Until i = 119 'stop when i reaches the final row on the spreadsheet
    ele = arr1(i) 'ele will be used to hold each element's symbol, the end of each button and label
    If frmTable.Controls("btn" & ele).BackColor = &HFFFF& Then 'buttons with this color are for alkaline earth metals
        frmTable.Controls("btn" & ele).Visible = True 'make only them visible
        frmTable.Controls("lblAto" & ele).Visible = True 'also make sure their atomic masses are visible
        Else 'if the button is not of this color, it is a different class...
            frmTable.Controls("btn" & ele).Visible = False 'make them invisible
            frmTable.Controls("lblAto" & ele).Visible = False 'also make sure their atomic masses are invisible
End If
i = i + 1 'add 1 onto i to advance the array
Loop 'loop until the do requirement is satisfied
End Sub

Private Sub btnCalculate_Click()
frmCalculate.Show 'Open the calculation Userform
End Sub

Private Sub btnGases_Click()
Dim i, ele 'i will be used to add onto arrays, ele will store each element's symbol
Dim arr1(1 To 119) 'one array will be used for the element symbols

i = 1 'reset the value of i
Do Until i = 119 'setting up arrays until i reaches the end of the 118 rows
    arr1(i) = Cells(i + 1, 2).Text 'array 1 is the symbols of each element, the second half of each button name
    i = i + 1 'Add to i to repeat process on new row
Loop 'There are 119 rows, so stop at the end

i = 1 'reset i value

Do Until i = 119 'stop when i reaches the final row on the spreadsheet
    ele = arr1(i) 'ele will be used to hold each element's symbol, the end of each button and label
    If frmTable.Controls("btn" & ele).ForeColor = vbRed Then 'buttons with text of this color have a gaseous state
        frmTable.Controls("btn" & ele).Visible = True 'make only them visible
        frmTable.Controls("lblAto" & ele).Visible = True 'also make sure their atomic masses are visible
        Else 'if the button is not of this color, it is a different state...
            frmTable.Controls("btn" & ele).Visible = False 'make them invisible
            frmTable.Controls("lblAto" & ele).Visible = False 'also make sure their atomic masses are invisible
End If
i = i + 1 'add 1 onto i to advance the array
Loop 'loop until the do requirement is satisfied
End Sub

Private Sub btnH_Click()
frmHydrogen.Show 'Open the Hydrogen Userform
End Sub


Private Sub btnHe_Click()
frmHelium.Show 'Open the HeliumUserform
End Sub

Private Sub btnLanth_Click()
Dim i, ele 'i will be used to add onto arrays, ele will store each element's symbol
Dim arr1(1 To 119) 'one array will be used for the element symbols

i = 1 'reset the value of i
Do Until i = 119 'setting up arrays until i reaches the end of the 118 rows
    arr1(i) = Cells(i + 1, 2).Text 'array 1 is the symbols of each element, the second half of each button name
    i = i + 1 'Add to i to repeat process on new row
Loop 'There are 119 rows, so stop at the end

i = 1 'reset i value

Do Until i = 119 'stop when i reaches the final row on the spreadsheet
    ele = arr1(i) 'ele will be used to hold each element's symbol, the end of each button and label
    If frmTable.Controls("btn" & ele).BackColor = &H80C0FF Then 'buttons with this color are for lanthanoids
        frmTable.Controls("btn" & ele).Visible = True 'make only them visible
        frmTable.Controls("lblAto" & ele).Visible = True 'also make sure their atomic masses are visible
        Else 'if the button is not of this color, it is a different class...
            frmTable.Controls("btn" & ele).Visible = False 'make them invisible
            frmTable.Controls("lblAto" & ele).Visible = False 'also make sure their atomic masses are invisible
End If
i = i + 1 'add 1 onto i to advance the array
Loop 'loop until the do requirement is satisfied
End Sub

Private Sub btnLiquids_Click()
Dim i, ele 'i will be used to add onto arrays, ele will store each element's symbol
Dim arr1(1 To 119) 'one array will be used for the element symbols

i = 1 'reset the value of i
Do Until i = 119 'setting up arrays until i reaches the end of the 118 rows
    arr1(i) = Cells(i + 1, 2).Text 'array 1 is the symbols of each element, the second half of each button name
    i = i + 1 'Add to i to repeat process on new row
Loop 'There are 119 rows, so stop at the end

i = 1 'reset i value

Do Until i = 119 'stop when i reaches the final row on the spreadsheet
    ele = arr1(i) 'ele will be used to hold each element's symbol, the end of each button and label
    If frmTable.Controls("btn" & ele).ForeColor = vbBlue Then 'buttons with text of this color have a liquid state
        frmTable.Controls("btn" & ele).Visible = True 'make only them visible
        frmTable.Controls("lblAto" & ele).Visible = True 'also make sure their atomic masses are visible
        Else 'if the button is not of this color, it is a different state...
            frmTable.Controls("btn" & ele).Visible = False 'make them invisible
            frmTable.Controls("lblAto" & ele).Visible = False 'also make sure their atomic masses are invisible
End If
i = i + 1 'add 1 onto i to advance the array
Loop 'loop until the do requirement is satisfied
End Sub

Private Sub btnMetalloid_Click()
Dim i, ele 'i will be used to add onto arrays, ele will store each element's symbol
Dim arr1(1 To 119) 'one array will be used for the element symbols

i = 1 'reset the value of i
Do Until i = 119 'setting up arrays until i reaches the end of the 118 rows
    arr1(i) = Cells(i + 1, 2).Text 'array 1 is the symbols of each element, the second half of each button name
    i = i + 1 'Add to i to repeat process on new row
Loop 'There are 119 rows, so stop at the end

i = 1 'reset i value

Do Until i = 119 'stop when i reaches the final row on the spreadsheet
    ele = arr1(i) 'ele will be used to hold each element's symbol, the end of each button and label
    If frmTable.Controls("btn" & ele).BackColor = &HC0FFC0 Then 'buttons with this color are for metalloids
        frmTable.Controls("btn" & ele).Visible = True 'make only them visible
        frmTable.Controls("lblAto" & ele).Visible = True 'also make sure their atomic masses are visible
        Else 'if the button is not of this color, it is a different class...
            frmTable.Controls("btn" & ele).Visible = False 'make them invisible
            frmTable.Controls("lblAto" & ele).Visible = False 'also make sure their atomic masses are invisible
End If
i = i + 1 'add 1 onto i to advance the array
Loop 'loop until the do requirement is satisfied
End Sub

Private Sub btnNoble_Click()
Dim i, ele 'i will be used to add onto arrays, ele will store each element's symbol
Dim arr1(1 To 119) 'one array will be used for the element symbols

i = 1 'reset the value of i
Do Until i = 119 'setting up arrays until i reaches the end of the 118 rows
    arr1(i) = Cells(i + 1, 2).Text 'array 1 is the symbols of each element, the second half of each button name
    i = i + 1 'Add to i to repeat process on new row
Loop 'There are 119 rows, so stop at the end

i = 1 'reset i value

Do Until i = 119 'stop when i reaches the final row on the spreadsheet
    ele = arr1(i) 'ele will be used to hold each element's symbol, the end of each button and label
    If frmTable.Controls("btn" & ele).BackColor = &HFFFF00 Then 'buttons with this color are for noble gases
        frmTable.Controls("btn" & ele).Visible = True 'make only them visible
        frmTable.Controls("lblAto" & ele).Visible = True 'also make sure their atomic masses are visible
        Else 'if the button is not of this color, it is a different class...
            frmTable.Controls("btn" & ele).Visible = False 'make them invisible
            frmTable.Controls("lblAto" & ele).Visible = False 'also make sure their atomic masses are invisible
End If
i = i + 1 'add 1 onto i to advance the array
Loop 'loop until the do requirement is satisfied
End Sub

Private Sub btnO_Click()
frmOxygen.Show 'Open the Oxygen Userform
End Sub

Private Sub btnOther_Click()
Dim i, ele 'i will be used to add onto arrays, ele will store each element's symbol
Dim arr1(1 To 119) 'one array will be used for the element symbols

i = 1 'reset the value of i
Do Until i = 119 'setting up arrays until i reaches the end of the 118 rows
    arr1(i) = Cells(i + 1, 2).Text 'array 1 is the symbols of each element, the second half of each button name
    i = i + 1 'Add to i to repeat process on new row
Loop 'There are 119 rows, so stop at the end

i = 1 'reset i value

Do Until i = 119 'stop when i reaches the final row on the spreadsheet
    ele = arr1(i) 'ele will be used to hold each element's symbol, the end of each button and label
    If frmTable.Controls("btn" & ele).BackColor = &HFF00& Then 'buttons with this color are for other nonmetals
        frmTable.Controls("btn" & ele).Visible = True 'make only them visible
        frmTable.Controls("lblAto" & ele).Visible = True 'also make sure their atomic masses are visible
        Else 'if the button is not of this color, it is a different class...
            frmTable.Controls("btn" & ele).Visible = False 'make them invisible
            frmTable.Controls("lblAto" & ele).Visible = False 'also make sure their atomic masses are invisible
End If
i = i + 1 'add 1 onto i to advance the array
Loop 'loop until the do requirement is satisfied
End Sub

Private Sub btnPostTrans_Click()
Dim i, ele 'i will be used to add onto arrays, ele will store each element's symbol
Dim arr1(1 To 119) 'one array will be used for the element symbols

i = 1 'reset the value of i
Do Until i = 119 'setting up arrays until i reaches the end of the 118 rows
    arr1(i) = Cells(i + 1, 2).Text 'array 1 is the symbols of each element, the second half of each button name
    i = i + 1 'Add to i to repeat process on new row
Loop 'There are 119 rows, so stop at the end

i = 1 'reset i value

Do Until i = 119 'stop when i reaches the final row on the spreadsheet
    ele = arr1(i) 'ele will be used to hold each element's symbol, the end of each button and label
    If frmTable.Controls("btn" & ele).BackColor = &H80FFFF Then 'buttons with this color are for post-transition metals
        frmTable.Controls("btn" & ele).Visible = True 'make only them visible
        frmTable.Controls("lblAto" & ele).Visible = True 'also make sure their atomic masses are visible
        Else 'if the button is not of this color, it is a different class...
            frmTable.Controls("btn" & ele).Visible = False 'make them invisible
            frmTable.Controls("lblAto" & ele).Visible = False 'also make sure their atomic masses are invisible
End If
i = i + 1 'add 1 onto i to advance the array
Loop 'loop until the do requirement is satisfied
End Sub

Private Sub btnSearch_Click()
Dim i, j, search, sym, ele 'i and j will be used to add onto arrays, ele and sym will store each element's symbol, search will be for an inputbox
Dim arr1(1 To 119), arr2(1 To 119), arr3(1 To 119) 'one array will be used for the element names, one for symbols, one for atomic numbers

i = 1 'set up initial i value
j = 1 'set up initial j value
Do Until i = 119 And j = 119 'setting up arrays
    arr1(i) = Cells(i + 1, 1).Text 'array 1 is the names of each element
    arr1(j) = Cells(j + 1, 1).Text 'array 1 is the names of each element
    arr2(i) = Cells(i + 1, 2).Text 'array 2 is the symbols of each element
    arr2(j) = Cells(j + 1, 2).Text 'array 2 is the symbols of each element
    arr3(i) = Cells(i + 1, 3).Text 'array 3 is the atomic numbers of each element
    arr3(j) = Cells(j + 1, 3).Text 'array 3 is the atomic numbers of each element
    i = i + 1 'Add to i to repeat process on new row
    j = j + 1
Loop 'There are 119 rows, so stop at the end

search = InputBox("Enter an element's symbol, name, or atomic number.") 'create an instruction over the input box

i = 1 'reset i value

Do Until sym = Cells(i + 1, 2).Text Or i > 119
    If search = arr1(i) Or search = arr2(i) Or search = arr3(i) Then 'If the value entered is equal to a value in any of the three arrays...
        sym = arr2(i) 'sym will be used as the "correct" symbol, the one entered
        j = 1 'reset j value
        Do Until j = 119 'j will be added onto like i, so set the appropriate loop condition
        ele = arr2(j) 'ele will be the symbol tested against sym
        If frmTable.Controls("btn" & ele).Caption = sym Then 'when the button tested in the loop has a caption equal to the one entered...
            frmTable.Controls("btn" & ele).Visible = True 'make only it visible
            frmTable.Controls("lblAto" & ele).Visible = True 'also make sure its atomic mass is visible
            ElseIf sym = "" Then 'if nothing was entered, don't do anything
                Else 'if the button is not of this element, it should not show...
                frmTable.Controls("btn" & ele).Visible = False 'make it invisible
                frmTable.Controls("lblAto" & ele).Visible = False 'also make sure its atomic mass is invisible
            End If
            j = j + 1 'add 1 onto j to advance the array
        Loop
        End If
i = i + 1 'add 1 onto i to advance the arrays
Loop
End Sub

Private Sub btnSolid_Click()
Dim i, ele 'i will be used to add onto arrays, ele will store each element's symbol
Dim arr1(1 To 119) 'one array will be used for the element symbols

i = 1 'reset the value of i
Do Until i = 119 'setting up arrays until i reaches the end of the 118 rows
    arr1(i) = Cells(i + 1, 2).Text 'array 1 is the symbols of each element, the second half of each button name
    i = i + 1 'Add to i to repeat process on new row
Loop 'There are 119 rows, so stop at the end

i = 1 'reset i value

Do Until i = 119 'stop when i reaches the final row on the spreadsheet
    ele = arr1(i) 'ele will be used to hold each element's symbol, the end of each button and label
    If frmTable.Controls("btn" & ele).ForeColor = vbBlack Then 'buttons with text of this color have a solid state
        frmTable.Controls("btn" & ele).Visible = True 'make only them visible
        frmTable.Controls("lblAto" & ele).Visible = True 'also make sure their atomic masses are visible
        Else 'if the button is not of this color, it is a different state...
            frmTable.Controls("btn" & ele).Visible = False 'make them invisible
            frmTable.Controls("lblAto" & ele).Visible = False 'also make sure their atomic masses are invisible
End If
i = i + 1 'add 1 onto i to advance the array
Loop 'loop until the do requirement is satisfied

End Sub

Private Sub btnTemp_Click()
Dim i, ele, temp As Double, mp As Double, bp As Double 'i will be used to add onto arrays, ele will store each element's symbol, the rest hold numbers
Dim arr1(1 To 119), arr2(1 To 119), arr3(1 To 119) 'one array will be used for the element symbols, one for melting points, one for boiling points
Dim msgError 'in case of an error, a message box will be used

i = 1 'set initial value for i
Do Until i = 119 'setting up arrays
    arr1(i) = Cells(i + 1, 2).Text 'array 1 is the symbols of each element, the second half of each button name
    arr2(i) = Cells(i + 1, 8).Text 'array 2 is the melting point of each element
    arr3(i) = Cells(i + 1, 9).Text 'array 3 is the boiling point of each element
    i = i + 1 'Add to i to repeat process on new row
Loop 'There are 119 rows, so stop at the end

i = 1 'reset i value

If IsNumeric(txtTemp.Text) Then 'if the value entered into the temperature's text box is numeric...
temp = txtTemp.Text 'assign the temp variable to it
Do Until i = 119 'stop when i reaches the final row on the spreadsheet
    ele = arr1(i) 'make ele equal to the symbol array
    If arr2(i) <> "Uncertain" Then 'if the melting point is uncertain, mp cannot be compared with temp
    mp = arr2(i) 'it that isn't the case, make it equal to the melting point array
    End If
    If arr3(i) <> "Uncertain" Then 'if the boiling point is uncertain, bp cannot be compared with temp
    bp = arr3(i) 'it that isn't the case, make it equal to the boiling point array
    End If
    If arr2(i) = "Uncertain" Then 'if the melting point is uncertain, the state of the element cannot be determined
        frmTable.Controls("btn" & ele).ForeColor = vbMagenta 'make it magenta, the color of the unknown
        ElseIf temp < mp Then 'if temp is less than the melting point, it will be solid
            frmTable.Controls("btn" & ele).ForeColor = vbBlack 'make it black, the color of the solid
            ElseIf temp >= mp And arr3(i) = "Uncertain" Then 'if temp is greater than the melting point but the boiling point is unknown, the state of the element cannot be determined
                frmTable.Controls("btn" & ele).ForeColor = vbMagenta 'make it magenta, the color of the unknown
                ElseIf temp >= mp And temp < bp Then 'if temp is greater than the melting point and less than the boiling point, it will be liquid
                    frmTable.Controls("btn" & ele).ForeColor = vbBlue 'make it blue, the color of the liquid
                    ElseIf temp >= mp And temp >= bp Then 'if temp is greater than the melting point and less than the boiling point, it will be gaseous
                        frmTable.Controls("btn" & ele).ForeColor = vbRed 'make it red, the color of the gas
                        Else 'if, for whatever reason, none of this applies to the element, the state cannot be determined
                            frmTable.Controls("btn" & ele).ForeColor = vbMagenta 'make it magenta, the color of the unknown
End If
i = i + 1 'add 1 onto i to advance the arrays
Loop 'loop until the do requirement is satisfied

Else 'if the text box contains characters other than numbers...
    msgError = MsgBox("Sorry, you can only enter numbers.", vbCritical, "Entry Error") 'an error message box will appear
End If

End Sub

Private Sub btnTempReset_Click()
Dim i, ele
Dim arr1(1 To 119), arr2(1 To 119)

i = 1
Do Until i = 119 'setting up arrays
    arr1(i) = Cells(i + 1, 2).Text 'array 1 is the symbols of each element, the second half of each button name
    arr2(i) = Cells(i + 1, 6).Text 'array 2 is the color of each button
    i = i + 1 'Add to I1 to repeat process on new row
Loop 'There are 119 rows, so stop at the end

i = 1 'reset i value

Do Until i = 119
    ele = arr1(i)
    frmTable.Controls("btn" & ele).Visible = True
    frmTable.Controls("lblAto" & ele).Visible = True
    If arr2(i) = "Black" Then
      frmTable.Controls("btn" & ele).ForeColor = vbBlack
      ElseIf arr2(i) = "Blue" Then
        frmTable.Controls("btn" & ele).ForeColor = vbBlue
        ElseIf arr2(i) = "Red" Then
          frmTable.Controls("btn" & ele).ForeColor = vbRed
          ElseIf arr2(i) = "Gray" Then
             frmTable.Controls("btn" & ele).ForeColor = vbMagenta
End If
i = i + 1
Loop 'loop until the do requirement is satisfied

End Sub

Private Sub btnTrans_Click()
Dim i, ele 'i will be used to add onto arrays, ele will store each element's symbol
Dim arr1(1 To 119) 'one array will be used for the element symbols

i = 1 'reset the value of i
Do Until i = 119 'setting up arrays until i reaches the end of the 118 rows
    arr1(i) = Cells(i + 1, 2).Text 'array 1 is the symbols of each element, the second half of each button name
    i = i + 1 'Add to i to repeat process on new row
Loop 'There are 119 rows, so stop at the end

i = 1 'reset i value

Do Until i = 119 'stop when i reaches the final row on the spreadsheet
    ele = arr1(i) 'ele will be used to hold each element's symbol, the end of each button and label
    If frmTable.Controls("btn" & ele).BackColor = &HC0C0FF Then 'buttons with this color are for transition metals
        frmTable.Controls("btn" & ele).Visible = True 'make only them visible
        frmTable.Controls("lblAto" & ele).Visible = True 'also make sure their atomic masses are visible
        Else 'if the button is not of this color, it is a different class...
            frmTable.Controls("btn" & ele).Visible = False 'make them invisible
            frmTable.Controls("lblAto" & ele).Visible = False 'also make sure their atomic masses are invisible
End If
i = i + 1 'add 1 onto i to advance the array
Loop 'loop until the do requirement is satisfied
End Sub

Private Sub btnUnkClass_Click()
Dim i, ele 'i will be used to add onto arrays, ele will store each element's symbol
Dim arr1(1 To 119) 'one array will be used for the element symbols

i = 1 'reset the value of i
Do Until i = 119 'setting up arrays until i reaches the end of the 118 rows
    arr1(i) = Cells(i + 1, 2).Text 'array 1 is the symbols of each element, the second half of each button name
    i = i + 1 'Add to i to repeat process on new row
Loop 'There are 119 rows, so stop at the end

i = 1 'reset i value

Do Until i = 119 'stop when i reaches the final row on the spreadsheet
    ele = arr1(i) 'ele will be used to hold each element's symbol, the end of each button and label
    If frmTable.Controls("btn" & ele).BackColor = &HE0E0E0 Then 'buttons with this color are for unknown classes
        frmTable.Controls("btn" & ele).Visible = True 'make only them visible
        frmTable.Controls("lblAto" & ele).Visible = True 'also make sure their atomic masses are visible
        Else 'if the button is not of this color, it is a different class...
            frmTable.Controls("btn" & ele).Visible = False 'make them invisible
            frmTable.Controls("lblAto" & ele).Visible = False 'also make sure their atomic masses are invisible
End If
i = i + 1 'add 1 onto i to advance the array
Loop 'loop until the do requirement is satisfied
End Sub

Private Sub btnUnkState_Click()
Dim i, ele 'i will be used to add onto arrays, ele will store each element's symbol
Dim arr1(1 To 119) 'one array will be used for the element symbols

i = 1 'reset the value of i
Do Until i = 119 'setting up arrays until i reaches the end of the 118 rows
    arr1(i) = Cells(i + 1, 2).Text 'array 1 is the symbols of each element, the second half of each button name
    i = i + 1 'Add to i to repeat process on new row
Loop 'There are 119 rows, so stop at the end

i = 1 'reset i value

Do Until i = 119 'stop when i reaches the final row on the spreadsheet
    ele = arr1(i) 'ele will be used to hold each element's symbol, the end of each button and label
    If frmTable.Controls("btn" & ele).ForeColor = vbMagenta Then 'buttons with text of this color have an unknown state
        frmTable.Controls("btn" & ele).Visible = True 'make only them visible
        frmTable.Controls("lblAto" & ele).Visible = True 'also make sure their atomic masses are visible
        Else 'if the button is not of this color, it is a different state...
            frmTable.Controls("btn" & ele).Visible = False 'make them invisible
            frmTable.Controls("lblAto" & ele).Visible = False 'also make sure their atomic masses are invisible
End If
i = i + 1 'add 1 onto i to advance the array
Loop 'loop until the do requirement is satisfied
End Sub

