' SPDX-License-Identifier: Apache-2.0
'
' Copyright 2024 Lou Lerner
'
' This source code is licensed under the Apache License, Version 2.0.
' A copy of the license is included in the root directory of this project as LICENSE.txt.

Function CLD(ModalityValuesRange As Range, TestResultsRange As Range, ModalityCell As Range, Optional Alpha As Variant, Optional Order As Variant)

'------------------------------------------------------------
On Error GoTo E:

Dim Failure As Boolean
'------------------------------------------------------------

' Variables declaration

Dim SortOrder As String

Dim i As Long, j As Long, k As Long, l As Long

Dim Key As Variant

Dim ModalityValuesArray() As Variant

Dim RowsNbr As Long, ColumnsNbr As Long
Dim EmptyRow As Boolean, EmptyColumn As Boolean

Dim TempArray1 As Variant
Dim TempArray2 As Variant

Dim Matrix() As Variant
Dim MatrixSort() As Variant

Dim DifferencesList As Object

Dim Counter As Integer

Dim TreatmentCounter As Integer
Dim PresentTreatmentCounter As Integer

Dim LettersArray() As Variant

Dim TargetModality As Integer
Dim ModalityLetters As String

Dim ModalityPresent As Boolean

' Configuration of the "Alpha" argument which allows the user to choose a pvalue threshold

    ' If the user did not specify the alpha -> default valut (5%, meaning pvalue of 0,05)

    If IsMissing(Alpha) Then
    
        Alpha = 5 ' Default value
        
    End If

    ' Else check if the specified value is between 1 and 100 and is numeric.

    If Not IsNumeric(Alpha) Or Alpha < 1 Or Alpha > 100 Then
    
        CLD = "" ' Function returns an empty cell if Alpha is incorrect
        
        Exit Function
        
    End If
    
' Configuration of the "Order" argument which allows user to choose if the highest modalities get the lowest letters in the alphabet (descending order)

    ' Check if "Order" is present and assign "descending" as default value if necessary.
    
    If IsMissing(Order) Then
    
        Order = "descending"
        
    End If

    ' Check the "Order" argument value and assign appropriate value to the "SortOrder" variable
    
    Select Case UCase(Order)
    
        Case "DESCENDING"
        
            SortOrder = "descending"
            
        Case "NORMAL"
        
            SortOrder = "normal"
            
        Case Else
        
            ' If the Order argument is not "descending" or "normal", function returns empty cell
            
            CLD = ""
            
            Exit Function
            
    End Select

' Create a 2 column array "ModalityValuesArray" (defined by user with 1st argument) containing the list of modalities and their means

ModalityValuesArray = ModalityValuesRange.Value
   
' Create a "Matrix" array containing the table indicating the pair-wise post-hoc results for each pair of modalities (defined by the 2nd argument of the function)

    ' Define array from range
 
    Matrix() = TestResultsRange.Value
    
    ' Replace empty values under the diagonal with the tag "EmptyCell", which will be usefull later to ignore these cells
    
    For i = 2 To UBound(Matrix, 1)
        For j = 2 To UBound(Matrix, 2) - (UBound(Matrix, 2) - i)
            Matrix(i, j) = "EmptyCell"
        Next j
    Next i

' Delete all modalities which have not been tested: this allows the function to work even with partial data

RowsNbr = UBound(Matrix, 1)
ColumnsNbr = UBound(Matrix, 2)

    Do
        EmptyRow = False
        
        EmptyColumn = False
        
        For i = RowsNbr To 2 Step -1
        
            ' Check if i row is empty
            
            EmptyRow = True
            
            For j = 2 To ColumnsNbr
            
                If Matrix(i, j) <> "" And Matrix(i, j) <> "EmptyCell" Then
                
                    EmptyRow = False
                    
                    Exit For
                    
                End If
                
            Next j
            
            ' Check if corresponding column is also empty
            
            If EmptyRow Then
            
                EmptyColumn = True
                
                For j = 2 To ColumnsNbr
                
                    If Matrix(j, i) <> "" And Matrix(j, i) <> "EmptyCell" Then
                    
                        EmptyColumn = False
                        
                        Exit For
                        
                    End If
                    
                Next j
                
            End If
            
            ' Remove row and corresponding column if both are empty
            
            If EmptyRow And EmptyColumn Then
            
                Matrix = RemoveRow(Matrix, i) ' Use of the RemoveRow and RemoveColum functions to delete a column/row from array without leaving empty space
                
                Matrix = RemoveColumn(Matrix, i)
                
                ModalityValuesArray = RemoveRow(ModalityValuesArray, i - 1) ' Also delete the corresponding modality in "ModalityValuesArray" so that it is considered missing
                
                RowsNbr = RowsNbr - 1
                
                ColumnsNbr = ColumnsNbr - 1

                EmptyRow = True
                
            End If
  
        Next i
        
    Loop While EmptyRow Or EmptyColumn

' If the "Order" argument is defined to "normal", do not modify the modality order

If SortOrder = "normal" Then

' If the "Order" argument is defined to "descending", sort the array in descending order, depending on the values of the modalities.
' This later allow the "a" letter to be on the highest modality for example.

ElseIf SortOrder = "descending" Then

    For i = LBound(ModalityValuesArray) To UBound(ModalityValuesArray) - 1
    
        For j = i + 1 To UBound(ModalityValuesArray)
        
            If ModalityValuesArray(i, 2) > ModalityValuesArray(j, 2) Then
            
                TempArray1 = ModalityValuesArray(j, 1)
                TempArray2 = ModalityValuesArray(j, 2)
                
                ModalityValuesArray(j, 1) = ModalityValuesArray(i, 1)
                ModalityValuesArray(j, 2) = ModalityValuesArray(i, 2)
                
                ModalityValuesArray(i, 1) = TempArray1
                ModalityValuesArray(i, 2) = TempArray2
                
            End If
            
        Next j
        
    Next i
    
    ' Sort the matrix depending on the sorted array

        ' Create a copy of "Matrix" named "MatrixSort" and sort it with "ModalityValuesArray"
    
        ReDim MatrixSort(1 To UBound(Matrix, 1), 1 To UBound(Matrix, 2))
            
            For i = 2 To UBound(Matrix, 1)
            
                For k = LBound(ModalityValuesArray) To UBound(ModalityValuesArray)
                
                    If Matrix(i, 1) = ModalityValuesArray(k, 1) Then
                    
                        MatrixSort(k + 1, 1) = Matrix(i, 1)
                        
                        Exit For
                        
                    End If
                    
                Next k
                
                For j = 2 To UBound(Matrix, 2)
                
                    For l = LBound(ModalityValuesArray) To UBound(ModalityValuesArray)
                    
                        If Matrix(1, j) = ModalityValuesArray(l, 1) Then
                        
                            MatrixSort(1, l + 1) = Matrix(1, j)
                            
                            Exit For
                            
                        End If
                        
                    Next l
                    
                    MatrixSort(k + 1, l + 1) = Matrix(i, j)
                    
                Next j
                
            Next i
            
        ' Replace the content of "Matrix" with the content of "MatrixSort"
        
        For i = LBound(Matrix, 1) To UBound(Matrix, 1)
        
            For j = LBound(Matrix, 2) To UBound(Matrix, 2)
            
                Matrix(i, j) = MatrixSort(i, j)
                
            Next j
            
        Next i
        
End If

' Create a scripting dictionary to list the significant differences. Each key element of the list
' includes a modality and all the modalities not linked to it (listed in an array)

Set DifferencesList = CreateObject("Scripting.Dictionary")

    ' Populate the differences list by reading the "Matrix" array
    
    For i = 2 To UBound(Matrix, 1) ' Read each line of the matrix
   
            For j = 2 To UBound(Matrix, 2) ' read each column of the matrix
            
                ' If one cell is empty, the function returns an empty value
                
                If Matrix(i, j) = "" Then
                
                    CLD = ""
                    
                    Exit Function
                   
                ' If the observed value is "EmptyCell", do nothing because it means we are out of zone
            
                ElseIf Matrix(i, j) = "EmptyCell" Then
            
                ' If a modality is significantly different from another, add both modalities in a dictionary key
            
                ElseIf (IsNumeric(Matrix(i, j)) And Matrix(i, j) <= Alpha / 100) Or (Matrix(i, j) = "<0,001" Or _
                                                                                    Matrix(i, j) = "<0,005" Or _
                                                                                    Matrix(i, j) = "<0,05" Or _
                                                                                    Matrix(i, j) = "<0,01" Or _
                                                                                    Matrix(i, j) = "< 0,001" Or _
                                                                                    Matrix(i, j) = "< 0,005" Or _
                                                                                    Matrix(i, j) = "< 0,05" Or _
                                                                                    Matrix(i, j) = "< 0,01") Then
                
                    Counter = DifferencesList.Count + 1
                    
                    DifferencesList.Add Counter, Array()
                    
                    ' Use of the "PopulateArray" function to add an element to a dictionary key
                    
                    DifferencesList(Counter) = PopulateArray(DifferencesList(Counter), Matrix(i, 1))

                    DifferencesList(Counter) = PopulateArray(DifferencesList(Counter), Matrix(1, j))
                    
                ' If a modality is not significantly different from another, do nothing
                    
                ElseIf IsNumeric(Matrix(i, j)) And Matrix(i, j) > Alpha / 100 Then
                
                ' If cell is filled with unexpecter value, function returns an empty value
                                
                Else
                
                    CLD = ""
                    
                    Exit Function
                    
                End If
   
            Next j
       
    Next i
      
' "Insert & Absorb" phase: create an array with common "1s" shared between modalities in different columns

    ' Create the first column with only 1s.
    
    ReDim LettersArray(1 To UBound(ModalityValuesArray, 1), 1 To 1)
       
    For i = 1 To UBound(LettersArray)

    LettersArray(i, 1) = "1"
    
    Next i
        
    ' For each pair of significantly different modalities, insert and aborb
        
    For Each Key In DifferencesList.Keys ' For each  pair

        i = 1
    
        ' Check for each column if it verifies the difference between the two modalities
    
        Do Until i = UBound(LettersArray, 2) + 1 ' For each column
        
                ' If the assertion is false, insert and absorb
  
                If AssertionVerified(DifferencesList(Key), LettersArray, i, ModalityValuesArray) = False Then ' Use of the "AssertionVerified" function to check if assertion is true
                
                    InsertAbsorb DifferencesList(Key), LettersArray, i, ModalityValuesArray ' Use of the "InsertAbsorb" function
                                       
                    i = 1
                    
                Else
                
                    i = i + 1
                
                End If
        Loop

    Next Key

' "Sweeping" phase: after the previous step, some useless "1s" remain, and must be deleted
    
For i = 1 To UBound(LettersArray, 2) ' For each column

    For j = 1 To UBound(LettersArray, 1) ' For each modality (cell)
        
        If LettersArray(j, i) = 0 Then ' If cell value is 0, do nothing
        
        Else ' If cell value is not 0
        
            TreatmentCounter = 1
            PresentTreatmentCounter = 1
            
            For k = 1 To UBound(LettersArray, 1) ' For each modality to compare (row)
            
                If k = j Then ' (do not compare modality with itself)
                
                Else
                
                    If LettersArray(k, i) = 1 Then ' If the treatment's AND the treatment to be compared with cell values are 1
                                            
                        TreatmentCounter = TreatmentCounter + 1
                                            
                        For l = 1 To UBound(LettersArray, 2) ' For each column
                        
                            If l = i Then ' (Do not compare the column to itself)
                            
                            Else
                        
                                If LettersArray(j, l) = 1 And LettersArray(k, l) = 1 Then ' If both treatments are in an other column
                                
                                    PresentTreatmentCounter = PresentTreatmentCounter + 1 ' Count
                                
                                    Exit For
                                
                                End If
                                
                            End If
                        
                        Next l
                    
                    End If
                
                End If
                         
            Next k
            
            If PresentTreatmentCounter = 1 And TreatmentCounter = 1 Then ' If there is only one "1"
            
                For m = 1 To UBound(LettersArray, 2) ' For each column
                
                    If m = i Then ' (no not compare the column to itself)
                    
                    Else
                    
                        If LettersArray(j, m) = 1 Then ' If an other "1" is present in the same row, the isolated "1" can be deleted
                        
                            LettersArray(j, i) = 0
                            
                            Exit For
                            
                        End If
                    
                    End If
                                    
                Next m
        
            ElseIf PresentTreatmentCounter = TreatmentCounter Then ' If all equalities are already indicated in an other column, the "1" can be deleted
            
                LettersArray(j, i) = 0
                
            End If
            
            PresentTreatmentCounter = 1 ' Reset counter
        
        End If

    Next j

Next i

' Empty columns removal: after sweeping, some columns have no "1s" and must be removed

i = 1

Do Until i > UBound(LettersArray, 2) ' For each column

    TreatmentCounter = 0

    For j = 1 To UBound(LettersArray, 1) ' For each modality (row)

        If LettersArray(j, i) = 1 Then ' If cell equals 1

            TreatmentCounter = TreatmentCounter + 1 ' Count

        End If

    Next j

    If TreatmentCounter = 0 Then ' If column contains no "1"
        
        LettersArray = RemoveColumn(LettersArray, i) ' Remove column
                
        i = 1
        
    Else
    
    i = i + 1
    
    End If

Loop

' Order columns depending on the order of modalities: for instance, for descending CLD, if we want the highest modality
' to have the lowest letters, the table must be sorted so that the highest modalities in the table have the first columns

If SortOrder = "descending" Then

SortLettersDescending LettersArray ' Use of the "SortLetters" function to sort the columns

ElseIf SortOrder = "normal" Then

SortLettersNormal LettersArray

End If

' Creation of the concatenated letters

' Define the target modality depending on the cell selected by the user (3rd argument of the function)

ModalityPresent = False

For i = 1 To UBound(ModalityValuesArray, 1)

    If ModalityCell.Value = ModalityValuesArray(i, 1) Then
    
        TargetModality = i
        
        ModalityPresent = True ' Check if the indicated modality exists in our table
    
    End If

Next i

' If not, it means that the modality has not been tested, and then the function must return an empty cvalue

If ModalityPresent = False Then

CLD = ""

Exit Function

End If

' Final step: in the row of the target modality, if one or more columns contain "1", concatenate a letter, depending of the column number (col 1 = a, col 1+3 = ac...)

For i = 1 To UBound(LettersArray, 2)

If LettersArray(TargetModality, i) = 1 Then

ModalityLetters = ModalityLetters + (Chr(96 + i)) 'Chr(96) = letter a

End If

Next i

' End: return the concatenated letters as result of the function

CLD = ModalityLetters

'------------------------------------------------------------ Error Handling

Finalization:

On Error Resume Next:

If Not Failure Then

Else

CLD = "" ' in case of unexpected error, return empty value

End If

Exit Function

'------------------------------------------------------------

E:

Failure = True

Resume Finalization:
      
End Function

Function PopulateArray(ByVal arr As Variant, ByVal element As Variant) As Variant

' Function used to add elements to an array

Dim i As Integer
Dim ArrayModif() As Variant

' Creation of a new array from the original (defined by the 1st argument of the function), with one more empty space

ReDim ArrayModif(LBound(arr) To UBound(arr) + 1)

' Copy of the elements from the original array to the new

For i = LBound(arr) To UBound(arr)
    ArrayModif(i) = arr(i)
Next i

' Add the element (defined by the 2nd argument of the function) at the end of the array

ArrayModif(UBound(ArrayModif)) = element

' Value returned by the function is now the modified array

PopulateArray = ArrayModif
    
End Function

Function AssertionVerified(ByVal Pair As Variant, ByRef LettersArray As Variant, ByVal TargetColumn As Integer, ByVal ModalityValuesArray) As Boolean

' This function checks if two modalities both have a "1" in a same column. If so it goes against the assertion indicating
' a significant difference between both modalities. The function therefore returns "AssertionVerified = False".

Dim i As Integer
Dim Moda1Present As Boolean
Dim Moda2Present As Boolean

' Check for the presence of the 1st modality of the pair in the target column

For i = 1 To UBound(LettersArray, 1)
    If ModalityValuesArray(i, 1) = Pair(0) Then
        If LettersArray(i, TargetColumn) = 1 Then
            Moda1Present = True
        Else
            Moda1Present = False
        End If
    End If
Next i

' Check for the presence of the 2nd modality of the pair in the target column

For i = 1 To UBound(LettersArray, 1)
    If ModalityValuesArray(i, 1) = Pair(1) Then
        If LettersArray(i, TargetColumn) = 1 Then
            Moda2Present = True
        Else
            Moda2Present = False
        End If
    End If
Next i

' If both modalities are present, function returns "False", else "True"

If Moda1Present = True And Moda2Present = True Then
    AssertionVerified = False
Else
    AssertionVerified = True
End If

End Function

Sub InsertAbsorb(ByVal Pair As Variant, ByRef LettersArray As Variant, ByVal TargetColumn As Integer, ByVal ModalityValuesArray)

' Function used to duplicate a column in an array, absorb it if needed and remove the "1"s corresponding to the modalities

Dim i As Integer, j As Integer

Dim ArrayModified() As Variant

Dim Absorb As Boolean

' Copy the column

    ' Create an array with one supplementary column
    
    ReDim ArrayModified(LBound(LettersArray, 1) To UBound(LettersArray, 1), LBound(LettersArray, 2) To UBound(LettersArray, 2) + 1)
    
    ' Duplicate target column
    
    For i = LBound(LettersArray, 1) To UBound(LettersArray, 1)
        For j = LBound(LettersArray, 2) To UBound(LettersArray, 2)
            If j < TargetColumn Then
                ArrayModified(i, j) = LettersArray(i, j)
            ElseIf j = TargetColumn Then
                ArrayModified(i, j) = LettersArray(i, j)
                ArrayModified(i, j + 1) = LettersArray(i, j)
            ElseIf j > TargetColumn Then
                ArrayModified(i, j + 1) = LettersArray(i, j)
            End If
        Next j
    Next i
    
    ' Apply modifications to the original array
    
    LettersArray = ArrayModified

' In the first column delete the "1" of the first modality

For i = 1 To UBound(LettersArray, 1)
    If ModalityValuesArray(i, 1) = Pair(0) Then
        LettersArray(i, TargetColumn) = 0
    End If
Next i

' Check if the modified original column can be absorbed by a previous column

    ' First column can never be absorbed
    
    If TargetColumn <= 1 Then
    
    Absorb = False
    
    Else

        ' Run through all columns preceding target column
        
        For j = 1 To TargetColumn - 1
        
        Absorb = True
        
            ' Run through rows
            
            For i = 1 To UBound(LettersArray, 1)
            
                ' In a row, if the cell of the target column is solely filled, no absorption
    
                If LettersArray(i, TargetColumn) = 1 And LettersArray(i, j) = 0 Then
                
                    Absorb = False
                    
                End If
    
            Next i
            
        Next j
    
    End If

' If the original column can be asborded, absorb (delete) it.

If Absorb = True Then

LettersArray = RemoveColumn(LettersArray, TargetColumn) ' Use of the "removecolumn" function to delete the column without leaving empty space

TargetColumn = TargetColumn - 1 ' If the original column has been absorbed, the added column moves back for 1 step

End If

' In the added column, delete the 2nd modality

For i = 1 To UBound(LettersArray, 1)

    If ModalityValuesArray(i, 1) = Pair(1) Then
    
        LettersArray(i, TargetColumn + 1) = 0
        
    End If
    
Next i

' Check if the added modified column can be absorbed by a previous column

    ' First column can never be absorbed
    
    If TargetColumn + 1 <= 1 Then
    
    Absorb = False
    
    Else
    
    Absorb = True
    
        ' Run through all columns preceding target column
        
        For j = 1 To TargetColumn + 1 - 1
        
            ' Run through rows
            
            For i = 1 To UBound(LettersArray, 1)
            
                ' In a row, if the cell of the target column is solely filled, no absorption
    
                If LettersArray(i, TargetColumn + 1) = 1 And LettersArray(i, j) = 0 Then
                
                    Absorb = False
                    
                End If
    
            Next i
            
        Next j
    
    End If

' If the original column can be asborded, absorb (delete) it.

If Absorb = True Then

LettersArray = RemoveColumn(LettersArray, TargetColumn + 1)

End If

End Sub

Function RemoveColumn(ByRef TargetArray As Variant, ByVal TargetColumn As Long) As Variant

' Function used to remove a column from an array without leaving empty space

Dim ArrayModified() As Variant

Dim i As Long, j As Long, k As Long

' Create an array with one less column

ReDim ArrayModified(LBound(TargetArray, 1) To UBound(TargetArray, 1), LBound(TargetArray, 2) To UBound(TargetArray, 2) - 1)

' If target column is the last of the array, just copy all the original array content

If TargetColumn = UBound(TargetArray, 2) Then

    For i = LBound(TargetArray, 1) To UBound(TargetArray, 1)
        For j = LBound(TargetArray, 2) To UBound(TargetArray, 2) - 1
            ArrayModified(i, j) = TargetArray(i, j)
        Next j
    Next i
    
' Else, copy the content of the original array until the column preceding the target column, then copy the end of the original array from the target column until the end.

Else

    For i = LBound(TargetArray, 1) To UBound(TargetArray, 1)
        For j = LBound(TargetArray, 2) To TargetColumn - 1
            ArrayModified(i, j) = TargetArray(i, j)
        Next j
    Next i
    
    For i = LBound(TargetArray, 1) To UBound(TargetArray, 1)
        For j = TargetColumn + 1 To UBound(TargetArray, 2)
            ArrayModified(i, j - 1) = TargetArray(i, j)
        Next j
    Next i
    
End If

' Apply modifications

RemoveColumn = ArrayModified
    
End Function

Function RemoveRow(ByRef TargetArray As Variant, ByVal TargetRow As Long) As Variant

' Function used to remove a row from an array without leaving empty space

Dim ArrayModified() As Variant

Dim i As Long, j As Long, k As Long

' Create an array with one less row

ReDim ArrayModified(LBound(TargetArray, 1) To UBound(TargetArray, 1) - 1, LBound(TargetArray, 2) To UBound(TargetArray, 2))

' If target row is the last of the array, just copy all the original array content

If TargetRow = UBound(TargetArray, 1) Then

    For i = LBound(TargetArray, 1) To UBound(TargetArray, 1) - 1
        For j = LBound(TargetArray, 2) To UBound(TargetArray, 2)
            ArrayModified(i, j) = TargetArray(i, j)
        Next j
    Next i

' Else, copy the content of the original array until the row preceding the target row, then copy the end of the original array from the target row until the end.

Else

    For i = LBound(TargetArray, 1) To TargetRow - 1
        For j = LBound(TargetArray, 2) To UBound(TargetArray, 2)
            ArrayModified(i, j) = TargetArray(i, j)
        Next j
    Next i

    For i = TargetRow + 1 To UBound(TargetArray, 1)
        For j = LBound(TargetArray, 2) To UBound(TargetArray, 2)
            ArrayModified(i - 1, j) = TargetArray(i, j)
        Next j
    Next i

End If

' Apply modifications

RemoveRow = ArrayModified

End Function

Sub SortLettersDescending(ByRef LettersArray As Variant)

' This function sorts the letters array so that the assigned letters depend on the order of modalities

Dim LettersArrayTri As Variant

Dim i As Long, j As Long, k As Long

Dim NbrColumns As Long

Dim ColumnSum() As Double

Dim tempArray() As Variant

' The higher a modality is, the lower it is on the table. For the first letters to be associated with the highest values,
' column ranks must be created depending on the presence of "1"s and their position in the column. To do that, a table is
' created where the "1"s are replaced by "2^x", where x is the number of the row. Thus the sum of theses values in a column
' allows to sort the columns. /!\ 2^x being an exponential function, the number of treatements dramaticaly impacts calulation.

' Creating a table with the 2^x values instead of the 1s
    
LettersArrayTri = LettersArray

For i = 1 To UBound(LettersArrayTri, 1)

    For j = 1 To UBound(LettersArrayTri, 2)
    
        If LettersArrayTri(i, j) = 1 Then
        
            LettersArrayTri(i, j) = 2 ^ i
        
        End If
    
    Next j

Next i

' Sort the original table with this table

    NbrColumns = UBound(LettersArrayTri, 2)
    
    ' Calculate the sum of each column
    
    ReDim ColumnSum(1 To NbrColumns)
    
    For i = 1 To NbrColumns
    
        For j = 1 To UBound(LettersArrayTri, 1)
        
            ColumnSum(i) = ColumnSum(i) + LettersArrayTri(j, i)
            
        Next j
        
    Next i
    
    ' Reorganize the columns depending on the sums
    
    For i = 1 To NbrColumns
    
        For j = i To NbrColumns
        
            If ColumnSum(j) > ColumnSum(i) Then
            
                ' Swap sums
                
                Dim temp As Double
                
                temp = ColumnSum(i)
                
                ColumnSum(i) = ColumnSum(j)
                
                ColumnSum(j) = temp
    
                ' Swap columns in the table
                
                ReDim tempArray(1 To UBound(LettersArray, 1))
                
                For k = 1 To UBound(LettersArray, 1)
                
                    tempArray(k) = LettersArray(k, i)
                    
                    LettersArray(k, i) = LettersArray(k, j)
                    
                    LettersArray(k, j) = tempArray(k)
                    
                Next k
                
            End If
            
        Next j
        
    Next i

End Sub

Sub SortLettersNormal(ByRef LettersArray As Variant)

' This function sorts the letters array so that the assigned letters depend on the order of modalities

Dim LettersArrayTri As Variant

Dim i As Long, j As Long, k As Long

Dim NbrColumns As Long

Dim ColumnSum() As Double

Dim tempArray() As Variant

' The higher a modality is, the lower it is on the table. For the first letters to be associated with the highest values,
' column ranks must be created denpending on the presence of "1"s and their position in the column. To do that, a table is
' created where the "1"s are replaced by "2^x", where x is the number of the row. Thus the sum of theses values in a column
' allows to sort the columns. /!\ 2^x being an exponential function, the number of treatements dramaticaly impacts calulation.

' Creating a table with the 2^x values instead of the 1s
    
LettersArrayTri = LettersArray

For i = 1 To UBound(LettersArrayTri, 1)

    For j = 1 To UBound(LettersArrayTri, 2)
    
        If LettersArrayTri(i, j) = 1 Then
        
            LettersArrayTri(i, j) = 2 ^ i
        
        End If
    
    Next j

Next i

' Sort the original table with this table

    NbrColumns = UBound(LettersArrayTri, 2)
    
    ' Calculate the sum of each column
    
    ReDim ColumnSum(1 To NbrColumns)
    
    For i = 1 To NbrColumns
    
        For j = 1 To UBound(LettersArrayTri, 1)
        
            ColumnSum(i) = ColumnSum(i) + LettersArrayTri(j, i)
            
        Next j
        
    Next i
    
    ' Reorganize the columns depending on the sums
    
    For i = 1 To NbrColumns
    
        For j = i To NbrColumns
        
            If ColumnSum(j) < ColumnSum(i) Then
            
                ' Swap sums
                
                Dim temp As Double
                
                temp = ColumnSum(i)
                
                ColumnSum(i) = ColumnSum(j)
                
                ColumnSum(j) = temp
    
                ' Swap columns in the table
                
                ReDim tempArray(1 To UBound(LettersArray, 1))
                
                For k = 1 To UBound(LettersArray, 1)
                
                    tempArray(k) = LettersArray(k, i)
                    
                    LettersArray(k, i) = LettersArray(k, j)
                    
                    LettersArray(k, j) = tempArray(k)
                    
                Next k
                
            End If
            
        Next j
        
    Next i

End Sub
