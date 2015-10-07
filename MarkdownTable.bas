Attribute VB_Name = "MarkdownTableModule"
Sub MarkdownTable()
' Sylvain for markdown (fork from Jens Bodal 2/13/2015)
' This macro requires the Microsoft Forms 2.0 Object Library, it will likely
' need to be added manually from Tools => References => Browse => FM20.DLL

    ' Data Object for copying Reddit Table Code to clipboard
    Dim dataObject As New MSForms.dataObject
    ' Stores current selection as a Variant Array
    Dim tableArray As Variant
    Dim startCol As Integer
    Dim endCol As Integer
    Dim startRow As Integer
    Dim endRow As Integer
    ' Stores the value of each cell as we iterate through tableArray
    Dim entry As String
    ' This array will be converted to final outputString
    Dim outputArray() As String
    ' This will be the output string that is copied to the clipboard
    Dim outputString As String
    ' Determines alignment of the column based on alignment of row 1 in column
    Dim colAlignment As String
    
    Dim rowSpaceHeader As String
    
    If Selection.Count <= 1 Then
        MsgBox ("Selection is not an Array")
        End
    End If
        
    'Create array from selection, assume 2d array (row, column)
    tableArray = Selection.Value
    startCol = LBound(tableArray, 2)
    endCol = UBound(tableArray, 2)
    startRow = LBound(tableArray, 1)
    endRow = UBound(tableArray, 1)
    
    ReDim outputArray(0 To endRow) As String
    
    ' Store max text length for each column
    Dim outputLength() As Integer
    ReDim outputLength(0 To endCol) As Integer
    
    Dim currentColumn As Integer
    Dim currentRow As Integer
            
    
    For mCol = startCol To endCol
    
        currentColumn = Selection.Column + mCol - 1
    
        For mRow = startRow To endRow
    
            currentRow = Selection.Row + mRow - 1
    
            currentLenght = Len(tableArray(mRow, mCol))
            
            isBold = Cells(currentRow, currentColumn).Font.Bold
            isItalic = Cells(currentRow, currentColumn).Font.Italic
            isUnderline = (Cells(currentRow, currentColumn).Font.Underline = 2)
            
            If isBold Then
                currentLenght = currentLenght + 4
            ElseIf isItalic Then
                currentLenght = currentLenght + 2
            ElseIf isUnderline Then
                currentLenght = currentLenght + 2
            End If

            MaxLenght = outputLength(mCol)
            
            If currentLenght > MaxLenght Then
                outputLength(mCol) = currentLenght
            End If
            
            Next mRow
    
            If outputLength(mCol) = 0 Then
                outputLength(mCol) = 10
            End If
    
        Next mCol
    
    Dim SeparatorStart As String
    Dim SeparatorEnd As String
    
    For mCol = startCol To endCol
        
        Dim colLength As Integer
        colLength = outputLength(mCol)
        
        currentColumn = Selection.Column + mCol - 1
        cellAlignment = Range(Cells(Selection.Row, currentColumn), Cells(Selection.Row, currentColumn)).HorizontalAlignment
        
        SeparatorStart = IIf(mCol = startCol, "|", "") & IIf(cellAlignment = xlCenter, ":", " ")
        SeparatorEnd = IIf(cellAlignment = xlCenter Or cellAlignment = xlRight, ":", " ") & "|"
    
        rowSpaceHeader = SeparatorStart + String(colLength, "-") + SeparatorEnd
   
        For mRow = startRow To endRow
            entry = tableArray(mRow, mCol)
            
            If Not mRow = startRow Then
           
                currentRow = Selection.Row + mRow - 1
                isBold = Cells(currentRow, currentColumn).Font.Bold
                isItalic = Cells(currentRow, currentColumn).Font.Italic
                isUnderline = (Cells(currentRow, currentColumn).Font.Underline = 2)
             
                If isBold Then
                    entry = "**" & entry & "**"
                ElseIf isItalic Then
                    entry = "*" & entry & "*"
                ElseIf isUnderline Then
                    entry = "`" & entry & "`"
                End If
            End If
            
            mIndex = mRow
            ' First row has index of 0.  As 2nd row in Table formatting
            ' defines column alignment the rest of the indices are equal to
            ' the actual row number
            If mRow = startRow Then
                mIndex = mRow - 1
            End If
            ' Adding new column notation to end of entry
            
            If mCol = startCol Then
                outputArray(mIndex) = "| "
            End If
            
            outputArray(mIndex) = outputArray(mIndex) + Left(entry & Space(colLength), colLength) + " | "
            If mCol = endCol Then
                outputArray(mIndex) = outputArray(mIndex) + vbCrLf
            End If
    
            Next mRow
            ' For each column need to assign formatting in 2nd table row
            outputArray(1) = outputArray(1) + rowSpaceHeader
            
        Next mCol
        
    ' Add line break at end of 2nd table row
    outputArray(1) = outputArray(1) + vbCrLf
    
    For Each Item In outputArray
        outputString = outputString + Item
        Next Item
    
    MsgBox ("COPIED TO CLIPBOARD" + vbCrLf + outputString)
    dataObject.SetText outputString
    dataObject.PutInClipboard
    
End Sub

