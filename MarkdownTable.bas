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
    
    For mCol = startCol To endCol
    
        For mRow = startRow To endRow
    
            currentLenght = Len(tableArray(mRow, mCol))
            MaxLenght = outputLength(mCol)
            
            If currentLenght > MaxLenght Then
                outputLength(mCol) = currentLenght
            End If
            
            Next mRow
    
        Next mCol
    
    For mCol = startCol To endCol
        Dim currentColumn As Integer
        
        Dim colLength As Integer
        colLength = outputLength(mCol)
        
        currentColumn = Selection.Column + mCol - 1
        
        If mCol = startCol Then
            rowSpaceHeader = "|" + String(colLength + 2, "-") + "|"
        Else
            rowSpaceHeader = String(colLength + 2, "-") + "|"
        End If
        
        For mRow = startRow To endRow
            entry = tableArray(mRow, mCol)
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

