Attribute VB_Name = "StringOperations"
Option Explicit
Option Compare Text

Function AddRowToBOM(name As String, idFormatTitles As profileFormatTitles, _
    appendEmptyLines As Boolean, idFormatStrings As profileFormatStrings, _
    isTitle As Boolean) As Boolean

    Const rowIndex As Integer = 1
    Dim colIndex As Integer
    Const missingColIndex As Integer = -1
    Dim format As TextFormat
    Dim i As Integer
        
    colIndex = missingColIndex
    For i = 0 To swTable.TotalColumnCount - 1
        If swTable.text(0, i) Like "*наименование*" Then
            colIndex = i
            Exit For
        End If
    Next
    If colIndex = missingColIndex Then
        colIndex = IIf(swTable.TotalColumnCount < 7, 3, 4)
    End If
    
    AddRowToBOM = swTable.InsertRow(swTableItemInsertPosition_Before, rowIndex)
    Set format = swTable.GetCellTextFormat(rowIndex, colIndex)
    If AddRowToBOM Then
        If isTitle Then
            Select Case idFormatTitles
                Case idUnderline
                    format.Underline = True
                    swTable.SetCellTextFormat rowIndex, colIndex, False, format
                Case idUpper
                    name = UCase(name)
                    swTable.SetCellTextFormat rowIndex, colIndex, True, format
            End Select
            swTable.CellTextHorizontalJustification(rowIndex, colIndex) = swTextJustificationCenter
            swTable.text(rowIndex, colIndex) = name
            
            If appendEmptyLines Then
                'порядок важен!
                swTable.InsertRow swTableItemInsertPosition_After, rowIndex
                swTable.InsertRow swTableItemInsertPosition_Before, rowIndex
            End If
        Else
            swTable.SetCellTextFormat rowIndex, colIndex, True, format
            Select Case idFormatStrings
                Case idLeft
                    swTable.CellTextHorizontalJustification(rowIndex, colIndex) = swTextJustificationLeft
                Case idCenter
                    swTable.CellTextHorizontalJustification(rowIndex, colIndex) = swTextJustificationCenter
            End Select
            swTable.text(rowIndex, colIndex) = name
        End If
    End If
End Function
