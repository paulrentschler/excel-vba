Attribute VB_Name = "Module1"
Option Explicit

'How many blank values in a row causes the search to stop (must be greater than zero)
Const CONCURRENT_BLANK_VALUES = 100

'12345678901234567890123456789012345678901234567890123456789012345678901234567890

Public Function ConvertCharColumnToNumber(ByVal column_letter As String) As Long
    'When provided a column letter it returns the column's number
    '
    '@param   column_letter  the letter designation for a column
    '@return  Long  the numeric representation for the column
    
    ConvertCharColumnToNumber = ActiveSheet.Range(column_letter & "1").column
End Function


Public Function DetermineColumnNumber(ByVal column As Variant) As Long
    'Returns the column number referenced by `column`
    '
    '@param   column  a letter or number representation of a column
    '@return  Long  the numeric representation for `Column`
    
    If IsNumeric(column) Then
        DetermineColumnNumber = column
    Else
        DetermineColumnNumber = ConvertCharColumnToNumber(column)
End Function


Public Function FindMatchingRow(ByRef haystack As Worksheet, ByVal haystack_column As Variant, ByVal needle As Variant, Optional ByVal start_row As Long = 2, Optional ByVal blanks_allowed As Boolean = False) As Long
    'Returns the row in worksheet `haystack` with a value in `haystack_column`
    '  that matches `needle`.
    '
    '@param   haystack         the worksheet object that should be searched
    '@param   haystack_column  the column (number or letter) that should be
    '                          searched
    '@param   needle           the value to find in `haystack_column`
    '@param   start_row        the row that the search should start with.
    '                          Defaults to 2 (assumes a headings row).
    '@param   blanks_allowed   a boolean indicating whether blank values in
    '                          `haystack_column` should stop the search for
    '                          `needle`. Default is False (a blank value will
    '                          stop the search.
    '@return  Long  the number of the row in `haystack` that contains the
    '               matching value to `needle`. Returns `0` if not found.
    
    Dim current_row As Long
    Dim blank_count As Integer
    
    haystack_column = DetermineColumnNumber(haystack_column)
    With haystack
        blank_count = 0
        current_row = start_row
        Do While blank_count < CONCURRENT_BLANK_VALUES And current_row <= haystack.Rows.Count
            If Trim(.Cells(current_row, haystack_column)) <> "" Then
                If blanks_allowed Then
                    blank_count = 0
                End If
                
                If Trim(.Cells(current_row, haystack_column)) = needle Then
                    FindMatchingRow = current_row
                    Exit Do
                End If
            Else
                If blanks_allowed Then
                    blank_count = blank_count + 1
                Else
                    blank_count = CONCURRENT_BLANK_VALUES + 1
                End If
            End If
                
            current_row = current_row + 1
            DoEvents
        Loop
    End With
    
    FindMatchingRow = 0
End Function
