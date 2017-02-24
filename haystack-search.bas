Attribute VB_Name = "HaystackSearch"
Option Explicit

'
' Provides functionality for searching a spreadsheet to find matching values
'


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
    End If
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

    FindMatchingRow = 0
    
    haystack_column = DetermineColumnNumber(haystack_column)
    With haystack
        blank_count = 0
        current_row = start_row
        Do While blank_count < CONCURRENT_BLANK_VALUES And current_row <= haystack.rows.Count
            If Trim(.Cells(current_row, haystack_column)) <> "" Then
                If blanks_allowed Then
                    blank_count = 0
                End If

                If CStr(Trim(.Cells(current_row, haystack_column))) = CStr(needle) Then
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
End Function


Public Function FindMatchingRows(ByRef haystack As Worksheet, ByVal haystack_column As Variant, ByVal needle As Variant, Optional ByVal start_row As Long = 2, Optional ByVal blanks_allowed As Boolean = False) As Collection
    'Returns a collection of rows in worksheet `haystack` with a value in
    '  `haystack_column` that matches `needle`.
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
    '@return  Collection  a collection of row numbers in `haystack` that
    '                     contains the matching value to `needle`. Returns an
    '                     empty collection if no matches are found.

    Dim rows As Collection
    Dim current_row As Long
    Dim blank_count As Integer

    Set rows = New Collection
    
    haystack_column = DetermineColumnNumber(haystack_column)
    With haystack
        blank_count = 0
        current_row = start_row
        Do While blank_count < CONCURRENT_BLANK_VALUES And current_row <= haystack.rows.Count
            If Trim(.Cells(current_row, haystack_column)) <> "" Then
                If blanks_allowed Then
                    blank_count = 0
                End If

                If CStr(Trim(.Cells(current_row, haystack_column))) = CStr(needle) Then
                    rows.Add current_row
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
    
    Set FindMatchingRows = rows
End Function
