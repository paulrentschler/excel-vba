Attribute VB_Name = "SpecialSauce"
Option Explicit

'
' Provides basic functionality that really should be part of VBA but isn't :(
'

Public Function Implode(ByVal entries As Collection, Optional ByVal delimiter As String = ", ") As String
    'Combines the values in a collection into a single string
    '
    'Individual values are separated by `delimiter`
    '
    '@param   entries    a Collection of values that will be combined into a
    '                    single string
    '@param   delimiter  a string indicating what character(s) should be used
    '                    to separate the values in `entries`
    '@return  String  a string composed of all the items in the collection
    '                 separated by `delimiter`

    Dim items() As String
    Dim entry As Variant
    Dim index As Long

    If entries.Count > 0 Then
        'convert the collection to an array so Join works
        ReDim items(entries.Count - 1)

        index = 0
        For Each entry In entries
            items(index) = entry
            index = index + 1
        Next

        Implode = Join(items, delimiter)
    Else
        Implode = ""
    End If
End Function

Sub ToggleR1C1()
'
' ToggleR1C1 Macro
' Toggle between R1C1 and A1 notation
'
' Keyboard Shortcut: Ctrl+Shift+R
'
' Courtesy of: https://gist.github.com/jakelosh/5851415
'
    If Application.ReferenceStyle = xlR1C1 Then
        Application.ReferenceStyle = xlA1
    Else
        Application.ReferenceStyle = xlR1C1
    End If
End Sub
