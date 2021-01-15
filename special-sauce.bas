Attribute VB_Name = "SpecialSauce"
Option Explicit

'
' Provides basic functionality that really should be part of VBA but isn't :(
'

Public Sub ArrayRemoveItem(ItemArray As Variant, ByVal ItemElement As Long)
'
'PURPOSE:       Remove an item from an array, then
'               resize the array
'
'PARAMETERS:    ItemArray: Array, passed by reference, with
'               item to be removed.  Array must not be fixed
'
'               ItemElement: Element to Remove
'
'EXAMPLE:
'           dim iCtr as integer
'           Dim sTest() As String
'           ReDim sTest(2) As String
'           sTest(0) = "Hello"
'           sTest(1) = "World"
'           sTest(2) = "!"
'           ArrayRemoveItem sTest, 1
'           for iCtr = 0 to ubound(sTest)
'               Debug.print sTest(ictr)
'           next
'
'           Prints
'
'           "Hello"
'           "!"
'           To the Debug Window
'
' Courtesy of: https://www.freevbcode.com/ShowCode.asp?ID=585
'
Dim lCtr As Long
Dim lTop As Long
Dim lBottom As Long

If Not IsArray(ItemArray) Then
    Err.Raise 13, , "Type Mismatch"
    Exit Sub
End If

lTop = UBound(ItemArray)
lBottom = LBound(ItemArray)

If ItemElement < lBottom Or ItemElement > lTop Then
    Err.Raise 9, , "Subscript out of Range"
    Exit Sub
End If

For lCtr = ItemElement To lTop - 1
    ItemArray(lCtr) = ItemArray(lCtr + 1)
Next
On Error GoTo ErrorHandler:

ReDim Preserve ItemArray(lBottom To lTop - 1)

Exit Sub
ErrorHandler:
  'An error will occur if array is fixed
    Err.Raise Err.Number, , _
       "You must pass a resizable array to this function"
End Sub

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
