Attribute VB_Name = "Regex"
' Regex
' =====
' Implements regular expressions, mimicking features of the Python `re` module:
'  https://docs.python.org/3/library/re.html
'
' For:
'  - Microsoft Excel

Option Explicit

Function RegexSearch(search_pattern As String, search_string As String) As String
    ' Return the first match of `search_pattern` applied to `search_string`.
    '
    ' Examples
    ' --------
    ' >>> RegexSearch("c[ao]t", "cat")
    ' "cat"
    '
    ' >>> RegexSearch("c[ao]t", "cut")
    ' ""
    '
    ' >>> RegexSearch("c[ao]t", "catcot")
    ' "cat"
    '
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    With re
        .pattern = search_pattern
        .IgnoreCase = False
        .MultiLine = True
        .Global = False
    End With

    Dim result As String

    If re.Test(search_string) Then
        Dim matches As Object
        Set matches = re.Execute(search_string)

        Dim match As Object
        For Each match In matches
            result = match.Value
        Next
    Else
        result = ""
    End If

    RegexSearch = result
End Function
