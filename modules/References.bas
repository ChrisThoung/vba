Attribute VB_Name = "References"
' References
' ==========
' Functions to operate on cell references.
'
' For:
'  - Microsoft Excel

Option Explicit

Function AlphaToInt(ColRef As String) As Integer
    ' Return an alphabetic column reference (e.g. "AA") as an integer column
    ' number (27 in the previous example).
    '
    ' Examples
    ' --------
    ' >>> AlphaToInt("A")
    ' 1
    '
    ' >>> AlphaToInt("AA")
    ' 27
    '
    ' >>> AlphaToInt("AZ")
    ' 52
    '
    ' See also
    ' --------
    ' IntToAlpha() : The reverse of this function
    '
    ColRef = UCase(ColRef)

    Dim num_chars As Integer
    num_chars = Len(ColRef)

    Dim result As Integer

    If num_chars = 1 Then
        ' Simplest case: One character to convert to integer
        result = InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", ColRef)
    Else
        ' Otherwise, sum across character string (in base-26)
        result = 0

        Dim i As Integer
        Dim multiplier As Integer
        For i = 1 To num_chars
            multiplier = 26 ^ (num_chars - i)
            result = result + multiplier * AlphaToInt(Mid(ColRef, i, 1))
        Next
    End If

    AlphaToInt = result
End Function

Function IntToAlpha(ColNumber As Integer) As String
    ' Return an integer column number (e.g. 27) as an alphabetic column
    ' reference ("AA" in the previous example).
    '
    ' Examples
    ' --------
    ' >>> IntToAlpha(1)
    ' "A"
    '
    ' >>> IntToAlpha(27)
    ' "AA"
    '
    ' >>> IntToAlpha(52)
    ' "AZ"
    '
    ' See also
    ' --------
    ' AlphaToInt() : The reverse of this function
    '
    Dim result As String

    If ColNumber < 1 Then
        result = CVErr(xlErrValue)
    ElseIf ColNumber < 27 Then
        ' Simplest case: Convert single character to integer
        result = Chr(64 + ColNumber)
    Else
        ' Otherwise, assemble string recursively

        ' Switch to zero-based numbering to avoid borrow operation later on
        ColNumber = ColNumber - 1

        Dim Quotient As Integer
        Dim Remainder As Integer
        Quotient = Int(ColNumber / 26)
        Remainder = ColNumber - (Quotient * 26)

        ' Add 65, rather than 64, because of switch to zero base
        result = IntToAlpha(Quotient) & Chr(65 + Remainder)
    End If

    IntToAlpha = result
End Function
