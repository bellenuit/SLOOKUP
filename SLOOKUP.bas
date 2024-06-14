Attribute VB_Name = "Modul3"
' SLOOKUP and LLOOKUP for Excel
' Functions to lookup for similar values
'
' LLOOKUP(needle, haystack, result, optional threshold = 0.75, optional partialstring = true, optional simplestring = true)
' Finds the best match for needle in the haystack range and returns the corresponding item in the result range using the Levensthein similarity
' Needle is interpreted as string.
' Haystack and result must be ranges with a single column and equal number of rows.
' Haystack can be of any order, but the search stops at an empty cell, if the range is the entire column (eg. A:A)
' Threshold fine tunes the sensitivity (for values between 0 and 1). If the score of no value in haystack exceeds threshold, the result is empty.
' Partialstring looks for asymetric matches where needle is included in haystack but the haystack not in needle
' Simplestring uses URL-representation of the needle and the haystack (all lowercase, no special characters)
'
' SLOOKUP(needle, haystack, result, optional threshold = 0.75, optional partialstring = true, optional simplestring = true)
' Finds the best match for needle in the haystack range and returns the corresponding item in the result range using the simpleSim similarity
' Needle is interpreted as string.
' Haystack and result must be ranges with a single column and equal number of rows.
' Haystack can be of any order, but the search stops at an empty cell, if the range is the entire column (eg. A:A)
' Threshold fine tunes the sensitivity (for values between 0 and 1). If the score of no value in haystack exceeds threshold, the result is empty.
' Partialstring looks for asymetric matches where needle is included in haystack but the haystack not in needle
' Simplestring uses URL-representation of the needle and the haystack (all lowercase, no special characters)
'
' LevenshteinDistance(needle, haystack, optional partialstring = true, optional simplestring = true)
' Calculates the edit distance of needle and haystack.
' Partialstring looks for asymetric matches where needle is included in haystack but the haystack not in needle
' Simplestring uses URL-representation of the needle and the haystack (all lowercase, no special characters)
'
' LevenshteinSimilarity(needle, haystack, optional partialstring = true, optional simplestring = true)
' Calculates the edit distance of needle and haystack as a similarity measure (0 not similar, 1 identical)
' Partialstring looks for asymetric matches where needle is included in haystack but the haystack not in needle
' Simplestring uses URL-representation of the needle and the haystack (all lowercase, no special characters)
'
' SimpleSimilarity(needle, haystack, result, optional partialstring = true, optional simplestring = true)
' Calculates the similarity of needle with haystack between 0 (not at all similar) and 1 (identical) based on custom algorithm
' Partialstring looks for asymetric matches where needle is included in haystack but the haystack not in needle
' Simplestring uses URL-representation of the needle and the haystack (all lowercase, no special characters)
'
' SimpleURL(text)
' Converts a text to an URL-compatible form (all lowercase and all non-alphanumeric characters replaced by "-")
'
' The functions do not have side effects
'
' Freeware
' Version 1.0 2024-06-14
' Author: matti@belle-nuit.com 2024

Option Explicit

Function LLOOKUP(needle As String, haystack As Range, result As Range, Optional threshold As Double = 0.75, Optional partialstring As Boolean = True, Optional simplestring As Boolean = True) As Variant

' Finds the best match for needle in the haystack range and returns the corresponding item in the result range using the Levensthein similarity
' Needle is interpreted as string.
' Haystack and result must be ranges with a single column and equal number of rows.
' Haystack can be of any order, but the search stops at an empty cell, if the range is the entire column (eg. A:A)
' Threshold fine tunes the sensitivity (for values between 0 and 1). If the score of no value in haystack exceeds threshold, the result is empty.
' Partialstring looks for asymetric matches where needle is included in haystack but the haystack not in needle
' Simplestring uses URL-representation of the needle and the haystack (all lowercase, no special characters)
'
' The algorithm calculates the LevenshteinSimilarity() score between needle and each haystack item and chooses the best, if threshold is met
'
' Complexity (for needle length n and haystack items m of length h)
' The comlexity is O(n*m*h)

Dim i, r, found As Long
Dim larr() As Variant
Dim rarr() As Variant
Dim score, newscore As Double
Dim h As String

larr = haystack
rarr = result

If UBound(larr, 2) <> 1 Then
    LLOOKUP = "#ERROR invalid lookup_array column count"
    Return
End If

If UBound(rarr, 2) <> 1 Then
    LLOOKUP = "#ERROR invalid return_array column count"
    Return
End If

If UBound(larr, 1) <> UBound(rarr, 1) Then
    LLOOKUP = "#ERROR row count if input and return array do not match"
    Return
End If

r = UBound(larr, 1)

score = 0#
found = 0

For i = 1 To r
    h = larr(i, 1)
    If h <> "" Then
        newscore = LevenshteinSimilarity(needle, h, partialstring, simplestring)
        If newscore > score Then
            found = i
            score = newscore
        End If
        ' stop when it cannot get better
        If score = 1 Then Exit For
    Else
        'stop on empty value on ranges that are the entire column (currently 1024*1024)
        If r > 1000000 Then Exit For
    End If
Next

If score >= threshold Then
    LLOOKUP = rarr(found, 1)
Else
    LLOOKUP = "" ' not found
End If


End Function




Function SLOOKUP(needle As String, haystack As Range, result As Range, Optional threshold As Double = 0.75, Optional partialstring As Boolean = True, Optional simplestring As Boolean = True) As Variant

' Finds the best match for needle in the haystack range and returns the corresponding item in the result range using the simpleSim similarity
' Needle is interpreted as string.
' Haystack and result must be ranges with a single column and equal number of rows.
' Haystack can be of any order, but the search stops at an empty cell, if the range is the entire column (eg. A:A)
' Threshold fine tunes the sensitivity (for values between 0 and 1). If the score of no value in haystack exceeds threshold, the result is empty.
' Partialstring looks for asymetric matches where needle is included in haystack but the haystack not in needle
' Simplestring uses URL-representation of the needle and the haystack (all lowercase, no special characters)
'
' The algorithm calculates the SimpleSim() score between needle and each haystack item and chooses the best, if threshold is met
'
' Complexity (for needle length n and haystack items m of length h)
' For exact matches (best case), the comlexity is O(n*m)
' For no matches (worst case), the complexity is O(n*m*h)
' For random needles and haystacks of sufficient length, the complexity is O(n*m*18) = O(n*m)

Dim i, r, found As Long
Dim larr() As Variant
Dim rarr() As Variant
Dim score, newscore As Double
Dim h As String

larr = haystack
rarr = result

If UBound(larr, 2) <> 1 Then
    SLOOKUP = "#ERROR invalid lookup_array column count"
    Return
End If

If UBound(rarr, 2) <> 1 Then
    SLOOKUP = "#ERROR invalid return_array column count"
    Return
End If

If UBound(larr, 1) <> UBound(rarr, 1) Then
    SLOOKUP = "#ERROR row count if input and return array do not match"
    Return
End If

r = UBound(larr, 1)

score = 0#
found = 0

For i = 1 To r
    h = larr(i, 1)
    If h <> "" Then
        newscore = SimpleSim(needle, h, simplestring)
        If newscore > score Then
            found = i
            score = newscore
        End If
        ' stop when it cannot get better
        If score = 1 Then Exit For
    Else
        'stop on empty value on ranges that are the entire column (currently 1024*1024)
        If r > 1000000 Then Exit For
    End If
Next

If score >= threshold Then
    SLOOKUP = rarr(found, 1)
Else
    SLOOKUP = "" ' not found
End If


End Function

Function SimpleSim(needle As String, haystack As String, Optional partialstring As Boolean = True, Optional simplestring As Double = True) As Double

' Calculates the similarity of needle with haystack between 0 (not at all similar) and 1 (identical) based on custom algorithm
' Partialstring looks for asymetric matches where needle is included in haystack but the haystack not in needle
' Simplestring uses URL-representation of the needle and the haystack (all lowercase, no special characters)
'
' Complexity (for needle length n and haystack length h)
' Best case, needle is start of haystack O(n)
' Worst case no character of needle is present in haystack O(n*m)
' Random needles and haystacks O(n*36) = O(n)

Dim score As Double
Dim Offset, HaystackLength, NeedleLength, i, j, test As Integer
Dim n, h As String

If simplestring Then
    haystack = SimpleURL(haystack)
    needle = SimpleURL(needle)
End If

score = 0
Offset = 0
HaystackLength = Len(haystack)
NeedleLength = Len(needle)

If NeedleLength = 0 Or HaystackLength = 0 Then
    Exit Function
End If
    
' loop through needle character
For i = 1 To NeedleLength
n = Mid(needle, i, 1)
    ' loop through haystack character at the last offset of last match
    For j = Offset + 1 To Offset + HaystackLength
        test = j
        ' wrap over length, restart at beginning
        If test > HaystackLength Then
            test = test - HaystackLength
        End If
        h = Mid(haystack, test, 1)
        If n = h Then
            ' partial score per character is maximal 1 if the match is immediately
            ' the score per character diminishes if the algorithm has to look longer
            score = score + 1 / (j - Offset)
            Offset = test
            Exit For
        End If
    Next j
Next i

SimpleSim = score / NeedleLength

' calculate average of sim in both direction
If Not partialstring Then
    SimpleSim = (SimpleSim + SimpleSim(needle, haystack)) / 2
End If

End Function


Function SimpleURL(text As String) As String

' Converts a text to an URL-compatible form (all lowercase and all non-alphanumeric characters replaced by "-")
'
' Complexity (for text length n)
' The complexity is O(n)

Dim Bytes() As Byte
Dim i As Integer

Bytes = LCase(text)

' string is UTF16, so we need to read 2 bytes at a time
For i = 0 To UBound(Bytes) Step 2
    Select Case Bytes(i)
    Case &H0, &H30 To &H39, &H61 To &H7A
        ' pass &H0 = null, &H30 = 0, &H39 = 9, &H61  = a, &H7A = z
    Case Else
        ' &H2D = -
        Bytes(i) = &H2D
        Bytes(i + 1) = 0
End Select
Next

SimpleURL = Bytes

End Function


Function LevenshteinDistance(s As String, t As String, Optional partialstring As Boolean = True, Optional simplestring As Boolean = True) As Long

' Calculates the edit distance of needle and haystack.
' Partialstring looks for asymetric matches where needle is included in haystack but the haystack not in needle
' Simplestring uses URL-representation of the needle and the haystack (all lowercase, no special characters)
'
' Wagner-Fischer algorithm of Levenshtein distance
'
' Complexity (for text length n)
' The complexity is O(n)

Dim d() As Long
Dim i, j, m, n As Long
Dim cost As Long

If s = "" Then
    LevenshteinDistance = Len(t)
    Exit Function
End If
If t = "" Then
    LevenshteinDistance = Len(s)
    Exit Function
End If

If simplestring Then
    s = SimpleURL(s)
    t = SimpleURL(t)
End If

m = Len(s)
n = Len(t)

ReDim d(m, n)

For i = 1 To m
    d(i, 0) = i
Next

cost = 0


For j = 1 To n
       If partialstring Then
             d(0, j) = 0
       Else
            d(0, j) = j
      End If
Next

For j = 1 To n
    For i = 1 To m
        If StrComp(Mid(s, i, 1), Mid(t, j, 1), 0) = 0 Then
            cost = 0
        Else
            cost = 1
        End If
        d(i, j) = minimal3(d(i - 1, j) + 1, d(i, j - 1) + 1, d(i - 1, j - 1) + cost)
    Next
Next


If partialstring Then
    LevenshteinDistance = d(m, 0)
    For j = 1 To n
        If LevenshteinDistance > d(m, j) Then
            LevenshteinDistance = d(m, j)
        End If
    Next
Else
    LevenshteinDistance = d(m, n)
End If

End Function


Private Function minimal3(x, y, z As Variant) As Variant
   minimal3 = IIf(x < y, x, y)
   minimal3 = IIf(minimal3 < z, minimal3, z)
End Function


Function LevenshteinSimilarity(s As String, t As String, Optional partialstring As Boolean = True, Optional simplestring As Boolean = True) As Double

' Calculates the edit distance of needle and haystack as a similarity measure (0 not similar, 1 identical)
' Partialstring looks for asymetric matches where needle is included in haystack but the haystack not in needle
' Simplestring uses URL-representation of the needle and the haystack (all lowercase, no special characters)

If s = "" Or t = "" Then
    LevenshteinSimilarity = 0#
    Exit Function
End If

If partialstring Then
    LevenshteinSimilarity = 1 - LevenshteinDistance(s, t, True, simplestring) / Len(s)
' the bigger of both string lengths is the maximal possible distance
ElseIf Len(s) > Len(t) Then
    LevenshteinSimilarity = 1 - LevenshteinDistance(s, t, False, simplestring) / Len(s)
Else
    LevenshteinSimilarity = 1 - LevenshteinDistance(s, t, False, simplestring) / Len(t)
End If

End Function





