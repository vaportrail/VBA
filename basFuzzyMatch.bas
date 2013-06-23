Attribute VB_Name = "basSoundex"
Option Compare Database
Option Explicit

Public Function Soundex(varText As Variant) As Variant
On Error GoTo Error_Handler
    'Purpose:   Return Soundex value for the text passed in.
    'Return:    Soundex code, or Null for Error, Null or zero-length string.
    'Argument:  The value to generate the Soundex for.
    'Author:    Allen Browne (allen@allenbrowne.com), November 2007.
    'Algorithm: Based on http://en.wikipedia.org/wiki/Soundex
    Dim strSource As String     'varText as a string.
    Dim strOut As String        'Output string to build up.
    Dim strValue As String      'Value for current character.
    Dim strPriorValue As String 'Value for previous character.
    Dim lngPos As Long          'Position in source string
    
    'Do not process Error, Null, or zero-length strings.
    If Not IsError(varText) Then
        strSource = Trim$(Nz(varText, vbNullString))
        If strSource <> vbNullString Then
            'Retain the initial character, and process from 2nd.
            strOut = left$(strSource, 1&)
            strPriorValue = SoundexValue(strOut)
            lngPos = 2&
            
            'Examine a character at a time, until we output 4 characters.
            Do
                strValue = SoundexValue(Mid$(strSource, lngPos, 1&))
                'Omit repeating values (except the zero for padding.)
                If ((strValue <> strPriorValue) And (strValue <> vbNullString)) Or (strValue = "0") Then
                    strOut = strOut & strValue
                    strPriorValue = strValue
                End If
                lngPos = lngPos + 1&
            Loop Until Len(strOut) >= 4&
        End If
    End If
    
    'Return the output string, or Null if nothing generated.
    If strOut <> vbNullString Then
        Soundex = strOut
    Else
        Soundex = Null
    End If
    
Exit_Handler:
    Exit Function
    
Error_Handler:
    MsgBox "Error " & Err.number & ": " & Err.Description, vbExclamation, "Soundex()"
    'Call LogError(Err.Number, Err.Description, conMod & ".Soundex")
    Resume Exit_Handler
End Function
Private Function SoundexValue(strChar As String) As String
    Select Case strChar
        Case "B", "F", "P", "V"
            SoundexValue = "1"
        Case "C", "G", "J", "K", "Q", "S", "X", "Z"
            SoundexValue = "2"
        Case "D", "T"
            SoundexValue = "3"
        Case "L"
            SoundexValue = "4"
        Case "M", "N"
            SoundexValue = "5"
        Case "R"
            SoundexValue = "6"
        Case vbNullString
            'Pad trailing zeros if no more characters.
            SoundexValue = "0"
        Case Else
        'Return nothing for "A", "E", "H", "I", "O", "U", "W", "Y", non-alpha.
    End Select
End Function

Public Function Levenshtein(ByVal s As String, ByVal t As String) As Integer

    Dim d() As Integer  ' matrix
    Dim m As Integer    ' length of t
    Dim n As Integer    ' length of s
    Dim i As Integer    ' iterates through s
    Dim j As Integer    ' iterates through t
    Dim s_i As String   ' ith character of s
    Dim t_j As String   ' jth character of t
    Dim cost As Integer ' cost
  
    ' Step 1
    n = Len(s)
    m = Len(t)
    
    If n = 0 Then
        Levenshtein = m
        Exit Function
    End If
    
    If m = 0 Then
        Levenshtein = n
        Exit Function
    End If
    
    ReDim d(0 To n, 0 To m) As Integer
  
    ' Step 2
    For i = 0 To n
        d(i, 0) = i
    Next i
  
    For j = 0 To m
        d(0, j) = j
    Next j

    ' Step 3
    For i = 1 To n
        s_i = Mid$(s, i, 1)
        ' Step 4
        For j = 1 To m
            t_j = Mid$(t, j, 1)
            ' Step 5
            If s_i = t_j Then
                cost = 0
            Else
                cost = 1
            End If
            ' Step 6
            d(i, j) = Minimum(d(i - 1, j) + 1, d(i, j - 1) + 1, d(i - 1, j - 1) + cost)
        Next j
    Next i
  
    ' Step 7
    Levenshtein = d(n, m)
    Erase d

End Function

Private Function Minimum(ByVal a As Integer, _
                         ByVal b As Integer, _
                         ByVal c As Integer) As Integer
    Dim mi As Integer

    mi = a
    If b < mi Then
        mi = b
    End If

    If c < mi Then
        mi = c
    End If
  
    Minimum = mi

End Function

Public Function Simil(strTxt1 As String, strTxt2 As String) As Double
'Determine match percentage between two strings. between 0 (no match) en 1 (identical)

    Dim intTot   As Integer
    Dim strMatch As String
    
    intTot = Len(strTxt1 & strTxt2) 'len(strtxt1) + len(strtxt2) 'Which is faster?
    
    strMatch = GetBiggest(strTxt1, strTxt2)
    
    Simil = CDbl(Len(strMatch) * 2) / CDbl(intTot)

End Function

Private Function GetBiggest(strTxt1 As String, strTxt2 As String) As String
'Returnvalue is all matching strings
'?GetBiggest("Pennsylvania","Pencilvaneya")
'lvanPena

    Dim intLang    As Integer
    Dim intKort    As Integer
    Dim intPos     As Integer
    Dim intX       As Integer
    Dim strLangste As String
    Dim strSearch  As String
    Dim strLang    As String
    Dim strKort    As String
    Dim strTotal1 As String
    Dim strTotal2 As String
    
    intKort = Len(strTxt1)
    intLang = Len(strTxt2)
    
    If intLang > intKort Then
        strLang = strTxt2
        strKort = strTxt1
    ElseIf intKort = 0 Or intLang = 0 Then
        Exit Function
    Else
        strLang = strTxt1
        strKort = strTxt2
        intX = intKort
        intKort = intLang
        intLang = intX
    End If
        
    For intPos = 1 To intKort 'Compare string based on the shortest.
        intX = 0
        Do
            intX = intX + 1
            strSearch = Mid$(strKort, intPos, intX) 'Determine part of string to search for
            If Len(strSearch) <> intX Then
                Exit Do 'end of string
            End If
        Loop While InStr(1, strLang, strSearch) > 0 'Part of string found in other string, increase size of partstring and try again.
        intX = intX - 1
        If intX > Len(strLangste) Then 'Longest substring found
            strLangste = Mid$(strKort, intPos, intX)
        End If
        If intX = 0 Then intX = 1
        intPos = intPos + intX - 1
    Next intPos

    If Len(strLangste) = 0 Then
        GetBiggest = "" 'No matching substring found
    Else 'Substring match found.
        'Split substring in left and right part.
        strTotal1 = Replace(strTxt1, strLangste, "|")
        strTotal2 = Replace(strTxt2, strLangste, "|")
            
        'Recursive part: Try again and paste result to returnvalue.
        GetBiggest = strLangste & _
                        GetBiggest(CStr(Split(strTotal1, "|")(0)), CStr(Split(strTotal2, "|")(0))) & _
                        GetBiggest(CStr(Split(strTotal1, "|")(1)), CStr(Split(strTotal2, "|")(1)))
    End If
    
End Function
