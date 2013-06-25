Attribute VB_Name = "basSimil"
Option Compare Database
Option Explicit

Private Type structSubString
    o1 As Integer
    o2 As Integer
    len1 As Integer
End Type

Public Function Simil(ByVal s1 As String, ByVal s2 As String) As Double
    Dim returnValue As Double
    Dim tlen As Integer
    Dim s1len As Integer
    Dim s2len As Integer

    ' Added by Jamie West, removes spaces before and after string
    s1 = Trim(s1)
    s2 = Trim(s2)

    returnValue = 0
    s1len = Len(s1)
    s2len = Len(s2)

    If s1len = 0 Or s2len = 0 Then
        returnValue = 0#
    Else
        Dim tcnt As Integer
        tcnt = 0
        tlen = s1len + s2len
        Call rsimil(s1, s1len, s2, s2len, tcnt)
        returnValue = tcnt / tlen
    End If
        Simil = returnValue
End Function

Private Sub rsimil(ByVal s1 As String, ByVal s1len As Integer, ByVal s2 As String, ByVal s2len As Integer, ByRef tcnt As Integer)
    Dim ss As structSubString

    If s1len = 0 Or s2len = 0 Then Exit Sub

    Call find_biggest_substring(s1, s1len, s2, s2len, ss)
    
    If ss.len1 <> 0 Then
        tcnt = tcnt + (ss.len1 * 2)

        'Check left half...
        Call rsimil(s1, ss.o1, s2, ss.o2, tcnt)

        'Check right half...
        Dim delta1 As Integer
        delta1 = ss.o1 + ss.len1
        
        Dim delta2 As Integer
        delta2 = ss.o2 + ss.len1
        
        If delta1 < s1len And delta2 < s2len Then
            Call rsimil(Mid(s1, delta1 + 1, Len(s1) - delta1), s1len - delta1, Mid(s2, delta2 + 1, Len(s2) - delta2), s2len - delta2, tcnt)
        End If
    End If
End Sub

Private Sub find_biggest_substring(ByVal s1 As String, ByVal s1len As Integer, ByVal s2 As String, ByVal s2len As Integer, ByRef ss As structSubString)
    Dim i As Integer
    Dim j As Integer
    Dim size As Integer
    
    size = 1

    ss.o2 = -1
    i = 0
    Do While i <= (s1len - size)
        j = 0
        Do While j <= (s2len - size)
            Dim test_size As Integer
            test_size = size
            Do While (1)
                If ((test_size <= (s1len - i)) And (test_size <= (s2len - j))) Then
                    'While things match, keep trying...
                    'Note: String.Equals performs an ordinal (case-sensitive and culture-insensitive) comparison.
                    If Mid(s1, i + 1, test_size) = Mid(s2, j + 1, test_size) Then
                        If ((test_size > size) Or (ss.o2 < 0)) Then
                            ss.o1 = i
                            ss.o2 = j
                            size = test_size
                        End If
                        test_size = test_size + 1
                        Else
                        'Not equal
                        Exit Do
                    End If
                    Else
                    'Gone past the end of a string - we're done.
                    Exit Do
                End If
            Loop
            j = j + 1
        Loop
        i = i + 1
    Loop
    If ss.o2 < 0 Then
        ss.len1 = 0
    Else
        ss.len1 = size
    End If
End Sub
