Attribute VB_Name = "basFuzzyAddress"
Option Compare Database
Option Explicit

Public Function fuzzyCompare(v1 As String, v2 As String) As Double
    Dim len1, len2 As Integer
    Dim v1c As String
    Dim v2c As String
    Dim fc As Integer
    Dim fcs As Double
    Dim lc As Integer
    Dim lcs As Double
    
    v1c = collapseString(v1)
    v2c = collapseString(v2)
    
    v1c = replaceWords(v1c, "avenue", "ave")
    v1c = replaceWords(v1c, "road", "rd")
    v1c = replaceWords(v1c, "street", "st")
    v1c = replaceWords(v1c, "east", "e")
    v1c = replaceWords(v1c, "place", "pl")
    v1c = replaceWords(v1c, "north", "n")
    v1c = replaceWords(v1c, "south", "s")
    
    v2c = replaceWords(v2c, "avenue", "ave")
    v2c = replaceWords(v2c, "road", "rd")
    v2c = replaceWords(v2c, "street", "st")
    v2c = replaceWords(v2c, "east", "e")
    v2c = replaceWords(v2c, "place", "pl")
    v2c = replaceWords(v2c, "north", "n")
    v1c = replaceWords(v1c, "south", "s")

    len1 = Len(v1c)
    len2 = Len(v2c)
    
    fc = compareFrequency(v1c, v2c)
    fcs = 1 - fc / (len1 + len2)
    
    Dim mlen As Integer
    If len1 > len2 Then
        mlen = len1
    Else
        mlen = len2
    End If
    
    lc = compareLongest(v1c, v2c)
    
    lcs = lc / mlen
    ' MsgBox mlen
    
    fuzzyCompare = 0.2 * lcs + 0.8 * fcs
    
End Function

Private Function replaceWords(ByVal v As String, w1 As String, w2 As String) As String
    Dim i As Integer
    Dim ln As Integer
    Dim wln As Integer
    Dim r As String
    Dim s As String
    
    ln = Len(v)
    wln = Len(w1)
    
    For i = 1 To ln
        s = Mid(v, i, wln)
        If s = w1 Then
            r = r + w2
            i = i + wln - 1
        Else
            r = r + Mid(v, i, 1)
        End If
    Next i
    replaceWords = r
End Function

Private Function compareLongest(ByVal v1 As String, ByVal v2 As String) As Integer
    Dim len1 As Integer
    Dim len2 As Integer
    Dim mlen As Integer
    Dim i As Integer
    Dim c1 As Integer
    Dim s1 As String
    
    Dim c2 As Integer
    Dim s2 As String
    
    len1 = Len(v1)
    len2 = Len(v2)
    
    If len1 > len2 Then
        mlen = len2
    Else
        mlen = len1
    End If
    
    
    Dim c As Integer
    
    For i = 1 To mlen
        s1 = Mid(v1, i, 1)
        s2 = Mid(v2, i, 1)
        
        c1 = Asc(s1)
        c2 = Asc(s2)
        
        If c1 = c2 Then
            c = c + 1
        End If
        
    Next i
    compareLongest = c
End Function

Private Function compareFrequency(ByVal v1 As String, ByVal v2 As String) As Integer
    Dim a(0 To 255) As Integer
    Dim i As Integer
    Dim len1 As Integer
    Dim len2 As Integer
    Dim c As Integer
    Dim s As String
    Dim sum As Integer
    
    len1 = Len(v1)
    
    For i = 1 To len1
        s = Mid(v1, i, 1)
        c = Asc(s)
        If i < 6 Then
            a(c) = a(c) + 2
        Else
            
        End If
    Next i
    
    len2 = Len(v2)
    
    For i = 1 To len2
        s = Mid(v2, i, 1)
        c = Asc(s)
        
        If i < 6 Then
            a(c) = a(c) - 2
        Else
            
        End If
    Next i
    
    For i = 48 To 57
        sum = sum + Abs(a(i))
    Next i
    
    For i = 97 To 122
        sum = sum + Abs(a(i))
    Next i
    
    compareFrequency = sum
    
End Function

Private Function collapseString(ByVal v As String) As String
    Dim i As Integer
    Dim ln As Integer
    ln = Len(v)
    Dim r As String
    Dim s As String
    Dim c As Integer
    
    For i = 1 To ln
        s = LCase(Mid(v, i, 1))
        c = Asc(s)
        If (c >= 48 And c <= 57) Or (c >= 97 And c <= 122) Then
           r = r + s
        End If
        
    Next i
    
    collapseString = r
End Function
