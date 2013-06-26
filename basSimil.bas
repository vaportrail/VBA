Attribute VB_Name = "basSimil"
Option Compare Database
Option Explicit

'   This is a vba port of Tom Van Stiphout's VB.Net/DLL implementation of Simil that was
'   written by Steve Grubb as an interpretation of the Ratcliff/Obershelp algorithm
'   for pattern recognition published in Dr. Dobb’s Journal in 1988 by Ratcliff and Metzener.
'   (Ratcliff and Metzener, “Pattern Matching: The Gestalt Approach”)

'   Tom wrote an amazing summary on using his implementation of this algorithm located on his
'   blog at (URL: http://www.accessmvp.com/tomvanstiphout/simil.htm)

'   I left most of Tom's comments in place, however please consult his blog for discrepancies

'   Steve Grubb's header below from simil.c
'   URL http://web.archive.org/web/20050213075957/www.gate.net/~ddata/utilities/simil.c

'   Simil - This is my hack at the Ratcliff/Obershelp Pattern
'   Recognition Algorithm as described in the July 1988 issue
'   of Dr. Dobbs Journal. This algorithm differs from strcmp in
'   that it returns a measure in percentage about how similar
'   two strings are. This rendition is purely in C for portability
'   reasons. The original was published in assembler for the small
'   memory model of the 386. The original used home brew stacks
'   for working space while this version handles everything by
'   recursion.

'   Copyright 1999  Steve Grubb  <linux_4ever@yahoo.com>

'   This program is free software; you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published  by
'   the Free Software Foundation; either version 2 of the License, or
'   (at your option) any later version.

'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.

'   You should have received a copy of the GNU General Public License
'   along with this program; if not, write to the Free Software
'   Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307  USA

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
                    'Note:
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
