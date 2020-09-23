Attribute VB_Name = "modRegkey"
Option Explicit
Global tmpSearchCount As Long
Global tmpAddCount As Long

Function GenerateKey() As String

'Variables
    Dim I As Integer
    Dim intTemp As Integer
    Dim strTemp As String
    Dim strTemp2 As String
    Dim strTemp3 As String
    
'Randomly Randomize
    Randomize
    For I = 1 To (Rnd * 25)
        Randomize
    Next
    
'Generate 10 Random Letters
    For I = 1 To 10
        strTemp = strTemp & Chr((Rnd * 25) + 65)
    Next
    
'Replace "0" with "T" (O looks like 0)
    strTemp = Replace(strTemp, "O", "T")
    
'Catch Last Numerical Character of each letter
    For I = 1 To Len(strTemp)
        strTemp2 = strTemp2 & Right(CStr(Asc(Mid(strTemp, I, 1))), 1)
    Next
    
'Add Up Each Individual Number into a Sum
    For I = 1 To Len(strTemp2)
        intTemp = intTemp + CInt(Mid(strTemp2, I, 1))
    Next
   
'Convert Numbers to Characters
    For I = 1 To Len(strTemp2)
        strTemp3 = strTemp3 & Chr(65 + CInt(Mid(strTemp2, I, 1)))
    Next
    strTemp2 = strTemp3
    
'Convert MM to String
    strTemp3 = ""
    For I = 1 To Len(CStr(intTemp))
        strTemp3 = strTemp3 & Chr(65 + CInt(CStr(Mid(intTemp, I, 1))))
    Next
    
'Return Keys
    GenerateKey = Mid(strTemp, 1, 3) & Mid(strTemp2, 1, 2) & "-" & _
                  Mid(strTemp, 4, 2) & Mid(strTemp2, 3, 1) & Mid(strTemp, 6, 1) & Mid(strTemp2, 4, 1) & "-" & _
                  Mid(strTemp2, 5, 3) & Mid(strTemp, 7, 2) & "-" & _
                  Mid(strTemp, 9, 1) & Mid(strTemp2, 8, 2) & Mid(strTemp, 10, 1) & Mid(strTemp2, 10, 1) & "-" & _
                  Mid(strTemp, 5, 1) & Mid(strTemp2, 1, 1) & strTemp3 & Mid(strTemp, 1, 1)

'Pattern:         AAANN-AANAN-NNNAA-ANNAN-ANXXN
    
End Function

Function ValidateKey(strKey1 As String, strKey2 As String, strKey3 As String, _
                     strKey4 As String, strKey5 As String) As Boolean
'Variables
    Dim I As Integer
    Dim strKey As String
    Dim strChr As String
    Dim strNum As String
    Dim strCode As String
    Dim strTemp As String
    Dim strTemp2 As String
    Dim intTemp As Integer

'No Blanks
    If Trim(strKey1) = "" Or Trim(strKey2) = "" Or Trim(strKey3) = "" Or _
       Trim(strKey4) = "" Or Trim(strKey5) = "" Then
            ValidateKey = False
            Exit Function
    End If
    
'All Must Be 5Chrs
    If Len(strKey1) <> 5 Or Len(strKey2) <> 5 Or Len(strKey3) <> 5 Or _
       Len(strKey4) <> 5 Or Len(strKey5) <> 5 Then
            ValidateKey = False
            Exit Function
    End If
    
'Last Char of Key5 Must be the First Char of Key1
    If Mid(strKey1, 1, 1) <> Right(strKey5, 1) Then
        ValidateKey = False
        Exit Function
    End If
    
'Setup Key
    strKey = strKey1 & strKey2 & strKey3 & strKey4 & strKey5

'Assemble String of 10 Characters
    strChr = Mid(strKey, 1, 3) & Mid(strKey, 6, 2) & Mid(strKey, 9, 1) & _
              Mid(strKey, 14, 3) & Mid(strKey, 19, 1)
              
'Assemble String of What will be 10 Numbers
    strNum = Mid(strKey, 4, 2) & Mid(strKey, 8, 1) & Mid(strKey, 10, 4) & _
             Mid(strKey, 17, 2) & Mid(strKey, 20, 1)
             
'Collect What will be the code
    strCode = Right(strKey, 5)
    
'Make Sure 3 Characters in strCode Match with Appropriate Number and Alpha Keys
    If Mid(strCode, 1, 1) <> Mid(strChr, 5, 1) Then
        ValidateKey = False
        Exit Function
    End If
    If Mid(strCode, 2, 1) <> Mid(strNum, 1, 1) Then
        ValidateKey = False
        Exit Function
    End If
    If Right(strCode, 1) <> Mid(strChr, 1, 1) Then
        ValidateKey = False
        Exit Function
    End If
    
'Convert strNum Characters to ASC and Keep Last Digit
    For I = 1 To Len(strNum)
        strTemp = strTemp & Right(CStr((Asc(Mid(strNum, I, 1)) - 65)), 1)
    Next
    
'Convert strChr Characters to ASC and Keep Last Digit
    For I = 1 To Len(strChr)
        strTemp2 = strTemp2 & Right(CStr((Asc(Mid(strChr, I, 1)))), 1)
    Next
    
'The ASC in strTemp must match the asc in strTemp2
    If strTemp <> strTemp2 Then
        ValidateKey = False
        Exit Function
    End If

'Add Up SUm of Each Unique Number in ASC of Chr String
    For I = 1 To Len(strNum)
        intTemp = intTemp + CInt(Right(CStr((Asc(Mid(strNum, I, 1)) - 65)), 1))
    Next
    strTemp = CStr(intTemp)
    
'Resize the Code & Extract Num Val
    strCode = Mid(strCode, 3, 2)
    strCode = Right(CStr((Asc(Mid(strCode, 1, 1)) - 65)), 1) & _
              Right(CStr((Asc(Mid(strCode, 2, 1)) - 65)), 1)
    
'Does the Extracted Num match the Group Sum?
    If strTemp <> strCode Then
        ValidateKey = False
        Exit Function
    End If
    
'Key Looks Good
    ValidateKey = True

End Function
