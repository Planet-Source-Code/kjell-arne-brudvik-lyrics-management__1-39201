Attribute VB_Name = "modFileInfo"
Public Function GetFilePath(FileNamePath As String) As String

    On Error GoTo FunctionError:
    
    Dim x
    Dim tString As String

    For x = Len(FileNamePath) To 0 Step -1
        tString = Mid$(FileNamePath, x, 1)
        If tString = "\" Then
            GetFilePath = Left(FileNamePath, x)
            Exit Function
        End If
    Next x

FunctionError:
    
End Function

Public Function GetFileName(FileNamePath As String) As String
    
    On Error GoTo FunctionError:
    
    Dim x
    Dim tString As String
    Dim tType As String

    For x = Len(FileNamePath) To 0 Step -1
        tString = Mid$(FileNamePath, x, 1)
        If tString = "\" Then
            tString = x + 1
            Exit For
        End If
    Next x

    For x = Len(FileNamePath) To 0 Step -1
        tType = Mid$(FileNamePath, x, 1)
        If tType = "." Then
            tType = x
            Exit For
        End If
    Next x
    GetFileName = Mid$(FileNamePath, tString, tType - tString)
    Exit Function

FunctionError:

End Function


Public Function GetFileNameAndType(FileNamePath As String) As String
    
    On Error GoTo FunctionError:
    
    Dim x
    Dim tString As String

    For x = Len(FileNamePath) To 0 Step -1
        tString = Mid$(FileNamePath, x, 1)
        If tString = "\" Then
            GetFileNameAndType = Right(FileNamePath, Len(FileNamePath) - x)
            Exit Function
        End If
    Next x

FunctionError:

End Function

Public Function GetFileType(FileNamePath As String) As String
    
    On Error GoTo FunctionError:
    
    Dim x
    Dim tString As String

    For x = Len(FileNamePath) To 0 Step -1
        tString = Mid$(FileNamePath, x, 1)
        If tString = "." Then
            GetFileType = Right(FileNamePath, Len(FileNamePath) - x)
            Exit Function
        End If
    Next x

FunctionError:

End Function


Public Function FileLenght(FileNamePath As String) As String
    
    On Error GoTo FunctionError:
    FileLenght = FileLen(FileNamePath)

FunctionError:

End Function

Public Function FileDate(FileNamePath As String) As String
    
    On Error GoTo FunctionError:
    FileDate = FileDateTime(FileNamePath)

FunctionError:

End Function
