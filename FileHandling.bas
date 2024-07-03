Attribute VB_Name = "FileHandling"
Option Compare Database

Public Function GetFilePath() As String
    
    Dim fd As Object
    Dim selectedFile As String
    
    'create a filedialog object
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    'offer the options
    With fd
        .Title = "Select a file"
        .Filters.Clear
        .Filters.Add "Image Files", "*.png; *.jpg; *.jpeg"
        .Filters.Add "Word Documents", "*.doc; *.docx"
        .Filters.Add "PDF Files", "*.pdf"
        .Filters.Add "PowerPoint Presentations", "*.ppt"
        .Filters.Add "Text Files", "*.txt"
        
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
        Else
            selectedFile = ""
        End If
    End With
    Set fd = Nothing
    GetFilePath = selectedFile
End Function

Public Function VerifyFilePath(inputPath As String) As Boolean
    Dim FileName As String

    ' Check if the inputPath is not empty
    If inputPath <> "" Then
        ' Use Dir to check if the file exists
        FileName = Dir(inputPath)

        ' Return True if the file exists, otherwise False
        If FileName <> "" Then
            VerifyFilePath = True
        Else
            VerifyFilePath = False
        End If
    Else
        ' Return False if the inputPath is empty
        VerifyFilePath = False
    End If
End Function

Public Function GetFileNameFromPath(filePath As String) As String
    Dim pos As Integer
    Dim FileName As String
    'find the position of the last specified character
    pos = InStrRev(filePath, "\")
    If pos > 0 Then
        'extract the substring that starts after the position of the last backslash
        FileName = Mid(filePath, pos + 1)
        GetFileNameFromPath = FileName
        
    Else
        GetFileNameFromPath = "Unknown"
    End If
End Function

Public Function GetFileExtensionFromFileName(FileName As String) As String
    'get the file extension from the path if one is included
    Dim pos As Integer
    Dim Extension As String
    
    pos = InStrRev(FileName, ".")
    If pos > 0 Then
        'extract the extension after the last period
        Extension = Mid(FileName, pos + 1)
        GetFileExtensionFromFileName = Extension
    Else
        GetFileExtensionFromFileName = "Unknown"
    End If
End Function

Public Function GetFolderPathFromPath(filePath As String) As String
    Dim pos As Integer

    ' Find the position of the last backslash
    pos = InStrRev(filePath, "\")

    ' Extract the folder path from the full file path
    If pos > 0 Then
        GetFolderPathFromPath = Left(filePath, pos - 1)  ' Exclude the trailing backslash
    Else
        ' No backslash found, handle the path (e.g., root folder)
        GetFolderPathFromPath = ""
    End If
End Function

Public Function GetStorageFromPath(filePath As String) As String
    ' Ensure the path is long enough to evaluate

    If Len(filePath) >= 3 Then
        ' Check if the first two characters are double backslashes (UNC path)
        If Left(filePath, 2) = "\\" Then

            GetStorageFromPath = "network"
            Exit Function
        End If
        
        ' Check if the path starts with a letter, followed by a colon and a backslash (drive path)
        If Mid(filePath, 2, 1) = ":" And Mid(filePath, 3, 1) = "\" Then
            Dim firstChar As String
            firstChar = Left(filePath, 1)
            
            ' Check if the first character is a letter from A to Z (case insensitive)
            If firstChar Like "[A-Z]" Then

                GetStorageFromPath = "drive"
                Exit Function
            End If
        End If
    End If
    
    ' If none of the above conditions are met, the path format is unknown
    GetStorageFromPath = "unknown"
End Function

Public Function FilePathToUrl(ByVal rawPath As String) As String
    Dim validCharacters As String
    Dim encodedUrl As String
    Dim char As String
    Dim i As Integer
    Dim filePath As String
    
    filePath = CleanFilePath(rawPath)
    
    ' Valid URL characters including the forward slash
    validCharacters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_.~/"
    
    ' Replace backslashes with forward slashes
    filePath = Replace(filePath, "\", "/")
    
    encodedUrl = ""
    
    ' Loop through each character in the filePath
    For i = 1 To Len(filePath)
        char = Mid(filePath, i, 1)
        ' Check if the character is valid
        If InStr(validCharacters, char) > 0 Then
            encodedUrl = encodedUrl & char
        Else
            ' Encode invalid characters
            encodedUrl = encodedUrl & "%" & Right("0" & Hex(Asc(char)), 2)
        End If
    Next i
    
    FilePathToUrl = encodedUrl
End Function

Public Function CleanFilePath(filePath As String) As String
    Dim cleanPath As String
    
    ' Check if the filePath starts with "\\", remove it if it does
    If Mid(filePath, 1, 2) = "\\" Then
        filePath = Mid(filePath, 3, Len(filePath))
        
    End If
    
    ' Check if the filePath starts with a drive letter, colon, and backslash (e.g., C:\), remove it if it does
    If filePath Like "[A-Za-z]:\*" Then
        filePath = Mid(filePath, 4, Len(filePath))
        
    End If
    
    ' Set the cleaned path to the result
    cleanPath = filePath
    
    ' Return the cleaned file path
    CleanFilePath = cleanPath
End Function

