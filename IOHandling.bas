Attribute VB_Name = "IOHandling"
Option Compare Database

Public Function SaveToFile(ByVal ArticleId As Long, FolderPath As String, FileType As String) As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim field As DAO.field
    Dim sql As String
    Dim output As String
    Dim filePath As String
    Dim FileName As String
    Dim fso As Object
    Dim ts As Object
    Dim timestamp As String


    'select the data to build a text file with the corresponding markdown code so it can be easily edited in the markdown form
    
    sql = "SELECT " & _
            "Articles.Id AS Id, Articles.Title As Title, Articles.Body, Articles.Created, Articles.Updated AS [Last Update] " & _
            "FROM ((Articles " & _
            "INNER JOIN AuthorsArticles ON Articles.Id = AuthorsArticles.Article_Id) " & _
            "INNER JOIN Users ON Users.Id = AuthorsArticles.User_Id) " & _
            "WHERE Articles.Id = " & ArticleId

    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql)
    
    If rs.EOF Then
        SaveToFile = "No records found for ArticleId: " & ArticleId
        Exit Function
    End If

    Do Until rs.EOF
        ' Loop through each field in the recordset and append its value to the output string
        For Each field In rs.fields
            If Not IsNull(field.Value) Then
                
                If field.name = "Title" Then
                    output = output & "# " & field.Value & vbCrLf
                    FileName = field.Value & "." & FileType
                ElseIf field.name = "Body" Then
                
                    output = output & field.Value & vbCrLf
                ElseIf field.name = "Created" Or field.name = "Last Update" Then
                    output = output & "- " & field.name & " : " & field.Value & vbCrLf
                    
                Else
                    output = output & field.name & " : " & field.Value & vbCrLf
                    
                End If
                
                
            Else
                output = output & "NULL" & vbCrLf
                If field.name = "Title" Then
                    FileName = "No_Title" & "." & FileType
                End If
            End If
        Next field
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    'get the associated files from the files table so we can add them to the resulting file
    
    output = output & GetFiles(ArticleId)


    ' add a timestamp to avoid file name collision
    
    timestamp = Format(Now(), "yyyyMMdd_HHmmss")
    
    ' add a generic name if there is no title associated with the article
    
    If FileName = "" Then
        FileName = "No_Title" & "." & FileType
    End If

    'create the filepath  for the filesystem object
    
    filePath = FolderPath & "\" & timestamp & "_" & FileName

    ' Create a FileSystemObject to write to a text file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(filePath, True)

    ' Write the output string to the text file
    ts.Write output

    ' Close the text stream
    ts.Close

    'return a filepath after saving the file
    SaveToFile = filePath
    
End Function

Public Function GetFiles(ByVal ArticleId As Long) As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim field As DAO.field
    Dim sql As String
    Dim output As String
    Dim fileId As String
    Dim sequence As String
    Dim WebPath As String
    Dim name As String
    
    'get the corresponding data for associated files based on the article id passed to the function
    
     sql = "SELECT Files.Id As FileId, Files.WebPath, Files.FileName As Name, ArticlesFiles.Sequence FROM (Files " & _
            "INNER JOIN ArticlesFiles ON Files.Id = ArticlesFiles.File_Id) " & _
            "WHERE ArticlesFiles.Article_Id = " & ArticleId & " " & _
            "ORDER BY ArticlesFiles.Sequence"
          
    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql)
    
    If rs.EOF Then
        GetFiles = ""
        Exit Function
    End If
    
    'for every record create a formatted markdown string so it can be easily copied and pasted into the markdown form
    
    Do Until rs.EOF
        For Each field In rs.fields
            
            If field.name = "FileId" Then
            
                fileId = field.Value
            
            End If
            
            If field.name = "WebPath" Then
            
                WebPath = field.Value
            
            End If
            
            If field.name = "Sequence" Then
                
                sequence = field.Value
            
            End If
            
            If field.name = "Name" Then
                name = field.Value
            End If
        
        Next field
        
        output = output & "![" & fileId & "_" & name & "_" & sequence & "](" & WebPath & ") " & vbCrLf & vbCrLf
        
        rs.MoveNext
    Loop
    
    GetFiles = output
          
End Function

