Attribute VB_Name = "Module1"
Dim FolderPathLCQD As String
Dim FolderPathQTON As String

Sub InitializeFolderPath(sheetName As String)
    Dim FolderPath As String
    

    FolderPath = InputBox("Enter the folder path for " & sheetName & ":")
    If FolderPath <> "" Then
       
        With ThisWorkbook.Sheets(sheetName)
            .Cells(1, 1).Value = "Folder Path"
            .Cells(1, 2).Value = FolderPath
        End With
    End If
End Sub

Sub CountFolders(sheetName As String)
    Dim FolderPath As String
    Dim FolderCount As Integer
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objSubFolder As Object
    
   
    FolderPath = ThisWorkbook.Sheets(sheetName).Cells(1, 2).Value
    
  
    If FolderPath = "" Then
        MsgBox "Folder path for " & sheetName & " is not set. Please initialize it first.", vbExclamation
        Exit Sub
    End If


    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(FolderPath)


    FolderCount = 0

    For Each objSubFolder In objFolder.SubFolders
        FolderCount = FolderCount + 1
    Next objSubFolder

    With ThisWorkbook.Sheets(sheetName)
        .Cells(2, 1).Value = "Folder Count"
        .Cells(2, 2).Value = FolderCount
    End With


    Set objSubFolder = Nothing
    Set objFolder = Nothing
    Set objFSO = Nothing
End Sub

Sub UpdateFolderCount()
 
    CountFolders "LCQD"
    CountFolders "QTON"
    Application.OnTime Now + TimeValue("00:01:00"), "UpdateFolderCount" '
End Sub

Sub Auto_Open()

    If ThisWorkbook.Sheets("LCQD").Cells(1, 2).Value = "" Then
        InitializeFolderPath "LCQD"
    End If
    If ThisWorkbook.Sheets("QTON").Cells(1, 2).Value = "" Then
        InitializeFolderPath "QTON"
    End If
    UpdateFolderCount
End Sub

Sub Workbook_Open()
    Auto_Open
End Sub
