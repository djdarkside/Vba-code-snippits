Attribute VB_Name = "folderbrowsing"
Option Compare Database

Function FileExists(FILENAME) As Boolean
    FileExists = (Dir(FILENAME) <> "")
End Function
'NEW FOLDER BROWSING FUNCTION
Function GetFolder(strPath As String) As String
    Dim fldr As FileDialog
    Dim sItem As String
    Dim sInitDir As String
    sInitDir = CurDir ' Store initial directory
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = strPath
    If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    ' Reset directory before exit
    ChDrive sInitDir ' Return to the Initial Drive
    ChDir sInitDir ' Resets directory for Initial Drive
    GetFolder = sItem
    Set fldr = Nothing
End Function

Sub ListFilesInFolder(SourceFolderName As String, IncludeSubfolders As Boolean)
    Dim Application As New Application
    Dim FSO As New FileSystemObject, SourceFolder As Folder, Subfolder As Folder, FileItem As File
    Dim lngCount As Long, strSQL, strSQL1 As String
    
    Set SourceFolder = FSO.GetFolder(SourceFolderName)
    
    For Each FileItem In SourceFolder.Files
        Select Case Right(FileItem.Name, 3)
            Case "JPG"
                strSQL = "INSERT INTO MainData (IMAGE,IMAGE_PATH,[Image Name]) SELECT '" & Replace(FileItem.Name, ".jpg", "") & "','" & SourceFolder.Path & "','" & FileItem.Name & "'"
                DoCmd.RunSQL strSQL
                DoCmd.SetWarnings False 'disable warnings
        End Select
        lngCount = lngCount + 1
    Next FileItem
    If IncludeSubfolders Then
        For Each Subfolder In SourceFolder.Subfolders
            ListFilesInFolder Subfolder.Path, True
            DoCmd.SetWarnings False 'disable warnings
        Next Subfolder
    End If
    Set FileItem = Nothing
    Set SourceFolder = Nothing
    DoCmd.SetWarnings True

End Sub

Sub ExportRunSheet()
    Dim dbs As Database
    Dim PCKExport As DAO.Recordset
    Dim PCKLoc As DAO.Recordset
    Dim Path As String
    Dim PCKExportName As String
    Set dbs = CurrentDb
    Set PCKExport = dbs.OpenRecordset("MainData", dbOpenDynaset)
    Set PCKLoc = dbs.OpenRecordset("FTP_DATA", dbOpenDynaset)

    Path = PCKLoc![LOCALDIR]
'SPECIFY AT END .txt, .csv, or .xlx
    PCKExportName = Path + "\" + PCKExport![School Name] + "-" + PCKExport![SPORT] + " Package Export.csv"
    DoCmd.TransferText acExportDelim, , "03 - Test Union", PCKExportName, True
    MsgBox "Sports Pack Export has been saved to " & Path & " ", vbInformation, "Sports Pack Export"
End Sub

Public Sub startup()
    DoCmd.ShowToolbar "Ribbon", acToolbarNo
    DoCmd.ShowToolbar "Status Bar", acToolbarNo
    DoCmd.NavigateTo "acNavigationCategoryObjectType"
    DoCmd.RunCommand acCmdWindowHide
End Sub
