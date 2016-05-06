VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_05 - FTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CloseButton_Click()
    DoCmd.Close
    DoCmd.OpenForm "00 - MainForm", acNormal
End Sub

Private Sub SendFtp_Click()
    Dim strDirectoryList As String
    Dim lStr_Dir As String
    Dim lInt_FreeFile01 As Integer
    Dim dbs As Database
    Dim rst As DAO.Recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("MainData", dbOpenDynaset)
    
    'Error Handler
    On Error GoTo Err_Handler
    
    'Asks If you want to run the method
    If MsgBox("Are you ready to send the images to PCL", vbYesNo, "PCL FTP") = vbYes Then
        'Check If you want to export the runsheet
        If MsgBox("Do you want to Export the Excel RunSheet?  If you have already click No.", vbYesNo, "PCL FTP") = vbYes Then
            Call folderbrowsing.ExportRunSheet
        End If
    
    'Send Email Calls EMAILPCL MODULE
        Call EMAILPCL.EMAILPCL
        
    ' Create script file with FTP commands
        lStr_Dir = rst![IMAGE_PATH]
        lInt_FreeFile01 = FreeFile
        strDirectoryList = "C:\ExportText" & "\FTP_TEXTFILE"
        Open strDirectoryList & ".txt" For Output As #lInt_FreeFile01
        Print #lInt_FreeFile01, "open " & DLookup("[FTP_SERVER]", "FTP_DATA")
        Print #lInt_FreeFile01, "user " & DLookup("[USER]", "FTP_DATA") & " " & DLookup("[PASSWORD]", "FTP_DATA")
        Print #lInt_FreeFile01, "lcd " & Chr(34) & lStr_Dir & Chr(34)
        Print #lInt_FreeFile01, "cd " & Chr(34) & "&STUDIO1_TO_PCL/HS_SPORTS/" & Chr(34)
        Print #lInt_FreeFile01, "mkdir " & Chr(34) & rst![School Name] & " " & rst![SPORT] & Chr(34)
        Print #lInt_FreeFile01, "cd " & Chr(34) & rst![School Name] & " " & rst![SPORT] & Chr(34)
        Print #lInt_FreeFile01, "binary"
        Print #lInt_FreeFile01, "prompt"
        Print #lInt_FreeFile01, "mput *"
        Print #lInt_FreeFile01, "bye"
        Close #lInt_FreeFile01
        
    'Send the files and run the script Uses WScript
        Dim FTPcommand As String
        Dim wsh As Object
        Dim errorCode As Integer
        FTPcommand = "ftp -n -s:" & Chr(34) & strDirectoryList & ".txt" & Chr(34)
        
        Set wsh = CreateObject("WScript.Shell")
        errorCode = wsh.run(FTPcommand, 1, True)
    
    'Checks if everything worked
        If errorCode = 0 Then
            MsgBox "Images have been uploaded", vbInformation, "PCL FTP"
        Else
            MsgBox "There were errors contact Mitchell to check them", vbCritical, "PCL FTP"
        End If
        
    'Deletes the text file
        If FileExists(strDirectoryList & ".txt") Then
            SetAttr strDirectoryList & ".txt", vbNormal
            Kill strDirectoryList & ".txt"
        End If
    Else
        MsgBox "Transfer Canceled!", vbCritical, "PCL FTP"
        Exit Sub
    End If
bye:
        Exit Sub
Err_Handler:
    MsgBox "Error : " & Err.Number & vbCrLf & "Description : " & Err.Description, vbCritical
    Resume bye
End Sub

