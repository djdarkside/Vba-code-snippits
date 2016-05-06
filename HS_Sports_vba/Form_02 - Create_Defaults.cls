VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_02 - Create_Defaults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CloseButton_Click()
    DoCmd.Close
    DoCmd.OpenForm "00 - MainForm", acNormal
End Sub

Private Sub Command5_Click()
    DoCmd.OpenQuery "06 - SelectTeams", acNormal
End Sub

Private Sub Command6_Click()
    DoCmd.OpenQuery "AppendTeam", acNormal
End Sub

Private Sub Form_AfterUpdate()
    Me.Refresh
End Sub


Private Sub GEN_INDV_Click()
  
    Dim Counter As Integer, MyFile As String, strFileName As String, MissingFlag As Boolean
    Dim OutputFolder As String

    Dim dbs As Database, rst As DAO.Recordset, log As DAO.Recordset, rs2 As DAO.Recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("MainData", dbOpenDynaset)
    Set rst2 = dbs.OpenRecordset("FTP_DATA", dbOpenDynaset)
    
On Error GoTo ErrorHandler
    MissingFlag = False

    'If rst.RecordCount <> 0 Then
    '    MsgBox "You cannot add on top of data in the Table. Please delete all records in Table", vbCritical
    '    Exit Sub
    'End If

    OutputFolder = GetFolder("\\S1P-SPECTRA1\SpectraPhotos\HS SPORTS\15-16 High School Sports")
    rst2.Edit
    rst2![LOCALDIR] = OutputFolder
    rst2.Update

    If IsNull(OutputFolder) Or OutputFolder = "" Then
        MsgBox "The Schools R2P folder must be specified to proceed!", vbExclamation
        Exit Sub
    End If

'FUNCTION
    'False Turns of Browesing subfolders
    Call ListFilesInFolder(OutputFolder, False)
'END FUNCTION

    MsgBox "Main Table Created, Verify all Images are present before Continuing", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
End Sub






