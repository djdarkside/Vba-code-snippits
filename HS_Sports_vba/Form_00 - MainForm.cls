VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_00 - MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command1_Click()
    DoCmd.Close
    DoCmd.OpenForm "01 - DefaultValue", acNormal
End Sub

Private Sub Command11_Click()
    Dim do_Mydata As DataObject
    Dim str_test As String
    Set do_Mydata = New DataObject
    str_test = CurrentDb.Name
    With do_Mydata
        .SetText str_test
        .PutInClipboard
    End With
End Sub

Private Sub Command12_Click()
    DoCmd.Close
    DoCmd.OpenForm "05 - FTP", acNormal
End Sub

Private Sub Command2_Click()
    DoCmd.Close
    DoCmd.OpenForm "02 - Create_Defaults", acNormal
End Sub

Private Sub Command3_Click()
    DoCmd.Close
    DoCmd.OpenForm "03 - DataEntry", acNormal
End Sub

Private Sub Command4_Click()
    DoCmd.Close
    DoCmd.OpenForm "04 - Run", acNormal
End Sub

Private Sub Command5_Click()
    DoCmd.Close
    DoCmd.OpenReport "05 - PACKING LABELS", acViewPreview
End Sub

Private Sub Command6_Click()
    DoCmd.OpenReport "03 - Proofing Report by Image", acViewPreview
End Sub

Private Sub Command7_Click()
    DoCmd.OpenQuery "03 - PROOFING CORRECTION"
End Sub

Private Sub Command8_Click()
    DoCmd.Close
    DoCmd.OpenReport "11 - FINANCE REPORT", acViewPreview
End Sub

Private Sub EMAILPCL_Click()
    Call EMAILPCL.EMAILPCL
End Sub
