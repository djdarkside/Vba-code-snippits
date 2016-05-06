VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_04 - Run"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CloseButton_Click()
    DoCmd.Close
    DoCmd.OpenForm "00 - MainForm", acNormal
End Sub

Private Sub run_Click()
    Call folderbrowsing.ExportRunSheet
End Sub
