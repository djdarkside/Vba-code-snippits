VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_05 - PACKING LABELS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Close()
    DoCmd.OpenForm "00 - MainForm", acNormal
End Sub

Private Sub Report_NoData(Cancel As Integer)
    Cancel = True
End Sub


