VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_01 - DefaultValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub CloseButton_Click()
    DoCmd.Close
    DoCmd.OpenForm "00 - MainForm", acNormal
End Sub

Private Sub SetDef_Click()
    On Error GoTo getout
    Dim str As String
    
    If IsNull(Me.SPORT) Or Me.SPORT = "" Then
        MsgBox "Missing School Sport Type", vbCritical
        Exit Sub
    End If
    If IsNull(Me.SCHOOL) Or (Me.SCHOOL = "") Then
        MsgBox "Missing School Name", vbCritical
        Exit Sub
    End If
    
    If MsgBox("Verify all and set default values?", vbExclamation + vbYesNo) = vbYes Then
        CurrentDb.TableDefs("MainData").Fields("SPORT").DefaultValue = """" & Me.SPORT & """"
        CurrentDb.TableDefs("MainData").Fields("School Name").DefaultValue = """" & Me.SCHOOL & """"
        CurrentDb.TableDefs("MainData").Fields("PRICELIST").DefaultValue = """" & Me.PRICELIST & """"
    End If
    
    MsgBox "All required default values were set", vbInformation
    DoCmd.Close
    DoCmd.OpenForm "00 - MainForm", acNormal
    Exit Sub

getout:
    MsgBox "Table default values were not set correctly. Try again or see Mitchell for help.", vbExclamation

End Sub





