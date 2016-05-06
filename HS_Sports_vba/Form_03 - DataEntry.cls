VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_03 - DataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Check19_Click()
If Me.Check19 = -1 Then
    If MsgBox("Do you want to set the current team as Default?   Please make sure you know what you are doing or see Mitchell before proceeding!!!     ", vbYesNo, "Team Default Names") = vbYes Then
        Me.Check19 = -1
    Else
        Me.Check19 = 0
    End If
Else
Exit Sub
End If
End Sub



Private Sub Check46_Click()
    If Me.Check46 = True Then
        Me.Text0 = "COACH"
        Me.Text2 = "COMP"
        Me.Check17 = True
        Me.Package.SetFocus
        Me.Refresh
    ElseIf Me.Check46 = False Then
        Me.Text0 = ""
        Me.Text2 = ""
        Me.Check17 = False
        Me.Refresh
    End If
    Me.Refresh
End Sub

Private Sub CloseButton_Click()
    DoCmd.Close
    DoCmd.OpenForm "00 - MainForm", acNormal
End Sub

Private Sub Command16_Click()
    DoCmd.GoToRecord , , acNewRec
End Sub

Private Sub Command19_Click()
    Dim FlyerToOpen As String
On Error GoTo errorhandle
    FlyerToOpen = Trim(DLookup("[PDFpath]", "PriceList", "[PriceCode] = '" & Trim(Me.PRICELIST) & "'"))
    'MsgBox FlyerToOpen
    
    'verify file before open
    If FileExists(FlyerToOpen) Then
        Shell "C:\WINDOWS\explorer.exe """ & FlyerToOpen & "", vbNormalFocus
    Else
        MsgBox "Your flyer PDF can't be located at " & FlyerToOpen, vbCritical, "Invalid file path"
        Exit Sub
    End If
    
    Exit Sub
errorhandle:
    MsgBox Err.Description + "  . Please see Mitchell", vbCritical, "PDF Location Error"
End Sub

Private Sub Form_Current()
    
    Dim Path As String
    Dim lastteam As String
    Dim dbs As Database
    Dim rs As DAO.Recordset
    Set dbs = CurrentDb
    Set rs = dbs.OpenRecordset("00 - All Data", dbOpenDynaset)
    Path = "\\S1P-SPECTRA1\SpectraPhotos\HS SPORTS\15-16 High School Sports\NoImage.JPG"
    
    ''''''''''''''''''''''''''''''''''''SETS PIC AND IF NO PIC OR TEAM ONLY SHOWS NOPIC.JPG
    If Me.ORDER_SEQ <> "X" Then
        If Me.ORDER_SEQ = "TEAM ONLY" Then
            Me.StudImg.Picture = Path
            Exit Sub
        Else
        Me.StudImg.Picture = Me.ImgPath & "/" & Me.ORDER_SEQ & ".JPG"
        Me.TEAM.SetFocus
        End If
    Else: Me.StudImg.Picture = Path
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''''''''SETS DEFAULT TEAM
    If Me.Check19 = -1 Then
            lastteam = DLookup("[TEAM]", "00 - All Data", "[ID]=Forms![03 - DataEntry]![ID]-1")
            Me.TEAM = lastteam
    End If
    ''''''''''''''''''''''''''''''''''''END DEFAULT TEAM
    Exit Sub
errorhandle:
    If Err.Number = 2220 Then
        MsgBox "Your image can't be located at " & Me.ImgPath, vbCritical, "Invalid file path"
    Else
        MsgBox "Error not handled.", vbCritical
    End If
    Exit Sub
    
End Sub

Private Sub Form_Load()
    Call folderbrowsing.startup
End Sub

Private Sub GrandTotal_AfterUpdate()
    Me.Refresh
End Sub
Private Sub GrandTotal_Change()
    Me.Refresh
End Sub
Private Sub GrandTotal_Enter()
    Me.Refresh
End Sub
Private Sub ImageCopy_Click()
    Dim FullImagePath, DestinationFolder, DestinationFullPath As String
On Error GoTo errorhandle
    If Me.ORDER_SEQ <> "X" Then
        FullImagePath = Me.ImgPath + "\" + Me.ORDER_SEQ + ".jpg"
        If FileExists(FullImagePath) Then
            DestinationFolder = GetFolder("\\S1P-SPECTRA1\SpectraPhotos\HS SPORTS\15-16 High School Sports")
            
            'CHECK FOR CANCEL BUTTON-------------------------------------------------------------------------------------
            If IsNull(DestinationFolder) Or DestinationFolder = "" Then
                MsgBox "An destination folder must be specified to perform copy function!", vbExclamation, "Destination Path Missing"
                Exit Sub
            Else
                If Right(DestinationFolder, 1) = "\" Then
                    DestinationFolder = Left(DestinationFolder, Len(DestinationFolder) - 1)
                End If
                
                'check to see if user wanna append school name as prefix
                If MsgBox("Do you want to use Team Name as prefix?", vbQuestion + vbYesNo, "Prefix confirmation") = vbYes Then
                    DestinationFullPath = DestinationFolder + "\" + Me.Text10 + "_" + Me.Text12 + "_" + Me.TEAM + "_" + Me.ORDER_SEQ + ".jpg"
                Else
                    DestinationFullPath = DestinationFolder + "\" + Me.ORDER_SEQ + ".jpg"
                End If

                FileCopy FullImagePath, DestinationFullPath
            End If
        Else
            MsgBox "Your image can not be verified. Please check your path", vbCritical, "Missing image jpg"
            Exit Sub
        End If
        Me.Refresh
    End If
    
    'open destination folder
    'Shell "C:\WINDOWS\explorer.exe """ & DestinationFolder & "", vbNormalFocus
    'Exit Sub
errorhandle:
    MsgBox "No Image Found.  ", vbCritical
End Sub

Private Sub NameLookup_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object
    Dim FullImgPath As String
    
    Set rs = Me.Recordset.Clone
    rs.FindFirst "[LASTFIRST] = '" & Me![NameLookup] & "'"
    If rs.NoMatch Then
        'Me.NameLookup.SetFocus
        MsgBox "Student is not found.  Please verify name spelling", vbCritical, "Name not found"
    Else
        Me.Bookmark = rs.Bookmark
        Me.NameLookup = ""
    End If
    
End Sub

Private Sub NextRec_Click()
On Error GoTo Err_NextRec_Click

    DoCmd.GoToRecord , , acNext

Exit_NextRec_Click:
    Exit Sub

Err_NextRec_Click:
    MsgBox Err.Description
    Resume Exit_NextRec_Click
End Sub

Private Sub OpenImgFolder_Click()
    Dim ImageFolder As String
    ImageFolder = Me.ImgPath
    
On Error GoTo errorhandle
    'open destination folder
    'MsgBox ImageFolder
    Shell "C:\WINDOWS\explorer.exe """ & ImageFolder & "", vbNormalFocus
    Exit Sub
    
errorhandle:
    MsgBox "No Image or Path Found.", vbCritical

End Sub

Private Sub PACKAGE_AfterUpdate()
    Me.Package = UCase(Me.Package)
End Sub

Private Sub PACKAGE_BeforeUpdate(Cancel As Integer)
    Me.Refresh
End Sub

Private Sub PACKAGE_Change()
    Me.Refresh
End Sub

Private Sub PACKAGE_Exit(Cancel As Integer)
    If Me.Package <> "" And Not IsNull(Me.Package) Then
        Me.PckTotal = CalculatePrice(Me.Package, Me.PRICELIST)
    Else
        Me.PckTotal = 0
    End If
    Me.GrandTotal = Me.PckTotal
    Me.Refresh
End Sub

Private Sub PrevRec_Click()
On Error GoTo Err_PrevRec_Click

    DoCmd.GoToRecord , , acPrevious

Exit_PrevRec_Click:
    Exit Sub

Err_PrevRec_Click:
    MsgBox Err.Description
    Resume Exit_PrevRec_Click
End Sub

Private Sub Slate_Click()
    If Me.Slate = True Then
        Me.Text0 = "SCHOOL"
        Me.Text2 = "COMP"
        Me.Check17 = True
        Me.Package.SetFocus
        Me.Refresh
    ElseIf Me.Slate = False Then
        Me.Text0 = ""
        Me.Text2 = ""
        Me.Check17 = False
        Me.Refresh
    End If
    Me.Refresh
End Sub

Private Sub StudImg_DblClick(Cancel As Integer)
    Dim ImgFullPath As String
    ImgFullPath = Me.ImgPath & "\" & Me.ORDER_SEQ & ".jpg"
    'open image
    If ImgFullPath = "" Or IsNull(ImgFullPath) Then
        Exit Sub
    Else
        Shell "C:\WINDOWS\explorer.exe """ & ImgFullPath & "", vbNormalFocus
    End If
End Sub
