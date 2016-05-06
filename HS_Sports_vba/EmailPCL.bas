Attribute VB_Name = "EmailPCL"
Option Compare Database

Public Sub EMAILPCL()
    Dim dbs As Database
    Dim rst As DAO.Recordset
    Dim oApp As New Outlook.Application
    Dim oEmail As Outlook.MailItem
    
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("00 - All Data", dbOpenDynaset)
'Loops through the recordset
    If rst.RecordCount = 0 Then
        MsgBox "There is nothing here!", vbInformation, "Email Sender"
        Exit Sub
    Else
        Set oEmail = oApp.CreateItem(olMailItem)
            oEmail.To = "orders@pclwest.com"
            oEmail.CC = "marvin@pclwest.com" & ";" & "claudia@pclwest.com"
            oEmail.Subject = rst![School Name] + " " + rst![SPORT] + " on FTP"
            oEmail.Body = "Hello PCL, " & Chr(13) & Chr(13) & _
                          rst![School Name] + " " + rst![SPORT] & " is now on the FTP for processing at &STUDIO1_TO_PCL/HS_SPORTS/" & rst![School Name] + " " + rst![SPORT] & Chr(13) & _
                          "For further assistance, please give us a call at (925)361-0430." & Chr(13) & Chr(13) & _
                          "Studio One Photography Production Department"
            oEmail.Send
    End If
    MsgBox "Email has been sent to PCL", vbInformation, "Email Sender"
End Sub

