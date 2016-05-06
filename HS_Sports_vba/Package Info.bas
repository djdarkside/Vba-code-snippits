Attribute VB_Name = "Package Info"
Option Compare Database

Function PackageExtension(Package As Variant) As String
    Dim pkglength, X As Integer, Keepsake As String
    Dim PCK As String
    Dim dbs As Database
    Dim Keepsakes As DAO.Recordset
    
    Set dbs = CurrentDb
    Set Keepsakes = dbs.OpenRecordset("PackageProducts", dbOpenDynaset)
    
    pkglength = 0
    X = 1
    pkglength = Int(Len(Package))
    While X <= pkglength
        PCK = Right(Left(Package, X), 1)
        'ignore space in package string
        If (PCK <> " ") Then
            Keepsakes.FindFirst "[CODES] = '" & PCK & "'"
            If Not (Keepsakes.NoMatch) Then
                Keepsake = Trim(Keepsake) + Trim(PCK) + Trim("-1;")
            End If
        End If
        X = X + 1
    Wend
    PackageExtension = Keepsake
    
End Function

Function PACKAGECOUNT(Package As Variant, Pkgcounted As String) As Integer
    Dim count As Integer
    Dim pkglength, X As Integer
    Dim packageOrdered As String
    
    count = 0
    pkglength = 0
    X = 1
    pkglength = Int(Len(Trim(Package)))
    While X <= pkglength
        packageOrdered = Right(Left(Trim(Package), X), 1)
        If packageOrdered = Pkgcounted Then
           count = count + 1
        End If
        X = X + 1
    Wend
    PACKAGECOUNT = count
End Function

Function ComboTeam(D, E, F, G As Double) As String
    Dim dbs As Database
    Set dbs = CurrentDb
    Dim TeamPck As DAO.Recordset
    Set TeamPck = dbs.OpenRecordset("01a - TeamPackage Counts Totals", dbOpenDynaset)
    Dim out1, out2, out3, out4 As String
    If TeamPck![D] > 0 Then
        out1 = "D-" & TeamPck![D]
    Else
        out1 = ""
    End If
    If TeamPck![E] > 0 Then
        out2 = "E-" & TeamPck![E]
    Else
        out2 = ""
    End If
    If TeamPck![F] > 0 Then
        out3 = "F-" & TeamPck![F]
    Else
        out3 = ""
    End If
    If TeamPck![G] > 0 Then
        out4 = "G-" & TeamPck![G]
    Else
        out4 = ""
    End If
    PckTotal = Trim(PckTotal) + out1 & ";" & out2 & ";" & out3 & ";" & out4 & ";"
    ComboTeam = PckTotal
End Function

Function SheetCountFull(Package As Variant) As Integer

    Dim currentCount As Integer
    Dim pkglength, X As Integer
    Dim PCK As String
    Dim count As Integer
    Dim dbs As Database
    Dim SheetList As DAO.Recordset
    
    Set dbs = CurrentDb
    Set SheetList = dbs.OpenRecordset("PackageSheets", dbOpenDynaset)
    
    count = 0
    pkglength = 0
    X = 1
    pkglength = Int(Len(Package))
    
        While X <= pkglength
            PCK = Right(Left(Package, X), 1)

            If (PCK <> " " And Not IsNull(PCK)) Then
                SheetList.FindFirst "[Package] = '" & PCK & "'"

                If Not (SheetList.NoMatch) Then
                    currentCount = IIf(IsNull(PCK), 0#, DLookup("[sheetCount]", "PackageSheets", "[Package]='" & PCK & "'"))
                    count = count + currentCount
                End If
            End If
            X = X + 1
        Wend
            
    SheetCountFull = count

End Function

Function SheetCountHalf(Package As Variant) As Integer

    Dim currentCount As Integer
    Dim pkglength, X As Integer
    Dim PCK As String
    Dim count As Integer
    Dim dbs As Database
    Dim SheetList As DAO.Recordset
    
    Set dbs = CurrentDb
    Set SheetList = dbs.OpenRecordset("PackageHalfSheets", dbOpenDynaset)
    
    count = 0
    pkglength = 0
    X = 1
    pkglength = Int(Len(Package))
    
        While X <= pkglength
            PCK = Right(Left(Package, X), 1)

            If (PCK <> " " And Not IsNull(PCK)) Then
                SheetList.FindFirst "[Package] = '" & PCK & "'"

                If Not (SheetList.NoMatch) Then
                    currentCount = IIf(IsNull(PCK), 0#, DLookup("[sheetCount]", "PackageHalfSheets", "[Package]='" & PCK & "'"))
                    count = count + currentCount
                End If
            End If
            X = X + 1
        Wend
            
    SheetCountHalf = count

End Function
