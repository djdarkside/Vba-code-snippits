Attribute VB_Name = "Calculate"
Option Compare Database

Function CalculatePrice(Package As Variant, COM As String) As Currency
    Dim TOTAL As Currency
    Dim currentPckPrice As Currency
    Dim pkglength, X As Integer
    Dim PCK As String
    
    TOTAL = 0
    pkglength = 0
        X = 1
        pkglength = Int(Len(Package))
        'MsgBox (pkglength)
        While X <= pkglength
            PCK = Right(Left(Package, X), 1)
            'ignore space in package string
            'MsgBox (Pck)
            If (PCK <> " ") Then
                    currentPckPrice = IIf(IsNull(PCK), 0#, DLookup("[" & PCK & "]", "PriceList", "[PriceCode]='" & COM & "'"))
                    TOTAL = TOTAL + currentPckPrice
                    'MsgBox (Pck) + " calculated"
            End If
            X = X + 1
        Wend
            
    CalculatePrice = TOTAL

End Function

Function CalculateKeepsake(Package As Variant, COM As String) As Currency
    Dim TOTAL As Currency
    Dim currentPckPrice As Currency
    Dim pkglength, X As Integer
    Dim PCK As String
    
    Dim dbs As Database
    Dim Keepsakes As DAO.Recordset
    
    Set dbs = CurrentDb
    Set Keepsakes = dbs.OpenRecordset("kEEPSAKEProducts", dbOpenDynaset)
    
    TOTAL = 0
    pkglength = 0
        X = 1
        pkglength = Int(Len(Package))
        'MsgBox (pkglength)
        While X <= pkglength
            PCK = Right(Left(Package, X), 1)
            'ignore space in package string
            'MsgBox (Pck)
            If (PCK <> " " And Not IsNull(PCK)) Then
                Keepsakes.FindFirst "[CODES] = '" & PCK & "'"
                'if valid keepsake found in table copy file
                If Not (Keepsakes.NoMatch) Then
                    currentPckPrice = IIf(IsNull(PCK), 0#, DLookup("[" & PCK & "]", "PriceList", "[PriceCode]='" & COM & "'"))
                    TOTAL = TOTAL + currentPckPrice
                End If
            End If
            X = X + 1
        Wend
            
    CalculateKeepsake = TOTAL

End Function

Function Calculate_PCL_Price(Package As Variant, COM As String) As Currency
    Dim TOTAL As Currency
    Dim currentPCLPrice As Currency
    Dim pkglength, X As Integer
    Dim PCK As String
    
    TOTAL = 0
    pkglength = 0
        X = 1
        pkglength = Int(Len(Package))
        'MsgBox (pkglength)
        While X <= pkglength
            PCK = Right(Left(Package, X), 1)
            'ignore space in package string
            'MsgBox (Pck)
            If (PCK <> " ") Then
                    currentPCLPrice = IIf(IsNull(PCK), 0#, DLookup("[" & PCK & "]", "PriceList", "[PriceCode]='PCLCOST'"))
                    TOTAL = TOTAL + currentPCLPrice
                    'MsgBox (Pck) + " calculated"
            End If
            X = X + 1
        Wend
            
    Calculate_PCL_Price = TOTAL

End Function

Function Calculate_PCL_Keepsake(Package As Variant, COM As String) As Currency
    Dim TOTAL As Currency
    Dim currentPckPrice As Currency
    Dim pkglength, X As Integer
    Dim PCK As String
    
    Dim dbs As Database
    Dim Keepsakes As DAO.Recordset
    
    Set dbs = CurrentDb
    Set Keepsakes = dbs.OpenRecordset("kEEPSAKEProducts", dbOpenDynaset)
    
    TOTAL = 0
    pkglength = 0
        X = 1
        pkglength = Int(Len(Package))
        'MsgBox (pkglength)
        While X <= pkglength
            PCK = Right(Left(Package, X), 1)
            'ignore space in package string
            'MsgBox (Pck)
            If (PCK <> " " And Not IsNull(PCK)) Then
                Keepsakes.FindFirst "[CODES] = '" & PCK & "'"
                'if valid keepsake found in table copy file
                If Not (Keepsakes.NoMatch) Then
                    currentPckPrice = IIf(IsNull(PCK), 0#, DLookup("[" & PCK & "]", "PriceList", "[PriceCode]='PCLCOST'"))
                    TOTAL = TOTAL + currentPckPrice
                End If
            End If
            X = X + 1
        Wend
            
    Calculate_PCL_Keepsake = TOTAL

End Function

Function Calculate_PORTRAIT_PRICE(Package As Variant, COM As String) As Currency
    Dim TOTAL As Currency
    Dim currentPckPrice As Currency
    Dim pkglength, X As Integer
    Dim PCK As String
    
    Dim dbs As Database
    Dim Keepsakes As DAO.Recordset
    
    Set dbs = CurrentDb
    Set Keepsakes = dbs.OpenRecordset("PortraitProducts", dbOpenDynaset)
    
    TOTAL = 0
    pkglength = 0
        X = 1
        pkglength = Int(Len(Package))
        'MsgBox (pkglength)
        While X <= pkglength
            PCK = Right(Left(Package, X), 1)
            'ignore space in package string
            'MsgBox (Pck)
            If (PCK <> " " And Not IsNull(PCK)) Then
                Keepsakes.FindFirst "[CODES] = '" & PCK & "'"
                'if valid keepsake found in table copy file
                If Not (Keepsakes.NoMatch) Then
                    currentPckPrice = IIf(IsNull(PCK), 0#, DLookup("[" & PCK & "]", "PriceList", "[PriceCode]='" & COM & "'"))
                    TOTAL = TOTAL + currentPckPrice
                End If
            End If
            X = X + 1
        Wend
            
    Calculate_PORTRAIT_PRICE = TOTAL

End Function

