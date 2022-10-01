Attribute VB_Name = "fctRGB"
Function GetRGBColor_Fill(ByVal MyCell As Range) As Variant
'PURPOSE: Output the RGB color code from the specified cell
'REFERENCE: www.TheSpreadsheetGuru.com

    Dim temparray(1 To 3) As Variant
    Dim HEXcolor As String
    Dim RGBcolor As String
    
    HEXcolor = Right("000000" & Hex(MyCell.Interior.Color), 6)
    
    temparray(1) = CInt("&H" & Right(HEXcolor, 2))
    temparray(2) = CInt("&H" & Mid(HEXcolor, 3, 2))
    temparray(3) = CInt("&H" & Left(HEXcolor, 2))
    
    GetRGBColor_Fill = RGB(temparray(1), temparray(2), temparray(3))

End Function

