Attribute VB_Name = "Module9"

Public Enum FontColours

    vbFontBlack = &H0
    vbFontWhite = &HF9FEFF
    vbFontGreen = &HD0FFCC
    vbFontYellow = &HE1FAFF
    vbFontRed = &HE1E1FF
    vbFontGray = &HC3C3C3
    vbDeepSkyBlue = &HFDA760

End Enum

Function placa_madre() As String

    Dim wmi   As Object

    Dim mos   As Object

    Dim mo    As Object

    Dim Text1 As String
    
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set mos = wmi.ExecQuery("Select * from Win32_Baseboard")
    Text1 = ""

    For Each mo In mos

        Text1 = mo.SerialNumber
        'Text1 = Text1 & "Serial Number: " & mo.SerialNumber & vbCrLf
        'Text1 = Text1 & "Manufacturer: " & mo.Manufacturer & vbCrLf
        'Text1 = Text1 & "Product: " & mo.Product
    Next
    placa_madre = Text1

End Function
