Attribute VB_Name = "Module23"
Option Explicit

'diccionario de datos
Global dicigv    As String

Global dicmoneda As String

Global dicruc    As String

Sub dicargentina()
    dicigv = "IVA"
    dicmoneda = "$/."
    dicruc = "CUIT"

End Sub

Sub dicchile()
    dicigv = "IVA"
    dicmoneda = "$/."
    dicruc = "CUIT"

End Sub

Sub dicperu()
    dicigv = "IGV"
    dicmoneda = "S/"
    dicruc = "RUC"

End Sub

Function VerificarCUIT(CUIT As String) As Integer

    Dim m, n, X, a As Integer

    a = 0

    For X = 1 To 11
        n = Val(Mid(CUIT, X, 1))
        n = n + 48

        Select Case X

            Case 1: m = 5

            Case 2: m = 4

            Case 3: m = 3

            Case 4: m = 2

            Case 5: m = 7

            Case 6: m = 6

            Case 7: m = 5

            Case 8: m = 4

            Case 9: m = 3

            Case 10: m = 2

            Case 11: m = 1

        End Select

        a = a + n * m
    Next
    a = a Mod 11

    If a = 3 Then
        VerificarCUIT = 1
    Else
        VerificarCUIT = 0

    End If

    'MsgBox VerificarCUIT
End Function

