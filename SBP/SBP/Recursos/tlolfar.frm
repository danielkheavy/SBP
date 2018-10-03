VERSION 5.00
Begin VB.Form tlolfar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recogiendo Datos de Lolfar"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   8535
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Menu lo444 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tlolfar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    lolfar_productos

    'lolfar_familias
    'lolfar_marcas
End Sub

Sub lolfar_productos()

    Dim mytablex As Table

    Dim mytabley As New ADODB.Recordset

    Dim mydbx    As Database

    Dim vr

    sum1 = 1
    cn.Execute ("delete from producto")
    cn.Execute ("delete from precios")
    cn.Execute ("delete from marca")
   
    MsgBox "eNTER"
   
    Set mydbx = OpenDatabase("d:\lolfar", False, False, "foxpro 2.5;")
    Set mytablex = mydbx.OpenTable("surf10")

    Do
        vr = DoEvents()
        Command1.Caption = "" & sum1

        If mytablex.EOF Then Exit Do
        '------------------------------
        mytabley.Open "select * from producto where producto='" & Trim("" & mytablex.Fields("codpro")) & "'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount = 0 Then
            mytabley.AddNew
            graba_lolfar mytablex, mytabley
            mytabley.Update

        End If

        mytabley.Close
        '------------------------------
        sum1 = sum1 + 1
        mytablex.MoveNext
    Loop
    mytablex.Close
    mydbx.Close
   
    MsgBox "Marca"

    Set mydbx = OpenDatabase("d:\lolfar", False, False, "foxpro 2.5;")
    Set mytablex = mydbx.OpenTable("surf15")

    Do
        vr = DoEvents()
   
        If mytablex.EOF Then Exit Do
        '------------------------------
        mytabley.Open "select * from marca where marca='" & Trim("" & mytablex.Fields("codlab")) & "'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount = 0 Then
            mytabley.AddNew
            mytabley.Fields("marca") = Trim("" & mytablex.Fields("codlab"))
            mytabley.Fields("descripcio") = Trim("" & mytablex.Fields("deslab"))
            mytabley.Update

        End If

        mytabley.Close
        '------------------------------
        sum1 = sum1 + 1
        mytablex.MoveNext
    Loop
    mytablex.Close
    mydbx.Close
    MsgBox "Fin"

End Sub

Sub graba_lolfar(mytablex As Table, mytabley As ADODB.Recordset)

    Dim sdx1     As Double

    Dim buf      As String

    Dim mytablea As New ADODB.Recordset

    mytabley.Fields("producto") = Trim("" & mytablex.Fields("codpro"))
    mytabley.Fields("barras") = ""
    mytabley.Fields("descripcio") = Trim("" & mytablex.Fields("despro"))
    mytabley.Fields("descorto") = Mid$(Trim("" & mytablex.Fields("despro")), 1, 20)

    mytabley.Fields("presenta") = "" '& mytabley.Fields("presentaci")

    mytabley.Fields("familia") = Trim("" & mytablex.Fields("codtip"))
    mytabley.Fields("subfamilia") = ""
    mytabley.Fields("seccion") = ""
    mytabley.Fields("marca") = "" & mytablex.Fields("codlab")
    mytabley.Fields("categoria") = ""
    mytabley.Fields("linea") = ""
    mytabley.Fields("color") = ""
    mytabley.Fields("fabrica") = ""
    mytabley.Fields("serie") = ""
    mytabley.Fields("peso") = ""
    mytabley.Fields("servicio") = ""
    mytabley.Fields("vecaja") = "S"
    mytabley.Fields("igv") = 18
    mytabley.Fields("isc") = 0
    mytabley.Fields("pesokgr") = 0.001
    mytabley.Fields("comision") = 0
    mytabley.Fields("monedac") = "S"
    mytabley.Fields("monedav") = "S"

    If Val("" & mytablex.Fields("stkfra")) <= 1 Then
        mytabley.Fields("unidad") = "PZA"
        mytabley.Fields("factor") = "1"
        mytabley.Fields("costou") = Val("" & mytablex.Fields("costod"))
        mytabley.Fields("costop") = Val("" & mytablex.Fields("costod"))
        mytabley.Fields("costoini") = Val("" & mytablex.Fields("costod"))
        mytabley.Fields("cospaqu") = Val("" & mytablex.Fields("costod"))
        mytabley.Fields("cospaqp") = Val("" & mytablex.Fields("costod"))
        mytabley.Fields("cospaqi") = Val("" & mytablex.Fields("costod"))

    End If

    If Val("" & mytablex.Fields("stkfra")) > 1 Then
        mytabley.Fields("unidad") = "CAJA"
        mytabley.Fields("factor") = "1"
          
        mytabley.Fields("unidad") = "CAJ"
        mytabley.Fields("factor") = Val("" & mytablex.Fields("stkfra"))
          
        sdx1 = Val("" & mytablex.Fields("costod")) / Val("" & mytablex.Fields("stkfra"))
        sdx1 = Format(sdx1, "0.000")
        mytabley.Fields("costou") = sdx1
        mytabley.Fields("costop") = sdx1
        mytabley.Fields("costoini") = sdx1
        mytabley.Fields("cospaqu") = sdx1
        mytabley.Fields("cospaqp") = sdx1
        mytabley.Fields("cospaqi") = sdx1

    End If

    'GRABANDO PRECIOS
    mytablea.Open "select * from precios where producto='" & "" & Trim("" & mytablex.Fields("codpro")) & "' and local='01'", cn, adOpenStatic, adLockOptimistic

    If mytablea.RecordCount = 0 Then
        If Val("" & mytablex.Fields("stkfra")) <= 1 Then
            mytablea.AddNew
            mytablea.Fields("local") = "01"
            mytablea.Fields("producto") = Trim("" & mytablex.Fields("codpro"))
            mytablea.Fields("factor1") = 1 'Val("" & mytabley.Fields("factor1"))
            mytablea.Fields("unidad1") = "UND" '& mytabley.Fields("unidad1")
            mytablea.Fields("pventa1") = Val("" & mytablex.Fields("prisal"))
            mytablea.Update

        End If

        If Val("" & mytablex.Fields("stkfra")) > 1 Then
            mytablea.AddNew
            mytablea.Fields("local") = "01"
            mytablea.Fields("producto") = Trim("" & mytablex.Fields("codpro"))
            mytablea.Fields("factor1") = Val("" & mytablex.Fields("stkfra"))
            mytablea.Fields("unidad1") = "CAJA" '& mytabley.Fields("unidad1")
            mytablea.Fields("pventa1") = Val("" & mytablex.Fields("prisal"))
   
            sdx1 = Val("" & mytablex.Fields("prisal")) / Val("" & mytablex.Fields("stkfra"))
            sdx1 = Format(sdx1, "0.00")
            mytablea.Fields("pventa2") = sdx1
            mytablea.Fields("unidad2") = "PZA"
            mytablea.Fields("factor2") = 1
   
            mytablea.Update

        End If

    End If

    mytablea.Close
       
End Sub

