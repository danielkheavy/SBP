VERSION 5.00
Begin VB.Form trecisc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recalcular exonerado"
   ClientHeight    =   3090
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      Height          =   615
      Left            =   3960
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox fechaf 
      Height          =   495
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox fechai 
      Height          =   495
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fechaf"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fechai"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu dflo23232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "trecisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub proceso_refrescar()
Dim xtotal As Double
Dim xsubtotal As Double
Dim xivap As Double
Dim descuento As Double
Dim xneto As Double
Dim ximpuesto As Double
Dim xisc As Double
Dim vr
Dim mytablex As Table
Dim mytabley As Table
Dim mytablez As Table
Dim xgravado As Double

xgravado = 0
xtotal = 0
xsubtotal = 0
xivap = 0
xdescuento = 0
xneto = 0
ximpuesto = 0
xisc = 0

Set mytablez = mydbxglo.OpenTable("producto")
mytablez.Index = "producto"



Set mytablex = mydbxglo.OpenTable("factura")
mytablex.Index = "fecha"

Set mytabley = mydbxglo.OpenTable("detalle")
mytabley.Index = "tdetalle"

mytablex.Seek ">=", fechai

Do
If mytablex.EOF Then Exit Do
vr = DoEvents()
Label3 = "" & mytablex.Fields("fecha")
xgravado = 0
xtotal = 0
xsubtotal = 0
xivap = 0
xdescuento = 0
xneto = 0
ximpuesto = 0
xisc = 0
mytabley.Seek "=", "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo"), "" & mytablex.Fields("serie"), "" & mytablex.Fields("numero")
If Not mytabley.NoMatch Then
   Do
   If mytabley.EOF Then Exit Do
   If "" & mytablex.Fields("local") = "" & mytabley.Fields("local") And "" & mytablex.Fields("tipo") = "" & mytabley.Fields("tipo") And "" & mytablex.Fields("serie") = "" & mytabley.Fields("serie") And "" & mytablex.Fields("numero") = "" & mytabley.Fields("numero") Then
      '----------------------
      mytabley.Edit
      busca_producto mytablez, mytabley
      calcula_igv mytabley
      ximpuesto = ximpuesto + Val("" & mytabley.Fields("impuesto"))
      xtotal = xtotal + Val("" & mytabley.Fields("total"))
      xsubtotal = xsubtotal + Val("" & mytabley.Fields("subtotal"))
      xivap = xivap + Val("" & mytabley.Fields("tivap"))
      xdescuento = xdescuento + Val("" & mytabley.Fields("descuento"))
      xneto = xneto + Val("" & mytabley.Fields("neto"))
      xisc = xisc + Val("" & mytabley.Fields("tisc"))
      If Val("" & mytabley.Fields("igv")) = 0 Then
         xgravado = xgravado + Val("" & mytabley.Fields("total"))
      End If

      mytabley.Update
      '----------------------
      Else
      Exit Do
   End If
   mytabley.MoveNext
   Loop
   mytablex.Edit
   mytablex.Fields("neto") = xneto
   mytablex.Fields("descuento") = xdescuento
   mytablex.Fields("subtotal") = xsubtotal
   mytablex.Fields("impuesto") = ximpuesto
   mytablex.Fields("tisc") = xisc
   mytablex.Fields("gravado") = xgravado
   mytablex.Fields("tivap") = xivap
   mytablex.Update
End If
mytablex.MoveNext
Loop
mytablex.Close
mytabley.Close
MsgBox "fin"
End Sub
Sub calcula_igv(mytabley As Table)
Dim sdx As Double
Dim sdx1 As Double
Dim sdx2 As Double
Dim tdscto As Double
Dim tdscto1 As Double
Dim found As Integer
Dim xtivap As Double
Dim xtisc As Double
xtivap = Val("" & mytabley.Fields("total")) * Val("" & mytabley.Fields("ivap")) / 100
mytabley.Fields("tivap") = xtivap
tdscto = Val("" & mytabley.Fields("total")) * Val("" & mytabley.Fields("deslipo")) / 100
mytabley.Fields("descuento") = tdscto
mytabley.Fields("total") = Val("" & mytabley.Fields("total")) - Val("" & mytabley.Fields("descuento"))
mytabley.Fields("subtotal") = Val("" & mytabley.Fields("total"))
mytabley.Fields("impuesto") = 0
mytabley.Fields("neto") = Val("" & mytabley.Fields("subtotal")) + Val("" & mytabley.Fields("descuento"))
sdx2 = 1 + Val("" & mytabley.Fields("igv")) / 100
sdx1 = Val("" & mytabley.Fields("total")) / sdx2
mytabley.Fields("subtotal") = sdx1
sdx = Val("" & mytabley.Fields("total")) - Val("" & mytabley.Fields("subtotal"))
mytabley.Fields("impuesto") = sdx
'MsgBox sdx
'End
mytabley.Fields("descuento") = tdscto
mytabley.Fields("neto") = Val("" & mytabley.Fields("subtotal")) + Val("" & mytabley.Fields("descuento"))
xtisc = Val("" & mytabley.Fields("subtotal")) * Val("" & mytabley.Fields("isc")) / 100
mytabley.Fields("tisc") = xtisc


End Sub




Private Sub Command1_Click()
proceso_refrescar
End Sub

Private Sub Form_Load()
fechai = Format(Now, "dd/mm/yyyy")
fechaf = Format(Now, "dd/mm/yyyy")
End Sub
Sub busca_producto(mytablez As Table, mytabley As Table)
mytablez.Seek "=", "" & mytabley.Fields("producto")
If Not mytablez.NoMatch Then
   mytabley.Fields("ivap") = Val("" & mytablez.Fields("ivap"))
   mytabley.Fields("isc") = Val("" & mytablez.Fields("isc"))
   mytabley.Fields("igv") = Val("" & mytablez.Fields("igv"))
   
End If
End Sub
