VERSION 5.00
Begin VB.Form tdatarep 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control Reportes"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox fechaf 
      Height          =   375
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox fechai 
      Height          =   375
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ejecutar"
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label nregistro 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label producto 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Producto"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "tdatarep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    Dim mytablex  As New ADODB.Recordset

    Dim mytabley  As New ADODB.Recordset

    Dim mytablez  As New ADODB.Recordset

    Dim ecantidad As Double

    Dim ecosto    As Double

    Dim ecostot   As Double

    Dim scantidad As Double

    Dim scosto    As Double

    Dim scostot   As Double

    Dim tcantidad As Double

    Dim tcosto    As Double

    Dim tcostot   As Double

    Dim vr

    cn.Execute ("delete from sunat_kardex")
    mytablez.Open "select * from producto where producto='" & producto & "'", cn, adOpenStatic, adLockOptimistic
    mytabley.Open "select * from sunat_kardex", cn, adOpenStatic, adLockOptimistic
    'MsgBox "xx"
    'Exit Sub
    Do

        If mytablez.EOF Then Exit Do
        ecantidad = 0
        ecosto = 0
        ecostot = 0

        scantidad = 0
        scosto = 0
        scostot = 0

        tcantidad = 0
        tcosto = 0
        tcostot = 0
        buf = "select local,tipo,serie,numero,fecha,producto,cantidad,factor,precio,total,acu from detalle where "
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "'"
        buf = buf & " and producto='" & "" & mytablez.Fields("producto") & "'"
        buf = buf & " and (acu='A' or acu='B' or acu='C' or acu='D' or acu='J' or acu='K' or acu='L' or acu='M')"
        'buf = buf & " order by fecha"
        MsgBox buf
        mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
        MsgBox "xxy"
        nregistro = "" & mytablex.RecordCount
        vr = DoEvents()
        'Exit Sub

        Do

            If mytablex.EOF Then Exit Do
            'nregistro = "" & mytablex.RecordCount
            vr = DoEvents()
            mytabley.AddNew
            mytabley.Fields("producto") = "" & mytablex.Fields("producto")
            mytabley.Fields("fecha") = Format("" & mytablex.Fields("fecha"), "dd/mm/yyyy")
            mytabley.Fields("tipo") = "" & mytablex.Fields("tipo")
            mytabley.Fields("serie") = "" & mytablex.Fields("serie")
            mytabley.Fields("numero") = "" & mytablex.Fields("numero")
            mytabley.Fields("tipoope") = "" & mytablex.Fields("tipo")

            If "" & mytablex.Fields("acu") = "A" Or "" & mytablex.Fields("acu") = "B" Or "" & mytablex.Fields("acu") = "C" Or "" & mytablex.Fields("acu") = "D" Then
                mytabley.Fields("ecantidad") = Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
                mytabley.Fields("ecosto") = Val("" & mytablex.Fields("precio"))
                mytabley.Fields("ecostot") = Val("" & mytabley.Fields("ecosto")) * Val("" & mytabley.Fields("ecantidad"))

                mytabley.Fields("scantidad") = 0
                mytabley.Fields("scosto") = 0
                mytabley.Fields("scostot") = 0
                tcantidad = tcantidad - Val("" & mytabley.Fields("ecantidad"))

            End If

            If "" & mytablex.Fields("acu") = "J" Or "" & mytablex.Fields("acu") = "K" Or "" & mytablex.Fields("acu") = "L" Or "" & mytablex.Fields("acu") = "M" Then
                mytabley.Fields("ecantidad") = 0
                mytabley.Fields("ecosto") = 0
                mytabley.Fields("ecostot") = 0

                mytabley.Fields("scantidad") = Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
                mytabley.Fields("scosto") = Val("" & mytablex.Fields("precio"))
                mytabley.Fields("scostot") = Val("" & mytabley.Fields("scosto")) * Val("" & mytabley.Fields("scantidad"))
                tcantidad = tcantidad + Val("" & mytabley.Fields("ecantidad"))

            End If

            mytabley.Fields("tcantidad") = tcantidad
            mytabley.Fields("tcosto") = 0
            mytabley.Fields("tcostot") = 0

            mytabley.Update
            mytablex.MoveNext
        Loop
        mytablex.Close
        mytablez.MoveNext
    Loop
    mytabley.Close
    mytablez.Close

    Set Rst = cn.Execute("SELECT * FROM sunat_kardex group  by producto")

    'Asigna el recordset al reporte
    'tinvsun.Orientation = rptOrientLandscape
    Set tinvsun.DataSource = Rst
    ' Muestra el reporte
    tinvsun.Show vbModal

End Sub

Private Sub Form_Load()
    fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    fechaf = Format(Now, "dd/mm/yyyy")

End Sub
