VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{7187619F-D732-11D2-8A16-00000E84DA63}#1.0#0"; "ezAVI26.ocx"
Begin VB.Form vidconcar 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema Auditoria Camaras"
   ClientHeight    =   7860
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   14835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   14835
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox sentido 
      Height          =   375
      Left            =   12360
      MaxLength       =   1
      TabIndex        =   64
      Text            =   "*"
      Top             =   600
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Caption         =   "Visor de Video"
      Height          =   6855
      Left            =   1800
      TabIndex        =   61
      Top             =   600
      Visible         =   0   'False
      Width           =   11895
      Begin VB.CommandButton Command5 
         Caption         =   "Cerrar Ventana"
         Height          =   495
         Left            =   9840
         TabIndex        =   63
         Top             =   6120
         Width           =   1575
      End
      Begin AVIPlay.ezAVIWnd ezAVIWnd1 
         CausesValidation=   0   'False
         Height          =   5625
         Left            =   0
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   360
         Width           =   11520
         _ExtentX        =   20320
         _ExtentY        =   9922
         Filename        =   ""
         AutoSize        =   0   'False
         Volume          =   100
      End
   End
   Begin VB.TextBox serie 
      Height          =   375
      Left            =   4920
      MaxLength       =   4
      TabIndex        =   60
      Text            =   "*"
      Top             =   240
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Preparacion archivo Sunat"
      Height          =   4095
      Left            =   3120
      TabIndex        =   31
      Top             =   1800
      Visible         =   0   'False
      Width           =   8535
      Begin VB.TextBox ruc 
         Height          =   375
         Left            =   1320
         MaxLength       =   11
         TabIndex        =   46
         Text            =   "20420605006"
         Top             =   2760
         Width           =   2775
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cerrar ventana"
         Height          =   615
         Left            =   5880
         TabIndex        =   45
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Generar Archivo"
         Height          =   615
         Left            =   5880
         TabIndex        =   44
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox XX 
         Height          =   375
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   43
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox DD 
         Height          =   375
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   41
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox MM 
         Height          =   375
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   39
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox AAAA 
         Height          =   375
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   37
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox BB 
         Height          =   375
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   35
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox CDR 
         Height          =   375
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   33
         Text            =   "CDR"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RUC"
         Height          =   375
         Left            =   240
         TabIndex        =   47
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "XX"
         Height          =   375
         Left            =   240
         TabIndex        =   42
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DD"
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MM"
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AAAA"
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BB"
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CDR"
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox detraccion 
      Height          =   375
      Left            =   12360
      MaxLength       =   1
      TabIndex        =   30
      Text            =   "N"
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox central 
      Height          =   375
      Left            =   10560
      MaxLength       =   1
      TabIndex        =   26
      Text            =   "S"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox moneda 
      Height          =   375
      Left            =   10560
      MaxLength       =   1
      TabIndex        =   24
      Text            =   "*"
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ver Video"
      Enabled         =   0   'False
      Height          =   495
      Left            =   13080
      TabIndex        =   23
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   14520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "vidconcar.frx":0000
      Height          =   1815
      Left            =   120
      OleObjectBlob   =   "vidconcar.frx":0014
      TabIndex        =   22
      Top             =   6000
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Seleccionando.."
      Height          =   495
      Left            =   13320
      TabIndex        =   21
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox estado 
      Height          =   375
      Left            =   10560
      MaxLength       =   1
      TabIndex        =   20
      Text            =   "*"
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox cajero 
      Height          =   375
      Left            =   8400
      MaxLength       =   11
      TabIndex        =   18
      Text            =   "*"
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox turno 
      Height          =   375
      Left            =   8400
      MaxLength       =   1
      TabIndex        =   16
      Text            =   "*"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox caja 
      Height          =   375
      Left            =   8400
      MaxLength       =   2
      TabIndex        =   14
      Text            =   "*"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox numero 
      Height          =   375
      Left            =   5640
      MaxLength       =   11
      TabIndex        =   12
      Text            =   "*"
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox nombre 
      Height          =   375
      Left            =   4920
      MaxLength       =   60
      TabIndex        =   10
      Text            =   "*"
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox codigo 
      Height          =   375
      Left            =   4920
      MaxLength       =   11
      TabIndex        =   8
      Text            =   "*"
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox tipo 
      Height          =   375
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   6
      Text            =   "*"
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox fechaf 
      Height          =   375
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox fechai 
      Height          =   375
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "vidconcar.frx":1233
      Height          =   4455
      Left            =   120
      OleObjectBlob   =   "vidconcar.frx":1247
      TabIndex        =   0
      Top             =   1440
      Width           =   14535
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   14400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label26 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sentido"
      Height          =   375
      Left            =   11040
      TabIndex        =   65
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label ntpeaje 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8280
      TabIndex        =   59
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label ntdetraccion 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8280
      TabIndex        =   58
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label ntcobrado 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8280
      TabIndex        =   57
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label ntnocobrado 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8280
      TabIndex        =   56
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label tnocobrado 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6960
      TabIndex        =   55
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label tcobrado 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6960
      TabIndex        =   54
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label tdetraccion 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6960
      TabIndex        =   53
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label tpeaje 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6960
      TabIndex        =   52
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DetNocobrado"
      Height          =   375
      Left            =   5640
      TabIndex        =   51
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DetCobrado"
      Height          =   375
      Left            =   5640
      TabIndex        =   50
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TotalDocumento"
      Height          =   375
      Left            =   5640
      TabIndex        =   49
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TotalPeaje"
      Height          =   375
      Left            =   5640
      TabIndex        =   48
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SoloDetracciones"
      Height          =   375
      Left            =   11040
      TabIndex        =   29
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "numero"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   13440
      TabIndex        =   28
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Centralizado"
      Height          =   375
      Left            =   9600
      TabIndex        =   27
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Moneda"
      Height          =   375
      Left            =   9600
      TabIndex        =   25
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado 2,0,1"
      Height          =   375
      Left            =   9600
      TabIndex        =   19
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cajero"
      Height          =   375
      Left            =   7080
      TabIndex        =   17
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Turno"
      Height          =   375
      Left            =   7080
      TabIndex        =   15
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Caja"
      Height          =   375
      Left            =   7080
      TabIndex        =   13
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serie Numero"
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TipoDocumento"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Menu dj8222 
      Caption         =   "&menu"
      Begin VB.Menu dk8822 
         Caption         =   "&1.Exportar Excell"
      End
      Begin VB.Menu dsumay 
         Caption         =   "&2.Exportar para Sunat"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu flo33 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "vidconcar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cadiario As String
Dim dediario As String
Dim fpdiario As String
Dim globalvid As String

Private Sub AAAA_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
MM.SetFocus

End Sub

Private Sub BB_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
AAAA.SetFocus

End Sub

Private Sub CDR_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
BB.SetFocus
End Sub

Private Sub Command1_Click()
Dim buf As String
Dim xtpeaje As Double
Dim xtdetraccion As Double
Dim xtcobrado As Double
Dim xtnocobrado As Double

Dim xntpeaje As Double
Dim xntdetraccion As Double
Dim xntcobrado As Double
Dim xntnocobrado As Double


tpeaje = ""
tdetraccion = ""
tcobrado = ""
tnocobrado = ""
xtpeaje = 0
xtdetraccion = 0
xtcobrado = 0
xtnocobrado = 0

ntpeaje = ""
ntdetraccion = ""
ntcobrado = ""
ntnocobrado = ""
xntpeaje = 0
xntdetraccion = 0
xntcobrado = 0
xntnocobrado = 0


cadiario = "factura"
dediario = "detalle"
globalvid = globaldat & "\video"
If central <> "S" Then
   cadiario = "cadiario"
   dediario = "dediario"
   globalvid = globaldat & "\cavideo"
End If
buf = "select * from " & cadiario & " where "
buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
If detraccion = "S" Then
   buf = buf & " and Tdetra>0 "
End If
If tipo <> "*" Then
   buf = buf & " and tipo like '" & tipo & "'"
End If
If caja <> "*" Then
   buf = buf & " and caja like '" & caja & "'"
End If
If serie <> "*" Then
   buf = buf & " and serie like '" & serie & "'"
End If
If sentido <> "*" Then
   buf = buf & " and sentido like '" & sentido & "'"
End If


If numero <> "*" Then
   buf = buf & " and numero like '" & numero & "'"
End If
If codigo <> "*" Then
   buf = buf & " and codigo like '" & codigo & "'"
End If
If nombre <> "*" Then
   buf = buf & " and nombreb like '" & nombre & "'"
End If
If moneda <> "*" Then
   buf = buf & " and moneda like '" & moneda & "'"
End If
If estado <> "*" Then
   buf = buf & " and estado like '" & estado & "'"
End If
If cajero <> "*" Then
buf = buf & " and usuario like '" & cajero & "'"
End If


               Data1.Connect = "foxpro 2.5;"
               Data1.DatabaseName = globaldat
               Data1.RecordSource = buf
               Data1.Refresh
               Do
               If Data1.Recordset.EOF Then Exit Do
               If "" & Data1.Recordset.Fields("estado") = "2" Then
                  xntpeaje = xntpeaje + 1
                  xtpeaje = xtpeaje + Val("" & Data1.Recordset.Fields("xneto"))
                  xntdetraccion = xntdetraccion + 1
                  xtdetraccion = xtdetraccion + Val("" & Data1.Recordset.Fields("total"))
                  If Val("" & Data1.Recordset.Fields("tdetra")) > 0 Then
                     
                     If Val("" & Data1.Recordset.Fields("xneto")) = Val("" & Data1.Recordset.Fields("total")) Then
                        xtnocobrado = xtnocobrado + Val("" & Data1.Recordset.Fields("tdetra"))
                        xntnocobrado = xntnocobrado + 1
                     Else
                        xtcobrado = xtcobrado + Val("" & Data1.Recordset.Fields("tdetra"))
                        xntcobrado = xntcobrado + 1
                     End If
                   End If
               End If
               Data1.Recordset.MoveNext
               Loop
 tpeaje = Format(xtpeaje, "0.00")
tdetraccion = Format(xtdetraccion, "0.00")
tcobrado = Format(xtcobrado, "0.00")
tnocobrado = Format(xtnocobrado, "0.00")

ntpeaje = Format(xntpeaje, "0")
ntdetraccion = Format(xntdetraccion, "0")
ntcobrado = Format(xntcobrado, "0")
ntnocobrado = Format(xntnocobrado, "0")

               

End Sub

Private Sub Command2_Click()
Dim buf As String
On Error GoTo cmd2_err
Frame2.Visible = True
buf = globaldat & "\video\" & Data1.Recordset.Fields("serie") & "-" & Data1.Recordset.Fields("numero")
Frame2.Caption = buf
If central <> "S" Then
   buf = globaldat & "\cavideo\" & Data1.Recordset.Fields("serie") & "-" & Data1.Recordset.Fields("numero")
End If
ezAVIWnd1.Filename = buf
ezAVIWnd1.Play
Exit Sub
cmd2_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Frame2.Visible = False
Frame2.Caption = "Visor de Video"
Exit Sub
End Sub

Private Sub Command3_Click()
Dim buf As String
Dim buf1 As String
Dim found As Integer
Dim xplaca As String
Dim xeje As String
Dim mytablex As Table
Dim mydbx As Database

Dim i As Integer
On Error GoTo cmd4322_err
    If CDR <> "CDR" Then
       CDR.SetFocus
       Exit Sub
    End If
    If Len(BB) = 0 Then
       BB.SetFocus
       Exit Sub
    End If
    If Len(AAAA) <> 4 Then
       AAAA.SetFocus
       Exit Sub
    End If
    If Not IsNumeric(AAAA) Then
      AAAA.SetFocus
    End If
    If Len(DD) <> 2 Then
       DD.SetFocus
       Exit Sub
    End If
    If Not IsNumeric(DD) Then
      DD.SetFocus
      Exit Sub
    End If
    If Val(DD) < 1 And Val(DD) > 31 Then
       DD.SetFocus
       Exit Sub
    End If
    If Not IsNumeric(XX) Then
       XX.SetFocus
       Exit Sub
    End If
    Data1.Refresh
    
Set mydbx = OpenDatabase(dediario, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("detalle")
mytablex.Index = "tdetalle"

    buf = CDR + BB + AAAA + MM + DD + XX
    
    Open globaldir & "\sunat\" & buf & ".txt" For Append As #1
    
   'cabecera----------------------------------------
   buf1 = ruc & ","
   buf1 = buf1 & Format(Val(ntdetraccion), "000000") & ","
   buf1 = buf1 & Format(Val(ntcobrado), "000000") & ","
   buf1 = buf1 & convierte_decimal(Val(tdetraccion)) & ","
   buf1 = buf1 & convierte_decimal(Val(tcobrado))
   Print #1, buf1
   buf1 = "" & Data1.Recordset.Fields("codigo") & ","

Do
If Data1.Recordset.EOF Then Exit Do
 mytablex.Seek "=", "" & Data1.Recordset.Fields("local"), "" & Data1.Recordset.Fields("tipo"), "" & Data1.Recordset.Fields("serie"), "" & Data1.Recordset.Fields("numero")
 If Not mytablex.NoMatch Then
    xplaca = "" & mytablex.Fields("placa")
    xeje = "" & mytablex.Fields("subfamilia")
 End If


'detalle
   
   buf1 = buf1 + xplaca + ","
   buf1 = buf1 + xeje + ","
   buf1 = buf1 + "0,"
   buf1 = buf1 + "000000000000000,"
   buf1 = buf1 + "2042060506,"
   buf1 = buf1 + "000,"
   buf1 = buf1 + "000,"
   buf1 = buf1 + "0,"
   buf1 = buf1 + "00000000,"
   buf1 = buf1 + "00000000,"
   buf1 = buf1 & "000000000000000,"
   buf1 = buf1 & "000000000000000,"
   buf1 = buf1 & "000000000000000"
   Print #1, buf1
 Data1.Recordset.MoveNext
Loop
    Close #1
    MsgBox "proceso Terminado"
    Exit Sub
cmd4322_err:
    MsgBox "Error " & error$, 24, "Aviso"
    Close #1
    Exit Sub

End Sub


Private Sub Command4_Click()
Frame1.Visible = False
End Sub

Private Sub Command5_Click()
Frame2.Visible = False
Frame2.Caption = "Visor de Video"
End Sub

Private Sub DBGrid1_DblClick()
visualiza_detalle
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
visualiza_detalle
End Sub

Private Sub DD_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
XX.SetFocus
End Sub

Private Sub dk8822_Click()
excel_paso
End Sub

Private Sub dsumay_Click()
Frame1.Visible = True
CDR = "CDR"
BB = "XX"
AAAA = Format(Year(Now), "YYYY")
MM = Format(Month(Now), "MM")
DD = Format(Day(Now), "DD")
XX = "01"
End Sub

Private Sub flo33_Click()
vidconcar.Hide
Unload vidconcar
End Sub

Sub visualiza_detalle()
Dim buf As String
Dim ufile As String
On Error GoTo cmd1_err
ufile = globaldir & "\video\" & Data1.Recordset.Fields("serie") & "-" & Data1.Recordset.Fields("numero")
If Dir(ufile) = "" Then 'si no existe
     Command2.Enabled = False
     Else
     Command2.Enabled = True
End If

buf = "select * from " & dediario & " where "
buf = buf & " tipo='" & Data1.Recordset.Fields("tipo") & "'"
buf = buf & " and numero like '" & Data1.Recordset.Fields("numero") & "'"
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldat
               Data2.RecordSource = buf
               Data2.Refresh
               Exit Sub
cmd1_err:
'MsgBox "Seleccione un dato " + Error$, 48, "Aviso"
               Exit Sub

End Sub

Private Sub WindowsMediaPlayer1_OpenStateChange(ByVal NewState As Long)

End Sub
Sub excel_paso()
Dim sdx As String
On Error GoTo cmd81_err
sdx = "" & Data1.Recordset.Fields("numero")
conteo_excell
Exit Sub
cmd81_err:
MsgBox "Elegir un dato ", 48, "Aviso"
Exit Sub
End Sub
Sub conteo_excell()
 Dim mydbx As Database
 Dim mytablex As Table
 Dim v, h As Integer
 Dim found As Integer
 Dim i As Integer
 Dim sdx As Double
 Dim sdx1 As Double
 Dim sdx2 As Double
 Dim sdx3 As Double
 Dim sdx4 As Double
 Dim sdxx As Double
 Dim vprecios(14) As String
    Dim Heading(25) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
    On Error GoTo cmd5612_err
   'Data1.Refresh
   Heading(1) = "Fecha"
   Heading(2) = "Fechav"

   Heading(3) = "Lo"
   Heading(4) = "Cajero"
   Heading(5) = "Caja"
   Heading(6) = "Turno"
    Heading(7) = "S"
    Heading(8) = "Tp"
    Heading(9) = "Serie"
    Heading(10) = "Numero"
    Heading(11) = "Codigo"
    Heading(12) = "Nombre"
    Heading(13) = "M"
    Heading(14) = "Subtotal"
    Heading(15) = "Igv"
    Heading(16) = "Peaje"
    Heading(17) = "Detra"
    Heading(18) = "Total"
    Heading(19) = "E"
    Heading(20) = "CobroDet"
    Heading(21) = "Nrodetra"
    Heading(22) = "Hora"
    Heading(23) = "Vehiculo"
    Heading(24) = "CodVehi"
    Heading(25) = "Placa"
   
    Call Inicio_Excel 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excel1(25, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    

Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable(dediario)
mytablex.Index = "tdetalle"
v = 5
h = 1
sdx = 0
sdx1 = 0
sdx2 = 0
sdx3 = 0
sdx4 = 0

Data1.Refresh
Do
If Data1.Recordset.EOF Then Exit Do
            objExcel.ActiveSheet.Cells(v, h + 0) = "'" & Data1.Recordset.Fields("Fecha")
            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & Data1.Recordset.Fields("fechae")
            objExcel.ActiveSheet.Cells(v, h + 2) = "'" & Data1.Recordset.Fields("local")
            objExcel.ActiveSheet.Cells(v, h + 3) = "'" & Data1.Recordset.Fields("usuario")
            objExcel.ActiveSheet.Cells(v, h + 4) = "'" & Data1.Recordset.Fields("caja")
            objExcel.ActiveSheet.Cells(v, h + 5) = "'" & Data1.Recordset.Fields("turno")
            objExcel.ActiveSheet.Cells(v, h + 6) = "'" & Data1.Recordset.Fields("sentido")
            objExcel.ActiveSheet.Cells(v, h + 7) = "'" & Data1.Recordset.Fields("tipo")
            objExcel.ActiveSheet.Cells(v, h + 8) = "'" & Data1.Recordset.Fields("serie")
            objExcel.ActiveSheet.Cells(v, h + 9) = "'" & Data1.Recordset.Fields("Numero")
            objExcel.ActiveSheet.Cells(v, h + 10) = "'" & Data1.Recordset.Fields("Codigo")
            objExcel.ActiveSheet.Cells(v, h + 11) = "'" & Data1.Recordset.Fields("nombre")
            objExcel.ActiveSheet.Cells(v, h + 12) = "'" & Data1.Recordset.Fields("moneda")
            
            objExcel.ActiveSheet.Cells(v, h + 13) = "" & Data1.Recordset.Fields("Subtotal")
            objExcel.ActiveSheet.Cells(v, h + 14) = "" & Data1.Recordset.Fields("Impuesto")
            
            objExcel.ActiveSheet.Cells(v, h + 15) = "" & Data1.Recordset.Fields("xneto")
            objExcel.ActiveSheet.Cells(v, h + 16) = "" & Data1.Recordset.Fields("tdetra")
            objExcel.ActiveSheet.Cells(v, h + 17) = "" & Data1.Recordset.Fields("total")
            objExcel.ActiveSheet.Cells(v, h + 18) = "'" & Data1.Recordset.Fields("estado")
            sdxx = Val("" & Data1.Recordset.Fields("xneto")) + Val("" & Data1.Recordset.Fields("tdetra"))
            If sdxx = Val("" & Data1.Recordset.Fields("total")) Then
               objExcel.ActiveSheet.Cells(v, h + 19) = "'"
               Else
               'MsgBox "xx"
               objExcel.ActiveSheet.Cells(v, h + 19) = "'Nocobra"
            End If
            objExcel.ActiveSheet.Cells(v, h + 20) = "'" & Data1.Recordset.Fields("denumero")
            objExcel.ActiveSheet.Cells(v, h + 21) = "'" & Data1.Recordset.Fields("hora")
            'MsgBox "" & Data1.Recordset.Fields("local") & " " & Data1.Recordset.Fields("tipo") & " " & Data1.Recordset.Fields("serie") & " " & Data1.Recordset.Fields("numero")
            mytablex.Seek "=", "" & Data1.Recordset.Fields("local"), "" & Data1.Recordset.Fields("tipo"), "" & Data1.Recordset.Fields("serie"), "" & Data1.Recordset.Fields("numero")
            If Not mytablex.NoMatch Then
               objExcel.ActiveSheet.Cells(v, h + 22) = "'" & mytablex.Fields("descripcio")
               objExcel.ActiveSheet.Cells(v, h + 23) = "'" & mytablex.Fields("producto")
               objExcel.ActiveSheet.Cells(v, h + 24) = "'" & mytablex.Fields("placa")
            End If
            v = v + 1
            If "" & Data1.Recordset.Fields("estado") = "2" Then
               sdx = sdx + Val("" & Data1.Recordset.Fields("xneto"))
               sdx1 = sdx1 + Val("" & Data1.Recordset.Fields("tdetra"))
               sdx2 = sdx2 + Val("" & Data1.Recordset.Fields("total"))
               sdx3 = sdx3 + Val("" & Data1.Recordset.Fields("subtotal"))
               sdx4 = sdx4 + Val("" & Data1.Recordset.Fields("impuesto"))
            End If
            
'mytablex.Seek "=", "" & Data1.Recordset.Fields("tipo"), "" & Data1.Recordset.Fields("numero")
'If Not mytablex.NoMatch Then
'    Do
'     If mytablex.EOF Then Exit Do
'
'     If "" & mytablex.Fields("tipo") = "" & Data1.Recordset.Fields("tipo") And "" & mytablex.Fields("numero") = "" & Data1.Recordset.Fields("numero") Then
'            sdx = sdx + Val("" & mytablex.Fields("cantidad"))
'            sdx1 = sdx1 + Val("" & mytablex.Fields("total"))
'            objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("producto")
'            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("descripcio")
'            objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablex.Fields("unidad")
'            objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytablex.Fields("factor")
'            objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("cantidad")
'            objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("precio")
'            objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytablex.Fields("total")
'            v = v + 1
'            Else: Exit Do
'     End If
'     mytablex.MoveNext
'     Loop
' End If
'            objExcel.ActiveSheet.Cells(v, h) = ""
'            objExcel.ActiveSheet.Cells(v, h + 1) = ""
'            objExcel.ActiveSheet.Cells(v, h + 2) = ""
'            objExcel.ActiveSheet.Cells(v, h + 3) = ""
'            objExcel.ActiveSheet.Cells(v, h + 4) = "" & sdx
'            objExcel.ActiveSheet.Cells(v, h + 5) = ""
'            objExcel.ActiveSheet.Cells(v, h + 6) = "" & sdx1
'            v = v + 1
  Data1.Recordset.MoveNext
  Loop
            objExcel.ActiveSheet.Cells(v, h + 0) = "'"
            objExcel.ActiveSheet.Cells(v, h + 1) = "'"
            objExcel.ActiveSheet.Cells(v, h + 2) = "'"
            objExcel.ActiveSheet.Cells(v, h + 3) = "'"
            objExcel.ActiveSheet.Cells(v, h + 4) = "'"
            objExcel.ActiveSheet.Cells(v, h + 5) = "'"
            
            objExcel.ActiveSheet.Cells(v, h + 6) = "'"
            objExcel.ActiveSheet.Cells(v, h + 7) = "'"
            objExcel.ActiveSheet.Cells(v, h + 8) = "'"
            objExcel.ActiveSheet.Cells(v, h + 9) = "'"
            objExcel.ActiveSheet.Cells(v, h + 10) = "'"
            objExcel.ActiveSheet.Cells(v, h + 11) = "'"
            objExcel.ActiveSheet.Cells(v, h + 12) = "'"
            objExcel.ActiveSheet.Cells(v, h + 13) = "" & sdx3
            objExcel.ActiveSheet.Cells(v, h + 14) = "" & sdx4
            
            objExcel.ActiveSheet.Cells(v, h + 15) = "" & sdx
            objExcel.ActiveSheet.Cells(v, h + 16) = "" & sdx1
            objExcel.ActiveSheet.Cells(v, h + 17) = "" & sdx2
            objExcel.ActiveSheet.Cells(v, h + 18) = "'"
  
  
 Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
 mytablex.Close
 Exit Sub
cmd5612_err:
MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
Exit Sub
End Sub
Public Function Formato_Excel1(Num_Campos As Integer, Nombre_Campos() As String) As Boolean
Dim i As Integer
With objExcel.ActiveSheet
        'MsgBox Num_Campos
        'Formato de las celdas de los titulos
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 8)).Font.Bold = True
        
    For i = 1 To Num_Campos Step 1
        .Cells(3, i) = Nombre_Campos(i)
    Next i
        'hasta aki pa colocar los titulos
        
        'a partir de aki ta claro que        es pa darle el ancho a las celdas ;-)
        .Columns("A").ColumnWidth = 10
        .Columns("B").ColumnWidth = 10
        .Columns("C").ColumnWidth = 2
        .Columns("D").ColumnWidth = 10
        .Columns("E").ColumnWidth = 4
        .Columns("F").ColumnWidth = 3
       
        .Columns("G").ColumnWidth = 2
        .Columns("H").ColumnWidth = 5
        .Columns("I").ColumnWidth = 7
        .Columns("J").ColumnWidth = 15
        .Columns("K").ColumnWidth = 15
        .Columns("L").ColumnWidth = 30
        .Columns("M").ColumnWidth = 1
        .Columns("N").ColumnWidth = 10
        .Columns("O").ColumnWidth = 10
        .Columns("P").ColumnWidth = 10
        
        .Columns("Q").ColumnWidth = 10
        .Columns("R").ColumnWidth = 10
        .Columns("S").ColumnWidth = 3
        .Columns("T").ColumnWidth = 10
        .Columns("U").ColumnWidth = 15
        .Columns("V").ColumnWidth = 15
        .Columns("W").ColumnWidth = 30
        .Columns("X").ColumnWidth = 15
        .Columns("Y").ColumnWidth = 15
    
   
     '9890-07477
End With
End Function






Private Sub prefijo_Change()

End Sub

Private Sub prefijo_KeyPress(KeyAscii As Integer)

End Sub

Private Sub Form_Load()
fechai = Format(Now, "dd/mm/yyyy")
fechaf = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub MM_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
DD.SetFocus

End Sub
Function convierte_decimal(sdx As Double) As String
Dim buf As String
Dim bufd As String
buf = Format(sdx, "0000000000000.00")
bufd = Mid$(buf, Len(buf) - 1, 2)
MsgBox bufd
buf = Mid$(buf, 1, 13) + bufd
convierte_decimal = buf
End Function
Function campo_detalle()

End Function
