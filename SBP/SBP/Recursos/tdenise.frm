VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tdenise 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Minimarket Denisse"
   ClientHeight    =   8580
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   13680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   13680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Ver datos"
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13455
      Begin VB.CommandButton Command4 
         Caption         =   "Subfamilias"
         Height          =   495
         Left            =   10440
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Familias"
         Height          =   495
         Left            =   8880
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Productos"
         Height          =   495
         Left            =   6840
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   495
         Left            =   4560
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4215
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   7435
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   2055
         Left            =   240
         TabIndex        =   5
         Top             =   5880
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   3625
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Barras"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   5520
         Width           =   2295
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Buscar"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Menu flo543 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tdenise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn1 As New ADODB.Connection
Dim registros As New ADODB.Recordset

Private Sub Command1_Click()
minimarket_denisse
End Sub

Private Sub Command2_Click()
graba_producto

End Sub

Private Sub Command3_Click()
pasar_familias
End Sub

Private Sub Command4_Click()
pasar_subfamilias
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
codigo_barras "" & registros.Fields("codart")
End Sub

Private Sub flo543_Click()
tdenise.Hide
Unload tdenise
End Sub

Sub minimarket_denisse()
Dim buf As String
If cnn1.State = 1 Then
   cnn1.Close
End If
'Z:\Microsoft Office\Nueva carpeta\SICR\DATA01
cnn1.CursorLocation = adUseClient
cnn1.Provider = "Microsoft.ACE.OLEDB.12.0"
'cnn1.Properties("Data Source") = "d:\SICR\data01\tab2010.mdb"
cnn1.Properties("Data Source") = "D:\SICR\DATA01\tab2010.mdb"
cnn1.Properties("Jet OLEDB:Database Password") = "RCRJJF"
cnn1.Open
registros.CursorLocation = adUseClient
Frame1.Visible = True
If Len(Trim(Text1)) = 0 Then
registros.Open "select * from articulos ", cnn1, adOpenKeyset, adLockOptimistic
Else
registros.Open "select * from articulos where denart like '" & Text1 & "'", cnn1, adOpenKeyset, adLockOptimistic
End If
Set DataGrid1.DataSource = registros

End Sub
Sub codigo_barras(buf As String)
Dim registros1 As New ADODB.Recordset
registros1.CursorLocation = adUseClient
registros1.Open "select * from cod_barra where codart='" & buf & "'", cnn1, adOpenKeyset, adLockOptimistic
Set DataGrid2.DataSource = registros1

End Sub

Sub graba_producto()
Dim mytablex As New ADODB.Recordset  'productos
Dim vr
Dim sdx As Double
sdx = 1
cn.Execute ("delete from producto")
cn.Execute ("delete from precios")

mytablex.Open "select * from producto ", cn, adOpenStatic, adLockOptimistic
registros.MoveFirst
Do
If registros.EOF Then Exit Do
   mytablex.AddNew
   pone_registro mytablex, registros
   mytablex.Update
sdx = sdx + 1
vr = DoEvents()
dd = "" & sdx
registros.MoveNext
Loop
MsgBox "Producto proceso Terminado", 48, "Aviso"

End Sub
Sub pone_registro(mytablex As ADODB.Recordset, mytabley As ADODB.Recordset)
Dim mytablea As New ADODB.Recordset
Dim mytablez As New ADODB.Recordset

mytablex.Fields("producto") = "" & mytabley.Fields("codart")
pone_denisse_barras mytablex, mytabley
mytablex.Fields("descripcio") = Mid$("" & mytabley.Fields("denart"), 1, 60)
mytablex.Fields("descorto") = Mid$("" & mytabley.Fields("dabart"), 1, 20)

mytablex.Fields("presenta") = ""
mytablex.Fields("dsctoref") = 0

mytablex.Fields("familia") = pone_familias_denisse(Trim("" & mytabley.Fields("codgru")))
mytablex.Fields("subfamilia") = pone_subfamilias_denisse(Trim("" & mytabley.Fields("codgru")), Trim("" & mytabley.Fields("coddiv")))
mytablex.Fields("seccion") = ""
mytablex.Fields("marca") = ""
mytablex.Fields("categoria") = ""
mytablex.Fields("linea") = ""
mytablex.Fields("color") = ""
mytablex.Fields("fabrica") = ""
mytablex.Fields("serie") = ""
mytablex.Fields("peso") = "N"
mytablex.Fields("servicio") = ""
mytablex.Fields("vecaja") = "S"
mytablex.Fields("igv") = 18
mytablex.Fields("isc") = 0
mytablex.Fields("pesokgr") = 0.001
mytablex.Fields("comision") = 0
mytablex.Fields("monedac") = "S"
mytablex.Fields("unidad") = "UND"
mytablex.Fields("factor") = 1
mytablex.Fields("costou") = Val("" & mytabley.Fields("PRECPR"))
mytablex.Fields("costop") = Val("" & mytabley.Fields("PRECPR"))
mytablex.Fields("monedav") = "S"
mytablex.Fields("estado") = "S"
mytablex.Fields("minimo") = 10
mytablex.Fields("maximo") = 100

'grabando precios al local nro 1
mytablea.Open "select * from precios where producto='" & "" & mytabley.Fields("codart") & "' and local='01'", cn, adOpenStatic, adLockOptimistic
If mytablea.RecordCount = 0 Then
   mytablea.AddNew
   pone_detalle01 mytablea, mytabley, "01", mytabley.Fields("codart")
   mytablea.Update
End If
mytablea.Close

End Sub
Sub pone_detalle01(mytablex As ADODB.Recordset, mytabley As ADODB.Recordset, buf As String, sdx As String)
mytablex.Fields("local") = buf
mytablex.Fields("producto") = "" & sdx
mytablex.Fields("ccosto") = ""
mytablex.Fields("factor1") = 1
mytablex.Fields("unidad1") = "UND"
mytablex.Fields("pventa1") = Val("" & mytabley.Fields("prevta"))

End Sub
Sub pone_denisse_barras(mytablexx As ADODB.Recordset, mytabley As ADODB.Recordset)
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from cod_barra where codart='" & "" & mytabley.Fields("codart") & "'", cnn1, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   'If "" & mytabley.Fields("unimed") <> "PESABLE" Then
   mytablexx.Fields("barras") = "" & mytablex.Fields("codrel")
   'Else
   'mytablexx.Fields("codigobalanza") = Right$(mytablex.Fields("codrel"), 4)
   'End If
End If
mytablex.Close
End Sub
Function pone_familias_denisse(buf) As String
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from grupos where codgru='" & buf & "'", cnn1, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
pone_familias_denisse = Trim(Mid$(Trim("" & mytablex.Fields("dabgru")), 1, 6))
End If
mytablex.Close

End Function
Function pone_subfamilias_denisse(buf, buf1) As String
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from divisiones where codgru='" & buf & "' and coddiv='" & buf1 & "'", cnn1, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
pone_subfamilias_denisse = Trim(Mid$(Trim("" & mytablex.Fields("dabdiv")), 1, 6))
End If
mytablex.Close

End Function
Sub pasar_familias()
Dim mytablex As New ADODB.Recordset
Dim mytabley As New ADODB.Recordset
cn.Execute ("delete from familia")
mytabley.Open "select * from familia ", cn, adOpenStatic, adLockOptimistic
mytablex.Open "select * from grupos ", cnn1, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
mytabley.AddNew
mytabley.Fields("familia") = Trim(Mid$(Trim("" & mytablex.Fields("dabgru")), 1, 6))
mytabley.Fields("descripcio") = Trim(Mid$(Trim("" & mytablex.Fields("dengru")), 1, 15))
mytabley.Update
mytablex.MoveNext
Loop
mytablex.Close
End Sub
Sub pasar_subfamilias()
Dim mytablex As New ADODB.Recordset
Dim mytabley As New ADODB.Recordset
cn.Execute ("delete from subfamil")
mytabley.Open "select * from subfamil ", cn, adOpenStatic, adLockOptimistic
mytablex.Open "select * from divisiones ", cnn1, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
mytabley.AddNew
mytabley.Fields("familia") = pone_familias_denisse(Trim(Mid$(Trim("" & mytablex.Fields("codgru")), 1, 6)))
mytabley.Fields("subfamilia") = Trim(Mid$(Trim("" & mytablex.Fields("dabdiv")), 1, 6))
mytabley.Fields("descripcio") = Trim(Mid$(Trim("" & mytablex.Fields("dendiv")), 1, 15))
mytabley.Update
mytablex.MoveNext
Loop
mytablex.Close
End Sub








