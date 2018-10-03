VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Texcell 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar"
   ClientHeight    =   8670
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   15570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   15570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Borrar Tablas"
      Height          =   495
      Left            =   13800
      TabIndex        =   10
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox xclave 
      Height          =   495
      Left            =   5640
      MaxLength       =   8
      TabIndex        =   7
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox xcolumna 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Text            =   "18"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox xfila 
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Text            =   "2"
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cargar"
      Height          =   615
      Left            =   13560
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CargarDesdeExcell"
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   7455
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   13150
      _Version        =   393216
      HeadLines       =   2
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
            LCID            =   10250
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
            LCID            =   10250
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
   Begin VB.Label dd 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   13800
      TabIndex        =   9
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clave"
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Columnas"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FilaUltima"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.Menu fki44 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "Texcell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim found As Integer

    'MsgBox globaldat & "\excell\demo.xls"
    found = verifica_usuario("" & xclave)

    If found = 0 Then
        MsgBox "Usuario No existe ", 48, "Aviso"
        Exit Sub

    End If

    Excel_a_Access globaldat & "\excell\demo.xls", xfila, xcolumna
    cargar_bd

End Sub

Function verifica_usuario(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from vendedor where clave='" & buf & "' and conexionremota='S'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        verifica_usuario = 1

    End If

    mytablex.Close

End Function

Private Sub Excel_a_Access(Path_XLS As String, Filas As Integer, columnas As Integer)

    Dim Obj_Excel      As Object

    Dim Obj_Hoja       As Object

    Dim cn_Ado         As ADODB.Connection

    Dim rst_Ado        As ADODB.Recordset

    Dim Fila_Actual    As Integer

    Dim Columna_Actual As Integer

    Dim DATO           As Variant

    Dim mytablex       As New ADODB.Recordset

    On Error GoTo cmd908912_err

    Screen.MousePointer = vbHourglass
  
    'Nueva instancia de Excel
    Set Obj_Excel = CreateObject("Excel.Application")
  
    ' Abre el libro de Excel
    Obj_Excel.Workbooks.Open FileName:=Path_XLS
  
    ' si es la versión de Excel 97, asigna la hoja activa ( ActiveSheet )
    If Val(Obj_Excel.Application.Version) >= 8 Then
        Set Obj_Hoja = Obj_Excel.ActiveSheet
    Else
        Set Obj_Hoja = Obj_Excel

    End If

    cn.Execute ("delete from formatoexcell")
    mytablex.Open "select * from  formatoexcell", cn, adOpenStatic, adLockOptimistic
      
    'Se posiciona al final    If rst_Ado.RecordCount <> 0 Then rst_Ado.MoveLast
    ' Recorre las filas y columnas de la hoja
    For Fila_Actual = 2 To Filas
        'Nuevo registro
        mytablex.AddNew

        For Columna_Actual = 0 To columnas
            ' Va leyendo los datos de la celda indicada
            DATO = Trim$(Obj_Hoja.Cells(Fila_Actual, Columna_Actual + 1))

            'Agrega los datos al campo indicado
            Select Case Columna_Actual

                Case 0
                    mytablex.Fields(Columna_Actual).Value = Mid(Trim(DATO), 1, 15)

                Case 1
                    mytablex.Fields(Columna_Actual).Value = Mid(Trim(DATO), 1, 15)

                Case 2
                    mytablex.Fields(Columna_Actual).Value = Mid(Trim(DATO), 1, 80)

                Case 3
                    mytablex.Fields(Columna_Actual).Value = Mid(Trim(DATO), 1, 6)

                Case 7
                    mytablex.Fields(Columna_Actual).Value = Mid(Trim(DATO), 1, 6)

                Case 8
                    mytablex.Fields(Columna_Actual).Value = Mid(Trim(DATO), 1, 6)

                Case 9
                    mytablex.Fields(Columna_Actual).Value = Mid(Trim(DATO), 1, 6)

                Case 10
                    mytablex.Fields(Columna_Actual).Value = Mid(Trim(DATO), 1, 6)

                Case 11
                    mytablex.Fields(Columna_Actual).Value = Mid(Trim(DATO), 1, 6)

                Case 12
                    mytablex.Fields(Columna_Actual).Value = Mid(Trim(DATO), 1, 6)
            
                Case 4, 5, 6, 13, 14, 15, 16, 17, 18
                    mytablex.Fields(Columna_Actual).Value = Val(DATO)
            
            End Select
            
        Next
        mytablex.Update
    Next
    mytablex.Close
    Call Descargar_Objetos(Obj_Excel, Obj_Hoja)
    Screen.MousePointer = vbDefault
    MsgBox " Datos copiados ", vbInformation
    Exit Sub
cmd908912_err:
    MsgBox "No se copiar desde el excell " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub Descargar_Objetos(Obj_Excel As Object, Obj_Hoja As Object)
    Obj_Excel.ActiveWorkbook.Close False
    Obj_Excel.Quit
    Set Obj_Hoja = Nothing
    Set Obj_Excel = Nothing
  
End Sub

Sub cargar_bd()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from  formatoexcell", cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = mytablex
   
End Sub

Private Sub Command2_Click()

    Dim found As Integer

    found = verifica_usuario("" & xclave)

    If found = 0 Then
        MsgBox "Usuario No existe ", 48, "Aviso"
        Exit Sub

    End If

    dd = ""
    graba_david

End Sub

Private Sub fki44_Click()
    Texcell.Hide
    Unload Texcell

End Sub

Private Sub Form_Load()
    cargar_bd

End Sub

Sub graba_david()

    Dim mytablex As New ADODB.Recordset  'productos

    Dim mytabley As New ADODB.Recordset

    Dim vr

    Dim sdx As Double

    sdx = 0

    If Check1.Value = 1 Then
        cn.Execute ("delete from producto")
        cn.Execute ("delete from precios")
        cn.Execute ("delete from FAMILIA")
        cn.Execute ("delete from SUBFAMILIA")

    End If

    mytabley.Open "select * from formatoexcell", cn, adOpenStatic, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        mytablex.Open "select * from producto where producto='" & Trim("" & mytabley.Fields("producto")) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.AddNew
            pone_david mytablex, mytabley
            mytablex.Update
        Else
            pone_david mytablex, mytabley
            mytablex.Update

        End If

        mytablex.Close
        sdx = sdx + 1
        vr = DoEvents()
        dd = "" & sdx
        mytabley.MoveNext
    Loop
    mytabley.Close
    MsgBox "Producto proceso Terminado", 48, "Aviso"

End Sub

Sub pone_david(mytablex As ADODB.Recordset, mytabley As ADODB.Recordset)

    Dim mytablea As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    mytablex.Fields("producto") = Trim("" & mytabley.Fields("producto"))
    mytablex.Fields("barras") = Trim("" & mytabley.Fields("barras"))
    mytablex.Fields("marca") = Trim("" & mytabley.Fields("marca"))

    mytablex.Fields("descripcio") = Mid$(Trim("" & mytabley.Fields("descripcion")), 1, 80)
    mytablex.Fields("descorto") = Mid$(Trim("" & mytabley.Fields("descripcion")), 1, 20)
    mytablex.Fields("presenta") = ""
    mytablex.Fields("dsctoref") = 0
    mytablex.Fields("familia") = Mid$(Trim("" & mytabley.Fields("familia")), 1, 6)
    mytablex.Fields("subfamilia") = Mid$(Trim("" & mytabley.Fields("subfamilia")), 1, 6)
    mytablex.Fields("vecaja") = "S"
    mytablex.Fields("igv") = 18
    mytablex.Fields("isc") = 0
    mytablex.Fields("pesokgr") = 0.001
    mytablex.Fields("comision") = 0
    mytablex.Fields("monedac") = "S"
    mytablex.Fields("unidad") = Mid$(Trim("" & mytabley.Fields("unidad")), 1, 6)
    mytablex.Fields("factor") = Val("" & mytabley.Fields("factor"))
    mytablex.Fields("costou") = Val("" & mytabley.Fields("costou"))
    mytablex.Fields("costop") = Val("" & mytabley.Fields("costop"))
    mytablex.Fields("monedav") = "S"
    mytablex.Fields("estado") = "S"

    mytablez.Open "select * from familia where familia='" & Mid$(Trim("" & mytabley.Fields("FAMILIA")), 1, 6) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablez.RecordCount = 0 Then
        mytablez.AddNew
        mytablez.Fields("familia") = Mid$(Trim("" & mytabley.Fields("FAMILIA")), 1, 6)
        mytablez.Fields("descripcio") = Mid$(Trim("" & mytabley.Fields("FAMILIA")), 1, 30)
        mytablez.Fields("vetouch") = "S"
        mytablez.Update
    Else
        mytablez.Fields("familia") = Mid$(Trim("" & mytabley.Fields("FAMILIA")), 1, 6)
        mytablez.Fields("descripcio") = Mid$(Trim("" & mytabley.Fields("FAMILIA")), 1, 30)
        mytablez.Fields("vetouch") = "S"
        mytablez.Update

    End If

    mytablez.Close

    mytablez.Open "select * from subfamil where familia='" & Mid$(Trim("" & mytabley.Fields("FAMILIA")), 1, 6) & "' and subfamilia='" & Mid$(Trim("" & mytabley.Fields("subFAMILIA")), 1, 6) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablez.RecordCount = 0 Then
        mytablez.AddNew
        mytablez.Fields("subfamilia") = Mid$(Trim("" & mytabley.Fields("subFAMILIA")), 1, 6)
        mytablez.Fields("familia") = Mid$(Trim("" & mytabley.Fields("FAMILIA")), 1, 6)
        mytablez.Fields("descripcio") = Mid$(Trim("" & mytabley.Fields("FAMILIA")), 1, 30)
   
        mytablez.Update
    Else
        mytablez.Fields("subfamilia") = Mid$(Trim("" & mytabley.Fields("subFAMILIA")), 1, 6)
        mytablez.Fields("familia") = Mid$(Trim("" & mytabley.Fields("FAMILIA")), 1, 6)
        mytablez.Fields("descripcio") = Mid$(Trim("" & mytabley.Fields("FAMILIA")), 1, 30)
   
        mytablez.Update

    End If

    mytablez.Close

    mytablea.Open "select * from precios where producto='" & Trim("" & mytabley.Fields("producto")) & "' and local='01'", cn, adOpenStatic, adLockOptimistic

    If mytablea.RecordCount = 0 Then
        mytablea.AddNew
        pone_david01 mytablea, mytabley, "01"
        mytablea.Update
    Else
        pone_david01 mytablea, mytabley, "01"
        mytablea.Update

    End If

    mytablea.Close

End Sub

Sub pone_david01(mytablex As ADODB.Recordset, mytabley As ADODB.Recordset, buf As String)
    mytablex.Fields("local") = buf
    mytablex.Fields("producto") = Trim("" & mytabley.Fields("producto"))
    mytablex.Fields("factor1") = Val("" & mytabley.Fields("factor1"))
    mytablex.Fields("unidad1") = Trim("" & mytabley.Fields("unidad1"))
    mytablex.Fields("pventa1") = Val("" & mytabley.Fields("pventa1"))

    mytablex.Fields("factor2") = Val("" & mytabley.Fields("factor2"))
    mytablex.Fields("unidad2") = Trim("" & mytabley.Fields("unidad2"))
    mytablex.Fields("pventa2") = Val("" & mytabley.Fields("pventa2"))

    mytablex.Fields("factor3") = Val("" & mytabley.Fields("factor3"))
    mytablex.Fields("unidad3") = Trim("" & mytabley.Fields("unidad3"))
    mytablex.Fields("pventa3") = Val("" & mytabley.Fields("pventa3"))

End Sub

