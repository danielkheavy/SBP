VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form tpuente 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar desde Excell Conteos Fisicos Externos"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   9375
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   6975
   End
End
Attribute VB_Name = "tpuente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Agregar la referencia de ADO
'----------------------------------------------------

Private Sub Excel_a_Access(PATH_XLS As String)


Dim Obj_Excel As Object
Dim Obj_Hoja As Object
Dim Fila_Actual As Integer
Dim Columna_Actual As Integer
Dim Dato As Variant
Dim mytablex As New ADODB.Recordset

    Screen.MousePointer = vbHourglass
    Set Obj_Excel = CreateObject("Excel.Application")
    Obj_Excel.Workbooks.Open Filename:=PATH_XLS
    If Val(Obj_Excel.Application.Version) >= 8 Then
        Set Obj_Hoja = Obj_Excel.ActiveSheet
    Else
        Set Obj_Hoja = Obj_Excel
    End If
    
        mytablex.Open "select * from saldoini where producto='" & Trim$(Obj_Hoja.Cells(1, 1)) & "' and  local='" & tsaldoin.LOCAL1 & "' and bodega='" & extra_loquesea(tsaldoin.bodega) & "' and fecha='" & tsaldoin.fecha & "'", cn, adOpenStatic, adLockOptimistic
        If mytablex.RecordCount > 0 Then
            Dato = Trim$(Obj_Hoja.Cells(1, 1))
            mytablex.Fields("cantidad") = Val(Dato)
            mytablex.Update
        End If
    Call Descargar_Objetos(Obj_Excel, Obj_Hoja)
    Screen.MousePointer = vbDefault
    MsgBox " Datos copiados ", vbInformation
Exit Sub
'Error
ErrSub:

Call Descargar_Objetos(Obj_Excel, Obj_Hoja)
MsgBox Err.Description, vbCritical
Screen.MousePointer = vbDefault
    
End Sub

'Descarga los objetos y los cierra
Sub Descargar_Objetos(Obj_Excel As Object, Obj_Hoja As Object)

    
    Obj_Excel.ActiveWorkbook.Close False
    Obj_Excel.Quit
    Set Obj_Hoja = Nothing
    Set Obj_Excel = Nothing

End Sub



Private Sub Command1_Click()


' Pasar como parámetro el nombre y path de la _
  base de datos y del libro excel, el nombre de la tabla _
  y la cantidad de filas y columnas de la hoja a leer
  
Dim pdtcarga As String
CommonDialog1.DialogTitle = "Seleccione un archivo Grafico"
CommonDialog1.InitDir = globaldir & "\excell"
CommonDialog1.Filter = "Archivos Excell|*.xls"
CommonDialog1.ShowOpen
'Si seleccionamos un archivo mostramos la ruta
If CommonDialog1.Filename <> "" Then
   pdtcarga = CommonDialog1.Filename
   Call Excel_a_Access(pdtcarga)

   'foto = LoadPicture(fotonombre)
Else
   'Si no mostramos un texto de advertencia de que no se seleccionó _   ninguno, ya que FileName devuelve una cadena vacía
   'Label1 = "No se seleccionó ningún archivo"
End If


                    
End Sub

Private Sub Form_Load()
Label1 = "PASOS PARA CARGAR FORMATOS EXCELL" + Chr$(10) + Chr$(13)
Label1 = Label1 & "1.DEBE ESTAR CARGADO LOS PRODUCTOS  " + Chr$(10) + Chr$(13)
Label1 = Label1 & "" + Chr$(10) + Chr$(13)
Label1 = Label1 & "FORMATO EN EXCELL" + Chr$(10) + Chr$(13)
Label1 = Label1 & "COLUMNA :A1=CODIGOPRODUCTO" + Chr$(10) + Chr$(13)
Label1 = Label1 & "COLUMNA :B1=CANTIDAD CONTADO" + Chr$(10) + Chr$(13)
End Sub
