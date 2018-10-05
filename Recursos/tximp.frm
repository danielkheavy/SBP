VERSION 5.00
Begin VB.Form tximp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Importacion de Datos de otros sistemas"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label indx 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "tximp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    Dim cnx      As New ADODB.Connection

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    Dim sdx      As Double

    On Error GoTo cmd1_err

    cn.Execute "update producto set estado='S'"
    Exit Sub

    indx = ""
    cnx.CursorLocation = adUseClient
    cnx.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=aries;Data Source=(local)"

    'pasamos los productos
    mytabley.Open "SELECT * FROM tab_productos ", cnx, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount = 0 Then  'si existe
        MsgBox "No hay Datos"
        Exit Sub

    End If
   
    cn.Execute " delete from precios"
    cn.Execute " delete from producto"
    mytablex.Open "SELECT * FROM producto ", cn, adOpenDynamic, adLockOptimistic
    mytablez.Open "SELECT * FROM precios ", cn, adOpenDynamic, adLockOptimistic
   
    sdx = 0
    Do

        If mytabley.EOF Then Exit Do
        sdx = sdx + 1
        indx = "" & sdx
        vr = DoEvents
   
        mytablex.AddNew
        mytablex.Fields("producto") = "" & Val(Trim("" & mytabley.Fields("cod_item")))
        mytablex.Fields("barras") = ""
        mytablex.Fields("descripcio") = Mid$(Trim("" & mytabley.Fields("desc_articulo")), 1, 60)
        mytablex.Fields("descorto") = Mid$(Trim("" & mytabley.Fields("desc_articulo")), 1, 20)
        mytablex.Fields("familia") = Mid$(Trim("" & mytabley.Fields("cod_familia")), 1, 6)
        mytablex.Fields("subfamilia") = Mid$(Trim("" & mytabley.Fields("cod_subfamilia")), 1, 6) & Mid$(Trim("" & mytabley.Fields("cod_grupo")), 1, 6)
        mytablex.Fields("seccion") = ""
        mytablex.Fields("marca") = Mid$(Trim("" & mytabley.Fields("marca")), 1, 6)
        mytablex.Fields("categoria") = ""
        mytablex.Fields("linea") = ""
        mytablex.Fields("color") = ""
        mytablex.Fields("fabrica") = ""
        mytablex.Fields("serie") = ""
        mytablex.Fields("peso") = "N"
        mytablex.Fields("servicio") = ""
        mytablex.Fields("vecaja") = "S"
        mytablex.Fields("igv") = 18#
        mytablex.Fields("isc") = 0#

        mytablex.Fields("pesokgr") = 0.001
        mytablex.Fields("comision") = 0#

        If Val("" & mytabley.Fields("moneda")) = 0 Or Val("" & mytabley.Fields("moneda")) = 2 Then
            mytablex.Fields("monedac") = "S"
            mytablex.Fields("monedav") = "S"
        Else
            mytablex.Fields("monedac") = "D"
            mytablex.Fields("monedav") = "D"

        End If

        mytablex.Fields("unidad") = Mid$(Trim("" & mytabley.Fields("unidad_medida")), 1, 6)
        mytablex.Fields("factor") = 1
        mytablex.Fields("costou") = 0
        mytablex.Fields("costop") = 0
        mytablex.Fields("estado") = "S"

        mytablex.Update

        'grabando precios
        mytablez.AddNew
        mytablez.Fields("local") = "01"
        mytablez.Fields("producto") = Trim("" & Val(Trim("" & mytabley.Fields("cod_item"))))
        mytablez.Fields("ccosto") = ""
        mytablez.Fields("factor1") = 1
        mytablez.Fields("unidad1") = Trim("" & mytabley.Fields("unidad_medida"))
        mytablez.Fields("pventa1") = 0
        mytablez.Update

        mytabley.MoveNext
    Loop
    Exit Sub

    'pasamos los clientes
    mytabley.Open "SELECT * FROM tab_clientes ", cnx, adOpenKeyset, adLockOptimistic
   
    cn.Execute " delete from clientes "
    mytablex.Open "SELECT * FROM clientes ", cn, adOpenDynamic, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        mytablex.Fields("codigo") = Mid$(Trim("" & mytabley.Fields("num_ruc")), 1, 11)
        mytablex.Fields("nombre") = Mid$(Trim("" & mytabley.Fields("razon_social")), 1, 60)
        mytablex.Fields("direccion") = Mid$(Trim("" & mytabley.Fields("direccion")), 1, 60)
        mytablex.Fields("telefono") = Mid$(Trim("" & mytabley.Fields("telefono1")), 1, 11)
        mytablex.Fields("estado") = "ACTIVO"
        mytablex.Fields("moneda") = "S"
        'mytablex.Fields("fechaalta") = Format(Now, "dd/mm/yyyy")
        mytablex.Update
        mytabley.MoveNext
    Loop
    mytabley.Close
    mytablex.Close
    Exit Sub

    'pasamos los proveedores
    mytabley.Open "SELECT * FROM tab_proveedor ", cnx, adOpenKeyset, adLockOptimistic
   
    cn.Execute " delete from proveedo "
    mytablex.Open "SELECT * FROM proveedo ", cn, adOpenDynamic, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        mytablex.Fields("codigo") = Mid$(Trim("" & mytabley.Fields("num_ruc")), 1, 11)
        mytablex.Fields("nombre") = Mid$(Trim("" & mytabley.Fields("razon_social")), 1, 60)
        mytablex.Fields("direccion") = Mid$(Trim("" & mytabley.Fields("direccion")), 1, 60)
        mytablex.Fields("telefono") = Mid$(Trim("" & mytabley.Fields("telefono1")), 1, 11)
        mytablex.Fields("estado") = "ACTIVO"
        mytablex.Fields("moneda") = "S"
        'mytablex.Fields("fechaalta") = Format(Now, "dd/mm/yyyy")
        mytablex.Update
        mytabley.MoveNext
    Loop
    mytabley.Close
    mytablex.Close
    Exit Sub

    'pasamos familias
    mytabley.Open "SELECT * FROM tab_familias_items ", cnx, adOpenKeyset, adLockOptimistic
   
    cn.Execute " delete from subfamil"
    mytablex.Open "SELECT * FROM subfamil ", cn, adOpenDynamic, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        mytablex.Fields("familia") = Mid$(Trim("" & mytabley.Fields("cod_familia")), 1, 6)
        mytablex.Fields("subfamilia") = Mid$(Trim("" & mytabley.Fields("cod_subfamilia")), 1, 6) & Mid$(Trim("" & mytabley.Fields("cod_grupo")), 1, 6)
        mytablex.Fields("descripcio") = Mid$(Trim("" & mytabley.Fields("descripcion")), 1, 15)
        mytablex.Update
        mytabley.MoveNext
    Loop
    mytabley.Close
    mytablex.Close
   
    'pasamos MARCAS
    mytabley.Open "SELECT * FROM tab_marcas ", cnx, adOpenKeyset, adLockOptimistic
    cn.Execute " delete from MARCA"
   
    mytablex.Open "SELECT * FROM marca ", cn, adOpenDynamic, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        mytablex.Fields("marca") = Mid$(Trim("" & mytabley.Fields("marca")), 1, 6)
        mytablex.Fields("descripcio") = Mid$(Trim("" & mytabley.Fields("descripcion")), 1, 15)
        mytablex.Update
        mytabley.MoveNext
    Loop
    mytabley.Close
    mytablex.Close
   
    '---------
   
    MsgBox "Proceso Terminado"
   
    Exit Sub
cmd1_err:
    MsgBox "No se puede conectar sql " + error$, 48, "Aviso"
    Exit Sub

End Sub

