VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form treevuti 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Utilidad"
   ClientHeight    =   9240
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Registrar Documentos Notas de Credito y Débito"
      Height          =   600
      Left            =   5400
      TabIndex        =   11
      Top             =   2500
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Borrar y Registra Tipos de Contingencia"
      Height          =   600
      Left            =   5400
      TabIndex        =   10
      Top             =   1850
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Borrar y Registrar Tipos de Notas de Credito y Débito"
      Height          =   600
      Left            =   5400
      TabIndex        =   9
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox respuesta 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   5280
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "treevuti.frx":0000
      Top             =   4320
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.CommandButton btnsalir 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   8640
      Picture         =   "treevuti.frx":0006
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir todo"
      Top             =   120
      Width           =   1695
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   14420
      _Version        =   393217
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label RucSql 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7200
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label RucYml 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5400
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Servicio de Fact.Electrónica"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5280
      TabIndex        =   6
      Top             =   3720
      Width           =   2820
   End
   Begin VB.Label Servicio 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8040
      TabIndex        =   5
      Top             =   3600
      Width           =   1905
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   15
      Left            =   1920
      TabIndex        =   3
      Top             =   3840
      Width           =   135
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   8175
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12975
   End
   Begin VB.Menu d89 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "treevuti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Facturacion Electronica Servicio 26/03/2018
Dim WithEvents objServ As servicios
Attribute objServ.VB_VarHelpID = -1

Dim rptaServicio       As String

'Facturacion Electronica Servicio 26/03/2018

'Testing Facturacion Electronica 14/03/2018
Dim rpta               As String

'Testing Facturacion Electronica 14/03/2018

Private Sub btnsalir_Click()
    d89_Click

End Sub

' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
Private Sub Command1_Click()

    On Error GoTo cmd989900_err

    If MsgBox("Desea Realizar Registro?", 1, "Aviso") <> 1 Then Exit Sub
    cn.Execute ("delete from TIPONCD")
    cn.Execute ("insert into TIPONCD values('01','Anulación de la operación','NC')")
    cn.Execute ("insert into TIPONCD values('02','Anulación por error en el RUC','NC')")
    cn.Execute ("insert into TIPONCD values('03','Corrección por error en la descripción','NC')")
    cn.Execute ("insert into TIPONCD values('04','Descuento global','NC')")
    cn.Execute ("insert into TIPONCD values('05','Descuento por Item','NC')")
    cn.Execute ("insert into TIPONCD values('06','Devolución total','NC')")
    cn.Execute ("insert into TIPONCD values('07','Devolución parcial','NC')")
    cn.Execute ("insert into TIPONCD values('08','Bonificación','NC')")
    cn.Execute ("insert into TIPONCD values('09','Disminución en el valor','NC')")
  
    cn.Execute (" insert into TIPONCD values('01','Intereses por mora','ND')")
    cn.Execute (" insert into TIPONCD values('02','Aumento en el valor','ND')")
    cn.Execute (" insert into TIPONCD values('03','Penalidades','ND')")

    MsgBox ("Proceso Correcto")

    Exit Sub
cmd989900_err:
    MsgBox "Aviso en borrar pedidoventa " + error$, 48, "Aviso"
    Exit Sub

End Sub

'Plan de Contingencia 07/05/2018
Private Sub Command2_Click()

    On Error GoTo cmd989900_err

    If MsgBox("Desea Realizar Registro?", 1, "Aviso") <> 1 Then Exit Sub
    cn.Execute ("  IF  exists (SELECT *FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME ='tipocontingencia' )   delete from tipocontingencia ")
  
    cn.Execute ("insert into tipocontingencia values('1',' Conexión a internet')")
    cn.Execute ("insert into tipocontingencia values('2','Fallas fluido eléctrico')")
    cn.Execute ("insert into tipocontingencia values('3','Desastres naturales')")
    cn.Execute ("insert into tipocontingencia values('4','Robo')")
    cn.Execute ("insert into tipocontingencia values('5','Fallas en el sistema de facturación')")
    cn.Execute ("insert into tipocontingencia values('6','Venta itinerante')")
    cn.Execute ("insert into tipocontingencia values('7','Otros')")

    MsgBox ("Proceso Correcto")

    Exit Sub
cmd989900_err:
    MsgBox "Aviso en borrar pedidoventa " + error$, 48, "Aviso"
    Exit Sub

End Sub

'Plan de Contingencia 07/05/2018

Private Sub Command3_Click()
    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

    On Error GoTo cmd989900_err

    If MsgBox("Desea Realizar Registro?", 1, "Aviso") <> 1 Then Exit Sub
    cn.Execute ("IF not exists (SELECT * FROM TIPO WHERE TIPO ='81') bEGIN  insert into TIPO VALUES  ('81','NOTA DE DEBITO BOLETA',  'F','08','B001','0','','','','','','F','S', '','','','','','',  '','','','','','','','','','','','','0','')   END")
    cn.Execute ("IF not exists (SELECT * FROM TIPO WHERE TIPO ='82') bEGIN  insert into TIPO VALUES  ('82','NOTA DE DEBITO FACTURA','F','08','F001','0','','','','','','F','S', '','','','','','','','','','','','','','','','','','','0','')  END")
    cn.Execute ("IF not exists (SELECT * FROM TIPO WHERE TIPO ='71') bEGIN  insert into TIPO VALUES  ('71','NOTA DE CREDITO BOLETA','E','07','B001','0','','','','','','E','E', '','','','','','','','','','','','','','','','','','','0','')  END")
    cn.Execute ("IF not exists (SELECT * FROM TIPO WHERE TIPO ='72') bEGIN  insert into TIPO VALUES  ('72','NOTA CREDITO FACTURA','E','07','F001','0','','','','','','E','E', '','','','','','','','','','','','','','','','','','','0','')  END")
    MsgBox ("Proceso Correcto")

    Exit Sub
cmd989900_err:
    MsgBox "Error!!! " + error$, 48, "Aviso"
    Exit Sub
    ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018

End Sub

Private Sub d89_Click()
    treevuti.Hide
    Unload treevuti

End Sub

Private Sub Form_Load()

    Dim sp  As String

    Dim sh  As String

    Dim sp1 As String

    Dim sh1 As String

    Dim sp2 As String

    Dim sh2 As String

    Dim sp3 As String

    Dim sh3 As String

    ' 26/07/2018 Desactivar Facturacion Electronica
    'Testing Proyecto Facturacion Electronica 05/04/2018
    V_EstadoSistema = Obtiene_EstadoSistema

    If V_EstadoSistema = "FE BYH" Then
        Set objServ = New servicios

        rptaServicio = objServ.ObtenerEstado("facturador-local")

        If rptaServicio = "El servicio está detenido" Then
            servicio = "Servicio Detenido"
        ElseIf rptaServicio = "El servicio está activo" Then
            servicio = "Servicio Activo"

        End If

    End If

    'Facturacion Electronica Servicio 26/03/2018
    ' 26/07/2018 Desactivar Facturacion Electronica

    sp = "sp"
    sp1 = "sp1"
    sp2 = "sp2"
    sp3 = "sp3"

    With TreeView1
        .Style = tvwTreelinesPlusMinusText
        .LineStyle = tvwRootLines
        .PathSeparator = "\"
        .Indentation = Screen.TwipsPerPixelX * 5 '256
        '
        ' No permitir la edición automática del texto
        .LabelEdit = tvwManual
        ' Para que se pueda expandir al seleccionar un nodo,
        ' cambia este valor a True,
        ' si se deja en False, tendrás que hacer doble-click
        .SingleSel = False
        ' Para que al perder el foco,
        ' se siga viendo el que está seleccionado
        .HideSelection = False
        '
        .refresh

    End With

    TreeView1.Nodes.Add , , sp, "Utilidades"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Registrar Sistema"
    '  TreeView1.Nodes.Add sp, tvwChild, sh, "Registrar Sistema Centralizador"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Copia Seguridad"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Copia Seguridad Correo"
    ' TreeView1.Nodes.Add sp, tvwChild, sh, "Mensaje Celular"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Enviar Correo"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Inicializa"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Excell->VisiOrion" 'Orion.V5"
    ' TreeView1.Nodes.Add sp, tvwChild, sh, "OrionV2"
    'TreeView1.Nodes.Add sp, tvwChild, sh, "OrionV4"
    ' TreeView1.Nodes.Add sp, tvwChild, sh, "OrionV5"
    '  TreeView1.Nodes.Add sp, tvwChild, sh, "Uniflex"
    ' TreeView1.Nodes.Add sp, tvwChild, sh, "Monica"
    ' TreeView1.Nodes.Add sp, tvwChild, sh, "Siscont"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Importar sql-sql"
    'TreeView1.Nodes.Add sp, tvwChild, sh, "Yactayo"
    'TreeView1.Nodes.Add sp, tvwChild, sh, "Lolfar"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Ejecutar"
    ' TreeView1.Nodes.Add sp, tvwChild, sh, "RECAVE"
    'TreeView1.Nodes.Add sp, tvwChild, sh, "Prueba Balanza"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Prueba Sql"
    ' TreeView1.Nodes.Add sp, tvwChild, sh, "ANT"
    TreeView1.Nodes.Add sp, tvwChild, sh, "CONSULTAS SQL - EXCEL"
    TreeView1.Nodes.Add sp, tvwChild, sh, "LIMPIA REGISTRO DE TABLAS"
    TreeView1.Nodes.Add sp, tvwChild, sh, "ACTUALIZAR FACTURACION ELECTRONICA BYH"
    TreeView1.Nodes.Add sp, tvwChild, sh, "UPDATE CONSULTAS 2017-2018"
  
    'inicio 25/04/2017 pll
    'TreeView1.Nodes.Add sp, tvwChild, sh, "Backup tablas"
    'fin 25/04/2017 pll
    'inicio 02/01/2018 pll para el proceso comprobantes electronicos sunat
    TreeView1.Nodes.Add sp, tvwChild, sh, "Comprobantes Electronicos Sunat"
    'fin 02/01/2018 pll
    'inicio 09/01/2018 pll
    TreeView1.Nodes.Add sp, tvwChild, sh, "Inicializa Stock"
    'fin 09/01/2018 pll
    'inicio 01/02/2018 pll
    ' TreeView1.Nodes.Add sp, tvwChild, sh, "Backup Comprobantes Electronicos"
    'fin 01/02/2018 pll
    
    'Testing Facturacion Electronica 14/03/2018
    TreeView1.Nodes.Add sp, tvwChild, sh, "Verificar configuracion electronica"
    'Testing Facturacion Electronica 14/03/2018
    
    '''27/07/2017 kenyo Testing Completo al Sistema
    For I = 1 To TreeView1.Nodes.count - 1
        'TreeView1.Nodes(i).ExpandedImage = "Open"
        TreeView1.Nodes(I).Expanded = True
    Next I

    Exit Sub
    '''27/07/2017 kenyo Testing Completo al Sistema

    'cmdLlenarTree_Click

End Sub

Sub actualiza_estadotipo()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM PARAMECA where  CAJA='01'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
           
        If (mytablex.Fields("FTB") = "S") Then cn.Execute ("update TIPO SET ESTADOT='2' WHERE tipo='1'") Else cn.Execute ("update TIPO SET ESTADOT='1' WHERE tipo='1'")
        If (mytablex.Fields("FTF") = "S") Then cn.Execute ("update TIPO SET ESTADOT='2' WHERE tipo='2' ") Else cn.Execute ("update TIPO SET ESTADOT='1' WHERE tipo='2'")
        If (mytablex.Fields("FNV") = "S") Then cn.Execute ("update TIPO SET ESTADOT='2' WHERE tipo='5' ") Else cn.Execute ("update TIPO SET ESTADOT='1' WHERE tipo='5'")
        If (mytablex.Fields("FBM") = "S") Then cn.Execute ("update TIPO SET ESTADOT='2' WHERE tipo='3' ") Else cn.Execute ("update TIPO SET ESTADOT='1' WHERE tipo='3'")
        If (mytablex.Fields("FFM") = "S") Then cn.Execute ("update TIPO SET ESTADOT='2' WHERE tipo='4'") Else cn.Execute ("update TIPO SET ESTADOT='1' WHERE tipo='4'")
    
        '''27/07/2017 kenyo Testing Completo al Sistema
        If (mytablex.Fields("FNC") = "S") Then cn.Execute ("update TIPO SET ESTADOT='2' WHERE tipo='4'") Else cn.Execute ("update TIPO SET ESTADOT='1' WHERE tipo='4'")
    
        '''27/07/2017 kenyo Testing Completo al Sistema
   
    End If

    mytablex.Close

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim salida           As Boolean

    Dim llego            As Boolean

    Dim my_cantidad_file As Integer

    Dim con_internet     As Boolean

    If Node = "Lolfar" Then
        If InputBox("Clave de Paso", "Aviso") = "KALIPOS" Then
            tlolfar.Show 1

        End If

    End If

    If Node = "LIMPIA REGISTRO DE TABLAS" Then
        If InputBox("Clave de Paso", "Aviso", "******") = "KALIPO" Then
            FrmLimpiaTablas.Show 1

        End If

    End If

    'inicio 25/04/2017 pll
    If Node = "Backup tablas" Then
        Frm_backup.Show 1

    End If

    'fin 25/04/2017 pll
    'inicio 02/01/2018 pll
    If Node = "Comprobantes Electronicos Sunat" Then
        con_internet = CBool(tptovta.Online())

        If con_internet = False Then
            MsgBox "No tiene internet en estos momentos"
        Else
            'inicio 09/02/2018 pll para ver la caja
            Call read_caja(my_caja)
            Call Datos_Empresa(my_struc_datos_empresa(), my_caja, salida, 0)

            If my_struc_datos_empresa(0).esunat = "A" Then
                MsgBox "No se Podrá enviar desde esta opcion esta modalidad Automatico"
            ElseIf my_struc_datos_empresa(0).esunat = "M" Then
                frm_ESunat.Show 1

            End If

        End If

    End If

    'fin 02/01/2018 pll
    'inicio 01/02/2018 pll
    If Node = "Backup Comprobantes Electronicos" Then
        Call bkp_en_crear
        Call bkp_en_out_PROCESADO
        MsgBox "Proceso Realizado ", 48, "Aviso"

    End If

    'fin 01/02/2018 pll

    'Testing Facturacion Electronica 14/03/2018
    If Node = "Verificar configuracion electronica" Then
        respuesta.Visible = True
        Call valida_facturacionElectronica

        If rpta = "" Then
            respuesta.Text = "TODO OK"
        Else
            respuesta.Text = rpta

        End If
       
    End If

    'Testing Facturacion Electronica 14/03/2018

    'inicio 09/01/2018 pll
    If Node = "Inicializa Stock" Then
        If MsgBox("Desea Inicializar Stock", 1, "Aviso") <> 1 Then Exit Sub
        If InputBox("Llave de Paso ", "Control", "") <> "KALIPOS" Then Exit Sub
        cn.Execute ("Update almacen set saldo='0',salida='0',svirtual='0'")
        cn.Execute ("Update producto set SALDOINI='0'")
        cn.Execute ("DELETE almacen WHERE BODEGA='02'")
        MsgBox "Proceso Realizado ", 48, "Aviso"

    End If

    'fin 09/01/2018 pll

    If Node = "UPDATE CONSULTAS 2017-2018" Then

        Dim buf As String

        'PRODUCTO
        buf = buf & "IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='producto' and  COLUMN_NAME ='DIASALERTA') BEGIN  ALTER TABLE PRODUCTO add DIASALERTA NVARCHAR(8) END "

        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='parame' and  COLUMN_NAME ='formatocierre') BEGIN  ALTER TABLE parame add formatocierre NVARCHAR(1) END "

        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='usercajapedido') BEGIN  "
        buf = buf & " CREATE TABLE usercajapedido ( caja varCHAR(3),p1 varCHAR(3),p2 varCHAR(3),p3 varCHAR(3),p4 varCHAR(3),p5 varCHAR(3),p6 varCHAR(3),p7 varCHAR(3),p8 varCHAR(3),p9 varCHAR(3),p10 varCHAR(3),p11 varCHAR(3),p12 varCHAR(3),p13 varCHAR(3),p14 varCHAR(3),p15 varCHAR(3),p16 varCHAR(3),p17 varCHAR(3),p18 varCHAR(3),p19 varCHAR(3),p20 varCHAR(3)) end "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='parame' and COLUMN_NAME ='cajapedido') BEGIN  ALTER TABLE parame  add cajapedido  CHAR(1) END "

        buf = buf & " IF not exists (SELECT *FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME ='sexo') BEGIN  create table sexo (sexo nvarchar(6),descripcio nvarchar(30)) END "
        buf = buf & " IF not exists (SELECT *FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME ='proyecto') BEGIN  create table proyecto (proyecto nvarchar(6),descripcio nvarchar(30)) END "
        buf = buf & " IF not exists (SELECT *FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME ='procedencia') BEGIN  create table procedencia (procedencia nvarchar(6),descripcio nvarchar(30)) END"
        buf = buf & " IF not exists (SELECT *FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME ='talla') BEGIN  create table talla (talla nvarchar(6),descripcio nvarchar(30)) END "

        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='producto' and  COLUMN_NAME ='sexo') BEGIN  ALTER TABLE PRODUCTO add sexo NVARCHAR(6) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='producto' and  COLUMN_NAME ='proyecto') BEGIN  ALTER TABLE PRODUCTO add proyecto NVARCHAR(6) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='producto' and  COLUMN_NAME ='procedencia') BEGIN  ALTER TABLE PRODUCTO add procedencia NVARCHAR(6) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='producto' and  COLUMN_NAME ='talla') BEGIN  ALTER TABLE PRODUCTO add talla NVARCHAR(6) END "
        buf = buf & " ALTER TABLE producto ALTER COLUMN  procedencia nvarchar(6) "

        buf = buf & " ALTER TABLE RECIBO ALTER COLUMN  CSECCION9 nvarchar(10) "
        buf = buf & " ALTER TABLE RECIBO ALTER COLUMN  CSECCION10 nvarchar(10) "
        buf = buf & " ALTER TABLE TLOCAL ALTER COLUMN  CABECERA char(300) "

        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='PARAMECA' and COLUMN_NAME ='obligaclavemesa') BEGIN  ALTER TABLE PARAMECA  add obligaclavemesa  CHAR(1) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='PARAMECA' and COLUMN_NAME ='puertocua') BEGIN  ALTER TABLE PARAMECA  add puertocua  nvarchar(80) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='PARAMECA' and COLUMN_NAME ='colacua') BEGIN  ALTER TABLE PARAMECA  add colacua   CHAR(1) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='PARAMECA' and COLUMN_NAME ='obligavdelivery') BEGIN  ALTER TABLE PARAMECA  add obligavdelivery  CHAR(1) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='TIPO' and COLUMN_NAME ='ESTADOT') BEGIN  ALTER TABLE tipo  add ESTADOT  CHAR(1) END "

        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='parameca' and COLUMN_NAME ='multicomanda') BEGIN  ALTER TABLE parameca  add multicomanda  CHAR(1) END "

        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='parame' and COLUMN_NAME ='saldocierre') BEGIN  ALTER TABLE parame  add saldocierre  CHAR(1) END "

        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='parame' and COLUMN_NAME ='vemesa') BEGIN  ALTER TABLE parame  add vemesa  CHAR(1) END "

        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='ALMACEN' and COLUMN_NAME ='svirtual') BEGIN  ALTER TABLE ALMACEN   add svirtual  FLOAT END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='cordenc' and COLUMN_NAME ='personas') BEGIN  ALTER TABLE cordenc   add personas  FLOAT END "
        buf = buf & " ALTER TABLE CORREOS ALTER COLUMN  txtfromname char(150) "
        buf = buf & " ALTER TABLE CORREOS ALTER COLUMN  txtfromemail char(150)"
        buf = buf & " ALTER TABLE CORREOS ALTER COLUMN  txtto char(150)"

        buf = buf & "  IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='producto' and  COLUMN_NAME ='obligacomentario') BEGIN  ALTER TABLE PRODUCTO add obligacomentario NVARCHAR(1) END "

        buf = buf & "  IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='producto' and  COLUMN_NAME ='obligacomentario') BEGIN  ALTER TABLE PRODUCTO add obligacomentario NVARCHAR(1) END "

        buf = buf & " ALTER TABLE dcomanda DROP COLUMN indx  "
        buf = buf & " ALTER TABLE dcomanda add indx bigint   identity "

        buf = buf & "IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='parameca' and COLUMN_NAME ='puertodelivery') BEGIN  ALTER TABLE parameca   add puertodelivery  nvarchar(80) END "
        buf = buf & "IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='parameca' and COLUMN_NAME ='coladelivery') BEGIN  ALTER TABLE parameca   add coladelivery  char(1) END "
   
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='parame' and  COLUMN_NAME ='tiporeceta') BEGIN  ALTER TABLE parame add tiporeceta NVARCHAR(1) END "
   
        ' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018
        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
        buf = buf & "  IF not exists (SELECT *FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME ='TIPONCD') BEGIN  create table TIPONCD (codigo nvarchar(3),descripcio nvarchar(60),tipo nvarchar(3)) END "
  
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='FACTURA' and COLUMN_NAME ='ESTADO_SUNAT') BEGIN  ALTER TABLE FACTURA   add ESTADO_SUNAT  nvarchar(80) END  "
        
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='FACTURA' and COLUMN_NAME ='TIPONCD') BEGIN  ALTER TABLE FACTURA   add TIPONCD  nvarchar(3) END "
  
        buf = buf & "   ALTER TABLE TIPO ALTER COLUMN  SERIE nvarchar(4) "
        buf = buf & "   ALTER TABLE FACTURA ALTER COLUMN  ESTADO_SUNAT nvarchar(80) "
        ' Testing Proyecto Facturacion Electronica Nota de Credito 16/04/2018
        ' Testing Proyecto Facturacion Electronica Nota de Credito 06/2018

        'Plan de Contingencia 07/05/2018
        buf = buf & "   IF not exists (SELECT *FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME ='tipocontingencia')   BEGIN  create table tipocontingencia (codigo nvarchar(6),descripcion nvarchar(60)) END   "
        'Plan de Contingencia 07/05/2018
    
        ' Varios Locales FE 18/05/2018
        buf = buf & "  IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='TLOCAL' and  COLUMN_NAME ='codsede') BEGIN  ALTER TABLE TLOCAL add codsede NVARCHAR(4) END "
        ' Varios Locales FE 18/05/2018
    
        '05/03/2018 Crea automaticamente tabla vendedorcomision
        buf = buf & " IF not exists (SELECT *FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME ='vendedorcomision') BEGIN create table vendedorcomision( producto nchar(15),descripcion nchar(120),codigo nchar(11),nombre  nchar(60),comision float)END  "
        '05/03/2018 Crea automaticamente tabla vendedorcomision
    
        '02/05/2018 Recargo %
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='FACTURA' and COLUMN_NAME ='RECARGO')  BEGIN  ALTER TABLE FACTURA   add RECARGO  FLOAT END "
        '02/05/2018 Recargo %
    
        ' Mesas Nombre 21/05/2018
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='MESA' and COLUMN_NAME ='ATENCION')  BEGIN  ALTER TABLE MESA   add   ATENCION  char(1) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='MESA' and COLUMN_NAME ='CODIGO')  BEGIN  ALTER TABLE MESA   add   CODIGO char(11) END "
        ' Mesas Nombre 21/05/2018

        'Color por familia y producto  30/05/2018
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='paramecacolor' and  COLUMN_NAME ='c35') BEGIN  ALTER TABLE paramecacolor add c35  char(10) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='paramecacolor' and  COLUMN_NAME ='d35') BEGIN  ALTER TABLE paramecacolor add d35  char(10) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='paramecacolor' and  COLUMN_NAME ='e35') BEGIN  ALTER TABLE paramecacolor add e35  char(10) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='paramecacolor' and  COLUMN_NAME ='f35') BEGIN  ALTER TABLE paramecacolor add f35  char(10) END "
  
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='paramecacolor' and  COLUMN_NAME ='c36') BEGIN  ALTER TABLE paramecacolor add c36  char(10) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='paramecacolor' and  COLUMN_NAME ='d36') BEGIN  ALTER TABLE paramecacolor add d36  char(10) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='paramecacolor' and  COLUMN_NAME ='e36') BEGIN  ALTER TABLE paramecacolor add e36  char(10) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='paramecacolor' and  COLUMN_NAME ='f36') BEGIN  ALTER TABLE paramecacolor add f36  char(10) END "
         
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='paramecacolor' and  COLUMN_NAME ='c37') BEGIN  ALTER TABLE paramecacolor add c37  char(10) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='paramecacolor' and  COLUMN_NAME ='d37') BEGIN  ALTER TABLE paramecacolor add d37  char(10) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='paramecacolor' and  COLUMN_NAME ='e37') BEGIN  ALTER TABLE paramecacolor add e37  char(10) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='paramecacolor' and  COLUMN_NAME ='f37') BEGIN  ALTER TABLE paramecacolor add f37  char(10) END "
    
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='familia' and  COLUMN_NAME ='c') BEGIN  ALTER TABLE familia add c  char(10) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='familia' and  COLUMN_NAME ='d') BEGIN  ALTER TABLE familia add d  char(10) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='familia' and  COLUMN_NAME ='e') BEGIN  ALTER TABLE familia add e  char(10) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='familia' and  COLUMN_NAME ='f') BEGIN  ALTER TABLE familia add f  char(10) END "
      
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='parame' and  COLUMN_NAME ='colorproductofamilia') BEGIN  ALTER TABLE parame add colorproductofamilia NVARCHAR(1) END "
 
        'Color por familia y producto  30/05/2018

        '' 10/07/2018 Edicion Comanda
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='parame' and COLUMN_NAME ='comanda') BEGIN  ALTER TABLE parame  add comanda  CHAR(2) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='parame' and COLUMN_NAME ='opcionnombre') BEGIN  ALTER TABLE parame  add opcionnombre  CHAR(2) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='parame' and COLUMN_NAME ='tamanocomanda') BEGIN  ALTER TABLE parame  add tamanocomanda  CHAR(2) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='parameca' and COLUMN_NAME ='formatocomanda') BEGIN  ALTER TABLE parameca  add formatocomanda  CHAR(1) END "
        '' 10/07/2018 Edicion Comanda
    
        '''' 17/07/2018 Factura de Exportación
        buf = buf & " IF not exists (SELECT *FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME ='TblTipoOperacion')   BEGIN  create table TblTipoOperacion (tipooperacion nvarchar(6),descripcio nvarchar(30),orden int ) END "
        buf = buf & " IF not exists (SELECT *FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME ='TblTipoIgv')   BEGIN  create table TblTipoIgv (tipoigv nvarchar(6),descripcio nvarchar(30),orden int ) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='parame' and COLUMN_NAME ='opcionexportacion') BEGIN  ALTER TABLE parame  add opcionexportacion  NVARCHAR(1) END "
        '''' 17/07/2018 Factura de Exportación

        '''' 19/07/2018 Campo Provincia en Cliente
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='clientes' and COLUMN_NAME ='provincia') BEGIN  ALTER TABLE clientes  add provincia  NVARCHAR(15) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='proveedo' and COLUMN_NAME ='provincia') BEGIN  ALTER TABLE proveedo  add provincia  NVARCHAR(15) END "
        '''' 19/07/2018 Campo Provincia en Cliente

        ''''' 25/07/2018 Delivery y Para Llevar desde mozo
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='salon' and COLUMN_NAME ='tipo') BEGIN  ALTER TABLE salon  add tipo  NVARCHAR(1) END "
        ''''' 25/07/2018 Delivery y Para Llevar desde mozo

        ' 26/07/2018 Desactivar Facturacion Electronica
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='parame' and COLUMN_NAME ='estadosistema') BEGIN  ALTER TABLE parame  add estadosistema  NVARCHAR(6) END "
        ' 26/07/2018 Desactivar Facturacion Electronica
    
        '07/08/2018 No descuenta stock en guia de remision
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='TIPO' and COLUMN_NAME ='descuentastock') BEGIN  ALTER TABLE tipo  add descuentastock  CHAR(1) END "
        '07/08/2018 No descuenta stock en guia de remision
    
        '13/08/2018 Integración FE - Pizzeria
        '''' 11/12/2017 SubReceta
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='parame' and  COLUMN_NAME ='tiporeceta') BEGIN  ALTER TABLE parame add tiporeceta NVARCHAR(1) END "

        buf = buf & "  IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='producto' and  COLUMN_NAME ='costoreceta') BEGIN  ALTER TABLE PRODUCTO add costoreceta NVARCHAR(1) END "
        '''' 11/12/2017 SubReceta
    
        '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='parame' and COLUMN_NAME ='tcostoreceta') BEGIN  ALTER TABLE parame  add tcostoreceta  CHAR(2) END "
        '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
        '13/08/2018 Integración FE - Pizzeria
    
        '15/08/2018 Cambiar Descripcion de producto venta de ventas
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='parame' and COLUMN_NAME ='cambiadescripcion') BEGIN  ALTER TABLE parame  add cambiadescripcion  CHAR(1) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='parame' and COLUMN_NAME ='nuevoproducto') BEGIN  ALTER TABLE parame  add nuevoproducto  CHAR(1) END "
        '15/08/2018 Cambiar Descripcion de producto venta de ventas

        'Balanza 2/3 dígitos
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='parameca' and COLUMN_NAME ='digitos') BEGIN  ALTER TABLE parameca  add digitos  CHAR(1) END "
        'Balanza 2/3 dígitos
    
        '24/08/2018  Delivery por mesa
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='mesa' and COLUMN_NAME ='dnombre') BEGIN  ALTER TABLE mesa  add dnombre  NVARCHAR(60) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='mesa' and COLUMN_NAME ='telefono') BEGIN  ALTER TABLE mesa  add telefono  NVARCHAR(11) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='mesa' and COLUMN_NAME ='ddireccion') BEGIN  ALTER TABLE mesa  add ddireccion  NVARCHAR(200) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='mesa' and COLUMN_NAME ='referencia') BEGIN  ALTER TABLE mesa  add referencia  NVARCHAR(120) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='dcomanda' and COLUMN_NAME ='nombre') BEGIN  ALTER TABLE dcomanda  add nombre  NVARCHAR(60) END "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='logcomanda' and COLUMN_NAME ='nombre') BEGIN  ALTER TABLE logcomanda  add nombre  NVARCHAR(60) END "
        '24/08/2018  Delivery por mesa
       
        '27/08/2018 Producto delivery automatico
        buf = buf & "IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='producto' and  COLUMN_NAME ='deliveryautom') BEGIN  ALTER TABLE PRODUCTO add deliveryautom NVARCHAR(1) END "
        '27/08/2018 Producto delivery automatico
    
        'Reporte de ingresos (Cobranzas) CONTASIS
        buf = buf & "IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='fpago' and  COLUMN_NAME ='cuentacontable') BEGIN  ALTER TABLE fpago add cuentacontable NVARCHAR(50) END "
        'Reporte de ingresos (Cobranzas) CONTASIS
    
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='FACTURA' and COLUMN_NAME ='E_SUNAT') BEGIN  ALTER TABLE FACTURA   add E_SUNAT  nvarchar(5) END  "
        buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='FACTURA' and COLUMN_NAME ='CDR') BEGIN  ALTER TABLE FACTURA   add CDR  nvarchar(50) END  "
    
        cn.Execute (buf)
   
        'Color por familia y producto  30/05/2018
        Dim buf2 As String

        buf2 = buf2 & " update FAMILIA set c=( SELECT top 1 colorfamilia1 FROM paramecacolor where  caja='01'), d=( SELECT top 1 colorfamilia2 FROM paramecacolor where  caja='01'), e=( SELECT top 1 colorfamilia3 FROM paramecacolor where  caja='01'), f='S' where d is null "
        buf2 = buf2 & " update parame set colorproductofamilia='N' where colorproductofamilia is null "
   
        buf2 = buf2 & " update parame set comanda='DL' where comanda is null "
        buf2 = buf2 & " update parame set opcionnombre='DL' where opcionnombre is null "
        buf2 = buf2 & " update parame set tamanocomanda='12' where tamanocomanda is null "
   
        '''' 17/07/2018 Factura de Exportación
        buf2 = buf2 & " delete from TblTipoOperacion where tipooperacion='01' "
        buf2 = buf2 & " delete from TblTipoOperacion where tipooperacion='02' "
        buf2 = buf2 & " delete from TblTipoOperacion where tipooperacion='03' "
        buf2 = buf2 & " delete from TblTipoOperacion where tipooperacion='04' "
        buf2 = buf2 & " delete from TblTipoOperacion where tipooperacion='05' "
    
        buf2 = buf2 & " insert into  TblTipoOperacion values('01','Venta Interna','1') "
        buf2 = buf2 & " insert into  TblTipoOperacion values('02','Exportación','2' ) "
        buf2 = buf2 & " insert into  TblTipoOperacion values('03','No Domiciliados','3') "
        buf2 = buf2 & " insert into  TblTipoOperacion values('04','Venta Interna-Anticipos','4') "
        buf2 = buf2 & " insert into  TblTipoOperacion values('05','Venta Itinerante','5')"
     
        buf2 = buf2 & " delete from TblTipoIgv where tipoigv='10' "
        buf2 = buf2 & " delete from TblTipoIgv where tipoigv='40' "
    
        buf2 = buf2 & " insert into  TblTipoIgv values('10','Gravado-Operación Onerosa','1') "
        buf2 = buf2 & " insert into  TblTipoIgv values('40','Exportación','2')"
     
        ''''' 25/07/2018 Delivery y Para Llevar desde mozo
        buf2 = buf2 & " update SALON set TIPO='C' where TIPO is null "
        ''''' 25/07/2018 Delivery y Para Llevar desde mozo

        ' 26/07/2018 Desactivar Facturacion Electronica
        buf2 = buf2 & " update parame set estadosistema='FE BYH' where estadosistema is null "
        buf2 = buf2 & " update parame set opcionexportacion='N' where opcionexportacion is null "
        ' 26/07/2018 Desactivar Facturacion Electronica
  
        '07/08/2018 No descuenta stock en guia de remision
        buf2 = buf2 & " update tipo set descuentastock='-' where descuentastock is null "
        '07/08/2018 No descuenta stock en guia de remision
  
        '13/08/2018 Integración FE - Pizzeria
        '''' 11/12/2017 SubReceta
        buf2 = buf2 & " update producto set costoreceta='S' where costoreceta is null "
        '''' 11/12/2017 SubReceta
   
        '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
        buf2 = buf2 & " update parame set tcostoreceta='CU' where tcostoreceta is null "
        buf2 = buf2 & " update parame set tiporeceta='E' where tiporeceta is null "
        buf2 = buf2 & " update parame set tiporeceta='E' where tiporeceta='' "
    
        buf2 = buf2 & " update parame set vemesa='S' where vemesa is null "
        buf2 = buf2 & " update parame set vemesa='S' where vemesa='' "
        '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
        '13/08/2018 Integración FE - Pizzeria
   
        '15/08/2018 Cambiar Descripcion de producto venta de ventas
        buf2 = buf2 & " update parame set cambiadescripcion='N' where cambiadescripcion is null "
        buf2 = buf2 & " update parame set nuevoproducto='N' where nuevoproducto is null "
        '15/08/2018 Cambiar Descripcion de producto venta de ventas

        'Balanza 2/3 dígitos
        buf2 = buf2 & " update parameca set digitos='2' where digitos is null "
        'Balanza 2/3 dígitos
    
        '10/07/2018 Edicion Comanda
        buf2 = buf2 & " update parameca set formatocomanda='D' where formatocomanda is null "
        '10/07/2018 Edicion Comanda
   
        '24/08/2018  Delivery por mesa
        buf2 = buf2 & " update mesa set dnombre='' where dnombre is null "
        buf2 = buf2 & " update mesa set telefono='' where telefono is null "
        buf2 = buf2 & " update mesa set ddireccion='' where ddireccion is null "
        buf2 = buf2 & " update mesa set referencia='' where referencia is null "
        buf2 = buf2 & " update dcomanda set nombre='' where nombre is null "
        buf2 = buf2 & " update logcomanda set nombre='' where nombre is null "
        '24/08/2018  Delivery por mesa
      
        '27/08/2018 Producto delivery automatico
        buf2 = buf2 & " update producto set deliveryautom='N' where deliveryautom is null "
        '27/08/2018 Producto delivery automatico

        'Reporte de ingresos (Cobranzas) CONTASIS
        buf2 = buf2 & " update fpago set cuentacontable='' where cuentacontable is null "

        'Reporte de ingresos (Cobranzas) CONTASIS
        
        'danielkheavy tlocal agregado campos
        
        'buf = buf & "   ALTER TABLE tlocal  ADD toperacion varchar(50) "
        'buf = buf & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='tlocal' and COLUMN_NAME ='toperacion') BEGIN  ALTER TABLE tlocal  add toperacion varchar(50) END "
        'buf = buf & "   ALTER TABLE tlocal ALTER COLUMN  timpresion varchar(50) "
        'buf = buf & "   ALTER TABLE tlocal ALTER COLUMN  esunat varchar(50) "
        
        'danielkheavy tlocal agregado campos

        cn.Execute (buf2)
        'Color por familia y producto  30/05/2018
  
        MsgBox ("Proceso Correcto")

    End If

    '07/08/2018 ACTUALIZAR BD A FACTURACION ELECTRONICA BYH
    If Node = "ACTUALIZAR BD A FE BYH" Then

        Dim buf3 As String

        buf3 = buf3 & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='TLOCAL' and COLUMN_NAME ='toperacion')  BEGIN  ALTER TABLE TLOCAL   add toperacion  VARCHAR(5) END "
        buf3 = buf3 & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='TLOCAL' and COLUMN_NAME ='timpresion')  BEGIN  ALTER TABLE TLOCAL   add timpresion  VARCHAR(5) END "
        buf3 = buf3 & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='TLOCAL' and COLUMN_NAME ='esunat')  BEGIN  ALTER TABLE TLOCAL   add esunat  VARCHAR(5) END  "
        buf3 = buf3 & " IF not exists (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='TLOCAL' and COLUMN_NAME ='codsede') BEGIN  ALTER TABLE TLOCAL add codsede NVARCHAR(4) END "
        cn.Execute (buf3)
   
        Dim buf4 As String

        buf4 = buf4 & " update TLOCAL set toperacion='E' where toperacion is null "
        buf4 = buf4 & " update TLOCAL set timpresion='G' where timpresion is null "
        buf4 = buf4 & " update TLOCAL set ESUNAT='A' where ESUNAT is null "
        buf4 = buf4 & " update TLOCAL set codsede='' where codsede is null "
         
        cn.Execute (buf4)
        MsgBox ("Proceso Correcto")

    End If

    '07/08/2018 ACTUALIZAR BD A FACTURACION ELECTRONICA BYH

    If Node = "CONSULTAS SQL - EXCEL" Then
        Form1.Show 1

    End If

    If Node = "Prueba Sql" Then
        tfsql.Show 1

    End If

    If Node = "Yactayo" Then
        tdenise.Show 1

    End If

    If Node = "Excell->Orion.V5" Then
        Texcell.Show 1

    End If

    If Node = "RECAVE" Then
        tinterfa.Show 1

    End If

    If Node = "Prueba Balanza" Then
        tprcom.Show 1

    End If

    If Node = "Enviar Correo" Then
        Tsms.Show 1

    End If

    If Node = "ANT" Then
        TUNIFLEX.Show 1

    End If

    If Node = "OrionV5" Then
        If InputBox("Clave de Paso", "Aviso") = "KALIPOS" Then
            tinterfa.Show 1

        End If

    End If

    If Node = "Uniflex" Then
        If InputBox("Clave de Paso", "Aviso") = "VICKY" Then
            TUNIFLEX.Show 1

        End If

    End If

    If Node = "Monica" Then
        If InputBox("Clave de Paso", "Aviso") = "VICKY" Then
            tinterfa.Show 1

        End If

    End If

    If Node = "Siscont" Then
        If InputBox("Clave de Paso", "Aviso") = "VICKY" Then
            tinterfa.Show 1

        End If

    End If

    If Node = "Registrar Sistema" Then
        regsiste.xlicencia = "LICENCIA"
        regsiste.Show 1

    End If

    If Node = "Registrar Sistema Centralizador" Then
        regsiste.xlicencia = "LICENCIACENTRALIZADO"
        regsiste.Show 1

    End If

    If Node = "Copia Seguridad" Then
        CreateBackup

    End If

    If Node = "Copia Seguridad Correo" Then
        Call CreateBackupBd
        Call ComprimeBackupBd
        Call eliminar_BackupBd

        If MsgBox("Desea enviar Backup al correo ", 1, "Aviso") = vbOK Then
            MsgBox ("Enviando Correo...Espere")
            Call envio_correosBackupBd

        End If

        Call eliminar_BackupBd

    End If

    If Node = "Reportes" Then
        frmVisualQry.Show 1

    End If

    If Node = "Mensaje Celular" Then
        frmsms.Show 1

    End If

    If Node = "Importar sql-sql" Then
        texporta.Show 1
        Exit Sub

    End If

    If Node = "OrionV4" Then
        If InputBox("Clave de Paso", "Aviso") <> "KALIPO" Then
            Exit Sub

        End If

        tinterfa.Show 1

    End If

    If Node = "OrionV2" Then
        If InputBox("Clave de Paso", "Aviso") <> "KALIPO" Then
            Exit Sub

        End If

        tinterfa.Show 1

    End If

    If Node = "Inicializa" Then
        If InputBox("Clave de Paso", "Aviso") <> "KALIPO" Then
            Exit Sub
   
        End If

        Dim sw       As Integer

        Dim mytablex As New ADODB.Recordset

        sw = 0
        mytablex.Open "select * from vendedor where codigo='" & gusuario & "' and inicializa='S'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            sw = 1

        End If

        mytablex.Close

        If sw = 1 Then
            tinicia.Show 1

        End If

    End If

    If Node = "Ejecutar" Then
        If InputBox("Clave de Paso", "Aviso") = "VICKY" Then
            tejecuta.Show 1

        End If

    End If

End Sub

'Testing Facturacion Electronica 14/03/2018
Public Function valida_facturacionElectronica()

    Dim mytable As New ADODB.Recordset

    Dim mysql   As String

    Dim k       As Integer

    Dim salida  As String

    rpta = ""

    ' Valida RUC en LOCAL
    mysql = ""
    mysql = "SELECT  codigo1  from Tlocal"
    mytable.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytable.RecordCount > 0 Then  'si existe
        Do

            If mytable.EOF Then Exit Do
            If mytable.Fields("CODIGO1") = "" Then
                rpta = "* FALTA AGREGAR RUC (CODIGO1) A LOCAL"
                Exit Do

            End If

            mytable.MoveNext
        Loop

    End If

    mytable.Close
 
    ' Valida SERIE en CAJA
    mysql = ""
    mysql = "SELECT CAJA,DESCRIPCIO,SERIETB,SERIETF from PARAMECA WHERE CAJA <>00"
    mytable.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytable.RecordCount > 0 Then  'si existe
        Do

            If mytable.EOF Then Exit Do
            If Len(mytable.Fields("serietb")) <> 0 Then
                If Len(mytable.Fields("serietb")) <> 4 And Len(mytable.Fields("serietb")) = 0 Then
                    rpta = rpta & " * VERIFICAR SERIE DE DOCUMENTO"
                    Exit Do

                End If

            End If

            mytable.MoveNext
        Loop

    End If

    mytable.Close

    'Facturacion Electronica Servicio 26/03/2018
    If servicio = "Servicio Detenido" Then
        Shell ("C:\BYH\iniciar.bat"), vbNormalFocus
        servicio = "Servicio Activo"

    End If

    'Facturacion Electronica Servicio 26/03/2018
 
    'Testing Proyecto Facturacion Electronica 05/04/2018
    mysql = ""
    mysql = "SELECT  codigo1  from Tlocal where codigo='01' "
    mytable.Open mysql, cn, adOpenStatic, adLockOptimistic

    If mytable.RecordCount > 0 Then  'si existe
        RucSql = mytable.Fields("codigo1")

    End If

    mytable.Close
 
    Call lee_conf_RUC("C:\BYH\SERVICIO\VISITEC\application.yml", "A")

    If RucYml <> RucSql Then
        rpta = rpta & " * RUC de Local DIFERENTE A RUC de Yml"

    End If

    'Testing Proyecto Facturacion Electronica 05/04/2018

    'Valida DIRECTORIOS
    Const ATTR_DIRECTORY = 16

    If Dir$("D:\ce_output", ATTR_DIRECTORY) = "" Then
        rpta = rpta & " * Carpeta D:\ce_output NO EXISTE"

    End If
 
    If Dir$("D:\ce_output\CREA", ATTR_DIRECTORY) = "" Then
        rpta = rpta & " * Carpeta D:\ce_output\CREA NO EXISTE"

    End If

    If Dir$("D:\ce_output\ERROR", ATTR_DIRECTORY) = "" Then
        rpta = rpta & " * Carpeta D:\ce_output\ERROR NO EXISTE"

    End If

    If Dir$("D:\ce_output\FIRMADO", ATTR_DIRECTORY) = "" Then
        rpta = rpta & " * Carpeta D:\ce_output\FIRMADO NO EXISTE"

    End If

    If Dir$("D:\ce_output\PROCESADO", ATTR_DIRECTORY) = "" Then
        rpta = rpta & " * Carpeta D:\ce_output\PROCESADO NO EXISTE"

    End If

    Exit Function

End Function

'Testing Facturacion Electronica 14/03/2018
