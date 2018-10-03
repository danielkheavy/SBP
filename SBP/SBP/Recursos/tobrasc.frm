VERSION 5.00
Begin VB.Form tobrasc 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de Control Obras"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   12030
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFF00&
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   11970
      TabIndex        =   0
      Top             =   0
      Width           =   12030
   End
   Begin VB.Menu dkl9922 
      Caption         =   "&1.Archivos"
      Begin VB.Menu yt67222 
         Caption         =   "&1.Trabajadores"
         Begin VB.Menu g18821 
            Caption         =   "&1.Empleados"
         End
         Begin VB.Menu d9922 
            Caption         =   "&2.Profesionales"
         End
         Begin VB.Menu cat7722 
            Caption         =   "&3.Categorias"
         End
         Begin VB.Menu h7744 
            Caption         =   "&4.Faltas"
         End
      End
      Begin VB.Menu dj8822 
         Caption         =   "&2.Agentes Externos"
         Begin VB.Menu dk8833 
            Caption         =   "&1.Proveedores"
         End
         Begin VB.Menu dk88221 
            Caption         =   "&2.Subcontratisttas"
         End
      End
      Begin VB.Menu d8833 
         Caption         =   "&3.Direcciones"
         Begin VB.Menu ddl8822 
            Caption         =   "&1.Codigos Postales"
         End
         Begin VB.Menu dk8822w 
            Caption         =   "&2.Poblaciones"
         End
         Begin VB.Menu t848422 
            Caption         =   "&3.Provincias"
         End
      End
      Begin VB.Menu ju7833 
         Caption         =   "&4.Relacion de Obras"
         Begin VB.Menu cdj8343 
            Caption         =   "&1.Centro de TRabajo"
         End
         Begin VB.Menu sh6622 
            Caption         =   "&2.Fases"
         End
         Begin VB.Menu dj77223 
            Caption         =   "&3.Tareas"
         End
      End
      Begin VB.Menu dfk8822 
         Caption         =   "&5.Otros"
         Begin VB.Menu dlo8822 
            Caption         =   "&1.Materiales"
         End
         Begin VB.Menu am88221 
            Caption         =   "&2.Maquinas"
         End
      End
   End
   Begin VB.Menu dk8822ss 
      Caption         =   "&Partes"
      Begin VB.Menu dk8822 
         Caption         =   "&1.Operarios"
      End
      Begin VB.Menu ni8811 
         Caption         =   "&2.Maquinas"
      End
      Begin VB.Menu mai811 
         Caption         =   "&3.Materiales"
      End
      Begin VB.Menu io9maq 
         Caption         =   "&4.Operarios y Maquinas"
      End
      Begin VB.Menu dj882200 
         Caption         =   "&5.Subcontrataciones"
      End
      Begin VB.Menu dj88imp 
         Caption         =   "&6.Importes facturadas"
      End
      Begin VB.Menu polk8822 
         Caption         =   "&7.Planificacion"
      End
   End
   Begin VB.Menu df9922 
      Caption         =   "&Listados"
      Begin VB.Menu dfi1 
         Caption         =   "&1.Fichas de Operarios"
      End
      Begin VB.Menu fima9 
         Caption         =   "&2.Fichas de maquinas"
      End
      Begin VB.Menu dl8822 
         Caption         =   "&3.Ficha de materiales"
      End
      Begin VB.Menu dk89ce 
         Caption         =   "&4.Costes por centro"
      End
      Begin VB.Menu k89fac 
         Caption         =   "&5.Listado de facturaciion"
      End
      Begin VB.Menu dkires9 
         Caption         =   "&6.Comparativa de Resultados"
      End
      Begin VB.Menu dk8811 
         Caption         =   "&7.Archivos generales"
      End
   End
   Begin VB.Menu fdj8822 
      Caption         =   "&Utilidades"
      Begin VB.Menu ci822 
         Caption         =   "&1.Configuracion"
      End
      Begin VB.Menu dj77822 
         Caption         =   "&2.Copia Seguridad"
      End
   End
   Begin VB.Menu dfolo232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tobrasc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
