VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H80000009&
      Caption         =   "Exit"
      Height          =   975
      Left            =   6600
      Picture         =   "botones.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5040
      Width           =   915
   End
   Begin VB.CommandButton CmdAdd 
      BackColor       =   &H80000009&
      Caption         =   "Add"
      Height          =   975
      Left            =   5640
      Picture         =   "botones.frx":1B42
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5040
      Width           =   915
   End
   Begin VB.CommandButton CmdBottom 
      BackColor       =   &H80000009&
      Caption         =   "Bottom"
      Height          =   975
      Left            =   3000
      Picture         =   "botones.frx":3684
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5040
      Width           =   915
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H80000009&
      Caption         =   "Delete"
      Height          =   975
      Left            =   4680
      Picture         =   "botones.frx":51C6
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5040
      Width           =   915
   End
   Begin VB.CommandButton CmdNext 
      BackColor       =   &H80000009&
      Caption         =   "Next"
      Height          =   975
      Left            =   2040
      Picture         =   "botones.frx":6D08
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5040
      Width           =   915
   End
   Begin VB.CommandButton CmdPrevious 
      BackColor       =   &H80000009&
      Caption         =   "Previous"
      Height          =   975
      Left            =   1080
      Picture         =   "botones.frx":884A
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5040
      Width           =   915
   End
   Begin VB.CommandButton CmdTop 
      BackColor       =   &H80000009&
      Caption         =   "Top"
      Height          =   975
      Left            =   120
      Picture         =   "botones.frx":A38C
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5040
      Width           =   915
   End
   Begin VB.TextBox txtFax 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Fax"
      DataMember      =   "Customers"
      DataSource      =   "deSample"
      Height          =   285
      Left            =   2490
      TabIndex        =   21
      Top             =   4470
      Width           =   2415
   End
   Begin VB.TextBox txtPhone 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Phone"
      DataMember      =   "Customers"
      DataSource      =   "deSample"
      Height          =   285
      Left            =   2490
      TabIndex        =   19
      Top             =   4095
      Width           =   2415
   End
   Begin VB.TextBox txtCountry 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Country"
      DataMember      =   "Customers"
      DataSource      =   "deSample"
      Height          =   285
      Left            =   2490
      TabIndex        =   17
      Top             =   3720
      Width           =   2385
   End
   Begin VB.TextBox txtPostalCode 
      BackColor       =   &H00C0C0C0&
      DataField       =   "PostalCode"
      DataMember      =   "Customers"
      DataSource      =   "deSample"
      Height          =   285
      Left            =   2490
      TabIndex        =   15
      Top             =   3330
      Width           =   1050
   End
   Begin VB.TextBox txtRegion 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Region"
      DataMember      =   "Customers"
      DataSource      =   "deSample"
      Height          =   285
      Left            =   2490
      TabIndex        =   13
      Top             =   2955
      Width           =   1035
   End
   Begin VB.TextBox txtCity 
      BackColor       =   &H00C0C0C0&
      DataField       =   "City"
      DataMember      =   "Customers"
      DataSource      =   "deSample"
      Height          =   285
      Left            =   2490
      TabIndex        =   11
      Top             =   2580
      Width           =   2115
   End
   Begin VB.TextBox txtAddress 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Address"
      DataMember      =   "Customers"
      DataSource      =   "deSample"
      Height          =   285
      Left            =   2490
      TabIndex        =   9
      Top             =   2190
      Width           =   3975
   End
   Begin VB.TextBox txtContactTitle 
      BackColor       =   &H00C0C0C0&
      DataField       =   "ContactTitle"
      DataMember      =   "Customers"
      DataSource      =   "deSample"
      Height          =   285
      Left            =   2490
      TabIndex        =   7
      Top             =   1800
      Width           =   3975
   End
   Begin VB.TextBox txtContactName 
      BackColor       =   &H00C0C0C0&
      DataField       =   "ContactName"
      DataMember      =   "Customers"
      DataSource      =   "deSample"
      Height          =   285
      Left            =   2490
      TabIndex        =   5
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox txtCompanyName 
      BackColor       =   &H00C0C0C0&
      DataField       =   "CompanyName"
      DataMember      =   "Customers"
      DataSource      =   "deSample"
      Height          =   285
      Left            =   2490
      TabIndex        =   3
      Top             =   1050
      Width           =   4575
   End
   Begin VB.TextBox txtCustomerID 
      BackColor       =   &H00C0C0C0&
      DataField       =   "CustomerID"
      DataMember      =   "Customers"
      DataSource      =   "deSample"
      Height          =   285
      Left            =   2490
      TabIndex        =   1
      Top             =   675
      Width           =   2145
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Visit http://www.planet-source-code.com"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2400
      MouseIcon       =   "botones.frx":C7CE
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   120
      Width           =   3585
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fax:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   2115
      TabIndex        =   20
      Top             =   4515
      Width           =   345
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Phone:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   1890
      TabIndex        =   18
      Top             =   4140
      Width           =   570
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Country:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   1740
      TabIndex        =   16
      Top             =   3765
      Width           =   720
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "PostalCode:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   1470
      TabIndex        =   14
      Top             =   3375
      Width           =   990
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Region:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   1830
      TabIndex        =   12
      Top             =   3000
      Width           =   630
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "City:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   2085
      TabIndex        =   10
      Top             =   2625
      Width           =   375
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   1725
      TabIndex        =   8
      Top             =   2235
      Width           =   735
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ContactTitle:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   1380
      TabIndex        =   6
      Top             =   1860
      Width           =   1080
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ContactName:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1275
      TabIndex        =   4
      Top             =   1485
      Width           =   1185
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CompanyName:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1140
      TabIndex        =   2
      Top             =   1095
      Width           =   1320
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CustomerID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1395
      TabIndex        =   0
      Top             =   720
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   6735
      Left            =   0
      Picture         =   "botones.frx":CAD8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7935
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      FillStyle       =   2  'Horizontal Line
      Height          =   375
      Left            =   2160
      Top             =   6360
      Width           =   3975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'API for opening a browser
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                (ByVal hWnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) As Long

Private Sub CmdAdd_Click()
    deSample.rsCustomers.AddNew
End Sub

Private Sub CmdBottom_Click()
    deSample.rsCustomers.MoveLast
End Sub

Private Sub CmdDelete_Click()
Dim StrDel As String
    StrDel = MsgBox("Are you sure you want to delete?", vbOKCancel + vbQuestion, "confirmation")
        If vbOK Then
            deSample.rsCustomers.Delete adAffectCurrent
        End If
End Sub

Private Sub CmdExit_Click()
    End
End Sub

Private Sub CmdNext_Click()
    deSample.rsCustomers.MoveNext
End Sub

Private Sub CmdPrevious_Click()
    deSample.rsCustomers.MovePrevious
End Sub

Private Sub CmdTop_Click()
    deSample.rsCustomers.MoveFirst
End Sub



Private Sub Form_Load()
'centering the form on the screen
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

'gradient title bar
GradForceColors = True
'Replace the 'True' below with 'False' if you want that the gradient will
'be drawn  horizonally.
GradVerticalGradient = True
'Set colors for active caption
GradForcedText = vbWhite
'Replace the two color values below to change the active title bar color
GradForcedFirst = &H800000
GradForcedSecond = &H8000
'Set colors for Inactive caption
GradForcedTextA = &HC0C0C0
'Replace the two color values below to change the inactive title bar color
GradForcedFirstA = vbBlack
GradForcedSecondA = vbBlue
GradientGetCapsFont
GradientForm Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Answer As Integer
    Answer = MsgBox("Thank you for downloading!, program by Froilan C. Alejando", vbOKOnly, "Thank you!")
        If vbOK Then
        BrowseTo "http://www.planet-source-code.com"
            Unload Me
        End If

'release gadient form
GradientReleaseForm Me
End Sub



Private Sub Image1_Click()

End Sub

Private Sub Label2_Click()
    BrowseTo "http://www.planet-source-code.com"
End Sub

Private Sub BrowseTo(ByRef pstrURL As String)
    ' Opens users default web browser and navigates to the selected URL
    Call ShellExecute(Me.hWnd, "Open", pstrURL, "", "", True)
End Sub

Private Sub DisplayAsURL(ByRef Link As VB.Label)
    ' Changes a link to look like a URL
    Link.Font.Underline = True
    Link.ForeColor = vbBlue
    Link.MousePointer = vbCustom
    Link.MouseIcon = LoadPicture(App.Path & "\Hand.cur")
End Sub


