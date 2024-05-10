VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "FluWork v10"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   -540
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command12 
      BackColor       =   &H80000010&
      Caption         =   "CONTI&NUAR BUSQ."
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command11 
      Caption         =   "IMPRIMIR &EXTERNO"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   17
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdimprimir 
      Caption         =   "IMPRIMIR &INTERNO"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   1935
   End
   Begin ComctlLib.StatusBar msgestado 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   45
      Top             =   7935
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   5345
            Text            =   "ULTIMA FICHA CREADA O MODIFICADA:"
            TextSave        =   "ULTIMA FICHA CREADA O MODIFICADA:"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Trae la ultima ficha creada, modificada o impresa."
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdprintpublic 
      Caption         =   "IMPRIMIR &PUBLICO"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   18
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H000000C0&
      Caption         =   "&CANCELAR"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      MaskColor       =   &H000000C0&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdgrabar 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "&GRABAR"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H8000000E&
      Height          =   10095
      Left            =   -120
      TabIndex        =   41
      Top             =   360
      Width           =   16000
      Begin VB.Frame Frame3 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   4695
         Left            =   240
         TabIndex        =   96
         Top             =   3000
         Width           =   11895
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "ANULADA"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   12
         Left            =   4320
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "PRESUP"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   480
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Tecnico x Estado"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9720
         TabIndex        =   93
         Top             =   7080
         Width           =   2175
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1200
         TabIndex        =   44
         Top             =   1560
         Width           =   3255
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Tecnico x Presupuesto"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9720
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   6600
         Width           =   2175
      End
      Begin VB.TextBox tce1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   360
         Left            =   8280
         TabIndex        =   76
         TabStop         =   0   'False
         Text            =   "100"
         Top             =   6480
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Buscar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   7080
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdnuevo 
         Caption         =   "&Nueva Ficha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10080
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1200
         TabIndex        =   46
         Top             =   2160
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1200
         TabIndex        =   43
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox bficha 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1200
         TabIndex        =   42
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox tce2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   360
         Left            =   11160
         TabIndex        =   80
         TabStop         =   0   'False
         Text            =   "100"
         Top             =   6000
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   360
         Left            =   6600
         TabIndex        =   75
         TabStop         =   0   'False
         Text            =   "Tecnico"
         Top             =   4560
         Width           =   2415
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Buscar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   7080
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   375
         Left            =   5400
         TabIndex        =   90
         TabStop         =   0   'False
         Text            =   "  100"
         Top             =   6480
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Buscar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   67
         TabStop         =   0   'False
         Tag             =   "POR VER"
         Top             =   5160
         Width           =   1695
      End
      Begin VB.TextBox tce 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   360
         Left            =   5400
         TabIndex        =   66
         TabStop         =   0   'False
         Text            =   "100"
         Top             =   4560
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "REP.EXT."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   4560
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "ENTREGADA"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   4560
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "STD/BY"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "LISTA"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "REPARANDO"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   480
         MaskColor       =   &H00808080&
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   4320
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "POR VER"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   480
         MaskColor       =   &H00808080&
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   3600
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "CONFIRMADO"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   86
         Top             =   6480
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "NO CONFIRMADO"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   87
         Top             =   6720
         Width           =   2775
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "NO DISPONIBLE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   92
         Top             =   6960
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "CHEQUEO"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   480
         MaskColor       =   &H00808080&
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   4800
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "PV EXT."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "LISTA NR"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   2520
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "LISTA BRGS"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   2520
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   4320
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "ENTREGAR"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   11
         Left            =   2520
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   4800
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "DEPOSITO"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   14
         Left            =   2520
         MaskColor       =   &H00808080&
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   5040
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "DIAGNOSTIC"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   13
         Left            =   480
         MaskColor       =   &H00808080&
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Label Label36 
         BackColor       =   &H00808080&
         Caption         =   "Mostrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   91
         Top             =   6600
         Width           =   735
      End
      Begin VB.Line Line44 
         X1              =   6120
         X2              =   6120
         Y1              =   6360
         Y2              =   6960
      End
      Begin VB.Line Line43 
         X1              =   4440
         X2              =   6120
         Y1              =   6960
         Y2              =   6960
      End
      Begin VB.Line Line42 
         X1              =   4440
         X2              =   6120
         Y1              =   6360
         Y2              =   6360
      End
      Begin VB.Line Line41 
         X1              =   4440
         X2              =   4440
         Y1              =   6360
         Y2              =   6960
      End
      Begin VB.Label Label35 
         BackColor       =   &H00808080&
         Caption         =   "Seleccione un Tecnico para la busqueda."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   6600
         TabIndex        =   89
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Image Image8 
         Height          =   480
         Left            =   6480
         Picture         =   "frmMain.frx":0442
         Top             =   3240
         Width           =   480
      End
      Begin VB.Label Label34 
         BackColor       =   &H00808080&
         Caption         =   "Busqueda Por Tecnico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   7080
         TabIndex        =   85
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Line Line40 
         X1              =   9120
         X2              =   9120
         Y1              =   3000
         Y2              =   7560
      End
      Begin VB.Line Line39 
         X1              =   6360
         X2              =   9120
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line38 
         X1              =   6360
         X2              =   9120
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Line Line37 
         X1              =   6360
         X2              =   6360
         Y1              =   3000
         Y2              =   7560
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   480
         Picture         =   "frmMain.frx":0884
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label33 
         BackColor       =   &H00808080&
         Caption         =   "Seleccione un Estado y un Tecnico o un Presupuesto y un Tecnico luego presione el boton correspondiente."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   9480
         TabIndex        =   82
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Line Line36 
         X1              =   11880
         X2              =   11880
         Y1              =   5880
         Y2              =   6480
      End
      Begin VB.Line Line35 
         X1              =   10200
         X2              =   10200
         Y1              =   5880
         Y2              =   6480
      End
      Begin VB.Line Line34 
         X1              =   10200
         X2              =   11880
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line Line33 
         X1              =   10200
         X2              =   11880
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Label Label32 
         BackColor       =   &H00808080&
         Caption         =   "Mostrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   10320
         TabIndex        =   81
         Top             =   6120
         Width           =   735
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   9360
         Picture         =   "frmMain.frx":0CC6
         Top             =   3240
         Width           =   480
      End
      Begin VB.Label Label31 
         BackColor       =   &H00808080&
         Caption         =   "Busqueda Relacionada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   9960
         TabIndex        =   79
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Line Line32 
         X1              =   9240
         X2              =   12000
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Line Line31 
         X1              =   12000
         X2              =   12000
         Y1              =   3000
         Y2              =   7560
      End
      Begin VB.Line Line30 
         X1              =   9240
         X2              =   12000
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line29 
         X1              =   9240
         X2              =   9240
         Y1              =   3000
         Y2              =   7560
      End
      Begin VB.Label Label30 
         BackColor       =   &H00808080&
         Caption         =   "Mostrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7440
         TabIndex        =   77
         Top             =   6600
         Width           =   735
      End
      Begin VB.Line Line20 
         X1              =   9000
         X2              =   9000
         Y1              =   6360
         Y2              =   6960
      End
      Begin VB.Line Line19 
         X1              =   7320
         X2              =   9000
         Y1              =   6960
         Y2              =   6960
      End
      Begin VB.Line Line18 
         X1              =   7320
         X2              =   7320
         Y1              =   6360
         Y2              =   6960
      End
      Begin VB.Line Line17 
         X1              =   7320
         X2              =   9000
         Y1              =   6360
         Y2              =   6360
      End
      Begin VB.Line Line16 
         X1              =   4440
         X2              =   6120
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line15 
         X1              =   4440
         X2              =   4440
         Y1              =   4440
         Y2              =   5040
      End
      Begin VB.Line Line14 
         X1              =   4440
         X2              =   6120
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line13 
         X1              =   6120
         X2              =   6120
         Y1              =   4440
         Y2              =   5040
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   480
         Picture         =   "frmMain.frx":1108
         Top             =   5880
         Width           =   480
      End
      Begin VB.Label Label11 
         BackColor       =   &H00808080&
         Caption         =   "Busqueda Por Presupuesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   74
         Top             =   6000
         Width           =   3735
      End
      Begin VB.Line Line12 
         X1              =   12000
         X2              =   12000
         Y1              =   120
         Y2              =   2760
      End
      Begin VB.Line Line11 
         X1              =   240
         X2              =   12000
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line10 
         X1              =   240
         X2              =   12000
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line9 
         X1              =   240
         X2              =   240
         Y1              =   120
         Y2              =   2760
      End
      Begin VB.Line Line8 
         X1              =   240
         X2              =   6240
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Line Line7 
         X1              =   240
         X2              =   6240
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Line Line6 
         X1              =   6240
         X2              =   6240
         Y1              =   5640
         Y2              =   7560
      End
      Begin VB.Line Line5 
         X1              =   240
         X2              =   240
         Y1              =   5640
         Y2              =   7560
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         Caption         =   "Mostrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   73
         Top             =   4680
         Width           =   735
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   240
         Y1              =   3000
         Y2              =   5640
      End
      Begin VB.Line Line2 
         X1              =   6240
         X2              =   6240
         Y1              =   3000
         Y2              =   5640
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   480
         Picture         =   "frmMain.frx":154A
         Top             =   3120
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   480
         Picture         =   "frmMain.frx":198C
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   480
         Picture         =   "frmMain.frx":1DCE
         Top             =   2040
         Width           =   480
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   6240
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label29 
         BackColor       =   &H00808080&
         Caption         =   "Busqueda Por Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   68
         Top             =   3240
         Width           =   2775
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   480
         Picture         =   "frmMain.frx":2210
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   6975
      Left            =   120
      TabIndex        =   26
      Top             =   600
      Width           =   11655
      Begin VB.TextBox tproblema 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   2400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   4200
         Width           =   8655
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   360
         ItemData        =   "frmMain.frx":2652
         Left            =   8880
         List            =   "frmMain.frx":2654
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   2760
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Trabajos Reali&zados"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   6240
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.TextBox tdireccion 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   2400
         TabIndex        =   1
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox tconfirmacion 
         BackColor       =   &H80000008&
         BorderStyle     =   0  'None
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   6240
         Width           =   2655
      End
      Begin VB.TextBox Tavisadopor 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   2400
         TabIndex        =   3
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox Tavisadoeldia 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   6600
         TabIndex        =   12
         Top             =   3720
         Width           =   1575
      End
      Begin VB.TextBox Tllamareldia 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   2400
         TabIndex        =   8
         Top             =   3720
         Width           =   2415
      End
      Begin VB.TextBox temail 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   6600
         TabIndex        =   4
         Top             =   2280
         Width           =   4455
      End
      Begin VB.TextBox tadjuntos 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   2400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   5400
         Width           =   3615
      End
      Begin VB.TextBox tnserie 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   6600
         TabIndex        =   7
         Top             =   3240
         Width           =   4455
      End
      Begin VB.TextBox tmodelo 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   2400
         TabIndex        =   6
         Top             =   3240
         Width           =   2415
      End
      Begin VB.TextBox tsolucion 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   2400
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   6480
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.TextBox ttecnico 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   6600
         TabIndex        =   11
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox tatendidopor 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   2400
         TabIndex        =   5
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox tprecio 
         BackColor       =   &H80000008&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C0C0C0&
         Height          =   345
         Left            =   9000
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox tpresupuesto 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   8400
         TabIndex        =   14
         Top             =   5280
         Width           =   2655
      End
      Begin VB.TextBox ttelefono 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   2400
         TabIndex        =   2
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox tfechaegreso 
         BackColor       =   &H80000008&
         BorderStyle     =   0  'None
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox tfechaingreso 
         BackColor       =   &H80000008&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox tfullname 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   2400
         TabIndex        =   0
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox tficha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000008&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox Tcontroladopor 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   9600
         TabIndex        =   13
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Line Line28 
         BorderColor     =   &H80000004&
         X1              =   11640
         X2              =   11640
         Y1              =   0
         Y2              =   6960
      End
      Begin VB.Line Line27 
         BorderColor     =   &H80000004&
         X1              =   0
         X2              =   11640
         Y1              =   6960
         Y2              =   6960
      End
      Begin VB.Line Line26 
         BorderColor     =   &H8000000A&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   6960
      End
      Begin VB.Line Line25 
         BorderColor     =   &H80000004&
         X1              =   0
         X2              =   11640
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line24 
         BorderColor     =   &H80000004&
         X1              =   11040
         X2              =   11040
         Y1              =   240
         Y2              =   2160
      End
      Begin VB.Line Line23 
         BorderColor     =   &H80000004&
         X1              =   7200
         X2              =   11040
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line22 
         BorderColor     =   &H80000004&
         X1              =   7200
         X2              =   11040
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line21 
         BorderColor     =   &H80000004&
         X1              =   7200
         X2              =   7200
         Y1              =   240
         Y2              =   2160
      End
      Begin VB.Label Label28 
         BackColor       =   &H00000000&
         Caption         =   "entregada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   7440
         TabIndex        =   61
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label27 
         BackColor       =   &H00000000&
         Caption         =   "fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   6480
         TabIndex        =   59
         Top             =   6360
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "NO DISPONIBLE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   8400
         TabIndex        =   57
         Top             =   5760
         Width           =   2655
      End
      Begin VB.Label Label26 
         BackColor       =   &H00000000&
         Caption         =   "direccion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label25 
         BackColor       =   &H00000000&
         Caption         =   "estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   6480
         TabIndex        =   54
         Top             =   5880
         Width           =   1335
      End
      Begin VB.Label Label24 
         BackColor       =   &H00000000&
         Caption         =   "control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   8280
         TabIndex        =   52
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label23 
         BackColor       =   &H00000000&
         Caption         =   "telefono alt."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label22 
         BackColor       =   &H00000000&
         Caption         =   "ubicacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   7440
         TabIndex        =   50
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackColor       =   &H00000000&
         Caption         =   "avisado el dia por"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   5160
         TabIndex        =   49
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label20 
         BackColor       =   &H00000000&
         Caption         =   "llamar el dia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label19 
         BackColor       =   &H00000000&
         Caption         =   "e-mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   5160
         TabIndex        =   47
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Caption         =   "tecnico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   5160
         TabIndex        =   35
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackColor       =   &H00000000&
         Caption         =   "adjuntos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   5520
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "no. serie"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   5160
         TabIndex        =   39
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "modelo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label15 
         BackColor       =   &H00000000&
         Caption         =   "descripcion del problema"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   240
         TabIndex        =   37
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label lbllista 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         Caption         =   " LISTA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   9000
         TabIndex        =   36
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         Caption         =   "atendido por"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "presupuesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   6480
         TabIndex        =   33
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "telefono"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "nombre/apellido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "finalizada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   7440
         TabIndex        =   30
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "ingreso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7440
         TabIndex        =   29
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   " nro. ficha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "&Archivo"
      Index           =   2
      Begin VB.Menu mnusalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnubusqueda 
      Caption         =   "&Busqueda"
      Begin VB.Menu mnubavanzada 
         Caption         =   "A&vanzada"
      End
   End
   Begin VB.Menu mnuherramientas 
      Caption         =   "&Herramientas"
      Begin VB.Menu mnuconfig 
         Caption         =   "&Configuracion"
      End
      Begin VB.Menu mnuexportar 
         Caption         =   "&Exportar a .CSV"
      End
      Begin VB.Menu Repext 
         Caption         =   "&Lista de Rep. Ext."
      End
      Begin VB.Menu utilizardatos 
         Caption         =   "&Utilizar Datos"
      End
   End
   Begin VB.Menu mnuayuda 
      Caption         =   "A&cerca de.."
      Begin VB.Menu mnuAcerca 
         Caption         =   "FluWork"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bficha_GotFocus()

bficha.BackColor = QBColor(7)
Text2.BackColor = QBColor(15)
Text1.BackColor = QBColor(15)
Text3.BackColor = QBColor(15)
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

End Sub

Private Sub bficha_KeyPress(KeyAscii As Integer)

If KeyAscii <> 8 And KeyAscii <> 13 And Len(bficha.Text) > 9 Then KeyAscii = 0
If (KeyAscii = 13) And Val(Trim(bficha.Text)) > 0 Then Command1.Value = True
If Not ((KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub cmdcancel_Click()

If MsgBox("�Confirma que quiere cancelar?", vbYesNo, "Confirmacion") = vbNo Then Exit Sub
mnubusqueda.Enabled = True
mnuexportar.Enabled = True
Form1.utilizardatos.Enabled = False
Frame1.Visible = False
tconfirmacion.Enabled = False
tdireccion.Enabled = False
temail.Enabled = False
Tllamareldia.Enabled = False
Tcontroladopor.Enabled = False
Tavisadoeldia.Enabled = False
Tavisadopor.Enabled = False
tficha.Enabled = False
tfullname.Enabled = False
tprecio.Enabled = False
tpresupuesto.Enabled = False
tproblema.Enabled = False
tsolucion.Enabled = False
ttelefono.Enabled = False
tfechaingreso.Enabled = False
tfechaegreso.Enabled = False
ttecnico.Enabled = False
tatendidopor.Enabled = False
Command12.Visible = False
Command12.Enabled = False
cmdgrabar.Visible = False
cmdgrabar.Enabled = False
cmdimprimir.Enabled = False
cmdimprimir.Visible = False
cmdprintpublic.Enabled = False
cmdprintpublic.Visible = False
Command11.Enabled = False
Command11.Visible = False
cmdcancel.Enabled = False
cmdcancel.Visible = False
Frame2.Visible = True
tficha.Text = ""
tfullname.Text = ""
tprecio.Text = ""
tpresupuesto.Text = ""
tproblema.Text = ""
tsolucion.Text = ""
ttelefono.Text = ""
tfechaingreso.Text = ""
tfechaegreso.Text = ""
ttecnico.Text = ""
tatendidopor.Text = ""
tnserie.Text = ""
tmodelo.Text = ""
tconfirmacion.Text = ""
tdireccion.Text = ""
temail.Text = ""
Tllamareldia.Text = ""
Tcontroladopor.Text = ""
Tavisadoeldia.Text = ""
Tavisadopor.Text = ""
Label27.Visible = True
tconfirmacion.Visible = True
cmdnuevo.Visible = True
cmdnuevo.Enabled = True
tmpficha = 0
touchedreally = False
If showres = True Then
Form3.Show
End If
End Sub

Private Sub cmdgrabar_Click()


If tficha.Text = "" Then
registro.ficha = getlastfichanumber() + 1
Else
registro.ficha = tficha.Text
End If
mnubusqueda.Enabled = True
mnuexportar.Enabled = True
Form1.utilizardatos.Enabled = False

msgestado.Panels(2).Text = registro.ficha
tficha.Text = registro.ficha
registro.fullname = tfullname.Text
registro.precio = tprecio.Text
registro.presupuesto = tpresupuesto.Text
registro.problema = tproblema.Text
registro.solucion = tsolucion.Text
registro.telefono = ttelefono.Text
registro.adjuntos = tadjuntos.Text
registro.fechaingreso = tfechaingreso.Text
registro.fechaegreso = tfechaegreso.Text
registro.tecnico = ttecnico.Text
registro.atendidopor = tatendidopor.Text
registro.estado = lbllista.Caption
registro.modelo = tmodelo.Text
registro.nserie = tnserie.Text
registro.confirmacion = tconfirmacion.Text
registro.direccion = tdireccion.Text
registro.email = temail.Text
registro.llamareldia = Tllamareldia.Text
registro.controladopor = Tcontroladopor.Text
registro.avisadoeldia = Tavisadoeldia.Text
registro.avisadopor = Tavisadopor.Text

Dim archivo As String
archivo = Trim(dbpath + "\" + Trim(tficha.Text))

Open archivo For Random As #1 Len = Len(registro)
Put #1, , registro
Close #1

regindex.ficha = registro.ficha
regindex.fullname = StrConv(registro.fullname, 1)
regindex.telefono = registro.telefono
regindex.modelo = registro.modelo
regindex.fecha = registro.fechaingreso
regindex.estado = registro.estado
regindex.tecnico = registro.tecnico
regindex.confirmacion = registro.confirmacion

Open Trim(dbpath + "\index.dat") For Random As #1 Len = Len(regindex)
Put #1, Val(registro.ficha) - BASE, regindex
Close #1

Command12.Visible = False
Command12.Enabled = False
cmdgrabar.Visible = False
cmdgrabar.Enabled = False
cmdimprimir.Enabled = False
cmdimprimir.Visible = False
cmdprintpublic.Enabled = False
cmdprintpublic.Visible = False
Command11.Enabled = False
Command11.Visible = False
cmdcancel.Enabled = False
cmdcancel.Visible = False
Frame1.Visible = False
Frame2.Visible = True
cmdnuevo.Visible = True
cmdnuevo.Enabled = True
tmpficha = 0
If showres = True Then
Form3.Show
End If
End Sub

Private Sub cmdimprimir_Click()

On Error Resume Next
Dim returnvalue, i
showres = False

If tficha.Text = "" Then
registro.ficha = getlastfichanumber() + 1
Else
registro.ficha = tficha.Text
End If
tficha.Text = Trim(registro.ficha)
mnubusqueda.Enabled = True
mnuexportar.Enabled = True
Form1.utilizardatos.Enabled = False

msgestado.Panels(2).Text = Trim(registro.ficha)
registro.fullname = Trim(tfullname.Text)
registro.precio = Trim(tprecio.Text)
registro.presupuesto = Trim(tpresupuesto.Text)
registro.problema = Trim(tproblema.Text)
registro.solucion = Trim(tsolucion.Text)
registro.telefono = Trim(ttelefono.Text)
registro.adjuntos = Trim(tadjuntos.Text)
registro.fechaingreso = Trim(tfechaingreso.Text)
registro.fechaegreso = Trim(tfechaegreso.Text)
registro.tecnico = Trim(ttecnico.Text)
registro.atendidopor = Trim(tatendidopor.Text)
registro.estado = Trim(lbllista.Caption)
registro.modelo = Trim(tmodelo.Text)
registro.nserie = Trim(tnserie.Text)
registro.confirmacion = Trim(tconfirmacion.Text)
registro.direccion = Trim(tdireccion.Text)
registro.email = Trim(temail.Text)
registro.llamareldia = Trim(Tllamareldia.Text)
registro.controladopor = Trim(Tcontroladopor.Text)
registro.avisadoeldia = Trim(Tavisadoeldia.Text)
registro.avisadopor = Trim(Tavisadopor.Text)

Dim archivo As String
archivo = Trim(dbpath + "\" + Trim(tficha.Text))

Open archivo For Random As #1 Len = Len(registro)
Put #1, , registro
Close #1

regindex.ficha = registro.ficha
regindex.fullname = StrConv(registro.fullname, 1)
regindex.telefono = registro.telefono
regindex.modelo = registro.modelo
regindex.fecha = registro.fechaingreso
regindex.estado = registro.estado
regindex.tecnico = registro.tecnico
regindex.confirmacion = registro.confirmacion

Open Trim(dbpath + "\index.dat") For Random As #1 Len = Len(regindex)
Put #1, Val(registro.ficha) - BASE, regindex
Close #1

'Sleep 1000
Kill "c:\laficha.html"
'Sleep 1000

Close #7

Open "c:\laficha.html" For Output As #7

Print #7, "<html><head><title></title></head><body bgcolor=""#FFFFFF"">"

'BLOQUE IMAGENES

Print #7, "<table border=""0"" cellpadding=""0"" cellspacing=""1"" width=""100%""><tr><td width=""25%"" rowspan=""3""><center><IMG border=""0"" SRC=""file:"
Print #7, Trim(dbpath) + "\mag.logo.gif"
Print #7, """ width=""190"" height=""159""></center></td><td width=""25%"" valign=""top"" align=""left"" height=""2""><img border=""0"" src=""file:"
Print #7, Trim(dbpath) + "\direccion.gif"
Print #7, """ width=""200"" height=""69""></td><td width=""25%"" valign=""middle"" align=""center""><p align=""right""><img border=""0"" src=""file:"
Print #7, Trim(dbpath) + "\epson.logo.gif"
Print #7, """ width=""128"" height=""40""></td><td width=""25%"" valign=""middle"" align=""center""><img border=""0"" src=""file:"
Print #7, Trim(dbpath) + "\hp.logo.gif"

Print #7, """ width=""111"" height=""115"">"
Print #7, "</td></tr><tr><td width=""75%"" valign=""top"" align=""left"" height=""1"" colspan=""3""><p align=""center""><b><font size=""4"">FICHA DE SERVICIO TECNICO</font></b></p>"
Print #7, "</td></tr><tr><td width=""75%"" valign=""top"" align=""left"" colspan=""3"">&nbsp;</td></tr></table>"

'FIN BLOQUE IMAGENES

Print #7, "<table border=""2"" width=""100%"" cellpadding=""0"" cellspacing=""5"" bordercolor=""#000000"" bordercolordark=""#000000"" bordercolorlight=""#000000""><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">FECHA RECEPCION:</font></b></td>"
Print #7, "<td width=""25%"" align=""center"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"

If (Trim(registro.fechaingreso) <> "") Then
Print #7, registro.fechaingreso
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If

Print #7, "</td><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2""></font></b><p align=""center""><font size=""2""><b>N� CONTROL:</b></font></td>"
Print #7, "<td width=""25%"" align=""right"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000""><b><font size=""5"">"

Print #7, registro.ficha

Print #7, "</font></b></td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">NOMBRE Y APELLIDO:</font></b></td><td width=""75%"" colspan=""3"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"

Print #7, registro.fullname

Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">DIRECCION:</font></b></td>"
Print #7, "<td width=""75%"" colspan=""3"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"

If (Trim(registro.direccion) <> "") Then
Print #7, registro.direccion
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">TELEFONO:</font></b></td><td width=""25%"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"

If (Trim(registro.telefono) <> "") Then
Print #7, registro.telefono
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "</td><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><p align=""center""><b><font size=""2"">E-MAIL:</font></b></td><td width=""30%"">"

If (Trim(registro.email) <> "") Then
Print #7, registro.email
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">DESCRIPCION DEL</font></b></td>"
Print #7, "<td width=""75%"" colspan=""3"" rowspan=""3"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000""><b>MODELO: </b>"

If (Trim(registro.modelo) <> "") Then
Print #7, registro.modelo
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "<br><b>ADJUNTOS:</b>"

If (Trim(registro.adjuntos) <> "") Then
Print #7, registro.adjuntos
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "<BR></td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">EQUIPO:</font></b></td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td></tr><tr>"
Print #7, "<td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">N</font></b><font size=""2""><b>� DE </b></font><b><font size=""2"">SERIE :</font></b></td><td width=""25%"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"

If (Trim(registro.nserie) <> "") Then
Print #7, registro.nserie
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "</td><td width=""55%"" colspan=""2"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"" valign=""top"">"
Print #7, "<p align=""center""><b><font size=""2""><sup>(datos informados por el cliente a confirmar)</sup></font></b>"
Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">DESCRIPCION DEL</font></b></td><td width=""75%"" colspan=""3"" rowspan=""5"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"

If (Trim(registro.problema) <> "") Then
Print #7, enter_a_br(registro.problema)
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">PROBLEMA:</font></b></td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td>"
Print #7, "</tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td></tr>"
Print #7, "<tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">LLAMAR"
Print #7, "EL DIA:</font></b></td><td width=""25%"" align=""center"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"
          
If (Trim(registro.llamareldia) <> "") Then
Print #7, registro.llamareldia
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If

          
Print #7, "</td><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><p align=""center""><b><font size=""2"">ATENDIDO POR:</font></b></p></td><td width=""25%"" align=""center"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"

If (Trim(registro.atendidopor) <> "") Then
Print #7, registro.atendidopor
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "</td></tr>"
Print #7, "<td width=""25%"" align=""center"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;"
Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">TAREAS REALIZADAS:</font></b></td><td width=""75%"" colspan=""3"" rowspan=""5"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"

If (Trim(registro.solucion) <> "") Then
Print #7, pipe_a_br(registro.solucion)
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td></tr>"
Print #7, "<tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td></tr>"
Print #7, "<tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">UBICACION:</font></b></td><td width=""25%"" align=""center"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"
          
If (Trim(registro.estado) <> "") Then
Print #7, registro.estado
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If

          
Print #7, "</td><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">  <p align=""center""><b><font size=""2"">CONTROLO:</font></b></p></td><td width=""25%"" align=""center"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"

If (Trim(registro.controladopor) <> "") Then
Print #7, registro.controladopor
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">PRESUPUESTO:</font></b></td><td width=""25%"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000""><p align=""center"">"

If (Trim(registro.presupuesto) <> "") Then
Print #7, registro.presupuesto
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "</p></td><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td><td width=""25%"" align=""center"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">AVISADO EL DIA POR:</font></b></td>"
Print #7, "<td width=""25%"" align=""center"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"
          
If (Trim(registro.avisadoeldia) <> "") Then
Print #7, registro.avisadoeldia
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If

         

Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">CONFIRMACION:</font></b></td><td width=""25%"" align=""center"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"
          
If (Trim(registro.confirmacion) <> "") Then
Print #7, registro.confirmacion
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If

          
Print #7, "</td><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><p align=""center""><b><font size=""2"">RETIRO EL DIA:</font></b></p></td><td width=""25%"" align=""center"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"

If (Trim(registro.fechaegreso) <> "") Then
Print #7, registro.fechaegreso
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If

Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">RETIRADO POR: </font></b></td>"
Print #7, "<td width=""75%"" colspan=""3"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"
Print #7, "<font size=""7""><font color=""#FFFFFF"">.</font>"

Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">FIRMA:</font></b></td><td width=""25%"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"
Print #7, "<font size=""7""><font color=""#FFFFFF"">.</font>"


Print #7, "</td></tr></table></body></html>"

Close #7

returnvalue = Shell("c:\Archivos de Programa\Internet Explorer\iexplore.exe C:\laficha.html", vbMaximizedFocus)

If returnvalue = 0 Then
returnvalue = Shell("start c:\laficha.html")
If returnvalue = 0 Then MsgBox "Ha ocurrido un error al intentar cargar la ficha en el navegador", vbExclamation, "Error Internet Explorer"
End If

Command12.Visible = False
Command12.Enabled = False
cmdgrabar.Visible = False
cmdgrabar.Enabled = False
cmdimprimir.Enabled = False
cmdimprimir.Visible = False
cmdcancel.Enabled = False
cmdcancel.Visible = False
cmdprintpublic.Enabled = False
cmdprintpublic.Visible = False
Command11.Enabled = False
Command11.Visible = False
Frame1.Visible = False
Frame2.Visible = True
cmdnuevo.Visible = True
cmdnuevo.Enabled = True
tmpficha = 0

End Sub

Private Sub cmdnuevo_Click()

showres = False
tmpficha = 0
Frame2.Visible = False
cmdnuevo.Visible = False
cmdnuevo.Enabled = False
Combo1.ListIndex = 0
mnubusqueda.Enabled = False
mnuexportar.Enabled = False
Form1.utilizardatos.Enabled = False

Frame1.Visible = True

tconfirmacion.Enabled = True
tdireccion.Enabled = True
temail.Enabled = True
Tllamareldia.Enabled = True
Tcontroladopor.Enabled = True
Tavisadoeldia.Enabled = True
Tavisadopor.Enabled = True
tficha.Enabled = True
tfullname.Enabled = True
tprecio.Enabled = True
tpresupuesto.Enabled = True
tproblema.Enabled = True
tadjuntos.Enabled = True
tsolucion.Enabled = True
ttelefono.Enabled = True

tfechaingreso.Enabled = True
tfechaegreso.Enabled = True

tnserie.Enabled = True
tmodelo.Enabled = True
ttecnico.Enabled = True
tatendidopor.Enabled = True
Command12.Enabled = False
Command12.Visible = False
cmdgrabar.Visible = True
cmdgrabar.Enabled = True
cmdcancel.Enabled = True
cmdcancel.Visible = True
cmdimprimir.Visible = True
cmdimprimir.Enabled = True
cmdprintpublic.Enabled = True
cmdprintpublic.Visible = True
Command11.Enabled = True
Command11.Visible = True

tficha.Text = ""
tfullname.Text = ""
tprecio.Text = ""
tpresupuesto.Text = ""
tproblema.Text = ""
tsolucion.Text = ""
tadjuntos.Text = ""
ttelefono.Text = ""
tfechaingreso.Text = Trim(str(Date))
tfechaegreso.Text = ""
ttecnico.Text = ""
tnserie.Text = ""
tmodelo.Text = ""
tatendidopor.Text = ""
tfullname.SetFocus
lbllista.Caption = "POR VER"
lbllista.ForeColor = QBColor(11)
tconfirmacion.Text = ""
tdireccion.Text = ""
temail.Text = ""
Tllamareldia.Text = ""
Tcontroladopor.Text = ""
Tavisadoeldia.Text = ""
Tavisadopor.Text = ""
Frame1.Width = 11660
Frame1.Height = 6990

End Sub

Private Sub cmdprintpublic_Click()

On Error Resume Next
Dim returnvalue, i
showres = False

If tficha.Text = "" Then
registro.ficha = getlastfichanumber() + 1
Else
registro.ficha = tficha.Text
End If
mnubusqueda.Enabled = True
mnuexportar.Enabled = True
Form1.utilizardatos.Enabled = False

tficha.Text = registro.ficha
msgestado.Panels(2).Text = registro.ficha
registro.fullname = tfullname.Text
registro.precio = tprecio.Text
registro.presupuesto = tpresupuesto.Text
registro.problema = tproblema.Text
registro.solucion = tsolucion.Text
registro.telefono = ttelefono.Text
registro.adjuntos = tadjuntos.Text
registro.fechaingreso = tfechaingreso.Text
registro.fechaegreso = tfechaegreso.Text
registro.tecnico = ttecnico.Text
registro.atendidopor = tatendidopor.Text
registro.estado = lbllista.Caption
registro.modelo = tmodelo.Text
registro.nserie = tnserie.Text
registro.confirmacion = tconfirmacion.Text
registro.direccion = tdireccion.Text
registro.email = temail.Text
registro.llamareldia = Tllamareldia.Text
registro.controladopor = Tcontroladopor.Text
registro.avisadoeldia = Tavisadoeldia.Text
registro.avisadopor = Tavisadopor.Text

Dim archivo As String
archivo = Trim(dbpath + "\" + Trim(tficha.Text))

Open archivo For Random As #1 Len = Len(registro)
Put #1, , registro
Close #1

regindex.ficha = registro.ficha
regindex.fullname = StrConv(registro.fullname, 1)
regindex.telefono = registro.telefono
regindex.modelo = registro.modelo
regindex.fecha = registro.fechaingreso
regindex.estado = registro.estado
regindex.tecnico = registro.tecnico
regindex.confirmacion = registro.confirmacion

Open Trim(dbpath + "\index.dat") For Random As #1 Len = Len(regindex)
Put #1, Val(registro.ficha) - BASE, regindex
Close #1

'Sleep 1000
Kill "c:\laficha.html"
'Sleep 1000
Close #7

Open "c:\laficha.html" For Output As #7

Print #7, "<html><head><title></title></head><body bgcolor=""#FFFFFF"">"

'BLOQUE IMAGENES

Print #7, "<table border=""0"" cellpadding=""0"" cellspacing=""1"" width=""100%""><tr><td width=""25%"" rowspan=""3""><center><IMG border=""0"" SRC=""file:"
Print #7, Trim(dbpath) + "\mag.logo.gif"
Print #7, """ width=""190"" height=""159""></center></td><td width=""25%"" valign=""top"" align=""left"" height=""2""><img border=""0"" src=""file:"
Print #7, Trim(dbpath) + "\direccion.gif"
Print #7, """ width=""200"" height=""69""></td><td width=""25%"" valign=""middle"" align=""center""><p align=""right""><img border=""0"" src=""file:"
Print #7, Trim(dbpath) + "\epson.logo.gif"
Print #7, """ width=""128"" height=""40""></td><td width=""25%"" valign=""middle"" align=""center""><img border=""0"" src=""file:"
Print #7, Trim(dbpath) + "\hp.logo.gif"
Print #7, """ width=""111"" height=""115"">"
Print #7, "</td></tr><tr><td width=""75%"" valign=""top"" align=""left"" height=""1"" colspan=""3""><p align=""center""><b><font size=""4"">FICHA DE SERVICIO TECNICO</font></b></p>"
Print #7, "</td></tr><tr><td width=""75%"" valign=""top"" align=""left"" colspan=""3"">&nbsp;</td></tr></table>"

'FIN BLOQUE IMAGENES

Print #7, "<table border=""2"" width=""100%"" cellpadding=""0"" cellspacing=""5"" bordercolor=""#000000"" bordercolordark=""#000000"" bordercolorlight=""#000000""><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">FECHA RECEPCION:</font></b></td>"
Print #7, "<td width=""25%"" align=""center"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"


If (Trim(registro.fechaingreso) <> "") Then
Print #7, registro.fechaingreso
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If

Print #7, "</td><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2""></font></b>"
Print #7, "<p align=""center""><font size=""2""><b>N� CONTROL:</b></font></td><td width=""25%"" align=""right"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000""><b><font size=""5"">"

If (Trim(registro.ficha) <> "") Then
Print #7, registro.ficha
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "</font></b></td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">NOMBRE Y APELLIDO:</font></b></td>"
Print #7, "<td width=""75%"" colspan=""3"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"

If (Trim(registro.fullname) <> "") Then
Print #7, registro.fullname
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">DIRECCION:</font></b></td>"
Print #7, "<td width=""75%"" colspan=""3"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"

If (Trim(registro.direccion) <> "") Then
Print #7, registro.direccion
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">TELEFONO:</font></b></td><td width=""25%"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"

If (Trim(registro.telefono) <> "") Then
Print #7, registro.telefono
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "</td><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><p align=""center""><b><font size=""2"">E-MAIL:</font></b></td><td width=""30%"">"

If (Trim(registro.email) <> "") Then
Print #7, registro.email
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">DESCRIPCION DEL</font></b></td><td width=""75%"" colspan=""3"" rowspan=""3"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000""><b>MODELO: </b>"

If (Trim(registro.modelo) <> "") Then
Print #7, registro.modelo
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "<br><b>ADJUNTOS:</b>"

If (Trim(registro.adjuntos) <> "") Then
Print #7, registro.adjuntos
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "<BR></td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">EQUIPO:</font></b></td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td></tr>"
Print #7, "<tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">N</font></b><font size=""2""><b>�"
Print #7, "DE </b></font><b><font size=""2"">SERIE :</font></b></td><td width=""25%"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"

If (Trim(registro.nserie) <> "") Then
Print #7, registro.nserie
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "</td><td width=""55%"" colspan=""2"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"" valign=""top"">"
Print #7, "<p align=""center""><b><font size=""2""><sup>(datos informados por el cliente a confirmar)</sup></font></b></td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">DESCRIPCION DEL</font></b></td>"
Print #7, "<td width=""75%"" colspan=""3"" rowspan=""5"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"

If (Trim(registro.problema) <> "") Then
Print #7, enter_a_br(registro.problema)
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If

Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">PROBLEMA:</font></b></td></tr><tr>"
Print #7, "<td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td></tr><tr>"
Print #7, "<td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td></tr><tr>"
Print #7, "<td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td></tr><tr>"
Print #7, "<td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">NOTA:"
Print #7, "</font></b></td><td width=""75%"" colspan=""3"" rowspan=""2"" bgcolor=""#C0C0C0"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000""><p align=""left""><font size=""1"">"
Print #7, "<b>El cliente, al realizar cualquier tr�mite relacionado con el equipo en reparaci�n, deber� presentar INDEFECTIBLEMENTE este comprobante.<BR>"
Print #7, "Se informa, que si a los 90 d�as a partir del d�a de la fecha, no procede al retiro del equipo dejado en reparaci�n, se entender� abandono del mismo (art. 2526 del C. Civil). En este caso Magenta Computacion adquiere el dominio del bien,"
Print #7, "quedando facultada para su libre disposici�n, perdiendo el cliente todo derecho a indemnizaci�n o reclamo alguno (art. 2525, 2526 del C. Civil)"
Print #7, "<BR>Se cobrar� un valor de $85,- + I.V.A en concepto de presupuesto y diagn�stico del equipo, s�lo en caso en que el cliente opte por retirar el"
Print #7, "equipo sin efectuar la reparaci�n del mismo.</b></font></td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td></tr><tr>"
Print #7, "<td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">LLAMAR EL DIA:</font></b></td>"
Print #7, "<td width=""25%"" align=""center"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"
          
If (Trim(registro.llamareldia) <> "") Then
Print #7, registro.llamareldia
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If

          
Print #7, "</td><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><p align=""center""><b><font size=""2"">ATENDIDO POR:</font></b></p></td><td width=""25%"" align=""center"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"


If (Trim(registro.atendidopor) <> "") Then
Print #7, registro.atendidopor
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If

Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">FIRMA DEL CLIENTE:</font></b></td><td width=""25%"" rowspan=""2"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">&nbsp;</td>"
Print #7, "<td width=""50%"" colspan=""2"" rowspan=""2"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000""><p align=""center""><font size=""1"">Se deja constancia de la aceptacion por parte del cliente que la empresa no se responsabiliza por la informacion contenida en el equipo.</font></td>"
Print #7, "</tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td></tr>"

Print #7, "<tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">ACLARACION:</font></b></td><td width=""25%"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000""><p align=""center"">&nbsp;</p> </td><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td>"
Print #7, "<td width=""25%"" align=""center"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;"

Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">RETIRADO POR: </font></b></td>"
Print #7, "<td width=""75%"" colspan=""3"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"
Print #7, "<font size=""7""><font color=""#FFFFFF"">.</font>"

Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">FIRMA:</font></b></td><td width=""25%"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"
Print #7, "<font size=""7""><font color=""#FFFFFF"">.</font>"

Print #7, "</td></tr></table></body></html>"

Close #7

returnvalue = Shell("c:\Archivos de Programa\Internet Explorer\iexplore.exe C:\laficha.html", vbMaximizedFocus)

If returnvalue = 0 Then
returnvalue = Shell("start c:\laficha.html")
If returnvalue = 0 Then MsgBox "Ha ocurrido un error al intentar cargar la ficha en el navegador", vbExclamation, "Error Internet Explorer"
End If

Command12.Visible = False
Command12.Enabled = False
cmdgrabar.Visible = False
cmdgrabar.Enabled = False
cmdimprimir.Enabled = False
cmdimprimir.Visible = False
cmdcancel.Enabled = False
cmdcancel.Visible = False
cmdprintpublic.Enabled = False
cmdprintpublic.Visible = False
Command11.Enabled = False
Command11.Visible = False
Frame1.Visible = False
Frame2.Visible = True
cmdnuevo.Visible = True
cmdnuevo.Enabled = True
tmpficha = 0

End Sub

Private Sub Combo1_click()

If Combo1.ListIndex = 0 Then
ttecnico.Text = ""
Else
ttecnico.Text = Combo1.Text
End If
End Sub

Private Sub Command1_Click()

Dim cantdefichas As Long
onlyone = True

Combo1.ListIndex = 0
bficha.BackColor = QBColor(15)
mnubusqueda.Enabled = False
mnuexportar.Enabled = False
Form1.utilizardatos.Enabled = True

Unload Form3

If bficha.Text = "" Then
Form1.mnubusqueda.Enabled = True
Form1.utilizardatos.Enabled = False
Form1.mnuexportar.Enabled = True
Exit Sub
End If


If (Trim(bficha.Text) > Val(str(BASE))) And (Trim(bficha.Text) <= Val(str(getlastfichanumber))) Then
    tmpficha = Val(str(bficha.Text))
    touchedreally = True
    Form1.bficha.Text = ""
    MostrarFicha (tmpficha)
    Else
    mnubusqueda.Enabled = True
    mnuexportar.Enabled = True
    Form1.utilizardatos.Enabled = False
    MsgBox "No hay fichas definidas", vbExclamation, "Busqueda"
    End If


End Sub

Private Sub Command10_Click()

Dim i As Long
Dim cantres As Long
Dim cantdefichas As Long
Dim rsindex As Tindexregistro
Dim b As Integer
onlyone = False


Combo1.ListIndex = 0
Text2.BackColor = QBColor(15)
mnubusqueda.Enabled = False
mnuexportar.Enabled = False

Unload Form3
cantres = 0

Open Trim(dbpath + "\index.dat") For Random As #1 Len = Len(regindex)
cantdefichas = getlastfichanumber - BASE
If (cantdefichas <= 0) Then
MsgBox "No hay fichas definidas", vbExclamation, "Busqueda"

Exit Sub
End If

For i = cantdefichas To 1 Step -1
If i > 0 Then

Get #1, i, rsindex

If InStr(1, Trim(rsindex.estado), Trim(Command5.Tag)) And InStr(1, Trim(rsindex.tecnico), Trim(Combo2.Text)) > 0 Then
tmpficha = Val(str(rsindex.ficha))
Form3.Label1.Caption = cantres + 1
If Trim(rsindex.estado) <> Trim("ANULADA") Then
If (cantres < Trim(Form1.tce2.Text)) Then
    cantres = cantres + 1
    Form3.MSFlexGrid1.AddItem rsindex.ficha
    Form3.MSFlexGrid1.Row = cantres
    Form3.MSFlexGrid1.Col = 0
    Form3.MSFlexGrid1.Text = str(rsindex.ficha)
    Form3.MSFlexGrid1.Col = 1
    Form3.MSFlexGrid1.Text = rsindex.fecha
    Form3.MSFlexGrid1.Col = 2
    Form3.MSFlexGrid1.Text = rsindex.fullname
    Form3.MSFlexGrid1.Col = 3
    Form3.MSFlexGrid1.Text = rsindex.telefono
    Form3.MSFlexGrid1.Col = 4
    Form3.MSFlexGrid1.Text = rsindex.modelo
    Form3.MSFlexGrid1.Col = 5
    Form3.MSFlexGrid1.Text = rsindex.estado
   
Else
    Exit For
End If
End If
End If
End If

Next i

For b = 1 To cantres
Form3.MSFlexGrid1.Row = b
Form3.MSFlexGrid1.Col = 5
If Trim(Form3.MSFlexGrid1.Text) = Trim("POR VER") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(11)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("REPARANDO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(12)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("CHEQUEO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(14)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("ENTREGADA") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(8)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("STD/BY") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(13)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("REP.EXT.") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(12)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA NR") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("PV EXT.") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(11)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA BRGS") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("ENTREGAR") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(9)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("PRESUP") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(15)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("DEPOSITO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(8)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("DIAGNOSTIC") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(14)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If


Next b

If cantres > 0 Then Text2.Text = ""
If cantres > 1 Then
    Form3.mnuprint.Enabled = True
    Form3.MSFlexGrid1.RemoveItem cantres + 1
    Form3.Show vbModal
ElseIf cantres = 1 Then
    touchedreally = True
    onlyone = True
    MostrarFicha (tmpficha)
Else
MsgBox "No hay resultados", vbExclamation, "Busqueda"
Form1.mnubusqueda.Enabled = True
Form1.mnuexportar.Enabled = True
End If
Close #1

End Sub

Private Sub Command11_Click()

On Error Resume Next
Dim returnvalue, i
showres = False

If tficha.Text = "" Then
registro.ficha = getlastfichanumber() + 1
Else
registro.ficha = tficha.Text
End If

mnubusqueda.Enabled = True
mnuexportar.Enabled = True
Form1.utilizardatos.Enabled = False

tficha.Text = registro.ficha
msgestado.Panels(2).Text = registro.ficha
registro.fullname = tfullname.Text
registro.precio = tprecio.Text
registro.presupuesto = tpresupuesto.Text
registro.problema = tproblema.Text
registro.solucion = tsolucion.Text
registro.telefono = ttelefono.Text
registro.adjuntos = tadjuntos.Text
registro.fechaingreso = tfechaingreso.Text
registro.fechaegreso = tfechaegreso.Text
registro.tecnico = ttecnico.Text
registro.atendidopor = tatendidopor.Text
registro.estado = lbllista.Caption
registro.modelo = tmodelo.Text
registro.nserie = tnserie.Text
registro.confirmacion = tconfirmacion.Text
registro.direccion = tdireccion.Text
registro.email = temail.Text
registro.llamareldia = Tllamareldia.Text
registro.controladopor = Tcontroladopor.Text
registro.avisadoeldia = Tavisadoeldia.Text
registro.avisadopor = Tavisadopor.Text

Dim archivo As String
archivo = Trim(dbpath + "\" + Trim(tficha.Text))

Open archivo For Random As #1 Len = Len(registro)
Put #1, , registro
Close #1

regindex.ficha = registro.ficha
regindex.fullname = StrConv(registro.fullname, 1)
regindex.telefono = registro.telefono
regindex.modelo = registro.modelo
regindex.fecha = registro.fechaingreso
regindex.estado = registro.estado
regindex.tecnico = registro.tecnico
regindex.confirmacion = registro.confirmacion

Open Trim(dbpath + "\index.dat") For Random As #1 Len = Len(regindex)
Put #1, Val(registro.ficha) - BASE, regindex
Close #1

'Sleep 1000
Kill "c:\laficha.html"
'Sleep 1000
Close #7

Open "c:\laficha.html" For Output As #7

Print #7, "<html><head><title></title></head><body bgcolor=""#FFFFFF"">"

Print #7, "</td></tr><tr><td width=""75%"" valign=""top"" align=""left"" height=""1"" colspan=""3""><p align=""center""><b><font size=""4"">FICHA DE SERVICIO TECNICO <BR>"
If Trim(Form1.Combo1.Text) = "Seleccionar" Then
Print #7, StrConv(Trim(registro.tecnico), vbUpperCase)
Else
Print #7, StrConv(Trim(Form1.Combo1.Text), vbUpperCase)
End If
Print #7, "</font></b></p>"
Print #7, "</td></tr><tr><td width=""75%"" valign=""top"" align=""left"" colspan=""3"">&nbsp;</td></tr></table>"

Print #7, "<table border=""2"" width=""100%"" cellpadding=""0"" cellspacing=""5"" bordercolor=""#000000"" bordercolordark=""#000000"" bordercolorlight=""#000000""><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">FECHA RECEPCION:</font></b></td>"
Print #7, "<td width=""25%"" align=""center"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"


If (Trim(registro.fechaingreso) <> "") Then
Print #7, registro.fechaingreso
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If

Print #7, "</td><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2""></font></b>"
Print #7, "<p align=""center""><font size=""2""><b>N� CONTROL:</b></font></td><td width=""25%"" align=""right"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000""><b><font size=""5"">"

If (Trim(registro.ficha) <> "") Then
Print #7, registro.ficha
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "</font></b></td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">NOMBRE Y APELLIDO:</font></b></td>"
Print #7, "<td width=""75%"" colspan=""3"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"

If (Trim(registro.fullname) <> "") Then
Print #7, "Magenta Computacion"
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">DIRECCION:</font></b></td>"
Print #7, "<td width=""75%"" colspan=""3"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"

If (Trim(registro.direccion) <> "") Then
Print #7, "Av. Maipu 2062- Olivos -Buenos Aires - Argentina"
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">TELEFONO:</font></b></td><td width=""25%"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"

If (Trim(registro.telefono) <> "") Then
Print #7, "4795-6727"
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "</td><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><p align=""center""><b><font size=""2"">E-MAIL:</font></b></td><td width=""30%"">"

Print #7, "info@magentacomputacion.com.ar"

Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">DESCRIPCION DEL</font></b></td><td width=""75%"" colspan=""3"" rowspan=""3"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000""><b>MODELO: </b>"

If (Trim(registro.modelo) <> "") Then
Print #7, registro.modelo
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "<br><b>ADJUNTOS:</b>"

If (Trim(registro.adjuntos) <> "") Then
Print #7, registro.adjuntos
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "<BR></td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">EQUIPO:</font></b></td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td></tr>"
Print #7, "<tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">N</font></b><font size=""2""><b>�"
Print #7, "DE </b></font><b><font size=""2"">SERIE :</font></b></td><td width=""25%"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"

If (Trim(registro.nserie) <> "") Then
Print #7, registro.nserie
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If


Print #7, "</td><td width=""55%"" colspan=""2"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"" valign=""top"">"
Print #7, "<p align=""center""><b><font size=""2""><sup>(datos informados por el cliente a confirmar)</sup></font></b></td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">DESCRIPCION DEL</font></b></td>"
Print #7, "<td width=""75%"" colspan=""3"" rowspan=""5"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"

If (Trim(registro.problema) <> "") Then
Print #7, enter_a_br(registro.problema)
Else
Print #7, "<font color=""#FFFFFF"">-</font>"
End If

Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">PROBLEMA:</font></b></td></tr><tr>"
Print #7, "<td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td></tr><tr>"
Print #7, "<td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td></tr><tr>"
Print #7, "<td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF"">&nbsp;</td></tr><tr>"

Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">RETIRADO POR: </font></b></td>"
Print #7, "<td width=""75%"" colspan=""3"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"
Print #7, "<font size=""7""><font color=""#FFFFFF"">.</font>"

Print #7, "</td></tr><tr><td width=""25%"" bordercolor=""#FFFFFF"" bordercolorlight=""#FFFFFF"" bordercolordark=""#FFFFFF""><b><font size=""2"">FIRMA:</font></b></td><td width=""25%"" bordercolor=""#000000"" bordercolorlight=""#000000"" bordercolordark=""#000000"">"
Print #7, "<font size=""7""><font color=""#FFFFFF"">.</font>"

Print #7, "</td></tr></table></body></html>"

Close #7

returnvalue = Shell("c:\Archivos de Programa\Internet Explorer\iexplore.exe C:\laficha.html", vbMaximizedFocus)

If returnvalue = 0 Then
returnvalue = Shell("start c:\laficha.html")
If returnvalue = 0 Then MsgBox "Ha ocurrido un error al intentar cargar la ficha en el navegador", vbExclamation, "Error Internet Explorer"
End If

Command12.Visible = False
Command12.Enabled = False
cmdgrabar.Visible = False
cmdgrabar.Enabled = False
cmdimprimir.Enabled = False
cmdimprimir.Visible = False
cmdcancel.Enabled = False
cmdcancel.Visible = False
cmdprintpublic.Enabled = False
cmdprintpublic.Visible = False
Command11.Enabled = False
Command11.Visible = False
Frame1.Visible = False
Frame2.Visible = True
cmdnuevo.Visible = True
cmdnuevo.Enabled = True
tmpficha = 0

End Sub

Private Sub Command12_Click()
If showres = True Then
Command12.BackColor = QBColor(8)
showres = False
Else
Command12.BackColor = QBColor(14)
showres = True
End If
End Sub

Private Sub Command2_Click()

Dim i As Long
Dim cantres As Long
Dim rsindex As Tindexregistro
Dim cantdefichas As Long
Dim b As Integer
onlyone = False

Combo1.ListIndex = 0
Text1.BackColor = QBColor(15)
mnubusqueda.Enabled = False
mnuexportar.Enabled = False
Form1.utilizardatos.Enabled = True

Unload Form3

If Text1.Text = "" Then
Form1.mnubusqueda.Enabled = True
Form1.utilizardatos.Enabled = False
Form1.mnuexportar.Enabled = True
Exit Sub
End If

cantres = 0
Open Trim(dbpath + "\index.dat") For Random As #1 Len = Len(regindex)
cantdefichas = getlastfichanumber - BASE
If (cantdefichas <= 0) Then
MsgBox "No hay fichas definidas", vbExclamation, "Busqueda"
Exit Sub
End If

For i = cantdefichas To 1 Step -1
If i > 0 Then
Get #1, i, rsindex
If InStr(1, rsindex.fullname, StrConv(Text1.Text, vbUpperCase)) > 0 Then
tmpficha = Val(str(rsindex.ficha))
Form3.Label1.Caption = cantres + 1
If Trim(rsindex.estado) <> Trim("ANULADA") Then
If (cantres < MAX_CANT_RESULTS) Then
    cantres = cantres + 1
    Form3.MSFlexGrid1.AddItem rsindex.ficha
    Form3.MSFlexGrid1.Row = cantres
    Form3.MSFlexGrid1.Col = 0
    Form3.MSFlexGrid1.Text = str(rsindex.ficha)
    Form3.MSFlexGrid1.Col = 1
    Form3.MSFlexGrid1.Text = rsindex.fecha
    Form3.MSFlexGrid1.Col = 2
    Form3.MSFlexGrid1.Text = rsindex.fullname
    Form3.MSFlexGrid1.Col = 3
    Form3.MSFlexGrid1.Text = rsindex.telefono
    Form3.MSFlexGrid1.Col = 4
    Form3.MSFlexGrid1.Text = rsindex.modelo
    Form3.MSFlexGrid1.Col = 5
    Form3.MSFlexGrid1.Text = rsindex.estado
     
Else
    Exit For
End If
End If
End If
End If
Next i

For b = 1 To cantres
Form3.MSFlexGrid1.Row = b
Form3.MSFlexGrid1.Col = 5
If Trim(Form3.MSFlexGrid1.Text) = Trim("POR VER") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(11)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("REPARANDO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(12)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("CHEQUEO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(14)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("ENTREGADA") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(8)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("STD/BY") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(13)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("REP.EXT.") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(12)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA NR") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("PV EXT.") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(11)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA BRGS") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("ENTREGAR") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(9)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("PRESUP") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(15)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("DEPOSITO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(8)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("DIAGNOSTIC") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(14)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If

Next b

If cantres > 0 Then Text1.Text = ""

If cantres > 1 Then
    Form3.mnulista.Enabled = False
    Form3.MSFlexGrid1.RemoveItem cantres + 1
    Form3.Show vbModal
ElseIf cantres = 1 Then
    touchedreally = True
    onlyone = True
    MostrarFicha (tmpficha)
    Else
    MsgBox "No hay resultados", vbExclamation, "Busqueda"
    Form1.mnubusqueda.Enabled = True
    Form1.mnuexportar.Enabled = True
End If
Close #1

End Sub

Private Sub Command3_Click()

Dim i As Long
Dim cantres As Long
Dim cantdefichas As Long
Dim rsindex As Tindexregistro
Dim b As Integer
onlyone = False

Combo1.ListIndex = 0
Text2.BackColor = QBColor(15)
mnubusqueda.Enabled = False
mnuexportar.Enabled = False
Form1.utilizardatos.Enabled = True

Unload Form3

If Text2.Text = "" Then
Form1.mnubusqueda.Enabled = True
Form1.utilizardatos.Enabled = False
Form1.mnuexportar.Enabled = True
Exit Sub
End If

cantres = 0

Open Trim(dbpath + "\index.dat") For Random As #1 Len = Len(regindex)
cantdefichas = getlastfichanumber - BASE
If (cantdefichas <= 0) Then
MsgBox "No hay fichas definidas", vbExclamation, "Busqueda"
Exit Sub
End If

For i = cantdefichas To 1 Step -1
If i > 0 Then

Get #1, i, rsindex
If InStr(1, rsindex.telefono, Text2.Text) > 0 Then
tmpficha = Val(str(rsindex.ficha))
Form3.Label1.Caption = cantres + 1
If Trim(rsindex.estado) <> Trim("ANULADA") Then
If (cantres < MAX_CANT_RESULTS) Then
    cantres = cantres + 1
    Form3.MSFlexGrid1.AddItem rsindex.ficha
    Form3.MSFlexGrid1.Row = cantres
    Form3.MSFlexGrid1.Col = 0
    Form3.MSFlexGrid1.Text = str(rsindex.ficha)
    Form3.MSFlexGrid1.Col = 1
    Form3.MSFlexGrid1.Text = rsindex.fecha
    Form3.MSFlexGrid1.Col = 2
    Form3.MSFlexGrid1.Text = rsindex.fullname
    Form3.MSFlexGrid1.Col = 3
    Form3.MSFlexGrid1.Text = rsindex.telefono
    Form3.MSFlexGrid1.Col = 4
    Form3.MSFlexGrid1.Text = rsindex.modelo
    Form3.MSFlexGrid1.Col = 5
    Form3.MSFlexGrid1.Text = rsindex.estado
    
Else
    Exit For
End If
End If
End If
End If
Next i

For b = 1 To cantres
Form3.MSFlexGrid1.Row = b
Form3.MSFlexGrid1.Col = 5
If Trim(Form3.MSFlexGrid1.Text) = Trim("POR VER") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(11)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("REPARANDO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(12)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("CHEQUEO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(14)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("ENTREGADA") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(8)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("STD/BY") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(13)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("REP.EXT.") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(12)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA NR") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("PV EXT.") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(11)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA BRGS") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("ENTREGAR") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(9)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("PRESUP") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(15)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("DEPOSITO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(8)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("DIAGNOSTIC") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(14)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If

Next b

If cantres > 0 Then Text2.Text = ""
If cantres > 1 Then
    Form3.mnulista.Enabled = False
    Form3.MSFlexGrid1.RemoveItem cantres + 1
    Form3.Show vbModal
ElseIf cantres = 1 Then
    touchedreally = True
    onlyone = True
    MostrarFicha (tmpficha)
Else
MsgBox "No hay resultados", vbExclamation, "Busqueda"
Form1.mnubusqueda.Enabled = True
Form1.mnuexportar.Enabled = True
End If
Close #1

End Sub

Private Sub Command4_Click()

trabajos.Show vbModal

End Sub

Private Sub Command5_Click()

Dim i As Long
Dim cantres As Long
Dim cantdefichas As Long
Dim rsindex As Tindexregistro
Dim b As Integer
onlyone = False


Combo1.ListIndex = 0
Text2.BackColor = QBColor(15)
mnubusqueda.Enabled = False
mnuexportar.Enabled = False

Unload Form3
cantres = 0

Open Trim(dbpath + "\index.dat") For Random As #1 Len = Len(regindex)
cantdefichas = getlastfichanumber - BASE
If (cantdefichas <= 0) Then
MsgBox "No hay fichas definidas", vbExclamation, "Busqueda"
Exit Sub
End If

For i = cantdefichas To 1 Step -1
If i > 0 Then

Get #1, i, rsindex
If Trim(rsindex.estado) = Trim(Command5.Tag) Then
tmpficha = Val(str(rsindex.ficha))
Form3.Label1.Caption = cantres + 1
If (cantres < Trim(Form1.tce)) Then
    cantres = cantres + 1
    Form3.MSFlexGrid1.AddItem rsindex.ficha
    Form3.MSFlexGrid1.Row = cantres
    Form3.MSFlexGrid1.Col = 0
    Form3.MSFlexGrid1.Text = str(rsindex.ficha)
    Form3.MSFlexGrid1.Col = 1
    Form3.MSFlexGrid1.Text = rsindex.fecha
    Form3.MSFlexGrid1.Col = 2
    Form3.MSFlexGrid1.Text = rsindex.fullname
    Form3.MSFlexGrid1.Col = 3
    Form3.MSFlexGrid1.Text = rsindex.telefono
    Form3.MSFlexGrid1.Col = 4
    Form3.MSFlexGrid1.Text = rsindex.modelo
    Form3.MSFlexGrid1.Col = 5
    Form3.MSFlexGrid1.Text = rsindex.estado
   
Else
    Exit For
End If
End If
End If
Next i

For b = 1 To cantres
Form3.MSFlexGrid1.Row = b
Form3.MSFlexGrid1.Col = 5
If Trim(Form3.MSFlexGrid1.Text) = Trim("POR VER") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(11)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("REPARANDO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(12)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("CHEQUEO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(14)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("ENTREGADA") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(8)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("STD/BY") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(13)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("REP.EXT.") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(12)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA NR") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("PV EXT.") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(11)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA BRGS") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("ENTREGAR") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(9)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("PRESUP") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(15)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("DEPOSITO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(8)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("DIAGNOSTIC") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(14)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
Next b

If cantres > 0 Then Text2.Text = ""
If cantres > 1 Then
    Form3.mnulista.Enabled = False
    Form3.MSFlexGrid1.RemoveItem cantres + 1
    Form3.Show vbModal
ElseIf cantres = 1 Then
    touchedreally = True
    onlyone = True
    MostrarFicha (tmpficha)
Else
MsgBox "No hay resultados", vbExclamation, "Busqueda"
Form1.mnubusqueda.Enabled = True
Form1.mnuexportar.Enabled = True
End If
Close #1

End Sub

Private Sub Command6_Click()

Dim i As Long
Dim cantres As Long
Dim cantdefichas As Long
Dim rsindex As Tindexregistro
Dim b As Integer
onlyone = False

Combo1.ListIndex = 0
mnubusqueda.Enabled = False
mnuexportar.Enabled = False

Unload Form3
cantres = 0

Open Trim(dbpath + "\index.dat") For Random As #1 Len = Len(regindex)
cantdefichas = getlastfichanumber - BASE

If (cantdefichas <= 0) Then
MsgBox "No hay fichas definidas", vbExclamation, "Busqueda"
Exit Sub
End If

For i = cantdefichas To 1 Step -1
If i > 0 Then

Get #1, i, rsindex

If InStr(1, Trim(rsindex.tecnico), Trim(Combo2.Text)) > 0 Then
tmpficha = Val(str(rsindex.ficha))
Form3.Label1.Caption = cantres + 1
If Trim(rsindex.estado) <> Trim("ANULADA") Then
If (cantres < Trim(Form1.tce1.Text)) Then
    cantres = cantres + 1
    Form3.MSFlexGrid1.AddItem rsindex.ficha
    Form3.MSFlexGrid1.Row = cantres
    Form3.MSFlexGrid1.Col = 0
    Form3.MSFlexGrid1.Text = str(rsindex.ficha)
    Form3.MSFlexGrid1.Col = 1
    Form3.MSFlexGrid1.Text = rsindex.fecha
    Form3.MSFlexGrid1.Col = 2
    Form3.MSFlexGrid1.Text = rsindex.fullname
    Form3.MSFlexGrid1.Col = 3
    Form3.MSFlexGrid1.Text = rsindex.telefono
    Form3.MSFlexGrid1.Col = 4
    Form3.MSFlexGrid1.Text = rsindex.modelo
    Form3.MSFlexGrid1.Col = 5
    Form3.MSFlexGrid1.Text = rsindex.estado
   
Else
    Exit For
End If
End If
End If
End If

Next i

For b = 1 To cantres
Form3.MSFlexGrid1.Row = b
Form3.MSFlexGrid1.Col = 5
If Trim(Form3.MSFlexGrid1.Text) = Trim("POR VER") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(11)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("REPARANDO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(12)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("CHEQUEO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(14)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("ENTREGADA") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(8)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("STD/BY") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(13)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("REP.EXT.") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(12)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA NR") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("PV EXT.") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(11)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA BRGS") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("ENTREGAR") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(9)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("PRESUP") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(15)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("DEPOSITO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(8)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("DIAGNOSTIC") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(14)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
Next b

If cantres > 0 Then Text2.Text = ""
If cantres > 1 Then
    Form3.mnulista.Enabled = False
    Form3.MSFlexGrid1.RemoveItem cantres + 1
    Form3.Show vbModal
ElseIf cantres = 1 Then
    touchedreally = True
    onlyone = True
    MostrarFicha (tmpficha)
Else
MsgBox "No hay resultados", vbExclamation, "Busqueda"
Form1.mnubusqueda.Enabled = True
Form1.mnuexportar.Enabled = True
End If
Close #1

End Sub

Private Sub Command7_Click()

Dim i As Long
Dim cantres As Long
Dim cantdefichas As Long
Dim rsindex As Tindexregistro
Dim b As Integer
onlyone = False


Combo1.ListIndex = 0
Text2.BackColor = QBColor(15)
mnubusqueda.Enabled = False
mnuexportar.Enabled = False

Unload Form3
cantres = 0

Open Trim(dbpath + "\index.dat") For Random As #1 Len = Len(regindex)
cantdefichas = getlastfichanumber - BASE
If (cantdefichas <= 0) Then
MsgBox "No hay fichas definidas", vbExclamation, "Busqueda"
Exit Sub
End If

For i = cantdefichas To 1 Step -1
If i > 0 Then

Get #1, i, rsindex

If InStr(1, Trim(rsindex.tecnico), Trim(Combo2.Text)) And InStr(1, Left(rsindex.confirmacion, 1), Trim(conpre)) > 0 Then
tmpficha = Val(str(rsindex.ficha))
Form3.Label1.Caption = cantres + 1
If Trim(rsindex.estado) <> Trim("ANULADA") Then
If (cantres < Trim(Form1.tce2.Text)) Then
    cantres = cantres + 1
    Form3.MSFlexGrid1.AddItem rsindex.ficha
    Form3.MSFlexGrid1.Row = cantres
    Form3.MSFlexGrid1.Col = 0
    Form3.MSFlexGrid1.Text = str(rsindex.ficha)
    Form3.MSFlexGrid1.Col = 1
    Form3.MSFlexGrid1.Text = rsindex.fecha
    Form3.MSFlexGrid1.Col = 2
    Form3.MSFlexGrid1.Text = rsindex.fullname
    Form3.MSFlexGrid1.Col = 3
    Form3.MSFlexGrid1.Text = rsindex.telefono
    Form3.MSFlexGrid1.Col = 4
    Form3.MSFlexGrid1.Text = rsindex.modelo
    Form3.MSFlexGrid1.Col = 5
    Form3.MSFlexGrid1.Text = rsindex.estado
   
Else
    Exit For
End If
End If
End If

End If
Next i

For b = 1 To cantres
Form3.MSFlexGrid1.Row = b
Form3.MSFlexGrid1.Col = 5
If Trim(Form3.MSFlexGrid1.Text) = Trim("POR VER") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(11)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("REPARANDO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(12)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("CHEQUEO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(14)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("ENTREGADA") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(8)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("STD/BY") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(13)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("REP.EXT.") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(12)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA NR") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("PV EXT.") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(11)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA BRGS") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("ENTREGAR") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(9)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("PRESUP") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(15)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("DEPOSITO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(8)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("DIAGNOSTIC") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(14)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If

Next b

If cantres > 0 Then Text2.Text = ""
If cantres > 1 Then
    Form3.mnulista.Enabled = False
    Form3.MSFlexGrid1.RemoveItem cantres + 1
    Form3.Show vbModal
ElseIf cantres = 1 Then
    onlyone = True
    touchedreally = True
    MostrarFicha (tmpficha)
Else
MsgBox "No hay resultados", vbExclamation, "Busqueda"
Form1.mnubusqueda.Enabled = True
Form1.mnuexportar.Enabled = True
End If
Close #1

End Sub


Private Sub Command8_Click()
Dim i As Long
Dim cantres As Long
Dim cantdefichas As Long
Dim rsindex As Tindexregistro
Dim b As Integer
onlyone = False

Combo1.ListIndex = 0
Text2.BackColor = QBColor(15)
mnubusqueda.Enabled = False
mnuexportar.Enabled = False
Form1.utilizardatos.Enabled = True

Unload Form3

If Text3.Text = "" Then
Form1.mnubusqueda.Enabled = True
Form1.utilizardatos.Enabled = False
Form1.mnuexportar.Enabled = True
Exit Sub
End If

cantres = 0
Open Trim(dbpath + "\index.dat") For Random As #1 Len = Len(regindex)
cantdefichas = getlastfichanumber - BASE
If (cantdefichas <= 0) Then
MsgBox "No hay fichas definidas", vbExclamation, "Busqueda"
Exit Sub
End If


For i = cantdefichas To 1 Step -1
If i > 0 Then
Get #1, i, rsindex
If InStr(1, rsindex.modelo, StrConv(Text3.Text, vbUpperCase)) > 0 Then
tmpficha = Val(str(rsindex.ficha))
Form3.Label1.Caption = cantres + 1
If Trim(rsindex.estado) <> Trim("ANULADA") Then
If (cantres < MAX_CANT_RESULTS) Then
    cantres = cantres + 1
    Form3.MSFlexGrid1.AddItem rsindex.ficha
    Form3.MSFlexGrid1.Row = cantres
    Form3.MSFlexGrid1.Col = 0
    Form3.MSFlexGrid1.Text = str(rsindex.ficha)
    Form3.MSFlexGrid1.Col = 1
    Form3.MSFlexGrid1.Text = rsindex.fecha
    Form3.MSFlexGrid1.Col = 2
    Form3.MSFlexGrid1.Text = rsindex.fullname
    Form3.MSFlexGrid1.Col = 3
    Form3.MSFlexGrid1.Text = rsindex.telefono
    Form3.MSFlexGrid1.Col = 4
    Form3.MSFlexGrid1.Text = rsindex.modelo
    Form3.MSFlexGrid1.Col = 5
    Form3.MSFlexGrid1.Text = rsindex.estado

    
    
Else
    Exit For
End If
End If
End If
End If
Next i

For b = 1 To cantres
Form3.MSFlexGrid1.Row = b
Form3.MSFlexGrid1.Col = 5
If Trim(Form3.MSFlexGrid1.Text) = Trim("POR VER") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(11)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("REPARANDO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(12)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("CHEQUEO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(14)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("ENTREGADA") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(8)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("STD/BY") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(13)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("REP.EXT.") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(12)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA NR") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("PV EXT.") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(11)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA BRGS") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("ENTREGAR") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(9)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("PRESUP") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(15)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("DEPOSITO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(8)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("DIAGNOSTIC") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(14)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If

Next b

If cantres > 0 Then Text2.Text = ""
If cantres > 1 Then
    Form3.mnulista.Enabled = False
    Form3.MSFlexGrid1.RemoveItem cantres + 1
    Form3.Show vbModal
ElseIf cantres = 1 Then
    touchedreally = True
    onlyone = True
    MostrarFicha (tmpficha)
Else
MsgBox "No hay resultados", vbExclamation, "Busqueda"
Form1.mnubusqueda.Enabled = True
Form1.mnuexportar.Enabled = True
End If
Close #1
End Sub

Private Sub Command9_Click()
Dim i As Long
Dim cantres As Long
Dim cantdefichas As Long
Dim rsindex As Tindexregistro
Dim b As Integer
onlyone = False

Combo1.ListIndex = 0
mnubusqueda.Enabled = False
mnuexportar.Enabled = False

Unload Form3

cantres = 0

Open Trim(dbpath + "\index.dat") For Random As #1 Len = Len(regindex)
cantdefichas = getlastfichanumber - BASE

If (cantdefichas <= 0) Then
MsgBox "No hay fichas definidas", vbExclamation, "Busqueda"
Exit Sub
End If

For i = cantdefichas To 1 Step -1
If i > 0 Then
Get #1, i, rsindex
If InStr(1, Left(rsindex.confirmacion, 1), conpre) > 0 Then
tmpficha = Val(str(rsindex.ficha))
Form3.Label1.Caption = cantres + 1
If Trim(rsindex.estado) <> Trim("ANULADA") Then
If (cantres < Trim(Form1.Text4.Text)) Then
    cantres = cantres + 1
    Form3.MSFlexGrid1.AddItem rsindex.ficha
    Form3.MSFlexGrid1.Row = cantres
    Form3.MSFlexGrid1.Col = 0
    Form3.MSFlexGrid1.Text = str(rsindex.ficha)
    Form3.MSFlexGrid1.Col = 1
    Form3.MSFlexGrid1.Text = rsindex.fecha
    Form3.MSFlexGrid1.Col = 2
    Form3.MSFlexGrid1.Text = rsindex.fullname
    Form3.MSFlexGrid1.Col = 3
    Form3.MSFlexGrid1.Text = rsindex.telefono
    Form3.MSFlexGrid1.Col = 4
    Form3.MSFlexGrid1.Text = rsindex.modelo
    Form3.MSFlexGrid1.Col = 5
    Form3.MSFlexGrid1.Text = rsindex.estado

   
Else
    Exit For
End If
End If
End If
End If
Next i

For b = 1 To cantres
Form3.MSFlexGrid1.Row = b
Form3.MSFlexGrid1.Col = 5
If Trim(Form3.MSFlexGrid1.Text) = Trim("POR VER") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(11)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("REPARANDO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(12)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("CHEQUEO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(14)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("ENTREGADA") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(8)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("STD/BY") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(13)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("REP.EXT.") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(12)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA NR") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("PV EXT.") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(11)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("LISTA BRGS") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(10)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("ENTREGAR") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(9)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("PRESUP") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(15)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("DEPOSITO") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(8)
Form3.MSFlexGrid1.CellForeColor = QBColor(15)
End If
If Trim(Form3.MSFlexGrid1.Text) = Trim("DIAGNOSTIC") Then
Form3.MSFlexGrid1.CellBackColor = QBColor(14)
Form3.MSFlexGrid1.CellForeColor = QBColor(0)
End If


Next b

If cantres > 0 Then Text2.Text = ""
If cantres > 1 Then
    Form3.mnulista.Enabled = False
    Form3.MSFlexGrid1.RemoveItem cantres + 1
    Form3.Show vbModal
ElseIf cantres = 1 Then
    touchedreally = True
    onlyone = True
    MostrarFicha (tmpficha)
Else
MsgBox "No hay resultados", vbExclamation, "Busqueda"
Form1.mnubusqueda.Enabled = True
Form1.mnuexportar.Enabled = True
End If
Close #1

End Sub

Private Sub Form_Load()
Dim pepe As String

On Error Resume Next

BASE = 0

Open pconfig For Input As #1
Line Input #1, dbpath
If dbpath = "" Then
MsgBox "Por favor, edite el archivo " & pconfig & " y en la primer linea ponga la ubicacion de la base de datos accesable desde todos lados. Por Ejemplo, \\SERVER\recurso\directorio . NO PONER UNA \ al FINAL.", vbCritical, "Error Importante"
End
Else
Close #1
End If

frmSplash.Show vbModal
Form1.utilizardatos.Enabled = False
showres = False

pepe = ""
Open dbpath + "\base.dat" For Input As #20
Line Input #20, pepe
If pepe = "" Then
MsgBox "Por favor, verifique la conexcion con el servidor. ", vbCritical, "ERROR IMPORTANTE"
End
End If

BASE = Val(pepe)
Close #20

Open dbpath + "\datos.dat" For Random As #24 Len = Len(persona)
Get #24, , persona
Close #24

Form1.Combo1.Clear
Form1.Combo2.Clear
    
Form1.Combo1.AddItem (persona.persona1)
Form1.Combo2.AddItem (persona.persona1)

If Not Form5.Text1.Text = "" Then
    Form1.Combo1.AddItem (Form5.Text1.Text)
    Form1.Combo2.AddItem (Form5.Text1.Text)
    End If

If Not Form5.Text2.Text = "" Then
    Form1.Combo1.AddItem (Form5.Text2.Text)
    Form1.Combo2.AddItem (Form5.Text2.Text)
    End If

If Not Form5.Text3.Text = "" Then
    Form1.Combo1.AddItem (Form5.Text3.Text)
    Form1.Combo2.AddItem (Form5.Text3.Text)
    End If

If Not Form5.Text4.Text = "" Then
    Form1.Combo1.AddItem (Form5.Text4.Text)
    Form1.Combo2.AddItem (Form5.Text4.Text)
    End If

If Not Form5.Text5.Text = "" Then
    Form1.Combo1.AddItem (Form5.Text5.Text)
    Form1.Combo2.AddItem (Form5.Text5.Text)
    End If
    
If Not Form5.Text6.Text = "" Then
    Form1.Combo1.AddItem (Form5.Text6.Text)
    Form1.Combo2.AddItem (Form5.Text6.Text)
    End If

If Not Form5.Text7.Text = "" Then
    Form1.Combo1.AddItem (Form5.Text7.Text)
    Form1.Combo2.AddItem (Form5.Text7.Text)
    End If

If Not Form5.Text8.Text = "" Then
    Form1.Combo1.AddItem (Form5.Text8.Text)
    Form1.Combo2.AddItem (Form5.Text8.Text)
    End If

If Not Form5.Text9.Text = "" Then
    Form1.Combo1.AddItem (Form5.Text9.Text)
    Form1.Combo2.AddItem (Form5.Text9.Text)
    End If

If Not Form5.Text10.Text = "" Then
    Form1.Combo1.AddItem (Form5.Text10.Text)
    Form1.Combo2.AddItem (Form5.Text10.Text)
    End If

If Not Form5.Text11.Text = "" Then
    Form1.Combo1.AddItem (Form5.Text11.Text)
    Form1.Combo2.AddItem (Form5.Text11.Text)
    End If

If Not Form5.Text12.Text = "" Then
    Form1.Combo1.AddItem (Form5.Text12.Text)
    Form1.Combo2.AddItem (Form5.Text12.Text)
    End If


Form1.Combo1.ListIndex = (0)
Form1.Combo2.ListIndex = (0)

Frame1.Visible = False
Frame1.Width = Form1.Width - 150
Frame1.Height = Form1.Height - Frame1.Top - 500

Command12.Visible = False
cmdimprimir.Visible = False
cmdcancel.Visible = False
cmdprintpublic.Visible = False
Command11.Visible = False

estados(1).txt = "POR VER"
estados(2).txt = "REPARANDO"
estados(3).txt = "LISTA"
estados(4).txt = "STD/BY"
estados(5).txt = "ENTREGADA"
estados(6).txt = "CHEQUEO"
estados(7).txt = "REP.EXT."
estados(8).txt = "PV EXT."
estados(9).txt = "LISTA NR"
estados(10).txt = "LISTA BRGS"
estados(11).txt = "ENTREGAR"
estados(12).txt = "PRESUP"
estados(13).txt = "ANULADA"
estados(14).txt = "DEPOSITO"
estados(15).txt = "DIAGNOSTIC"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("�Esta seguro de que desea salir?", vbYesNo, "Confirmacion de cierre") = vbYes Then
Else
Cancel = 1
End If
End Sub

Private Sub Form_Resize()

If Me.WindowState = 1 Then
Me.Caption = "Ficha: " & tmpficha
End If

If tmpficha = 0 Then
Me.Caption = "FluWork v10"
End If

If Me.WindowState = 2 Then
Me.Caption = "FluWork v10"
End If

If Me.WindowState = 0 Then Me.WindowState = vbMaximized
If Me.WindowState = 1 Or Me.WindowState = 2 Then Exit Sub

Frame1.Width = Form1.Width - 150
Frame1.Height = Form1.Height - Frame1.Top - 500
cmdgrabar.Left = Form1.Width - cmdgrabar.Width - 250
cmdcancel.Left = Form1.Width - cmdgrabar.Width - cmdcancel.Width - 350
tfechaingreso.Left = Form1.Width - tfechaingreso.Width - 350
tfechaegreso.Left = Form1.Width - tfechaegreso.Width - 350
tprecio.Left = Form1.Width - tprecio.Width - 350
lbllista.Left = Form1.Width - lbllista.Width - 350
tproblema.Width = Form1.Width - 500
tsolucion.Width = Form1.Width - 500
tadjuntos.Width = Form1.Width - 500

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload statsel
Unload Form3
Unload Form2
Unload Form5
Unload Me
End Sub


Private Sub Label16_Click()
confsel.Show vbModal
Label16.Caption = Trim(registro.confirmacion)
If Label16.Caption = "NO DISPONIBLE" Then
tconfirmacion.Visible = True
Label27.Visible = True
tconfirmacion.Text = " "
ElseIf Label16.Caption = "CONFIRMADO" Then
tconfirmacion.Visible = True
Label27.Visible = True
tconfirmacion.Text = "C-" & Date
ElseIf Label16.Caption = "NO CONFIRMADO" Then
tconfirmacion.Visible = True
Label27.Visible = True
tconfirmacion.Text = "N-" & Date
End If
End Sub

Private Sub lbllista_Click()
statsel.Show vbModal
lbllista.Caption = registro.estado
tfechaegreso.Text = registro.fechaegreso
tprecio.Text = registro.precio
End Sub

Private Sub mnuAcerca_Click()
Form2.Show vbModal
End Sub

Private Sub mnubavanzada_Click()
Dim i As Integer
If Frame3.Visible = True Then
Frame3.Visible = False
mnubavanzada.Checked = True

For i = 0 To 14
Option1(i).Enabled = True
Next i

Option2(1).Enabled = True
Option2(2).Enabled = True
Option2(3).Enabled = True

Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command9.Enabled = True
Command10.Enabled = True
Combo2.Enabled = True

tce.Enabled = True
tce1.Enabled = True
tce2.Enabled = True
Text4.Enabled = True
Else
For i = 0 To 14
Option1(i).Enabled = False
Next i
Option2(1).Enabled = False
Option2(2).Enabled = False
Option2(3).Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
Combo2.Enabled = False
tce.Enabled = False
tce1.Enabled = False
tce2.Enabled = False
Text4.Enabled = False
mnubavanzada.Checked = False
Frame3.Visible = True
End If
End Sub

Private Sub mnuconfig_Click()
Form5.Show vbModal
End Sub

Private Sub mnuexportar_Click()
Dim lineacsv As String
Dim registro As Tregistro
Dim i As Integer
Dim archivo As String
Dim encabezado As String
Dim tamanio As Long
Dim total As Long
unidad = ""
directorio = ""
nombrecompleto = ""

Form10.Show vbModal

If nombrecompleto = "" Then
Exit Sub
End If


encabezado = "Nombre,Direcci�n de correo electr�nico,Calle,Tel�fono particular"

Open pconfig For Input As #3
Line Input #3, dbpath
Close #3

tamanio = FileLen(Trim(dbpath + "\" + "Index.dat"))
total = tamanio / 225

Open nombrecompleto For Output As #2
Print #2, encabezado
For i = 1 To total
archivo = (dbpath + "\" + Trim(16800 + i))
Open archivo For Random As #1 Len = Len(registro)

Get #1, 1, registro

lineacsv = ""
lineacsv = Trim(registro.fullname) & "," & StrConv(Trim(registro.email), vbLowerCase) & "," & Trim(registro.direccion) & "," & Trim(registro.telefono)

Print #2, lineacsv


Close #1

Next i
Close #2
Form4.Show vbModal
End Sub

Private Sub mnusalir_Click()
Unload Me
End Sub

Private Sub msgestado_PanelClick(ByVal Panel As ComctlLib.Panel)
If Frame1.Visible = True Then Exit Sub
bficha.Text = Trim(msgestado.Panels(2).Text)
Command1.Value = True
Combo1.ListIndex = 0
End Sub

Private Sub Option1_Click(Index As Integer)
Command5.Tag = Option1(Index).Caption
End Sub

Private Sub Option2_Click(Index As Integer)
If Index = "1" Then
conpre = "C"
End If
If Index = "2" Then
conpre = "N"
End If
If Index = "3" Then
conpre = " "
End If
End Sub

Private Sub Repext_Click()

Dim returnvalue, e
Dim i As Long
Dim cantres As Long
Dim cantdefichas As Long
Dim rsindex As Tindexregistro
On Error Resume Next

Combo1.ListIndex = 0
cantres = 0

Open Trim(dbpath + "\index.dat") For Random As #1 Len = Len(regindex)
cantdefichas = getlastfichanumber - BASE

Kill "c:\repext.html"

Close #17

Open "c:\repext.html" For Output As #17

Print #17, ""
Print #17, "<table border=""0"" cellpadding=""0"" cellspacing=""1"" width=""100%""><tr><td width=""25%"" rowspan=""3""><center><IMG border=""0"" SRC=""file:"
Print #17, Trim(dbpath) + "\mag.logo.gif"
Print #17, """ width=""190"" height=""159""></center></td><td width=""25%"" valign=""top"" align=""left"" height=""2""><img border=""0"" src=""file:"
Print #17, Trim(dbpath) + "\direccion.gif"
Print #17, """ width=""200"" height=""69""></td><td width=""25%"" valign=""middle"" align=""center""><p align=""right""><img border=""0"" src=""file:"
Print #17, Trim(dbpath) + "\epson.logo.gif"
Print #17, """ width=""128"" height=""40""></td><td width=""25%"" valign=""middle"" align=""center""><img border=""0"" src=""file:"
Print #17, Trim(dbpath) + "\hp.logo.gif"

Print #17, """ width=""111"" height=""115"">"
Print #17, "</td></tr><tr><td width=""75%"" valign=""top"" align=""left"" height=""1"" colspan=""3""><p align=""center""><b><font size=""4"">RESUMEN DE REPARACIONES EXTERNAS</font></b></p>"
Print #17, "</td></tr><tr><td width=""75%"" valign=""top"" align=""left"" colspan=""3"">&nbsp;</td></tr></table>"

For i = cantdefichas To 1 Step -1
If i > 0 Then

Get #1, i, rsindex

If InStr(1, Trim(rsindex.estado), Trim("REP.EXT.")) > 0 Then
   
Print #17, "<table width= ""100%"" border =""1"" cellspacing =0> <tr> <td width=""5%"">"
Print #17, str(rsindex.ficha)
Print #17, "</td> <td width=""5%"" >"
Print #17, rsindex.fecha
Print #17, "</td> <td width=""30%"">"
Print #17, rsindex.fullname
Print #17, "</td> <td width=""5%"">"
Print #17, rsindex.telefono
Print #17, "</td> <td width=""30%"">"
Print #17, rsindex.modelo
Print #17, "</td> <td width=""5%"">"
Print #17, rsindex.estado
Print #17, "</td> </tr>"
End If
   
Else
    Exit For
End If

Next i
Close #1
Print #17, "</table></html>"
Close #17
returnvalue = Shell("c:\Archivos de Programa\Internet Explorer\iexplore.exe C:\repext.html", vbMaximizedFocus)

End Sub

Private Sub tadjuntos_GotFocus()
tadjuntos.BackColor = QBColor(7)
End Sub

Private Sub tadjuntos_Keypress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(tadjuntos.Text) > 1024 Then KeyAscii = 0
End Sub

Private Sub tadjuntos_LostFocus()
tadjuntos.BackColor = QBColor(15)
End Sub

Private Sub tatendidopor_GotFocus()
tatendidopor.BackColor = QBColor(7)
End Sub

Private Sub tatendidopor_Keypress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(tatendidopor.Text) > 50 Then KeyAscii = 0
End Sub

Private Sub tatendidopor_LostFocus()
tatendidopor.BackColor = QBColor(15)
End Sub

Private Sub Tavisadoeldia_GotFocus()
Tavisadoeldia.BackColor = QBColor(7)
End Sub

Private Sub Tavisadoeldia_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Tavisadoeldia.Text) > 30 Then KeyAscii = 0
End Sub

Private Sub Tavisadoeldia_LostFocus()
Tavisadoeldia.BackColor = QBColor(15)
End Sub

Private Sub Tavisadopor_GotFocus()
Tavisadopor.BackColor = QBColor(7)
End Sub

Private Sub Tavisadopor_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Tavisadopor.Text) > 30 Then KeyAscii = 0
End Sub

Private Sub Tavisadopor_LostFocus()
Tavisadopor.BackColor = QBColor(15)
End Sub

Private Sub tce_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(bficha.Text) > 9 Then KeyAscii = 0
If (KeyAscii = 13) And Val(Trim(tce.Text)) > 0 Then Command5.Value = True
If Not ((KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub tce1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(bficha.Text) > 9 Then KeyAscii = 0
If (KeyAscii = 13) And Val(Trim(tce1.Text)) > 0 Then Command6.Value = True
If Not ((KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub Tcontroladopor_GotFocus()
Tcontroladopor.BackColor = QBColor(7)
End Sub

Private Sub Tcontroladopor_keypress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Tcontroladopor.Text) > 50 Then KeyAscii = 0
End Sub

Private Sub Tcontroladopor_LostFocus()
Tcontroladopor.BackColor = QBColor(15)
End Sub

Private Sub tdireccion_GotFocus()
tdireccion.BackColor = QBColor(7)
End Sub

Private Sub tdireccion_keypress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(tdireccion.Text) > 200 Then KeyAscii = 0
End Sub

Private Sub tdireccion_LostFocus()
tdireccion.BackColor = QBColor(15)
End Sub

Private Sub temail_GotFocus()
temail.BackColor = QBColor(7)
End Sub

Private Sub temail_keypress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(temail.Text) > 75 Then KeyAscii = 0
End Sub

Private Sub temail_LostFocus()
temail.BackColor = QBColor(15)
End Sub

Private Sub Text1_GotFocus()
Text2.BackColor = QBColor(15)
bficha.BackColor = QBColor(15)
Text1.BackColor = QBColor(7)
Text3.BackColor = QBColor(15)
bficha.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text1.Text) > 49 Then KeyAscii = 0
If (KeyAscii = 13) And Len(Text1.Text) > 0 Then Command2.Value = True
If Not ((KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii > 96 And KeyAscii < 123) Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 241 Or KeyAscii = 209) Then KeyAscii = 0
End Sub

Private Sub Text2_GotFocus()
bficha.BackColor = QBColor(15)
Text2.BackColor = QBColor(7)
Text1.BackColor = QBColor(15)
Text3.BackColor = QBColor(15)
Text3.Text = ""
Text1.Text = ""
bficha.Text = ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text2.Text) > 14 Then KeyAscii = 0
If (KeyAscii = 13) And Val(Trim(Text2.Text)) > 0 Then Command3.Value = True
If Not ((KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub Text3_GotFocus()
Text3.BackColor = QBColor(7)
bficha.BackColor = QBColor(15)
Text2.BackColor = QBColor(15)
Text1.BackColor = QBColor(15)
Text1.Text = ""
bficha.Text = ""
Text2.Text = ""
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Text3.Text) > 49 Then KeyAscii = 0
If (KeyAscii = 13) And Len(Text3.Text) > 0 Then Command8.Value = True
If Not ((KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii > 96 And KeyAscii < 123) Or KeyAscii = 32 Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub Text3_LostFocus()
Text3.BackColor = QBColor(15)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(bficha.Text) > 9 Then KeyAscii = 0
If (KeyAscii = 13) And Val(Trim(Text4.Text)) > 0 Then Command9.Value = True
If Not ((KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub tfullname_GotFocus()
tfullname.BackColor = QBColor(7)
End Sub

Private Sub tfullname_keypress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(tfullname.Text) > 50 Then KeyAscii = 0
If Not ((KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii > 96 And KeyAscii < 123) Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 241 Or KeyAscii = 209) Then KeyAscii = 0
End Sub

Private Sub tfullname_LostFocus()
tfullname.BackColor = QBColor(15)
End Sub

Private Sub Tllamareldia_GotFocus()
Tllamareldia.BackColor = QBColor(7)
End Sub

Private Sub Tllamareldia_keypress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(Tllamareldia.Text) > 30 Then KeyAscii = 0
End Sub

Private Sub Tllamareldia_LostFocus()
Tllamareldia.BackColor = QBColor(15)
End Sub

Private Sub tmodelo_GotFocus()
tmodelo.BackColor = QBColor(7)
End Sub

Private Sub tmodelo_keypress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(tmodelo.Text) > 50 Then KeyAscii = 0
End Sub

Private Sub tmodelo_LostFocus()
tmodelo.BackColor = QBColor(15)
End Sub

Private Sub tnserie_GotFocus()
tnserie.BackColor = QBColor(7)
End Sub

Private Sub tnserie_keypress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(tnserie.Text) > 50 Then KeyAscii = 0
End Sub

Private Sub tnserie_LostFocus()
tnserie.BackColor = QBColor(15)
End Sub

Private Sub tprecio_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(tprecio.Text) > 10 Then KeyAscii = 0
End Sub

Private Sub tpresupuesto_GotFocus()
tpresupuesto.BackColor = QBColor(7)
End Sub

Private Sub tpresupuesto_keypress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(tpresupuesto.Text) > 10 Then KeyAscii = 0
End Sub

Private Sub tpresupuesto_LostFocus()
tpresupuesto.BackColor = QBColor(15)
End Sub

Private Sub tproblema_GotFocus()
tproblema.BackColor = QBColor(7)
End Sub

Private Sub tproblema_keypress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(tproblema.Text) > 1024 Then KeyAscii = 0
End Sub

Private Sub tproblema_LostFocus()
tproblema.BackColor = QBColor(15)
End Sub

Private Sub ttecnico_GotFocus()
ttecnico.BackColor = QBColor(7)
End Sub

Private Sub ttecnico_keypress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(ttecnico.Text) > 50 Then KeyAscii = 0
End Sub

Private Sub ttecnico_LostFocus()
ttecnico.BackColor = QBColor(15)
End Sub

Private Sub ttelefono_GotFocus()
ttelefono.BackColor = QBColor(7)
End Sub

Private Sub ttelefono_Keypress(KeyAscii As Integer)
If KeyAscii <> 8 And KeyAscii <> 13 And Len(ttelefono.Text) > 14 Then KeyAscii = 0
If Not ((KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub ttelefono_LostFocus()
ttelefono.BackColor = QBColor(15)
End Sub

Private Sub utilizardatos_Click()

Combo1.ListIndex = 0
tficha.Text = ""
tfechaingreso.Text = ""
tfechaegreso.Text = ""
tpresupuesto.Text = ""
tprecio.Text = ""
lbllista.Caption = "POR VER"
lbllista.ForeColor = QBColor(11)
tatendidopor.Text = ""
ttecnico.Text = ""
tmodelo.Text = ""
tnserie.Text = ""
Tllamareldia.Text = ""
Tcontroladopor.Text = ""
Tavisadoeldia.Text = ""
tadjuntos.Text = ""
tproblema.Text = ""
tsolucion.Text = ""
tconfirmacion.Text = ""
tfechaingreso.Text = Trim(str(Date))
Label16.Caption = "NO DISPONIBLE"
tconfirmacion.Text = ""
End Sub
