VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Principal 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Gestor de Ventas"
   ClientHeight    =   7170
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9825
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Artículos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      TabIndex        =   2
      Top             =   3600
      Width           =   2055
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2580
         ItemData        =   "Principal.frx":0000
         Left            =   120
         List            =   "Principal.frx":0002
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2580
         ItemData        =   "Principal.frx":0004
         Left            =   120
         List            =   "Principal.frx":0006
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ventas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   7455
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Principal.frx":0008
         Height          =   2655
         Left            =   120
         OleObjectBlob   =   "Principal.frx":001C
         TabIndex        =   66
         Top             =   2880
         Width           =   7215
      End
      Begin VB.Frame Frame6 
         Caption         =   "Articulos ocultos"
         Height          =   1215
         Left            =   -240
         TabIndex        =   57
         Top             =   0
         Visible         =   0   'False
         Width           =   5775
         Begin VB.TextBox txtreferencia 
            DataSource      =   "Data2"
            Height          =   375
            Left            =   3960
            TabIndex        =   64
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtstockmax 
            DataSource      =   "Data2"
            Height          =   375
            Left            =   2640
            TabIndex        =   62
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtstockmin 
            DataSource      =   "Data2"
            Height          =   375
            Left            =   1680
            TabIndex        =   60
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtstockreal 
            DataSource      =   "Data2"
            Height          =   360
            Left            =   240
            TabIndex        =   58
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Referencia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4080
            TabIndex        =   65
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Stockmaximo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   63
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Stockminimo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            TabIndex        =   61
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Stockactual"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   59
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "cuando hace click en Facturar guarda en la tabla FACTURA"
         Height          =   1095
         Left            =   480
         TabIndex        =   54
         Top             =   3240
         Visible         =   0   'False
         Width           =   5655
         Begin VB.TextBox txtfecha2 
            DataSource      =   "Data8"
            Height          =   360
            Left            =   600
            TabIndex        =   55
            Text            =   "Text8"
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   56
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Data Data8 
         Caption         =   "Data8"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Factura"
         Top             =   5880
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   2535
         Left            =   -6480
         TabIndex        =   38
         Top             =   -1200
         Visible         =   0   'False
         Width           =   6375
         Begin VB.TextBox txtcliente2 
            DataSource      =   "Data3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            MaxLength       =   13
            TabIndex        =   53
            Top             =   1080
            Width           =   3735
         End
         Begin VB.TextBox txtsuma2 
            DataSource      =   "Data7"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4680
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtiva2 
            DataSource      =   "Data7"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4680
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox txttotal2 
            DataSource      =   "Data7"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4680
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   2160
            Width           =   1335
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3840
            TabIndex        =   46
            Text            =   "IVA"
            Top             =   1800
            Width           =   735
         End
         Begin VB.TextBox txtformadepago 
            DataSource      =   "Data7"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox txtfecha 
            DataSource      =   "Data7"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtndeventa2 
            DataSource      =   "Data7"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   " Suma: $"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3840
            TabIndex        =   51
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Total: $"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3960
            TabIndex        =   50
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Forma de Pago:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   43
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de Venta:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.TextBox txtcliente 
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         MaxLength       =   13
         TabIndex        =   52
         Top             =   600
         Width           =   3735
      End
      Begin VB.Data Data7 
         Caption         =   "Data7"
         Connect         =   "Access"
         DatabaseName    =   "C:\utilesvbasic\Prg Ventas\VENTAS.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Librodeventas"
         Top             =   5640
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4920
         TabIndex        =   37
         Text            =   "% IVA"
         Top             =   6120
         Width           =   735
      End
      Begin VB.Data Data6 
         Caption         =   "Data6"
         Connect         =   "Access"
         DatabaseName    =   "C:\utilesvbasic\Prg Ventas\VENTAS.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Iva"
         Top             =   6480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtpreciodecosto 
         Height          =   345
         Left            =   3600
         TabIndex        =   34
         Top             =   1920
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Data Data5 
         Caption         =   "Data5"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Ventaporpantalla"
         Top             =   6240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Data Data3 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Ventaporpantalla"
         Top             =   5760
         Width           =   1815
      End
      Begin VB.TextBox txtcodigo 
         DataSource      =   "Data3"
         Height          =   360
         Left            =   1080
         TabIndex        =   33
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Data Data4 
         Caption         =   "Data4"
         Connect         =   "Access"
         DatabaseName    =   "C:\utilesvbasic\Prg Ventas\VENTAS.MDB"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Formadepago"
         Top             =   6000
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Left            =   4560
         TabIndex        =   32
         Top             =   1560
         Width           =   255
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Articulos"
         Top             =   6480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Clientes"
         Top             =   6240
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txttotal 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "$0,00"
         Top             =   6480
         Width           =   1335
      End
      Begin VB.TextBox txtiva 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "$0,00"
         Top             =   6120
         Width           =   1335
      End
      Begin VB.TextBox txtsuma 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "$0,00"
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5640
         TabIndex        =   27
         Top             =   6120
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Factura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   6240
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   24
         Top             =   5760
         Width           =   1335
      End
      Begin VB.TextBox txtaño 
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         MaxLength       =   4
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txtmes 
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtdia 
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtndelinea 
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox txtndeventa 
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6480
         Picture         =   "Principal.frx":09E7
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtprecio 
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtcantidad 
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtarticulo 
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         TabIndex        =   8
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtcuil 
         DataSource      =   "Data3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5760
         MaxLength       =   13
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txt21 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         DataField       =   "Iva"
         DataSource      =   "Data6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4560
         TabIndex        =   36
         Text            =   "21"
         Top             =   6110
         Width           =   375
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio de Costo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   35
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Total: $"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   28
         Top             =   6480
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   " Suma: $"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   26
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pago:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   5760
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   19
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de Art.:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Venta:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   12
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Cant."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Artículo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   855
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   0
         X2              =   7455
         Y1              =   2190
         Y2              =   2205
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   7440
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   7440
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   0
         X2              =   7455
         Y1              =   1200
         Y2              =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "C.U.I.L Nº:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Menu mnuprincipalventas 
      Caption         =   "Ventas"
      Begin VB.Menu mnunuevaventa 
         Caption         =   "Nueva Venta"
      End
      Begin VB.Menu mnufacturar 
         Caption         =   "Facturar"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuacercade 
         Caption         =   "Acerca de..."
      End
      Begin VB.Menu mnusalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuarticulos 
      Caption         =   "Artículos"
      Begin VB.Menu mnuarticulo 
         Caption         =   "Nuevo Artículo"
      End
      Begin VB.Menu mnuestadodealmacen 
         Caption         =   "Estado de Almacén"
      End
   End
   Begin VB.Menu mnuclientes 
      Caption         =   "Clientes"
      Begin VB.Menu mnunuevocliente 
         Caption         =   "Nuevo Cliente"
      End
      Begin VB.Menu mnulistadodecliente 
         Caption         =   "Listado de Clientes"
      End
   End
   Begin VB.Menu mnuproveedores 
      Caption         =   "Proveedores"
      Begin VB.Menu mnunuevoproveedor 
         Caption         =   "Nuevo Proveedor"
      End
      Begin VB.Menu mnulistadodeproveedores 
         Caption         =   "Listado de Proveedores"
      End
   End
   Begin VB.Menu mnucaja 
      Caption         =   "Caja"
      Begin VB.Menu mnugastar 
         Caption         =   "Gastar"
      End
      Begin VB.Menu mnuingresar 
         Caption         =   "Ingresar"
      End
      Begin VB.Menu mnulibrodecuentas 
         Caption         =   "Libro de Cuentas"
      End
      Begin VB.Menu mnulistadeventas 
         Caption         =   "Libro de Ventas"
      End
   End
   Begin VB.Menu mnuentorno 
      Caption         =   "Entorno"
      Begin VB.Menu mnudatosdelnegocio 
         Caption         =   "Datos del Negocio"
      End
      Begin VB.Menu mnufecha 
         Caption         =   "Fecha Actual"
      End
      Begin VB.Menu mnuivaactual 
         Caption         =   "I.V.A Actual"
      End
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programa de stock hecho por Julio R. Andres
'E-Mail julioan@topmail.com.ar
'faltan todavia algunos detalles como borrar los datos en la base en la
'tabla VENTAPORPANTALLA cada vez que el usuario hace click en FACTURA
'En la otra version lo terminaré
Private Sub Check1_Click()
If Check1.Value = 0 Then
Data5.DatabaseName = App.Path & "\VENTAS.mdb"
Data5.RecordSource = "SELECT  sum(Importe) As total from Ventaporpantalla "
Data5.Refresh
txtsuma.Text = Data5.Recordset!total
txtiva.Text = "00,00"
txttotal.Text = Val(Str(txtsuma.Text)) + Val(Str(txtiva.Text))
txttotal.Text = Format(txttotal.Text, "$##,##0.00")
txtndelinea.Text = Data3.Recordset.RecordCount
Else
Data5.DatabaseName = App.Path & "\VENTAS.mdb"
Data5.RecordSource = "SELECT  sum(Importe) As total from Ventaporpantalla "
Data5.Refresh
txtsuma.Text = Data5.Recordset!total
txtiva.Text = Val(Str(txtsuma.Text)) * Val(Str(txt21.Text)) / 100
txttotal.Text = Val(Str(txtsuma.Text)) + Val(Str(txtiva.Text))
txttotal.Text = Format(txttotal.Text, "$##,##0.00")
txtndelinea.Text = Data3.Recordset.RecordCount
End If
End Sub

Private Sub Command1_Click()
 ' On Error GoTo AddErr
'If Data3.Recordset.RecordCount = 0 Then Exit Sub
        
       ' With Data3
            
        '    .Recordset.Edit
         '   .Recordset.MoveFirst
'End With
If Check1.Value = 0 Then
Else
    
    Data3.Recordset.AddNew
    Data7.Recordset.AddNew
txtfecha.Text = txtdia.Text & "/" & txtmes.Text & "/" & txtaño.Text
txtformadepago.Text = Combo1.Text
txtprecio.Text = Format(txtprecio.Text, "$##,##0.00")
txtiva.Text = Format(txtiva.Text, "$##,##0.00")
txtsuma.Text = Format(txtsuma.Text, "$##,##0.00")
txttotal.Text = Format(txttotal.Text, "$##,##0.00")

txtprecio.Text = Val(Str(txtcantidad.Text)) * Val(Str(txtprecio.Text))
 Data3.Recordset.Fields("Articulo") = txtarticulo.Text
 Data3.Recordset.Fields("Codigo") = txtcodigo.Text
 Data3.Recordset.Fields("Cantidad") = txtcantidad.Text
 Data3.Recordset.Fields("Precio") = txtpreciodecosto.Text
 Data3.Recordset.Fields("Importe") = txtprecio.Text
 Data3.Recordset.Fields("Dia") = txtdia.Text
 Data3.Recordset.Fields("Mes") = txtmes.Text
 Data3.Recordset.Fields("Año") = txtaño.Text

  'Data3.Enabled = False
 Data3.Recordset.Update
'causa un error si no hay datos en la base el "MoveLast"
'  Data3.Recordset.MoveLast


Data3.DatabaseName = App.Path & "\VENTAS.mdb"

Data3.RecordSource = "Select * from Ventaporpantalla "

Data5.DatabaseName = App.Path & "\VENTAS.mdb"
Data5.RecordSource = "SELECT  sum(Importe) As total from Ventaporpantalla "
Data5.Refresh
txtsuma.Text = Data5.Recordset!total


txtiva.Text = Val(Str(txtsuma.Text)) * Val(Str(txt21.Text)) / 100
txttotal.Text = Val(Str(txtsuma.Text)) + Val(Str(txtiva.Text))
txtndelinea.Text = Data3.Recordset.RecordCount

With Data7
.Recordset.Fields("Ndecaja") = txtndeventa2.Text
.Recordset.Fields("Fecha") = txtfecha.Text
.Recordset.Fields("Cliente") = txtcliente2.Text
.Recordset.Fields("Fdepago") = txtformadepago.Text
.Recordset.Fields("Suma") = txtsuma2.Text
.Recordset.Fields("Iva") = txtiva2.Text
.Recordset.Fields("Total") = txttotal2.Text
.Refresh
End With
Data2.DatabaseName = App.Path & ("\VENTAS.mdb")
Data2.RecordSource = "SELECT * FROM Articulos order by Articulo"
    
    Data8.Recordset.AddNew
txtfecha2.Text = txtdia.Text & "/" & txtmes.Text & "/" & txtaño.Text
txtprecio.Text = Val(Str(txtcantidad.Text)) * Val(Str(txtprecio.Text))
 Data8.Recordset.Fields("Articulo") = txtarticulo.Text
 Data8.Recordset.Fields("Ndecaja") = txtndeventa.Text
 Data8.Recordset.Fields("Cantidad") = txtcantidad.Text
 Data8.Recordset.Fields("Precio") = txtprecio.Text
 Data8.Recordset.Fields("Importe") = txttotal.Text
  Data8.Recordset.Fields("Fdepago") = Combo1.Text
 Data8.Recordset.Fields("Fecha") = txtfecha2.Text
 Data8.Recordset.Fields("Cliente") = txtcliente.Text
 Data8.Recordset.Fields("Iva") = txtiva.Text
Data2.Recordset.Edit
Data2.Recordset.Fields("Stockactual") = txtstockreal.Text - 1
'Data2.Refresh
  'Data3.Enabled = False
 If txtstockreal.Text <= txtstockmin.Text Then
txtreferencia.Text = "PEDIR STOCK"
Data2.Recordset.Fields("Referencia") = txtreferencia.Text
End If

 Data8.Refresh
Data2.Refresh

Command1.Enabled = False
txtcantidad.Text = "1"
End If
 'AddErr:
 ' MsgBox Err.Description
 
  End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command2_Click()
Data2.DatabaseName = App.Path & ("\VENTAS.mdb")
Data2.RecordSource = "SELECT * FROM Articulos order by Articulo"
    
    Data8.Recordset.AddNew
txtfecha2.Text = txtdia.Text & "/" & txtmes.Text & "/" & txtaño.Text
txtprecio.Text = Val(txtcantidad.Text) * Val(txtprecio.Text)
 Data8.Recordset.Fields("Articulo") = txtarticulo.Text
 Data8.Recordset.Fields("Ndecaja") = txtndeventa.Text
 Data8.Recordset.Fields("Cantidad") = txtcantidad.Text
 Data8.Recordset.Fields("Precio") = txtprecio.Text
 Data8.Recordset.Fields("Importe") = txttotal.Text
  Data8.Recordset.Fields("Fdepago") = Combo1.Text
 Data8.Recordset.Fields("Fecha") = txtfecha2.Text
 Data8.Recordset.Fields("Cliente") = txtcliente.Text
 Data8.Recordset.Fields("Iva") = txtiva.Text
Data2.Recordset.Edit
Data2.Recordset.Fields("Stockactual") = txtstockreal.Text - 1
'Data2.Refresh
  'Data3.Enabled = False
 Data8.Refresh
 DBGrid1.Refresh
'Data3.Refresh
MsgBox "Los productos seleccionados para la venta en la grilla(DBGRI1) han sidos salvados... para imprimir está en CONSTRUCCION."
End Sub

Private Sub Data2_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Form_Load()

'Command10.Enabled = True
Data1.DatabaseName = App.Path & ("\VENTAS.mdb")
Data1.RecordSource = "SELECT * FROM Clientes order by Empresa"
Data2.DatabaseName = App.Path & ("\VENTAS.mdb")
Data2.RecordSource = "SELECT * FROM Articulos order by Articulo"
Data3.DatabaseName = App.Path & ("\VENTAS.mdb")
Data3.RecordSource = "SELECT * FROM Ventaporpantalla order by Articulo"
Data4.DatabaseName = App.Path & ("\VENTAS.mdb")
Data4.RecordSource = "SELECT * FROM Formadepago order by Pago"
Data6.DatabaseName = App.Path & ("\VENTAS.mdb")
Data6.RecordSource = "SELECT * FROM Iva "
Data7.DatabaseName = App.Path & ("\VENTAS.mdb")
Data7.RecordSource = "SELECT * FROM Librodeventas "
Data8.DatabaseName = App.Path & ("\VENTAS.mdb")
Data8.RecordSource = "SELECT * FROM Factura "

Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.Refresh
Data6.Refresh
Data7.Refresh
Data8.Refresh
'   Dim intRecord As Integer
'Dim intField As Integer
 ' intRecord = Data3.Recordset.RecordCount
  'intField = Data3.Recordset.Fields.Count
  'call the procedure here...
  'Call AdjustDataGridColumns(DBGrid1, Data3, intRecord, intField, True)
    Do While Not Data1.Recordset.EOF
               ' start from beginning of the records
               ' work to the end of the records
        
            
            List1.AddItem IIf(IsNull(Data1.Recordset("Empresa")), "", Data1.Recordset("Empresa")), i
            'List2.AddItem IIf(IsNull(Data1.Recordset("Origen")), "", Data1.Recordset("Origen"))
                      Data1.Recordset.MoveNext
            i = i + 1
            
        Loop
Call cargararticulos
Call Formadepago
If List1.ListCount = -0 Then
List1.AddItem "No hay Clientes..."
Else
End If
If List2.ListCount = -0 Then
List2.AddItem "No hay Artículos..."
Else
End If
Combo1.ListIndex = 3
txtdia = Format(Date, "DD")
txtmes = Format(Date, "MM")
txtaño = Format(Date, "YYYY")
txtndelinea.Text = Data3.Recordset.RecordCount

If Data3.Recordset.RecordCount = 0 Then Exit Sub
        
        With Data3
            .Recordset.Edit
            .Recordset.MoveFirst
End With
End Sub

Private Sub List1_Click()
Dim cdname As String
cdname = List1.List(List1.ListIndex)
   If Trim(cdname) <> "" Then
        If (Right(cdname, 1) <> "*") Then
        cdname = cdname '+ "*"
    End If
Data1.RecordSource = "SELECT * FROM Clientes WHERE Empresa like '" & cdname & "'order by Empresa"
    Data1.Refresh
   
    
            
            txtcuil.Text = IIf(IsNull(Data1.Recordset("cuil")), "", Data1.Recordset("cuil"))
txtcliente.Text = IIf(IsNull(Data1.Recordset("Empresa")), "", Data1.Recordset("Empresa"))

End If

End Sub

Private Sub List2_Click()
Dim cdname As String
cdname = List2.List(List2.ListIndex)
   If Trim(cdname) <> "" Then
        If (Right(cdname, 1) <> "*") Then
        cdname = cdname '+ "*"
    End If
Data2.RecordSource = "SELECT * FROM Articulos WHERE Articulo like '" & cdname & "'order by Articulo"
    Data2.Refresh
   
    
            
txtprecio.Text = IIf(IsNull(Data2.Recordset("Preciodecosto")), "", Data2.Recordset("Preciodecosto"))
txtarticulo.Text = IIf(IsNull(Data2.Recordset("Articulo")), "", Data2.Recordset("Articulo"))
txtcodigo.Text = IIf(IsNull(Data2.Recordset("Codigo")), "", Data2.Recordset("Codigo"))
txtpreciodecosto.Text = IIf(IsNull(Data2.Recordset("Preciodecosto")), "", Data2.Recordset("Preciodecosto"))
txtstockreal.Text = IIf(IsNull(Data2.Recordset("Stockactual")), "", Data2.Recordset("Stockactual"))
txtstockmax.Text = IIf(IsNull(Data2.Recordset("Stockmaximo")), "", Data2.Recordset("Stockmaximo"))
txtstockmin.Text = IIf(IsNull(Data2.Recordset("Stockminimo")), "", Data2.Recordset("Stockminimo"))
txtreferencia.Text = IIf(IsNull(Data2.Recordset("Referencia")), "", Data2.Recordset("Referencia"))

Command1.Enabled = True
txtprecio.Text = Format(txtprecio.Text, "$##,##0.00")

End If
End Sub

Private Sub mnuarticulo_Click()
Articulos.Show
End Sub

Private Sub mnudatosdelnegocio_Click()
datosdelaempresa.Show
End Sub

Private Sub mnuestadodealmacen_Click()
Almacen.Show
End Sub


Private Sub mnugastar_Click()
Gastar.Show
End Sub

Private Sub mnuingresar_Click()
Ingresar.Show
End Sub

Private Sub mnuivaactual_Click()
Iva.Show
End Sub

Private Sub mnulibrodecuentas_Click()
librodecuentas.Show
End Sub

Private Sub mnulistadeventas_Click()
Listadeventas.Show
End Sub

Private Sub mnulistadodecliente_Click()
listadodeclientes.Show
End Sub

Private Sub mnulistadodeproveedores_Click()
Listadodeproveedores.Show
End Sub

Private Sub mnunuevocliente_Click()
Clientes.Show
End Sub

Private Sub mnunuevoproveedor_Click()
Proveedores.Show
End Sub
Public Sub cargararticulos()
Data2.DatabaseName = App.Path & ("\VENTAS.mdb")
Data2.RecordSource = "SELECT * FROM Articulos order by Articulo"
    Data2.Refresh
   
    Do While Not Data2.Recordset.EOF
               ' start from beginning of the records
               ' work to the end of the records
        
            
            List2.AddItem IIf(IsNull(Data2.Recordset("Articulo")), "", Data2.Recordset("Articulo")), i
            'List2.AddItem IIf(IsNull(Data1.Recordset("Origen")), "", Data1.Recordset("Origen"))
                      Data2.Recordset.MoveNext
            i = i + 1
            
        Loop
        
End Sub

Private Sub Text2_Change()

End Sub

Private Sub txtcliente_Change()
txtcliente2.Text = txtcliente.Text
End Sub

Private Sub txtiva_Change()
txtiva2.Text = txtiva.Text
End Sub

Private Sub txtndeventa_Change()
txtndeventa2.Text = txtndeventa.Text
End Sub

Private Sub txtprecio_Change()

End Sub

Private Sub txtpreciodecosto_Change()

End Sub

Private Sub txtsuma_Change()
txtsuma2.Text = txtsuma.Text
End Sub

Private Sub txttotal_Change()
txttotal2.Text = txttotal.Text
End Sub

Private Sub VScroll1_Change()
txtcantidad.Text = VScroll1.Value

End Sub
Public Sub Formadepago()
Data4.DatabaseName = App.Path & ("\VENTAS.mdb")
Data4.RecordSource = "SELECT * FROM Formadepago order by Pago"
Combo1.Clear
Do While Not Data4.Recordset.EOF
               ' start from beginning of the records
               ' work to the end of the records
        
            
            Combo1.AddItem IIf(IsNull(Data4.Recordset("Pago")), "", Data4.Recordset("Pago")), i
            'List2.AddItem IIf(IsNull(Data1.Recordset("Origen")), "", Data1.Recordset("Origen"))
                      Data4.Recordset.MoveNext
            i = i + 1
            
        Loop
End Sub

Public Function RemoveComma(strNumber As String) As String

    Dim ColStr1 As String
    Dim ColStr2 As String
    Dim ColStr3 As String
    Dim tmpPos As Integer
    tmpPos = InStr(1, strNumber, ",")


    If tmpPos > 0 Then
        ColStr1 = Mid(strNumber, 1, tmpPos - 1)
        ColStr2 = Mid(strNumber, tmpPos + 1)
        ColStr3 = ColStr1 & ColStr2
        RemoveComma = ColStr3
    Else
        RemoveComma = strNumber
    End If

End Function

