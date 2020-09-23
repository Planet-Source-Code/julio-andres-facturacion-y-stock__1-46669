VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Listadeventas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LISTA DE VENTAS"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Listado de Ventas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   "C:\utilesvbasic\Prg Ventas\VENTAS.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Librodeventas"
         Top             =   5520
         Visible         =   0   'False
         Width           =   3375
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Listadeventas.frx":0000
         Height          =   3135
         Left            =   240
         OleObjectBlob   =   "Listadeventas.frx":0014
         TabIndex        =   17
         Top             =   360
         Width           =   9615
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Librodeventas"
         Top             =   5040
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Totales Cobrados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   5160
         TabIndex        =   11
         Top             =   3600
         Width           =   4695
         Begin VB.TextBox txtsuma 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   840
            TabIndex        =   14
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtiva 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            TabIndex        =   13
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txttotal 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3240
            TabIndex        =   12
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3600
            TabIndex        =   20
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "IVA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   19
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "SUMA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1080
            TabIndex        =   18
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Búsqueda"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   240
         TabIndex        =   4
         Top             =   3600
         Width           =   4335
         Begin VB.Data Data3 
            Caption         =   "Data3"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   300
            Left            =   1680
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   0
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1920
            TabIndex        =   16
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Clientes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   1800
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Diaria"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1920
            TabIndex        =   9
            Top             =   360
            Width           =   495
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Mensual"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   840
            Width           =   1455
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1920
            TabIndex        =   7
            Top             =   840
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Anual"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1920
            TabIndex        =   5
            Top             =   1320
            Width           =   495
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ordenar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5400
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5400
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5400
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Listadeventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
MsgBox "En construccion"

End Sub

Private Sub Command2_Click()
MsgBox "EN CONSTRUCCION"
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & ("\Ventas.mdb")
Data1.RecordSource = "SELECT * FROM Librodeventas order by Ndecaja"
Data2.DatabaseName = App.Path & ("\Ventas.mdb")
Data2.RecordSource = "SELECT * FROM Librodeventas order by Ndecaja"
Data3.DatabaseName = App.Path & ("\Ventas.mdb")
Data3.RecordSource = "SELECT * FROM Librodeventas order by Ndecaja"

Data1.Refresh
Data2.Refresh
Data3.Refresh
Dim intRecord As Integer
Dim intField As Integer
  intRecord = Data1.Recordset.RecordCount
  intField = Data1.Recordset.Fields.Count
  'call the procedure here...
  Call AdjustDataGridColumns _
  (DBGrid1, Data1, intRecord, intField, True)

Data1.DatabaseName = App.Path & "\VENTAS.mdb"

Data1.RecordSource = "Select * from Librodeventas "
Data3.DatabaseName = App.Path & "\VENTAS.mdb"

Data3.RecordSource = "Select * from Librodeventas "

Data2.DatabaseName = App.Path & "\VENTAS.mdb"
Data2.RecordSource = "SELECT  sum(Suma) As total from Librodeventas "
Data2.Refresh
txtsuma.Text = Data2.Recordset!total
Data3.DatabaseName = App.Path & "\VENTAS.mdb"
Data3.RecordSource = "SELECT  sum(Iva) As totaliva from Librodeventas "
Data3.Refresh
txtiva.Text = Data3.Recordset!totaliva
txttotal.Text = Val(txtsuma.Text) + Val(txtiva.Text)
End Sub



Private Sub Option1_Click()
MsgBox "En construccion"
End Sub

Private Sub Option2_Click()
MsgBox "En construccion"

End Sub

Private Sub Option3_Click()
MsgBox "En construccion"

End Sub
