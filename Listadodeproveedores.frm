VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Listadodeproveedores 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PROVEEDORES"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Listado de Proveedores"
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
      Width           =   8655
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5400
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ordenar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5280
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5280
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5280
         Width           =   1695
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Listadodeproveedores.frx":0000
         Height          =   4815
         Left            =   120
         OleObjectBlob   =   "Listadodeproveedores.frx":0014
         TabIndex        =   1
         Top             =   360
         Width           =   8415
      End
   End
End
Attribute VB_Name = "Listadodeproveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.DatabaseName = App.Path & ("\VENTAS.mdb")
Data1.RecordSource = "SELECT * FROM Proveedores order by Empresa"
Data1.Refresh
DBGrid1.Refresh

End Sub

Private Sub Command2_Click()
MsgBox "In Construction..."

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & ("\VENTAS.mdb")
Data1.RecordSource = "SELECT * FROM Proveedores order by Codigo"
Data1.Refresh
DBGrid1.Refresh
Dim intField As Integer
 Dim intRecord As Integer
  intRecord = Data1.Recordset.RecordCount
  intField = Data1.Recordset.Fields.Count
  'call the procedure here...
  Call AdjustDataGridColumns _
  (DBGrid1, Data1, intRecord, intField, True)
End Sub
