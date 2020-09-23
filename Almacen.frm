VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Almacen 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Almacén"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Almacén"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   "C:\utilesvbasic\Prg Ventas\VENTAS.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   4200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Proveedores"
         Top             =   6000
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Almacen.frx":0000
         Height          =   4575
         Left            =   120
         OleObjectBlob   =   "Almacen.frx":0014
         TabIndex        =   8
         Top             =   360
         Width           =   8775
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\utilesvbasic\Prg Ventas\VENTAS.MDB"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   3240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Articulos"
         Top             =   5040
         Visible         =   0   'False
         Width           =   2535
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
         Height          =   495
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5520
         Width           =   1335
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
         Height          =   495
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ordenar Todo"
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
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5520
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Búsqueda"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   5040
         Width           =   3855
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
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
            Left            =   1560
            TabIndex        =   3
            Top             =   480
            Width           =   2175
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Proveedor"
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
            TabIndex        =   2
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   """Pedir Stock"" : Stock Agotado"
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
         Left            =   5880
         TabIndex        =   7
         Top             =   5040
         Width           =   3015
      End
   End
End
Attribute VB_Name = "Almacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
Combo1.Enabled = True
Else
Combo1.Enabled = False
End If
End Sub

Private Sub Combo1_Click()
Data1.DatabaseName = App.Path & ("\VENTAS.mdb")
Data1.RecordSource = "SELECT * FROM Articulos order by Articulo"
Dim cdname As String
cdname = Combo1.List(Combo1.ListIndex)
   If Trim(cdname) <> "" Then
        If (Right(cdname, 1) <> "*") Then
        cdname = cdname + "*"
    End If
Data1.RecordSource = "SELECT * FROM Articulos WHERE Proveedor like '" & cdname & "'order by Articulo"
    Data1.Refresh
DBGrid1.Refresh
Dim intRecord As Integer
Dim intField As Integer
  intRecord = Data1.Recordset.RecordCount
  intField = Data1.Recordset.Fields.Count
  'call the procedure here...
  Call AdjustDataGridColumns _
  (DBGrid1, Data1, intRecord, intField, True)
End If
End Sub

Private Sub Command1_Click()
Data1.DatabaseName = App.Path & ("\VENTAS.mdb")
Data1.RecordSource = "SELECT * FROM Articulos order by Articulo"
Data1.Refresh
DBGrid1.Refresh
Dim intRecord As Integer
Dim intField As Integer
  intRecord = Data1.Recordset.RecordCount
  intField = Data1.Recordset.Fields.Count
  'call the procedure here...
  Call AdjustDataGridColumns _
  (DBGrid1, Data1, intRecord, intField, True)
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & ("\VENTAS.mdb")
Data1.RecordSource = "SELECT * FROM Articulos order by Articulo"
Data2.DatabaseName = App.Path & ("\VENTAS.mdb")
Data2.RecordSource = "SELECT * FROM Proveedores "
Data2.Refresh

Data1.Refresh
  Dim intRecord As Integer
Dim intField As Integer
  intRecord = Data1.Recordset.RecordCount
  intField = Data1.Recordset.Fields.Count
  'call the procedure here...
  Call AdjustDataGridColumns _
  (DBGrid1, Data1, intRecord, intField, True)
Do While Not Data2.Recordset.EOF
               ' start from beginning of the records
               ' work to the end of the records
        
            
            Combo1.AddItem IIf(IsNull(Data2.Recordset("Empresa")), "", Data2.Recordset("Empresa")), i
            'List2.AddItem IIf(IsNull(Data1.Recordset("Origen")), "", Data1.Recordset("Origen"))
                      Data2.Recordset.MoveNext
            i = i + 1
            
        Loop
If Combo1.ListIndex - 1 Then
Combo1.Text = "PROVEEDOR"
Else
Combo1.ListIndex = 0
End If
End Sub
