VERSION 5.00
Begin VB.Form Articuloseditar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Articulos - EDITAR"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "EDITAR"
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.Data Data2 
         Caption         =   "Data2"
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
         Top             =   3600
         Width           =   2535
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\utilesvbasic\Prg Ventas\VENTAS.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   3120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Articulos"
         Top             =   3600
         Width           =   2655
      End
      Begin VB.TextBox txtcodigo 
         DataField       =   "Codigo"
         DataSource      =   "Data1"
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
         Left            =   1200
         TabIndex        =   11
         Top             =   360
         Width           =   1695
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
         Left            =   4800
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtarticulo 
         DataField       =   "Articulo"
         DataSource      =   "Data1"
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
         Left            =   1200
         TabIndex        =   9
         Top             =   1200
         Width           =   5055
      End
      Begin VB.TextBox txtpreciodecosto 
         DataField       =   "Preciodecosto"
         DataSource      =   "Data1"
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
         Left            =   2160
         TabIndex        =   8
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtstockmin 
         DataField       =   "Stockminimo"
         DataSource      =   "Data1"
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
         Left            =   5640
         TabIndex        =   7
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtstockmax 
         DataField       =   "Stockmaximo"
         DataSource      =   "Data1"
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
         Left            =   5640
         TabIndex        =   6
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtpreciodeventa 
         DataField       =   "Preciodeventa"
         DataSource      =   "Data1"
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
         Left            =   2160
         TabIndex        =   5
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtstockreal 
         DataField       =   "stockactual"
         DataSource      =   "Data1"
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
         Left            =   5640
         TabIndex        =   4
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox txtreferencia 
         DataField       =   "Referencia"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtproveedor 
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Guardar"
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
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
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
         Left            =   3360
         TabIndex        =   18
         Top             =   360
         Width           =   1215
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
         Left            =   360
         TabIndex        =   17
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio de Costo:  $"
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
         TabIndex        =   16
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Mínimo:"
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
         Left            =   4200
         TabIndex        =   15
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Máximo:"
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
         Left            =   4200
         TabIndex        =   14
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio de VENTA: $"
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
         TabIndex        =   13
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Real:"
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
         Left            =   4440
         TabIndex        =   12
         Top             =   2880
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Articuloseditar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
Data1.DatabaseName = App.Path & ("\VENTAS.mdb")
Data1.RecordSource = "SELECT * FROM Articulos order by Codigo"
Dim cdname As String
cdname = Articulos.List1.List(Articulos.List1.ListIndex)
   If Trim(cdname) <> "" Then
        If (Right(cdname, 1) <> "*") Then
        cdname = cdname '+ "*"
    End If
MsgBox "ADVERTENCIA!!!,los cambios realizados se actualizarán cuando salga de ARTICULOS y eliga nuevamente NUEVO ARTICULO", vbInformation, "INFORMACION DE ACTUALIZACIÓN"
If txtstockreal.Text <= txtstockmin.Text Then
txtreferencia.Text = "PEDIR STOCK"
Else
txtreferencia.Text = ""
End If
Articulos.Data1.RecordSource = "SELECT * FROM Articulos WHERE Articulo like '" & cdname & "'order by Articulo"
    Articulos.Data1.Refresh
   
    
            
Articulos.txtpreciodeventa.Text = IIf(IsNull(Articulos.Data1.Recordset("Preciodecosto")), "", Articulos.Data1.Recordset("Preciodecosto"))
Articulos.txtarticulo.Text = IIf(IsNull(Articulos.Data1.Recordset("Articulo")), "", Articulos.Data1.Recordset("Articulo"))
Articulos.txtcodigo.Text = IIf(IsNull(Articulos.Data1.Recordset("Codigo")), "", Articulos.Data1.Recordset("Codigo"))
Articulos.txtpreciodecosto.Text = IIf(IsNull(Articulos.Data1.Recordset("Preciodecosto")), "", Articulos.Data1.Recordset("Preciodecosto"))
Articulos.txtstockreal.Text = IIf(IsNull(Articulos.Data1.Recordset("Stockactual")), "", Articulos.Data1.Recordset("Stockactual"))
Articulos.txtstockmax.Text = IIf(IsNull(Articulos.Data1.Recordset("Stockmaximo")), "", Articulos.Data1.Recordset("Stockmaximo"))
Articulos.txtstockmin.Text = IIf(IsNull(Articulos.Data1.Recordset("Stockminimo")), "", Articulos.Data1.Recordset("Stockminimo"))
Articulos.txtreferencia.Text = IIf(IsNull(Articulos.Data1.Recordset("Referencia")), "", Articulos.Data1.Recordset("Referencia"))
Articulos.Data1.UpdateRecord
Articulos.Data1.Refresh
Unload Me

End If
Unload Me
Articulos.Command1.Enabled = True
Articulos.Command3.Enabled = False

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & ("\VENTAS.mdb")
Data1.RecordSource = "SELECT * FROM Articulos order by Codigo"
Data2.DatabaseName = App.Path & ("\VENTAS.mdb")
Data2.RecordSource = "SELECT * FROM Proveedores "
Data1.Refresh
Data2.Refresh

Call cargarbaseproveedores
End Sub

Public Sub cargarbaseproveedores()
Data2.DatabaseName = App.Path & ("\VENTAS.mdb")
Data2.RecordSource = "SELECT * FROM Proveedores "
Data2.Refresh
Combo1.Clear
Do While Not Data2.Recordset.EOF
               ' start from beginning of the records
               ' work to the end of the records
        
            
            Combo1.AddItem IIf(IsNull(Data2.Recordset("Empresa")), "", Data2.Recordset("Empresa")), i
            'List2.AddItem IIf(IsNull(Data1.Recordset("Origen")), "", Data1.Recordset("Origen"))
                      Data2.Recordset.MoveNext
            i = i + 1
            
        Loop

End Sub


