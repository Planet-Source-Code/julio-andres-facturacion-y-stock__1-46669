VERSION 5.00
Begin VB.Form Articulos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artículos"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
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
      Height          =   3630
      ItemData        =   "Articulos.frx":0000
      Left            =   360
      List            =   "Articulos.frx":0002
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Artículos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Editar"
         Enabled         =   0   'False
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
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Proveedores"
         Top             =   4200
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txtproveedor 
         DataSource      =   "Data1"
         Height          =   285
         Left            =   4080
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtreferencia 
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2520
         TabIndex        =   20
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   2640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Articulos"
         Top             =   3840
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
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
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Archivar"
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
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox txtstockreal 
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
         Left            =   7800
         TabIndex        =   17
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox txtpreciodeventa 
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
         Left            =   4320
         TabIndex        =   15
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtstockmax 
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
         Left            =   7800
         TabIndex        =   13
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox txtstockmin 
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
         Left            =   7800
         TabIndex        =   11
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtpreciodecosto 
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
         Left            =   4320
         TabIndex        =   9
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtarticulo 
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
         Left            =   3360
         TabIndex        =   7
         Top             =   1440
         Width           =   5055
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
         Left            =   6960
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   3360
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Para editar haga click"
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
         TabIndex        =   23
         Top             =   360
         Width           =   2295
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
         Left            =   6600
         TabIndex        =   16
         Top             =   3120
         Width           =   1215
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
         Left            =   2400
         TabIndex        =   14
         Top             =   2640
         Width           =   1935
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
         Left            =   6360
         TabIndex        =   12
         Top             =   2640
         Width           =   1455
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
         Left            =   6360
         TabIndex        =   10
         Top             =   2160
         Width           =   1455
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
         Left            =   2400
         TabIndex        =   8
         Top             =   2160
         Width           =   2055
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
         Left            =   2520
         TabIndex        =   6
         Top             =   1440
         Width           =   975
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
         Left            =   5520
         TabIndex        =   4
         Top             =   600
         Width           =   1215
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
         Left            =   2520
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "Articulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
txtproveedor.Text = Combo1.Text
End Sub

Private Sub Command1_Click()
'On Error GoTo error
Data1.DatabaseName = App.Path & ("\VENTAS.mdb")
Data1.RecordSource = "SELECT * FROM Articulos order by articulo"
If txtstockreal.Text <= txtstockmin.Text Then
txtreferencia.Text = "Pedir Stock"
Else
txtreferencia.Text = ""
End If
If Val(txtstockreal.Text) > Val(txtstockmax.Text) Or Val(txtstockreal.Text) = Val(txtstockmax.Text) Then
MsgBox "El stock real es MAYOR o IGUAL al stock máximo", vbInformation, "ERROR"
txtstockreal.SetFocus
Else
Data1.Recordset.AddNew


With Data1

.Recordset.Fields("Codigo").Value = UCase(txtcodigo.Text)
.Recordset.Fields("Articulo").Value = UCase(txtarticulo.Text)
.Recordset.Fields("Preciodecosto").Value = txtpreciodecosto.Text
.Recordset.Fields("Preciodeventa").Value = txtpreciodeventa.Text
.Recordset.Fields("Stockminimo").Value = txtstockmin.Text
.Recordset.Fields("Stockmaximo").Value = txtstockmax.Text
.Recordset.Fields("stockactual").Value = txtstockreal.Text
.Recordset.Fields("Referencia").Value = UCase(txtreferencia.Text)
.Recordset.Fields("Proveedor").Value = (txtproveedor.Text)
.Refresh
End With
List1.Clear
Principal.List2.Clear
Do While Not Data1.Recordset.EOF
               ' start from beginning of the records
               ' work to the end of the records
        
            
            List1.AddItem IIf(IsNull(Data1.Recordset("Articulo")), "", Data1.Recordset("Articulo")), i
            Principal.List2.AddItem IIf(IsNull(Data1.Recordset("Articulo")), "", Data1.Recordset("Articulo"))
                      Data1.Recordset.MoveNext
            i = i + 1
            
        Loop

 txtcodigo.Text = ""
 txtarticulo.Text = ""
 txtpreciodecosto.Text = ""
 txtpreciodeventa.Text = ""
 txtstockmin.Text = ""
 txtstockmax.Text = ""
 txtstockreal.Text = ""
 txtreferencia.Text = ""
 txtproveedor.Text = ""
End If

End Sub

Private Sub Command2_Click()
txtcodigo.Text = ""
 txtarticulo.Text = ""
 txtpreciodecosto.Text = ""
 txtpreciodeventa.Text = ""
 txtstockmin.Text = ""
 txtstockmax.Text = ""
 txtstockreal.Text = ""
 txtreferencia.Text = ""
 txtproveedor.Text = ""
Unload Me
End Sub

Private Sub Command3_Click()
Dim cdname As String
cdname = List1.List(List1.ListIndex)
   If Trim(cdname) <> "" Then
        If (Right(cdname, 1) <> "*") Then
        cdname = cdname '+ "*"
    End If
Articuloseditar.Data1.RecordSource = "SELECT * FROM Articulos WHERE Articulo like '" & cdname & "'order by Articulo"
    Articuloseditar.Data1.Refresh
   
    
            
Articuloseditar.txtpreciodeventa.Text = IIf(IsNull(Articuloseditar.Data1.Recordset("Preciodecosto")), "", Articuloseditar.Data1.Recordset("Preciodecosto"))
Articuloseditar.txtarticulo.Text = IIf(IsNull(Articuloseditar.Data1.Recordset("Articulo")), "", Articuloseditar.Data1.Recordset("Articulo"))
Articuloseditar.txtcodigo.Text = IIf(IsNull(Articuloseditar.Data1.Recordset("Codigo")), "", Articuloseditar.Data1.Recordset("Codigo"))
Articuloseditar.txtpreciodecosto.Text = IIf(IsNull(Articuloseditar.Data1.Recordset("Preciodecosto")), "", Articuloseditar.Data1.Recordset("Preciodecosto"))
Articuloseditar.txtstockreal.Text = IIf(IsNull(Articuloseditar.Data1.Recordset("Stockactual")), "", Articuloseditar.Data1.Recordset("Stockactual"))
Articuloseditar.txtstockmax.Text = IIf(IsNull(Articuloseditar.Data1.Recordset("Stockmaximo")), "", Articuloseditar.Data1.Recordset("Stockmaximo"))
Articuloseditar.txtstockmin.Text = IIf(IsNull(Articuloseditar.Data1.Recordset("Stockminimo")), "", Articuloseditar.Data1.Recordset("Stockminimo"))
Articuloseditar.txtreferencia.Text = IIf(IsNull(Articuloseditar.Data1.Recordset("Referencia")), "", Articuloseditar.Data1.Recordset("Referencia"))

Articuloseditar.Show

End If
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & ("\VENTAS.mdb")
Data1.RecordSource = "SELECT * FROM Articulos order by Codigo"
Data2.DatabaseName = App.Path & ("\VENTAS.mdb")
Data2.RecordSource = "SELECT * FROM Proveedores "
Data2.Refresh
Data1.Refresh
If Data1.Recordset.RecordCount = 0 Then Exit Sub
        
        With Data1
            
            .Recordset.Edit
            .Recordset.MoveFirst
End With
Do While Not Data1.Recordset.EOF
               ' start from beginning of the records
               ' work to the end of the records
        
            
            List1.AddItem IIf(IsNull(Data1.Recordset("Articulo")), "", Data1.Recordset("Articulo")), i
            'List2.AddItem IIf(IsNull(Data1.Recordset("Origen")), "", Data1.Recordset("Origen"))
                      Data1.Recordset.MoveNext
            i = i + 1
            
        Loop
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

Private Sub List1_Click()
Command1.Enabled = False
Command3.Enabled = True
Dim cdname As String
cdname = List1.List(List1.ListIndex)
   If Trim(cdname) <> "" Then
        If (Right(cdname, 1) <> "*") Then
        cdname = cdname '+ "*"
    End If
Data2.RecordSource = "SELECT * FROM Articulos WHERE Articulo like '" & cdname & "'order by Articulo"
    Data2.Refresh
   
    
            
txtpreciodecosto.Text = IIf(IsNull(Data2.Recordset("Preciodecosto")), "", Data2.Recordset("Preciodecosto"))
txtarticulo.Text = IIf(IsNull(Data2.Recordset("Articulo")), "", Data2.Recordset("Articulo"))
txtcodigo.Text = IIf(IsNull(Data2.Recordset("Codigo")), "", Data2.Recordset("Codigo"))
txtpreciodeventa.Text = IIf(IsNull(Data2.Recordset("Preciodeventa")), "", Data2.Recordset("Preciodeventa"))
txtstockreal.Text = IIf(IsNull(Data2.Recordset("Stockactual")), "", Data2.Recordset("Stockactual"))
txtstockmax.Text = IIf(IsNull(Data2.Recordset("Stockmaximo")), "", Data2.Recordset("Stockmaximo"))
txtstockmin.Text = IIf(IsNull(Data2.Recordset("Stockminimo")), "", Data2.Recordset("Stockminimo"))
txtreferencia.Text = IIf(IsNull(Data2.Recordset("Referencia")), "", Data2.Recordset("Referencia"))

End If
End Sub

