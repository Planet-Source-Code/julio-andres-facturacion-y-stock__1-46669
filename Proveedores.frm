VERSION 5.00
Begin VB.Form Proveedores 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PROVEEDORES"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.TextBox txtlocalidad 
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
         Left            =   7200
         TabIndex        =   20
         Top             =   2760
         Width           =   855
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Proveedores"
         Top             =   3600
         Width           =   1935
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
         Left            =   3600
         MaxLength       =   7
         TabIndex        =   17
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtcuil 
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
         Left            =   6360
         TabIndex        =   16
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtempresa 
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
         Left            =   3600
         TabIndex        =   15
         Top             =   1320
         Width           =   4455
      End
      Begin VB.TextBox txtdireccion 
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
         Left            =   3600
         TabIndex        =   14
         Top             =   1800
         Width           =   4455
      End
      Begin VB.TextBox txttelefono 
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
         Left            =   3600
         TabIndex        =   13
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txtemail 
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
         Left            =   3600
         TabIndex        =   12
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox txtcp 
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
         Left            =   7200
         TabIndex        =   11
         Top             =   2280
         Width           =   855
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
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3480
         Width           =   1455
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
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3480
         Width           =   1455
      End
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
         Height          =   3210
         ItemData        =   "Proveedores.frx":0000
         Left            =   240
         List            =   "Proveedores.frx":0002
         TabIndex        =   1
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Localidad:"
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
         Left            =   6240
         TabIndex        =   21
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Si quiere puede seguir el orden del último codigo escrito :"
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
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label lblcodigo 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Left            =   2760
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "CU.I.T:"
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
         Left            =   5640
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa:"
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
         Left            =   2640
         TabIndex        =   6
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
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
         Left            =   2640
         TabIndex        =   5
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono:"
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
         Left            =   2640
         TabIndex        =   4
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail:"
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
         Left            =   2880
         TabIndex        =   3
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "C.p:"
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
         Left            =   6720
         TabIndex        =   2
         Top             =   2280
         Width           =   495
      End
   End
End
Attribute VB_Name = "Proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Data1.DatabaseName = App.Path & ("\VENTAS.mdb")
Data1.RecordSource = "SELECT * FROM Proveedores order by Codigo"
  If txtcodigo.Text = "" Then
  MsgBox "Se olvidó de completar el casillero de CODIGO", vbCritical, "CLIENTES - ERROR"
  txtcodigo.SetFocus
ElseIf txtcuil.Text = "" Then
  MsgBox "Se olvidó de completar el casillero de CUIL", vbCritical, "CLIENTES - ERROR"
  txtcuil.SetFocus
ElseIf txtempresa.Text = "" Then
  MsgBox "Se olvidó de completar el casillero de EMPRESA", vbCritical, "CLIENTES - ERROR"
  txtempresa.SetFocus
ElseIf txtdireccion.Text = "" Then
  MsgBox "Se olvidó de completar el casillero de DIRECCION", vbCritical, "CLIENTES - ERROR"
  txtdireccion.SetFocus
  Else
Open App.Path & "\Codigo.txt" For Output As #1
Print #1, LTrim(RTrim(txtcodigo.Text))
Close #1
Data1.Recordset.AddNew
With Data1
.Recordset.Fields("Codigo").Value = txtcodigo.Text
.Recordset.Fields("Cuil").Value = txtcuil.Text
.Recordset.Fields("Empresa").Value = UCase(txtempresa.Text)
.Recordset.Fields("Direccion").Value = UCase(txtdireccion.Text)
.Recordset.Fields("Tel").Value = txttelefono.Text
.Recordset.Fields("Email").Value = UCase(txtemail.Text)
.Recordset.Fields("Cp").Value = UCase(txtcp.Text)
.Recordset.Fields("Localidad").Value = UCase(txtlocalidad.Text)
End With
Data1.Refresh
List1.Clear
Principal.List1.Clear
Do While Not Data1.Recordset.EOF
               ' start from beginning of the records
               ' work to the end of the records
        
            
            List1.AddItem IIf(IsNull(Data1.Recordset("Empresa")), "", Data1.Recordset("Empresa")), i
            Principal.List1.AddItem IIf(IsNull(Data1.Recordset("Empresa")), "", Data1.Recordset("Empresa"))
                      Data1.Recordset.MoveNext
            i = i + 1
            
        Loop

 txtcodigo.Text = ""
 txtcuil.Text = ""
 txtempresa.Text = ""
 txtdireccion.Text = ""
 txttelefono.Text = ""
 txtemail.Text = ""
 txtcp.Text = ""
 
Open App.Path & "\Codigo.txt" For Input As #1
lblcodigo.Caption = Input(LOF(1), 1)
Close #1
txtcodigo.Text = lblcodigo.Caption
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo error
Data1.DatabaseName = App.Path & ("\VENTAS.mdb")
Data1.RecordSource = "SELECT * FROM Proveedores order by Codigo"
Data1.Refresh
'If Data1.Recordset.RecordCount = 0 Then Exit Sub
        
 '       With Data1
            
  '          .Recordset.Edit
   '         .Recordset.MoveFirst
'End With

Do While Not Data1.Recordset.EOF
               ' start from beginning of the records
               ' work to the end of the records
        
            
            List1.AddItem IIf(IsNull(Data1.Recordset("Empresa")), "", Data1.Recordset("Empresa")), i
            'List2.AddItem IIf(IsNull(Data1.Recordset("Origen")), "", Data1.Recordset("Origen"))
                      Data1.Recordset.MoveNext
            i = i + 1
            
        Loop
Open App.Path & "\Codigo.txt" For Input As #1
lblcodigo.Caption = Input(LOF(1), 1)
Close #1
txtcodigo.Text = lblcodigo.Caption
error:
If Err.Number = 53 Then
Open App.Path & "\Codigo.txt" For Output As #1
Print #1, "0000001"
Close #1
Open App.Path & "\Codigo.txt" For Input As #1
lblcodigo.Caption = Input(LOF(1), 1)
Close #1
End If
End Sub

Private Sub Label8_Click()

End Sub
