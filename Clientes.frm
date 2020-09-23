VERSION 5.00
Begin VB.Form Clientes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.TextBox txtformadepago 
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
         Left            =   2880
         TabIndex        =   20
         Top             =   3360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   "C:\utilesvbasic\Prg Ventas\VENTAS.MDB"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Formadepago"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\utilesvbasic\Prg Ventas\VENTAS.MDB"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Clientes"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1695
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aceptar"
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3240
         Width           =   1455
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
         Left            =   3720
         TabIndex        =   17
         Top             =   2520
         Width           =   855
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
         Left            =   6240
         TabIndex        =   16
         Top             =   2040
         Width           =   1815
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
         Left            =   3720
         TabIndex        =   15
         Top             =   2040
         Width           =   1695
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
         Left            =   3720
         TabIndex        =   14
         Top             =   1560
         Width           =   4335
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
         Left            =   3720
         TabIndex        =   13
         Top             =   1080
         Width           =   4335
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Clientes.frx":0000
         Left            =   6720
         List            =   "Clientes.frx":0002
         TabIndex        =   12
         Top             =   2640
         Width           =   1455
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
         MaxLength       =   11
         TabIndex        =   5
         Top             =   480
         Width           =   1695
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
         Left            =   3720
         MaxLength       =   7
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3180
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2535
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
         TabIndex        =   28
         Top             =   240
         Width           =   1695
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
         TabIndex        =   27
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8160
         TabIndex        =   26
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8040
         TabIndex        =   25
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8040
         TabIndex        =   24
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8040
         TabIndex        =   23
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5400
         TabIndex        =   22
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "* NO OLVIDARSE DE COMPLETAR LOS CAMPOS REQUERIDOS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2760
         TabIndex        =   21
         Top             =   3000
         Width           =   5655
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pago:"
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
         Left            =   5280
         TabIndex        =   11
         Top             =   2640
         Width           =   1695
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
         Left            =   3240
         TabIndex        =   10
         Top             =   2520
         Width           =   495
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
         Left            =   5640
         TabIndex        =   9
         Top             =   2040
         Width           =   735
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
         Left            =   2760
         TabIndex        =   8
         Top             =   2040
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
         Left            =   2760
         TabIndex        =   7
         Top             =   1560
         Width           =   975
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
         Left            =   2760
         TabIndex        =   6
         Top             =   1080
         Width           =   975
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
         TabIndex        =   4
         Top             =   480
         Width           =   735
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
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim codigo As String
Private Sub Combo1_Click()
'Text1.Text = Combo1.Text
End Sub

Private Sub Command1_Click()
Data1.DatabaseName = App.Path & ("\Ventas.mdb")
Data1.RecordSource = "SELECT * FROM Clientes order by Codigo"
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
txtformadepago.Text = Combo1.Text
.Recordset.Fields("Codigo").Value = txtcodigo.Text
.Recordset.Fields("Cuil").Value = txtcuil.Text
.Recordset.Fields("Empresa").Value = txtempresa.Text
.Recordset.Fields("Direccion").Value = txtdireccion.Text
.Recordset.Fields("Tel").Value = txttelefono.Text
.Recordset.Fields("Email").Value = txtemail.Text
.Recordset.Fields("Cp").Value = txtcp.Text
.Recordset.Fields("Formadepago").Value = txtformadepago.Text
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
 txtformadepago.Text = ""
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
Data1.DatabaseName = App.Path & ("\Ventas.mdb")
Data1.RecordSource = "SELECT * FROM Clientes order by Codigo"
Data2.DatabaseName = App.Path & ("\Ventas.mdb")
Data2.RecordSource = "SELECT * FROM Formadepago "
Data2.Refresh
Data1.Refresh
If Data1.Recordset.RecordCount = 0 Then Exit Sub
        
        With Data1
            
            .Recordset.Edit
            .Recordset.MoveFirst
End With
If Data2.Recordset.RecordCount = 0 Then Exit Sub
        
        With Data2
            
            .Recordset.Edit
            .Recordset.MoveFirst
End With

Do While Not Data1.Recordset.EOF
               ' start from beginning of the records
               ' work to the end of the records
        
            
            List1.AddItem IIf(IsNull(Data1.Recordset("Empresa")), "", Data1.Recordset("Empresa")), i
            'List2.AddItem IIf(IsNull(Data1.Recordset("Origen")), "", Data1.Recordset("Origen"))
                      Data1.Recordset.MoveNext
            i = i + 1
            
        Loop
Call cargarbaseformadepago
Combo1.ListIndex = 0
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
Public Sub cargarbaseformadepago()
Data2.DatabaseName = App.Path & ("\Ventas.mdb")
Data2.RecordSource = "SELECT * FROM Formadepago "
Data2.Refresh
Combo1.Clear
Do While Not Data2.Recordset.EOF
               ' start from beginning of the records
               ' work to the end of the records
        
            
            Combo1.AddItem IIf(IsNull(Data2.Recordset("Pago")), "", Data2.Recordset("Pago")), i
            'List2.AddItem IIf(IsNull(Data1.Recordset("Origen")), "", Data1.Recordset("Origen"))
                      Data2.Recordset.MoveNext
            i = i + 1
            
        Loop

End Sub

