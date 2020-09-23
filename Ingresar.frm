VERSION 5.00
Begin VB.Form Ingresar 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INGRESAR"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Concepto de Caja"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.TextBox txtimporte 
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
         Left            =   1680
         TabIndex        =   12
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Cajaingresar"
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox txtndecaja 
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
         Left            =   1680
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtconcepto 
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
         Left            =   1680
         TabIndex        =   6
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox txtdia 
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
         Left            =   4080
         MaxLength       =   2
         TabIndex        =   5
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtmes 
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
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   4
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtaño 
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
         Left            =   5040
         MaxLength       =   4
         TabIndex        =   3
         Top             =   720
         Width           =   735
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
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2520
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Caja:"
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
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto:"
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
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe  :   $ "
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
         TabIndex        =   9
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label4 
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
         Left            =   3360
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Ingresar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.DatabaseName = App.Path & ("\Ventas.mdb")
Data1.RecordSource = "SELECT * FROM Cajaingresar order by Ndecaja"
 ' If txtcodigo.Text = "" Then
  'MsgBox "Se olvidó de completar el casillero de CODIGO", vbCritical, "CLIENTES - ERROR"
  'txtcodigo.SetFocus
'ElseIf txtcuil.Text = "" Then
 ' MsgBox "Se olvidó de completar el casillero de CUIL", vbCritical, "CLIENTES - ERROR"
  'txtcuil.SetFocus
'ElseIf txtempresa.Text = "" Then
 ' MsgBox "Se olvidó de completar el casillero de EMPRESA", vbCritical, "CLIENTES - ERROR"
  'txtempresa.SetFocus
'ElseIf txtdireccion.Text = "" Then
 ' MsgBox "Se olvidó de completar el casillero de DIRECCION", vbCritical, "CLIENTES - ERROR"
  'txtdireccion.SetFocus
  'Else
Data1.Recordset.AddNew
With Data1
.Recordset.Fields("Ndecaja").Value = txtndecaja.Text
.Recordset.Fields("Dia").Value = txtdia.Text
.Recordset.Fields("Mes").Value = txtmes.Text
.Recordset.Fields("Año").Value = txtaño.Text
.Recordset.Fields("Concepto").Value = UCase(txtconcepto.Text)
.Recordset.Fields("Importe").Value = txtimporte.Text

End With
Data1.Recordset.Update
txtndecaja.Text = ""
txtdia.Text = ""
txtmes.Text = ""
txtaño.Text = ""
txtconcepto.Text = ""
txtimporte.Text = ""
txtndecaja.SetFocus
End Sub

Private Sub Command2_Click()
txtndecaja.Text = ""
txtdia.Text = ""
txtmes.Text = ""
txtaño.Text = ""
txtconcepto.Text = ""
txtimporte.Text = ""
Unload Me
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & ("\Ventas.mdb")
Data1.RecordSource = "SELECT * FROM Cajaingresar order by Ndecaja"
Data1.Refresh
If Data1.Recordset.RecordCount = 0 Then Exit Sub
        
        With Data1
            
            .Recordset.Edit
            .Recordset.MoveFirst
End With
txtdia.Text = Format(Date, "DD")
txtmes.Text = Format(Date, "MM")
txtaño.Text = Format(Date, "YYYY")
End Sub
