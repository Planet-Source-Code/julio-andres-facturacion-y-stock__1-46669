VERSION 5.00
Begin VB.Form Gastar 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GASTO"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6360
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
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   "C:\utilesvbasic\Prg Ventas\VENTAS.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Librodecuentas"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   2655
         Left            =   5640
         TabIndex        =   14
         Top             =   120
         Width           =   3135
         Begin VB.TextBox txtsaldo 
            DataSource      =   "Data2"
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
            Left            =   1800
            TabIndex        =   20
            Top             =   2280
            Width           =   975
         End
         Begin VB.TextBox txtndecaja2 
            DataSource      =   "Data2"
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
            Left            =   120
            TabIndex        =   19
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtconcepto2 
            DataSource      =   "Data2"
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
            Left            =   120
            TabIndex        =   18
            Top             =   1680
            Width           =   4095
         End
         Begin VB.TextBox txtimporte2 
            DataSource      =   "Data2"
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
            Left            =   120
            TabIndex        =   17
            Top             =   2280
            Width           =   1455
         End
         Begin VB.TextBox txtfecha 
            DataSource      =   "Data2"
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo"
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
            Left            =   1800
            TabIndex        =   24
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label9 
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
            Left            =   120
            TabIndex        =   23
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label8 
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
            Left            =   120
            TabIndex        =   22
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label7 
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
            Left            =   120
            TabIndex        =   21
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Cajagastar"
         Top             =   2760
         Width           =   1815
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
         TabIndex        =   12
         Top             =   2520
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
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtaño 
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
         Left            =   5040
         MaxLength       =   4
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtmes 
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
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   9
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtdia 
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
         Left            =   4080
         MaxLength       =   2
         TabIndex        =   8
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtimporte 
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
         Left            =   1680
         TabIndex        =   6
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtconcepto 
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
         Left            =   1680
         TabIndex        =   5
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox txtndecaja 
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
         Left            =   1680
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "EJ: 200,00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   1920
         Width           =   1095
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
         TabIndex        =   7
         Top             =   720
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
         TabIndex        =   3
         Top             =   1920
         Width           =   1335
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
         TabIndex        =   2
         Top             =   1320
         Width           =   1095
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
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Gastar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.DatabaseName = App.Path & ("\Ventas.mdb")
Data1.RecordSource = "SELECT * FROM Cajagastar order by Ndecaja"
Data2.DatabaseName = App.Path & ("\Ventas.mdb")
Data2.RecordSource = "SELECT * FROM Librodecuentas order by Ndecaja"
txtfecha.Text = txtdia.Text & "/" & txtmes.Text & "/" & txtaño.Text
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
Data2.Recordset.AddNew
With Data1
.Recordset.Fields("Ndecaja").Value = txtndecaja2.Text
.Recordset.Fields("Dia").Value = txtdia.Text
.Recordset.Fields("Mes").Value = txtmes.Text
.Recordset.Fields("Año").Value = txtaño.Text
.Recordset.Fields("Concepto").Value = UCase(txtconcepto.Text)
.Recordset.Fields("Importe").Value = txtimporte.Text
.Refresh
End With
With Data2
.Recordset.Fields("Ndecaja").Value = txtndecaja.Text
.Recordset.Fields("Concepto").Value = UCase(txtconcepto2.Text)
.Recordset.Fields("Gastos").Value = txtimporte2.Text
.Recordset.Fields("Fecha").Value = txtfecha.Text
.Refresh
End With
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
Data1.RecordSource = "SELECT * FROM Cajagastar order by Ndecaja"
Data2.DatabaseName = App.Path & ("\Ventas.mdb")
Data2.RecordSource = "SELECT * FROM Librodecuentas order by Ndecaja"

Data1.Refresh
Data2.Refresh
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
txtdia.Text = Format(Date, "DD")
txtmes.Text = Format(Date, "MM")
txtaño.Text = Format(Date, "YYYY")
End Sub

Private Sub txtconcepto_Change()
txtconcepto2.Text = txtconcepto.Text
End Sub

Private Sub txtimporte_Change()
txtimporte2.Text = txtimporte.Text
End Sub

Private Sub txtndecaja_Change()
txtndecaja2.Text = txtndecaja.Text
End Sub
