VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form librodecuentas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CAJA DIARIA"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin VB.Data Data5 
         Caption         =   "Data5"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   7320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Librodeventas"
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Data Data4 
         Caption         =   "Data4"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   7320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Librodeventas"
         Top             =   2520
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "librodecuentas.frx":0000
         Height          =   1335
         Left            =   240
         OleObjectBlob   =   "librodecuentas.frx":0014
         TabIndex        =   21
         Top             =   2280
         Width           =   9615
      End
      Begin VB.Data Data3 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1920
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1560
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Cajagastar"
         Top             =   1200
         Visible         =   0   'False
         Width           =   2775
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
         TabIndex        =   17
         Top             =   4920
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
         TabIndex        =   16
         Top             =   4920
         Width           =   1455
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
         TabIndex        =   15
         Top             =   4920
         Width           =   1455
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
         Height          =   1695
         Left            =   240
         TabIndex        =   8
         Top             =   3600
         Width           =   4335
         Begin VB.TextBox txtfecha 
            Height          =   285
            Left            =   2640
            TabIndex        =   25
            Top             =   240
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "BUSCAR..."
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
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtaño 
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
            Height          =   285
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   14
            Top             =   1320
            Width           =   615
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
            TabIndex        =   13
            Top             =   1320
            Width           =   1455
         End
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
            Left            =   1920
            TabIndex        =   12
            Top             =   840
            Width           =   735
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
            TabIndex        =   11
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtdia 
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
            Height          =   285
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   10
            Top             =   360
            Width           =   495
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
            TabIndex        =   9
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Totales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5160
         TabIndex        =   4
         Top             =   3600
         Width           =   4695
         Begin VB.TextBox txtsaldo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3240
            TabIndex        =   7
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtingresos 
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
            TabIndex        =   6
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtgastos 
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
            TabIndex        =   5
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "SALDO"
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
            Left            =   3480
            TabIndex        =   20
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "INGRESOS"
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
            Left            =   2160
            TabIndex        =   19
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "GASTOS"
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
            Left            =   960
            TabIndex        =   18
            Top             =   240
            Width           =   855
         End
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "librodecuentas.frx":09DF
         Height          =   1335
         Left            =   240
         OleObjectBlob   =   "librodecuentas.frx":09F3
         TabIndex        =   3
         Top             =   720
         Width           =   9615
      End
      Begin VB.TextBox Text1 
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
         Left            =   8520
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "INGRESOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "GASTOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Anterior:   $"
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
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "librodecuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.DatabaseName = App.Path & ("\Ventas.mdb")
Data1.RecordSource = "SELECT * FROM Cajagastar order by Ndecaja"
Data1.Refresh

End Sub

Private Sub Command2_Click()
MsgBox "En construccion"

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
If txtdia.Text = "" Then
MsgBox "Por Favor indique un día válido...", vbInformation, "ERROR- EN DIA"
txtdia.SetFocus
ElseIf Combo1.Text = "" Then
MsgBox "Por Favor indique un mes válido...", vbInformation, "ERROR- EN MES"
Combo1.SetFocus
ElseIf txtaño.Text = "" Then
MsgBox "Por Favor indique un año válido...", vbInformation, "ERROR- EN AÑO"
txtaño.SetFocus
Else
txtfecha.Text = txtdia.Text & "/" & Combo1.Text & "/" & txtaño.Text
Data1.DatabaseName = App.Path & ("\VENTAS.mdb")
Data1.RecordSource = "SELECT * FROM librodecuentas order by Ndecaja"
Dim cdname
cdname = txtfecha.Text
Data1.RecordSource = "SELECT * FROM Librodecuentas WHERE Fecha like '" & cdname & "'order by Fecha"
Data1.Refresh
DBGrid1.Refresh
Dim intRecord As Integer
Dim intField As Integer
  intRecord = Data1.Recordset.RecordCount
  intField = Data1.Recordset.Fields.Count
    Call AdjustDataGridColumns _
  (DBGrid1, Data1, intRecord, intField, True)
txtdia.Enabled = False
txtaño.Enabled = False
Combo1.Enabled = False
End If
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & ("\Ventas.mdb")
Data1.RecordSource = "SELECT * FROM Librodecuentas order by Ndecaja"
Data2.DatabaseName = App.Path & ("\Ventas.mdb")
Data2.RecordSource = "SELECT * FROM Librodecuentas order by Ndecaja"
Data3.DatabaseName = App.Path & ("\Ventas.mdb")
Data3.RecordSource = "SELECT * FROM Librodecuentas order by Ndecaja"
Data4.DatabaseName = App.Path & ("\Ventas.mdb")
Data4.RecordSource = "SELECT * FROM Librodeventas order by Ndecaja"
Data5.DatabaseName = App.Path & ("\Ventas.mdb")
Data5.RecordSource = "SELECT * FROM Librodeventas order by Ndecaja"

Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.Refresh
Data5.Refresh
If Data1.Recordset.RecordCount = 0 Then Exit Sub
        
        With Data1
            
            .Recordset.Edit
            .Recordset.MoveFirst
End With
Dim intField As Integer
 Dim intRecord As Integer
 Dim intField2 As Integer
 Dim intRecord2 As Integer

  intRecord = Data1.Recordset.RecordCount
  intField = Data1.Recordset.Fields.Count
  intRecord2 = Data4.Recordset.RecordCount
  intField2 = Data4.Recordset.Fields.Count
  'call the procedure here...
  Call AdjustDataGridColumns _
  (DBGrid1, Data1, intRecord, intField, True)
Call AdjustDataGridColumns _
  (DBGrid2, Data4, intRecord2, intField2, True)
Data2.DatabaseName = App.Path & "\VENTAS.mdb"
Data2.RecordSource = "Select * from Librodecuentas "

Data3.DatabaseName = App.Path & "\VENTAS.mdb"
Data3.RecordSource = "Select * from Librodecuentas "

Data4.DatabaseName = App.Path & "\VENTAS.mdb"
Data4.RecordSource = "Select * from Librodeventas "

Data5.DatabaseName = App.Path & "\VENTAS.mdb"
Data5.RecordSource = "SELECT  sum(Total) As totallibrodeventas from Librodeventas "
Data5.Refresh
Data2.DatabaseName = App.Path & "\VENTAS.mdb"
Data2.RecordSource = "SELECT  sum(Gastos) As totalgastos from Librodecuentas "
Data2.Refresh
txtgastos.Text = IIf(IsNull(Data2.Recordset("totalgastos")), "", Data2.Recordset("totalgastos"))

txtingresos.Text = IIf(IsNull(Data5.Recordset("totallibrodeventas")), "", Data5.Recordset("totallibrodeventas"))
txtsaldo.Text = Val(txtingresos.Text) - Val(txtgastos.Text)
If Val(txtsaldo.Text) < Val(txtgastos) Then
txtsaldo.ForeColor = &HFF&
Else
txtsaldo.ForeColor = &H0&
End If
txtdia.Text = Format(Date, "DD")
txtaño.Text = Format(Date, "YYYY")
txtdia.Enabled = False
txtaño.Enabled = False
Combo1.AddItem "01"
Combo1.AddItem "02"
Combo1.AddItem "03"
Combo1.AddItem "04"
Combo1.AddItem "05"
Combo1.AddItem "06"
Combo1.AddItem "07"
Combo1.AddItem "08"
Combo1.AddItem "09"
Combo1.AddItem "10"
Combo1.AddItem "11"
Combo1.AddItem "12"
Combo1.ListIndex = 0
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
txtdia.Enabled = True
Else
txtdia.Enabled = False

End If

End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Combo1.Enabled = True
Else
Combo1.Enabled = False
End If

End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
txtaño.Enabled = True
Else
txtaño.Enabled = False
End If

End Sub
