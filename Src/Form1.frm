VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDisplay1 
      Caption         =   "Load"
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton cmdexit1 
      Caption         =   "Exit"
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   5280
      Width           =   1815
   End
   Begin FPSpread.vaSpread fpSpread1 
      Height          =   3855
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   8175
      _Version        =   393216
      _ExtentX        =   14420
      _ExtentY        =   6800
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDisplay1_Click()
Form2.Show
End Sub

Private Sub Form_Load()
' Display only 4 columns and rows
fpSpread1.VisibleRows = 4
fpSpread1.VisibleCols = 4
fpSpread1.AutoSize = True
' Put text in the first column
fpSpread1.Col = 1
fpSpread1.Row = 1
fpSpread1.Text = "1st Quarter"
fpSpread1.Row = 2
fpSpread1.Text = "2nd Quarter"
fpSpread1.Row = 3
fpSpread1.Text = "3rd Quarter"
fpSpread1.Row = 4
fpSpread1.Text = "4th Quarter"
' Set up the locked cells
fpSpread1.LockBackColor = RGB(192, 192, 192)
fpSpread1.Col = 1
fpSpread1.Row = 1
fpSpread1.Col2 = 1
fpSpread1.Row2 = 4
fpSpread1.BlockMode = True
fpSpread1.Lock = True
fpSpread1.BlockMode = False
fpSpread1.Protect = True



If fpSpread1.Col = 1 Then
fpSpread1.Text = "right"

Else

fpSpread1.Text = "WRONG"

End If
End Sub
