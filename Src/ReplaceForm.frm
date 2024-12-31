VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ReplaceForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Replace"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame ReplaceFrame 
      Caption         =   "Replace Method"
      Height          =   1245
      Left            =   60
      TabIndex        =   13
      Top             =   2610
      Width           =   8085
      Begin VB.ComboBox MethodStyle 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "ReplaceForm.frx":0000
         Left            =   3690
         List            =   "ReplaceForm.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   570
         Width           =   1395
      End
      Begin VB.OptionButton MethodOptionCompute 
         Caption         =   "Replace with the current value of the column"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   600
         Width           =   3525
      End
      Begin VB.TextBox ReplaceText 
         Height          =   315
         Left            =   1890
         TabIndex        =   15
         Top             =   210
         Width           =   4365
      End
      Begin VB.OptionButton MethodOptionValue 
         Caption         =   "Replace with value"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   1665
      End
      Begin MSMask.MaskEdBox MethodValue 
         Height          =   315
         Left            =   5160
         TabIndex        =   18
         Top             =   570
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   "_"
      End
   End
   Begin VB.Frame CriteriaFrame 
      Caption         =   "Criteria"
      Height          =   1125
      Left            =   60
      TabIndex        =   9
      Top             =   1470
      Width           =   8085
      Begin VB.TextBox FindText 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3030
         TabIndex        =   12
         Top             =   480
         Width           =   4365
      End
      Begin VB.OptionButton CriteriaOptionSelective 
         Caption         =   "Replace only where column equals"
         Height          =   255
         Left            =   210
         TabIndex        =   11
         Top             =   510
         Width           =   2805
      End
      Begin VB.OptionButton CriteriaOptionAll 
         Caption         =   "Replace all values in column"
         Height          =   255
         Left            =   210
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   2655
      End
   End
   Begin VB.Frame RangeFrame 
      Caption         =   "Record Range"
      Height          =   975
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   8085
      Begin MSMask.MaskEdBox RangeStart 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   510
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   "_"
      End
      Begin VB.OptionButton RangeOptionRange 
         Caption         =   "Record #"
         Height          =   345
         Left            =   210
         TabIndex        =   5
         Top             =   510
         Width           =   1065
      End
      Begin VB.OptionButton RangeOptionAll 
         Caption         =   "All Records"
         Height          =   345
         Left            =   210
         TabIndex        =   4
         Top             =   210
         Value           =   -1  'True
         Width           =   1275
      End
      Begin MSMask.MaskEdBox RangeStop 
         Height          =   315
         Left            =   2010
         TabIndex        =   7
         Top             =   510
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "to"
         Height          =   255
         Left            =   1770
         TabIndex        =   8
         Top             =   570
         Width           =   225
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   1530
      Top             =   3870
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReplaceForm.frx":003C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReplaceForm.frx":035E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolBar 
      Height          =   840
      Left            =   60
      TabIndex        =   0
      Top             =   3840
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   1482
      ButtonWidth     =   1191
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "E&xit"
            Description     =   "Close this window"
            Object.ToolTipText     =   "Close this window"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Replace"
            Object.ToolTipText     =   "Replace values as specified"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label FindInLabel 
      Caption         =   "Replace In:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   1275
   End
   Begin VB.Label FindInColumnLabel 
      Caption         =   "collabel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1380
      TabIndex        =   1
      Top             =   60
      Width           =   2775
   End
End
Attribute VB_Name = "ReplaceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: Shooter Data Manager (SDM)                                **
'**                                                                        **
'** Module.....: ReplaceForm.frm - The SDM Replace Function (Edit Menu).   **
'**                                                                        **
'** Description: This form presents a simple "Replace" Window.             **
'**                                                                        **
'** History....:                                                           **
'**    03/20/02 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2002 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit
Private SpreadSheet As FPSpread.vaSpread                'Pointer to spreadsheet to work in

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Form_Load                                             **
'**                                                                        **
'**  Description..:  Initialize form controls                              **
'**                                                                        **
'****************************************************************************
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":Form_Load()"
    '---- Initialize the form controls
    RangeOptionAll_Click                                'Default to all records
    CriteriaOptionAll_Click                             'Default to replace all
    MethodOptionValue_Click                             'Default method to replace with text
    ViewLog.Log logdebug, Me.Name & ":Form_Load Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "ReplaceForm:Form_Load", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub


'****************************************************************************
'**                                                                        **
'**  Procedure....:  ToolBar_ButtonClick                                   **
'**                                                                        **
'**  Description..:  Provide form exit on toolbar.                         **
'**                                                                        **
'****************************************************************************
Private Sub ToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":ToolBar_ButtonClick(Button=" + str(Button.Index) + ")"
    Select Case Button.Index
        Case 1
            Me.Hide
        Case 2
            If MsgBox("Are you sure?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "Replace") = vbYes Then
                ReplaceValues
            End If
    End Select
    ViewLog.Log logdebug, Me.Name & ":ToolBar_ButtonClick Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "ReplaceForm:ToolBar_ButtonClick", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

Private Sub RangeOptionAll_Click()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":RangeOptionAll_Click()"
    '---- User selected all records
    RangeStart.Enabled = False
    RangeStart.BackColor = &H8000000F
    RangeStop.Enabled = False
    RangeStop.BackColor = &H8000000F
    ViewLog.Log logdebug, Me.Name & ":RangeOptionAll_Click Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "ReplaceForm:RangeOptionAll_Click", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

Private Sub RangeOptionRange_Click()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":RangeOptionRange_Click()"
    '---- User selected Range of Records
    RangeStart.Enabled = True
    RangeStart.BackColor = &H80000005
    RangeStop.Enabled = True
    RangeStop.BackColor = &H80000005
    RangeStart.SetFocus
    ViewLog.Log logdebug, Me.Name & ":RangeOptionRange_Click Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "ReplaceForm:RangeOptionRange_Click", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

Private Sub CriteriaOptionAll_Click()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":CriteriaOptionAll_Click()"
    FindText.Enabled = False
    FindText.BackColor = &H8000000F
    ViewLog.Log logdebug, Me.Name & ":CriteriaOptionAll_Click Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "ReplaceForm:CriteriaOptionAll_Click", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

Private Sub CriteriaOptionSelective_Click()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":CriteriaOptionSelective_Click()"
    FindText.Enabled = True
    FindText.BackColor = &H80000005
    ViewLog.Log logdebug, Me.Name & ":CriteriaOptionSelective_Click Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "ReplaceForm:CriteriaOptionSelective_Click", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

Private Sub MethodOptionValue_Click()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":MethodOPtionValue_Click()"
    ReplaceText.Enabled = True
    ReplaceText.BackColor = &H80000005
    MethodStyle.Enabled = False
    MethodStyle.BackColor = &H8000000F
    MethodValue.Enabled = False
    MethodValue.BackColor = &H8000000F
    ViewLog.Log logdebug, Me.Name & ":MethodOptionValue_Click Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "ReplaceForm:MethodOptionValue_Click", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

Private Sub MethodOptionCompute_Click()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":MethodOptionCompute()"
    ReplaceText.Enabled = False
    ReplaceText.BackColor = &H8000000F
    MethodStyle.Enabled = True
    MethodStyle.BackColor = &H80000005
    MethodValue.Enabled = True
    MethodValue.BackColor = &H80000005
    ViewLog.Log logdebug, Me.Name & ":MethodOptionCompute Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "ReplaceForm:MethodOptionCompute_Click", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

Public Sub DoReplace(SourceSpreadSheet As FPSpread.vaSpread)
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":DoReplace(SourceSpreadSheet)"
    '---- Set pointer to source spreadsheet (where replace will take place)
    Set SpreadSheet = SourceSpreadSheet
    SpreadSheet.Row = SpreadSheet.ActiveRow
    SpreadSheet.Col = SpreadSheet.ActiveCol
    FindText.MaxLength = Len(SourceSpreadSheet.Text)                'Max length of column
    FindText.Text = SourceSpreadSheet.Text                          'Get the value of the column
    ReplaceText.MaxLength = FindText.MaxLength
    ReplaceText.Text = SourceSpreadSheet.Text
    SpreadSheet.Row = 1
    FindInColumnLabel.Caption = SpreadSheet.Text                    'Column to work in
    Me.Show vbModal
    ViewLog.Log logdebug, Me.Name & ":DoReplace Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "ReplaceForm:DoReplace", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

Private Sub ReplaceValues()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":ReplaceValues()"
    Dim ColLength As Integer
    Dim ColText As String
    Dim StartRow As Integer
    Dim EndRow As Integer
    Dim MaxRows As Integer
    Dim RowNum As Integer
    Dim ReplacedIt As Boolean
            
    If SpreadSheet.Name = "FrameSpread" Or SpreadSheet.Name = "PackageSpread" Then
        SpreadSheet.Col = SpreadSheet.ActiveCol                     'Go to Active Column
        SpreadSheet.Row = SpreadSheet.ActiveRow                     'Go to Active row
        ColLength = Len(SpreadSheet.TypePicMask)
        ColText = SpreadSheet.Text
        MaxRows = SpreadSheet.MaxRows
        '--- Validate row range
        StartRow = IIf(Me.RangeOptionAll.Value = True, 1, Val(Me.RangeStart.Text))
        EndRow = IIf(Me.RangeOptionAll.Value = True, SpreadSheet.MaxRows, Val(Me.RangeStop.Text))
        If EndRow < StartRow Then
            MsgBox "Ending record must be higher than starting record.", vbApplicationModal + vbInformation + vbOKOnly, "Error"
            Exit Sub
        End If
        '--- Loop for records in range
        ReplacedIt = False
        For RowNum = StartRow To EndRow
            SpreadSheet.Row = RowNum
            
            If CriteriaOptionSelective.Value = True Then            'If selective replacement chosen
                If Trim(SpreadSheet.Text) = Trim(FindText.Text) Then
                    '--- Perform Replace
                    ReplacedIt = True
                    ReplaceAsSpecified
                End If
            Else
                '--- perform replace
                ReplacedIt = True
                ReplaceAsSpecified
            End If
        Next
        If ReplacedIt = False Then
            MsgBox "Replace did not find any matching records.", vbApplicationModal + vbInformation + vbOKOnly, "Replace"
        End If
    End If
    ViewLog.Log logdebug, Me.Name & ":ReplaceValues Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "ReplaceForm:ReplaceValues", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

Private Sub ReplaceAsSpecified()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":ReplaceAsSpecified()"
    If Me.MethodOptionValue.Value = True Then
        SpreadSheet.Text = ReplaceText.Text                         'Simply replace with value
    Else                                                            'Else replace with computation
        Select Case Me.MethodStyle.ListIndex                        'Get computation type
            Case 0
                SpreadSheet.Text = Trim(str(Val(SpreadSheet.Text) + Val(MethodValue.Text)))
            Case 1
                SpreadSheet.Text = Trim(str(Val(SpreadSheet.Text) - Val(MethodValue.Text)))
            Case 2
                SpreadSheet.Text = Trim(str(Val(SpreadSheet.Text) * Val(MethodValue.Text)))
            Case 3
                SpreadSheet.Text = Trim(str(Val(SpreadSheet.Text) / Val(MethodValue.Text)))
        End Select
    End If
    '--- Handle Padding of Numeric Columns
    If SpreadSheet.Name = "PackageSpread" Then
        '--- All package columns are simply numeric - 2 digit
        SpreadSheet.Text = Format(Val(SpreadSheet.Text), "00")
    Else
        '--- Special handling for frame columns
        Select Case SpreadSheet.Col
            Case 1              'Pkg
            If SpreadSheet.Text <> "BL" And SpreadSheet.Text <> "SL" And SpreadSheet.Text <> "NP" Then
                SpreadSheet.Text = Format(Val(SpreadSheet.Text), "00")
            End If
        End Select
    End If
    ViewLog.Log logdebug, Me.Name & ":ReplaceAsSpecified Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "ReplaceForm:ReplaceAsSpecified", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub
