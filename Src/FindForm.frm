VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FindForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox FindText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1380
      TabIndex        =   2
      Top             =   540
      Width           =   6075
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   720
      Top             =   1290
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FindForm.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolBar 
      Height          =   840
      Left            =   60
      TabIndex        =   0
      Top             =   1290
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1482
      ButtonWidth     =   1032
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "E&xit"
            Description     =   "Close this window"
            Object.ToolTipText     =   "Close this window"
            ImageIndex      =   1
         EndProperty
      EndProperty
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
      Height          =   405
      Left            =   1380
      TabIndex        =   4
      Top             =   60
      Width           =   2775
   End
   Begin VB.Label FindInLabel 
      Caption         =   "Find In:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   1275
   End
   Begin VB.Label FindLabel 
      Caption         =   "Find Value:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   60
      TabIndex        =   1
      Top             =   570
      Width           =   1275
   End
End
Attribute VB_Name = "FindForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: Shooter Data Manager (SDM)                                **
'**                                                                        **
'** Module.....: FindForm.frm - The SDM Find Function (Edit Menu).         **
'**                                                                        **
'** Description: This form presents a simple "Find" Window.                **
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
'**  Procedure....:  FindText_KeyDown                                      **
'**                                                                        **
'**  Description..:  Provide form exit on Enter key.                       **
'**                                                                        **
'****************************************************************************
Private Sub FindText_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":FindText_KeyDown(KeyCode=" + str(KeyCode) + ",Shift=" + str(Shift)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            DoFind
            Me.Hide
    End Select
    ViewLog.Log logdebug, Me.Name & ":FindText_KeyDown Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "FindForm:FindText_KeyDown", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
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
            DoFind
            Me.Hide
    End Select
    ViewLog.Log logdebug, Me.Name & ":ToolBar_ButtonClick Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "FindForm:ToolBar_ButtonClick", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

Public Sub DoFindText(SourceSpreadSheet As FPSpread.vaSpread)
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":DoFindText(SourceSpreadSheet)"
    '---- Set pointer to source spreadsheet (where replace will take place)
    Set SpreadSheet = SourceSpreadSheet
    SpreadSheet.Row = SpreadSheet.ActiveRow
    SpreadSheet.Col = SpreadSheet.ActiveCol
    FindText.MaxLength = Len(SpreadSheet.Text)                      'Max length of column
    FindText.Text = SpreadSheet.Text                                'Get the value of the column
    SpreadSheet.Row = 0
    FindInColumnLabel.Caption = SpreadSheet.Text                    'Column to work in
    Me.Show vbModal
    ViewLog.Log logdebug, Me.Name & ":DoFindText Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "ReplaceForm:DoReplace", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

Private Sub DoFind()

    Dim ColLength As Integer
    Dim StartRow As Integer
    Dim MaxRows As Integer
    Dim RowNum As Integer
    Dim FoundIt As Boolean
            
    If SpreadSheet.Name = "FrameSpread" Or SpreadSheet.Name = "PackageSpread" Then
        StartRow = SpreadSheet.ActiveRow
        MaxRows = SpreadSheet.MaxRows
        
        FoundIt = False
        For RowNum = (StartRow + 1) To MaxRows
            SpreadSheet.Row = RowNum
            If SpreadSheet.Text = FindForm.FindText.Text Then
                SpreadSheet.SetActiveCell SpreadSheet.Col, RowNum
                FoundIt = True
                Exit For
            End If
        Next
        
        If FoundIt = False Then
            MsgBox "Value not found.", vbApplicationModal + vbInformation + vbOKOnly, "Find"
        End If
        
    End If
End Sub


