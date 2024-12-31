VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form GoToRecordForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Go to Record"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2790
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   2790
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox RecordNum 
      Height          =   435
      Left            =   1590
      TabIndex        =   2
      Top             =   30
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   767
      _Version        =   393216
      MaxLength       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "000"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   780
      Top             =   780
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
            Picture         =   "GoToRecordForm.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolBar 
      Height          =   840
      Left            =   60
      TabIndex        =   1
      Top             =   750
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
   Begin VB.Label GoToRecordLabel 
      Caption         =   "Go to record#"
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
      TabIndex        =   0
      Top             =   60
      Width           =   1515
   End
End
Attribute VB_Name = "GoToRecordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: Shooter Data Manager (SDM)                                **
'**                                                                        **
'** Module.....: GoToRecordForm.frm - The SDM Go To Record Form (Edit Menu)**
'**                                                                        **
'** Description: This form provides the Go To Record routine.              **
'**                                                                        **
'** History....:                                                           **
'**    03/20/02 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2002 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit

'****************************************************************************
'**                                                                        **
'**  Procedure....:  RecordNum_KeyDown                                     **
'**                                                                        **
'**  Description..:  Provide form exit on Enter key.                       **
'**                                                                        **
'****************************************************************************
Private Sub RecordNum_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandler
    
    ViewLog.Log logdebug, Me.Name & ":RecordNum_KeyDown(KeyCode=" & str(KeyCode) & ",Shift=" & str(Shift)
    
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            Me.Hide
    End Select
    
    ViewLog.Log logdebug, Me.Name & ":RecordNum_KeyDown Exiting"
    
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "GoToRecordForm:RecordNum_KeyDown", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
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
    End Select
    
    ViewLog.Log logdebug, Me.Name & ":ToolBar_ButtonClick Exiting"
    
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "GoToRecordForm:ToolBar_ButtonClick", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub
