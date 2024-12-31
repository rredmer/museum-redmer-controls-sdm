VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form FileManager 
   Caption         =   "File Manager"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   ScaleHeight     =   9935.747
   ScaleMode       =   0  'User
   ScaleWidth      =   14430
   StartUpPosition =   3  'Windows Default
   Begin FPSpread.vaSpread FileMangerSpread 
      Height          =   6975
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   14415
      _Version        =   393216
      _ExtentX        =   25426
      _ExtentY        =   12303
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
      ScrollBars      =   2
      SpreadDesigner  =   "FileManager.frx":0000
   End
   Begin MSComctlLib.Toolbar ToolBar 
      Height          =   840
      Left            =   120
      TabIndex        =   0
      Top             =   7800
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   1482
      ButtonWidth     =   1429
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "E&xit"
            Description     =   "Close this window"
            Object.ToolTipText     =   "Close this window"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Add"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Delete "
            Object.ToolTipText     =   "Erase lens drawer setup"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "A&rchive"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Notes"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Tracking"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   13200
      Top             =   8280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   13680
      Top             =   8160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FileManager.frx":4626
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FileManager.frx":4948
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FileManager.frx":4C6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FileManager.frx":4F8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FileManager.frx":52A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FileManager.frx":55C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "SDM Files"
      Height          =   7575
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   14535
   End
End
Attribute VB_Name = "FileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'****************************************************************************
'**                                                                        **
'** Project....: Shooter Data Manager (SDM)                                **
'**                                                                        **
'** Module.....: FileManager.frm - Handles SDM Files.                      **
'**                                                                        **
'** Description: Lets user select from SDM Files and make changes          **
'**                                                                        **
'** History....:                                                           **
'**    03/20/02 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2002 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit

'Private FileManagerSpread As vaSpread

Private Sub FileMangerSpread_Advance(ByVal AdvanceNext As Boolean)


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
        Case 1                                  'Exit
            Me.Hide
        Case 2                                  'Add
            'FileManagerSpread.MaxRows = 501
            'FileManagerSpread.InsertRows 500, 1
            
        
            'With FileManagerSpread
            '    .InsertRows 3, 1
            
        Case 3                                  'Delete
        
           ' FileManagerSpread.DeleteRows
           ' End With
        Case 4                                  'Archive
         
        Case 5                                  'Notes
        
        Case 6                                  'Tracking
        
           
    End Select
    
   ViewLog.Log logdebug, Me.Name & ":ToolBar_ButtonClick Exiting"
    
    Exit Sub
    
ErrorHandler:
    ErrorForm.ReportError "FileManager:ToolBar_ButtonClick", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

    

