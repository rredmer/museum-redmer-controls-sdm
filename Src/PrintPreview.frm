VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form PrintPreview 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Preview"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame OrientationFrame 
      Caption         =   "Orientation"
      Height          =   825
      Left            =   1410
      TabIndex        =   2
      Top             =   8040
      Width           =   2115
      Begin VB.OptionButton OptionLandscape 
         Caption         =   "Landscape"
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   510
         Width           =   1875
      End
      Begin VB.OptionButton OptionPortrait 
         Caption         =   "Portrait"
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   210
         Value           =   -1  'True
         Width           =   1875
      End
   End
   Begin FPSpread.vaSpreadPreview vaSpreadPreview 
      Height          =   7905
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   10875
      _Version        =   393216
      _ExtentX        =   19182
      _ExtentY        =   13944
      _StockProps     =   96
      BorderStyle     =   1
      AllowUserZoom   =   -1  'True
      GrayAreaColor   =   8421504
      GrayAreaMarginH =   720
      GrayAreaMarginType=   0
      GrayAreaMarginV =   720
      PageBorderColor =   8388608
      PageBorderWidth =   2
      PageShadowColor =   0
      PageShadowWidth =   2
      PageViewPercentage=   100
      PageViewType    =   0
      ScrollBarH      =   1
      ScrollBarV      =   1
      ScrollIncH      =   360
      ScrollIncV      =   360
      PageMultiCntH   =   1
      PageMultiCntV   =   1
      PageGutterH     =   -1
      PageGutterV     =   -1
      ScriptEnhanced  =   0   'False
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   10290
      Top             =   8070
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrintPreview.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrintPreview.frx":0322
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrintPreview.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrintPreview.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrintPreview.frx":0FD0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolBar 
      Height          =   840
      Left            =   60
      TabIndex        =   1
      Top             =   8010
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   1482
      ButtonWidth     =   1032
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
            Caption         =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "PrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: Shooter Data Manager (SDM)                                **
'**                                                                        **
'** Module.....: PrintPreview.frm - Preview a spreadsheet for printing.    **
'**                                                                        **
'** Description: This form provides the FarPoint Preview ActiveX Control.  **
'**    **** DO NOT CALL WITH FORM.SHOW, USE PREPARETOPREVIEW METHOD ***    **
'**                                                                        **
'** History....:                                                           **
'**    03/20/02 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2002 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit

Private SpreadSheetToPrint As FPSpread.vaSpread                     'Pointer to spreadsheet to print

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Form_Activate                                         **
'**                                                                        **
'**  Description..:  Updates the preview control on form activation.       **
'**                                                                        **
'****************************************************************************
Private Sub Form_Activate()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":Form_Activate()"
    UpdatePreview
    ViewLog.Log logdebug, Me.Name & ":Form_Activate Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "PrintPreview:Form_Activate", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  OptionLandscape_Click                                 **
'**                                                                        **
'**  Description..:  Changes the preview mode to landscape & updates screen**
'**                                                                        **
'****************************************************************************
Private Sub OptionLandscape_Click()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":OptionLandscape_Click()"
    SpreadSheetToPrint.PrintOrientation = PrintOrientationLandscape
    UpdatePreview
    ViewLog.Log logdebug, Me.Name & ":OptionLandscape_Click()"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "PrintPreview:OptionLandscape_Click", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  OptionPortrait_Click                                  **
'**                                                                        **
'**  Description..:  Changes the preview mode to portrait & updates screen **
'**                                                                        **
'****************************************************************************
Private Sub OptionPortrait_Click()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":OptionPortrait_Click()"
    SpreadSheetToPrint.PrintOrientation = PrintOrientationPortrait
    UpdatePreview
    ViewLog.Log logdebug, Me.Name & ":OptionPortrait_Click"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "PrintPreview:OptionPortrait_Click", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
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
        Case 1                                          'Hide
            Me.Hide
        Case 2
        If MsgBox("Are you sure?", vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "Print") = vbYes Then
            SpreadSheetToPrint.PrintSheet 1
        End If
    End Select
    ViewLog.Log logdebug, Me.Name & ":ToolBar_ButtonClick Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "PrintPreview:ToolBar_ButtonClick", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  UpdatePreview                                         **
'**                                                                        **
'**  Description..:  Updates the preview control with current settings     **
'**                                                                        **
'****************************************************************************
Private Sub UpdatePreview()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":UpdatePreview()"
    With vaSpreadPreview
        .hWndSpread = SpreadSheetToPrint.hWnd
        .PageViewType = PageViewTypeMultiplePages
        .AllowUserZoom = True
    End With
    ViewLog.Log logdebug, Me.Name & ":UpdatePreview Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "PrintPreview:UpdatePreview", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  PrepareToPreview                                      **
'**                                                                        **
'**  Description..:  Called from Main Menu - performs the preview!         **
'**                                                                        **
'****************************************************************************
Public Sub PrepareToPreview(SourceSpreadSheet As FPSpread.vaSpread)
    On Error GoTo ErrorHandler
    
    ViewLog.Log logdebug, Me.Name & ":PrepareToPreview(SourceSpreadSheet)"
    
    Set SpreadSheetToPrint = SourceSpreadSheet
    Me.Show vbModal
    
    ViewLog.Log logdebug, Me.Name & ":PrepareToPreview(SourceSpreadSheet)"
    
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "PrintPreview:PrepareToPreview", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub
