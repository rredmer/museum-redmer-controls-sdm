VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form TimingForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Application Timing Subsystem"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList 
      Left            =   2010
      Top             =   2490
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
            Picture         =   "TimingForm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TimingForm.frx":0322
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Height          =   840
      Left            =   60
      TabIndex        =   0
      Top             =   2250
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   1482
      ButtonWidth     =   1508
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
            Caption         =   "Test Timer"
            Description     =   "Test the application timer"
            Object.ToolTipText     =   "Test the application timer"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "TimingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: Shooter Data Manager (SDM)                                **
'**                                                                        **
'** Module.....: TimingForm.frm - The application timer interface.         **
'**                                                                        **
'** Description: This form provides a consistent interface to the timer(s) **
'**                                                                        **
'** History....:                                                           **
'**    03/20/02 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2002 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit

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
            MsgBox "Coming soon!"
    End Select
    ViewLog.Log logdebug, Me.Name & ":ToolBar_ButtonClick Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "TimingForm:ToolBar_ButtonClick", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**                                                                        **
'**                     APPLICATION TIMING SUBSYSTEM                       **
'**                                                                        **
'**                                                                        **
'****************************************************************************
'****************************************************************************
'**                                                                        **
'** Subroutine.: GetTimer                                                  **
'**                                                                        **
'** Description: This routine provides # seconds elapsed since app started.**
'**                                                                        **
'****************************************************************************
Public Function GetTimer() As Single
    GetTimer = Timer                               'Return timer value
End Function
'****************************************************************************
'**                                                                        **
'** Subroutine.: Delay                                                     **
'**                                                                        **
'** Description: This routine delays for specified # milliseconds.         **
'**                                                                        **
'****************************************************************************
Public Sub Delay(DelayInMilliseconds As Single)
    Dim StartTime As Single, CurrentTime As Single
    StartTime = Timer
    Do
        CurrentTime = Timer
        DoEvents
    Loop Until CurrentTime >= StartTime + (DelayInMilliseconds / 1000#) Or (CurrentTime < StartTime And CurrentTime > StartTime + (DelayInMilliseconds / 1000#) - 86400)
End Sub


