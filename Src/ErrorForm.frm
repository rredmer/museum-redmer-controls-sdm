VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ErrorForm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Problem Report"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog SaveDialog 
      Left            =   2640
      Top             =   3810
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "txt"
      DialogTitle     =   "Save Error Report"
      FileName        =   "SDM_Error.txt"
      Filter          =   "txt"
   End
   Begin VB.TextBox ErrorText 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   720
      Width           =   6345
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   5850
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ErrorForm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ErrorForm.frx":0322
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ErrorForm.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ErrorForm.frx":0CB6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolBar 
      Height          =   840
      Left            =   60
      TabIndex        =   3
      Top             =   3780
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   1482
      ButtonWidth     =   1032
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "E&xit"
            Description     =   "Close this window"
            Object.ToolTipText     =   "Close this window"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Copy"
            Object.ToolTipText     =   "Copy to Windows Clipboard"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "E-Mail"
            Object.ToolTipText     =   "E-Mail report to technical support"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Object.ToolTipText     =   "Save"
            Object.Tag             =   "Save report to disk"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "problem are shown below:"
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
      TabIndex        =   1
      Top             =   330
      Width           =   6375
   End
   Begin VB.Label TitleLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "A problem has occured in the program.  The details of this"
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
      TabIndex        =   0
      Top             =   60
      Width           =   6375
   End
End
Attribute VB_Name = "ErrorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: Shooter Data Manager (SDM)                                **
'**                                                                        **
'** Module.....: ErrorForm.frm - The application error handler.            **
'**                                                                        **
'** Description: This form provides a consistent error reporting method.   **
'**                                                                        **
'** History....:                                                           **
'**    03/20/02 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2002 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ReportError                                           **
'**                                                                        **
'**  Description..:  This is the main error reporting routine for the      **
'**                  entire application.  It's main purpose is to ensure   **
'**                  that messages are reported in consistent format.      **
'**                                                                        **
'****************************************************************************
Public Sub ReportError(ErrProcName As String, ErrNum As Integer, ErrDllNum As Integer, ErrSource As String, ErrDescript As String, ErrShowWin As Boolean)
    On Error GoTo ErrorHandler
    '---- Format the Error Message (this provides consistency throughout the program)
    Dim ErrMsg As String
    
    ErrMsg = ""
    ErrMsg = ErrMsg & "Report time.: " & Now & vbCrLf                       'The time the error occured
    ErrMsg = ErrMsg & "Error number: " & Trim(str(ErrNum)) & vbCrLf         'The VB Error Number
    ErrMsg = ErrMsg & "DLL Error...: " & Trim(str(ErrDllNum)) & vbCrLf      'The last DLL Error that occured
    ErrMsg = ErrMsg & "In procedure: " & ErrProcName & vbCrLf               'The Procure in which the error occured
    ErrMsg = ErrMsg & "Source. ....: " & ErrSource & vbCrLf                 'The source of the error
    ErrMsg = ErrMsg & "Description.: " & ErrDescript                        'The VB description of the error

    ErrorText.Text = ErrMsg
    ErrorText.Refresh
    
    If ErrShowWin = True Then
        Me.Show vbModal
    End If
    '---- Also report error to application log file
    ViewLog.Log LogError, "Error (" & Trim(str(ErrNum)) & ")(" & Trim(str(ErrDllNum)) & ") In Procedure (" & ErrProcName & ") source (" & ErrSource & ") description (" & ErrDescript & ")."
    ErrMsg = ""
    Exit Sub
ErrorHandler:
    MsgBox "An error occured formatting the error message.  Please contact technical support.", vbApplicationModal + vbOKOnly + vbInformation, "Error"
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  ToolBar_ButtonClick                                   **
'**                                                                        **
'**  Description..:  This routine handles user clicks on the toolbar.      **
'**                                                                        **
'****************************************************************************
Private Sub ToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo ErrorHandler
    
    ViewLog.Log logdebug, Me.Name & ":ToolBar_ButtonClick(Button=" + str(Button.Index) + ")"
    Select Case Button.Index
        Case 1                                          'Exit button
            Me.Hide                                     'Simply hide this form - it needs to stay loaded for possible future errors
        Case 2                                          'Copy to clipboard
            Clipboard.Clear                             'Clear the clipboard contents
            Clipboard.SetText ErrorText.Text, vbCFText  'Save the text to the clipboard
        Case 3                                          'Email Error Report (Future)
            MsgBox "Please use save button to save to file then e-mail file to technical support.", vbApplicationModal + vbOKOnly + vbInformation, "Feature coming soon"
        Case 4                                          'Save to disk
            SaveErrorTextToFile                         'This routine uses the VB SaveDialog to prompt for saving file
    End Select
    
    ViewLog.Log logdebug, Me.Name & ":ToolBar_ButtonClick Exiting"
    
    Exit Sub
ErrorHandler:
    MsgBox "An error occured in the Error Toolbar.  Please contact technical support.", vbApplicationModal + vbOKOnly + vbInformation, "Error"
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  SaveErrorTextToFile                                   **
'**                                                                        **
'**  Description..:  This routine stores error text to user-named file.    **
'**                                                                        **
'****************************************************************************
Private Sub SaveErrorTextToFile()
    On Error GoTo ErrorHandler
    
    ViewLog.Log logdebug, Me.Name & ":SaveErrorTextToFile()"
    
    '---- Show the save file dialog (VB6 has custom dialogs, .NET uses a common windows dialog)
    SaveDialog.ShowSave
    If SaveDialog.FileName <> "" Then
        Dim ErrFileSystem As New Scripting.FileSystemObject     'Pointer to Error File System Object
        Dim ErrFile As Scripting.TextStream                     'Pointer to Error Report File
        
        '---- Get pointer to current log file and create it if necessary
        Set ErrFile = ErrFileSystem.CreateTextFile(SaveDialog.FileName, True, False)
        ErrFile.Write ErrorText.Text
        ErrFile.Close
        Set ErrFile = Nothing
        Set ErrFileSystem = Nothing
        MsgBox "Created Error Report: " & SaveDialog.FileName, vbApplicationModal + vbOKOnly + vbInformation, "Finished"
    End If
    
    ViewLog.Log logdebug, Me.Name & ":SaveErrorTextToFile Exiting"
    
    Exit Sub
ErrorHandler:
    MsgBox "An error occured writing to text file.  Please contact technical support.", vbApplicationModal + vbOKOnly + vbInformation, "Error"
End Sub
