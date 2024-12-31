VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form LensDrawerSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lens Drawer Setup"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   2550
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.FileListBox FileList 
      Height          =   4575
      Left            =   60
      Pattern         =   "*.len"
      TabIndex        =   2
      Top             =   360
      Width           =   3255
   End
   Begin FPSpread.vaSpread LensSpreadsheet 
      Height          =   4575
      Left            =   3360
      TabIndex        =   1
      Top             =   360
      Width           =   8325
      _Version        =   393216
      _ExtentX        =   14684
      _ExtentY        =   8070
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
      SpreadDesigner  =   "LensDrawerSetup.frx":0000
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   1920
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LensDrawerSetup.frx":01D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LensDrawerSetup.frx":04F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LensDrawerSetup.frx":0818
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolBar 
      Height          =   840
      Left            =   60
      TabIndex        =   0
      Top             =   5010
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1482
      ButtonWidth     =   1032
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "E&xit"
            Description     =   "Close this window"
            Object.ToolTipText     =   "Close this window"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Object.ToolTipText     =   "Create new lens drawer setup"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Erase"
            Object.ToolTipText     =   "Erase lens drawer setup"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Label LensFileLabel 
      Caption         =   "Lens Drawer Setup Files:"
      Height          =   285
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   2355
   End
End
Attribute VB_Name = "LensDrawerSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: Shooter Data Manager (SDM)                                **
'**                                                                        **
'** Module.....: LensDrawerSetup.frm - Setup Form for Lens Drawer Files.   **
'**                                                                        **
'** Description: This form provides configuration using FarPoint Spread.   **
'**                                                                        **
'** History....:                                                           **
'**    03/20/02 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2002 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit

Private Const DefaultFile As String = "DefaultLens.Len"     'Default name of Lens Drawer Files
Private FileLocation As String                              'Location of Lens Drawer Files
Private CurrentFile As String                               'Name of the current file
Private FileSystem As New Scripting.FileSystemObject        'Pointer to Log File System Object

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
        Case 2                                  'New
            PromptForFileName
            
        Case 3                                  'Erase
            If FileList.FileName = DefaultFile Then
                MsgBox "Can't erase default lens file.", vbApplicationModal + vbInformation + vbOKOnly, "Error"
            Else
                If MsgBox("Delete " & FileList.FileName & "?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "Are you sure?") = vbYes Then
                    If FileSystem.FileExists(FileList.Path & "\" & FileList.FileName) = True Then
                        FileSystem.DeleteFile FileList.Path & "\" & FileList.FileName
                        FileList.Refresh
                        FileList.ListIndex = 0
                    End If
                End If
            End If
    End Select
    ViewLog.Log logdebug, Me.Name & ":ToolBar_ButtonClick Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "LensDrawerSetup:ToolBar_ButtonClick", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub


Private Sub FileList_Click()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":FileList_Click"
    '---- Open the Lens Setup in the active spreadsheet
    CurrentFile = FileList.FileName
    LensSpreadsheet.Reset
    LensSpreadsheet.LoadTextFile FileList.Path & "\" & CurrentFile, ",", ",", vbCrLf, LoadTextFileColHeaders, ""
    ViewLog.Log logdebug, Me.Name & ":FileList_Click Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "LensDrawerSetup:FileList_Click", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

Private Sub LensSpreadsheet_LostFocus()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":LensSpreadsheet_LostFocus()"
    '---- Save the file
    LensSpreadsheet.ExportToTextFile FileList.Path & "\" & CurrentFile, ",", ",", vbCrLf, ExportToTextFileColHeaders, ""
    Exit Sub
    ViewLog.Log logdebug, Me.Name & ":CreateSpreadsheet_LostFocus Exiting"
ErrorHandler:
    ErrorForm.ReportError "LensDrawerSetup:LensSpreadsheet_LostFocus", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub


Public Sub Setup()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":Setup()"
    FileLocation = IIf(Right(Trim(App.Path), 1) = "\", Trim(App.Path), Trim(App.Path) & "\")
    FileLocation = FileLocation & "LensDrawers\"
    If FileSystem.FolderExists(FileLocation) = False Then
        FileSystem.CreateFolder FileLocation
    End If
    FileList.Path = FileLocation
    
    If FileSystem.FolderExists(FileLocation) = False Then
        FileSystem.CreateFolder FileLocation
    End If
    
    FileList.Refresh
    If FileList.ListCount > 0 Then
        FileList.ListIndex = 0
    Else
        CreateLensFile FileLocation & DefaultFile
        FileList.ListIndex = 0
    End If
    ViewLog.Log logdebug, Me.Name & ":CreateLensFile Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "LensDrawerSetup:Setup", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub


Public Sub CreateLensFile(FileName As String)
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":CreateLensFile(FileName=" & (FileName) & ")"
    Dim oFile As Scripting.TextStream
    Set oFile = FileSystem.CreateTextFile(FileName, True, False)
    oFile.WriteLine "Size,Paper Used,# Units"
    oFile.WriteLine "11x14,14,0"
    oFile.WriteLine "8x10,8,0"
    oFile.WriteLine "5x7,7,0"
    oFile.WriteLine "3.5x5,5,0"
    oFile.WriteLine "WALLT,10,0"
    oFile.WriteLine "SUBWLT,10,0"
    oFile.WriteLine "CHARMS,10,0"
    oFile.WriteLine "SPLIT,0,0"
    oFile.WriteLine "SPLIT,0,0"
    oFile.WriteLine "SPLIT,0,0"
    oFile.Close
    FileList.Refresh
    ViewLog.Log logdebug, Me.Name & ":CreateLensFile Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "LensDrawerSetup:CreateLensFile", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub


Private Sub PromptForFileName()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":PromptForFileName()"
    With CommonDialog
        .InitDir = FileList.Path                            'Set initial directory to lens path
        .CancelError = False
        .DefaultExt = ".Len"
        .Filter = "SDM Lens Drawer|*.len"
        .Flags = cdlOFNCreatePrompt Or cdlOFNOverwritePrompt
        .ShowSave
        If .FileName <> "" Then
            CreateLensFile .FileName
        End If
    ViewLog.Log logdebug, Me.Name & ":PromptForFileName Exiting"
    End With
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "LensDrawerSetup:PromptForFileName", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub


Public Function GetLensText(LensNumber As Integer) As String
    On Error GoTo ErrorHandler
   ViewLog.Log logdebug, Me.Name & ":GetLensText(LensNumber=" & (LensNumber) & ")"
    With LensSpreadsheet
        .Row = LensNumber
        .Col = 1
        GetLensText = Trim(Left(.Text, 10))
    End With
    ViewLog.Log logdebug, Me.Name & ":GetLensText Exiting = " + GetLensText
    Exit Function
ErrorHandler:
    ErrorForm.ReportError "LensDrawerSetup:GetLensText", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function
