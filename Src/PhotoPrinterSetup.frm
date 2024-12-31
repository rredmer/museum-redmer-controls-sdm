VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form PhotoPrinterSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Photo Printer Setup"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox FileList 
      Height          =   6720
      Left            =   60
      Pattern         =   "*.pps"
      TabIndex        =   0
      Top             =   330
      Width           =   3255
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   2580
      Top             =   7470
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpread.vaSpread PrinterSpreadsheet 
      Height          =   1725
      Left            =   3360
      TabIndex        =   1
      Top             =   330
      Width           =   8325
      _Version        =   393216
      _ExtentX        =   14684
      _ExtentY        =   3043
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
      SpreadDesigner  =   "PhotoPrinterSetup.frx":0000
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   1980
      Top             =   7380
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
            Picture         =   "PhotoPrinterSetup.frx":01D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhotoPrinterSetup.frx":04F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PhotoPrinterSetup.frx":0818
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolBar 
      Height          =   840
      Left            =   90
      TabIndex        =   2
      Top             =   7170
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
   Begin FPSpread.vaSpread SplitSpreadsheet 
      Height          =   4605
      Left            =   3360
      TabIndex        =   4
      Top             =   2430
      Width           =   8325
      _Version        =   393216
      _ExtentX        =   14684
      _ExtentY        =   8123
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
      SpreadDesigner  =   "PhotoPrinterSetup.frx":0B3A
   End
   Begin VB.Label SplitLabel 
      Caption         =   "Split Definitions"
      Height          =   225
      Left            =   3390
      TabIndex        =   5
      Top             =   2190
      Width           =   1125
   End
   Begin VB.Label LensFileLabel 
      Caption         =   "Photo Printer Setup Files:"
      Height          =   285
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   2355
   End
End
Attribute VB_Name = "PhotoPrinterSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: Shooter Data Manager (SDM)                                **
'**                                                                        **
'** Module.....: PhotoPrinterSetup.frm - Setup Form for Photo Printer Files**
'**                                                                        **
'** Description: This form provides configuration using FarPoint Spread.   **
'**                                                                        **
'** History....:                                                           **
'**    03/20/02 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2002 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit

Private Const DefaultFile As String = "DefaultPrinter.pps"  'Default name of Photo Printer Files
Private Const SplitExt As String = ".Splits"
Private FileLocation As String                              'Location of Photo Printer Files
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
                MsgBox "Can't erase default printer file.", vbApplicationModal + vbInformation + vbOKOnly, "Error"
            Else
                If MsgBox("Delete " & FileList.FileName & "?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "Are you sure?") = vbYes Then
                    If FileSystem.FileExists(FileList.Path & "\" & FileList.FileName) = True Then
                        FileSystem.DeleteFile FileList.Path & "\" & FileList.FileName
                        FileSystem.DeleteFile FileList.Path & "\" & FileList.FileName & SplitExt
                        FileList.Refresh
                        FileList.ListIndex = 0
                    End If
                End If
            End If
    End Select
    
   ViewLog.Log logdebug, Me.Name & ":ToolBar_ButtonClick Exiting"
    
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "PhotoPrinterSetup:ToolBar_ButtonClick", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub


Private Sub FileList_Click()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":FileList_Click()"
    '---- Open the Printer Setup in the active spreadsheet
    CurrentFile = FileList.FileName
    PrinterSpreadsheet.Reset
    PrinterSpreadsheet.LoadTextFile FileList.Path & "\" & CurrentFile, ",", ",", vbCrLf, LoadTextFileColHeaders, ""
    SplitSpreadsheet.Reset
    SplitSpreadsheet.LoadTextFile FileList.Path & "\" & CurrentFile & SplitExt, ",", ",", vbCrLf, LoadTextFileColHeaders, ""
    
    '---- Configure spreadsheet editing
    With PrinterSpreadsheet
        .BlockMode = True
        .Col = 1: .Col2 = 1
        .Row = 1: .Row2 = .MaxRows
        .Lock = True
        .LockBackColor = vbGrayText
        .ColWidth(-2) = 30
        .BlockMode = False
    End With
    ViewLog.Log logdebug, Me.Name & ":FileList_Click Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "PhotoPrinterSetup:FileList_Click", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

Private Sub PrinterSpreadsheet_LostFocus()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":PrinterSpreadsheet_LostFocus()"
    '---- Save the file
    PrinterSpreadsheet.ExportToTextFile FileList.Path & "\" & CurrentFile, ",", ",", vbCrLf, ExportToTextFileColHeaders, ""
    SplitSpreadsheet.ExportToTextFile FileList.Path & "\" & CurrentFile & SplitExt, ",", ",", vbCrLf, ExportToTextFileColHeaders, ""
    ViewLog.Log logdebug, Me.Name & ":PrinterSpreadsheet_LostFocus Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "PhotoPrinterSetup:PrinterSpreadsheet_LostFocus", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub


Public Sub Setup()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":Setup()"
    FileLocation = IIf(Right(Trim(App.Path), 1) = "\", Trim(App.Path), Trim(App.Path) & "\")
    FileLocation = FileLocation & "PhotoPrinters\"
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
        CreatePrinterFile FileLocation & DefaultFile
        FileList.ListIndex = 0
    End If
    ViewLog.Log logdebug, Me.Name & ":Setup Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "PhotoPrinterSetup:Setup", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub


Public Sub CreatePrinterFile(FileName As String)
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":CreatePrinterFile(FileName=" & (FileName) & ")"
    Dim oFile As Scripting.TextStream
    Dim oFileA As Scripting.TextStream
    Dim SplitCount As Integer
    
    Set oFile = FileSystem.CreateTextFile(FileName, True, False)
    oFile.WriteLine "Setting Name,Value"
    oFile.WriteLine "Printer Brand,Lucht VP-2"
    oFile.WriteLine "Lens Drawer,DefaultLens.len"
    oFile.WriteLine "Number of Lenses in Drawer,6"
    oFile.WriteLine "Paper Advance Speed (Inches Per Second),0.00"
    oFile.WriteLine "Film Drive Time     (Seconds Per Frame),0.00"
    oFile.WriteLine "Data Comm Time      (Seconds Per Frame),0.00"
    oFile.WriteLine "Film Drive Rotation (Seconds Per Frame),0.00"
    
    Set oFileA = FileSystem.CreateTextFile(FileName & SplitExt, True, False)
    oFileA.WriteLine "#Exposures,Paper Length,#Units,Deck Rotation"
    For SplitCount = 1 To 99
        oFileA.WriteLine "0,0.00,0.00,Y"
    Next
    oFile.Close
    oFileA.Close
    FileList.Refresh
    ViewLog.Log logdebug, Me.Name & ":CreatePrinterFile Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "PhotoPrinterSetup:CreatePrinterFile", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub


Private Sub PromptForFileName()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":PromptForFileName()"
    With CommonDialog
        .InitDir = FileList.Path                            'Set initial directory to Printer path
        .CancelError = False
        .DefaultExt = ".pps"
        .Filter = "SDM Photo Printer|*.pps"
        .Flags = cdlOFNCreatePrompt Or cdlOFNOverwritePrompt
        .ShowSave
        If .FileName <> "" Then
            CreatePrinterFile .FileName
        End If
    End With
    ViewLog.Log logdebug, Me.Name & ":PromptForFileName Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "PhotoPrinterSetup:PromptForFileName", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

Public Function GetNumberOfLenses() As Integer
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":GetNumberOfLenses()"
    With PrinterSpreadsheet
        .Col = 2
        .Row = 3
        GetNumberOfLenses = IIf(Val(.Text) > 10, 6, Val(.Text))
    End With
    ViewLog.Log logdebug, Me.Name & ":GetNumberOfLenses()"
    Exit Function
ErrorHandler:
    ErrorForm.ReportError "PhotoPrinterSetup:GetNumberOfLenses", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function
