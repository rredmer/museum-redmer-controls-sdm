VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm MainForm 
   BackColor       =   &H00808080&
   Caption         =   "SDM"
   ClientHeight    =   10155
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14085
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   30
      Top             =   9300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   9855
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu MainMenu 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu FileMenu 
         Caption         =   "&New"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu FileMenu 
         Caption         =   "&Open"
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu FileMenu 
         Caption         =   "&Close"
         Index           =   2
         Shortcut        =   ^X
      End
      Begin VB.Menu FileMenu 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu FileMenu 
         Caption         =   "&Save"
         Index           =   4
         Shortcut        =   ^S
      End
      Begin VB.Menu FileMenu 
         Caption         =   "Save &As"
         Index           =   5
      End
      Begin VB.Menu FileMenu 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu FileMenu 
         Caption         =   "&Print"
         Index           =   7
         Shortcut        =   ^P
      End
      Begin VB.Menu FileMenu 
         Caption         =   "Page Set&up"
         Index           =   8
         Shortcut        =   ^U
      End
      Begin VB.Menu FileMenu 
         Caption         =   "Production Reports"
         Index           =   9
      End
      Begin VB.Menu FileMenu 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu FileMenu 
         Caption         =   "File Manager"
         Index           =   11
      End
      Begin VB.Menu FileMenu 
         Caption         =   "DOS Disk Manager"
         Index           =   12
      End
      Begin VB.Menu FileMenu 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu FileMenu 
         Caption         =   "E&xit"
         Index           =   14
      End
   End
   Begin VB.Menu MainMenu 
      Caption         =   "&Edit"
      Index           =   1
      Begin VB.Menu EditMenu 
         Caption         =   "&Find"
         Index           =   2
         Shortcut        =   ^F
      End
      Begin VB.Menu EditMenu 
         Caption         =   "&Replace"
         Index           =   3
         Shortcut        =   ^H
      End
      Begin VB.Menu EditMenu 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu EditMenu 
         Caption         =   "&Go To Record"
         Index           =   5
         Shortcut        =   ^G
      End
      Begin VB.Menu EditMenu 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu EditMenu 
         Caption         =   "Printer Configurations"
         Index           =   7
      End
      Begin VB.Menu EditMenu 
         Caption         =   "Lens Drawers"
         Index           =   8
      End
      Begin VB.Menu EditMenu 
         Caption         =   "Pricing Tables"
         Index           =   9
      End
      Begin VB.Menu EditMenu 
         Caption         =   "Users"
         Index           =   10
      End
      Begin VB.Menu EditMenu 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu EditMenu 
         Caption         =   "Clear Packages"
         Index           =   12
      End
      Begin VB.Menu EditMenu 
         Caption         =   "Clear Frames"
         Index           =   13
      End
   End
   Begin VB.Menu MainMenu 
      Caption         =   "&Window"
      Index           =   2
      Begin VB.Menu WindowMenu 
         Caption         =   "Tile &Horiztonal"
         Index           =   0
      End
      Begin VB.Menu WindowMenu 
         Caption         =   "Tile &Vertical"
         Index           =   1
      End
      Begin VB.Menu WindowMenu 
         Caption         =   "&Cascade"
         Index           =   2
      End
      Begin VB.Menu WindowMenu 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu WindowMenu 
         Caption         =   "View Application Log File"
         Index           =   4
      End
      Begin VB.Menu WindowMenu 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu WindowMenu 
         Caption         =   "&Select"
         Index           =   6
         WindowList      =   -1  'True
      End
   End
   Begin VB.Menu MainMenu 
      Caption         =   "&Help"
      Index           =   3
      Begin VB.Menu HelpMenu 
         Caption         =   "About SDM"
         Index           =   0
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: Shooter Data Manager (SDM)                                **
'**                                                                        **
'** Module.....: MainForm.frm - The application main form (MDI Parent)     **
'**                                                                        **
'** Description: This form provides the application main window & menu.    **
'**                                                                        **
'** History....:                                                           **
'**    03/20/02 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2002 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit

'****************************************************************************
'**                                                                        **
'**  Procedure....:  MDIForm_Load                                          **
'**                                                                        **
'**  Description..:  Initialize form variables.                            **
'**                                                                        **
'****************************************************************************
Private Sub MDIForm_Load()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, "Loading " & Me.Name
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":MDIForm_Load", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub



'****************************************************************************
'**                                                                        **
'**  Procedure....:  MDIForm_QueryUnload                                   **
'**                                                                        **
'**  Description..:  Confirm application exit on close-window request.     **
'**                                                                        **
'****************************************************************************
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ErrorHandler
    
    ViewLog.Log logdebug, Me.Name & ":Form_QueryUnload(Cancel=" & str(Cancel) & ",UnloadMode=" & str(UnloadMode)
    
    If UnloadMode <> vbFormCode Then                                'Prevent recursive call from 'Exit' function in File Menu
        If EndTheProgram() = False Then                             'If the user chose not to exit the application
            Cancel = 1                                              'Cancel the request to exit
        End If
    End If
    
    ViewLog.Log logdebug, Me.Name & ":Form_QueryUnload Exiting"
    
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":MDIForm_QueryUnload", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

Public Sub UpdateMenus()

    On Error GoTo ErrorHandler
    
    ViewLog.Log logdebug, Me.Name & ":UpdateMenus()"
    
    If Forms.Count <= NumberOfAppForms Then                   'If no child documents are open
        FileMenu(2).Enabled = False         'Close
        FileMenu(4).Enabled = False         'Save
        FileMenu(5).Enabled = False         'Save As
        FileMenu(7).Enabled = False         'Save As
        MainMenu(1).Enabled = False         'Disable entire Edit Menu
        MainMenu(2).Enabled = False         'Disable entire Window Menu
    Else
        FileMenu(2).Enabled = True         'Close
        FileMenu(4).Enabled = True         'Save
        FileMenu(5).Enabled = True         'Save As
        FileMenu(7).Enabled = True         'Save As
        MainMenu(1).Enabled = True         'Disable entire Edit Menu
        MainMenu(2).Enabled = True         'Disable entire Window Menu
    End If
    
    ViewLog.Log logdebug, Me.Name & ":UpdateMenus Exiting"
    
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":UpdateMenus", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub



'****************************************************************************
'**                                                                        **
'**  Procedure....:  FileMenu_Click                                        **
'**                                                                        **
'**  Description..:  Handle menu selections on the File Menu.              **
'**                                                                        **
'****************************************************************************
Private Sub FileMenu_Click(Index As Integer)
    On Error GoTo ErrorHandler
    
    ViewLog.Log logdebug, Me.Name & ":FileMenu_Click(Index=" & str(Index) & ")"
    
    Select Case Index
        Case 0                                                      'New file
            CreateEditForm
        Case 1                                                      'Open
            OpenShooterFile
        Case 2                                                      'Close
            Unload Screen.ActiveForm
        Case 3                                                      'Line
        Case 4                                                      'Save
            If Left(Me.ActiveForm.Caption, 8) = "Untitled" Then     'If file has not been named yet
                SaveAsShooterFile                                   'Save with dialog
            Else                                                    'Else just save file
                SaveShooterFile
            End If
        Case 5                                                      'Save As
            SaveAsShooterFile
        Case 6                                                      'Line
        Case 7                                                      'Print
            PrintPreview.PrepareToPreview Me.ActiveForm.ActiveControl
        Case 8                                                      'Page Setup
            PageSetup
        Case 9                                                      'Production Reports
           ' ReportForm.Show vbModal
            MsgBox "Feature available in full version only."
        Case 10                                                     'Line
        Case 11                                                     'File Manager
           ' FileManager.Show vbModal
           MsgBox "Feature available in full version only."
        Case 12                                                     'DOS Disk Manager
            DosDiskManager.Show vbModal
        Case 13                                                     'Line
        Case 14                                                     'Exit Program
            EndTheProgram
    End Select
    
    ViewLog.Log logdebug, Me.Name & ":FileMenu_Click Exiting"
    
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":FileMenu_Click", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  EditMenu_Click                                        **
'**                                                                        **
'**  Description..:  Handle menu selections on the Edit Menu.              **
'**                                                                        **
'****************************************************************************
Private Sub EditMenu_Click(Index As Integer)
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":EditMenu_Click(Index=" & str(Index) & ")"
    Select Case Index
        Case 0              'Copy   (works by default with Farpoint Spread - no need to code here!)
        Case 1              'Paste  (works by default with Farpoint Spread - no need to code here!)
        Case 2              'Find
            FindForm.DoFindText Me.ActiveForm.ActiveControl
        Case 3              'Replace
            
                ReplaceForm.DoReplace Me.ActiveForm.ActiveControl
           
             
        Case 4              'Line
        Case 5              'Go to record
            If Screen.ActiveForm.ActiveControl.Name = "FrameSpread" Then
                '---- Goto record in Frame Spreadsheet
                With GoToRecordForm
                    .RecordNum.Mask = "###"
                    .RecordNum.MaxLength = 3
                    .GoToRecordLabel.Caption = "Frame#"
                    .Show vbModal
                    Screen.ActiveForm.ActiveControl.SetActiveCell 1, Val(.RecordNum.Text)
                End With
            Else
                '---- Goto record in Package Spreadsheet
                With GoToRecordForm
                    .RecordNum.Mask = "##"
                    .RecordNum.MaxLength = 2
                    .GoToRecordLabel.Caption = "Package#"
                    .Show vbModal
                    Screen.ActiveForm.ActiveControl.SetActiveCell 1, Val(.RecordNum.Text)
                End With
            End If
        Case 6              'Line
        Case 7              'Printer
            PhotoPrinterSetup.Show vbModal
            Me.ActiveForm.GetPackageLabels          'These may have changed in Setup - update display
            Me.ActiveForm.GetAlaCarteLabels
        Case 8              'Lens
            LensDrawerSetup.Show vbModal
            Me.ActiveForm.GetPackageLabels          'These may have changed in Setup - update display
            Me.ActiveForm.GetAlaCarteLabels
        Case 9              'Pricing
            MsgBox "Feature available in full version only."
        Case 10             'Users
            MsgBox "Feature available in full version only."
        Case 11             'Line
        Case 12             'Clear Packages
            If MsgBox("Are you sure?", vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "Clear Packages") = vbYes Then
                Me.ActiveForm.AddPackages
            End If
        Case 13             'Clear Frames
            If MsgBox("Are you sure?", vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "Clear Frames") = vbYes Then
                Me.ActiveForm.AddFrames
            End If
    End Select
    ViewLog.Log logdebug, Me.Name & ":EditMenu_Click Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":EditMenu_Click", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  WindowMenu_Click                                      **
'**                                                                        **
'**  Description..:  Handle menu selections on the Window Menu.            **
'**                                                                        **
'****************************************************************************
Private Sub WindowMenu_Click(Index As Integer)
    On Error GoTo ErrorHandler
   ViewLog.Log logdebug, Me.Name & ":WindowMenu_Click(Index=" & str(Index) & ")"
    Select Case Index
        Case 0
            Me.Arrange vbTileHorizontal
        Case 1
            Me.Arrange vbTileVertical
        Case 2
            Me.Arrange vbCascade
        Case 3              'Line
        Case 4              'Application Log File
            ViewLog.Show vbModal
    End Select
   ViewLog.Log logdebug, Me.Name & ":WindowMenu_Click Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":WindowMenu_Click", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  HelpMenu_Click                                        **
'**                                                                        **
'**  Description..:  Handle menu selections on the Help Menu.              **
'**                                                                        **
'****************************************************************************
Private Sub HelpMenu_Click(Index As Integer)
    On Error GoTo ErrorHandler
   ViewLog.Log logdebug, Me.Name & ":HelpMenu_Click(Index=" & str(Index) & ")"
    Select Case Index
        Case 0
            Splash.Show vbModal
    End Select
   ViewLog.Log logdebug, Me.Name & ":HelpMenu_Click Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":HelpMenu_Click", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub


'****************************************************************************
'**                                                                        **
'**  Procedure....:  SaveAsShooterFile                                     **
'**                                                                        **
'**  Description..:  Save shooter file with Dialog.                        **
'**                                                                        **
'****************************************************************************
Public Sub SaveAsShooterFile()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":SaveAsShooterFile()"
    With CommonDialog                                               'Windows Common control for Save As
        .DefaultExt = ".DAT"
        .Filter = "SDM Files (*.dat)|*.dat"
        .CancelError = False
        .DialogTitle = "Save File"
        .ShowSave                                                   'Show the Save As dialog
        If .FileName <> "" Then                                     'If the user chose a file
            Me.ActiveForm.Caption = .FileName                       'Set the window caption of the active window
            SaveShooterFile                                         'Save the file to disk
        End If
    End With
    ViewLog.Log logdebug, Me.Name & ":SaveAsShooterFile Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":SaveAsShooterFile", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  SaveShooterFile                                       **
'**                                                                        **
'**  Description..:  Save shooter file.                                    **
'**                                                                        **
'****************************************************************************
Public Sub SaveShooterFile()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":SaveShooterFile()"
    Dim RecordNum As Integer                                        'Counter for Packages & Frames
    Dim ColumnNum As Integer
    Dim SaveFileSystem As New Scripting.FileSystemObject     'Pointer to File System Object
    Dim SaveFile As Scripting.TextStream                     'Pointer to File to save
    Dim TextOutput As String
    Dim BytePos As Integer
    
    '---- Get pointer to current log file and create it if necessary
    Set SaveFile = SaveFileSystem.CreateTextFile(Screen.ActiveForm.Caption, True, False)
   
    '---- Save the Control Record
    With Me.ActiveForm.FileSpread
        
        .Col = 1: .Row = 1
        TextOutput = .Text
        .Col = 2
        TextOutput = TextOutput & .Text
        .Col = 3
        TextOutput = TextOutput & .Text
        .Col = 4
        TextOutput = TextOutput & .Text
        
        TextOutput = TextOutput & "0000"

        .Col = 5
        TextOutput = TextOutput & .Text
        .Col = 6
        TextOutput = TextOutput & .Text
        
        SaveFile.WriteLine TextOutput               'MUST be 24 bytes
    End With
    
    '---- Save the packages
    With Me.ActiveForm.PackageSpread
        For RecordNum = 1 To 99
            .Row = RecordNum
            TextOutput = ""
            For ColumnNum = 1 To 11
                .Col = ColumnNum
                TextOutput = TextOutput + .Text
            Next
            SaveFile.WriteLine TextOutput
        Next
    End With
    
    '---- Save the Frames
    With Me.ActiveForm.FrameSpread
        For RecordNum = 1 To 999
            .Row = RecordNum
            TextOutput = ""
            For ColumnNum = 1 To .MaxCols
                .Col = ColumnNum
                If ColumnNum >= 4 And ColumnNum <= 8 Then           'These are the signed columns +/-
                    If Val(.Text) >= 0 Then                         'If the value is zero or higher
                        TextOutput = TextOutput + "+" + .Text       'Put a Plus-sign on the front of it!
                    Else
                        TextOutput = TextOutput + .Text             'Else it already has a minus sign
                    End If
                Else                                                'All other columns are handled as-is
                    TextOutput = TextOutput + .Text
                End If
            Next
            SaveFile.WriteLine TextOutput
        Next
    End With
    SaveFile.Close
    Set SaveFile = Nothing
    Set SaveFileSystem = Nothing
    ViewLog.Log logdebug, Me.Name & ":SaveShooterFile Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":SaveShooterFile", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  OpenShooterFile                                       **
'**                                                                        **
'**  Description..:  Open existing shooter file & populate spreadsheet.    **
'**                                                                        **
'****************************************************************************
Public Sub OpenShooterFile()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":OpenShooterFile()"
    Dim RecordNum As Integer                                            'Counter for Packages & Frames
    Dim ColumnNum As Integer                                            'Counter for spreadsheet column
    With CommonDialog
        .DefaultExt = ".DAT"
        .Filter = "SDM Files (*.dat)|*.dat"
        .CancelError = False
        .DialogTitle = "Open File"
        .ShowOpen
        If .FileName <> "" Then                                         'If the user chose a file to open
            If CheckForFileOpen(.FileName) = False Then                 'If the file is not already open
                '---- Validate and open file
                Dim OpenFileSystem As New Scripting.FileSystemObject    'Pointer to File System Object
                Dim OpenFile As Scripting.TextStream                    'Pointer to File to read
                Dim TextInput As String                                 'Text read from file
                Dim BytePos As Integer                                  'Simple counter
                
                '---- Get pointer to current log file and create it if necessary
                Set OpenFile = OpenFileSystem.OpenTextFile(.FileName, ForReading, False, TristateUseDefault)
                TextInput = OpenFile.ReadLine
                If Len(TextInput) = 24 Then                              'Length of Control Record
                    '---- It is assumed the file is valid, create a new document instance and continue loading file
                    Dim EditFormPointer As New EditForm                     'Pointer to edit form
                    Dim EditFormKey As String                               'Key to edit form
                    EditFormKey = .FileName
                    EditFormPointer.Caption = EditFormKey                   'Assign the key to the form caption
                    EditFormPointer.Tag = EditFormKey
                    EditFormPointer.Show vbModeless                         'Show the form
                    
                    UpdateMenus

                    '---- Parse TextInput into the Control Columns
                    With Me.ActiveForm.FileSpread
                        .Col = 1: .Row = 1
                        .Text = Mid(TextInput, 1, 2)
                        .Col = 2
                        .Text = Mid(TextInput, 3, 2)
                        .Col = 3
                        .Text = Mid(TextInput, 5, 2)
                        .Col = 4
                        .Text = Mid(TextInput, 7, 2)
                        .Col = 5
                        .Text = Mid(TextInput, 13, 6)
                        .Col = 6
                        .Text = Mid(TextInput, 19, 6)
                    End With
                    
                    '---- Load and parse the packages
                    With Me.ActiveForm.PackageSpread
                        For RecordNum = 1 To 99
                            TextInput = OpenFile.ReadLine
                            If Len(TextInput) = 22 Then
                                .Row = RecordNum
                                BytePos = 1
                                For ColumnNum = 1 To 11
                                    .Col = ColumnNum
                                    .Text = Mid(TextInput, BytePos, 2)
                                    BytePos = BytePos + 2
                                Next
                            Else
                                MsgBox "The data file is damaged or invalid (Bad Package Record).", vbApplicationModal + vbInformation + vbOKOnly, "Error reading file"
                                Exit For
                            End If
                        Next
                    End With
                    
                    '---- Load and parse the frames
                    With Me.ActiveForm.FrameSpread
                        For RecordNum = 1 To 999
                            TextInput = OpenFile.ReadLine
                            If Len(TextInput) = 90 Then
                                .Row = RecordNum
                                .Col = 1                            'Package Code
                                .Text = Mid(TextInput, 1, 2)
                                .Col = 2                            'Print Code
                                .Text = Mid(TextInput, 3, 1)
                                .Col = 3                            'Reprint Code
                                .Text = Mid(TextInput, 4, 1)
                                .Col = 4                            'Density
                                .Text = Mid(TextInput, 5, 2)
                                .Col = 5                            'Y Axis
                                .Text = Mid(TextInput, 7, 2)
                                .Col = 6                            'CR
                                .Text = Mid(TextInput, 9, 2)
                                .Col = 7                            'CG
                                .Text = Mid(TextInput, 11, 2)
                                .Col = 8                            'CB
                                .Text = Mid(TextInput, 13, 2)
                                .Col = 9                            'VD
                                .Text = Mid(TextInput, 15, 3)
                                .Col = 10                           'VR
                                .Text = Mid(TextInput, 18, 2)
                                .Col = 11                           'VG
                                .Text = Mid(TextInput, 20, 2)
                                .Col = 12                           'VB
                                .Text = Mid(TextInput, 22, 2)
                                '---- The next 10 columns are alacarte sizes
                                BytePos = 24
                                For ColumnNum = 13 To 23
                                    .Col = ColumnNum
                                    .Text = Mid(TextInput, BytePos, 2)
                                    BytePos = BytePos + 2
                                Next
                                .Col = 23                           'Invoice #
                                .Text = Mid(TextInput, 44, 6)
                                .Col = 24                           'Twin Check
                                .Text = Mid(TextInput, 50, 6)
                                .Col = 25                           'Frame Number
                                .Text = Mid(TextInput, 56, 3)
                                .Col = 26                           'Film Type
                                .Text = Mid(TextInput, 59, 1)
                                .Col = 27                           'Rotation
                                .Text = Mid(TextInput, 60, 1)
                                .Col = 28
                                .Text = Mid(TextInput, 61, 30)
                            Else
                                If Len(TextInput) <> 0 Then         'If length is 0 then variable record file detected
                                    MsgBox "The data file is damaged or invalid (Bad Package Record).", vbApplicationModal + vbInformation + vbOKOnly, "Error reading file"
                                End If
                                Exit For
                            End If
                        Next
                    End With
                Else
                    '---- The Control Record is not the proper length, the file is not valid
                    MsgBox "The data file is damaged or invalid (Bad Control Record).", vbApplicationModal + vbInformation + vbOKOnly, "Error reading file"
                End If
                OpenFile.Close
                Set OpenFile = Nothing
                Set OpenFileSystem = Nothing
            End If
        Else
            MsgBox "No file selected."
        End If
    End With
   ViewLog.Log logdebug, Me.Name & ":OPenShooterFile Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":OpenShooterFile", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  PageSetup                                             **
'**                                                                        **
'**  Description..:  Perform printer Setup (File Menu)                     **
'**                                                                        **
'****************************************************************************
Private Sub PageSetup()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":PageSetup()"
    With CommonDialog                                       'uses Windows Common Dialog
        .CancelError = False
        .DialogTitle = "Page Setup"
        .Flags = cdlPDPrintSetup                            'Forces setup only
        .ShowPrinter
    End With
    ViewLog.Log logdebug, Me.Name & ":PageSetup Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError Me.Name & ":PageSetup", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub


