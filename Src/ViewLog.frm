VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form ViewLog 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Application Log"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   14400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox txtViewLog 
      Height          =   8265
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   14265
      _ExtentX        =   25162
      _ExtentY        =   14579
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      TextRTF         =   $"ViewLog.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgViewLog 
      Left            =   6090
      Top             =   8400
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
            Picture         =   "ViewLog.frx":0080
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewLog.frx":03A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewLog.frx":06C4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   885
      Left            =   2160
      TabIndex        =   1
      Top             =   8370
      Width           =   3885
      Begin VB.TextBox txtLocationID 
         Height          =   285
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "1"
         Top             =   510
         Width           =   915
      End
      Begin VB.CheckBox chkDebugEnable 
         Caption         =   "Enable Debug Mode"
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   210
         Width           =   2295
      End
      Begin VB.Label lblLabel 
         Caption         =   "Application Location Code"
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   540
         Width           =   1905
      End
   End
   Begin MSComctlLib.Toolbar tlbViewLog 
      Height          =   840
      Left            =   90
      TabIndex        =   0
      Top             =   8400
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   1482
      ButtonWidth     =   1746
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "imgViewLog"
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
            Caption         =   "&Erase Logs"
            Description     =   "Erase Application Log Files"
            Object.ToolTipText     =   "Erase application log files"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ViewLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: RC SeatTest                                               **
'**                                                                        **
'** Module.....: frmViewLog (Application Log/Timing Form)                   **
'**                                                                        **
'** Description: This form provides application tracking to flat file in   **
'**              addition to timing routines.                              **
'**                                                                        **
'**              The sub-systems include:                                  **
'**                Form Control - User Interface for Application View Log. **
'**                Log File Processing - File handling routines for log.   **
'**                Application Timing - 10ms Counter Routines.             **
'**                                                                        **
'**              Log File Format:                                          **
'**                System Date: 10 Bytes, mm/dd/yyyy                       **
'**                System Time: 8 bytes, hh:mm:ss                          **
'**                Elapsed Milliseconds: 12 Bytes, ssssssss.mmm            **
'**                Message Type: 5 bytes, enumerated constants (see below).**
'**                Message Text: Message text from calling function.       **
'**                                                                        **
'**              Message Types:                                            **
'**                INFOR: General information messages - always logged.    **
'**                DEBUG: Application debugging enabled, log for debug only**
'**                ERROR: Application error - always logged.               **
'**                                                                        **
'**              CONFIGURATION:                                            **
'**                DebugEnable - Set 1 to log DEBUG messages, 0 to ingore. **
'**                AutoCleanup - Set 1 to auto delete log files > 1 month. **
'**                MaxLogSize  - Determines maximum log file size.         **
'**                * Configurable values are stored in the Registry.       **
'**                                                                        **
'** History....:                                                           **
'**    03/20/02 v1.71 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2002 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit                                     'Require explicit variable declaration
Private Const CON_DIV = ":"                         'Log sectional divider
Private Const CON_LOGFILEDIR = "LogFiles"           'Log file directory
Private Const CON_MAXFILESIZE = 1400000             'Maximum log file size
Private oFileSystem As Scripting.FileSystemObject   'Pointer to Log File System Object
Private oFile As Scripting.TextStream               'Pointer to Log File
Private bLogFileConnected As Boolean                'Set TRUE when Log File is functional
Private sLogDir As String                           'Pointer to log file directory
Private sLogFileName As String                      'The name of the current log file
Private lMaxLogFileSize As Long                     'Maximum log file size
Private bAutoCleanup As Boolean                     'Indicates automatic log file cleanup
Enum LogMessageTypes                                'This enumeration helps programming
    logInformation = 0
    logdebug = 1
    LogError = 2
End Enum
Dim lMessageType As LogMessageTypes

'****************************************************************************
'**                                                                        **
'**                                                                        **
'**           FORM CONTROL SECTION - USER INTERFACE SUBSYSTEM              **
'**                                                                        **
'**                                                                        **
'****************************************************************************
'****************************************************************************
'**                                                                        **
'** Subroutine.: Form_Load                                                 **
'**                                                                        **
'** Description: This routine calls sub-system initialization routines.    **
'**                                                                        **
'****************************************************************************
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    txtLocationID.Text = GetSetting(App.CompanyName, App.ProductName, "LocationID", "1")
    If txtLocationID.Text = "1" Then SaveSetting App.CompanyName, App.ProductName, "LocationID", txtLocationID.Text     'Force write to create key (in case not found)
    
    chkDebugEnable.Value = GetSetting(App.CompanyName, App.ProductName, "DebugEnable", 0)
    If chkDebugEnable.Value = 0 Then SaveSetting App.CompanyName, App.ProductName, "DebugEnable", chkDebugEnable.Value  'Force write to create key (in case not found)

    lMaxLogFileSize = GetSetting(App.CompanyName, App.ProductName, "MaxLogSize", CON_MAXFILESIZE)
    If lMaxLogFileSize = CON_MAXFILESIZE Then SaveSetting App.CompanyName, App.ProductName, "MaxLogSize", lMaxLogFileSize  'Force write to create key (in case not found)
    
    bAutoCleanup = IIf(UCase(GetSetting(App.CompanyName, App.ProductName, "AutoCleanup", "TRUE")) = "TRUE", True, False)
    If bAutoCleanup = True Then SaveSetting App.CompanyName, App.ProductName, "AutoCleanup", "TRUE"  'Force write to create key (in case not found)
    
    Set oFileSystem = New Scripting.FileSystemObject
    LogFileConnect
    
    Exit Sub
ErrorHandler:
End Sub
'****************************************************************************
'**                                                                        **
'** Subroutine.: Form_Unload                                               **
'**                                                                        **
'** Description: This routine clears objects used in this formset.         **
'**                                                                        **
'****************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    LogFileDisconnect
    Set oFileSystem = Nothing
End Sub
'****************************************************************************
'**                                                                        **
'** Subroutine.: Form_Activate                                             **
'**                                                                        **
'** Description: This routine refreshed the display with current log file. **
'**                                                                        **
'****************************************************************************
Private Sub Form_Activate()
    On Error GoTo ErrorHandler
    txtViewLog.LoadFile sLogFileName
    Exit Sub
ErrorHandler:
    txtViewLog.Text = ""
End Sub
Private Sub tlbViewLog_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1                                      'Exit the log file form
            SaveSetting App.CompanyName, App.ProductName, "LocationID", Trim(txtLocationID.Text)
            SaveSetting App.CompanyName, App.ProductName, "DebugEnable", chkDebugEnable.Value
            Me.Hide
        Case 2                                      'Delete all log files
            If MsgBox("Delete all log files?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, "WARNING!") = vbYes Then
                Dim lFileCount As Long              'Number of files deleted
                LogFileDisconnect
                lFileCount = DeleteLogFiles(False)  'Delete Log Files w/ Auto flag off
                MsgBox "Deleted " & Trim(str(lFileCount)) & " log files.", vbApplicationModal + vbInformation + vbOKOnly, "Completed."
                LogFileConnect
            End If
    End Select
End Sub
Public Property Get MessageType() As LogMessageTypes
    MessageType = lMessageType
End Property
Public Property Let MessageType(lMsgType As LogMessageTypes)
    lMessageType = lMsgType
End Property
Public Property Get LocationID() As Integer
    LocationID = Val(txtLocationID.Text)
End Property
'****************************************************************************
'**                                                                        **
'**                                                                        **
'**                     LOG FILE PROCESSING SUBSYSTEM                      **
'**                                                                        **
'**                                                                        **
'****************************************************************************
'****************************************************************************
'**                                                                        **
'** Subroutine.: LogFileConnect                                            **
'**                                                                        **
'** Description: This routine creates a new log file instance.             **
'**                                                                        **
'****************************************************************************
Private Sub LogFileConnect()
    On Error GoTo ErrorHandler
    bLogFileConnected = False
    
    '---- Get pointer to filesystem object and verify log file folder
    sLogDir = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & CON_LOGFILEDIR
    If oFileSystem.FolderExists(sLogDir) = False Then
        oFileSystem.CreateFolder sLogDir
    End If
    
    '---- Get pointer to current log file and create it if necessary
    sLogFileName = sLogDir & "\LOG_" & Format(Now, "mmddyyyy") & "_" & Format(Now, "hhmmss") & ".TXT"
    Set oFile = oFileSystem.CreateTextFile(sLogFileName, True, False)
    oFile.Close
    Log logInformation, "Created log file (" & sLogFileName & ")"
    bLogFileConnected = True
    
    '---- Clean up log files if auto flag is active
    If bAutoCleanup Then
        DeleteLogFiles True                         'Delete Log Files w/ Auto flag ON
    End If
    Exit Sub
ErrorHandler:
End Sub
'****************************************************************************
'**                                                                        **
'** Subroutine.: LogFileDisconnect                                         **
'**                                                                        **
'** Description: This routine closes the currently open log file.          **
'**                                                                        **
'****************************************************************************
Private Sub LogFileDisconnect()
    On Error GoTo ErrorHandler
    Set oFile = Nothing
    bLogFileConnected = False
    Exit Sub
ErrorHandler:
End Sub
'****************************************************************************
'**                                                                        **
'** Subroutine.: Log                                                       **
'**                                                                        **
'** Description: This routine writes a new line of text to the log file.   **
'**                                                                        **
'****************************************************************************
Public Sub Log(lMsgType As LogMessageTypes, sLogMessage As String)
    
    '---- If debug message sent and debugging disabled, exit sub
    If chkDebugEnable.Value = 0 And lMsgType = logdebug Then
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    '---- Choose message type
    Dim sMessageType As String
    Select Case lMsgType
        Case 0
            sMessageType = "INFOR"
        Case 1
            sMessageType = "DEBUG"
        Case 2
            sMessageType = "ERROR"
    End Select
    
    '---- Write message into log file
    Dim sMsg As String
    Set oFile = oFileSystem.OpenTextFile(sLogFileName, ForAppending, False, TristateFalse)
    sMsg = Format(Now, "mm/dd/yyyy") & CON_DIV & Format(Now, "hh:mm:ssAMPM") & CON_DIV & Format(Timer / 1000, "00000000.000") & CON_DIV & sMessageType & CON_DIV & Trim(sLogMessage)
    
    oFile.WriteLine sMsg
    oFile.Close
    
    '---- Create new log file when current file reaches 1.4MB (this provides for copying files to Floppy disk)
    Dim oFC As Scripting.File
    Set oFC = oFileSystem.GetFile(sLogFileName)
    If oFC.Size > lMaxLogFileSize Then
        LogFileDisconnect
        LogFileConnect
    End If
    Set oFC = Nothing
    DoEvents
    Exit Sub
ErrorHandler:
End Sub
'****************************************************************************
'**                                                                        **
'** Subroutine.: DeleteLogFiles                                            **
'**                                                                        **
'** Description: This routine deletes log files, either forcibly through   **
'**              the user interface or programmatically using AutoCleanup. **
'**                                                                        **
'** Returns....: The number of files deleted.                              **
'**                                                                        **
'****************************************************************************
Private Function DeleteLogFiles(bCalledByAutoCleanup As Boolean) As Long
    On Error GoTo ErrorHandler
    Dim lFileCount As Long
    Dim oFolder As Scripting.Folder
    Dim oFl As Scripting.File
    Set oFolder = oFileSystem.GetFolder(sLogDir)
    lFileCount = 0
    For Each oFl In oFolder.Files
        If UCase(Left(oFl.Name, 3)) = "LOG" And UCase(Right(oFl.Name, 3)) = "TXT" Then
            If bCalledByAutoCleanup Then
                '---- Delete all log files not created in current month
                If Month(Now) <> Month(oFl.DateLastModified) Then
                    oFl.Delete True
                End If
            Else
                oFl.Delete True
                lFileCount = lFileCount + 1
            End If
        End If
    Next
    Set oFl = Nothing
    Set oFolder = Nothing
    If Not bCalledByAutoCleanup Then txtViewLog.Text = ""
    DeleteLogFiles = lFileCount
    Exit Function
ErrorHandler:
End Function


