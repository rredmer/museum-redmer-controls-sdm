VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form DosDiskManager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DOS Disk Manager"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   4020
      Top             =   4890
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.img"
   End
   Begin VB.FileListBox FileList 
      Height          =   4185
      Left            =   60
      Pattern         =   "*.img"
      TabIndex        =   0
      Top             =   330
      Width           =   5400
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   3420
      Top             =   4800
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
            Picture         =   "DosDiskManager.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DosDiskManager.frx":0322
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DosDiskManager.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DosDiskManager.frx":1016
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolBar 
      Height          =   840
      Left            =   60
      TabIndex        =   1
      Top             =   4560
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   1482
      ButtonWidth     =   1455
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
            Caption         =   "Write Disk"
            Object.ToolTipText     =   "Write Image to Disk"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Read Disk"
            Object.ToolTipText     =   "Read Image from Disk"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Erase"
            Object.ToolTipText     =   "Erase Disk Image"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Label DiskImagesLocation 
      Caption         =   "Disk Image Files:"
      Height          =   225
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   3435
   End
End
Attribute VB_Name = "DosDiskManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const DefaultFile As String = "DefaultDisk.img"     'Default name of image file
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
        Case 2                                  'Read
            If FileList.FileName <> "" Then
                Call WriteImageToFloppy(FileList.Path & "\" & FileList.FileName, True)
            End If
        Case 3                                  'Write
            ReadDiskToFile
        Case 4                                  'Delete Image
            If FileList.FileName = DefaultFile Then
                MsgBox "Can't erase default printer file.", vbApplicationModal + vbInformation + vbOKOnly, "Error"
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
    ErrorForm.ReportError "DosDiskManager:ToolBar_ButtonClick", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

Public Sub Setup()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":Setup()"
    FileLocation = IIf(Right(Trim(App.Path), 1) = "\", Trim(App.Path), Trim(App.Path) & "\")
    FileLocation = FileLocation & "DiskImages\"
    If FileSystem.FolderExists(FileLocation) = False Then
        FileSystem.CreateFolder FileLocation
    End If
    FileList.Path = FileLocation
   
    If FileSystem.FolderExists(FileLocation) = False Then
        FileSystem.CreateFolder FileLocation
    End If
    FileList.Refresh
     ViewLog.Log logdebug, Me.Name & ":Setup Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "DosDiskManager:Setup", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

Private Sub ReadDiskToFile()
    On Error GoTo ErrorHandler
     ViewLog.Log logdebug, Me.Name & ":ReadDiskToFile()"
    With CommonDialog
        .InitDir = FileList.Path                            'Set initial directory to Printer path
        .CancelError = False
        .DefaultExt = ".img"
        .Filter = "Shooter Disk Images|*.img"
        .Flags = cdlOFNCreatePrompt Or cdlOFNOverwritePrompt
        .ShowSave
        If .FileName <> "" Then
            DoEvents
            Call ReadFloppyToFile(.FileName, "SDM", True)
            FileList.Refresh
        End If
    End With
    ViewLog.Log logdebug, Me.Name & ":ReadDiskToFile Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "DosDiskManager:ReadDiskToFile", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub
