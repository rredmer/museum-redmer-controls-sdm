VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form EditForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Untitled"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   9090
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   8175
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
      EndProperty
   End
   Begin VB.Frame ControlFrame 
      Appearance      =   0  'Flat
      Caption         =   "File Info"
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   30
      TabIndex        =   4
      Top             =   30
      Width           =   9045
      Begin FPSpread.vaSpread FileSpread 
         Height          =   495
         Left            =   60
         TabIndex        =   5
         Top             =   210
         Width           =   8925
         _Version        =   393216
         _ExtentX        =   15743
         _ExtentY        =   873
         _StockProps     =   64
         BorderStyle     =   0
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         EditModePermanent=   -1  'True
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   -2147483633
         ScrollBars      =   0
         SelectBlockOptions=   0
         SpreadDesigner  =   "EditForm.frx":0000
      End
   End
   Begin VB.Frame PackageFrame 
      Caption         =   "Packages"
      Height          =   2445
      Left            =   30
      TabIndex        =   2
      Top             =   810
      Width           =   9045
      Begin FPSpread.vaSpread PackageSpread 
         Height          =   2175
         Left            =   60
         TabIndex        =   3
         Top             =   210
         Width           =   8925
         _Version        =   393216
         _ExtentX        =   15743
         _ExtentY        =   3836
         _StockProps     =   64
         BorderStyle     =   0
         ColHeaderDisplay=   0
         EditModePermanent=   -1  'True
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SelectBlockOptions=   0
         SpreadDesigner  =   "EditForm.frx":0218
      End
   End
   Begin VB.Frame ExposureFrame 
      Caption         =   "Exposures"
      Height          =   4845
      Left            =   30
      TabIndex        =   0
      Top             =   3270
      Width           =   9045
      Begin FPSpread.vaSpread FrameSpread 
         Height          =   4485
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   8925
         _Version        =   393216
         _ExtentX        =   15743
         _ExtentY        =   7911
         _StockProps     =   64
         BorderStyle     =   0
         ColHeaderDisplay=   0
         EditModePermanent=   -1  'True
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SelectBlockOptions=   0
         SpreadDesigner  =   "EditForm.frx":0417
         UserResize      =   0
      End
   End
End
Attribute VB_Name = "EditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'**                                                                        **
'** Project....: Shooter Data Manager (SDM)                                **
'**                                                                        **
'** Module.....: EditForm.frm - The SDM File Editing form (MDI Child)      **
'**                                                                        **
'** Description: This form provides the application document window.       **
'**                                                                        **
'** History....:                                                           **
'**    03/20/02 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2002 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit

Private Sub Form_Activate()
    FrameSpread.SetFocus
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Form_Load                                             **
'**                                                                        **
'**  Description..:  Initialize form variables.                            **
'**                                                                        **
'****************************************************************************
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":Form_Load()"
    AddFileInfo                                             'Setup the File Info (Control Record) Spreadsheet
    AddPackages                                             'Setup the Packages Spreadsheet
    AddFrames                                               'Setup the Frames Spreadsheet
    StatusBar.Panels(1).Width = Me.Width - 5                'Set the width of the statusbar panel to take up full form
    ViewLog.Log logdebug, Me.Name & ":Form_Load Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "EditForm:Form_Load", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Form_QueryUnload                                      **
'**                                                                        **
'**  Description..:  Confirm closing of the form from the exit button.     **
'**                                                                        **
'****************************************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ErrorHandler
        If MsgBox("Are you sure?", vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "Close") = vbYes Then
        Unload Screen.ActiveForm
            'UnloadEditForm
        Else
            Cancel = 1
        End If
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "EditForm:Form_QueryUnload", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

Private Sub Form_Terminate()
    MainForm.UpdateMenus
End Sub


'****************************************************************************
'**                                                                        **
'**  Procedure....:  FrameSpread_LeaveCell                                 **
'**                                                                        **
'**  Description..:  Validate cells in the Frame Spreadsheet.              **
'**                                                                        **
'****************************************************************************
Private Sub FrameSpread_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":FrameSpread_LeaveCell(Col=" + str(Col) + ",Row=" + str(Row) + ",NewCol=" + str(NewCol) + ",NewRow=" + str(NewRow) + "," + IIf(Cancel, "true", "false")
    '---- Validate Cells
    With FrameSpread
        .Col = Col
        .Row = Row
        Select Case Col
            Case 1
                '---- Validate package code
                If .Text <> "BL" And .Text <> "SL" And .Text <> "NP" Then
                    If IsNumeric(.Text) = False Then
                        StatusBar.Panels(1).Text = "Package code must be 00-99, BL= Blink, SL= Slate, or NP= No Package"
                        Cancel = True
                    End If
                End If
            Case 2
                Select Case .Text
                    Case "N", "P", "R", "T", "X"
                    Case Else
                        StatusBar.Panels(1).Text = "Print Code must be N= Not Printed, P=Printed, R= Re-Printed, T= Re-Printed Twice, or X=Damaged."
                        Cancel = True
                End Select
            Case 3
                Select Case .Text
                    Case "D", "C", "Q", "Y"
                    Case Else
                        StatusBar.Panels(1).Text = "Reprint Code must be D= Dirt, C= Color, Q= Quantity, or Y= Y-Axis."
                        Cancel = True
                End Select
            Case 4 To 8
                If Val(.Text) < -9 Or Val(.Text) > 9 Then
                    StatusBar.Panels(1).Text = "Value must be -9 to 9."
                    Cancel = True
                End If
        End Select
    End With
    ViewLog.Log logdebug, Me.Name & ":FrameSpread_LeaveCell Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "EditForm:FrameSpread_LeaveCell", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  AddFileInfo                                           **
'**                                                                        **
'**  Description..:  Configure the File Info (Control Record) Spreadsheet. **
'**                                                                        **
'****************************************************************************
Private Sub AddFileInfo()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":AddFileInfo()"
    With FileSpread
        .MaxCols = 6
        .MaxRows = 1
        .BlockMode = True
        .Row = 1
        .Col = 1
        .Col2 = 6
        .Row2 = 1
        .CellType = CellTypePic
        .TypePicMask = "99"
        .Text = "00"
        .BlockMode = False
        .Col = 5
        .TypePicMask = "999999"
        .Text = "000000"
        .Col = 6
        .TypePicMask = "999999"
        .Text = "000000"
        .SetText 1, 0, "Pacer"
        .SetText 2, 0, "Pacer Size"
        .SetText 3, 0, "Slate"
        .SetText 4, 0, "Mode"
        .SetText 5, 0, "Job#"
        .SetText 6, 0, "Control#"
        .RetainSelBlock = False
        .ProcessTab = True
        .EditModeReplace = True
        .ArrowsExitEditMode = True
        .OperationMode = OperationModeNormal
        .AllowDragDrop = True
    End With
    ViewLog.Log logdebug, Me.Name & ":AddFileInfo Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "EditForm:AddFileInfo", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  AddPackages                                           **
'**                                                                        **
'**  Description..:  Configure the Packages Spreadsheet.                   **
'**                                                                        **
'****************************************************************************
Public Sub AddPackages()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":AddPackages()"
    With PackageSpread
        .MaxCols = 11
        .MaxRows = 99
        .BlockMode = True
        .Row = 1
        .Col = 1
        .Col2 = 11
        .Row2 = .MaxRows
        .CellType = CellTypePic
        .TypePicMask = "99"
        .Text = "00"
        .ColWidth(-1) = 6
        .BlockMode = False
        .RetainSelBlock = False
        .ProcessTab = True
        .EditModeReplace = True
        .ArrowsExitEditMode = True
        .OperationMode = OperationModeNormal
        .AllowDragDrop = True
        
        GetPackageLabels
        
    End With
    ViewLog.Log logdebug, Me.Name & ":AddPackages Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "EditForm:AddPackages", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  AddFrames                                             **
'**                                                                        **
'**  Description..:  Configure the Frames Spreadsheet.                     **
'**                                                                        **
'****************************************************************************
Public Sub AddFrames()
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":AddFrames()"
    With FrameSpread
        .MaxCols = 28
        .MaxRows = 999
        .SetText 1, 0, "Pkg"
        .SetText 2, 0, "Prnt"
        .SetText 3, 0, "Rep"
        .SetText 4, 0, "Den"
        .SetText 5, 0, "Y"
        .SetText 6, 0, "CR"
        .SetText 7, 0, "CG"
        .SetText 8, 0, "CB"
        .SetText 9, 0, "VD"
        .SetText 10, 0, "VR"
        .SetText 11, 0, "VG"
        .SetText 12, 0, "VB"
        .SetText 13, 0, "SZ1"
        .SetText 14, 0, "SZ2"
        .SetText 15, 0, "SZ3"
        .SetText 16, 0, "SZ4"
        .SetText 17, 0, "SZ5"
        .SetText 18, 0, "SZ6"
        .SetText 19, 0, "SZ7"
        .SetText 20, 0, "SZ8"
        .SetText 21, 0, "SZ9"
        .SetText 22, 0, "SZ10"
        .SetText 23, 0, "Invoice"
        .SetText 24, 0, "TwinCheck"
        .SetText 25, 0, "Frm#"
        .SetText 26, 0, "FilmType"
        .SetText 27, 0, "Rotate"
        .SetText 28, 0, "Text"
        
        .BlockMode = True
        .Row = 1: .Row2 = .MaxRows
        
        '---- Package Code Column
        .Col = 1: .Col2 = 1
        .CellType = CellTypePic
        .TypePicMask = "NN"
        .Text = "00"
        .ColWidth(-2) = 4
        
        '---- Print Code Column
        .Col = 2: .Col2 = 2
        .CellType = CellTypePic
        .TypePicMask = "N"
        .Text = "N"
        .ColWidth(-2) = 4

        '---- Reprint reason column
        .Col = 3: .Col2 = 3
        .CellType = CellTypePic
        .TypePicMask = "N"
        .Text = "D"
        .ColWidth(-2) = 4

        '---- Den, Y-Axis, CR, CG, CB Columns
        .Col = 4: .Col2 = 8
        .CellType = CellTypeNumber
        .TypeNumberMin = -9
        .TypeNumberMax = 9
        .TypeSpin = True
        .TypeNumberDecPlaces = 0
        .Text = "0"
        .ColWidth(-2) = 4
        .Col = 5
        .ColWidth(-2) = 3
        .Col = 6
        .ColWidth(-2) = 3
        .Col = 7
        .ColWidth(-2) = 3
        .Col = 8
        .ColWidth(-2) = 3
        
        '---- Density Column
        .Col = 9: .Col2 = 9
        .CellType = CellTypePic
        .TypePicMask = "999"
        .Text = "000"
        .ColWidth(-2) = 4
        
        '---- Video Red, Video Green, Video Blue, & Ala Carte Column
        .Col = 10: .Col2 = 22
        .CellType = CellTypePic
        .TypePicMask = "99"
        .Text = "00"
        Dim ColNum
        For ColNum = 10 To 22
            .Col = ColNum
            .ColWidth(-2) = 5
        Next
        
        '---- Invoice & Twin Check Columns
        .Col = 23: .Col2 = 24
        .CellType = CellTypePic
        .TypePicMask = "999999"
        .Text = "000000"
        .ColWidth(-2) = 6
        
        '---- Frame # Column
        .Col = 25: .Col2 = 25
        .CellType = CellTypePic
        .TypePicMask = "999"
        .Text = "000"
        .ColWidth(-2) = 4
        
        '---- Film Type Column
        .Col = 26: .Col2 = 26
        .CellType = CellTypePic
        .TypePicMask = "9"
        .Text = "0"
        .ColWidth(-2) = 4
        
        '---- Rotation Column
        .Col = 27: .Col2 = 27
        .CellType = CellTypePic
        .TypePicMask = "N"
        .Text = "Y"
        .ColWidth(-2) = 5
        
        '---- Text Column
        .Col = 28: .Col2 = 28
        .Text = Space(30)
        .ColWidth(-2) = 50
                
        .BlockMode = False
        .RetainSelBlock = False
        .ProcessTab = True
        .EditModeReplace = True
        .ArrowsExitEditMode = True
        .OperationMode = OperationModeNormal
        .AllowDragDrop = True
    End With
    ViewLog.Log logdebug, Me.Name & ":AddFrames Exiting"
    GetAlaCarteLabels

    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "EditForm:AddPackages", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub


Public Sub GetPackageLabels()
    '---- Configure Captions
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":GetPackageLabels()"
    Dim ColCount As Integer
    Dim NumberOfLenses As Integer                       'Number of Lenses in Printer Definition
    With PackageSpread
        .Row = 0
        NumberOfLenses = PhotoPrinterSetup.GetNumberOfLenses
        For ColCount = 1 To NumberOfLenses
            .Col = ColCount
            .Text = LensDrawerSetup.GetLensText(ColCount)
        Next
        For ColCount = NumberOfLenses + 1 To .MaxCols
            .Col = ColCount
            .Text = "SP-" & Trim(str(ColCount - NumberOfLenses))
        Next
    End With
    ViewLog.Log logdebug, Me.Name & ":GetPackageLabels Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "EditForm:GetPackageLabels", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

Public Sub GetAlaCarteLabels()
    '---- Configure Captions
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":GetAlaCarteLabels()"
    Dim ColCount As Integer
    Dim NumberOfLenses As Integer                       'Number of Lenses in Printer Definition
    With FrameSpread
        .Row = 0
        NumberOfLenses = PhotoPrinterSetup.GetNumberOfLenses
        For ColCount = 13 To (13 + NumberOfLenses - 1)
            .Col = ColCount
            .Text = LensDrawerSetup.GetLensText(ColCount - 12)
        Next
        For ColCount = 13 + NumberOfLenses To 22
            .Col = ColCount
            .Text = "SP-" & Trim(str(ColCount - NumberOfLenses - 12))
        Next
    End With
    ViewLog.Log logdebug, Me.Name & ":GetAlaCarteLabels Exiting"
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "EditForm:GetFrameLabels", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

