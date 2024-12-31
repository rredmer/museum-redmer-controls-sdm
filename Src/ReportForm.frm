VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ReportForm 
   Caption         =   "Report"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox ReportText 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   240
      Width           =   7665
   End
   Begin MSComDlg.CommonDialog SaveDialog 
      Left            =   2520
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "txt"
      DialogTitle     =   "Save Report"
      FileName        =   "SDM_Report.txt"
      Filter          =   "txt"
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   3120
      Top             =   5880
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
            Picture         =   "ReportForm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportForm.frx":0322
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportForm.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportForm.frx":0CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportForm.frx":1330
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReportForm.frx":164A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolBar 
      Height          =   840
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Width           =   2385
      _ExtentX        =   4207
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
            Caption         =   "Save"
            Object.ToolTipText     =   "Save"
            Object.Tag             =   "Save report to disk"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Object.ToolTipText     =   "Print File"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ReportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'****************************************************************************
'**                                                                        **
'** Project....: Shooter Data Manager (SDM)                                **
'**                                                                        **
'** Module.....: ReportForm.frm - The application ReportFile handler       **
'**                                                                        **
'** Description: This form shows the ReportFile information                **
'**                                                                        **
'** History....:                                                           **
'**    03/20/02 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1997-2002 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit

Private FileInfoSpread As FPSpread.vaSpread
Private FileFrameSpread As FPSpread.vaSpread
Private FilePackageSpread As FPSpread.vaSpread
Dim MsgText As String
Dim MaxRows As Integer
Dim StartRow As Integer
Dim EndRow As Integer
Dim ColCount As Integer        'Problem: Setting value in sub
Dim RowCount As Integer          'Problem: Setting value in Sub
Dim RowNum As Integer
Dim TotalNoPackage As Integer
Dim TotalSlaters As Integer
Dim TotalBlinks As Integer
Dim TotalPrinted As Integer
Dim TotalNotPrinted As Integer
Dim TotalReprinted As Integer
Dim TotalPrintedTwice As Integer
Dim TotalDamaged As Integer
Dim TotalDirt As Integer
Dim TotalColor As Integer
Dim TotalQuantity As Integer
Dim TotalYaxis As Integer
Dim JobNumber As Integer
Dim ControlNumber As Integer
Dim PrintedSlates As Integer
Dim PrintedPacers As Integer
Dim FramesEdited As Integer


'****************************************************************************
'**                                                                        **
'**  Procedure....:  ReportFile                                            **
'**                                                                        **
'**  Description..:  This is where the ReportFile displays the information.**
'**                                                                        **
'**                                                                        **
'****************************************************************************

Public Sub ReportFile(SourceInfoSpread As FPSpread.vaSpread, SourceFrameSpread As FPSpread.vaSpread, SourcePackageSpread As FPSpread.vaSpread)

   
    On Error GoTo ErrorHandler
    ViewLog.Log logdebug, Me.Name & ":ReportFile(SourceSpreadSheet)"
    '---- Set pointer to source spreadsheet (where replace will take place)
    Set FileInfoSpread = SourceInfoSpread
    Set FileFrameSpread = SourceFrameSpread
    Set FilePackageSpread = SourcePackageSpread
    
    'FileInfoSpread.Row = FileInfoSpread.ActiveRow
    'FileInfoSpread.Col = FileInfoSpread.ActiveCol
    'FileFrameSpread.Row = FileFrameSpread.ActiveRow
    'FileFrameSpread.Col = FileFrameSpread.ActiveCol
    'FilePackageSpread.Row = FilePackageSpread.ActiveRow
    'FilePackageSpread.Col = FilePackageSpread.ActiveCol
    'ReportText.MaxLength = Len(SourceInfoSpread.Text)                'Max length of column
    'ReportText.MaxLength = Len(SourceFrameSpread.Text)
    'ReportText.MaxLength = Len(SourcePackageSpread.Text)
    'ReportText.Text = SourceInfoSpread.Text                          'Get the value of the column
    'ReportText.Text = SourceFrameSpread.Text
    'ReportText.Text = SourcePackageSpread.Text
    'FileInfoSpread.Col = 1
    'FileFrameSpread.Col = 1
    'FilePackageSpread.Col = 1
    
'*******************************************************************
'                     Report Information column                   **
'*******************************************************************
    
    'Note: Might need- .Text
    'Note: Might need- With SpreadName
    'Note: Might Need- ColCont and RowCount= No Value
       
    If FileInfoSpread <> "000000" Then
       For RowNum = ColCount To RowCount
            '.Col = 1
            '.Row = 1
        FileInfoSpread = RowNum
        
        Select Case SourceInfoSpread.Text
        Case 1                                         'for job number
 
            JobNumber = JobNumber + 1
        Case 2                                         'for control number
            ControlNumber = ControlNumber + 1
        
            End Select
        Next
    End If
    'SpreadSheet.Text = ReplaceText.Text
    'ReplaceText.Text = SourceSpreadSheet.Text
    
    
'*******************************************************************
'                           Frames Edited                         **
'*******************************************************************

    
    
    
 'FramesEdited= 1,2,3. Shows which frames have been edited.
    
    
    
    
'*******************************************************************
'                     Package Code Column                         **
'*******************************************************************

    If FileFrameSpread <> "00" Then
         For RowNum = StartRow To EndRow
         FileFrameSpread.Row = RowNum
        'FileFrameSpread.Text = ReportText.Text
        '.Col = 1
    Select Case SourceFrameSpread.Text
        Case "BL"                                         'for total blinks
            
            TotalBlinks = TotalBlinks + 1
        Case "SL"                                         'for total slaters
            TotalSlaters = TotalSlaters + 1
        Case "NP"                                         'for total no package
            TotalNoPackage = TotalNoPackage + 1
    
            End Select
        Next
    End If
  
'*******************************************************************
'                     Print Code Column                           **
'*******************************************************************
  
    If FileFrameSpread <> "00" Then
        For RowNum = StartRow To EndRow
            FileFrameSpread.Row = RowNum
       'FileFrameSpread.Text = ReportText.Text
       '.Col = 2
    Select Case SourceFrameSpread.Text
        Case "P"                                           'for printed
            TotalPrinted = TotalPrinted + 1
        Case "N"                                           'for not printed
            TotalNotPrinted = TotalNotPrinted + 1
        Case "R"                                           'for reprinted
            TotalReprinted = TotalReprinted + 1
        Case "T"                                           'for printed twice
            TotalPrintedTwice = TotalPrintedTwice + 1
        Case "X"                                           'for damaged
            TotalDamaged = TotalDamaged + 1
            
            End Select
        Next
    End If
  
'*******************************************************************
'                       Reprint reason column                     **
'*******************************************************************

    If FileFrameSpread <> "00" Then
        For RowNum = StartRow To EndRow
         FileFrameSpread.Row = RowNum
        'FileFrameSpread.Text = ReportText.Text
        '.Col = 3
    Select Case SourceFrameSpread.Text
        Case "D"                                          'for dirt
            TotalDirt = TotalDirt + 1
        Case "C"                                          'for color
            TotalColor = TotalColor + 1
        Case "Q"                                          'reprints for quantity
            TotalQuantity = TotalQuantity + 1
        Case "Y"                                          'for y-axis
            TotalYaxis = TotalYaxis + 1
            
            End Select
        Next
    End If
    
'*******************************************************************
'              Printed # for Slates and Pacers Column             **
'*******************************************************************

     If FileInfoSpread <> "00" Then
        For RowNum = StartRow To EndRow
         FileInfoSpread.Row = RowNum
        'FileFrameSpread.Text = ReportText.Text
    Select Case SourceInfoSpread.Text
        Case 1                                          'for printed slates
            PrintedSlates = PrintedSlates + 1
        Case 2                                          'for printed pacers
            PrintedPacers = PrintedPacers + 1
      
            End Select
        Next
    End If
    
'Debug.Print "********TESTING FOR RESULTS*********"
'ReportText.Text = "Job #: " & JobNumber & vbNewLine
'ReportText.Text = "Control #: " & ControlNumber & vbNewLine
'ReportText.Text = "Frames edited :" & FramesEdited & vbNewLine
'ReportText.Text = "Frame printed: " & TotalPrinted & vbNewLine
'ReportText.Text = "Frames not printed: " & TotalNotPrinted & vbNewLine
'ReportText.Text = "Frames to be reprinted: " & TotalReprinted & vbNewLine
'ReportText.Text = "Frames printed twice" & TotalPrintedTwice & vbNewLine
'ReportText.Text = "Reprints for quantity: " & TotalQuantity & vbNewLine
'ReportText.Text = "         for Y-axis: " & TotalYaxis & vbNewLine
'ReportText.Text = "         for color: " & TotalColor & vbNewLine
'ReportText.Text = "         for dirt: " & TotalDirt & vbNewLine
'ReportText.Text = "Frames damaged: " & TotalDamaged & vbNewLine
'ReportText.Text = "Total blinks: " & TotalBlinks & vbNewLine
'ReportText.Text = "Total slaters: " & TotalSlaters & vbNewLine
'ReportText.Text = "Total No Package: " & TotalNoPackage & vbNewLine
'ReportText.Text = "Slates printed: " & SlatesPrinted & vbNewLine
'ReportText.Text = "Pacers printed: " & PacersPrinted & vbNewLine
    
ViewLog.Log logdebug, Me.Name & ":ReportFile Exiting"
Exit Sub
ErrorHandler:
    ErrorForm.ReportError "ReportForm:DoReplace", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub
                    
Public Sub ReportShow()

MsgText = ""
MsgText = MsgText & "Job #: " & JobNumber & vbCrLf
MsgText = MsgText & "Control #: " & ControlNumber & vbCrLf
MsgText = MsgText & "Frames edited : " & FramesEdited & vbCrLf
MsgText = MsgText & "Frame printed: " & TotalPrinted & vbCrLf
MsgText = MsgText & "Frames not printed: " & TotalNotPrinted & vbCrLf
MsgText = MsgText & "Frames to be reprinted: " & TotalReprinted & vbCrLf
MsgText = MsgText & "Frames printed twice: " & TotalPrintedTwice & vbCrLf
MsgText = MsgText & "Reprints for quantity: " & TotalQuantity & vbCrLf
MsgText = MsgText & "         for Y-axis: " & TotalYaxis & vbCrLf
MsgText = MsgText & "         for color: " & TotalColor & vbCrLf
MsgText = MsgText & "         for dirt: " & TotalDirt & vbCrLf
MsgText = MsgText & "Frames damaged: " & TotalDamaged & vbCrLf
MsgText = MsgText & "Total blinks: " & Trim(str(TotalBlinks)) & vbCrLf
MsgText = MsgText & "Total slaters: " & TotalSlaters & vbCrLf
MsgText = MsgText & "Total No Package: " & TotalNoPackage & vbCrLf
MsgText = MsgText & "Slates printed: " & PrintedSlates & vbCrLf
MsgText = MsgText & "Pacers printed: " & PrintedPacers & vbCrLf
ReportText.Text = MsgText
ReportText.Refresh

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
            Clipboard.SetText ReportText.Text, vbCFText  'Save the text to the clipboard
        Case 3                                          'Save to disk
            SaveReportTextToFile                         'This routine uses the VB SaveDialog to prompt for saving file
        Case 4
          '  If MsgBox("Are you sure?", vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "Print") = vbYes Then
          
            
    End Select
    
    ViewLog.Log logdebug, Me.Name & ":ToolBar_ButtonClick Exiting"
    
    Exit Sub
ErrorHandler:
    MsgBox ""
End Sub


'****************************************************************************
'**                                                                        **
'**  Procedure....:  SaveReportTextToFile                                  **
'**                                                                        **
'**  Description..:  This routine stores error text to user-named file.    **
'**                                                                        **
'****************************************************************************

Private Sub SaveReportTextToFile()


    On Error GoTo ErrorHandler
    
    ViewLog.Log logdebug, Me.Name & ":SaveReportTextToFile()"
    
    '---- Show the save file dialog (VB6 has custom dialogs, .NET uses a common windows dialog)
    SaveDialog.ShowSave
    If SaveDialog.FileName <> "" Then
        Dim RepFileSys As New Scripting.FileSystemObject     'Pointer to Report File System Object
        Dim RepFile As Scripting.TextStream                     'Pointer to Error Report File
        
    '---- Get pointer to current log file and create it if necessary
        Set RepFile = RepFileSys.CreateTextFile(SaveDialog.FileName, True, False)
        RepFile.Write ReportText.Text
        RepFile.Close
        Set RepFile = Nothing
        Set RepFileSys = Nothing
        MsgBox "Created Report: " & SaveDialog.FileName, vbApplicationModal + vbOKOnly + vbInformation, "Saved"
        
    End If
    
    ViewLog.Log logdebug, Me.Name & ":SaveReportTextToFile Exiting"
    
    Exit Sub
ErrorHandler:
    MsgBox ""
End Sub

Private Sub Form_Load()
ReportShow
'ReportFile
End Sub

