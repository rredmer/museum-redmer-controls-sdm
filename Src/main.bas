Attribute VB_Name = "Startup"
'****************************************************************************
'**                                                                        **
'** Project....: Shooter Data Manager (SDM)                                **
'**                                                                        **
'** Module.....: Main.bas - The application main module (startup)          **
'**                                                                        **
'** Description: This is the main application module, it loads all of the  **
'**              global forms and provides functions to handle the multi-  **
'**              document interface (Create & Unload child forms).         **
'**                                                                        **
'** History....:                                                           **
'**    03/20/02 v1.00 RDR Designed and programmed first release.           **
'**                                                                        **
'** (c) 1992-2002 Redmer Controls Inc.  All rights reserved.               **
'****************************************************************************
Option Explicit
Public EditFormCount As Integer
Public NumberOfAppForms As Integer

'****************************************************************************
'**                                                                        **
'**  Procedure....:  Main                                                  **
'**                                                                        **
'**  Description..:  This is the main application procedure (called from   **
'**                  Windows).  It loads the application's forms.          **
'**                                                                        **
'****************************************************************************
Public Sub Main()
    On Error GoTo ErrorHandler                              'Standard Error Handler
        
    '---- Check for a previous instance of the application (prevent multiple copies from running)
    If App.PrevInstance = True Then                         'If the application is already running
        MsgBox "SDM is already running.", vbSystemModal + vbCritical + vbOKOnly, "Error"
    End If
    
    '---- Load the Application forms
    Splash.Show vbModeless                                  'Show the splash window
    Load ViewLog                                            'Load the application log form (must be the first app form)
    Load TimingForm                                         'Load the application timing form
    TimingForm.Delay 1000                                   'One second delay for splash screen
    
    Load DosDiskManager
    DosDiskManager.Setup
    
    'adding ReportForm
    Load ReportForm
    Load PrintPreview
    Load GoToRecordForm
    Load FindForm
    Load ReplaceForm
    Load ErrorForm                                          'Load the application error handler form
    
    Load LensDrawerSetup
    LensDrawerSetup.Setup
    
    Load MainForm
    
    Load PhotoPrinterSetup
    PhotoPrinterSetup.Setup
    
    
    NumberOfAppForms = Forms.Count
    
    MainForm.Show vbModeless                                'Show the main form (MDI Parent Form)
    EditFormCount = 0                                       'Initialize the Edit Form Counter (# new documents)
    CreateEditForm                                          'Create the first edit form ("Untitled 1" window)
    Splash.Hide                                             'Hide the splash window
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "Main", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
    End
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  EndTheProgram                                         **
'**                                                                        **
'**  Description..:  This routine terminates the program cleanly.          **
'**                                                                        **
'****************************************************************************
Public Function EndTheProgram() As Boolean
    On Error Resume Next                                    'There is no stopping this routine!
    
    If MsgBox("Are you sure?", vbApplicationModal + vbQuestion + vbYesNo, "Exit SDM") = vbYes Then
        Unload Splash
        Unload TimingForm
        Unload DosDiskManager
        'added unload ReportForm
        Unload ReportForm
        Unload PrintPreview
        
        Unload ErrorForm
        Unload ViewLog
        Unload GoToRecordForm
        Unload FindForm
        Unload ReplaceForm
        Unload LensDrawerSetup
        Unload PhotoPrinterSetup
        
        Unload MainForm
        EndTheProgram = True
        End
    Else
        EndTheProgram = False
    End If
End Function




'****************************************************************************
'**                                                                        **
'**  Procedure....:  CreateEditForm                                        **
'**                                                                        **
'**  Description..:  This routine creates a new Edit Form, displays it and **
'**                  adds it to the form collection for proper handling.   **
'**                                                                        **
'****************************************************************************
Public Sub CreateEditForm()
    On Error GoTo ErrorHandler                              'Standard Error Handler
    Dim EditFormPointer As New EditForm                     'Pointer to edit form
    Dim EditFormKey As String                               'Key to edit form
    
    EditFormCount = EditFormCount + 1                       'Increment the edit form count ('Untitled X' caption)
    EditFormKey = "Untitled" & str(EditFormCount)           'Assign the default new edit form key
    EditFormPointer.Caption = EditFormKey                   'Assign the key to the form caption
    EditFormPointer.Tag = EditFormKey
    EditFormPointer.Show vbModeless                         'Show the form
    
    MainForm.UpdateMenus
    
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "CreateEditForm", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  UnloadEditForm                                        **
'**                                                                        **
'**  Description..:  This routine unloads the current edit form.           **
'**                                                                        **
'****************************************************************************
Public Sub UnloadEditForm()
    On Error GoTo ErrorHandler                              'Standard Error Handler
    
    Dim EditFormPointer As Form                             'Pointer to edit form
    Dim EditFormKey As String                               'Key to edit form
    DoEvents
    EditFormKey = Screen.ActiveForm.Caption                 'The key is the form caption for the active window
    For Each EditFormPointer In Forms                       'For each form in the collection of forms
        If EditFormPointer.Tag = EditFormKey Then           'If the form key is the active window
            Unload EditFormPointer                          'Unload the form (Required to prevent memory leak)
            Exit For
        End If
    Next
    DoEvents
    
    MainForm.UpdateMenus
    
    
    Exit Sub
ErrorHandler:
    ErrorForm.ReportError "UnloadEditForm", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Sub

'****************************************************************************
'**                                                                        **
'**  Procedure....:  CheckForOpenFile                                      **
'**                                                                        **
'**  Description..:  This routine checks for shooter file already open.    **
'**                                                                        **
'****************************************************************************
Public Function CheckForFileOpen(FileName As String) As Boolean
    On Error GoTo ErrorHandler                              'Standard Error Handler
    CheckForFileOpen = False                                'Initialize to file not open
    Dim EditFormPointer As Form                             'Pointer to edit form
    Dim EditFormKey As String                               'Key to edit form
    DoEvents
    EditFormKey = FileName
    For Each EditFormPointer In Forms                       'For each form in the collection of forms
        If UCase(EditFormPointer.Caption) = UCase(EditFormKey) Then   'If a form caption is already set to this filename
            MsgBox "'" + FileName + "' is already open.", vbInformation + vbApplicationModal + vbOKOnly, "Error"
            CheckForFileOpen = True
            Exit For
        End If
    Next
    Exit Function
ErrorHandler:
    ErrorForm.ReportError "UnloadEditForm", Err.Number, Err.LastDllError, Err.Source, Err.Description, True
End Function

