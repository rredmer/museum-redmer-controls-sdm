Attribute VB_Name = "WimaFunctions"
Function WriteImageToFloppy(ImageFiles As String, SeeWindowsProgess As Boolean)
    Dim blnFileCompressed As Boolean
    Dim Ima As Long
    Dim ReturnValue As Boolean
    Dim EntryInImage As Long
    Dim WindowsProgress As Integer
    
    WriteImageToFloppy = False
    
    If SeeWindowsProgess Then
        WindowsProgress = 0
    Else
        WindowsProgress = 1
    End If
    ' Check if the file exist
    If Not FileExist(ImageFiles) Then MsgBox "The file " & ImageFiles & " do not exist!"
    
    Ima = CreateMemFatHima()
    
    'ReturnValue = MakeEmptyImage(Ima, 6)
    
    blnFileCompressed = False
    ReturnValue = ReadImaFile(Ima, 0, ImageFiles, blnFileCompressed, 0)
    EntryInImage = GetNbEntryCurDir(Ima)
    Call WriteFloppy(Ima, WindowsProgress, 0, FL_ALL, FL_ALL, FL_ALL, 0)
End Function
Function ReadFloppyToFile(ImageFiles As String, ImageLabel As String, SeeWindowsProgess As Boolean) As Boolean
    Dim blnFileCompressed As Boolean
    Dim Ima As Long
    Dim ReturnValue As Boolean
    Dim ReturnValue2 As Boolean
    Dim EntryInImage As Long
    Dim WindowsProgress, I As Integer
    Dim MYDIRINFO() As DIRINFO
    
    ReadFloppyToFile = False
    
    If SeeWindowsProgess Then
        WindowsProgress = 0
    Else
        WindowsProgress = 1
    End If
    
    Ima = CreateMemFatHima()
    
    If ImageLabel = "" Then ImageLabel = "NO LABEL"
    ' Set the label
    SetLabel Ima, ImageLabel
    
    ' Read Floppy
    If ReadFloppy(Ima, WindowsProgress, 0, FL_USED) Then
    'If ReadFloppy(Ima, WindowsProgress, 0, FL_ALL) Then
        ' Write image to File
        EntryInImage = GetNbEntryCurDir(Ima)
        If WriteImaFile(Ima, WindowsProgress, ImageFiles, True, True, 5, 0, ImageFiles) Then
            ReadFloppyToFile = True
            'MsgBox "File " & ImageFiles & " is done."
        End If
    End If
    '''''''''''''''
    ' GetDirInfo : Get info about the entry of cur directory
'  LPDIRINFO : array of DIRINFO that will receive the info
'                  (use GetNbEntryCurDir for know the size needed)
'  bSort :     specify how the file must be sort
'          (SORT_NONE, SORT_NAME, SORT_EXT, SORT_SIZE or SORT_DATE)
' BOOL WIMAAPI GetDirInfo(HIMA hIma,LPDIRINFO lpdi,BYTE bSort);
'' GetDirInfo and Sort MUST BE CHECKED IN BASIC!!!
'ReDim MYDIRINFO(EntryInImage)
'For I = 1 To EntryInImage
 '   Call GetDirInfo(Ima, MYDIRINFO(I), SORT_NAME)
 '       Debug.Print MYDIRINFO(I).bAttr
 '       Debug.Print MYDIRINFO(I).cReserved
 '       Debug.Print MYDIRINFO(I).cReserved2
 '       Debug.Print MYDIRINFO(I).DosDate
 '       Debug.Print MYDIRINFO(I).DosTime
 '       Debug.Print MYDIRINFO(I).dwLocalisation
 '       Debug.Print MYDIRINFO(I).dwSize
 '       Debug.Print MYDIRINFO(I).dwTrueSize
'       Debug.Print MYDIRINFO(I).ext
 '       Debug.Print MYDIRINFO(I).fIsSubDir
'       Debug.Print MYDIRINFO(I).fLfnEntry
 '       Debug.Print MYDIRINFO(I).fSel
 '       Debug.Print MYDIRINFO(I).longname
'        Debug.Print MYDIRINFO(I).nom
'        Debug.Print MYDIRINFO(I).szCompactName
'        Debug.Print MYDIRINFO(I).uiPosInDir
'Next I
    
Call DeleteIma(Ima)
End Function
Function FileExist(File As String) As Boolean
    Dim Exist As Boolean
    Dim FileNumber As Integer
    
    FileNumber = FreeFile
    
    Exist = True
    On Error GoTo FileError
    Open File For Input As #FileNumber
    If Exist Then
        Close #FileNumber
        FileExist = True
        Exit Function
    Else
        FileExist = False
    End If
Exit Function
FileError:
    'MsgBox Err.Number & " " & Error(Err)
    Select Case Err.Number  ' Evaluate error number.
        Case 53 ' "File not Exist" error.
            Exist = False
        Case Else
            ' Handle other situations here...
    End Select
    Resume Next
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub test()
  Dim blnFileCompressed As Boolean
  Dim dwPos As Long
  Dim Ima As Long
  Dim res As Boolean
  Dim res2 As Boolean
  Dim ent As Long
  Dim str As String
  Dim strsav As String

  str = "q:\image\testsdk\tst.ima"

  Ima = CreateMemFatHima()
   Rem res = ReadImaFile(Ima, 0, str, blnFileCompressed, dwPos)

' iNotypeDisk : 4=720K,6=1440K,7=2880K,8=DMF2048,9=DMF1024,10=1680K
'                  0=160K,1=180K,2=320K,3=360K,5=1200K (old, no ! :-))

  res = MakeEmptyImage(Ima, 6 + (2 * 1))
  SetLabel Ima, "BasicSdk"
  ent = GetNbEntryCurDir(Ima)

  Rem Declare Function InjectFile Lib "wimadll.dll" (ByVal Ima As Long,
'ByVal lpDir As String, _
  rem      lpDwSize As Long, lpTooBig As Boolean, ByVal lpNameWhenInjected
'As String) As Boolean


    res = InjectFile(Ima, "c:\boot.ini", dwPos, blnFileCompressed, "boot.ini")
    res = InjectFile(Ima, "c:\command.com", dwPos, blnFileCompressed, "COMMAND.COM")


  strsav = "q:\image\testsdk\tst3.imz"
   res2 = WriteImaFile(Ima, 0, strsav, True, True, 5, 0, "tst2.ima")
  DeleteIma Ima

' WriteImaFile : WriteCompressed image
'  hWnd : parent window for progress window
'  lpFn : FileName
'  fTruncate : TRUE if you want truncate unused part of image
'  fCompress : TRUE if you want compress
'  iLevelCompress : used is fCompress is TRUE, level of compress (1 to 9)
'  dwPosBeginWrite : position in file (usualy 0)
'  lpNameInCompr : alternate name in compressed file (can be NULL)
'Declare Function WriteImaFile Lib "wimadll.dll" (ByVal Ima As Long, ByVal
'hWnd As Long, _
'        ByVal lpFn As String, ByVal fTruncate As Boolean, ByVal fCompr As
'Boolean, _
'        ByVal iLevelCompress As Long, ByVal dwPosBeginWrite As Long, _
'        ByVal lpNameInCompr As String) As Boolean

' Read an image file (.IMA or .IMZ)
'  hWnd : parent window for progress window
'  lpFn : FileName
'  lpfCompr : pointer to Boolean (will receive TRUE if file is compressed)
'  dwPosFileBegin : position in file (usualy 0, except in WLZ)
' Declare Function ReadImaFile Lib "wimadll.dll" (ByVal Ima As Long, ByVal
'hWnd As Long, _
'         ByVal lpFn As String, lpfCompr As Boolean, ByVal dwPosFileBegin As
'Long) As Boolean

End Sub



