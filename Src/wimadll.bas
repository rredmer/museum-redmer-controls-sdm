Attribute VB_Name = "Wimadll"

Public Const MAXLFN = 256

Public Const SORT_NONE = 72
Public Const SORT_NAME = 73
Public Const SORT_EXT = 74
Public Const SORT_SIZE = 75
Public Const SORT_DATE = 76

Public Const CDM_ROOT = 50
Public Const CDM_UPPER = 51
Public Const CDM_ENTRY = 52

'Values for Floppy Density
Public Const FLOPPY_160K = 0
Public Const FLOPPY_180K = 1
Public Const FLOPPY_320K = 2
Public Const FLOPPY_360K = 3
Public Const FLOPPY_720K = 4
Public Const FLOPPY_1200K = 5
Public Const FLOPPY_1440K = 6
Public Const FLOPPY_2880K = 7
Public Const FLOPPY_DMF2048K = 8
Public Const FLOPPY_DMF1024K = 9
Public Const FLOPPY_1680K = 10


' value for CaRead or CaCompare or CaWrite or CaFormat
Public Const FL_NOTHING = 0
Public Const FL_USED = 1
Public Const FL_ALL = 2
Public Const FL_BEGINFLOPPY = 3

Type DIRINFO
    nom(1 To 8) As Byte
    ext(1 To 3) As Byte
    szCompactName(1 To 13) As Byte
    bAttr As Byte

    dir_CreateMSec As Byte
    dir_CreateDate As Integer

    DosTime As Integer
    DosDate As Integer

    fIsSubDir As Long
    fSel As Long 'Boolean
    fLfnEntry As Long
    dwSize As Long
    uiPosInDir As Long

    dwLocalisation As Long
    dwTrueSize As Long
    longname(1 To MAXLFN) As Byte
    dir_CreateTime As Integer
    dir_LastAccessDate As Integer
End Type

Type ASPIINQUIRYTAB
    dwSizeStruct As Long
    dwHost As Long
    dwTargetID As Long
    dwTargetType As Long
    szDeviceName(1 To 32) As Byte
End Type

' CreateMemFatHima : Create an Image Object.
' you need call ReadImaFile, ReadFloppy or MakeEmptyImage
Declare Function CreateMemFatHima Lib "wimadll.dll" () As Long

' CreateMemHfsHima : Create an Image Object for Mac floppy.
' you need call ReadImaFile, ReadFloppy
' extract, inject... cannot be used
Declare Function CreateMemHfsHima Lib "wimadll.dll" () As Long

' CreateCDIsoIma : Create an Image Object by loading CDRom ISO image
'  lpFn : Filename of .ISO file
'  inject,...cannot be used
Declare Function CreateCDIsoIma Lib "wimadll.dll" (ByVal lpFn As String) As Long

' DeleteIma : Delete an Image Object.
Declare Sub DeleteIma Lib "wimadll.dll" (ByVal Ima As Long)


' ReadImaFile: Read an image file (.IMA or .IMZ)
'  hWnd : parent window for progress window
'  lpFn : FileName
'  lpfCompr : pointer to Boolean (will receive TRUE if file is compressed)
'  dwPosFileBegin : position in file (usualy 0, except in WLZ)
Declare Function ReadImaFile Lib "wimadll.dll" (ByVal Ima As Long, ByVal hWnd As Long, _
        ByVal lpFn As String, lpfCompr As Boolean, ByVal dwPosFileBegin As Long) As Boolean
Declare Function ReadImaFileEx Lib "wimadll.dll" (ByVal Ima As Long, ByVal hWnd As Long, _
        ByVal lpFn As String, lpfCompr As Boolean, ByVal dwPosFileBegin As Long, _
                ByVal lpszPassword As String) As Boolean

' WriteImaFile : WriteCompressed image
'  hWnd : parent window for progress window
'  lpFn : FileName
'  fTruncate : TRUE if you want truncate unused part of image
'  fCompress : TRUE if you want compress
'  iLevelCompress : used is fCompress is TRUE, level of compress (1 to 9)
'  dwPosBeginWrite : position in file (usualy 0)
'  lpNameInCompr : alternate name in compressed file (can be NULL)
Declare Function WriteImaFile Lib "wimadll.dll" (ByVal Ima As Long, ByVal hWnd As Long, _
        ByVal lpFn As String, ByVal fTruncate As Boolean, ByVal fCompr As Boolean, _
        ByVal iLevelCompress As Long, ByVal dwPosBeginWrite As Long, _
        ByVal lpNameInCompr As String) As Boolean
Declare Function WriteImaFileEx Lib "wimadll.dll" (ByVal Ima As Long, ByVal hWnd As Long, _
        ByVal lpFn As String, ByVal fTruncate As Boolean, ByVal fCompr As Boolean, _
        ByVal iLevelCompress As Long, ByVal dwPosBeginWrite As Long, _
        ByVal lpNameInCompr As String, ByVal lpszPassword As String) As Boolean



'  ReadFloppy : Read a floppy
'  hWnd : parent window for progress window
'  bFloppy : Floppy to read (0 for A:)
'  caRead : USED, or ALL (ALL if you want read unused part of floppy)
' BOOL WIMAAPI ReadFloppy(HIMA hIma,HWND hWnd,BYTE bFloppy,CHOICEAPP caRead);
Declare Function ReadFloppy Lib "wimadll.dll" (ByVal Ima As Long, ByVal hWnd As Long, _
        ByVal bFloppy As Byte, ByVal caRead As Long) As Boolean


' WriteFloppy : Write a floppy
'  hWnd : parent window for progress window
'  bFloppy : Floppy to write (0 for A:)
'  caFormat : NOTHING or ALL (ALL for format)
'  caWrite : USED or ALL
'  caCompare : NOTHING, USED or ALL
'  fCheckDiskBeforeWrite : if you want check disk is empty
'BOOL WIMAAPI WriteFloppy(HIMA hIma,HWND hWnd,BYTE bFloppy,CHOICEAPP caFormat,
'                        CHOICEAPP caWrite,CHOICEAPP caCompare,
'                        BYTE fCheckDiskBeforeWrite);alias
Declare Function WriteFloppy Lib "wimadll.dll" (ByVal Ima As Long, _
    ByVal hWnd As Long, ByVal bFloppy As Byte, _
    ByVal caFormat As Long, ByVal caWrite As Long, ByVal caCompare As Long, _
    ByVal fCheckDiskBeforeWrite As Byte) As Boolean


' Create a directory in the image
'  lpDir : Directory name
' BOOL WIMAAPI MkDir(HIMA hIma,LPSTR lpDir);
Declare Function MkDir Lib "wimadll.dll" (ByVal Ima As Long, ByVal lpDir As String) As Boolean

' Change current directory by name
'  lpDir : Directory name
' BOOL WIMAAPI ChszDir(HIMA hIma,LPSTR lpDir);
Declare Function ChszDir Lib "wimadll.dll" (ByVal Ima As Long, ByVal lpDir As String) As Boolean

' Change current directory by mode
'  bMode : CDM_ROOT or CDM_UPPER (equiv. to cd \ and cd ..)
' BOOL WIMAAPI ChDir(HIMA hIma,BYTE bMode);
Declare Function ChDir Lib "wimadll.dll" (ByVal Ima As Long, ByVal bMode As Byte) As Boolean

' InjectFile : Inject a file in floppy
'  lpFn : file to inject
'  lpDwSize : Pointer to DWORD that will receive the size. Can be NULL.
'  lpTooBig : Pointer to BOOL, become TRUE if file too big to be injected
'      (if InjectFile return FALSE). Can be NULL.
'  lpNameWhenInjected : if not NULL, contain a new name in the image
'      (if the file must have another name when injected). Can be NULL.
'BOOL WIMAAPI InjectFile(HIMA hIma,LPSTR lpFn,
'                        LPDWORD lpDwSize,LPBOOL lpTooBig,
'                        LPSTR lpNameWhenInjected);
Declare Function InjectFile Lib "wimadll.dll" (ByVal Ima As Long, ByVal lpDir As String, _
        lpdwSize As Long, lpTooBig As Boolean, ByVal lpNameWhenInjected As String) As Boolean


' MakeEmptyImage : make an empty image
' iNotypeDisk : 4=720K,6=1440K,7=2880K,8=DMF2048,9=DMF1024,10=1680K
'                  0=160K,1=180K,2=320K,3=360K,5=1200K (old, no ! :-))
'BOOL WIMAAPI MakeEmptyImage(HIMA hIma,int iNoTypeDisk);
Declare Function MakeEmptyImage Lib "wimadll.dll" (ByVal Ima As Long, ByVal iNoTypeDisk As Long) As Boolean

' InitWimaSdk : Init the DLL and use hinstdll for resource
'#define DEBENUSTD "ENU"
' #define BASEENUSTD (10000)
Const DEBENUSTD = "ENU"
Const BASEENUSTD = (10000)
' BOOL WIMAAPI InitWimaSdk(HINSTANCE hinstdll,LPSTR lpDeb,WORD wBase);
Declare Function InitWimaSdk Lib "wimadll.dll" (ByVal hinstdll As Long, ByVal lpDeb As String, _
        ByVal wBase As Integer) As Boolean

' GetCurDir : Get the name of current directory
'  lpBuf : buffer that will receive the name
'  uiMaxSize : the size of buffer
' BOOL WIMAAPI GetCurDir(HIMA hIma,LPSTR lpBuf,UINT uiMaxSize);
Declare Function GetCurDir Lib "wimadll.dll" (ByVal Ima As Long, ByVal lpBuf As String, _
                ByVal uiMaxSize As Long) As Boolean

' GetNbEntryCurDir : Get the number of entry of cur directory
' DWORD WIMAAPI GetNbEntryCurDir(HIMA hIma);
Declare Function GetNbEntryCurDir Lib "wimadll.dll" (ByVal Ima As Long) As Long

' GetDirInfo : Get info about the entry of cur directory
'  LPDIRINFO : array of DIRINFO that will receive the info
'                  (use GetNbEntryCurDir for know the size needed)
'  bSort :     specify how the file must be sort
'          (SORT_NONE, SORT_NAME, SORT_EXT, SORT_SIZE or SORT_DATE)
' BOOL WIMAAPI GetDirInfo(HIMA hIma,LPDIRINFO lpdi,BYTE bSort);
'' GetDirInfo and Sort MUST BE CHECKED IN BASIC!!!
Declare Function GetDirInfo Lib "wimadll.dll" (ByVal Ima As Long, di As DIRINFO, ByVal bSort As Byte) As Boolean

' Sort : Resort the array obtained by GetDirInfo
'BOOL WIMAAPI Sort(HIMA hIma,LPDIRINFO lpdi,BYTE bSort);
Declare Function Sort Lib "wimadll.dll" (ByVal Ima As Long, di As DIRINFO, ByVal bSort As Byte) As Boolean

' GetLabel : Get the label of Image
'  lpBuf : will receive the label
'BOOL WIMAAPI GetLabel(HIMA hIma,LPSTR lpBuf);
Declare Function GetLabel Lib "wimadll.dll" (ByVal Ima As Long, ByVal lpDir As String) As Boolean


' SetLabel : Set the label of Image
'  lpBuf : contain the new label
' BOOL WIMAAPI SetLabel(HIMA hIma,LPSTR lpBuf);
Declare Function SetLabel Lib "wimadll.dll" (ByVal Ima As Long, ByVal lpDir As String) As Boolean

' ExtractFile : Extract one file
'  unPosDir :  The uiPosInDir fields in DIRINFO structure that describe
'                  the file
'  lpPath :    Path where extract the file
'  lpFullName: will receive the exact full name of created file. Can be NULL
'BOOL WIMAAPI ExtractFile(HIMA hIma,UINT uiPosDir,LPSTR lpPath,LPSTR lpFullName);
Declare Function ExtractFile Lib "wimadll.dll" (ByVal Ima As Long, ByVal uiPosDir As Long, _
                                ByVal lpPath As String, ByVal lpFullName As String) As Boolean


' CheckSpaceForFile : Check you've space for inject a file of dwSize bytes
' BOOL WIMAAPI CheckSpaceForFile(HIMA hIma,DWORD dwSize);
Declare Function CheckSpaceForFile Lib "wimadll.dll" (ByVal Ima As Long, ByVal dwSize As Long) As Boolean

' to know if an inject is possible but need replace
'  lpFn : contain the name of file to be injected
'  lpDwSize : will receive the size of old file with same name. Can be NULL
'  lpNameWhenInjected : if not NULL, contain a new name in the image
'  lpShortName : will receive the short (8) name of file in image. Can be NULL
'  lpShortExt  : will receive the short (3) ext of file in image. Can be NULL
'      (if the file must have another name when injected)
'BOOL WIMAAPI IfInjectPossibleButNeedReplace(HIMA hIma,LPSTR lpFn,
'         LPDWORD lpDwSize,LPSTR lpShortName,
'         LPSTR lpShortExt,LPSTR lpNameWhenInjected);
Declare Function IfInjectPossibleButNeedReplace Lib "wimadll.dll" (ByVal Ima As Long, ByVal lpFn As String, _
                    lpdwSize As Long, ByVal lpShortName As String, ByVal lpShortExt As String, _
                    ByVal lpNameWhenInjected As String) As Boolean

' RmDir : Remove a directory
'  unPosDir :  The uiPosInDir fields in DIRINFO structure that describe
'                  the file
'BOOL WIMAAPI RmDir(HIMA hIma,UINT uiPosDir);
Declare Function RmDir Lib "wimadll.dll" (ByVal Ima As Long, ByVal uiPosDir As Long) As Boolean

' DeleteFileNameExt
' BOOL WIMAAPI DeleteFileNameExt(HIMA hIma,LPSTR lpNom,LPSTR lpExt,BOOL fRealDel);
Declare Function DeleteFileNameExt Lib "wimadll.dll" (ByVal Ima As Long, ByVal lpNom As String, _
                                ByVal lpExt As String, ByVal fRealDel As Boolean) As Boolean


' RenameFile :    Rename one file
'  uiPosDir :     The uiPosInDir fields in DIRINFO structure that describe
'                  the file
'  lpNewLongName: The new name of the file
' BOOL RenameFile(HIMA hIma,UINT uiPosDir,LPCSTR lpNewLongName);
Declare Function RenameFile Lib "wimadll.dll" (ByVal Ima As Long, ByVal uiPosDir As Long, _
                                               ByVal lpNewLongName As String) As Boolean

' ChangeDateAndAttribute :    Change the date and attribute of a File
'  uiPosDir :     The uiPosInDir fields in DIRINFO structure that describe
'                  the file
'  *lpbNewAttr:   Contain the new attribute of the file (or NULL to no change)
'  *lpNewDosDate,
'  *lpNewDosTime: Contain the Modified Date and Time (or NULL to no change)
'  *lpbNewdir_CreateMSec,*lpwNewdir_CreateTime,*lpwNewdir_CreateDate
'                 Contain the Created Date and Time (or NULL to no change)
'  *lpwNewdir_LastAccessDate : Contain the Last Access Date (or NULL...)
' BOOL ChangeDateAndAttribute(HIMA hIma,UINT uiPosDir,LPBYTE lpbNewAttr,
'                                     LPWORD lpNewDosDate,LPWORD lpNewDosTime,
'                                     LPBYTE lpbNewdir_CreateMSec,
'                                     LPWORD lpwNewdir_CreateTime,LPWORD lpwNewdir_CreateDate,
'                                     LPWORD lpwNewdir_LastAccessDate);
Declare Function ChangeDateAndAttribute Lib "wimadll.dll" (ByVal Ima As Long, ByVal uiPosDir As Long, _
                                     NewAttr As Byte, _
                                     NewDosDate As Integer, NewDosTime As Integer, _
                                     Newdir_CreateMSec As Byte, _
                                     lpwNewdir_CreateTime As Integer, Newdir_CreateDate As Integer, _
                                     lpwNewdir_LastAccessDate As Integer)

' ReadData : Direct read data in image.
'  dwPos :  begin position
'  dwSize : number of byte to copy (size of buffer)
'  lpBuf :  buffer that will receive data
'BOOL WIMAAPI ReadData(HIMA hIma,DWORD dwPos,DWORD dwSize,LPSTR lpBuf);
Declare Function ReadData Lib "wimadll.dll" (ByVal Ima As Long, ByVal dwPos As Long, _
                                ByVal dwSize As Long, ByVal lpBuf As String) As Boolean

' WriteData : Direct write data in image. Be carreful, WI don't refresh dir!
'  dwPos :  begin position
'  dwSize : number of byte to copy (size of buffer)
'  lpBuf :  buffer that contain data
'BOOL WIMAAPI WriteData(HIMA hIma,DWORD dwPos,DWORD dwSize,LPSTR lpBuf);
Declare Function WriteData Lib "wimadll.dll" (ByVal Ima As Long, ByVal dwPos As Long, _
                                ByVal dwSize As Long, ByVal lpBuf As String) As Boolean


' To be added : DRIVEINFO, GetFatImaSizeFileName, GetDriveInfo


'//
'// GetFatImaSizeFileName : Get information about UNCOMPRESSED Fat image
'//   lpfn :          FileName
'//   lpdwSize :      Will receive the size of the image, 32 bits low part of 64 bit data
'//   lpdwSize!high : Will receive the size of the image, 32 bits high part of 64 bit data
'//   lpfIsBigFat :   Boolean pointer, will receive TRUE if this is a large image (>2.88MB), not floppy image
'//   lpdwPosInFile : Will receive the position of the image
'BOOL WIMAAPI GetFatImaSizeFileName(LPCSTR lpFn,LPDWORD lpdwSize,LPDWORD lpdwSizeHigh,LPBOOL lpfIsBigFat,LPDWORD lpdwPosInFile);
Declare Function GetFatImaSizeFileName Lib "wimadll.dll" (ByVal lpFn As String, lpdwSize As Long, _
                                                          lpfIsBigFat As Boolean, lpdwPosInFile As Long) As Boolean


'// GetDriveInfo : Get info about drive type
'//  bDrive : number of driver (0 = 'A:', 1 = 'B:')
'//  return the kind of drive
'DriveInfo return:
'      NO_FLOPPY=0,
'      FLOPPY_360=1,
'      FLOPPY_12M=2,
'      FLOPPY_720=3,
'      FLOPPY_144=4,
'      FLOPPY_288=5,
'      LDISK_REMOVABLE=6,
'      LDISK_HARDDISK=7,
'      LDISK_CDROM=8,
'      FLOPPY_LS120=9
'DRIVEINFO WIMAAPI GetDriveInfo(BYTE bDrive);
Declare Function GetDriveInfo Lib "wimadll.dll" (ByVal bDrive As Byte) As Long


'// Fill the ASPI Inquiry array.
'// if lpAspiCdRomInquityTab is NULL AND dwMaxNumberInArray==0, just return the number of ASPI CDrom Unit.
'//  lpAspiCdRomInquityTab : Will receive the Array of SCSI Unit
'//  dwMaxNumberInArray : size of array (in number of ASPIINQUIRYTAB)
'DWORD WIMAAPI WimLargeAspiCdromInquiryFillArray(ASPIINQUIRYTAB* lpAspiCdRomInquityTab,DWORD dwMaxNumberInArray);
Declare Function WimLargeAspiCdromInquiryFillArray Lib "wimadll.dll" (AspiInqTab As ASPIINQUIRYTAB, _
                                                                      ByVal dwMaxNumberInArray As Long) As Long

'// Create a CDRom Image fro ASPI Unit, using dwHost and dwTargetID from AspiCdromInquiy
'//   lpFn : Filename to create
'//   lpdwTotal : will receive the filesize
'// Note : I suggest using WimLargeReadAspiCDImageIgnoreError with fIgnoreError at FALSE
'BOOL WIMAAPI WimLargeReadAspiCDImage(HWND hWnd,DWORD dwHost,DWORD dwTargetID,LPSTR lpFn,LPDWORD lpdwTotal);
Declare Function WimLargeReadAspiCDImage Lib "wimadll.dll" (ByVal hWnd As Long, _
                              ByVal dwHost As Long, ByVal dwTarget As Long, _
                                                          ByVal lpFn As String, lpdwTotal As Long) As Boolean


'// Like WimLargeReadAspiCDImage
'// fIgnoreError :
'//    FALSE : if there is error ignore it only if the error is after ISO9660 size (suggested)
'//    TRUE : Ignore all ISO 9660 error
'BOOL WIMAAPI WimLargeReadAspiCDImageIgnoreError(HWND hWnd,DWORD dwHost,DWORD dwTargetID,LPSTR lpFn,LPDWORD lpdwTotal,BOOL fIgnoreError);
Declare Function WimLargeReadAspiCDImageIgnoreError Lib "wimadll.dll" (ByVal hWnd As Long, _
                               ByVal dwHost As Long, ByVal dwTarget As Long, _
                               ByVal lpFn As String, lpdwTotal As Long, _
                                                           ByVal fIgnoreError As Boolean) As Boolean



'// return value != 0 if WimLargeReadLargeIma can be used with CDRom
'// (elsewhere, only hard disk partition)
'DWORD WIMAAPI WimLargeIsReadImaIsoPossible();
Declare Function WimLargeIsReadImaIsoPossible Lib "wimadll.dll" () As Long

'// Read Disk partition to image
'//  cDrive : disk letter ('C' for disk C:...)
'//  lpdwTotal : will receive number of byte processed
'//  caRead : USED, or ALL (ALL if you want read unused part of disk)
'BOOL WIMAAPI WimLargeReadLargeIma(HWND hWnd,char cDrive,LPSTR lpFn,LPDWORD lpdwTotal,CHOICEAPP caRead);
Declare Function WimLargeReadLargeIma Lib "wimadll.dll" (ByVal hWnd As Long, _
                               ByVal cDrive As Byte, _
                               ByVal lpFn As String, _
                               lpdwTotal As Long, _
                               ByVal caRead As Long) As Boolean

'// Write Disk partition from image
'//  cDrive : disk letter ('C' for disk C:...)
'//  lpdwTotal : will receive number of byte processed
'//  caWrite : USED or ALL
'//  fCheckDiskBeforeWrite : if you want check disk is empty
'BOOL WIMAAPI WimLargeWriteLargeIma(HIMA hIma,HWND hWnd,char cDrive,LPDWORD lpdwTotal,
'                                   CHOICEAPP caWrite,BOOL fCheckDiskBeforeWriteThis);
Declare Function WimLargeWriteLargeIma Lib "wimadll.dll" (ByVal Ima As Long, _
                               ByVal hWnd As Long, _
                               ByVal cDrive As Byte, lpdwTotal As Long, _
                                                           ByVal caWrite As Long, ByVal fCheckDiskBeforeWriteThis As Boolean) As Boolean

'// say if a letter if a CDRom
'BOOL WIMAAPI WimLargeIsIsoCDDrive(char cDrive);
Declare Function WimLargeIsIsoCDDrive Lib "wimadll.dll" (ByVal cDrive As Byte) As Boolean

'// Write the boot sector of an image
'BOOL WIMAAPI WriteSectBoot(HIMA hIma,const BYTE* lpBuf,DWORD dwSizeBuf);
Declare Function WriteSectBoot Lib "wimadll.dll" (ByVal Ima As Long, _
                     ByVal lpBuf As String, _
                                         ByVal dwSizeBuf As Long) As Boolean

'// Read the boot sector of an image
'BOOL WIMAAPI GetSectBoot(HIMA hIma,LPBYTE lpBuf,DWORD dwSizeBuf,LPDWORD lpdwSizeBoot);
Declare Function GetSectBoot Lib "wimadll.dll" (ByVal Ima As Long, _
                     ByVal lpBuf As String, _
                                         ByVal dwSizeBuf As Long, lpdwSizeBoot As Long) As Boolean

