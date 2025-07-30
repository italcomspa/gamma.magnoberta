Attribute VB_Name = "Mdl_Cmdialog"
Option Explicit

Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100


Const FW_NORMAL = 400
Const DEFAULT_CHARSET = 1
Const OUT_DEFAULT_PRECIS = 0
Const CLIP_DEFAULT_PRECIS = 0
Const DEFAULT_QUALITY = 0
Const DEFAULT_PITCH = 0
Const FF_ROMAN = 16
Const CF_PRINTERFONTS = &H2
Const CF_SCREENFONTS = &H1
Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Const CF_EFFECTS = &H100&
Const CF_FORCEFONTEXIST = &H10000
Const CF_INITTOLOGFONTSTRUCT = &H40&
Const CF_LIMITSIZE = &H2000&
Const REGULAR_FONTTYPE = &H400
Const LF_FACESIZE = 32
Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_ZEROINIT = &H40
Public Const DM_DUPLEX = &H1000&
Public Const DM_ORIENTATION = &H1&
Const PD_PRINTSETUP = &H40
Const PD_DISABLEPRINTTOFILE = &H80000

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type PRINTDLG_TYPE
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type


Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type PAGESETUPDLG
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    flags As Long
    ptPaperSize As POINTAPI
    rtMinMargin As RECT
    rtMargin As RECT
    hInstance As Long
    lCustData As Long
    lpfnPageSetupHook As Long
    lpfnPagePaintHook As Long
    lpPageSetupTemplateName As String
    hPageSetupTemplate As Long
End Type

Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 31
End Type

Private Type CHOOSEFONT
        lStructSize As Long
        hwndOwner As Long          '  caller's window handle
        hdc As Long                '  printer DC/IC or NULL
        lpLogFont As Long          '  ptr. to a LOGFONT struct
        iPointSize As Long         '  10 * size in points of selected font
        flags As Long              '  enum. type flags
        rgbColors As Long          '  returned text color
        lCustData As Long          '  data passed to hook fn.
        lpfnHook As Long           '  ptr. to hook function
        lpTemplateName As String     '  custom template name
        hInstance As Long          '  instance handle of.EXE that
                                       '    contains cust. dlg. template
        lpszStyle As String          '  return the style field here
                                       '  must be LF_FACESIZE or bigger
        nFontType As Integer          '  same value reported to the EnumFonts
                                       '    call back with the extra FONTTYPE_
                                       '    bits added
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long           '  minimum pt size allowed &
        nSizeMax As Long           '  max pt size allowed if
                                       '    CF_LIMITSIZE is used
End Type


Public Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type

Public Type DEVMODE_TYPE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type


Private Const FTP_TRANSFER_TYPE_UNKNOWN = &H0
Private Const FTP_TRANSFER_TYPE_ASCII = &H1
Private Const FTP_TRANSFER_TYPE_BINARY = &H2
Private Const INTERNET_DEFAULT_FTP_PORT = 21               ' default for FTP servers
Private Const INTERNET_SERVICE_FTP = 1
Private Const INTERNET_FLAG_PASSIVE = &H8000000            ' used for FTP connections
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0                    ' use registry configuration
Private Const INTERNET_OPEN_TYPE_DIRECT = 1                        ' direct to net
Private Const INTERNET_OPEN_TYPE_PROXY = 3                         ' via named proxy
Private Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY = 4   ' prevent using java/script/INS

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


Private Declare Function SHBrowseForFolder Lib "shell32" _
    (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
    (ByVal pidList As Long, _
    ByVal lpBuffer As String) As Long


Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
    (ByVal lpString1 As String, ByVal _
    lpString2 As String) As Long



Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long


Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Function PAGESETUPDLG Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PAGESETUPDLG) As Long
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Long
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean
Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hConnect As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Long, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByRef dwContext As Long) As Boolean
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hConnect As Long, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Boolean
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
'Private Const PassiveConnection As Boolean = True
Private PassiveConnection As Boolean


Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long



Public OFName As OPENFILENAME
Dim CustomColors() As Byte

Public hConnection As Long
Public hOpen As Long
Public WorkFolder As String


Public Function GetFolderName(Title As String) As String
    'Opens a Treeview control that displays
    '     the directories in a computer
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    szTitle = Title


    With tBrowseInfo
        .hwndOwner = 0 'Me.hwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)


    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    End If
    
    GetFolderName = sBuffer
End Function

Public Function ShowPrinter(frmOwner As Form, Optional PrintFlags As Long) As Boolean
    '-> Code by Donald Grover
    Dim PrintDlg As PRINTDLG_TYPE
    Dim DevMode As DEVMODE_TYPE
    Dim DevName As DEVNAMES_TYPE

    Dim lpDevMode As Long, lpDevName As Long
    Dim bReturn As Integer
    Dim objPrinter As Printer, NewPrinterName As String

    ' Use PrintDialog to get the handle to a memory
    ' block with a DevMode and DevName structures
    ShowPrinter = True
    PrintDlg.lStructSize = Len(PrintDlg)
    PrintDlg.hwndOwner = frmOwner.hwnd

    PrintDlg.flags = PrintFlags
    On Error Resume Next
    'Set the current orientation and duplex setting
    DevMode.dmDeviceName = Printer.DeviceName
    DevMode.dmSize = Len(DevMode)
    DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
    DevMode.dmPaperWidth = Printer.Width
    DevMode.dmOrientation = Printer.Orientation
    DevMode.dmPaperSize = Printer.PaperSize
    DevMode.dmDuplex = Printer.Duplex
    On Error GoTo 0

    'Allocate memory for the initialization hDevMode structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
    End If

    'Set the current driver, device, and port name strings
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With

    With Printer
        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With

    'Allocate memory for the initial hDevName structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If

    'Call the print dialog up and let the user make changes
    If PrintDialog(PrintDlg) <> 0 Then

        'First get the DevName structure.
        lpDevName = GlobalLock(PrintDlg.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree PrintDlg.hDevNames

        'Next get the DevMode structure and set the printer
        'properties appropriately
        lpDevMode = GlobalLock(PrintDlg.hDevMode)
        CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
        GlobalFree PrintDlg.hDevMode
        NewPrinterName = UCase$(Left(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
        If Printer.DeviceName <> NewPrinterName Then
            For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                    Set Printer = objPrinter
                    'set printer toolbar name at this point
                End If
            Next
        End If

        On Error Resume Next
        'Set printer object properties according to selections made
        'by user
        Printer.Copies = DevMode.dmCopies
        Printer.Duplex = DevMode.dmDuplex
        Printer.Orientation = DevMode.dmOrientation
        Printer.PaperSize = DevMode.dmPaperSize
        Printer.PrintQuality = DevMode.dmPrintQuality
        Printer.ColorMode = DevMode.dmColor
        Printer.PaperBin = DevMode.dmDefaultSource
        On Error GoTo 0
    Else
        ShowPrinter = False
    End If
End Function


Public Function OpenFtp(CurServer As String, CurUser As String, CurPwd As String, CurPort As String, CurFolder As String, Tipo As Integer) As Long
Dim sOrgPath  As String
Dim CurFileName As String
Dim Folders() As String
Dim i As Long
hOpen = InternetOpen("FTP SIFTE", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)

If Tipo = 0 Then
  PassiveConnection = True
Else
  PassiveConnection = False
End If

hConnection = InternetConnect(hOpen, CurServer, INTERNET_DEFAULT_FTP_PORT, CurUser, CurPwd, INTERNET_SERVICE_FTP, IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)
 
If NVL(CurFolder, "") <> "" Then
    Folders() = Split(CurFolder, "\")

    For i = 0 To UBound(Folders())
        OpenFtp = FtpSetCurrentDirectory(hConnection, Folders(i))
    Next
End If
OpenFtp = hConnection

End Function

Public Function DeleteFromFtp(DelFile As String) As Long
'FtpSetCurrentDirectory hConnection, DelFolder
DeleteFromFtp = FtpDeleteFile(hConnection, DelFile)
End Function

Public Function TransferFtp(CurPathFile As String, UploadFileName As String, UploadFileNameRen As String) As Long
  
  Dim ric As Boolean
  'Invio file
  TransferFtp = FtpPutFile(hConnection, CurPathFile, UploadFileName, FTP_TRANSFER_TYPE_UNKNOWN, 0)

  'Rinomino file
  ric = FtpRenameFile(hConnection, UploadFileName, UploadFileNameRen)

End Function


Public Function GetFileFromFtp(FileToGet As String, StorePath As String) As Long
'retrieve the file from the FTP server
GetFileFromFtp = FtpGetFile(hConnection, FileToGet, StorePath, False, 0, FTP_TRANSFER_TYPE_UNKNOWN, 0)

End Function

Public Function CloseFtp() As Long
On Error Resume Next
InternetCloseHandle hConnection
CloseFtp = InternetCloseHandle(hOpen)
End Function



Public Sub GetMultipleFileFromFtp(hConnection As Long, LocalDir As String, Estensione As String)
  Dim pData As WIN32_FIND_DATA, hFind As Long, lRet As Long
  
  Dim flgFileImport As Boolean
  Dim sOrgPath As String
  Dim pos As Integer
  Dim StripNull As String
  
  'create a buffer to store the original directory
  sOrgPath = String(MAX_PATH, 0)
  
  'get the directory
  Call FtpGetCurrentDirectory(hConnection, sOrgPath, Len(sOrgPath))
  
  'create a buffer
  pData.cFileName = String(MAX_PATH, 0)
  
  'find the first file
'  hFind = FtpFindFirstFile(hConnection, "*" & Estensione, pData, 0, 0)
  hFind = FtpFindFirstFile(hConnection, "*.txt", pData, 0, 0)
  
  'if there's no file, then exit sub
  If hFind = 0 Then
    Exit Sub
  Else
    pos = InStr(pData.cFileName, Chr$(0))
    
    If pos Then
       StripNull = Left$(pData.cFileName, pos - 1)
    Else
       StripNull = pData.cFileName
    End If
    
    flgFileImport = False
    'Se il file è di import lo scarico altrimenti non lo elaboro
    Select Case Mid(StripNull, 1, 2)
    Case "DR"
      '- Import ddt ricevimento (tipo DR)
      flgFileImport = True
    Case "RC"
      '- Import resi cliente (tipo RC)
      flgFileImport = True
    Case "DS"
      '- Import ddt cliente (tipo DS)
      flgFileImport = True
    Case "RT"
      flgFileImport = True
    End Select
    
    If flgFileImport Then
      'Copia il file nella directory
      'retrieve the file from the FTP server
      Call FtpGetFile(hConnection, StripNull, LocalDir & "\" & StripNull, False, 0, FTP_TRANSFER_TYPE_UNKNOWN, 0)
      'delete the file from the FTP server
      Call FtpDeleteFile(hConnection, StripNull)
    End If
    
    
    Do
      'create a buffer
      pData.cFileName = String(MAX_PATH, 0)
      'find the next file
      lRet = InternetFindNextFile(hFind, pData)
      
      'if there's no next file, exit do
      If lRet = 0 Then
        Exit Do
      Else
        pos = InStr(pData.cFileName, Chr$(0))
        
        If pos Then
           StripNull = Left$(pData.cFileName, pos - 1)
        Else
           StripNull = pData.cFileName
        End If
        
        flgFileImport = False
        'Se il file è di import lo scarico altrimenti non lo elaboro
        Select Case Mid(StripNull, 1, 2)
        Case "DR"
          '- Import ddt ricevimento (tipo DR)
          flgFileImport = True
        Case "RC"
          '- Import resi cliente (tipo RC)
          flgFileImport = True
        Case "DS"
          '- Import ddt cliente (tipo DS)
          flgFileImport = True
        Case "RT"
          flgFileImport = True
        End Select
        
        If flgFileImport Then
        
          'Copia il file nella directory
          'retrieve the file from the FTP server
          Call FtpGetFile(hConnection, StripNull, LocalDir & "\" & StripNull, False, 0, FTP_TRANSFER_TYPE_UNKNOWN, 0)
          
          'delete the file from the FTP server
          Call FtpDeleteFile(hConnection, StripNull)
        End If
      End If
    Loop
  End If
  
  Call InternetCloseHandle(hFind)
  Call InternetCloseHandle(lRet)
  
End Sub


Public Sub SpostaFile(FileOrigine As String, FileDestinazione As String)
  Dim Ret
  
  Ret = MoveFile(FileOrigine, FileDestinazione)
  
End Sub


Public Sub CopiaFile(FileOrigine As String, FileDestinazione As String)
  Dim Ret
  
  Ret = CopyFile(FileOrigine, FileDestinazione, True)
  
End Sub

Public Sub CancellaFile(File As String)
  Dim Ret
  
  Ret = DeleteFile(File)
  
End Sub

