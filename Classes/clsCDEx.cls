VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCDEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type BROWSEINFO
   hOwner           As Long
   pidlRoot         As Long
   pszDisplayName   As String
   lpszTitle        As String
   ulFlags          As Long
   lpfn             As Long
   lParam           As Long
   iImage           As Long
End Type

Private Type OPENFILENAME
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

Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    RGBResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Const LF_FACESIZE = 32
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
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type ChooseFont
    lStructSize As Long
    hwndOwner As Long
    Hdc As Long
    lpLogFont As Long
    iPointSize As Long
    flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

Const CF_INITTOLOGFONTSTRUCT = &H40&
Const SCREEN_FONTTYPE = &H2000
Const BOLD_FONTTYPE = &H100
Const FW_BOLD = 700
Const LOGPIXELSY = 90

Private Type PrintDlg
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    Hdc As Long
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

Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32
Private Type DEVMODE
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

Private Type DEVNAMES
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type

Const DM_DUPLEX = &H1000&
Const DM_ORIENTATION = &H1&

Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ChooseColorAPI Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long
Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PrintDlg) As Long
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hOwner As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

Private Declare Function SHShutDownDialog Lib "Shell32" Alias "#60" (ByVal YourGuess As Long) As Long
Private Declare Function SHRestartSystem Lib "Shell32" Alias "#59" (ByVal hOwner As Long, ByVal sPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function SHRunDialog Lib "Shell32" Alias "#61" (ByVal hOwner As Long, ByVal hIcon As Long, ByVal sDir As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function SHFormatDrive Lib "Shell32" (ByVal hwndOwner As Long, ByVal iDrive As Long, ByVal iCapacity As Long, ByVal iFormatType As Long) As Long
Private Declare Function SHBrowseForFolder Lib "Shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHChangeIconDialog Lib "Shell32" Alias "#62" (ByVal hOwner As Long, ByVal szFilename As String, ByVal Reserved As Long, lpIconIndex As Long) As Long
Private Declare Function SHObjectProperties Lib "Shell32" Alias "#178" (ByVal hOwner As Long, ByVal uFlags As Long, ByVal sName As String, ByVal sParam As String) As Long
Private Declare Function SHAbout Lib "Shell32" Alias "ShellAboutA" (ByVal hOwner As Long, ByVal sAppName As String, ByVal sPrompt As String, ByVal hIcon As Long) As Long

Private Declare Function SHSimpleIDListFromPath Lib "Shell32" Alias "#162" (ByVal szPath As String) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHFree Lib "Shell32" Alias "#196" ()
Private Declare Function ILFree Lib "Shell32" Alias "#195" (ByVal pidlFree As Long)
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function ExtractIconEx Lib "Shell32" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal Hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal Hdc As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal Hdc As Long, ByVal nIndex As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hOwner As Long) As Long

Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (p1 As Any, p2 As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)


Public Enum CdlgExt_Actions
       cdlgOpen = 1
       cdlgSave = 2
       cdlgColor = 3
       cdlgFont = 4
       cdlgPrinter = 5
       cdlgHelp = 6
       cdlgAbout = 7
       cdlgFolder = 8
       cdlgFormat = 9
       cdlgIcon = 10
       cdlgRestart = 12
       cdlgRun = 13
       cdlgShutDown = 14
 End Enum
 
 Public Enum CdlgEx_IconSize
       IconSizeSmall = 16
       IconSizeLarge = 32
 End Enum

Public Enum CdlgExt_Flags
 cdlOFNAllowMultiselect = &H200
 cdlOFNCreatePrompt = &H2000
 cdlOFNExplorer = &H80000
 cdlOFNExtensionDifferent = &H400
 cdlOFNFileMustExist = &H1000
 cdlOFNHelpButton = &H10
 cdlOFNHideReadOnly = &H4
 cdlOFNLongNames = &H200000
 cdlOFNNoChangeDir = &H8
 cdlOFNNoDereferenceLinks = &H100000
 cdlOFNNoLongNames = &H40000
 cdlOFNNoReadOnlyReturn = &H8000
 cdlOFNNoValidate = &H100
 cdlOFNOverwritePrompt = &H2
 cdlOFNPathMustExist = &H800
 cdlOFNReadOnly = &H1
 cdlOFNShareAware = &H4000

 cdlCCFullOpen = &H2
 cdlCCHelpButton = &H8
 cdlCCPreventFullOpen = &H4
 cdlCCRGBInit = &H1

 cdlPDAllPages = &H0
 cdlPDCollate = &H10
 cdlPDDisablePrintToFile = &H80000
 cdlPDHelpButton = &H800
 cdlPDHidePrintToFile = &H100000
 cdlPDNoPageNums = &H8
 cdlPDNoSelection = &H4
 cdlPDNoWarning = &H80
 cdlPDPageNums = &H2
 cdlPDPrintSetup = &H40
 cdlPDPrintToFile = &H20
 cdlPDReturnDC = &H100
 cdlPDReturnDefault = &H400
 cdlPDReturnIC = &H200
 cdlPDSelection = &H1
 cdlPDUseDevModeCopies = &H40000

 cdlCFANSIOnly = &H400
 cdlCFApply = &H200
 cdlCFBoth = &H3
 cdlCFEffects = &H100
 cdlCFFixedPitchOnly = &H4000
 cdlCFForceFontExist = &H10000
 cdlCFHelpButton = &H4
 cdlCFLimitSize = &H2000
 cdlCFNoFaceSel = &H80000
 cdlCFNoSimulations = &H1000
 cdlCFNoSizeSel = &H200000
 cdlCFNoStyleSel = &H100000
 cdlCFNoVectorFonts = &H800
 cdlCFPrinterFonts = &H2
 cdlCFScalableOnly = &H20000
 cdlCFScreenFonts = &H1
 cdlCFTTOnly = &H40000
 cdlCFWYSIWYG = &H8000

 Restart_Logoff = &H0
 Restart_ShutDown = &H1
 Restart_Reboot = &H2
 Restart_Force = &H4

 Run_NoBrowse = &H10
 Run_NoDefault = &H20
 Run_CalcDir = &H40
 Run_NoLable = &H80

 ObjProp_System = &H0
 ObjProp_Printer = &H100
 ObjProp_File = &H200
 ObjProp_Mouse = &H300
 ObjProp_Locale = &H400
 ObjProp_MMedia = &H500
 ObjProp_TimeDate = &H600
 ObjProp_Network = &H700
 ObjProp_Screen = &H800
 ObjProp_Internet = &H900

 Folder_COMPUTER = &H1000
 Folder_PRINTER = &H2000
 Folder_INCLUDEFILES = &H4001
End Enum

Public Enum CdlgExt_HelpCommand
 HelpCommandHelp = &H102&
 HelpContents = &H3&
 HelpContext = &H1
 HelpContextPOPUP = &H8&
 HelpForceFile = &H9&
 HelpHelpOnHelp = &H4
 HelpIndex = &H3
 HelpKeyHelp = &H101
 HelpPartialKey = &H105&
 HelpQuit = &H2
 HelpSetContents = &H5&
 HelpSetIndex = &H5
 HelpMultiKey = &H201&
 HelpSetWinPos = &H203&
End Enum

Private RetValue As Long
Const MAX_PATH = 260
Private OFN As OPENFILENAME

Private mFileName As String
Private mFileTitle As String
Private mhOwner As Long
Private mDialogTitle As String
Private mFilter As String
Private mInitDir As String
Private mSelDir As String
Private mDefaultExt As String
Private mFilterIndex As Long
Private mHelpFile As String
Private mHelpCommand As CdlgExt_HelpCommand
Private mHelpKey As Long
Private mRGBResult As Long
Private mItalic As Boolean
Private mUnderline As Boolean
Private mStrikethru As Boolean
Private mFontName As String
Private mFontSize As Long
Private mBold As Boolean
Private mDialogPrompt As String
Private mFlags As CdlgExt_Flags
Private mCancelError As Boolean
Private mhIcon As Long
Private mAppName As String
Private mIconSize As CdlgEx_IconSize

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Const WM_USER = &H400
Const BFFM_INITIALIZED = 1
Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Public Property Let Action(ByVal New_Action As CdlgExt_Actions)
   
    On Error GoTo ActionError
   
    Select Case New_Action
     Case 1
      ShowOpen
     Case 2
      ShowSave
     Case 3
      ShowColor
     Case 4
      ShowFont
     Case 5
      ShowPrinter
     Case 6
      ShowHelp
     Case 7
      ShowAbout
     Case 8
      ShowFolder
     Case 9
      ShowFormat
     Case 10
      ShowIcon
     Case 11
      ShowObjectProp
     Case 12
      ShowRestart
     Case 13
      ShowRun
     Case 14
      ShowShutDown
     Case Else
    End Select

ActionError:
    Exit Property

End Property

Public Property Let CancelError(ByVal vData As Boolean)
    mCancelError = vData
End Property

Public Property Get CancelError() As Boolean
    CancelError = mCancelError
End Property

Public Property Get hOwner() As Long
    hOwner = mhOwner
End Property

Public Property Let hOwner(ByVal New_hOwner As Long)
    mhOwner = New_hOwner
End Property

Public Property Get flags() As CdlgExt_Flags
    flags = mFlags
End Property

Public Property Let flags(ByVal New_Flags As CdlgExt_Flags)
    mFlags = New_Flags
End Property

Public Property Get DialogTitle() As String
    DialogTitle = mDialogTitle
End Property

Public Property Let DialogTitle(sTitle As String)
    mDialogTitle = sTitle
End Property

Public Property Get DialogPrompt() As String
    DialogPrompt = mDialogPrompt
End Property

Public Property Let DialogPrompt(ByVal New_Prompt As String)
    mDialogPrompt = New_Prompt
End Property

Public Property Get AppName() As String
    AppName = mAppName
End Property

Public Property Let AppName(ByVal New_AppName As String)
    mAppName = New_AppName
End Property

Public Property Let hIcon(ByVal vData As Long)
    mhIcon = vData
End Property

Public Property Get hIcon() As Long
    hIcon = mhIcon
End Property

Public Property Get Bold() As Boolean
    Bold = mBold
End Property

Public Property Let Bold(bBold As Boolean)
    mBold = bBold
End Property

Public Property Get FontName() As String
    FontName = mFontName
End Property

Public Property Let FontName(sName As String)
   mFontName = sName
End Property

Public Property Get FontSize() As Long
    FontSize = mFontSize
End Property

Public Property Let FontSize(lSize As Long)
    mFontSize = lSize
End Property

Public Property Get Italic() As Boolean
    Italic = mItalic
End Property

Public Property Let Italic(BItalic As Boolean)
    mItalic = BItalic
End Property

Public Property Get StrikeThru() As Boolean
   StrikeThru = mStrikethru
End Property

Public Property Let StrikeThru(bStrikethru As Boolean)
   mStrikethru = bStrikethru
End Property

Public Property Get Underline() As Boolean
   Underline = mUnderline
End Property

Public Property Let Underline(bUnderline As Boolean)
   mUnderline = bUnderline
End Property
Public Property Get DefaultExt() As String
    DefaultExt = mDefaultExt
End Property

Public Property Let DefaultExt(sDefExt As String)
    mDefaultExt = DefaultExt
End Property

Public Property Get FileName() As String
    FileName = mFileName
End Property

Public Property Let FileName(sFileName As String)
    mFileName = sFileName
End Property

Public Property Get FileTitle() As String
    FileTitle = mFileTitle
End Property

Public Property Let FileTitle(sTitle As String)
    mFileTitle = sTitle
End Property

Public Property Get Filter() As String
    Filter = mFilter
End Property

Public Property Let Filter(sFilter As String)
    mFilter = sFilter
End Property

Public Property Get FilterIndex() As Long
    FilterIndex = mFilterIndex
End Property

Public Property Let FilterIndex(lIndex As Long)
    mFilterIndex = lIndex
End Property

Public Property Get InitDir() As String
    InitDir = mInitDir
End Property

Public Property Let InitDir(sDir As String)
    mInitDir = sDir
End Property

Public Property Get SelDir() As String
    SelDir = mSelDir
End Property

Public Property Let SelDir(sDir As String)
    mSelDir = sDir
End Property

Public Property Get IconSize() As CdlgEx_IconSize
  
    If mIconSize <> IconSizeLarge And mIconSize <> IconSizeSmall Then mIconSize = IconSizeLarge
    IconSize = mIconSize

End Property

Public Property Let IconSize(nSize As CdlgEx_IconSize)
   
   If nSize <> IconSizeLarge And nSize <> IconSizeSmall Then nSize = IconSizeLarge
   mIconSize = nSize

End Property
Public Property Get HelpCommand() As CdlgExt_HelpCommand
    HelpCommand = mHelpCommand
End Property

Public Property Let HelpCommand(lCommand As CdlgExt_HelpCommand)
    mHelpCommand = lCommand
End Property

Public Property Get HelpFile() As String
    HelpFile = mHelpFile
End Property

Public Property Let HelpFile(sFile As String)
    mHelpFile = sFile
End Property

Public Property Get HelpKey() As Long
    HelpKey = mHelpKey
End Property

Public Property Let HelpKey(sKey As Long)
    mHelpKey = sKey
End Property

Public Property Get RGBResult() As Long
    RGBResult = mRGBResult
End Property

Public Property Let RGBResult(lValue As Long)
    mRGBResult = lValue
End Property
Public Function ShowShutDown()
   SHShutDownDialog mhOwner
End Function
Public Function ShowRestart()
    
    On Error GoTo ShowRestartError
    Dim uFlag As Long
    uFlag = mFlags And (&H0 Or &H1 Or &H2 Or &H4)
    SHRestartSystem mhOwner, mDialogPrompt, uFlag

ShowRestartError:
    Exit Function

End Function
Public Function ShowRun()
  
    On Error GoTo RunError
  
    Dim uFlag As Long
    uFlag = mFlags And (&H10 Or &H20 Or &H40 Or &H80)
    uFlag = uFlag / 16
    SHRunDialog mhOwner, mhIcon, 0, mDialogTitle, mDialogPrompt, uFlag

RunError:
    Exit Function

End Function

Public Function ShowFormat(Optional ByVal iDrive As Long, Optional ByVal iCapacity As Long, Optional ByVal iFormatType As Long) As Long
    ShowFormat = SHFormatDrive(mhOwner, iDrive, iCapacity, iFormatType)
End Function
Public Function ShowIcon()
    
    On Error GoTo ShowIconError
    
    Dim nIconIdx As Long, OldFileName As String
    Dim hSmallIcon As Long, hLargeIcon As Long, NewIcon As Long
    If Right(mFileName, 1) = "\" Then Exit Function
    OldFileName = mFileName
    mFileName = mFileName & String$(MAX_PATH - Len(mFileName), 0) 'FileName  must be maximum lenth
    
    If SHChangeIconDialog(0, mFileName, 0, nIconIdx) Then
     If ExtractIconEx(mFileName, nIconIdx, hLargeIcon, hSmallIcon, 1) > 0 Then
      NewIcon = IIf(mIconSize - 32, hSmallIcon, hLargeIcon)
      mhIcon = CopyIcon(NewIcon)
      DestroyIcon hSmallIcon
      DestroyIcon hLargeIcon
      End If
    End If
    mFileName = OldFileName

ShowIconError:
    Exit Function

End Function
Public Function ShowFolder() As String
  
    On Error GoTo ShowFolderError
    Dim bi As BROWSEINFO
    Dim pidl As Long, path As String, pos As Integer, uFlag As Long
    Dim TopFolder, SelFolder
    TopFolder = TopFolder & Chr$(0)
    SelFolder = SelFolder & Chr$(0)
    
    bi.hOwner = mhOwner
    bi.pidlRoot = SHSimpleIDListFromPath(mInitDir)
    bi.lpszTitle = mDialogPrompt
    
    'Dont remove the item below ! ! !
    
    'bi.lpfn = FARPROC(AddressOf BrowseCallbackProc)
    
    bi.lParam = SHSimpleIDListFromPath(SelDir)
    uFlag = mFlags And (&H1000 Or &H2000 Or &H4001)
    If uFlag < Folder_COMPUTER Then
     bi.ulFlags = &H1
    Else
     bi.ulFlags = uFlag
    End If
    pidl = SHBrowseForFolder(bi)
    path = String$(MAX_PATH, 0)

    If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
     pos = InStr(path, Chr$(0))
     InitDir = Left(path, pos - 1)
    Else
     If mCancelError Then Err.Raise 0
    End If
    Call CoTaskMemFree(pidl)

ShowFolderError:
    Exit Function

End Function
Public Function FARPROC(pfn As Long) As Long
    FARPROC = pfn
End Function
Public Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
   
    Select Case uMsg
     Case BFFM_INITIALIZED
     Call SendMessage(hwnd, BFFM_SETSELECTIONA, False, ByVal lpData)
    Case Else:
    End Select

End Function
Public Function ShowObjectProp(Optional ByVal sObjectName As String, Optional ByVal sTab As String)
    
    On Error GoTo ShowObjectError
    Dim uFlag As Long, sObj As String
    Dim sPath As String
    uFlag = mFlags And (&H100 Or &H200 Or &H300 Or &H400 Or &H500 Or &H600 Or &H700 Or &H800 Or &H900)
    uFlag = uFlag / 256
    
    Select Case uFlag
     Case 1, 2
      sObj = sObjectName
     Case 3
      uFlag = 0
      Call Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl,,0", vbNormalFocus)
     Case 4
      uFlag = 0
      Call Shell("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,0", vbNormalFocus)
     Case 5
      uFlag = 0
      Call Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,0", vbNormalFocus)
     Case 6
      uFlag = 0
      Call Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl,,0", vbNormalFocus)
     Case 7
      uFlag = 0
      Call Shell("rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl,,0", vbNormalFocus)
     Case 8
      uFlag = 0
      Call Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0", vbNormalFocus)
     Case 9
      uFlag = 0
      Call Shell("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", vbNormalFocus)
     Case Else
      uFlag = 2
      sObj = ""
    End Select
    
    If uFlag > 0 Then SHObjectProperties mhOwner, uFlag, sObj, sTab

ShowObjectError:
    Exit Function

End Function

Public Function ShowAbout()
        
    On Error GoTo ShowAboutError
    
    If mAppName = "" Then mAppName = Chr$(0)
    SHAbout mhOwner, mAppName, mDialogPrompt, mhIcon

ShowAboutError:
    Exit Function

End Function
Public Sub ShowOpen()
  
    On Error GoTo ShowOpenError
    
    Dim iDelim As Integer
    InitOFN
    RetValue = GetOpenFileName(OFN)
    If RetValue > 0 Then
     iDelim = InStr(OFN.lpstrFileTitle, vbNullChar)
     If iDelim Then mFileTitle = Left$(OFN.lpstrFileTitle, iDelim - 1)
     iDelim = InStr(OFN.lpstrFile, vbNullChar)
     If iDelim Then mFileName = Left$(OFN.lpstrFile, iDelim - 1)
    Else
     If mCancelError Then Err.Raise 0
    End If

ShowOpenError:
    Exit Sub

End Sub
Public Sub ShowSave()
  
    On Error GoTo ShowSaveError
    Dim iDelim As Integer
    InitOFN
    RetValue = GetSaveFileName(OFN)
    If RetValue > 0 Then
     iDelim = InStr(OFN.lpstrFileTitle, vbNullChar)
     If iDelim Then mFileTitle = Left$(OFN.lpstrFileTitle, iDelim - 1)
     iDelim = InStr(OFN.lpstrFile, vbNullChar)
     If iDelim Then mFileName = Left$(OFN.lpstrFile, iDelim - 1)
    Else
     If mCancelError Then Err.Raise 0
    End If

ShowSaveError:
    Exit Sub

End Sub
Private Sub InitOFN()
    
    On Error GoTo InitError
    Dim sTemp As String, i As Integer
    Dim uFlag As Long
    uFlag = mFlags And (&H1 Or &H2 Or &H4 Or &H8 Or &H10 Or &H100 Or &H200 Or &H400 Or &H800 Or &H1000 Or &H2000 Or &H4000 Or &H8000 Or &H40000 Or &H80000 Or &H100000 Or &H200000)
    
    With OFN
     .lStructSize = Len(OFN)
     .hwndOwner = mhOwner
     .flags = uFlag
     .lpstrDefExt = mDefaultExt
     sTemp = mInitDir
     If sTemp = "" Then sTemp = App.path
     .lpstrInitialDir = sTemp
     sTemp = mFileName
     .lpstrFile = sTemp & String$(255 - Len(sTemp), 0)
     .nMaxFile = 255
     .lpstrFileTitle = String$(255, 0)
     .nMaxFileTitle = 255
     sTemp = mFilter
     For i = 1 To Len(sTemp)
      If Mid(sTemp, i, 1) = "|" Then
      Mid(sTemp, i, 1) = vbNullChar
      End If
     Next
     sTemp = sTemp & String$(2, 0)
     .lpstrFilter = sTemp
     .nFilterIndex = mFilterIndex
     .lpstrTitle = mDialogTitle
     .hInstance = App.hInstance
    End With

InitError:
    Exit Sub

End Sub
Public Sub ShowHelp()
    RetValue = WinHelp(mhOwner, mHelpFile, mHelpCommand, mHelpKey)
End Sub
Public Sub ShowColor()
    
    On Error GoTo ShowColorError
    Dim CC As CHOOSECOLOR
    Dim CustomColors() As Byte
    Dim uFlag As Long
    Dim i
    ReDim CustomColors(0 To 16 * 4 - 1) As Byte
    
    For i = LBound(CustomColors) To UBound(CustomColors)
     CustomColors(i) = 255 ' white
    Next i
    
    uFlag = mFlags And (&H1 Or &H2 Or &H4 Or &H8)
    
    With CC
     .lStructSize = Len(CC)
     .hwndOwner = mhOwner
     .hInstance = App.hInstance
     .lpCustColors = StrConv(CustomColors, vbUnicode)
     .flags = uFlag
     .RGBResult = mRGBResult
     RetValue = ChooseColorAPI(CC)
     If RetValue = 0 Then
      If mCancelError Then Err.Raise (RetValue)
     Else
      CustomColors = StrConv(.lpCustColors, vbFromUnicode)
      mRGBResult = .RGBResult
     End If
    End With

ShowColorError:
    Exit Sub
     
End Sub
Public Sub ShowFont()
    
    On Error GoTo ShowFontError
    Dim CF As ChooseFont
    Dim LF As LOGFONT
    Dim TempByteArray() As Byte
    Dim ByteArrayLimit As Long
    Dim OldhDC As Long
    Dim FontToUse As Long
    Dim tbuf As String * 80
    Dim X As Long
    Dim uFlag As Long
    uFlag = mFlags And (&H1 Or &H2 Or &H3 Or &H4 Or &H100 Or &H200 Or &H400 Or &H800 Or &H1000 Or &H2000 Or &H4000 Or &H8000 Or &H10000 Or &H20000 Or &H40000 Or &H80000 Or &H100000 Or &H200000)
    TempByteArray = StrConv(mFontName & vbNullChar, vbFromUnicode)
    ByteArrayLimit = UBound(TempByteArray)
    
    With LF
     For X = 0 To ByteArrayLimit
      .lfFaceName(X) = TempByteArray(X)
     Next
     .lfHeight = mFontSize / 72 * GetDeviceCaps(GetDC(mhOwner), LOGPIXELSY)
     .lfItalic = mItalic * -1
     .lfUnderline = mUnderline * -1
     .lfStrikeOut = mStrikethru * -1
     If mBold Then .lfWeight = FW_BOLD
    End With
    
    With CF
     .lStructSize = Len(CF)
     .hwndOwner = mhOwner
     .Hdc = GetDC(mhOwner)
     .lpLogFont = lstrcpy(LF, LF)
     If Not uFlag Then
      .flags = cdlCFScreenFonts
     Else
      .flags = uFlag Or cdlCFWYSIWYG
     End If
     .flags = .flags Or cdlCFEffects Or CF_INITTOLOGFONTSTRUCT
     .rgbColors = mRGBResult
     .lCustData = 0
     .lpfnHook = 0
     .lpTemplateName = 0
     .hInstance = 0
     .lpszStyle = 0
     .nFontType = SCREEN_FONTTYPE
     .nSizeMin = 0
     .nSizeMax = 0
     .iPointSize = mFontSize * 10
    End With
    RetValue = ChooseFont(CF)
    
    If RetValue = 0 Then
     If mCancelError Then Err.Raise (RetValue)
    Else
     With LF
      mItalic = .lfItalic * -1
      mUnderline = .lfUnderline * -1
      mStrikethru = .lfStrikeOut * -1
     End With
     With CF
      mFontSize = .iPointSize \ 10
      mRGBResult = .rgbColors
      If .nFontType And BOLD_FONTTYPE Then
       mBold = True
      Else
       mBold = False
      End If
     End With
     FontToUse = CreateFontIndirect(LF)
    If FontToUse = 0 Then Exit Sub
     OldhDC = SelectObject(CF.Hdc, FontToUse)
     RetValue = GetTextFace(CF.Hdc, 79, tbuf)
     mFontName = Mid$(tbuf, 1, RetValue)
    End If

ShowFontError:
    Exit Sub

End Sub
Public Sub ShowPrinter()
  
    Dim PD As PrintDlg
    Dim DM As DEVMODE
    Dim DN As DEVNAMES
    Dim lpDevMode As Long, lpDevName As Long
    Dim objPrinter As Printer, NewPrinterName As String
    Dim strSetting As String
    Dim uFlag As Long
    uFlag = mFlags And (&H0 Or &H1 Or &H2 Or &H4 Or &H8 Or &H10 Or &H20 Or &H40 Or &H80 Or &H100 Or &H200 Or &H400 Or &H800 Or &H40000 Or &H80000 Or &H100000)

    With PD
     .lStructSize = Len(PD)
     .hwndOwner = mhOwner
     .Hdc = GetDC(mhOwner)
     .flags = uFlag
    End With

    On Error GoTo ErrorHandler
    With DM
     .dmDeviceName = Printer.DeviceName
     .dmSize = Len(DM)
     .dmFields = DM_ORIENTATION Or DM_DUPLEX
     .dmOrientation = Printer.Orientation
     On Error Resume Next
     .dmDuplex = Printer.Duplex
     On Error GoTo 0
    End With

    PD.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DM))
    lpDevMode = GlobalLock(PD.hDevMode)
    
    If lpDevMode > 0 Then
     CopyMemory ByVal lpDevMode, DM, Len(DM)
     RetValue = GlobalUnlock(lpDevMode)
    End If

    With DN
     .wDriverOffset = 8
     .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
     .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
     .wDefault = 0
    End With
    
    With Printer
     DN.extra = .DriverName & vbNullChar & .DeviceName & vbNullChar & .Port & vbNullChar
    End With

    PD.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DN))
    lpDevName = GlobalLock(PD.hDevNames)
    
    If lpDevName > 0 Then
     CopyMemory ByVal lpDevName, DN, Len(DN)
     RetValue = GlobalUnlock(lpDevName)
    End If

    RetValue = PrintDlg(PD)
    
    If RetValue = 0 Then
     If mCancelError Then Err.Raise (RetValue)
    Else
     mhOwner = PD.Hdc
     lpDevName = GlobalLock(PD.hDevNames)
     CopyMemory DN, ByVal lpDevName, 45
     RetValue = GlobalUnlock(lpDevName)
     GlobalFree PD.hDevNames
     lpDevMode = GlobalLock(PD.hDevMode)
     CopyMemory DM, ByVal lpDevMode, Len(DM)
     RetValue = GlobalUnlock(PD.hDevMode)
     GlobalFree PD.hDevMode
     NewPrinterName = UCase$(Left(DM.dmDeviceName, InStr(DM.dmDeviceName, vbNullChar) - 1))
    
     If Printer.DeviceName <> NewPrinterName Then
      For Each objPrinter In Printers
       If UCase$(objPrinter.DeviceName) = NewPrinterName Then
         Set Printer = objPrinter
       End If
      Next
     End If
     On Error Resume Next
    
     With Printer
      .Copies = DM.dmCopies
      .Duplex = DM.dmDuplex
      .Orientation = DM.dmOrientation
     End With
     On Error GoTo 0
    End If

ExitSub:
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Printer Error"
    Resume ExitSub

End Sub
