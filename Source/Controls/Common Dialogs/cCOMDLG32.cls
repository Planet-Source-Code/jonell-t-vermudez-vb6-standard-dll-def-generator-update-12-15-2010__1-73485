Attribute VB_Name = "cCOMDLG32"
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'  Persistable = 0  'NotPersistable
'  DataBindingBehavior = 0  'vbNone
'  DataSourceBehavior = 0   'vbNone
'  MTSTransactionMode = 0   'NotAnMTSObject
'End

Option Explicit


'=============================================================================================================
'
' cCOMDLG32 Class Module
' -----------------------
'
' Created By  : Kevin Wilson
'               http://www.TheVBZone.com   ( The VB Zone )
'               http://www.TheVBZone.net   ( The VB Zone .net )
'
' Last Update : June 14, 2000
'
' VB Versions : 5.0 / 6.0
'
' Requires    : modCOMDLG32.bas (Common Dialog module by Kevin Wilson)
'               COMDLG32.DLL    (Microsoft Common Dialog Library)
'               OLEPRO32.DLL    (OLE Automation)
'
' Description : This class module was created as a wrapper for the modCOMDLG32 standard module created by
'               Kevin Wilson.  This class module allows you to use this class module in EXACTLY the same
'               way you'd use the Microsoft Common Dialog ActiveX Control (COMDLG32.OCX).  This class module
'               has/uses the exact same properties and methods as the Microsoft Common Dialog Control, plus
'               a whole lot more for the other dialogs not available in the Microsoft Common Dialog Control.
'
'               See the description in the modCOMDLG32.bas file for more information on the COMDLG32.DLL and
'               the use of these dialogs.
'
'=============================================================================================================
'
' LEGAL:
'
' You are free to use this code as long as you keep the above heading information intact and unchanged. Credit
' given where credit is due.  Also, it is not required, but it would be appreciated if you would mention
' somewhere in your compiled program that that your program makes use of code written and distributed by
' Kevin Wilson (www.TheVBZone.com).  Feel free to link to this code via your web site or articles.
'
' You may NOT take this code and pass it off as your own.  You may NOT distribute this code on your own server
' or web site.  You may NOT take code created by Kevin Wilson (www.TheVBZone.com) and use it to create products,
' utilities, or applications that directly compete with products, utilities, and applications created by Kevin
' Wilson, TheVBZone.com, or Wilson Media.  You may NOT take this code and sell it for profit without first
' obtaining the written consent of the author Kevin Wilson.
'
' These conditions are subject to change at the discretion of the owner Kevin Wilson at any time without
' warning or notice.  Copyright© by Kevin Wilson.  All rights reserved.
'
'=============================================================================================================


Private cd_AppName As String
Private cd_CancelError As Boolean
Private cd_Collate As Boolean
Private cd_Color As Long
Private cd_ColorMode As Long
Private cd_Copies As Long
Private cd_DefaultExt As String
Private cd_DialogTitle As String
Private cd_DriveIndex As Long
Private cd_Duplex As Long
Private cd_FileName As String
Private cd_FileTitle As String
Private cd_Filter As String
Private cd_FindString As String
Private cd_Flags As Long
Private cd_FolderName As String
Private cd_FolderTitle As String
Private cd_FontBold As Boolean
Private cd_FontItalic As Boolean
Private cd_FontName As String
Private cd_FontSize As Long
Private cd_FontStrikethru As Boolean
Private cd_FontUnderline As Boolean
Private cd_FromPage As Long
Private cd_HelpCommand As Long
Private cd_HelpContext As Long
Private cd_HelpFile As String
Private cd_HelpKey As String
Private cd_hIcon As Long
Private cd_IconIndex As Long
Private cd_InitDir As String
Private cd_Margin_Bottom As Long
Private cd_Margin_Left As Long
Private cd_Margin_Right As Long
Private cd_Margin_Top As Long
Private cd_Max As Long
Private cd_Min As Long
Private cd_MinMargin_Bottom As Long
Private cd_MinMargin_Left As Long
Private cd_MinMargin_Right As Long
Private cd_MinMargin_Top As Long
Private cd_Orientation As Long
Private cd_OtherInfo As String
Private cd_OwnerHandle As Long
Private cd_PaperBin As Long
Private cd_PaperSize As Long
Private cd_PaperSize_Height As Long
Private cd_PaperSize_Width As Long
Private cd_PrinterDefault As Boolean
Private cd_PrinterName As String
Private cd_PrintQuality As Long
Private cd_Prompt As String
Private cd_ReplaceString As String
Private cd_RGB_Blue As Byte
Private cd_RGB_Green As Byte
Private cd_RGB_Red As Byte
Private cd_SavedSearch As String
Private cd_ToPage As Long


'=============================================================================================================
'                            CLASS EVENTS
'=============================================================================================================


Private Sub Class_Initialize()
On Error Resume Next
  
  ' Set initial values
  cd_AppName = App.Title
  cd_ColorMode = DMCOLOR_MONOCHROME
  cd_Copies = 1
  cd_Duplex = DMDUP_SIMPLEX
  cd_Filter = "All Files (*.*)|*.*"
  cd_FontName = "Arial"
  cd_FontSize = 10
  cd_FromPage = 1
  cd_HelpCommand = HELP_TAB
  cd_HelpFile = App.HelpFile
  cd_InitDir = CurDir
  cd_Max = 1
  cd_Min = 1
  cd_Orientation = DMORIENT_PORTRAIT
  cd_OwnerHandle = App.hInstance
  cd_PaperSize = DMPAPER_LETTER
  cd_PrinterName = Printer.DeviceName
  cd_PrintQuality = DMRES_HIGH
  cd_ToPage = 1
  
End Sub

Private Sub Class_Terminate()
  
  DoEvents
  
End Sub


'=============================================================================================================
'                            CLASS PROPERTIES
'=============================================================================================================


' Used in : CD_ShowAbout
Public Property Get AppName() As String
  
  AppName = cd_AppName
  
End Property
Public Property Let AppName(ByVal NewValue As String)
  
  cd_AppName = NewValue
  
End Property

' Used in : ALL
Public Property Get CancelError() As Boolean
  
  CancelError = cd_CancelError
  
End Property
Public Property Let CancelError(ByVal NewValue As Boolean)
  
  cd_CancelError = NewValue
  
End Property

' Used in : CD_ShowPrinter
Public Property Get Collate() As Boolean
  
  Collate = cd_Collate
  
End Property
Public Property Let Collate(ByVal NewValue As Boolean)
  
  cd_Collate = NewValue
  
End Property

' Used in : CD_ShowColor, CD_ShowFont
Public Property Get Color() As Long
  
  Color = cd_Color
  
End Property
Public Property Let Color(ByVal NewValue As Long)
  
  cd_Color = NewValue
  
End Property

' Used in : CD_ShowPrinter
Public Property Get ColorMode() As Long
  
  ColorMode = cd_ColorMode
  
End Property
Public Property Let ColorMode(ByVal NewValue As Long)
  
  cd_ColorMode = NewValue
  
End Property

' Used in : CD_ShowPrinter
Public Property Get Copies() As Long
  
  Copies = cd_Copies
  
End Property
Public Property Let Copies(ByVal NewValue As Long)
  
  cd_Copies = NewValue
  
End Property

' Used in : CD_ShowOpen_Save
Public Property Get DefaultExt() As String
  
  DefaultExt = cd_DefaultExt
  
End Property
Public Property Let DefaultExt(ByVal NewValue As String)
  
  cd_DefaultExt = NewValue
  
End Property

' Used in : CD_ShowAbout, CD_ShowOpen_Save, CD_ShowRun
Public Property Get DialogTitle() As String
  
  DialogTitle = cd_DialogTitle
  
End Property
Public Property Let DialogTitle(ByVal NewValue As String)
  
  cd_DialogTitle = NewValue
  
End Property

' Used in : CD_ShowFormat
Public Property Get DriveIndex() As Long
  
  DriveIndex = cd_DriveIndex
  
End Property
Public Property Let DriveIndex(ByVal NewValue As Long)
  
  cd_DriveIndex = NewValue
  
End Property

' Used in : CD_ShowPrinter
Public Property Get Duplex() As Long
  
  Duplex = cd_Duplex
  
End Property
Public Property Let Duplex(ByVal NewValue As Long)
  
  cd_Duplex = NewValue
  
End Property

' Used in : CD_ShowIcon, CD_ShowOpen_Save, CD_ShowProperties
Public Property Get FileName() As String
  
  FileName = cd_FileName
  
End Property
Public Property Let FileName(ByVal NewValue As String)
  
  cd_FileName = NewValue
  
End Property

' Used in : CD_ShowOpen_Save
' > READ-ONLY <
Public Property Get FileTitle() As String
  
  FileTitle = cd_FileTitle
  
End Property

' Used in : CD_ShowOpen_Save
Public Property Get Filter() As String
  
  Filter = cd_Filter
  
End Property
Public Property Let Filter(ByVal NewValue As String)
  
  cd_Filter = NewValue
  
End Property

' Used in : CD_ShowFind, CD_ShowFindReplace
Public Property Get FindString() As String
  
  FindString = cd_FindString
  
End Property
Public Property Let FindString(ByVal NewValue As String)
  
  cd_FindString = NewValue
  
End Property

' Used in : CD_ShowColor       ( CC_...  )
'           CD_ShowFolder      ( BIF_... )
'           CD_ShowFont        ( CF_...  )
'           CD_ShowOpen_Save   ( OFN_... )
'           CD_ShowPageSetup   ( PSD_... )
'           CD_ShowPrinter     ( PD_...  )
'           CD_ShowReboot      ( EWX_... )
'           CD_ShowRun         ( RFF_... )
'           CD_ShowFind        ( FR_...  )
'           CD_ShowFindReplace ( FR_...  )
Public Property Get Flags() As Long
  
  Flags = cd_Flags
  
End Property
Public Property Let Flags(ByVal NewValue As Long)
  
  cd_Flags = NewValue
  
End Property

' Used in : CD_ShowFolder
Public Property Get FolderName() As String
  
  FolderName = cd_FolderName
  
End Property
Public Property Let FolderName(ByVal NewValue As String)
  
  cd_FolderName = NewValue
  
End Property

' Used in : CD_ShowFolder
' > READ-ONLY <
Public Property Get FolderTitle() As String
  
  FolderTitle = cd_FolderTitle
  
End Property

' Used in : CD_ShowFont
Public Property Get FontBold() As Boolean
  
  FontBold = cd_FontBold
  
End Property
Public Property Let FontBold(ByVal NewValue As Boolean)
  
  cd_FontBold = NewValue
  
End Property

' Used in : CD_ShowFont
Public Property Get FontItalic() As Boolean
  
  FontItalic = cd_FontItalic
  
End Property
Public Property Let FontItalic(ByVal NewValue As Boolean)
  
  cd_FontItalic = NewValue
  
End Property

' Used in : CD_ShowFont
Public Property Get fontName() As String
  
  fontName = cd_FontName
  
End Property
Public Property Let fontName(ByVal NewValue As String)
  
  cd_FontName = NewValue
  
End Property

' Used in : CD_ShowFont
Public Property Get FontSize() As Long
  
  FontSize = cd_FontSize
  
End Property
Public Property Let FontSize(ByVal NewValue As Long)
  
  cd_FontSize = NewValue
  
End Property

' Used in : CD_ShowFont
Public Property Get FontStrikethru() As Boolean
  
  FontStrikethru = cd_FontStrikethru
  
End Property
Public Property Let FontStrikethru(ByVal NewValue As Boolean)
  
  cd_FontStrikethru = NewValue
  
End Property

' Used in : CD_ShowFont
Public Property Get FontUnderline() As Boolean
  
  FontUnderline = cd_FontUnderline
  
End Property
Public Property Let FontUnderline(ByVal NewValue As Boolean)
  
  cd_FontUnderline = NewValue
  
End Property

' Used in : CD_ShowPrinter
Public Property Get FromPage() As Long
  
  FromPage = cd_FromPage
  
End Property
Public Property Let FromPage(ByVal NewValue As Long)
  
  cd_FromPage = NewValue
  
End Property

' Used in : CD_ShowHelp
Public Property Get HelpCommand() As Long
  
  HelpCommand = cd_HelpCommand
  
End Property
Public Property Let HelpCommand(ByVal NewValue As Long)
  
  cd_HelpCommand = NewValue
  
End Property

' Used in : CD_ShowHelp
Public Property Get HelpContext() As Long
  
  HelpContext = cd_HelpContext
  
End Property
Public Property Let HelpContext(ByVal NewValue As Long)
  
  cd_HelpContext = NewValue
  
End Property

' Used in : CD_ShowHelp
Public Property Get HelpFile() As String
  
  HelpFile = cd_HelpFile
  
End Property
Public Property Let HelpFile(ByVal NewValue As String)
  
  cd_HelpFile = NewValue
  
End Property

' Used in : CD_ShowHelp
Public Property Get HelpKey() As String
  
  HelpKey = cd_HelpKey
  
End Property
Public Property Let HelpKey(ByVal NewValue As String)
  
  cd_HelpKey = NewValue
  
End Property

' Used in : CD_ShowAbout, CD_ShowRun
Public Property Get hIcon() As Long
  
  hIcon = cd_hIcon
  
End Property
Public Property Let hIcon(ByVal NewValue As Long)
  
  cd_hIcon = NewValue
  
End Property

' Used in : CD_ShowIcon
Public Property Get IconIndex() As Long
  
  IconIndex = cd_IconIndex
  
End Property
Public Property Let IconIndex(ByVal NewValue As Long)
  
  cd_IconIndex = NewValue
  
End Property

' Used in : CD_ShowFindFile, CD_ShowOpen_Save
Public Property Get InitDir() As String
  
  InitDir = cd_InitDir
  
End Property
Public Property Let InitDir(ByVal NewValue As String)
  
  cd_InitDir = NewValue
  
End Property

' Used in : CD_ShowPageSetup
Public Property Get Margin_Bottom() As Long
  
  Margin_Bottom = cd_Margin_Bottom
  
End Property
Public Property Let Margin_Bottom(ByVal NewValue As Long)
  
  cd_Margin_Bottom = NewValue
  
End Property

' Used in : CD_ShowPageSetup
Public Property Get Margin_Left() As Long
  
  Margin_Left = cd_Margin_Left
  
End Property
Public Property Let Margin_Left(ByVal NewValue As Long)
  
  cd_Margin_Left = NewValue
  
End Property

' Used in : CD_ShowPageSetup
Public Property Get Margin_Right() As Long
  
  Margin_Right = cd_Margin_Right
  
End Property
Public Property Let Margin_Right(ByVal NewValue As Long)
  
  cd_Margin_Right = NewValue
  
End Property

' Used in : CD_ShowPageSetup
Public Property Get Margin_Top() As Long
  
  Margin_Top = cd_Margin_Top
  
End Property
Public Property Let Margin_Top(ByVal NewValue As Long)
  
  cd_Margin_Top = NewValue
  
End Property

' Used in : CD_ShowPrinter
Public Property Get Max() As Long
  
  Max = cd_Max
  
End Property
Public Property Let Max(ByVal NewValue As Long)
  
  cd_Max = NewValue
  
End Property

' Used in : CD_ShowPrinter
Public Property Get Min() As Long
  
  Min = cd_Min
  
End Property
Public Property Let Min(ByVal NewValue As Long)
  
  cd_Min = NewValue
  
End Property

' Used in : CD_ShowPageSetup
Public Property Get MinMargin_Bottom() As Long
  
  MinMargin_Bottom = cd_MinMargin_Bottom
  
End Property
Public Property Let MinMargin_Bottom(ByVal NewValue As Long)
  
  cd_MinMargin_Bottom = NewValue
  
End Property

' Used in : CD_ShowPageSetup
Public Property Get MinMargin_Left() As Long
  
  MinMargin_Left = cd_MinMargin_Left
  
End Property
Public Property Let MinMargin_Left(ByVal NewValue As Long)
  
  cd_MinMargin_Left = NewValue
  
End Property

' Used in : CD_ShowPageSetup
Public Property Get MinMargin_Right() As Long
  
  MinMargin_Right = cd_MinMargin_Right
  
End Property
Public Property Let MinMargin_Right(ByVal NewValue As Long)
  
  cd_MinMargin_Right = NewValue
  
End Property

' Used in : CD_ShowPageSetup
Public Property Get MinMargin_Top() As Long
  
  MinMargin_Top = cd_MinMargin_Top
  
End Property
Public Property Let MinMargin_Top(ByVal NewValue As Long)
  
  cd_MinMargin_Top = NewValue
  
End Property

' Used in : CD_ShowPageSetup, CD_ShowPrinter
Public Property Get Orientation() As Long
  
  Orientation = cd_Orientation
  
End Property
Public Property Let Orientation(ByVal NewValue As Long)
  
  cd_Orientation = NewValue
  
End Property

' Used in : CD_ShowAbout
Public Property Get OtherInfo() As String
  
  OtherInfo = cd_OtherInfo
  
End Property
Public Property Let OtherInfo(ByVal NewValue As String)
  
  cd_OtherInfo = NewValue
  
End Property

' Used in : ALL except ShowFindComputer & ShowFindFile
Public Property Get OwnerHandle() As Long
  
  OwnerHandle = cd_OwnerHandle
  
End Property
Public Property Let OwnerHandle(ByVal NewValue As Long)
  
  cd_OwnerHandle = NewValue
  
End Property

' Used in : CD_ShowPrinter
' > READ-ONLY <
Public Property Get PaperBin() As Long
  
  PaperBin = cd_PaperBin
  
End Property

' Used in : CD_ShowPageSetup, CD_ShowPrinter
Public Property Get PaperSize() As Long
  
  PaperSize = cd_PaperSize
  
End Property
Public Property Let PaperSize(ByVal NewValue As Long)
  
  cd_PaperSize = NewValue
  
End Property

' Used in : CD_ShowPageSetup
' > READ-ONLY <
Public Property Get PaperSize_Height() As Long
  
  PaperSize_Height = cd_PaperSize_Height
  
End Property

' Used in : CD_ShowPageSetup
' > READ-ONLY <
Public Property Get PaperSize_Width() As Long
  
  PaperSize_Width = cd_PaperSize_Width
  
End Property

' Used in : CD_ShowPrinter
Public Property Get PrinterDefault() As Boolean
  
  PrinterDefault = cd_PrinterDefault
  
End Property
Public Property Let PrinterDefault(ByVal NewValue As Boolean)
  
  cd_PrinterDefault = NewValue
  
End Property

' Used in : CD_ShowPrinter, CD_ShowProperties
Public Property Get PrinterName() As String
  
  PrinterName = cd_PrinterName
  
End Property
Public Property Let PrinterName(ByVal NewValue As String)
  
  cd_PrinterName = NewValue
  
End Property

' Used in : CD_ShowPrinter
Public Property Get PrintQuality() As Long
  
  PrintQuality = cd_PrintQuality
  
End Property
Public Property Let PrintQuality(ByVal NewValue As Long)
  
  cd_PrintQuality = NewValue
  
End Property

' Used in : CD_ShowFolder, CD_ShowReboot, CD_ShowRun
Public Property Get Prompt() As String
  
  Prompt = cd_Prompt
  
End Property
Public Property Let Prompt(ByVal NewValue As String)
  
  cd_Prompt = NewValue
  
End Property

' Used in : CD_ShowFindReplace
Public Property Get ReplaceString() As String
  
  ReplaceString = cd_ReplaceString
  
End Property
Public Property Let ReplaceString(ByVal NewValue As String)
  
  cd_ReplaceString = NewValue
  
End Property

' Used in : CD_ShowColor
' > READ-ONLY <
Public Property Get RGB_Blue() As Long
  
  RGB_Blue = cd_RGB_Blue
  
End Property

' Used in : CD_ShowColor
' > READ-ONLY <
Public Property Get RGB_Green() As Long
  
  RGB_Green = cd_RGB_Green
  
End Property

' Used in : CD_ShowColor
' > READ-ONLY <
Public Property Get RGB_Red() As Long
  
  RGB_Red = cd_RGB_Red
  
End Property

' Used in : CD_ShowFindFile
Public Property Get SavedSearch() As String
  
  SavedSearch = cd_SavedSearch
  
End Property
Public Property Let SavedSearch(ByVal NewValue As String)
  
  cd_SavedSearch = NewValue
  
End Property

' Used in : CD_ShowPrinter
Public Property Get ToPage() As Long
  
  ToPage = cd_ToPage
  
End Property
Public Property Let ToPage(ByVal NewValue As Long)
  
  cd_ToPage = NewValue
  
End Property



'=============================================================================================================
'                            CLASS METHODS
'=============================================================================================================


Public Function ShowAbout() As Boolean
  
  If CD_ShowAbout(cd_OwnerHandle, cd_DialogTitle, cd_AppName, cd_OtherInfo, cd_hIcon) = False Then
    ShowAbout = False
    If cd_CancelError = True Then
      Err.Raise CDERR_CANCEL, "basCOMDLG32.bas - CD_ShowAbout", CDERR_CANCELMSG
    End If
  Else
    ShowAbout = True
  End If
  
End Function

Public Function ShowColor() As Boolean
  
  Dim TheRed As Byte
  Dim TheGreen As Byte
  Dim TheBlue As Byte
  Dim TheColor As Long
  Dim TheFlags As Long
  
  TheColor = cd_Color
  TheFlags = cd_Flags
  
  If CD_ShowColor(cd_OwnerHandle, TheFlags, TheColor, TheRed, TheGreen, TheBlue) = False Then
    ShowColor = False
    If cd_CancelError = True Then
      Err.Raise CDERR_CANCEL, "basCOMDLG32.bas - CD_ShowColor", CDERR_CANCELMSG
    End If
  Else
    ShowColor = True
    cd_Color = TheColor
    cd_Flags = TheFlags
    cd_RGB_Red = TheRed
    cd_RGB_Green = TheGreen
    cd_RGB_Blue = TheBlue
  End If
  
End Function

' NOTE - Once this is called, make sure to put a call to your custom function
'        that handles the Find dialog's events in the FindReplace_Event()
'        function in the modCOMDLG32.bas file
Public Function ShowFind() As Boolean
  
  CD_ShowFindReplace cd_OwnerHandle, cd_Flags, cd_FindString, False
  
End Function

Public Function ShowFindComputer() As Boolean
  
  If CD_ShowFindComputer = False Then
    ShowFindComputer = False
    If cd_CancelError = True Then
      Err.Raise CDERR_CANCEL, "basCOMDLG32.bas - CD_ShowFindComputer", CDERR_CANCELMSG
    End If
  Else
    ShowFindComputer = True
  End If
  
End Function

Public Function ShowFindFile() As Boolean
  
  If CD_ShowFindFile(cd_InitDir, cd_SavedSearch) = False Then
    ShowFindFile = False
    If cd_CancelError = True Then
      Err.Raise CDERR_CANCEL, "basCOMDLG32.bas - CD_ShowFindFile", CDERR_CANCELMSG
    End If
  Else
    ShowFindFile = True
  End If
  
End Function

' NOTE - Once this is called, make sure to put a call to your custom function
'        that handles the Find dialog's events in the FindReplace_Event()
'        function in the modCOMDLG32.bas file
Public Function ShowFindReplace() As Boolean
  
  CD_ShowFindReplace cd_OwnerHandle, cd_Flags, cd_FindString, True, cd_ReplaceString
  
End Function

Public Function ShowFolder() As Boolean
  
  Dim TheFlags As Long
  Dim TheFolderName As String
  Dim TheFolderTitle As String
  
  TheFlags = cd_Flags
  TheFolderName = cd_FolderName
  TheFolderTitle = cd_FolderTitle
  
  If CD_ShowFolder(cd_OwnerHandle, TheFlags, TheFolderName, TheFolderTitle, cd_Prompt) = False Then
    ShowFolder = False
    If cd_CancelError = True Then
      Err.Raise CDERR_CANCEL, "basCOMDLG32.bas - CD_ShowFolder", CDERR_CANCELMSG
    End If
  Else
    ShowFolder = True
    cd_Flags = TheFlags
    cd_FolderName = TheFolderName
    cd_FolderTitle = TheFolderTitle
  End If
  
End Function

Public Function ShowFont() As Boolean ' Memory Leak
  
  Dim TheFlags As Long
  Dim TheFontName As String
  Dim TheFontSize As Long
  Dim TheFontBold As Boolean
  Dim TheFontItalic As Boolean
  Dim TheFontStrikethru As Boolean
  Dim TheFontUnderline As Boolean
  Dim TheColor As Long
  
  TheFlags = cd_Flags
  TheFontName = cd_FontName
  TheFontSize = cd_FontSize
  TheFontBold = cd_FontBold
  TheFontItalic = cd_FontItalic
  TheFontStrikethru = cd_FontStrikethru
  TheFontUnderline = cd_FontUnderline
  TheColor = cd_Color
  
  If CD_ShowFont(cd_OwnerHandle, TheFlags, TheFontName, TheFontSize, TheFontBold, TheFontItalic, TheFontStrikethru, TheFontUnderline, TheColor) = False Then
    ShowFont = False
    If cd_CancelError = True Then
      Err.Raise CDERR_CANCEL, "basCOMDLG32.bas - CD_ShowFont", CDERR_CANCELMSG
    End If
  Else
    ShowFont = True
    cd_Flags = TheFlags
    cd_FontName = TheFontName
    cd_FontSize = TheFontSize
    cd_FontBold = TheFontBold
    cd_FontItalic = TheFontItalic
    cd_FontStrikethru = TheFontStrikethru
    cd_FontUnderline = TheFontUnderline
    cd_Color = TheColor
  End If
  
End Function

Public Function ShowFormat() As Boolean
  
  If CD_ShowFormat(cd_OwnerHandle, cd_DriveIndex) = False Then
    ShowFormat = False
    If cd_CancelError = True Then
      Err.Raise CDERR_CANCEL, "basCOMDLG32.bas - CD_ShowFormat", CDERR_CANCELMSG
    End If
  Else
    ShowFormat = True
  End If
  
End Function

Public Function ShowHelp() As Boolean
  
  If CD_ShowHelp(cd_OwnerHandle, cd_HelpFile, ConvertHelpConst(cd_HelpCommand), cd_HelpContext, cd_HelpKey) = False Then
    ShowHelp = False
    If cd_CancelError = True Then
      Err.Raise CDERR_CANCEL, "basCOMDLG32.bas - CD_ShowHelp", CDERR_CANCELMSG
    End If
  Else
    ShowHelp = True
  End If
  
End Function

Public Function ShowIcon() As Boolean
  
  Dim TheFileName As String
  Dim TheIconIndex As Long
  
  TheFileName = cd_FileName
  TheIconIndex = cd_IconIndex
  
  If CD_ShowIcon(cd_OwnerHandle, TheFileName, TheIconIndex) = False Then
    ShowIcon = False
    If cd_CancelError = True Then
      Err.Raise CDERR_CANCEL, "basCOMDLG32.bas - CD_ShowIcon", CDERR_CANCELMSG
    End If
  Else
    ShowIcon = True
    cd_FileName = TheFileName
    cd_IconIndex = TheIconIndex
  End If
  
End Function

Public Function ShowOpen() As Boolean
  
  Dim TheFlags As Long
  Dim TheFileName As String
  Dim TheFileTitle As String
  Dim TheFilter As String
  
  TheFlags = cd_Flags
  TheFileName = cd_FileName
  TheFileTitle = cd_FileTitle
  TheFilter = cd_Filter
  
  If CD_ShowOpen_Save(cd_OwnerHandle, TheFlags, TheFileName, TheFileTitle, cd_DefaultExt, cd_DialogTitle, TheFilter, cd_InitDir, True) = False Then
    ShowOpen = False
    If cd_CancelError = True Then
      Err.Raise CDERR_CANCEL, "basCOMDLG32.bas - CD_ShowOpen_Save", CDERR_CANCELMSG
    End If
  Else
    ShowOpen = True
    cd_Flags = TheFlags
    cd_FileName = TheFileName
    cd_FileTitle = TheFileTitle
    cd_Filter = TheFilter
  End If
  
End Function

Public Function ShowPageSetup() As Boolean ' Memory Leak
  
  Dim TheFlags As Long
  Dim TheOrientation As Long
  Dim ThePaperSize As Long
  Dim ThePaperSize_Height As Long
  Dim ThePaperSize_Width As Long
  Dim TheMargin_Left As Long
  Dim TheMargin_Top As Long
  Dim TheMargin_Right As Long
  Dim TheMargin_Bottom As Long
  
  TheFlags = cd_Flags
  TheOrientation = cd_Orientation
  ThePaperSize = cd_PaperSize
  ThePaperSize_Height = 0
  ThePaperSize_Width = 0
  TheMargin_Left = cd_Margin_Left
  TheMargin_Top = cd_Margin_Top
  TheMargin_Right = cd_Margin_Right
  TheMargin_Bottom = cd_Margin_Bottom
  
  If CD_ShowPageSetup(cd_OwnerHandle, TheFlags, TheOrientation, ThePaperSize, ThePaperSize_Height, ThePaperSize_Width, TheMargin_Left, TheMargin_Top, TheMargin_Right, TheMargin_Bottom, cd_MinMargin_Left, cd_MinMargin_Top, cd_MinMargin_Right, cd_MinMargin_Bottom) = False Then
    ShowPageSetup = False
    If cd_CancelError = True Then
      Err.Raise CDERR_CANCEL, "basCOMDLG32.bas - CD_ShowPageSetup", CDERR_CANCELMSG
    End If
  Else
    ShowPageSetup = True
    cd_Flags = TheFlags
    cd_Orientation = TheOrientation
    cd_PaperSize = ThePaperSize
    cd_PaperSize_Height = ThePaperSize_Height
    cd_PaperSize_Width = ThePaperSize_Width
    cd_Margin_Left = TheMargin_Left
    cd_Margin_Top = TheMargin_Top
    cd_Margin_Right = TheMargin_Right
    cd_Margin_Bottom = TheMargin_Bottom
  End If
  
End Function

Public Function ShowPrinter() As Boolean ' Memory Leak
  
  Dim TheFlags As Long
  Dim ThePrinterName As String
  Dim TheFromPage As Long
  Dim TheToPage As Long
  Dim TheMin As Long
  Dim TheMax As Long
  Dim TheCopies As Long
  Dim TheDuplex As Long
  Dim TheOrientation As Long
  Dim ThePaperSize As Long
  Dim ThePrintQuality As Long
  Dim TheColorMode As Long
  Dim ThePaperBin As Long
  Dim TheCollate As Boolean
  
  TheFlags = cd_Flags
  ThePrinterName = ""
  TheFromPage = cd_FromPage
  TheToPage = cd_ToPage
  TheMin = cd_Min
  TheMax = cd_Max
  TheCopies = cd_Copies
  TheDuplex = 0
  TheOrientation = cd_Orientation
  ThePaperSize = cd_PaperSize
  ThePrintQuality = cd_PrintQuality
  TheColorMode = cd_ColorMode
  ThePaperBin = 0
  TheCollate = cd_Collate
  
  If CD_ShowPrinter(cd_OwnerHandle, TheFlags, ThePrinterName, CInt(TheFromPage), CInt(TheToPage), CInt(TheMin), CInt(TheMax), CInt(TheCopies), CInt(TheDuplex), CInt(TheOrientation), CInt(ThePaperSize), CInt(ThePrintQuality), CInt(TheColorMode), CInt(ThePaperBin), TheCollate, cd_PrinterDefault) = False Then
    ShowPrinter = False
    If cd_CancelError = True Then
      Err.Raise CDERR_CANCEL, "basCOMDLG32.bas - CD_ShowPageSetup", CDERR_CANCELMSG
    End If
  Else
    ShowPrinter = True
    cd_Flags = TheFlags
    cd_PrinterName = ThePrinterName
    cd_FromPage = CLng(TheFromPage)
    cd_ToPage = CLng(TheToPage)
    cd_Min = CLng(TheMin)
    cd_Max = CLng(TheMax)
    cd_Copies = CLng(TheCopies)
    cd_Duplex = CLng(TheDuplex)
    cd_Orientation = CLng(TheOrientation)
    cd_PaperSize = CLng(ThePaperSize)
    cd_PrintQuality = CLng(ThePrintQuality)
    cd_ColorMode = CLng(TheColorMode)
    cd_PaperBin = CLng(ThePaperBin)
    cd_Collate = TheCollate
  End If
  
End Function

Public Function ShowProperties() As Boolean
  
  If CD_ShowProperties(cd_OwnerHandle, cd_FileName, cd_PrinterName) = False Then
    ShowProperties = False
    If cd_CancelError = True Then
      Err.Raise CDERR_CANCEL, "basCOMDLG32.bas - CD_ShowProperties", CDERR_CANCELMSG
    End If
  Else
    ShowProperties = True
  End If
  
End Function

Public Function ShowReboot() As Boolean
  
  If CD_ShowReboot(cd_OwnerHandle, cd_Flags, cd_Prompt) = False Then
    ShowReboot = False
    If cd_CancelError = True Then
      Err.Raise CDERR_CANCEL, "basCOMDLG32.bas - CD_ShowReboot", CDERR_CANCELMSG
    End If
  Else
    ShowReboot = True
  End If
  
End Function

Public Function ShowRun() As Boolean
  
  If CD_ShowRun(cd_OwnerHandle, cd_Flags, cd_Prompt, cd_DialogTitle, cd_hIcon) = False Then
    ShowRun = False
    If cd_CancelError = True Then
      Err.Raise CDERR_CANCEL, "basCOMDLG32.bas - CD_ShowRun", CDERR_CANCELMSG
    End If
  Else
    ShowRun = True
  End If
  
End Function

Public Function ShowSave() As Boolean
  
  Dim TheFlags As Long
  Dim TheFileName As String
  Dim TheFileTitle As String
  Dim TheDefaultExt As String
  Dim TheDialogTitle As String
  Dim TheFilter As String
  Dim TheInitDir As String
  
  TheFlags = cd_Flags
  TheFileName = cd_FileName
  TheFileTitle = cd_FileTitle
  TheDefaultExt = cd_DefaultExt
  TheDialogTitle = cd_DialogTitle
  TheFilter = cd_Filter
  TheInitDir = cd_InitDir
  
  If CD_ShowOpen_Save(cd_OwnerHandle, TheFlags, TheFileName, TheFileTitle, TheDefaultExt, TheDialogTitle, TheFilter, TheInitDir, False) = False Then
    ShowSave = False
    If cd_CancelError = True Then
      Err.Raise CDERR_CANCEL, "basCOMDLG32.bas - CD_ShowOpen_Save", CDERR_CANCELMSG
    End If
  Else
    ShowSave = True
    cd_Flags = TheFlags
    cd_FileName = TheFileName
    cd_FileTitle = TheFileTitle
    cd_DefaultExt = TheDefaultExt
    cd_DialogTitle = TheDialogTitle
    cd_Filter = TheFilter
    cd_InitDir = TheInitDir
  End If
  
End Function

Public Function ShowShutDown() As Boolean
  
  If CD_ShowShutDown(cd_OwnerHandle) = False Then
    ShowShutDown = False
    If cd_CancelError = True Then
      Err.Raise CDERR_CANCEL, "basCOMDLG32.bas - CD_ShowShutDown", CDERR_CANCELMSG
    End If
  Else
    ShowShutDown = True
  End If
  
End Function


'=============================================================================================================
'=============================================================================================================

Private Function ConvertHelpConst(ByVal HelpConst As Long) As HelpCommands
On Error Resume Next
  
  Select Case HelpConst
    Case &H102&  ' Run the specified help macro
      ConvertHelpConst = HELP_COMMAND
    Case &H1     ' Displays specified help
      ConvertHelpConst = HELP_CONTEXT
    Case &H8&    ' Displays specified help in a Pop-Up window
      ConvertHelpConst = HELP_CONTEXTPOPUP
    Case &H9&    ' Display the first page of the help system
      ConvertHelpConst = HELP_FORCEFILE
    Case &H4     ' Displays help on how to use the WinHelp system
      ConvertHelpConst = HELP_HELPONHELP
    Case &H2     ' Exits all WinHelp help systems
      ConvertHelpConst = HELP_QUIT
    Case &H101   ' Opens the help file to the "Index" tab
      ConvertHelpConst = HELP_KEY
    Case &HF     ' Opens the help file to the "Contents" tab
      ConvertHelpConst = HELP_TAB
  End Select
  
End Function
