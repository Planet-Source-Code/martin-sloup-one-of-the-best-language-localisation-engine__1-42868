VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*****************************************
'
' Loading Resources from Resource-Only DLL
'
'  Copyright 2000 by Zvonko Bostjancic
'  E-mail: zvonko.bostjancic@siol.net
'
'  This class (or DLL) is used to load resources
'  from resource-only DLL. You can load common
'  picture formats (BMP, ICO, CUR, JPEG, GIF)
'  directly to VB (through picture object). From
'  other resources you can load strings, HTML (as
'  string or to file) and play WAVEs. All other
'  resources can be saved to file on disk and loaded
'  then (including those that can be loaded directly
'  to VB.
'
'  It's not very good documented, but I thing it'll
'  do it for now.
'
'  Goals for next versions:
'    - Playing AVIs directly from resources
'
'  If you have any comments or requests, please
'  contact me on my e-mail address.
'
' Updated by Martin Sloup (01.02.2003)
'  * Added Alfanumeric key for resources
'
'****************************************

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As Guid, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Declare Function FreeResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As Any) As Long
Private Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Const SND_RESOURCE = &H40004 ' name is resource name or atom
Private Const SND_SYNC = &H0 ' play synchronously

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

Private Const IMAGE_BITMAP = 0
Private Const IMAGE_ICON = 1
Private Const IMAGE_CURSOR = 2

Private Const RT_CURSOR = 1&
Private Const RT_BITMAP = 2&
Private Const RT_ICON = 3&
Private Const RT_MENU = 4&
Private Const RT_DIALOG = 5&
Private Const RT_STRING = 6&
Private Const RT_FONTDIR = 7&
Private Const RT_FONT = 8&
Private Const RT_ACCELERATOR = 9&
Private Const RT_RCDATA = 10&
Private Const RT_MESSAGETABLE = 11&
Private Const RT_GROUP_CURSOR = RT_CURSOR + 11
Private Const RT_GROUP_ICON = RT_ICON + 11
Private Const RT_VERSION = 16&
Private Const RT_DLGINCLUDE = 17&
Private Const RT_PLUGPLAY = 19&
Private Const RT_VXD = 20&
Private Const RT_ANICURSOR = 21&
Private Const RT_ANIICON = 22&
Private Const RT_HTML = 23&

Public Enum CRStandardResourceTypeConstants
   crCursor = RT_CURSOR
   crBitmap = RT_BITMAP
   crIcon = RT_ICON
   crMenu = RT_MENU
   crDialog = RT_DIALOG
   crString = RT_STRING
   crFontDir = RT_FONTDIR
   crFont = RT_FONT
   crAccelerator = RT_ACCELERATOR
   crRCData = RT_RCDATA
   crMessageTable = RT_MESSAGETABLE
   crGroupCursor = RT_GROUP_CURSOR
   crGroupIcon = RT_GROUP_ICON
   crVersion = RT_VERSION
   crDlgInclude = RT_DLGINCLUDE
   crPlugPlay = RT_PLUGPLAY
   crVXD = RT_VXD
   crAniCursor = RT_ANICURSOR
   crAniIcon = RT_ANIICON
   crHTML = RT_HTML
End Enum

Private Type PictDesc
    cbSizeofStruct As Long
    PicType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type PicBmp
   Size As Long
   tType As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type

Private Type Guid
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Enum PicType
    resIcon = 1
    resBitmap = 2
    resCursor = 3
    resJPEG = 4
    resGIF = 5
End Enum

Private mvarDllName As String 'local copy

'Set name of resource-only DLL
Public Property Let DllName(ByVal vData As String)
Attribute DllName.VB_Description = "Gets/sets name of resource DLL."
    mvarDllName = vData
End Property

'Get name of resource-only DLL
Public Property Get DllName() As String
    DllName = mvarDllName
End Property

'This function loads string(s) from resource-only DLL
'You have to specify ID of resource.
Function LoadStringFromDLL(lngID As Long) As String
Attribute LoadStringFromDLL.VB_Description = "Loads a string from DLL."
    Dim hInst As Long
    Dim str As String * 64000
    
    'Check if DLLName is valid
    If CheckDllName = False Then Exit Function
    
    hInst = LoadLibrary(DllName) 'load DLL
    If hInst = 0 Then
        Err.Raise 27001, App.Title & ".LoadRes", WinError(Err.LastDllError)
        Exit Function
    End If
    
    Call LoadString(hInst, lngID, str, Len(str)) 'let's load string
    
    LoadStringFromDLL = Trim(str) 'set value of string to function

    FreeLibrary (hInst) 'close library
End Function

'This function loads picture from resource-only DLL
'You have to specify ID and type of picture
Function LoadPictureFromDLL(ID, Pic As PicType) As IPictureDisp
Attribute LoadPictureFromDLL.VB_Description = "Loads pictures (bitmaps, icons, corsors, jpegs and gifs) from resource-only DLL when they're using usal resource types."
    Dim sFile As String
    Dim sType As String
    Dim hInst As Long
    Dim bLoad As Long
    Dim oNewPic As Picture
    Dim tPicConv As PicBmp
    Dim IGuid As Guid
    Dim iLoad As Long
    Dim cLoad As Long
    Dim oNewIcon As Picture
    Dim tpic As PictDesc
    
    If CheckDllName = False Then Exit Function
    
    If IsNumeric(ID) Then ID = "#" & ID
    
    On Error GoTo Nnapaka
    
    hInst = LoadLibrary(DllName)
    If hInst = 0 Then
        Err.Raise 27001, App.Title & ".LoadRes", "Resource DLL could not be loaded."
        Exit Function
    End If
    
    Select Case Pic
        Case resIcon
            GoTo IconStart
        Case resBitmap
            GoTo BitmapStart
        Case resCursor
            GoTo CursorStart
        Case resJPEG
            GoTo JPEGStart
        Case resGIF
            GoTo GIFStart
    End Select
    Exit Function
    
BitmapStart:
    
    'load bitmap
    bLoad = LoadImage(hInst, ID, IMAGE_BITMAP, 0, 0, 0)
    If bLoad = 0 Then
        Err.Raise 27001, App.Title & ".LoadRes", WinError(Err.LastDllError)
        FreeLibrary (hInst)
        Exit Function
    End If

    With tPicConv 'set picture type
        .tType = vbPicTypeBitmap
        .Size = Len(tPicConv)
        .hBmp = bLoad
    End With
    
    With IGuid 'set bitmap properties
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    
    'create bitmap object
    OleCreatePictureIndirect tPicConv, IGuid, True, oNewPic
    
    Set LoadPictureFromDLL = oNewPic 'set function to bitmap object
    FreeLibrary (hInst) 'close library
    Exit Function


IconStart:
    
    'load icon
    iLoad = LoadImage(hInst, ID, IMAGE_ICON, 0, 0, 0)
    If iLoad = 0 Then
        Err.Raise 27001, App.Title & ".LoadRes", WinError(Err.LastDllError)
        FreeLibrary (hInst)
        Exit Function
    End If

    With tPicConv 'set picture type
        .Size = Len(tPicConv)
        .hBmp = iLoad
        .tType = vbPicTypeIcon
    End With
    
    With IGuid 'set icon properties
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    'create icon object
    OleCreatePictureIndirect tPicConv, IGuid, True, oNewIcon
    
    Set LoadPictureFromDLL = oNewIcon 'set object to function
    FreeLibrary (hInst) 'close library
    Exit Function

CursorStart:
    
    'load cursor
    cLoad = LoadImage(hInst, ID, IMAGE_CURSOR, 0, 0, 0)
    If cLoad = 0 Then
        Err.Raise 27001, App.Title & ".LoadRes", WinError(Err.LastDllError)
        FreeLibrary (hInst)
        Exit Function
    End If

    With tPicConv 'set cursor properties (just like icon)
        .Size = Len(tPicConv)
        .hBmp = cLoad
        .tType = vbPicTypeIcon
    End With
    
    With IGuid 'set cursor properties (same as icon)
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    'create cursor object
    OleCreatePictureIndirect tPicConv, IGuid, True, oNewIcon
    
    Set LoadPictureFromDLL = oNewIcon 'set cursor object to function
    FreeLibrary (hInst) 'close library
    Exit Function

JPEGStart:

    sFile = Trim(GetTempFile()) 'let's get free temp file name
    
    On Error GoTo Nnapaka
    
    FreeLibrary (hInst) 'we don't need library instance
                        'here so we close it
    
    'save resource to file
    If SaveAnyResource(ID, "JPEG", sFile) = False Then
        Err.Raise 27008, App.EXEName & ":LoadRes.LoadPictureFromDLL", WinError(Err.LastDllError)
        Exit Function
    End If

    'load JPEG from temporary file
    Set LoadPictureFromDLL = LoadPicture(sFile)
    
    Kill sFile 'delete temp file
    sFile = Empty
    Exit Function

GIFStart:

    sFile = Trim(GetTempFile()) 'let's get free temp file
    
    On Error GoTo Nnapaka
    
    FreeLibrary (hInst) 'we don't need library instance
                        'here so we close it
    
    'save resource to file
    If SaveAnyResource(ID, "GIF", sFile) = False Then
        Err.Raise 27008, App.EXEName & ".LoadRes", WinError(Err.LastDllError)
        Exit Function
    End If

    'load picture from temp file
    Set LoadPictureFromDLL = LoadPicture(sFile)
    
    Kill sFile 'delete temp file
    sFile = Empty
    Exit Function

Nnapaka:
    Err.Raise 27999, App.Title & ".LoadRes", WinError(Err.LastDllError)
    If hInst <> 0 Then FreeLibrary (hInst)
    Exit Function
End Function

'this function returns free temp file
'Prefix = first letter of a filename
Private Function GetTempFile(Optional Prefix = "&")
    Dim lngReturnVal As Long
    Dim strTempPath As String * 255
    Dim strTempFileName As String * 255
    
    lngReturnVal = GetTempPath(254, strTempPath)
    lngReturnVal = GetTempFileName(strTempPath & "\", Prefix, 0, strTempFileName)
    
    GetTempFile = Replace(strTempFileName, ".tmp", ".jpg", , , vbTextCompare)
End Function

Function LoadHtmlFromDLL(ID, Optional resType As String = RT_HTML) As String
Attribute LoadHtmlFromDLL.VB_Description = "Loads HTML from resource-only DLL."
    Dim hInst As Long
    Dim hFind As Long
    Dim hLoad As Long
    Dim hLock As Long
    Dim hSize As Long
    Dim strTemp As String
    Dim iFile As Integer

    If CheckDllName = False Then Exit Function
        
    On Error GoTo Nnapaka
    
    sName = ID
    If IsNumeric(ID) Then sName = "#" & CStr(ID)
    
    'load library
    hInst = LoadLibrary(DllName)
    If hInst = 0 Then
        Err.Raise 27001, LoadRes.LoadDLLPicture, WinError(Err.LastDllError)
        FreeLibrary (hInst)
        Exit Function
    End If
    
    'first we have to find resource
    If IsNumeric(resType) Then
        hFind = FindResource(hInst, ByVal sName, ByVal CLng(resType))
    Else
        hFind = FindResource(hInst, ByVal sName, ByVal resType)
    End If
    If hFind = 0 Then
        Err.Raise 27001, App.EXEName & ".LoadRes", "The specified resource could not be found."
        FreeLibrary (hInst)
        Exit Function
    End If
    
    'then we have to load it
    hLoad = LoadResource(hInst, hFind)
    If hLoad = 0 Then
        Err.Raise 27002, App.EXEName & ".LoadRes", WinError(Err.LastDllError)
        FreeLibrary (hInst)
        Exit Function
    End If
    
    'and lock it for access
    hLock = LockResource(hLoad)
    If hLock = 0 Then
        Err.Raise 27003, App.EXEName & ".LoadRes", WinError(Err.LastDllError)
        FreeLibrary (hInst)
        Exit Function
    End If
    
    'we have to determine size of resource
    hSize = SizeofResource(hInst, hFind)
    If hSize = 0 Then
        Err.Raise 27004, App.EXEName & ".LoadRes", WinError(Err.LastDllError)
        FreeLibrary (hInst)
        Exit Function
    End If
    
    'and then load it with help of CopyMemory function
    strTemp = String$(hSize + 1, Chr(0))
    CopyMemory ByVal strTemp, ByVal hLock, hSize
    
    FreeLibrary (hInst) 'we close library
    LoadHtmlFromDLL = strTemp 'and set function value

    Exit Function
    
Nnapaka:
    Err.Raise Err.Number, App.EXEName & ".LoadRes", Err.Description
    If Not (iFile = 0) Then
        Close #iFile
    End If
    If hInst <> 0 Then FreeLibrary (hInst)
    Exit Function
End Function

'Returns error description of an windows error that
'can't be returned through VB (loading library, resources...)
Private Function WinError(ByVal lLastDLLError As Long) As String
    Dim sBuff As String
    Dim lCount As Long
    
    sBuff = String$(256, 0)
    lCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
    If lCount Then
        WinError = Left$(sBuff, lCount)
    End If
End Function

'saves any resource to file on disk
'most of it is already explained in LoadHtmlFromDLL function
Public Function SaveAnyResource(ID, resType As String, strFileName As String) As Boolean
Attribute SaveAnyResource.VB_Description = "Saves any resource to with specified ID and type to file on hard disk."
    Dim hInst As Long
    Dim hFind As Long
    Dim hLoad As Long
    Dim hLock As Long
    Dim hSize As Long
    Dim sName As String
    Dim b() As Byte
    Dim iFile As Integer

    SaveAnyResource = False

    If CheckDllName = False Then Exit Function
    
    On Error GoTo Nnapaka
    
    sName = ID
    If IsNumeric(ID) Then sName = "#" & ID
    
    hInst = LoadLibrary(DllName)
    If hInst = 0 Then
        Err.Raise 27006, App.EXEName & ".LoadRes", WinError(Err.LastDllError)
        Exit Function
    End If
    
    hFind = FindResource(hInst, ByVal sName, ByVal resType)
    If hFind = 0 Then
        Err.Raise 27001, App.EXEName & ".LoadRes", "The specified resource could not be found."
        FreeLibrary (hInst)
        Exit Function
    End If
    
    hLoad = LoadResource(hInst, hFind)
    If hLoad = 0 Then
        Err.Raise 27002, App.EXEName & ".LoadRes", WinError(Err.LastDllError)
        FreeLibrary (hInst)
        Exit Function
    End If
    
    hLock = LockResource(hLoad)
    If hLock = 0 Then
        Err.Raise 27003, App.EXEName & ".LoadRes", WinError(Err.LastDllError)
        FreeLibrary (hInst)
        Exit Function
    End If
    
    hSize = SizeofResource(hInst, hFind)
    If hSize = 0 Then
        Err.Raise 27004, App.EXEName & ".LoadRes", WinError(Err.LastDllError)
        FreeLibrary (hInst)
        Exit Function
    End If
    
    'we redim variable
    ReDim b(0 To hSize) As Byte
    'and copy memory from resources to our variable
    CopyMemory b(0), ByVal hLock, hSize
    
    On Error Resume Next
    Kill strFileName 'if file exists we delete it
    On Error GoTo Nnapaka
    
    'then we create file with contenst of our resource
    'loaded from resource file
    iFile = FreeFile
    Open strFileName For Binary Access Write Lock Read As iFile
        Put #iFile, , b
    Close #iFile
    iFile = 0
    
    FreeLibrary hInst 'close library
    SaveAnyResource = True 'save was successful
    Exit Function
    
Nnapaka:
    Err.Raise Err.Number, App.EXEName & ".LoadRes", Err.Description
    If Not (iFile = 0) Then
        Close #iFile
    End If
    If hInst <> 0 Then FreeLibrary (hInst)
    SaveAnyResource = False
    Exit Function
End Function

'Here we check for DLL (if it's entered and if it's valid)
Private Function CheckDllName() As Boolean
    CheckDllName = True
    
    If Len(DllName) = 0 Or DllName = vbNullString Then
        Err.Raise 27000, App.EXEName & ".LoadRes", "Ime datoteke ni vpisano!"
        CheckDllName = False
        Exit Function
    End If
    If Dir(DllName) = "" Then
        Err.Raise 27005, App.EXEName & ".LoadRes", "Datoteka " & vbNewLine & DllName & vbNewLine & " ne obstaja!"
        CheckDllName = False
        Exit Function
    End If
End Function

'this function saves HTML resources to file on disk
'most of it has already been explaind before
Function SaveHtml(ID, strFileName As String, Optional resType As String = "RT_HTML") As String
Attribute SaveHtml.VB_Description = "Saves HTML from resource-only DLL to a file on hard disk."
    Dim hInst As Long
    Dim hFind As Long
    Dim hLoad As Long
    Dim hLock As Long
    Dim hSize As Long
    Dim sName As String
    Dim b() As Byte
    Dim iFile As Integer

    If CheckDllName = False Then Exit Function
        
    On Error GoTo Nnapaka
    
    sName = ID
    If IsNumeric(ID) Then sName = "#" & ID
    
    hInst = LoadLibrary(DllName)
    If hInst = 0 Then
        Err.Raise 27006, App.EXEName & ".LoadRes", WinError(Err.LastDllError)
        Exit Function
    End If
    
    hFind = FindResource(hInst, ByVal sName, ByVal resType)
    If hFind = 0 Then
        Err.Raise 27001, App.EXEName & ".LoadRes", "The specified resource could not be found."
        FreeLibrary (hInst)
        Exit Function
    End If
    
    hLoad = LoadResource(hInst, hFind)
    If hLoad = 0 Then
        Err.Raise 27002, App.EXEName & ".LoadRes", WinError(Err.LastDllError)
        FreeLibrary (hInst)
        Exit Function
    End If
    
    hLock = LockResource(hLoad)
    If hLock = 0 Then
        Err.Raise 27003, App.EXEName & ".LoadRes", WinError(Err.LastDllError)
        FreeLibrary (hInst)
        Exit Function
    End If
    
    hSize = SizeofResource(hInst, hFind)
    If hSize = 0 Then
        Err.Raise 27004, App.EXEName & ".LoadRes", WinError(Err.LastDllError)
        FreeLibrary (hInst)
        Exit Function
    End If
    
    ReDim b(0 To lsize) As Byte
    CopyMemory b(0), hLock, hSize
    
    On Error Resume Next
    Kill strFileName
    On Error GoTo Nnapaka
    
    iFile = FreeFile
    Open strFileName For Binary Access Write Lock Read As iFile
        Put #iFile, , b
    Close #iFile
    iFile = 0
    
    FreeLibrary hInst
    LoadDLLHtml = b

    Exit Function
    
Nnapaka:
    Err.Raise Err.Number, App.EXEName & ".LoadRes", Err.Description
    If Not (iFile = 0) Then
        Close #iFile
    End If
    If hInst <> 0 Then FreeLibrary (hInst)
    Exit Function
End Function

'This function loads any picture from DLL (if it's valid
'VB picture of course). It can be used if you have JPEGs
'in resource file under resource type other than "JPEG" and
'other similar situations.
'Most of it has already been explained before.
Public Function LoadAnyPictureFromDLL(ID As Long, resType As String) As IPictureDisp
Attribute LoadAnyPictureFromDLL.VB_Description = "Loads any valid picture format (bitmap, icon, cursor, jpeg, gif). You can use it when resources  doesn't use usual resource types."
    Dim sFile As String

    sFile = Trim(GetTempFile())
    
    On Error GoTo Nnapaka
    
    If SaveAnyResource(ID, resType, sFile) = False Then
        Err.Raise 27008, App.EXEName & ".LoadRes", WinError(Err.LastDllError)
        FreeLibrary (hInst)
        Exit Function
    End If

    Set LoadAnyPictureFromDLL = LoadPicture(sFile)
    
    Kill sFile
    sFile = Empty
    Exit Function

Nnapaka:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Sub About()
Attribute About.VB_Description = "Shows about box."
Attribute About.VB_UserMemId = -552
    If Right(App.EXEName, 3) = "dll" Then
        frmAbout.Show vbModal
    Else
        MsgBox "ResDllLoad 1.0" & vbNewLine & "Class for loading resources from resource-only DLL." & vbNewLine & "Copyright 2000 by Zvonko Bostjancic - zvonko.bostjancic@siol.net", vbInformation, "LoadRes"
    End If
End Sub

'This sub plays wave from resource-only DLL
Sub PlayWaveFromDLL(lngID As Long, Optional resType As String = "WAVE")
Attribute PlayWaveFromDLL.VB_Description = "Plays wave, stored in resource DLL."
    Dim hInst As Long
    Dim sName As String

    If CheckDllName = False Then Exit Sub
    
    On Error GoTo Nnapaka
    
    sName = "#" & lngID
    
    hInst = LoadLibrary(DllName)
    If hInst = 0 Then
        Err.Raise 27006, App.EXEName & ".LoadRes", WinError(Err.LastDllError)
        Exit Sub
    End If

    'play sound
    r = PlaySound(sName, hInst, SND_RESOURCE Or SND_SYNC)
    If r = 0 Then
        Err.Raise 27010, App.EXEName & ".LoadRes", WinError(Err.LastDllError)
        FreeLibrary (hInst)
        Exit Sub
    End If
    
    FreeLibrary (hInst)
    Exit Sub
    
Nnapaka:
    Err.Raise Err.Number, App.EXEName & ".LoadRes", Err.Description
    If Not (iFile = 0) Then
        Close #iFile
    End If
    If hInst <> 0 Then FreeLibrary (hInst)
    Exit Sub
End Sub
