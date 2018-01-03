VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cShellLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------
'- Public enums...
'---------------------------------------------------------------
Public Enum STGM
    STGM_DIRECT = &H0&
    STGM_TRANSACTED = &H10000
    STGM_SIMPLE = &H8000000
    STGM_READ = &H0&
    STGM_WRITE = &H1&
    STGM_READWRITE = &H2&
    STGM_SHARE_DENY_NONE = &H40&
    STGM_SHARE_DENY_READ = &H30&
    STGM_SHARE_DENY_WRITE = &H20&
    STGM_SHARE_EXCLUSIVE = &H10&
    STGM_PRIORITY = &H40000
    STGM_DELETEONRELEASE = &H4000000
    STGM_CREATE = &H1000&
    STGM_CONVERT = &H20000
    STGM_FAILIFTHERE = &H0&
    STGM_NOSCRATCH = &H100000
End Enum

Public Enum SHELLFOLDERS            ' Shell Folder Path Constants...
    CSIDL_DESKTOP = &H0&            ' ..\WinNT\profiles\username\Desktop
    CSIDL_PROGRAMS = &H2&           ' ..\WinNT\profiles\username\Start Menu\Programs
    CSIDL_CONTROLS = &H3&           ' No Path
    CSIDL_PRINTERS = &H4&           ' No Path
    CSIDL_PERSONAL = &H5&           ' ..\WinNT\profiles\username\Personal
    CSIDL_FAVORITES = &H6&          ' ..\WinNT\profiles\username\Favorites
    CSIDL_STARTUP = &H7&            ' ..\WinNT\profiles\username\Start Menu\Programs\Startup
    CSIDL_RECENT = &H8&             ' ..\WinNT\profiles\username\Recent
    CSIDL_SENDTO = &H9&             ' ..\WinNT\profiles\username\SendTo
    CSIDL_BITBUCKET = &HA&          ' No Path
    CSIDL_STARTMENU = &HB&          ' ..\WinNT\profiles\username\Start Menu
    CSIDL_DESKTOPDIRECTORY = &H10&  ' ..\WinNT\profiles\username\Desktop
    CSIDL_DRIVES = &H11&            ' No Path
    CSIDL_NETWORK = &H12&           ' No Path
    CSIDL_NETHOOD = &H13&           ' ..\WinNT\profiles\username\NetHood
    CSIDL_FONTS = &H14&             ' ..\WinNT\fonts
    CSIDL_TEMPLATES = &H15&         ' ..\WinNT\ShellNew
    CSIDL_COMMON_STARTMENU = &H16&  ' ..\WinNT\profiles\All Users\Start Menu
    CSIDL_COMMON_PROGRAMS = &H17&   ' ..\WinNT\profiles\All Users\Start Menu\Programs
    CSIDL_COMMON_STARTUP = &H18&    ' ..\WinNT\profiles\All Users\Start Menu\Programs\Startup
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19& '..\WinNT\profiles\All Users\Desktop
    CSIDL_APPDATA = &H1A&           ' ..\WinNT\profiles\username\Application Data
    CSIDL_PRINTHOOD = &H1B&         ' ..\WinNT\profiles\username\PrintHood
End Enum

Public Enum SHOWCMDFLAGS
    SHOWNORMAL = 5
    SHOWMAXIMIZE = 3
    SHOWMINIMIZE = 7
End Enum

Const MAX_PATH = 260

'---------------------------------------------------------------
Public Function GetShellLinkInfo(LnkFile As String, ExeFile As String, WorkDir As String, _
                                 ExeArgs As String, IconFile As String, IconIdx As Long, _
                                 ShowCmd As Long, Description As String) As Long
'---------------------------------------------------------------
    Dim pidl As Long                                    ' Item id list
    Dim wHotKey As Long                                 ' Hotkey to shortcut...
    Dim FD As WIN32_FIND_DATA
    Dim buffLen As Long
    Dim cShellLink As ShellLinkA                        ' An explorer IShellLink instance
    Dim cPersistFile As IPersistFile                    ' An explorer IPersistFile instance
'---------------------------------------------------------------
    If (LnkFile = "") Then Exit Function                ' Validate min. input requirements.
    
    Set cShellLink = New ShellLinkA                     ' Create new IShellLink interface
    Set cPersistFile = cShellLink                       ' Implement cShellLink's IPersistFile interface
    
    ' Load Shortcut file...(must do this UNICODE hack!)
    On Error GoTo errhandler
    cPersistFile.Load StrConv(LnkFile, vbUnicode), STGM_DIRECT
    
    With cShellLink
        ' Get command line exe name & path of shortcut
        ExeFile = Space(260)
        buffLen = Len(ExeFile)
        .GetPath ExeFile, buffLen, FD, SLGP_UNCPRIORITY
        Dim s As String
        s = FD.cFileName                                ' Not returned to calling function
        
        ' Get working directory of shortcut
        WorkDir = Space(MAX_PATH)
        buffLen = Len(WorkDir)
        .GetWorkingDirectory WorkDir, buffLen
        
        ' Get command line arguments of shortcut
        ExeArgs = Space(MAX_PATH)
        buffLen = Len(ExeArgs)
        .GetArguments ExeArgs, buffLen
        
        ' Get description of shortcut
        Description = Space(MAX_PATH)
        buffLen = Len(Description)
        .GetDescription Description, buffLen
        
        ' Get the HotKey for shortcut
        .GetHotkey wHotKey                              ' Not returned to calling function
       
        ' Get shortcut icon location & index
        IconFile = Space(MAX_PATH)
        buffLen = Len(IconFile)
        .GetIconLocation IconFile, buffLen, IconIdx
        
        ' Get Item ID List...
        .GetIDList pidl                                 ' Not returned to calling function
                
        ' Set shortcut's startup mode (min,max,normal)
        .GetShowCmd ShowCmd
    End With

    GetShellLinkInfo = True                             ' Return Success
'---------------------------------------------------------------
errhandler:
'---------------------------------------------------------------
    Set cPersistFile = Nothing                          ' Destroy Object
    Set cShellLink = Nothing                            ' Destroy Object
'---------------------------------------------------------------
End Function
'---------------------------------------------------------------
