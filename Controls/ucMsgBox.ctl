VERSION 5.00
Begin VB.UserControl ucMsgBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ucMsgBox.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ucMsgBox.ctx":045C
End
Attribute VB_Name = "ucMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+  File Description:
'       ucMsgBox - A Selfsubclassed Theme Aware ucMsgBox Control which Provides Correct Theme Visualization
'
'   Product Name:
'       ucMsgBox.ctl
'
'   Compatability:
'       Widnows: 9x, ME, NT, 2K, XP
'
'   Software Developed by:
'       Paul R. Territo, Ph.D
'
'   Based on the following On-Line Articles
'       (Paul Caton - Self-Subclassser)
'           http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'       (Mario Flores - WinXPEngine)
'           http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=51400&lngWId=1
'       (Fred.cpp - isButton)
'           http://www.devx.com/vb2themax/Tip/19239
'       (Randy Birch - IsWinXP)
'           http://vbnet.mvps.org/code/system/getversionex.htm
'       (Randy Birch - TrimNull)
'           http://vbnet.mvps.org/code/core/trimnull.htm
'       (Randy Birch - Center MsgBox)
'           http://vbnet.mvps.org/code/hooks/messageboxhookcentre.htm
'       (Randy Birch - Timed MsgBox Destroy)
'           http://vbnet.mvps.org/code/hooks/messageboxhooktimerapi.htm
'       (Aftab Mahar - Set Process Priority)
'           http://www.developerspk.com/showthread.php?t=20
'       (Philip Manavopoulos & Aaron Young - MsgBox BackColor / ForeColor)
'           http://www.vbforums.com/showthread.php?t=329373
'       (MSDN eMagazine - Replacing the Icon)
'           http://msdn.microsoft.com/msdnmag/issues/02/11/CuttingEdge/
'
'   Legal Copyright & Trademarks:
'       Copyright © 2006, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'       Trademark ™ 2006, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'
'   Comments:
'       No claims or warranties are expressed or implied as to accuracy or fitness
'       for use of this software. Advance Research Systems shall not be liable for
'       any incidental or consequential damages suffered by any use of this software.
'       This software is owned by Paul R. Territo, Ph.D and is sold for use as a
'       license in accordance with the terms of the License Agreement in the
'       accompanying the documentation.
'
'   Contact Information:
'       For Technical Assistance:
'       pwterrito@insightbb.com
'
'   Modification(s) History:
'
'       13Aug06 - Initial Usercontrol Build (Modified from WinXPEngine)
'       15Aug06 - Fixed Button Focus Bug in Drawing Routine which painted them as StateNormal
'       16Aug06 - Fixed the BrowseForFolder (BFF) loading conflict which prevented the global
'                 subclassing of the MsgBox. The BFF uses the exact same ClassName, but different
'                 WndowCaption than the MsgBox, so this is how we can differentiate it...
'               - Added Alignment property and code to allow for MsgBox alignment
'               - Added DrawFilledRect and Translate Color routines
'               - Added DrawFilledRect to DrawWinXPButton routine to clean up the button painting
'                 along the edges and corners
'       20Aug06 - Fixed the Common Dialog (CD) and BFF ClassName conflict once again....it seems that
'                 all three window classes share the same base Class, but have different captions.
'                 Since the CD's caption can be changed this was not a reliable method for identification.
'                 Given this, and the fact that the Window Styles were unique between the objects, we
'                 could differentiate the MsgBox by ClassName and WindowStyle.
'               - Fixed Looping bug when stop subclassing the MsgBox
'               - Fixed Painting error when the mouse was down over a button and we moved...This
'                 caused the button to not be painted in the correct style and showed the default
'                 Win9x style instead of our down state.
'       11Oct06 - Moved Alignemt and Theme properties into the m_MsgBox structure
'               - Added SelfClosing property, type, and m_MsgBox structure elements
'               - Added Duration property, type, and m_MsgBox structure elements
'               - Added Dialog Prompt setting based on time left
'       28Nov06 - Added Parent and ParentType to the m_MsgBox structure to allow for tracking of the
'                 hWnd and Type of the Host Object
'               - Fixed lLeft, lTop, lWidth, and lHeight calculations in the zSubclass_Proc to provide
'                 a method to compute these when the host object is being destroyed and has gone
'                 out of scope. Under these conditions, the reference to UserControl.Parent.Left (.Top...etc)
'                 are not valid and throw an error which freezes the IDE since the MsgBox is Modal!!
'               - Added ThreadPriority property to allow for real-time priority for the thread which allows the
'                 Timer to have the New Priority other than "mbNormal"
'               - Added SetProcessPriority to set the process priority
'               - Fixed TimerProc bug which resulted in timer intervals which were incrementally
'                 shorter on subsequent calls....cause: passing ByVal 0& instead of AddressOf Proc pointer
'                 from sc_aSubData array.
'               - Added CenterCaption to permit message centering for all passed prompts
'       03Dec06 - Added AutoCenterPrompt property to allow for centering of the prompt from the property dialog
'               - Added ForeColor Property to allow for Prompt and Button text being changed
'               - Added BackColor Property to allow the Main form BackColor to be changed
'               - Optimized DrawWinXPButton to correctect for stray pixels around each button
'               - Optimized CenterCaption to now compute the correct spacing for centering of
'                 the passed Prompt string
'               - Set Init Properties to match the PropertyRead/Write defaults
'       07Dec06 - Removed CenterCaption Method
'               - Removed AutoCenterPrompt property
'               - Added AlignPrompt property to replace the two above methods/prop.
'               - Added an All API method (SetWindowLong) for Aligning the Prompt Text instead of padding the string.
'               - Added Icon Property to allow for custom MsgBox icon replacment
'               - Added DrawWin9xButton routine and associated methods to replicate the actual dialog
'                 3D buttons. This allows us complete control of the colors and painting.
'               - Added Support for Storing Either Win9x or WinXP States depending on the Theme
'       08Dec06 - Fixed Minor bug in the SDIHost_QueryUnload and MDIHost_QueryUnload which incorrectly
'                 set the subclassing flag and attempted to re-subclass the object on shutdown
'               - Added hDC Property to allow Painting via this Mechanism
'               - Added Additional Status Messages
'       11Dec06 - Fixed Minor bug in the zSubclass_Proc, section WM_CTLCOLORDLG, WM_CTLCOLORSTATIC
'                 where we were we passed invalid Icon.Pictures to the SendMessage API.
'       18Dec06 - Added CustomCaption property to allow the dynamic switching of Default or Customs Captions
'               - Added Caption Property and Structure element to permit custom button captions
'               - Updated DrawWin9xButton and DrawWinXPButton methods to not permit custom dialog
'                 button captions. Note: From a Software Engineering standpoint the CustomCaption structure
'                 elements should be in the Button button field (i.e. m_MsgBox.Button(n).CustomCaption), but these are
'                 dynamic arrays which are built just prior to the MsgBox being shown. As such, we have
'                 choosen to place this struture element at the m_MsgBox level (i.e. m_MsgBox.CustomCaption(n)
'                 so that they can be passed at runtime....
'               - Added Subclass_StopAll to the zIdx method to prevent dialog faults which hault the mouse
'                 and keyboard on subclassing errors. We also removed the Debug.Assert False statment
'                 as this is the main cause of issues when working with modal dialogs of this sort.
'       19Dec06 - Fixed minor bug in the SetProcessPriority method which hard coded the value....left over from testing ;-D
'               - Added With structure to WM_ACTIVATE message section to allow for greater clarity
'       24Feb07 - Added Expanded Theme Emulation for XP Luna Blue, HomeStead, and Metallic to the DrawXPButton routines
'               - Bug still remains in the DrawAlphaVGradient which provides the "highlight" alphablending
'                 between the background gradient and white
'       11Mar07 - Removed DrawAlphaVGradient and AlphaBlend methods and replaced with Direct Gradient calls
'                 to DrawVGradient and DrawVGradientEx for Metallic Luna Styles
'               - Updated all Luna Metallic painting routines to reflect the corrected emulation style
'               - Extended Theme Enum to encompass all possible default Luna Themes
'                 (mbBlue (Blue), mbHomeStead (Olive Green), mbMetallic (Silver))
'
'-  Notes:
'       This ucMsgBox control intercepts the ucMsgBox Call to the systems MessageBox API, and provides
'       the ability to Paint the controls surfaces as we see fit. In addtion, with a little care we
'       can position the control before it is shown and adjust the BackColors, ForeColors, Fonts
'       and even the MessageBox Icons. The System Icons are ~32x32 Pixels with 24BPP if you plan to
'       pass them from VB controls like ImageLists, Image, or PictureBoxes. The Icon window can handle
'       icons which are 48x48 Pixels (24BPP) as well, but care should be taken to test these out prior
'       to using them directly. The hDC is the DC to the Icon Window at "runtime" since the object does
'       not exist until the MsgBox is called from the Host Object
'
'       There is one known issues with this control that can cause the Modal dialog to hault.
'       One key issue is Double Subclassing or Failing to UnSubclass the MsgBox. This will allow the
'       dialog to be painted the first round, but will cause a Debug.Assert in the zIdx method on all
'       subsequent calls to the daialog. The result will be a hault of the subclasser and a Modal MsgBox
'       Dialog....the net result being that the mouse will not work in either the IDE, Dialog, or Host Object.
'       The only way to exit this is to F8 (Run) until you throw the exeption dialog in the IDE and click End.
'       Be forewarned this may cause the IDE to crash!! I have extensively tested this and only has this happen
'       when I did not follow the above rules....in short, the control is stable, provided the above caviate.
'
'       One Minor bug still exists when using fonts which are significantly different than Ambient.Font. In this
'       case we will need to adjust the Prompt Rect to fit the new Font proportions. This can be achieved by the
'       DrawText with the DT_CALCRECT flag. This is still under development....
'
'   Build Date & Time: 3/11/2007 11:52:32 PM
Const Major As Long = 1
Const Minor As Long = 0
Const Revision As Long = 240
Const DateTime As String = "3/11/2007 11:52:32 PM"
'
'   Force Declarations
Option Explicit

'   Private Constants
Private Const VER_PLATFORM_WIN32_NT As Long = 2

'   Msgbox Constant
Private Const WC_DIALOG As String = "#32770"        'Win32 Classname for MsgBox Dialog
Private Const MB_ICON As Long = &H14

'   Window Enum Constants
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)
Private Const GW_HWNDNEXT As Long = 2
Private Const GW_CHILD As Long = 5
Private Const GW_HWNDFIRST As Long = 0

'   DrawText Flags
Private Const DT_CENTER As Long = &H1
Private Const DT_VCENTER As Long = &H4

'   MsgBox Prompt Flag
Private Const IDPROMPT As Long = &HFFFF&
Private Const IDOK As Long = 1
Private Const IDCANCEL As Long = 2
Private Const IDABORT As Long = 3
Private Const IDRETRY As Long = 4
Private Const IDIGNORE As Long = 5
Private Const IDYES As Long = 6
Private Const IDNO As Long = 7

'   Private MsgBoxTimerID
Private Const MB_TIMERID As Long = 999

'   Private Class Priority Constants
Private Const REALTIME_PRIORITY_CLASS = &H100
Private Const HIGH_PRIORITY_CLASS = &H80
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const IDLE_PRIORITY_CLASS = &H40

'   Private TextAlign Flags for Static Windows
Private Const SS_LEFT As Long = &H0
Private Const SS_RIGHT As Long = &H2
Private Const SS_CENTER As Long = &H1

'   Private Icon Flag
Private Const STM_SETICON As Long = &H170

'   Os Window Types
Private Type OSVERSIONINFO
    OSVSize         As Long         'size, in bytes, of this data structure
    dwVerMajor      As Long         'ie NT 3.51, dwVerMajor = 3; NT 4.0, dwVerMajor = 4.
    dwVerMinor      As Long         'ie NT 3.51, dwVerMinor = 51; NT 4.0, dwVerMinor= 0.
    dwBuildNumber   As Long         'NT: build number of the OS
                                    'Win9x: build number of the OS in low-order word.
                                    '       High-order word contains major & minor ver nos.
    PlatformID      As Long         'Identifies the operating system platform.
    szCSDVersion    As String * 128 'NT: string, such as "Service Pack 3"
                                    'Win9x: string providing arbitrary additional information
End Type

'   Required Type Definitions
Private Type RGBQUAD
    rgbBlue                     As Byte
    rgbGreen                    As Byte
    rgbRed                      As Byte
    rgbReserved                 As Byte
End Type

Private Type POINT
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type LOGFONT
    lfHeight          As Long
    lfWidth           As Long
    lfEscapement      As Long
    lfOrientation     As Long
    lfWeight          As Long
    lfItalic          As Byte
    lfUnderline       As Byte     '   //--The LOGFONT structure defines the attributes of a font.
    lfStrikeOut       As Byte
    lfCharSet         As Byte
    lfOutPrecision    As Byte
    lfClipPrecision   As Byte
    lfQuality         As Byte
    lfPitchAndFamily  As Byte
    lfFacename        As String * 32
End Type

'   Private MsgBox Types for Holding Object Data
Private Type mbButtonType
    hWnd As Long                        'Handle of the Button
    State As mbState                    'Current State of the Button
    Caption As String                   'Caption of the Button
End Type

Private Type mbPromptType
    hWnd As Long                        'Handle to the Prompt
    hdc As Long                         'DC of the Prompt
    lpRect As RECT                      'Rect of the Prompt
    Caption As String * 1024            'Prompt Caption of the Main Dialog...
    ForeColor As OLE_COLOR              'ForeColor Property
    Alignment As mbPromptAlignEnum      'Prompt Aligment Flag
End Type

Private Type mbIcon
    hWnd As Long                        'Icon Window handle
    hdc As Long                         'DC of the Icon window
    lpRect As RECT                      'Rect of the Window
    Picture As StdPicture               'Picture Property of the Icon
    PrevPicture As Long                 'Handle to the Previous Picture
End Type

Private Type mbTimerType
    TimerDuration As Long               'How long a Timer Waits
    TimerID As Long                     'Unique ID for the Timer
    ThreadPriority As mbPriorityEnum    'Process Priority for MsgBox Timer
    TimerToken As String                'Token String for the CountDown Timer
End Type

Private Type mbMsgBoxType
    Alignment As mbAlignEnum            'MsgBox Alignment Type
    BackColor As OLE_COLOR              'BackColor Property
    Button() As mbButtonType            'Array of Buttons
    CustomCaption(3) As String          'Custom Caption for the Button
    CaptionType As mbCaptionEnum        'Type of Button Captions to Use
    Count As Long                       'Number of Buttons in the Array
    DefaultIndex As Long                'Index of the Default Button
    DownIndex As Long                   'Index of the Down Button
    Font As Font                        'Font to Use
    hdc As Long                         'MsgBox Form DC
    hWnd As Long                        'MsgBox Form Handle which Hosts the Buttons
    Icon As mbIcon                      'Handle to Icon Window
    lpRect As RECT                      'Rect of the MsgBox Dialog
    Parent As Long                      'Host Objects hWnd
    ParentType As mbParentTypeEnum      'Type of Object which hosts the MsgBox
    Prompt As mbPromptType              'Prompt Infomation
    SelfClosing As Boolean              'Use Timer for SelfClosing...
    Theme As mbThemeEnum                'Current Theme
    TimerInfo As mbTimerType            'Timer Specific Information....
End Type

Private Enum mbParentTypeEnum
    [mbSDIForm] = &H0
    [mbMDIForm] = &H1
End Enum
#If False Then
    Const mbSDIForm = &H0
    Const mbMDIForm = &H1
#End If

Private Enum mbState
    StateNormal = &H1
    StateHot = &H2
    StatePressed = &H3
    StateDisabled = &H4
    StateDefaulted = &H5
End Enum
#If False Then
    Const StateNormal = &H1
    Const StateHot = &H2
    Const StatePressed = &H3
    Const StateDisabled = &H4
    Const StateDefaulted = &H5
#End If

Public Enum mbAlignEnum
    [mbCenterScreen] = &H0
    [mbCenterOwner] = &H1
End Enum
#If False Then
    Const mbCenterScreen = &H0
    Const mbCenterOwner = &H1
#End If

Public Enum mbCaptionEnum
    [mbDefault] = &H0
    [mbCustom] = &H1
End Enum
#If False Then
    Const mbDefault = &H0
    Const mbCustom = &H1
#End If

Public Enum mbPromptAlignEnum
    [mbLeft] = SS_LEFT
    [mbRight] = SS_RIGHT
    [mbCenter] = SS_CENTER
End Enum
#If False Then
    Const mbLeft = SS_LEFT
    Const mbRight = SS_RIGHT
    Const mbCenter = SS_CENTER
#End If

Public Enum mbPriorityEnum
    [mbRealTime] = REALTIME_PRIORITY_CLASS          'Highest Priority
    [mbHigh] = HIGH_PRIORITY_CLASS                  'Second Highest Priority
    [mbNormal] = NORMAL_PRIORITY_CLASS              'Third Highest Priority
    [mbIdle] = IDLE_PRIORITY_CLASS                  'Lowest Priority
End Enum
#If False Then
    Const mbRealTime = REALTIME_PRIORITY_CLASS
    Const mbHigh = HIGH_PRIORITY_CLASS
    Const mbNormal = NORMAL_PRIORITY_CLASS
    Const mbIdle = IDLE_PRIORITY_CLASS
#End If

Public Enum mbThemeEnum
    [mbAuto] = &H0
    [mbClassic] = &H1
    [mbBlue] = &H2
    [mbHomeStead] = &H3
    [mbMetallic] = &H4
End Enum
#If False Then
    Const mbAuto = &H0
    Const mbClassic = &H1
    Const mbBlue = &H2
    Const mbHomeStead = &H3
    Const mbMetallic = &H4
#End If

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As String, ByVal dwMaxNameChars As Integer, ByVal pszColorBuff As String, ByVal cchMaxColorChars As Integer, ByVal pszSizeBuff As String, ByVal cchMaxSizeChars As Integer) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function GetDlgItemText Lib "user32" Alias "GetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function LStrLen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function OleTranslateColorA Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function OleTranslateColorEx Lib "oleaut32.dll" Alias "OleTranslateColor" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

'   Private variables
Private lzCount             As Long             'zSubclass Index to the Proc Address
Private lCount              As Long             'Counter for the Timer
Private bMsgBoxSubClass     As Boolean          'SubClass Flag of the MsgBox
Private bTimerRunning       As Boolean          'Timer Running Flag for SelfClose
Private m_AutoTheme         As String           'String Var to hold the Theme info
Private m_MsgBox            As mbMsgBoxType     'MsgBox Type which hold the MsgBox Object Data
Private m_Theme             As mbThemeEnum      'MsgBox Theme for property "Get"
Private bHasIcon            As Boolean          'Flag which indicates the the dialog is using the icon

Private WithEvents SDIHost  As Form
Attribute SDIHost.VB_VarHelpID = -1
Private WithEvents MDIHost  As MDIForm
Attribute MDIHost.VB_VarHelpID = -1

'==================================================================================================
' ucSubclass - A template UserControl for control authors that require self-subclassing without ANY
'              external dependencies. IDE safe.
'
' Paul_Caton@hotmail.com
' Copyright free, use and abuse as you see fit.
'
' v1.0.0000 20040525 First cut.....................................................................
' v1.1.0000 20040602 Multi-subclassing version.....................................................
' v1.1.0001 20040604 Optimized the subclass code...................................................
' v1.1.0002 20040607 Substituted byte arrays for strings for the code buffers......................
' v1.1.0003 20040618 Re-patch when adding extra hWnds..............................................
' v1.1.0004 20040619 Optimized to death version....................................................
' v1.1.0005 20040620 Use allocated memory for code buffers, no need to re-patch....................
' v1.1.0006 20040628 Better protection in zIdx, improved comments..................................
' v1.1.0007 20040629 Fixed InIDE patching oops.....................................................
' v1.1.0008 20040910 Fixed bug in UserControl_Terminate, zSubclass_Proc procedure hidden...........
'==================================================================================================
'Subclasser declarations

Public Event Status(ByVal sStatus As String)

Private Const WM_ACTIVATE               As Long = &H6
Private Const WM_CLOSE                  As Long = &H10
Private Const WM_COMMAND                As Long = &H111
Private Const WM_CREATE                 As Long = &H1
Private Const WM_CTLCOLORSTATIC         As Long = &H138
Private Const WM_CTLCOLORDLG            As Long = &H136
Private Const WM_EXITSIZEMOVE           As Long = &H232
Private Const WM_LBUTTONDOWN            As Long = &H201
Private Const WM_LBUTTONUP              As Long = &H202
Private Const WM_MOUSELEAVE             As Long = &H2A3
Private Const WM_MOUSEMOVE              As Long = &H200
Private Const WM_MOVING                 As Long = &H216
Private Const WM_NCPAINT                As Long = &H85
Private Const WM_PAINT                  As Long = &HF
Private Const WM_RBUTTONDBLCLK          As Long = &H206
Private Const WM_RBUTTONDOWN            As Long = &H204
Private Const WM_SETFOCUS               As Long = &H7
Private Const WM_SIZING                 As Long = &H214
Private Const WM_SYSCOLORCHANGE         As Long = &H15
Private Const WM_THEMECHANGED           As Long = &H31A
Private Const WM_TIMER                  As Long = &H113

Private Enum TRACKMOUSEEVENT_FLAGS
  TME_HOVER = &H1&
  TME_LEAVE = &H2&
  TME_QUERY = &H40000000
  TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
  cbSize                             As Long
  dwFlags                            As TRACKMOUSEEVENT_FLAGS
  hwndTrack                          As Long
  dwHoverTime                        As Long
End Type

Private bTrack                       As Boolean
Private bTrackUser32                 As Boolean
Private bInCtrl                      As Boolean
Private bSubClass                    As Boolean

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Enum eMsgWhen
    MSG_AFTER = 1                                                                   'Message calls back after the original (previous) WndProc
    MSG_BEFORE = 2                                                                  'Message calls back before the original (previous) WndProc
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                  'Message calls back before and after the original (previous) WndProc
End Enum
#If False Then
    Private Const MSG_AFTER = 1                                                                   'Message calls back after the original (previous) WndProc
    Private Const MSG_BEFORE = 2                                                                  'Message calls back before the original (previous) WndProc
    Private Const MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                  'Message calls back before and after the original (previous) WndProc
#End If

Private Const ALL_MESSAGES           As Long = -1                                   'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                    'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                   'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                   'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                   'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                  'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                  'Table A (after) entry count patch offset

Private Type tSubData                                                               'Subclass data type
    hWnd                               As Long                                      'Handle of the window being subclassed
    nAddrSub                           As Long                                      'The address of our new WndProc (allocated memory).
    nAddrOrig                          As Long                                      'The address of the pre-existing WndProc
    nMsgCntA                           As Long                                      'Msg after table entry count
    nMsgCntB                           As Long                                      'Msg before table entry count
    aMsgTblA()                         As Long                                      'Msg after table array
    aMsgTblB()                         As Long                                      'Msg Before table array
End Type

Private sc_aSubData()                As tSubData                                    'Subclass data array

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)

'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
    'Parameters:
        'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
        'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
        'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
        'hWnd     - The window handle
        'uMsg     - The message number
        'wParam   - Message related data
        'lParam   - Message related data
    'Notes:
        'If you really know what you're doing, it's possible to change the values of the
        'hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
        'values get passed to the default handler.. and optionaly, the 'after' callback
    Dim lhWnd As Long
    Dim ClassWindow As String * 50
    Dim ClassCaption As String * 100
    Dim i As Long
    Dim j As Long
    Dim lpMsgBoxRect As RECT
    Dim lLeft As Long
    Dim lTop As Long
    Dim lWidth As Long
    Dim lHeight As Long
    Dim lIndex As Long
    Dim newLeft As Long
    Dim newTop As Long
    Dim dlgWidth As Long
    Dim dlgHeight As Long
    Dim scrWidth As Long
    Dim scrHeight As Long
    Dim hWndStyle As Long
    Dim hWndStyleEx As Long
    Dim sPrompt As String
    
    Select Case uMsg
        Case WM_ACTIVATE
            '   Get the Theme and store it....
            m_AutoTheme = GetThemeInfo
            '   Get the Window Style and Extended Style
            hWndStyle = GetWindowLong(lParam, GWL_STYLE)
            hWndStyleEx = GetWindowLong(lParam, GWL_EXSTYLE)
            'Debug.Print "Window Style: &H" & Hex(hWndStyle), "&H" & Hex(hWndStyleEx), "&H" & (Hex(hWndStyle Or hWndStyleEx))
            '   Get the Class Window Styles and make sure this is not the BrowseForFolder (BFF) or the
            '   Common Dialog which also use this same #32770 ClassName. To differentiate these from
            '   MsgBox we get the window Styles which are different between the three objects...namely the
            '   WindowStyles (Style, ExStyle) are as follows: MsgBox (&H84C801C5, &H10101), and the
            '   BrowseForFolder(&H4C820CC, &H10501) and the Common Dailog (&H86CC20C4, &H10501), which then
            '   bitwise "OR" yeilds unique ID that allows us to ignore them when they occur!!
            If (InStr(ThisWindowClassName(lParam), WC_DIALOG) > 0) And ((hWndStyle Or hWndStyleEx) = &H84C901C5) Then
                '   ///////////////////////////////////////////////////////////////
                '   This is the original way Mario Flores was able to locate
                '   the Buttons in MsgBox, but this does not work in UserControls
                '   since the API utilizes a CallBack Function which we can not
                '   address directly....
                'EnumChildWindows lParam, AddressOf EnumChildProc, ByVal 0&
                '   ///////////////////////////////////////////////////////////////
                '   The following is an equivalant routine....but without CallBacks
                '
                '   Get the Prompt string and store the value
                Call GetDlgItemText(lParam, IDPROMPT, m_MsgBox.Prompt.Caption, 1024)
                If (Not bMsgBoxSubClass) Then
                    '   Enumerate child windows...the hard way!!!
                    '
                    '   The MsgBox hWnd is passed by the lParam value
                    '
                    '   Subclass the Main MsgBox Form
                    With m_MsgBox
                        '   Store the Handle of the MsgBox
                        .hWnd = lParam
                        Call Subclass_Start(.hWnd)
                        Call Subclass_AddMsg(.hWnd, WM_MOUSEMOVE, MSG_AFTER)
                        Call Subclass_AddMsg(.hWnd, WM_CLOSE, MSG_BEFORE)
                        '   Subclass the Msgbox Color Dialog messages so we can set the
                        '   BackColor and ForeColor
                        Call Subclass_AddMsg(.hWnd, WM_CTLCOLORDLG, MSG_AFTER)
                        Call Subclass_AddMsg(.hWnd, WM_CTLCOLORSTATIC, MSG_AFTER)
                        '   Now we find the First Button!
                        lhWnd = GetWindow(.hWnd, GW_CHILD)
                        '   Get the ClassName
                        GetClassName lhWnd, ClassWindow, 50
                        '   Get the Caption of the Object
                        GetWindowText lhWnd, ClassCaption, 100
                        'Debug.Print lhWnd, Left$(ClassWindow, 6), Left$(ClassCaption, 6)
                        '   Make sure we found a Button on th MsgBox
                        If (lhWnd <> 0) And (Left$(ClassWindow, 6) = "Button") Then
                            '   Store the First Button Value
                            ReDim .Button(i)
                            With .Button(i)
                                .hWnd = lhWnd
                                .Caption = TrimNull(ClassCaption)
                            End With
                            '   Now check for the remaining buttons, if any...
                            '   Note we only let this Do loop run 5 times....this
                            '   should prevent any endless loops...and besides there
                            '   are never more then 3 buttons ;-)
                            Do While (lhWnd <> 0) And (i <= 4)
                                '   Get the Next Item in the MsgBox
                                lhWnd = GetWindow(lhWnd, GW_HWNDNEXT)
                                '   Get the ClassName
                                GetClassName lhWnd, ClassWindow, 50
                                '   Check to see if this is a button....because
                                '   there are static windows which are used to
                                '   display Icons.....if one wished we could trap
                                '   these and change them by subclassing them
                                '   here as we are for the buttons....
                                If Left$(ClassWindow, 6) = "Button" Then
                                    '   We found one more....so increment things
                                    i = i + 1
                                    '   Now store the new value
                                    ReDim Preserve .Button(i)
                                    '   Get the Caption of the Object
                                    GetWindowText lhWnd, ClassCaption, 100
                                    With .Button(i)
                                        .hWnd = lhWnd
                                        .Caption = TrimNull(ClassCaption)
                                    End With
                                    'Debug.Print lhWnd, Left$(ClassWindow, 6), Left$(ClassCaption, 6)
                                End If
                            Loop
                            '   Store the Icon Window for Later
                            .Icon.hWnd = GetDlgItem(.hWnd, MB_ICON)
                            bHasIcon = .Icon.hWnd <> 0
                            '   "i" holds the number of buttons in the MsgBox
                            '   so we will use this to increment the subclassing
                            '   calls to the Thunk....
                            If i >= 0 Then
                                '   Sublass the Usecontrol
                                Call Subclass_Start(UserControl.hWnd)
                                Call Subclass_AddMsg(UserControl.hWnd, WM_TIMER, MSG_AFTER)
                                '   See if we are using SelfClosing Style
                                If .SelfClosing Then
                                    '   Get a Unique ID for the API Timer by getting the Time Since Midnight
                                    '   We need to keep this reasonable, so we will mod this value
                                    m_MsgBox.TimerInfo.TimerID = (Timer Mod 65535)
                                    'Debug.Print "TimerID: " & m_MsgBox.TimerInfo.TimerID
                                    '   Start the local timer to perform the count down
                                    With .TimerInfo
                                        '   Set the Process Priority for the Timer Thread
                                        Call SetProcessPriority(.ThreadPriority)
                                        '   Get the Index of the AddressOf CallBack pointer from the
                                        '   sc_aSubData array....only the first time through....
                                        If (lzCount = 0) Then
                                            'lzCount = UBound(sc_aSubData)
                                            lzCount = 2
                                            'Debug.Print lzCount
                                        End If
                                        .TimerID = SetTimer(UserControl.hWnd, .TimerID, 1000, sc_aSubData(lzCount).nAddrSub)
                                        'Debug.Print "Timer Started @ " & Now()
                                    End With
                                    '   Set the Inital Prompt to Include the Count Down Value
                                    sPrompt = Replace$(.Prompt.Caption, .TimerInfo.TimerToken, .TimerInfo.TimerDuration)
                                    '   Set that the timer is running
                                    bTimerRunning = True
                                Else
                                    sPrompt = .Prompt.Caption
                                End If
                                '   Store the Client hWnd and Rect for Measuring
                                .Prompt.hWnd = GetDlgItem(.hWnd, IDPROMPT)
                                Call GetClientRect(.Prompt.hWnd, .Prompt.lpRect)
                                '   Set the New Text
                                SetDlgItemText .hWnd, IDPROMPT, sPrompt
                                '   Start Subclassing the Buttons
                                For j = 0 To i
                                    '   One last sanity check for the hWnd values
                                    If .Button(i).hWnd <> 0 Then
                                        '   Start Subclassing them
                                        .Count = j + 1
                                        With .Button(j)
                                            Call Subclass_Start(.hWnd)
                                            Call Subclass_AddMsg(.hWnd, WM_MOUSEMOVE, MSG_BEFORE)
                                            Call Subclass_AddMsg(.hWnd, WM_LBUTTONDOWN, MSG_AFTER)
                                            Call Subclass_AddMsg(.hWnd, WM_LBUTTONUP, MSG_AFTER)
                                            Call Subclass_AddMsg(.hWnd, WM_MOUSELEAVE, MSG_AFTER)
                                            Call Subclass_AddMsg(.hWnd, WM_PAINT, MSG_AFTER)
                                            Call Subclass_AddMsg(.hWnd, WM_SETFOCUS, MSG_AFTER)
                                            .State = StateNormal
                                        End With
                                    End If
                                Next
                                '   The default button flag...set in the WM_SETFOCUS
                                .DefaultIndex = -1
                                '   The default down button index...set in the WM_LBUTTONDOWN
                                '   and reset in the WM_LBUTTONUP
                                .DownIndex = -1
                            End If
                        End If
                    End With
                    '   See if we need to Center the MsgBox to the Owner
                    If (m_MsgBox.Alignment = mbCenterOwner) And (m_MsgBox.hWnd) Then
                        '   Get the MsgBox Rect
                        Call GetWindowRect(m_MsgBox.hWnd, lpMsgBoxRect)
                        '   Store this for use....
                        m_MsgBox.lpRect = lpMsgBoxRect
                        '   Get the Host Objects RECT
                        If m_MsgBox.ParentType = mbSDIForm Then
                            lLeft = (SDIHost.Left \ Screen.TwipsPerPixelX)
                            lTop = (SDIHost.Top \ Screen.TwipsPerPixelY)
                            lWidth = (SDIHost.Width \ Screen.TwipsPerPixelX)
                            lHeight = (SDIHost.Height \ Screen.TwipsPerPixelX)
                        Else
                            lLeft = (MDIHost.Left \ Screen.TwipsPerPixelX)
                            lTop = (MDIHost.Top \ Screen.TwipsPerPixelY)
                            lWidth = (MDIHost.Width \ Screen.TwipsPerPixelX)
                            lHeight = (MDIHost.Height \ Screen.TwipsPerPixelX)
                        End If
                        '   Get the Dialogs Width and Height
                        dlgWidth = (lpMsgBoxRect.Right - lpMsgBoxRect.Left)
                        dlgHeight = (lpMsgBoxRect.Bottom - lpMsgBoxRect.Top)
                        '   Get the Screen Width and Height
                        scrWidth = (Screen.Width \ Screen.TwipsPerPixelX)
                        scrHeight = (Screen.Height \ Screen.TwipsPerPixelY)
                        '   Compute the New Top and Left
                        newLeft = lLeft + ((lWidth - dlgWidth) \ 2)
                        newTop = lTop + ((lHeight - dlgHeight) \ 2)
                        '   Keep the Dialog on the Screen!! If we don't do this then the
                        '   MsgBox could end up off screen and we would be stuck because it
                        '   is a Modal dialog!!!!
                        If newLeft < 0 Then newLeft = 0
                        If newLeft > scrWidth Then newLeft = scrWidth - dlgWidth
                        If newTop < 0 Then newTop = 0
                        If newTop > scrHeight Then newTop = scrHeight - dlgHeight
                        '   Now move the Window...before it is shown
                        Call MoveWindow(m_MsgBox.hWnd, newLeft, newTop, dlgWidth, dlgHeight, True)
                        RaiseEvent Status("MsgBox Moved")
                    End If
                    '   Store the flag for when the MsgBox Closes
                    bMsgBoxSubClass = True
                    RaiseEvent Status("MsgBox Initialized")
                Else
                    With m_MsgBox
                        '   Stop subclassing the Buttons
                        For j = 0 To i
                            With .Button(j)
                                Subclass_Stop .hWnd
                            End With
                        Next j
                        '   Stop the Timer if we are using SelfClosing
                        If (.SelfClosing) And (bTimerRunning) Then
                            lCount = 0
                            'Debug.Print "Timer Killed @ " & Now()
                            Call KillTimer(UserControl.hWnd, .TimerInfo.TimerID)
                            bTimerRunning = False
                        End If
                        '   Now stop the Parent Form
                        Subclass_Stop m_MsgBox.hWnd
                    End With
                    '   Reset the counter value to 0
                    i = 0
                    '   Reset our flag
                    bMsgBoxSubClass = False
                    RaiseEvent Status("MsgBox Terminated")
                End If
            End If
        
        Case WM_CLOSE
            'Debug.Print "Closing"
            RaiseEvent Status("MsgBox Closing")
            
        Case WM_CTLCOLORDLG, WM_CTLCOLORSTATIC
            With m_MsgBox
                '   Store the values for later
                .hdc = wParam
                '   Set the BackColor
                SetBkColor wParam, TranslateColor(.BackColor)
                RaiseEvent Status("BackColor Changed")
                '   Get the Prompt hWnd
                .Prompt.hWnd = GetDlgItem(.hWnd, IDPROMPT)
                '   Get the current Window Style
                hWndStyle = GetWindowLong(.Prompt.hWnd, GWL_STYLE)
                '   Set the New Window Style
                SetWindowLong .Prompt.hWnd, GWL_STYLE, .Prompt.Alignment Or hWndStyle
                RaiseEvent Status("PromptAlign Changed")
                '   Store the DC
                .Prompt.hdc = wParam
                '   Get our Rect
                Call GetClientRect(.Prompt.hWnd, .Prompt.lpRect)
                '   Set the BackColor
                SetBkColor .Prompt.hdc, TranslateColor(.BackColor)
                '   Set the ForeColor for the Text....
                '   Note: Button Text Color is set in the DrawButton
                SetTextColor .Prompt.hdc, TranslateColor(.Prompt.ForeColor)
                '   Set the Font
                SelectFont .Prompt.hdc, m_MsgBox.Font.Size, m_MsgBox.Font.Italic, m_MsgBox.Font.Name, m_MsgBox.Font.Underline
                '   Set the Icon the Picture is Valid
                If Not .Icon.Picture Is Nothing Then
                    '   Is there an Icon Window?
                    If (.Icon.hWnd) And (.Icon.Picture) Then
                        '   See if the Old Picture and New are Different
                        '   if so then change them.....since this is called
                        '   over and over via the changes set in the TimerProc
                        '   we need to be careful not to flooding the Static
                        '   window with Picture changes if not needed ;-)
                        'Debug.Print "Set New Image " & Now()
                        If .Icon.PrevPicture <> .Icon.Picture Then
                            '   Set the New Picture in the MsgBox
                            '   and store the old one....
                            .Icon.PrevPicture = SendMessage(.Icon.hWnd, STM_SETICON, .Icon.Picture, ByVal 0&)
                            RaiseEvent Status("Icon Changed")
                        End If
                    End If
                End If
                '   Pass back the handle to the new brush
                lReturn = CreateSolidBrush(TranslateColor(.BackColor))
            End With
                    
        Case WM_LBUTTONDOWN
            If (m_AutoTheme <> "None") And (m_MsgBox.Theme <> mbClassic) Then
                For i = 0 To m_MsgBox.Count - 1
                    If lng_hWnd = (m_MsgBox.Button(i).hWnd) Then
                        Call DrawWinXPButton(m_MsgBox.Button(i).hWnd, StatePressed)
                        m_MsgBox.Button(i).State = StatePressed
                        m_MsgBox.DownIndex = i
                        If (m_MsgBox.SelfClosing) And (bTimerRunning) Then
                            lCount = 0
                            Call KillTimer(UserControl.hWnd, m_MsgBox.TimerInfo.TimerID)
                            bTimerRunning = False
                        End If
                    Else
                        Call DrawWinXPButton(m_MsgBox.Button(i).hWnd, StateNormal)
                        m_MsgBox.Button(i).State = StateNormal
                    End If
                Next i
            Else
                For i = 0 To m_MsgBox.Count - 1
                    If lng_hWnd = (m_MsgBox.Button(i).hWnd) Then
                        Call DrawWin9xButton(m_MsgBox.Button(i).hWnd, StatePressed)
                        m_MsgBox.Button(i).State = StatePressed
                    End If
                Next i
            End If
            
        Case WM_LBUTTONUP
            If (m_AutoTheme <> "None") And (m_MsgBox.Theme <> mbClassic) Then
                m_MsgBox.DownIndex = -1
                For i = 0 To m_MsgBox.Count - 1
                    If lng_hWnd = (m_MsgBox.Button(i).hWnd) Then
                        If (m_MsgBox.Button(i).State <> StateHot) Then
                            Call DrawWinXPButton(m_MsgBox.Button(i).hWnd, StateHot)
                            m_MsgBox.Button(i).State = StateHot
                        End If
                    End If
                Next i
            Else
                For i = 0 To m_MsgBox.Count - 1
                    If lng_hWnd = (m_MsgBox.Button(i).hWnd) Then
                        Call DrawWin9xButton(m_MsgBox.Button(i).hWnd, StateNormal)
                        m_MsgBox.Button(i).State = StateNormal
                    End If
                Next i
            End If
                                    
        Case WM_MOUSELEAVE
            If (m_AutoTheme <> "None") And (m_MsgBox.Theme <> mbClassic) Then
                For i = 0 To m_MsgBox.Count - 1
                    If lng_hWnd = (m_MsgBox.Button(i).hWnd) Then
                        Call DrawWinXPButton(m_MsgBox.Button(i).hWnd, StateNormal)
                        m_MsgBox.Button(i).State = StateNormal
                    End If
                Next i
            Else
                For i = 0 To m_MsgBox.Count - 1
                    If lng_hWnd = (m_MsgBox.Button(i).hWnd) Then
                        Call DrawWin9xButton(m_MsgBox.Button(i).hWnd, StateNormal)
                        m_MsgBox.Button(i).State = StateNormal
                    End If
                Next i
            End If
        
        Case WM_MOUSEMOVE
            If (m_AutoTheme <> "None") And (m_MsgBox.Theme <> mbClassic) Then
                '   We handle this as a "Before" proc Event to prevent painting by the
                '   default handler and showing the old Win9x button style....
                For i = 0 To m_MsgBox.Count - 1
                    If lng_hWnd = (m_MsgBox.Button(i).hWnd) Then
                        If (m_MsgBox.DownIndex = -1) Then
                            '   There is no button Pressed, Just hovering
                            '   so paint it hot
                            If (m_MsgBox.Button(i).State <> StateHot) Then
                                Call DrawWinXPButton(m_MsgBox.Button(i).hWnd, StateHot)
                                m_MsgBox.Button(i).State = StateHot
                            End If
                        Else
                            '   There is a button down and moving....keep it pressed
                            If i = m_MsgBox.DownIndex Then
                                If (m_MsgBox.Button(i).State <> StatePressed) Then
                                    Call DrawWinXPButton(m_MsgBox.Button(i).hWnd, StatePressed)
                                End If
                            End If
                        End If
                    Else
                        If i = m_MsgBox.DefaultIndex Then
                            '   Paint the Focused colors
                            If m_MsgBox.Button(i).State <> StateDefaulted Then
                                Call DrawWinXPButton(m_MsgBox.Button(i).hWnd, StateDefaulted)
                                m_MsgBox.Button(i).State = StateDefaulted
                            End If
                        ElseIf (m_MsgBox.DownIndex = i) Then
                            '   There is a down button moving
                            If (m_MsgBox.Button(i).State <> StatePressed) Then
                                Call DrawWinXPButton(m_MsgBox.Button(i).hWnd, StatePressed)
                            End If
                        Else
                            'Normal...not focused, hot or down
                            If m_MsgBox.Button(i).State <> StateNormal Then
                                Call DrawWinXPButton(m_MsgBox.Button(i).hWnd, StateNormal)
                                m_MsgBox.Button(i).State = StateNormal
                            End If
                        End If
                    End If
                Next i
            Else
                
            End If
            '   Tell the window we will handle the event and pass back a null pointer
            '   to indicate that we are done with this proc event...
            bHandled = True
            lReturn = 0
            
        Case WM_PAINT
            '   Get the Theme Info....if needed
            If Len(m_AutoTheme) = 0 Then
                '   Fail safe to make sure we have a valid theme
                m_AutoTheme = GetThemeInfo
            End If
            '   If this XP then Paint the buttons, or if Classic do nothing...
            If (m_AutoTheme <> "None") And (m_MsgBox.Theme <> mbClassic) Then
                For i = 0 To m_MsgBox.Count - 1
                    If i = m_MsgBox.DefaultIndex Then
                        Call DrawWinXPButton(m_MsgBox.Button(i).hWnd, StateDefaulted)
                    Else
                        Call DrawWinXPButton(m_MsgBox.Button(i).hWnd, m_MsgBox.Button(i).State)
                    End If
                Next i
            Else
                For i = 0 To m_MsgBox.Count - 1
                    Call DrawWin9xButton(m_MsgBox.Button(i).hWnd, m_MsgBox.Button(i).State)
                Next i
            End If
            RaiseEvent Status("MsgBox Painted")
                
        Case WM_SETFOCUS
            If (m_AutoTheme <> "None") And (m_MsgBox.Theme <> mbClassic) Then
                For i = 0 To m_MsgBox.Count - 1
                    If lng_hWnd = (m_MsgBox.Button(i).hWnd) Then
                        If m_MsgBox.Button(i).State <> StateDefaulted Then
                            Call DrawWinXPButton(m_MsgBox.Button(i).hWnd, StateDefaulted)
                            m_MsgBox.Button(i).State = StateDefaulted
                            If m_MsgBox.DefaultIndex = -1 Then
                                m_MsgBox.DefaultIndex = i
                            End If
                        End If
                    Else
                        Call DrawWinXPButton(m_MsgBox.Button(i).hWnd, StateNormal)
                        m_MsgBox.Button(i).State = StateNormal
                    End If
                Next i
            Else
                For i = 0 To m_MsgBox.Count - 1
                    If lng_hWnd = (m_MsgBox.Button(i).hWnd) Then
                        Call DrawWin9xButton(m_MsgBox.Button(i).hWnd, StateDefaulted)
                        m_MsgBox.Button(i).State = StateDefaulted
                    End If
                Next i
            End If
            RaiseEvent Status("Focus Changed")
                    
        Case WM_SYSCOLORCHANGE
            Refresh
            RaiseEvent Status("SysColor Changed")
            
        Case WM_TIMER
            Call TimerProc
                        
        Case WM_THEMECHANGED
            Refresh
            RaiseEvent Status("Theme Changed")
            
    End Select
    
End Sub

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

    'Add a message to the table of those that will invoke a callback. You should Subclass_Subclass first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    'Parameters:
        'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
        'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
        'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If When And eMsgWhen.MSG_AFTER Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    'Parameters:
    'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
    'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
    'When      - Whether the msg is to be removed from the before, after or both callback tables
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If When And eMsgWhen.MSG_AFTER Then
            Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
    '   Determine if the passed function is supported
                                    
    On Error GoTo IsFunctionExported_Error
    
    Dim hmod        As Long
    Dim bLibLoaded  As Boolean
    
    hmod = GetModuleHandleA(sModule)
    
    If hmod = 0 Then
        hmod = LoadLibraryA(sModule)
        If hmod Then
            bLibLoaded = True
        End If
    End If
    
    If hmod Then
        If GetProcAddress(hmod, sFunction) Then
            IsFunctionExported = True
        End If
    End If
    
    If bLibLoaded Then
        Call FreeLibrary(hmod)
    End If
    
    Exit Function

IsFunctionExported_Error:
End Function

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
    'Parameters:
    'lng_hWnd  - The handle of the window to be subclassed
    'Returns;
    'The sc_aSubData() index
    Const CODE_LEN              As Long = 204                                       'Length of the machine code in bytes
    Const FUNC_CWP              As String = "CallWindowProcA"                       'We use CallWindowProc to call the original WndProc
    Const FUNC_EBM              As String = "EbMode"                                'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
    Const FUNC_SWL              As String = "SetWindowLongA"                        'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
    Const MOD_USER              As String = "user32"                                'Location of the SetWindowLongA & CallWindowProc functions
    Const MOD_VBA5              As String = "vba5"                                  'Location of the EbMode function if running VB5
    Const MOD_VBA6              As String = "vba6"                                  'Location of the EbMode function if running VB6
    Const PATCH_01              As Long = 18                                        'Code buffer offset to the location of the relative address to EbMode
    Const PATCH_02              As Long = 68                                        'Address of the previous WndProc
    Const PATCH_03              As Long = 78                                        'Relative address of SetWindowsLong
    Const PATCH_06              As Long = 116                                       'Address of the previous WndProc
    Const PATCH_07              As Long = 121                                       'Relative address of CallWindowProc
    Const PATCH_0A              As Long = 186                                       'Address of the owner object
    Static aBuf(1 To CODE_LEN)  As Byte                                             'Static code buffer byte array
    Static pCWP                 As Long                                             'Address of the CallWindowsProc
    Static pEbMode              As Long                                             'Address of the EbMode IDE break/stop/running function
    Static pSWL                 As Long                                             'Address of the SetWindowsLong function
    Dim i                       As Long                                             'Loop index
    Dim j                       As Long                                             'Loop index
    Dim nSubIdx                 As Long                                             'Subclass data index
    Dim sHex                    As String                                           'Hex code string
    
    'If it's the first time through here..
    If aBuf(1) = 0 Then
        
        'The hex pair machine code representation.
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
            "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
            "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
            "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
        
        'Convert the string from hex pairs to bytes and store in the static machine code buffer
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            aBuf(j) = Val("&H" & Mid$(sHex, i, 2))                                  'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
            i = i + 2
        Loop                                                                        'Next pair of hex characters
        
        'Get API function addresses
        If Subclass_InIDE Then                                                      'If we're running in the VB IDE
            aBuf(16) = &H90                                                         'Patch the code buffer to enable the IDE state code
            aBuf(17) = &H90                                                         'Patch the code buffer to enable the IDE state code
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                 'Get the address of EbMode in vba6.dll
            If pEbMode = 0 Then                                                     'Found?
                pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                             'VB5 perhaps
            End If
        End If

        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                        'Get the address of the CallWindowsProc function
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                        'Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData                                       'Create the first sc_aSubData element
    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If nSubIdx = -1 Then                                                        'If an sc_aSubData element isn't being re-cycled
            nSubIdx = UBound(sc_aSubData()) + 1                                     'Calculate the next element
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                    'Create a new sc_aSubData element
        End If
        Subclass_Start = nSubIdx
    End If
    
    '   Use the following debuging to indicate which index into the
    '   sc_aSubData array the AddressOf Pointer exists at....
    'Debug.Print "AddressOf Index: " & nSubIdx
    With sc_aSubData(nSubIdx)
        .hWnd = lng_hWnd                                                            'Store the hWnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                               'Allocate memory for the machine code WndProc
        .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)                  'Set our WndProc in place
        Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                      'Copy the machine code from the static byte array to the code array in sc_aSubData
        Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                             'Original WndProc address for CallWindowProc, call the original WndProc
        Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                   'Patch the relative address of the SetWindowLongA api function
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                             'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
        Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                   'Patch the relative address of the CallWindowProc api function
        Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                             'Patch the address of this object instance into the static machine code buffer
    End With
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
    Dim i As Long
    
    i = UBound(sc_aSubData())                                                       'Get the upper bound of the subclass data array
    Do While i >= 0                                                                 'Iterate through each element
        With sc_aSubData(i)
            If .hWnd <> 0 Then                                                      'If not previously Subclass_Stop'd
                Call Subclass_Stop(.hWnd)                                           'Subclass_Stop
            End If
        End With
        i = i - 1                                                                   'Next element
    Loop
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
    'Parameters:
    'lng_hWnd  - The handle of the window to stop being subclassed
    With sc_aSubData(zIdx(lng_hWnd))
        Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)                         'Restore the original WndProc
        Call zPatchVal(.nAddrSub, PATCH_05, 0)                                      'Patch the Table B entry count to ensure no further 'before' callbacks
        Call zPatchVal(.nAddrSub, PATCH_09, 0)                                      'Patch the Table A entry count to ensure no further 'after' callbacks
        Call GlobalFree(.nAddrSub)                                                  'Release the machine code memory
        .hWnd = 0                                                                   'Mark the sc_aSubData element as available for re-use
        .nMsgCntB = 0                                                               'Clear the before table
        .nMsgCntA = 0                                                               'Clear the after table
        Erase .aMsgTblB                                                             'Erase the before table
        Erase .aMsgTblA                                                             'Erase the after table
    End With
End Sub

Private Function TimerProc()
    Dim i As Long
    Dim sPrompt As String
    Dim bDefault As Boolean
    'Static lTimer As Long
    
    On Error Resume Next
    With m_MsgBox
        '   This is where we will check to see if the dialog needs to close
        If .hWnd Then
            If (lCount < .TimerInfo.TimerDuration) Then
                '   Increment the count
                lCount = (lCount + 1)
                '   Reset our Flag
                bTimerRunning = True
                '   Update the Prompt Info with the time left
                sPrompt = Replace$(.Prompt.Caption, m_MsgBox.TimerInfo.TimerToken, CStr((m_MsgBox.TimerInfo.TimerDuration - lCount)))
                '   Set the Dialog Prompt
                SetDlgItemText m_MsgBox.hWnd, IDPROMPT, sPrompt
                Debug.Print sPrompt
            Else
                '   Stop the timer
                KillTimer UserControl.hWnd, .TimerInfo.TimerID
                '   Reset our Flag
                bTimerRunning = False
                '   Reset our counter
                lCount = 0
                For i = 0 To .Count - 1
                    If .Button(i).State = StateDefaulted Then
                        bDefault = True
                        Exit For
                    End If
                Next
                If Not bDefault Then
                    i = 0
                End If
                '   Set the Focused Button
                Call PutFocus(.Button(i).hWnd)
                '   DoEvents to allow PutFocus to actually put focus
                DoEvents
                '   Now Push the Button via Code ;-)
                Call PostMessage(.Button(i).hWnd, WM_LBUTTONDOWN, 0, ByVal 0&)
                Call PostMessage(.Button(i).hWnd, WM_LBUTTONUP, 0, ByVal 0&)
            End If
        End If
    End With
End Function

'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
  Dim tme As TRACKMOUSEEVENT_STRUCT
  
    If bTrack Then
        With tme
            .cbSize = Len(tme)
            .dwFlags = TME_LEAVE
            .hwndTrack = lng_hWnd
        End With
    
        If bTrackUser32 Then
            Call TrackMouseEvent(tme)
        Else
            Call TrackMouseEventComCtl(tme)
        End If
    End If
End Sub

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for sc_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry  As Long                                                             'Message table entry index
    Dim nOff1   As Long                                                             'Machine code buffer offset 1
    Dim nOff2   As Long                                                             'Machine code buffer offset 2
    
    If uMsg = ALL_MESSAGES Then                                                     'If all messages
        nMsgCnt = ALL_MESSAGES                                                      'Indicates that all messages will callback
    Else                                                                            'Else a specific message number
        Do While nEntry < nMsgCnt                                                   'For each existing entry. NB will skip if nMsgCnt = 0
            nEntry = nEntry + 1
            
            If aMsgTbl(nEntry) = 0 Then                                             'This msg table slot is a deleted entry
                aMsgTbl(nEntry) = uMsg                                              'Re-use this entry
                Exit Sub                                                            'Bail
            ElseIf aMsgTbl(nEntry) = uMsg Then                                      'The msg is already in the table!
                Exit Sub                                                            'Bail
            End If
        Loop                                                                        'Next entry
        nMsgCnt = nMsgCnt + 1                                                       'New slot required, bump the table entry count
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                'Bump the size of the table.
        aMsgTbl(nMsgCnt) = uMsg                                                     'Store the message number in the table
    End If

    If When = eMsgWhen.MSG_BEFORE Then                                              'If before
        nOff1 = PATCH_04                                                            'Offset to the Before table
        nOff2 = PATCH_05                                                            'Offset to the Before table entry count
    Else                                                                            'Else after
        nOff1 = PATCH_08                                                            'Offset to the After table
        nOff2 = PATCH_09                                                            'Offset to the After table entry count
    End If

    If uMsg <> ALL_MESSAGES Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                            'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)                                           'Patch the appropriate table entry count
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc                                                          'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for sc_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry As Long
    
    If uMsg = ALL_MESSAGES Then                                                     'If deleting all messages
        nMsgCnt = 0                                                                 'Message count is now zero
        If When = eMsgWhen.MSG_BEFORE Then                                          'If before
            nEntry = PATCH_05                                                       'Patch the before table message count location
        Else                                                                        'Else after
            nEntry = PATCH_09                                                       'Patch the after table message count location
        End If
        Call zPatchVal(nAddr, nEntry, 0)                                            'Patch the table message count to zero
    Else                                                                            'Else deleteting a specific message
        Do While nEntry < nMsgCnt                                                   'For each table entry
            nEntry = nEntry + 1
            If aMsgTbl(nEntry) = uMsg Then                                          'If this entry is the message we wish to delete
                aMsgTbl(nEntry) = 0                                                 'Mark the table slot as available
                Exit Do                                                             'Bail
            End If
        Loop                                                                        'Next entry
    End If
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
    'Get the upper bound of sc_aSubData() - If you get an error here, you're probably sc_AddMsg-ing before Subclass_Start
    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0                                                              'Iterate through the existing sc_aSubData() elements
        With sc_aSubData(zIdx)
            If .hWnd = lng_hWnd Then                                                'If the hWnd of this element is the one we're looking for
                If Not bAdd Then                                                    'If we're searching not adding
                    Exit Function                                                   'Found
                End If
            ElseIf .hWnd = 0 Then                                                   'If this an element marked for reuse.
                If bAdd Then                                                        'If we're adding
                    Exit Function                                                   'Re-use it
                End If
            End If
        End With
    zIdx = zIdx - 1                                                                 'Decrement the index
    Loop
    
    If Not bAdd Then
        '   Never, Ever use this in a modal dialog or your system will hang!!!
        'Debug.Assert False     'hWnd not found, programmer error
        '   Instead, we need a way to get out gracefully, so stop everything
        '   and continue on processing the requests....
        Call Subclass_StopAll
        Debug.Print "Sublcassing Error....No Handle Located!!!"
    End If

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True
End Function
'======================================================================================================
'   End SubClass Sections
'======================================================================================================

Public Property Get Alignment() As mbAlignEnum

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Alignment = m_MsgBox.Alignment
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.Alignment", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let Alignment(ByVal New_Value As mbAlignEnum)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_MsgBox.Alignment = New_Value
    PropertyChanged "Alignment"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.Alignment", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get AlignPrompt() As mbPromptAlignEnum

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    AlignPrompt = m_MsgBox.Prompt.Alignment
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.AlignPrompt", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let AlignPrompt(ByVal New_Value As mbPromptAlignEnum)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_MsgBox.Prompt.Alignment = New_Value
    PropertyChanged "AlignPrompt"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.AlignPrompt", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

'draw a Line Using API call's
Private Sub APILine(X1 As Long, _
    Y1 As Long, _
    X2 As Long, _
    Y2 As Long, _
    lColor As Long, _
    Optional lhDC As Long)
    'Use the API LineTo for Fast Drawing
    
    Dim Pt As POINT
    Dim hPen As Long, hPenOld As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    If lhDC = 0 Then lhDC = UserControl.hdc
    
    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(lhDC, hPen)
    MoveToEx lhDC, X1, Y1, Pt
    LineTo lhDC, X2, Y2
    SelectObject lhDC, hPenOld
    DeleteObject hPen
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.APILine", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

' full version of APILine
Private Sub APILineEx(lhdcEx As Long, _
    X1 As Long, _
    Y1 As Long, _
    X2 As Long, _
    Y2 As Long, _
    lColor As Long)
    
    'Use the API LineTo for Fast Drawing
    Dim Pt As POINT
    Dim hPen As Long, hPenOld As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(lhdcEx, hPen)
    MoveToEx lhdcEx, X1, Y1, Pt
    LineTo lhdcEx, X2, Y2
    SelectObject lhdcEx, hPenOld
    DeleteObject hPen
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucButton.APILineEx", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Function APIRectangle(ByVal hdc As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal W As Long, _
    ByVal h As Long, _
    Optional lColor As OLE_COLOR = -1, _
    Optional lhDC As Long) As Long
    
    Dim hPen As Long, hPenOld As Long
    Dim Pt As POINT
    
    '   Handle Any Errors
    On Error GoTo Func_ErrHandler

    If lhDC = 0 Then lhDC = UserControl.hdc
    
    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(lhDC, hPen)
    MoveToEx lhDC, X, Y, Pt
    LineTo lhDC, X + W, Y
    LineTo lhDC, X + W, Y + h
    LineTo lhDC, X, Y + h
    LineTo lhDC, X, Y
    SelectObject lhDC, hPenOld
    DeleteObject hPen
    
Func_ErrHandlerExit:
    Exit Function
Func_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.APIRectangle", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Func_ErrHandlerExit:
End Function

Public Property Get BackColor() As OLE_COLOR
    
    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler
    
    BackColor = m_MsgBox.BackColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.BackColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    
    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler
    
    m_MsgBox.BackColor = NewValue
    PropertyChanged "BackColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.BackColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get Caption(ByVal Index As Long) As String
    '   Keep the range reasonable....
    If Index < 0 Then Index = 0
    If Index > 3 Then Index = 3
    '   Get the array data
    Caption = m_MsgBox.CustomCaption(Index)
End Property

Public Property Let Caption(ByVal Index As Long, NewValue As String)
    '   Keep the range reasonable....
    If Index < 0 Then Index = 0
    If Index > 3 Then Index = 3
    '   Fill the array with the data
    m_MsgBox.CustomCaption(Index) = NewValue
    PropertyChanged "Caption" & Index
End Property

Public Property Get CaptionType() As mbCaptionEnum
    CaptionType = m_MsgBox.CaptionType
End Property

Public Property Let CaptionType(ByVal NewValue As mbCaptionEnum)
    m_MsgBox.CaptionType = NewValue
    PropertyChanged "CaptionType"
End Property

Private Sub DrawFilledRect(ByVal lhDC As Long, lpRect As RECT, ByVal lColor As Long)
    Dim hBrush As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    lColor = TranslateColor(lColor)
    hBrush = CreateSolidBrush(lColor)
    InflateRect lpRect, 1, 0
    FillRect lhDC, lpRect, hBrush
    DeleteObject hBrush
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.DrawFilledRect", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub DrawVGradient(lEndColor As Long, _
    lStartcolor As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long, _
    Optional lhDC As Long)
    ''Draw a Vertical Gradient in the current HDC
    
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    If lhDC = 0 Then lhDC = UserControl.hdc
    
    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    dR = (sR - eR) / Y2
    dG = (sG - eG) / Y2
    dB = (sB - eB) / Y2
    
    For ni = 0 To Y2
        APILine X, Y + ni, X2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB)), lhDC
    Next 'ni
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.DrawVGradient", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub DrawVGradientEx(lEndColor As Long, _
    lStartcolor As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long, _
    Optional ByVal lhDC As Long = -1)
    '   Draw a Vertical Gradient in the current HDC
    
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    dR = (sR - eR) / Y2
    dG = (sG - eG) / Y2
    dB = (sB - eB) / Y2
    
    If lhDC <> (-1) Then
        For ni = 0 To Y2 - Y
            APILineEx lhDC, X, Y + ni, X2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
        Next
    Else
        For ni = 0 To Y2 - Y
            APILineEx UserControl.hdc, X, Y + ni, X2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
        Next
    End If
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.DrawVGradientEx", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub DrawWin9xButton(ByVal lhWnd As Long, ByVal Mode As mbState)
    '   This Sub Draws the Win9x Style Button
    Dim lhDC As Long
    Dim tempColor As Long
    Dim lH As Long
    Dim lW As Long
    Dim lpRect As RECT
    Dim lpTmpRect As RECT
    Dim i As Long
    Dim tPt As POINT
    Dim lBkColor As Long
    Dim State As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    '   Get the Window RECT
    Call GetWindowRect(lhWnd, lpRect)
    tPt.X = lpRect.Right
    tPt.Y = lpRect.Bottom
    '   Convert these to Pixels
    ScreenToClient lhWnd, tPt
    '   Now store them for use
    lW = tPt.X
    lH = tPt.Y
    '   Get the DC for the this window
    lhDC = GetDC(lhWnd)
    '   Set a Temporary RECT
    SetRect lpTmpRect, 0, 0, lW, lH
    '   Set the BackColor...
    lBkColor = TranslateColor(m_MsgBox.BackColor)
    '   Set the Back Color based on the Form BackColor
    SetBkColor lhDC, lBkColor
    '   Start By Drawing the Backcolor over the OldButton
    DrawFilledRect lhDC, lpTmpRect, lBkColor
    '   Now Draw the New Button with our New Color
    '   Note: The DrawEdge or DrawFrameControl do not seem to yeild
    '         drawings which are the same as the native Win9x button
    '         so we will paint these the hard way....
DrawNormal:
    Select Case Mode
        Case StateDefaulted
            '   Outer Edges (ALL)
            tempColor = TranslateColor(vbBlack)
            APILine 0, 0, lW, 0, tempColor, lhDC
            APILine lW, 0, lW, lH, tempColor, lhDC
            APILine lW, lH, 0, lH, tempColor, lhDC
            APILine 0, lH, 0, 0, tempColor, lhDC
            '   Inner Edge Highlighted (TOP / LEFT)
            tempColor = TranslateColor(vb3DHighlight)
            APILine 1, 1, lW - 1, 1, tempColor, lhDC
            APILine 1, 2, lW - 1, 2, tempColor, lhDC
            APILine 1, 1, 1, lH - 1, tempColor, lhDC
            APILine 2, 1, 2, lH - 1, tempColor, lhDC
            '   Inner Edge Light Shadow (RIGHT)
            tempColor = TranslateColor(vb3DShadow)
            APILine 2, lH - 2, lW - 1, lH - 2, tempColor, lhDC
            APILine lW - 2, 2, lW - 2, lH - 1, tempColor, lhDC
            '   Inner Edge Dark Shadow (RIGHT)
            tempColor = TranslateColor(vb3DDKShadow)
            APILine 1, lH - 1, lW, lH - 1, tempColor, lhDC
            APILine lW - 1, 1, lW - 1, lH - 1, tempColor, lhDC
        Case StateNormal
            '   Inner Edge Highlighted (TOP / LEFT)
            tempColor = TranslateColor(vb3DHighlight)
            APILine 0, 0, lW, 0, tempColor, lhDC
            APILine 0, 1, lW, 1, tempColor, lhDC
            APILine 0, 0, 0, lH, tempColor, lhDC
            APILine 1, 0, 1, lH, tempColor, lhDC
            '   Inner Edge Light Shadow (RIGHT)
            tempColor = TranslateColor(vb3DShadow)
            APILine 1, lH - 1, lW, lH - 1, tempColor, lhDC
            APILine lW - 1, 1, lW - 1, lH, tempColor, lhDC
            '   Inner Edge Dark Shadow (RIGHT)
            tempColor = TranslateColor(vb3DDKShadow)
            APILine 0, lH, lW + 1, lH, tempColor, lhDC
            APILine lW, 0, lW, lH, tempColor, lhDC
        Case StatePressed
            '   Outer Edges (ALL)
            tempColor = TranslateColor(vbBlack)
            APILine 0, 0, lW, 0, tempColor, lhDC
            APILine lW, 0, lW, lH, tempColor, lhDC
            APILine lW, lH, 0, lH, tempColor, lhDC
            APILine 0, lH, 0, 0, tempColor, lhDC
            '   Inner Edges Light Shadow (ALL)
            tempColor = TranslateColor(vb3DShadow)
            APILine 1, 1, lW - 1, 1, tempColor, lhDC
            APILine lW - 1, 1, lW - 1, lH - 1, tempColor, lhDC
            APILine lW - 1, lH - 1, 1, lH - 1, tempColor, lhDC
            APILine 1, lH - 1, 1, 1, tempColor, lhDC
        Case Else
            '   Something went wrong...so paint it as a normal button
            Mode = StateNormal
            GoTo DrawNormal:
    End Select
    
    '   Set the BackMode to Transparent
    SetBkMode lhDC, 0
    '   See if the font has been set, if not then use a default
    If Len(m_MsgBox.Font.Name) = 0 Then
        '   Select the Correct Font....as close as we can if missing ;-)
        SelectFont lhDC, 8, False, "Tahoma", False
    Else
        '   Select the Passed Font set by the Developer
        SelectFont lhDC, m_MsgBox.Font.Size, m_MsgBox.Font.Italic, m_MsgBox.Font.Name, m_MsgBox.Font.Underline
    End If
    '   Set the ForeColor
    SetTextColor lhDC, TranslateColor(m_MsgBox.Prompt.ForeColor)
    '   Now set the Button Rects
    If Mode <> StatePressed Then
        SetRect lpRect, 4, 4, 72, 21
    Else
        SetRect lpRect, 5, 5, 73, 22
    End If
    '   Find out Button and Place the Caption
    For i = 0 To m_MsgBox.Count - 1
        If m_MsgBox.Button(i).hWnd = lhWnd Then
            '   If one wanted we could pass any caption to these buttons as we are
            '   repainting the text ourselves
            If m_MsgBox.CaptionType = mbDefault Then
                DrawText lhDC, m_MsgBox.Button(i).Caption, LStrLen(m_MsgBox.Button(i).Caption), lpRect, DT_CENTER Or DT_VCENTER
            Else
                DrawText lhDC, m_MsgBox.CustomCaption(i), LStrLen(m_MsgBox.CustomCaption(i)), lpRect, DT_CENTER Or DT_VCENTER
            End If
        End If
    Next
    '   Free the Resources
    ReleaseDC lhWnd, lhDC
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.DrawWin9xButton", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub DrawWinXPButton(lhWnd As Long, Mode As mbState)
    '' This Sub Draws the XPStyle Button
    Dim lhDC As Long
    Dim tempColor As Long
    Dim lH As Long
    Dim lW As Long
    Dim lpRect As RECT
    Dim lpTmpRect As RECT
    Dim i As Long
    Dim tPt As POINT
    Dim lBkColor As Long
    Dim AutoTheme As String
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    '   Get the Window RECT
    Call GetWindowRect(lhWnd, lpRect)
    tPt.X = lpRect.Right
    tPt.Y = lpRect.Bottom
    '   Convert these to Pixels
    ScreenToClient lhWnd, tPt
    '   Now store them for use
    lW = tPt.X
    lH = tPt.Y
    '   Get the DC for this window
    lhDC = GetDC(lhWnd)
    '   Set a Temporary RECT
    SetRect lpTmpRect, 0, 0, lW, lH
    '   Get the Current BackColor of the Buttons
    lBkColor = GetBkColor(lhDC)
    If Not lBkColor Then lBkColor = TranslateColor(vbButtonFace)
    '   Start By Drawing the Backcolor over the OldButton
    DrawFilledRect lhDC, lpTmpRect, lBkColor
    
    Select Case Mode
        Case StateNormal, StateHot, StateDefaulted
            'Main
            Select Case m_MsgBox.Theme
                Case mbBlue
                    DrawVGradient &HFBFCFC, &HF0F0F0, 1, 1, lW - 2, 4, lhDC
                    DrawVGradient &HF9FAFA, &HEAF0F0, 1, 4, lW - 2, lH - 8, lhDC
                    DrawVGradient &HE6EBEB, &HC5D0D6, 1, lH - 4, lW - 2, 3, lhDC
                    'right
                    DrawVGradient &HFAFBFB, &HDAE2E4, lW - 3, 3, lW - 2, lH - 5, lhDC
                    DrawVGradient &HF2F4F5, &HCDD7DB, lW - 2, 3, lW - 1, lH - 5, lhDC
                    'Border
                    APILine 1, 0, lW - 1, 0, &H743C00, lhDC
                    APILine 0, 1, 0, lH - 1, &H743C00, lhDC
                    APILine lW - 1, 1, lW - 1, lH - 1, &H743C00, lhDC
                    APILine 1, lH - 1, lW - 1, lH - 1, &H743C00, lhDC
                    'Corners
                    SetPixelV lhDC, 0, 0, lBkColor
                    SetPixelV lhDC, lW - 1, 0, lBkColor
                    SetPixelV lhDC, lW - 1, lH - 1, lBkColor
                    SetPixelV lhDC, 0, lH - 1, lBkColor
                    
                    SetPixelV lhDC, 1, 1, &H906E48
                    SetPixelV lhDC, 1, lH - 2, &H906E48
                    SetPixelV lhDC, lW - 2, 1, &H906E48
                    SetPixelV lhDC, lW - 2, lH - 2, &H906E48
                    'External Borders
                    SetPixelV lhDC, 0, 1, &HA28B6A
                    SetPixelV lhDC, 1, 0, &HA28B6A
                    SetPixelV lhDC, 1, lH - 1, &HA28B6A
                    SetPixelV lhDC, 0, lH - 2, &HA28B6A
                    SetPixelV lhDC, lW - 1, lH - 2, &HA28B6A
                    SetPixelV lhDC, lW - 2, lH - 1, &HA28B6A
                    SetPixelV lhDC, lW - 2, 0, &HA28B6A
                    SetPixelV lhDC, lW - 1, 1, &HA28B6A
                    'Internal Soft
                    SetPixelV lhDC, 1, 2, &HCAC7BF
                    SetPixelV lhDC, 2, 1, &HCAC7BF
                    SetPixelV lhDC, 2, lH - 2, &HCAC7BF
                    SetPixelV lhDC, 1, lH - 3, &HCAC7BF
                    SetPixelV lhDC, lW - 2, lH - 3, &HCAC7BF
                    SetPixelV lhDC, lW - 3, lH - 2, &HCAC7BF
                    SetPixelV lhDC, lW - 3, 1, &HCAC7BF
                    SetPixelV lhDC, lW - 2, 2, &HCAC7BF
                Case mbHomeStead
                    DrawVGradient &HF6FFFF, &HEFFDFE, 1, 1, lW - 2, 4, lhDC
                    DrawVGradient &HEFFDFE, &HDBEEF3, 1, 4, lW - 2, lH - 8, lhDC
                    DrawVGradient &HDBEEF3, &HB8D1E3, 1, lH - 4, lW - 2, 3, lhDC
                    'right
                    DrawVGradient &HFAFBFB, &HDAE2E4, lW - 3, 3, lW - 2, lH - 5, lhDC
                    DrawVGradient &HF2F4F5, &HCDD7DB, lW - 2, 3, lW - 1, lH - 5, lhDC
                    'Border
                    APILine 1, 0, lW - 1, 0, &H66237, lhDC
                    APILine 0, 1, 0, lH - 1, &H66237, lhDC
                    APILine lW - 1, 1, lW - 1, lH - 1, &H66237, lhDC
                    APILine 1, lH - 1, lW - 1, lH - 1, &H66237, lhDC
                    'Corners
                    SetPixelV lhDC, 0, 0, lBkColor
                    SetPixelV lhDC, lW - 1, 0, lBkColor
                    SetPixelV lhDC, lW - 1, lH - 1, lBkColor
                    SetPixelV lhDC, 0, lH - 1, lBkColor
                    
                    SetPixelV lhDC, 1, 1, &H906E48
                    SetPixelV lhDC, 1, lH - 2, &H906E48
                    SetPixelV lhDC, lW - 2, 1, &H906E48
                    SetPixelV lhDC, lW - 2, lH - 2, &H906E48
                    'External Borders
                    SetPixelV lhDC, 0, 1, &H5B8975
                    SetPixelV lhDC, 1, 0, &H5B8975
                    SetPixelV lhDC, 1, lH - 1, &H5B8975
                    SetPixelV lhDC, 0, lH - 2, &H5B8975
                    SetPixelV lhDC, lW - 1, lH - 2, &H5B8975
                    SetPixelV lhDC, lW - 2, lH - 1, &H5B8975
                    SetPixelV lhDC, lW - 2, 0, &H5B8975
                    SetPixelV lhDC, lW - 1, 1, &H5B8975
                    'Internal Soft
                    SetPixelV lhDC, 1, 2, &HBEDAE4
                    SetPixelV lhDC, 2, 1, &HBEDAE4
                    SetPixelV lhDC, 2, lH - 2, &HBEDAE4
                    SetPixelV lhDC, 1, lH - 3, &HBEDAE4
                    SetPixelV lhDC, lW - 2, lH - 3, &HBEDAE4
                    SetPixelV lhDC, lW - 3, lH - 2, &HBEDAE4
                    SetPixelV lhDC, lW - 3, 1, &HBEDAE4
                    SetPixelV lhDC, lW - 2, 2, &HBEDAE4
                Case mbMetallic
                    '   Main Gradients
                    DrawVGradientEx &HFFFFFF, &HFFFFFF, 2, 1, lW - 2, lH * 0.35, lhDC
                    DrawVGradientEx &HFFFFFF, &HC57272, 2, lH * 0.35, lW - 2, lH - 2, lhDC
                    '   Edge Gradients
                    DrawVGradient &HFFFFFF, &HFFFFFF, 1, 2, 2, lH - 5, lhDC
                    DrawVGradient &HFFFFFF, &HFFFFFF, lW - 2, 2, lW - 1, lH - 5, lhDC
                    'Border
                    APILine 1, 0, lW - 1, 0, &H743C00, lhDC
                    APILine 0, 1, 0, lH - 1, &H743C00, lhDC
                    APILine lW - 1, 1, lW - 1, lH - 1, &H743C00, lhDC
                    APILine 1, lH - 1, lW - 1, lH - 1, &H743C00, lhDC
                    '   Corner Colors
                    SetPixelV lhDC, 1, 1, &H906E48
                    SetPixelV lhDC, 1, lH - 2, &H906E48
                    SetPixelV lhDC, lW - 2, 1, &H906E48
                    SetPixelV lhDC, lW - 2, lH - 2, &H906E48
                    'External Borders
                    SetPixelV lhDC, 0, 1, &HA28B6A
                    SetPixelV lhDC, 1, 0, &HA28B6A
                    SetPixelV lhDC, 1, lH - 1, &HA28B6A
                    SetPixelV lhDC, 0, lH - 2, &HA28B6A
                    SetPixelV lhDC, lW - 1, lH - 2, &HA28B6A
                    SetPixelV lhDC, lW - 2, lH - 1, &HA28B6A
                    SetPixelV lhDC, lW - 2, 0, &HA28B6A
                    SetPixelV lhDC, lW - 1, 1, &HA28B6A
                    'Internal Soft
                    SetPixelV lhDC, 1, 2, &HCAC7BF
                    SetPixelV lhDC, 2, 1, &HCAC7BF
                    SetPixelV lhDC, 2, lH - 2, &HCAC7BF
                    SetPixelV lhDC, 1, lH - 3, &HCAC7BF
                    SetPixelV lhDC, lW - 2, lH - 3, &HCAC7BF
                    SetPixelV lhDC, lW - 3, lH - 2, &HCAC7BF
                    SetPixelV lhDC, lW - 3, 1, &HCAC7BF
                    SetPixelV lhDC, lW - 2, 2, &HCAC7BF
            End Select
            
            If Mode = StateHot Then
                Select Case m_MsgBox.Theme
                    Case mbBlue
                        DrawVGradient &H89D8FD, &H30B3F8, 1, 2, 3, lH - 5, lhDC
                        DrawVGradient &H89D8FD, &H30B3F8, lW - 3, 2, lW - 1, lH - 5, lhDC
                        APILine 2, 1, lW - 2, 1, &HCFF0FF, lhDC
                        APILine 2, 2, lW - 2, 2, &H89D8FD, lhDC
                        APILine 2, lH - 3, lW - 2, lH - 3, &H30B3F8, lhDC
                        APILine 2, lH - 2, lW - 2, lH - 2, &H1097E5, lhDC
                    Case mbHomeStead
                        DrawVGradient &H8BB8EB, &H5291E3, 1, 2, 3, lH - 5, lhDC
                        DrawVGradient &H8BB8EB, &H5291E3, lW - 3, 2, lW - 1, lH - 5, lhDC
                        APILine 2, 1, lW - 2, 1, &H95C5FC, lhDC
                        APILine 2, 2, lW - 2, 2, &H96BEED, lhDC
                        APILine 2, lH - 3, lW - 2, lH - 3, &H4E90E3, lhDC
                        APILine 2, lH - 2, lW - 2, lH - 2, &H2572CF, lhDC
                    Case mbMetallic
                        '   Main Gradients
                        DrawVGradientEx &HFFFFFF, &HFFFFFF, 2, 1, lW - 2, lH * 0.45, lhDC
                        DrawVGradientEx &HFFFFFF, &HC57777, 2, lH * 0.45, lW - 2, lH - 2, lhDC
                        '   Edge Gradients
                        DrawVGradient &H9ADFFE, &H49B5EF, 1, 2, 2, lH - 5, lhDC
                        DrawVGradient &H79D2FC, &H35B5F8, 2, 2, 3, lH - 5, lhDC
                        DrawVGradient &H79D2FC, &H35B5F8, lW - 3, 2, lW - 2, lH - 5, lhDC
                        DrawVGradient &H9ADFFE, &H49B5EF, lW - 2, 2, lW - 1, lH - 5, lhDC
                        '   Hot Borders
                        APILine 2, 1, lW - 2, 1, &HCFF0FF, lhDC
                        APILine 2, 2, lW - 2, 2, &H89D8FD, lhDC
                        APILine 2, lH - 3, lW - 2, lH - 3, &H2FB2F8, lhDC
                        APILine 2, lH - 2, lW - 2, lH - 2, &H1097E5, lhDC
                End Select
            ElseIf (Mode = StateDefaulted) Then
                Select Case m_MsgBox.Theme
                    Case mbBlue
                        DrawVGradient &HF6D4BC, &HE4AD89, 1, 2, 3, lH - 5, lhDC
                        DrawVGradient &HF6D4BC, &HE4AD89, lW - 3, 2, lW - 1, lH - 5, lhDC
                        APILine 2, 1, lW - 2, 1, &HFFE7CE, lhDC
                        APILine 2, 2, lW - 2, 2, &HF6D4BC, lhDC
                        APILine 2, lH - 3, lW - 2, lH - 3, &HE4AD89, lhDC
                        APILine 2, lH - 2, lW - 2, lH - 2, &HEE8269, lhDC
                    Case mbHomeStead
                        DrawVGradient &H54C190, &H7DCBB1, 1, 2, 3, lH - 5, lhDC
                        DrawVGradient &H54C190, &H7DCBB1, lW - 3, 2, lW - 1, lH - 5, lhDC
                        APILine 2, 1, lW - 2, 1, &H8FD1C2, lhDC
                        APILine 2, 2, lW - 2, 2, &H80CBB1, lhDC
                        APILine 2, lH - 3, lW - 2, lH - 3, &H54C190, lhDC
                        APILine 2, lH - 2, lW - 2, lH - 2, &H66A7A8, lhDC
                    Case mbMetallic
                        '   Main Gradient
                        DrawVGradientEx &HFFFFFF, &HFFFFFF, 2, 1, lW - 2, lH * 0.35, lhDC
                        DrawVGradientEx &HFFFFFF, &HC57272, 2, lH * 0.35, lW - 2, lH - 2, lhDC
                        '   Edge Gradients
                        DrawVGradient &HE4AD89, &HF5D3BA, 1, 2, 2, lH - 5, lhDC
                        DrawVGradient &HFFFFFF, &HFFFFFF, 2, 2, 3, lH - 5, lhDC
                        DrawVGradient &HFFFFFF, &HFFFFFF, lW - 3, 2, lW - 2, lH - 5, lhDC
                        DrawVGradient &HE4AD89, &HF5D3BA, lW - 2, 2, lW - 1, lH - 5, lhDC
                        APILine 2, 1, lW - 2, 1, &HFFE7CE, lhDC
                        APILine 2, 2, lW - 2, 2, &HF6D4BC, lhDC
                        APILine 2, lH - 3, lW - 2, lH - 3, &HE4AD89, lhDC
                        APILine 2, lH - 2, lW - 2, lH - 2, &HEE8269, lhDC
                End Select
            End If
        Case StatePressed
            Select Case m_MsgBox.Theme
                Case mbBlue
                    '   Main
                    DrawVGradient &HC1CCD1, &HDCE3E4, 2, 1, lW - 1, 4, lhDC
                    DrawVGradient &HDCE3E4, &HDBE2E3, 2, 4, lW - 1, lH - 8, lhDC
                    DrawVGradient &HDBE2E3, &HEEF1F2, 3, lH - 4, lW - 1, 3, lhDC
                    '   Left Edge
                    DrawVGradient &HCED8DA, &HDBE2E3, 1, 3, 2, lH - 5, lhDC
                    DrawVGradient &HCED8DA, &HDBE2E3, 2, 4, 3, lH - 7, lhDC
                    '   Border
                    APILine 1, 0, lW - 1, 0, &H743C00, lhDC
                    APILine 0, 1, 0, lH - 1, &H743C00, lhDC
                    APILine lW - 1, 1, lW - 1, lH - 1, &H743C00, lhDC
                    APILine 1, lH - 1, lW - 1, lH - 1, &H743C00, lhDC
                    '   Corners
                    SetPixelV lhDC, 1, 1, &H906E48
                    SetPixelV lhDC, 1, lH - 2, &H906E48
                    SetPixelV lhDC, lW - 2, 1, &H906E48
                    SetPixelV lhDC, lW - 2, lH - 2, &H906E48
                    '   External Borders
                    SetPixelV lhDC, 0, 1, &HA28B6A
                    SetPixelV lhDC, 1, 0, &HA28B6A
                    SetPixelV lhDC, 1, lH - 1, &HA28B6A
                    SetPixelV lhDC, 0, lH - 2, &HA28B6A
                    SetPixelV lhDC, lW - 1, lH - 2, &HA28B6A
                    SetPixelV lhDC, lW - 2, lH - 1, &HA28B6A
                    SetPixelV lhDC, lW - 2, 0, &HA28B6A
                    SetPixelV lhDC, lW - 1, 1, &HA28B6A
                    '   Internal Soft
                    SetPixelV lhDC, 1, 2, &HCAC7BF
                    SetPixelV lhDC, 2, 1, &HCAC7BF
                    SetPixelV lhDC, 2, lH - 2, &HCAC7BF
                    SetPixelV lhDC, 1, lH - 3, &HCAC7BF
                    SetPixelV lhDC, lW - 2, lH - 3, &HCAC7BF
                    SetPixelV lhDC, lW - 3, lH - 2, &HCAC7BF
                    SetPixelV lhDC, lW - 3, 1, &HCAC7BF
                    SetPixelV lhDC, lW - 2, 2, &HCAC7BF
                Case mbHomeStead
                    DrawVGradientEx &HCEE4EC, &HD2E6EE, 1, 1, lW - 1, lH - 2, lhDC
                    '   Left
                    DrawVGradient &HB0C4D4, &HBFD4E4, 1, 3, 2, lH - 5, lhDC
                    DrawVGradient &HB0C4D4, &HBFD4E4, 2, 4, 3, lH - 7, lhDC
                    '   Border
                    APILine 1, 0, lW - 1, 0, &H66237, lhDC
                    APILine 0, 1, 0, lH - 1, &H66237, lhDC
                    APILine lW - 1, 1, lW - 1, lH - 1, &H66237, lhDC
                    APILine 1, lH - 1, lW - 1, lH - 1, &H66237, lhDC
                    '   Corners
                    SetPixelV lhDC, 1, 1, &H906E48
                    SetPixelV lhDC, 1, lH - 2, &H906E48
                    SetPixelV lhDC, lW - 2, 1, &H906E48
                    SetPixelV lhDC, lW - 2, lH - 2, &H906E48
                    '   External Borders
                    SetPixelV lhDC, 0, 1, &H5B8975
                    SetPixelV lhDC, 1, 0, &H5B8975
                    SetPixelV lhDC, 1, lH - 1, &H5B8975
                    SetPixelV lhDC, 0, lH - 2, &H5B8975
                    SetPixelV lhDC, lW - 1, lH - 2, &H5B8975
                    SetPixelV lhDC, lW - 2, lH - 1, &H5B8975
                    SetPixelV lhDC, lW - 2, 0, &H5B8975
                    SetPixelV lhDC, lW - 1, 1, &H5B8975
                    '   Internal Soft
                    SetPixelV lhDC, 1, 2, &HBEDAE4
                    SetPixelV lhDC, 2, 1, &HBEDAE4
                    SetPixelV lhDC, 2, lH - 2, &HBEDAE4
                    SetPixelV lhDC, 1, lH - 3, &HBEDAE4
                    SetPixelV lhDC, lW - 2, lH - 3, &HBEDAE4
                    SetPixelV lhDC, lW - 3, lH - 2, &HBEDAE4
                    SetPixelV lhDC, lW - 3, 1, &HBEDAE4
                    SetPixelV lhDC, lW - 2, 2, &HBEDAE4
                Case mbMetallic
                    DrawVGradientEx &HB5B5B5, &HFDFDFD, 1, 1, lW - 1, lH - 5, lhDC
                    DrawVGradientEx &HFDFDFD, &HFFFFFF, 1, lH - 5, lW - 1, lH - 2, lhDC
                    '   Left
                    DrawVGradient &HCED8DA, &HDBE2E3, 1, 3, 2, lH - 5, lhDC
                    DrawVGradient &HCED8DA, &HDBE2E3, 2, 4, 3, lH - 7, lhDC
                    '   Border
                    APILine 1, 0, lW - 1, 0, &H743C00, lhDC
                    APILine 0, 1, 0, lH - 1, &H743C00, lhDC
                    APILine lW - 1, 1, lW - 1, lH - 1, &H743C00, lhDC
                    APILine 1, lH - 1, lW - 1, lH - 1, &H743C00, lhDC
                    '   Corners
                    SetPixelV lhDC, 1, 1, &H906E48
                    SetPixelV lhDC, 1, lH - 2, &H906E48
                    SetPixelV lhDC, lW - 2, 1, &H906E48
                    SetPixelV lhDC, lW - 2, lH - 2, &H906E48
                    '   External Borders
                    SetPixelV lhDC, 0, 1, &HA28B6A
                    SetPixelV lhDC, 1, 0, &HA28B6A
                    SetPixelV lhDC, 1, lH - 1, &HA28B6A
                    SetPixelV lhDC, 0, lH - 2, &HA28B6A
                    SetPixelV lhDC, lW - 1, lH - 2, &HA28B6A
                    SetPixelV lhDC, lW - 2, lH - 1, &HA28B6A
                    SetPixelV lhDC, lW - 2, 0, &HA28B6A
                    SetPixelV lhDC, lW - 1, 1, &HA28B6A
                    '   Internal Soft
                    SetPixelV lhDC, 1, 2, &HCAC7BF
                    SetPixelV lhDC, 2, 1, &HCAC7BF
                    SetPixelV lhDC, 2, lH - 2, &HCAC7BF
                    SetPixelV lhDC, 1, lH - 3, &HCAC7BF
                    SetPixelV lhDC, lW - 2, lH - 3, &HCAC7BF
                    SetPixelV lhDC, lW - 3, lH - 2, &HCAC7BF
                    SetPixelV lhDC, lW - 3, 1, &HCAC7BF
                    SetPixelV lhDC, lW - 2, 2, &HCAC7BF
            End Select
            AutoTheme = GetThemeInfo
            If (m_MsgBox.Theme = mbHomeStead) Or (AutoTheme = "HomeStead") Then
                SetRect lpRect, 2, 2, lW - 2, lH - 2
                DrawFocusRect lhDC, lpRect
            End If
        Case StateDisabled
            tempColor = &HEAF4F5
            UserControl.BackColor = tempColor
            lhDC = UserControl.hdc
            APIRectangle lhDC, 0, 0, lW - 1, lH - 1, &HBAC7C9
            tempColor = &HC7D5D8
            SetPixelV lhDC, 0, 1, tempColor
            SetPixelV lhDC, 1, 1, tempColor
            SetPixelV lhDC, 1, 0, tempColor
            SetPixelV lhDC, 0, lH - 2, tempColor
            SetPixelV lhDC, 1, lH - 2, tempColor
            SetPixelV lhDC, 1, lH - 1, tempColor
            SetPixelV lhDC, lW - 1, 1, tempColor
            SetPixelV lhDC, lW - 2, 1, tempColor
            SetPixelV lhDC, lW - 2, 0, tempColor
            SetPixelV lhDC, lW - 1, lH - 2, tempColor
            SetPixelV lhDC, lW - 2, lH - 2, tempColor
            SetPixelV lhDC, lW - 2, lH - 1, tempColor
    End Select

    '   Clean up Around Each Button
    tempColor = TranslateColor(m_MsgBox.BackColor)
    APILine -1, 0, -1, lH, tempColor, lhDC
    APILine lW, 0, lW, lH, tempColor, lhDC
    SetPixelV lhDC, 0, 0, tempColor
    SetPixelV lhDC, lW - 1, 0, tempColor
    SetPixelV lhDC, lW - 1, lH - 1, tempColor
    SetPixelV lhDC, 0, lH - 1, tempColor
    
    '   Set the BackMode to Transparent
    SetBkMode lhDC, 0
    
    If Len(m_MsgBox.Font.Name) = 0 Then
        '   Select the Correct Font....as close as we can if missing ;-)
        SelectFont lhDC, 8, False, "Tahoma", False
    Else
        '   Select the Passed Font set by the Developer
        SelectFont lhDC, m_MsgBox.Font.Size, m_MsgBox.Font.Italic, m_MsgBox.Font.Name, m_MsgBox.Font.Underline
    End If
    '   Set the ForeColor
    SetTextColor lhDC, TranslateColor(m_MsgBox.Prompt.ForeColor)
    '   Now set the Button Rects
    SetRect lpRect, 4, 4, 72, 21
    '   Find out Button and Place the Caption
    For i = 0 To m_MsgBox.Count - 1
        If m_MsgBox.Button(i).hWnd = lhWnd Then
            '   If one wanted we could pass any caption to these buttons as we are
            '   repainting the text ourselves
            If m_MsgBox.CaptionType = mbDefault Then
                DrawText lhDC, m_MsgBox.Button(i).Caption, LStrLen(m_MsgBox.Button(i).Caption), lpRect, DT_CENTER Or DT_VCENTER
            Else
                DrawText lhDC, m_MsgBox.CustomCaption(i), LStrLen(m_MsgBox.CustomCaption(i)), lpRect, DT_CENTER Or DT_VCENTER
            End If
        End If
    Next
    '   Free the Resources
    ReleaseDC lhWnd, lhDC
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.DrawWinXPButton", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Property Get Duration() As Long

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Duration = m_MsgBox.TimerInfo.TimerDuration
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.Duration", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let Duration(ByVal NewValue As Long)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_MsgBox.TimerInfo.TimerDuration = NewValue
    PropertyChanged "Duration"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.Duration", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get Font() As Font

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Set Font = m_MsgBox.Font
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.Font", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Set Font(NewValue As Font)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Set m_MsgBox.Font = NewValue
    PropertyChanged "Font"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.Font", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get ForeColor() As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    ForeColor = m_MsgBox.Prompt.ForeColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.ForeColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_MsgBox.Prompt.ForeColor = NewValue
    PropertyChanged "ForeColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.ForeColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Private Function GetObjectText(ByVal lhWnd As Long) As String
    
    '   Handle Any Errors
    On Error GoTo Func_ErrHandler

    GetObjectText = String(GetWindowTextLength(lhWnd) + 1, Chr$(0))
    GetWindowText lhWnd, GetObjectText, Len(GetObjectText)

Func_ErrHandlerExit:
    Exit Function
Func_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.GetObjectText", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Func_ErrHandlerExit:
End Function

Public Property Get hdc() As Long
    hdc = m_MsgBox.Icon.hdc
End Property

Public Property Get Icon() As StdPicture
    Set Icon = m_MsgBox.Icon.Picture
End Property

Public Property Set Icon(ByVal NewValue As StdPicture)
    Set m_MsgBox.Icon.Picture = NewValue
    PropertyChanged "Icon"
End Property

Private Function GetThemeInfo() As String
    Dim lResult As Long
    Dim sFileName As String
    Dim sColor As String
    Dim lPos As Long
    
    On Error Resume Next
    If IsWinXP Then
        '   Allocate Space
        sFileName = Space(255)
        sColor = Space(255)
        '   Read the data
        If GetCurrentThemeName(sFileName, 255, sColor, 255, vbNullString, 0) <> &H0 Then
            GetThemeInfo = "UxTheme_Error"
            Exit Function
        End If
        '   Find our trailing null terminator
        lPos = InStrRev(sColor, vbNullChar)
        '   Parse it....
        sColor = Mid(sColor, 1, lPos)
        '   Now replace the nulls....
        sColor = Replace(sColor, vbNullChar, "")
        If Trim$(sColor) = vbNullString Then sColor = "None"
        GetThemeInfo = sColor
    Else
        sColor = "None"
    End If
End Function

Public Function IsWinXP() As Boolean
    'returns True if running Windows XP
    Dim OSV As OSVERSIONINFO

    '   Handle Any Errors
    On Error GoTo Func_ErrHandler

    OSV.OSVSize = Len(OSV)
    If GetVersionEx(OSV) = 1 Then
        IsWinXP = (OSV.PlatformID = VER_PLATFORM_WIN32_NT) And _
            (OSV.dwVerMajor = 5 And OSV.dwVerMinor = 1) And _
            (OSV.dwBuildNumber >= 2600)
    End If
    
Func_ErrHandlerExit:
    Exit Function
Func_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.IsWinXP", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Func_ErrHandlerExit:
End Function

Private Sub MDIHost_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    '   We need to be sure and unsubclass on the Query else
    '   the host object may be destroyed before the MsgBox
    '   which results in a GPF ;-)
    If bSubClass And bMsgBoxSubClass And (Not Cancel) Then
        '   In most cases, we will be displaying a msgbox on close, so
        '   make sure we send an WM_ACTIVATE message to make sure it paints
        '   the MsgBox buttons correctly
        '   Send the Activate Message via code ;-)
        Call PostMessage(m_MsgBox.Parent, WM_ACTIVATE, 0, ByVal 0&)
        '   Now stop subclassing
        Call Subclass_StopAll
        '   Set our flags
        bSubClass = False
        bMsgBoxSubClass = False
    End If
End Sub

' Description: Refresh the control
Public Sub Refresh()
    Dim i As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    Select Case m_MsgBox.Theme
        Case [mbAuto]
            '   Sanity check, make sure we have a valid theme to use here....
            If Len(m_AutoTheme) = 0 Then
                m_AutoTheme = GetThemeInfo
            End If
            Select Case m_AutoTheme
                Case "None"
                    m_MsgBox.Theme = mbClassic
                    GoTo Classic
                Case "Normal Color"
                    m_MsgBox.Theme = mbBlue
                    GoTo WindowsXP
                Case "HomeStead"
                    m_MsgBox.Theme = mbHomeStead
                    GoTo WindowsXP
                Case "Metallic"
                    m_MsgBox.Theme = mbMetallic
                    GoTo WindowsXP
                Case Else
                    m_MsgBox.Theme = mbBlue
                    GoTo WindowsXP
            End Select
        Case [mbClassic]
Classic:
            'Classic Style (Win9x)
            If (bMsgBoxSubClass) Then
                Call Subclass_StopAll
                For i = 0 To m_MsgBox.Count - 1
                    '//---(invoke a Paint-event) turn to Normal ;-)
                    RedrawWindow m_MsgBox.Button(i).hWnd, ByVal 0&, ByVal 0&, &H1
                Next i
            End If
            
        Case Else '[mbWinXP]
WindowsXP:
            'WinXP (Emulated)
            If (bMsgBoxSubClass) Then
                For i = 0 To m_MsgBox.Count - 1
                    DrawWinXPButton m_MsgBox.Button(i).hWnd, m_MsgBox.Button(i).State
                Next
            End If
    End Select
      
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.Refresh", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub SDIHost_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    '   We need to be sure and unsubclass on the Query else
    '   the host object may be destroyed before the MsgBox
    '   which results in a GPF ;-)
    If bSubClass And bMsgBoxSubClass And (Not Cancel) Then
        '   In most cases, we will be displaying a msgbox on close, so
        '   make sure we send an WM_ACTIVATE message to make sure it paints
        '   the MsgBox buttons correctly
        '   Send the Activate Message via code ;-)
        Call PostMessage(m_MsgBox.Parent, WM_ACTIVATE, 0, ByVal 0&)
        '   Now stop subclassing
        Call Subclass_StopAll
        '   Set our flags
        bSubClass = False
        bMsgBoxSubClass = False
    End If
End Sub

Private Sub SelectFont(ByVal cHdc As Long, ByVal Size As Integer, ByVal Italic As Boolean, ByVal FontName As String, ByVal Underline As Boolean)
    Dim MyFont As LOGFONT
    Dim NewFont As Long

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    With MyFont
        .lfHeight = (Size * -20) / Screen.TwipsPerPixelY
        .lfCharSet = 1
        .lfItalic = Italic
        .lfUnderline = Underline
        .lfFacename = FontName & Chr$(0)
    End With
    NewFont = CreateFontIndirect(MyFont)
    SelectObject cHdc, NewFont
    DeleteObject NewFont
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.SelectFont", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Property Get SelfClosing() As Boolean

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    SelfClosing = m_MsgBox.SelfClosing
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.SelfClosing", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let SelfClosing(ByVal NewValue As Boolean)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_MsgBox.SelfClosing = NewValue
    PropertyChanged "SelfClosing"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.SelfClosing", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Private Sub SetProcessPriority(ByVal lPriority As mbPriorityEnum)
    Dim lRet As Long
    Dim lProcessID As Long
    Dim lProcessHandle As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    '   Get the Process Handled
    lProcessHandle = GetCurrentProcess
    'Debug.Print "lProcessHandle: " & lProcessHandle
    '   Sets the priority using any priority from the following
    '   Highest priority to the Lowest ;-)
    '   - REALTIME_PRIORITY_CLASS
    '   - HIGH_PRIORITY_CLASS
    '   - NORMAL_PRIORITY_CLASS
    '   - IDLE_PRIORITY_CLASS
    '   If lRet <> 0 then the changing operation action was successful
    lRet = SetPriorityClass(lProcessHandle, lPriority)

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.SetProcessPriority", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Property Get Theme() As mbThemeEnum

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Theme = m_Theme 'm_MsgBox.Theme
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.Theme", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let Theme(ByVal New_Theme As mbThemeEnum)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_MsgBox.Theme = New_Theme
    m_Theme = New_Theme
    Call Refresh
    PropertyChanged "Theme"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.Theme", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Private Function ThisWindowClassName(ByVal lhWnd As Long) As String
    Dim RetVal As Long
    Dim lpClassName As String
    
    '   Handle Any Errors
    On Error GoTo Func_ErrHandler
    
    '   Get the Classname of the Window
    lpClassName = Space(255)
    RetVal = GetClassName(lhWnd, lpClassName, 255)
    ThisWindowClassName = Left$(lpClassName, RetVal)

Func_ErrHandlerExit:
    Exit Function
Func_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.ThisWindowClassName", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Func_ErrHandlerExit:
End Function

Public Property Get ThreadPriority() As mbPriorityEnum

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    ThreadPriority = m_MsgBox.TimerInfo.ThreadPriority
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.ThreadPriority", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let ThreadPriority(ByVal NewValue As mbPriorityEnum)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_MsgBox.TimerInfo.ThreadPriority = NewValue
    PropertyChanged "ThreadPriority"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.ThreadPriority", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get TimerToken() As String

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    TimerToken = m_MsgBox.TimerInfo.TimerToken
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.TimerToken", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let TimerToken(ByVal NewValue As String)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_MsgBox.TimerInfo.TimerToken = NewValue
    PropertyChanged "TimerToken"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.TimerToken", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Private Function TranslateColor(ByVal lColor As Long) As Long
    If OleTranslateColorA(lColor, 0, TranslateColor) Then
        TranslateColor = -1
    End If
End Function

Private Function TrimNull(Item As String) As String
    Dim Pos As Integer
    
    '   Handle Any Errors
    On Error GoTo Func_ErrHandler
    
    '   Trim the Null Terminators
    Pos = InStr(Item, Chr$(0))

    If Pos Then
        TrimNull = Left$(Item, Pos - 1)
    Else
        TrimNull = Item
    End If

Func_ErrHandlerExit:
    Exit Function
Func_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.TrimNull", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Func_ErrHandlerExit:
End Function

Private Sub UserControl_InitProperties()
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    With m_MsgBox
        .Alignment = mbCenterScreen
        .BackColor = vbButtonFace
        Set .Font = Ambient.Font
        Set .Icon.Picture = Nothing
        .Prompt.Alignment = mbLeft
        .Prompt.ForeColor = &H0
        .SelfClosing = False
        .Theme = mbAuto
        .TimerInfo.ThreadPriority = mbRealTime
        .TimerInfo.TimerDuration = 10
        .TimerInfo.TimerToken = "%T"
    End With
    UserControl.BackColor = UserControl.Parent.BackColor

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.UserControl_InitProperties", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    Dim i As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    With PropBag
        m_MsgBox.Alignment = .ReadProperty("Alignment", [mbCenterScreen])
        m_MsgBox.Prompt.Alignment = .ReadProperty("AlignPrompt", [mbLeft])
        m_MsgBox.BackColor = .ReadProperty("BackColor", [vbButtonFace])
        For i = 0 To m_MsgBox.Count - 1
            m_MsgBox.CustomCaption(i) = .ReadProperty("Caption" & i, vbNullString)
        Next
        m_MsgBox.CaptionType = .ReadProperty("CaptionType", [mbDefault])
        m_MsgBox.TimerInfo.TimerDuration = .ReadProperty("Duration", 10)
        Set m_MsgBox.Font = .ReadProperty("Font", Ambient.Font)
        m_MsgBox.Prompt.ForeColor = .ReadProperty("ForeColor", &H0)
        Set m_MsgBox.Icon.Picture = .ReadProperty("Icon", Nothing)
        m_MsgBox.SelfClosing = .ReadProperty("SelfClosing", False)
        m_MsgBox.Theme = .ReadProperty("Theme", [mbAuto])
        m_MsgBox.TimerInfo.ThreadPriority = .ReadProperty("ThreadPriority", [mbRealTime])
        m_MsgBox.TimerInfo.TimerID = .ReadProperty("TimerID", MB_TIMERID)
        m_MsgBox.TimerInfo.TimerToken = .ReadProperty("TimerToken", "%T")
    End With
    
    If Ambient.UserMode Then 'If we're not in design mode
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
        
        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
                bTrack = False
            End If
        End If
        
        If bTrack Then
            'OS supports mouse leave so subclass for it
            With UserControl
                'Start subclassing the UserControl
                Call Subclass_Start(Parent.hWnd)
                Call Subclass_AddMsg(Parent.hWnd, WM_ACTIVATE, MSG_BEFORE)
                m_MsgBox.hWnd = Parent.hWnd
                If TypeOf .Parent Is Form Then
                    Set SDIHost = .Parent
                    '   Store the type of Parent Object
                    m_MsgBox.ParentType = mbSDIForm
                ElseIf TypeOf .Parent Is MDIForm Then
                    Set MDIHost = .Parent
                    '   Store the type of Parent Object
                    m_MsgBox.ParentType = mbMDIForm
                End If
                '   Store the hWnd of the Parent Object
                m_MsgBox.Parent = .Parent.hWnd
            End With
            RaiseEvent Status("HostObject Subclassing")
            
        End If
    End If
    UserControl.BackColor = UserControl.Parent.BackColor
    UserControl_Resize

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.UserControl_ReadProperties", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_Resize()
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    With UserControl
        .Width = 375
        .Height = 375
    End With
    Refresh
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.UserControl_Resize", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_Show()
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    UserControl.BackColor = UserControl.Parent.BackColor
    Refresh
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.UserControl_Show", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_Terminate()
    '   The control is terminating - a good place to stop the subclasser
    '
    '   We will keep this code in place...just in case,
    '   as the Host Objects QueryUnload should have Stopped All
    '   subclassing before we got here....
    On Error GoTo Catch
    If bSubClass Then
        '   Release all of the DCs
        With m_MsgBox
            ReleaseDC .Icon.hWnd, .Icon.hdc
            ReleaseDC .Prompt.hWnd, .Prompt.hdc
            ReleaseDC .hWnd, .hdc
        End With
        bMsgBoxSubClass = False
        Call Subclass_StopAll
        RaiseEvent Status("HostObject UnSubclassing")
    End If
Catch:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim i As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    With PropBag
        Call .WriteProperty("Alignment", m_MsgBox.Alignment, [mbCenterScreen])
        Call .WriteProperty("AlignPrompt", m_MsgBox.Prompt.Alignment, [mbLeft])
        Call .WriteProperty("BackColor", m_MsgBox.BackColor, [vbButtonFace])
        For i = 0 To m_MsgBox.Count - 1
            Call .WriteProperty("Caption" & i, m_MsgBox.CustomCaption(i), vbNullString)
        Next
        Call .WriteProperty("CaptionType", m_MsgBox.CaptionType, [mbDefault])
        Call .WriteProperty("Duration", m_MsgBox.TimerInfo.TimerDuration, 10)
        Call .WriteProperty("Font", m_MsgBox.Font, Ambient.Font)
        Call .WriteProperty("ForeColor", m_MsgBox.Prompt.ForeColor, &H0)
        Call .WriteProperty("Icon", m_MsgBox.Icon.Picture, Nothing)
        Call .WriteProperty("SelfClosing", m_MsgBox.SelfClosing, False)
        Call .WriteProperty("Theme", m_MsgBox.Theme, [mbAuto])
        Call .WriteProperty("TimerID", m_MsgBox.TimerInfo.TimerID, MB_TIMERID)
        Call .WriteProperty("ThreadPriority", m_MsgBox.TimerInfo.ThreadPriority, [mbRealTime])
        Call .WriteProperty("TimerToken", m_MsgBox.TimerInfo.TimerToken, "%T")
    End With
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucMsgBox.UserControl_WriteProperties", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Function Version(Optional ByVal bDateTime As Boolean) As String
    On Error GoTo Version_Error
    
    If bDateTime Then
        Version = Major & "." & Minor & "." & Revision & " (" & DateTime & ")"
    Else
        Version = Major & "." & Minor & "." & Revision
    End If
    Exit Function
    
Version_Error:
End Function



