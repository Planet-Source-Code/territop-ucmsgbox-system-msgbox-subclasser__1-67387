VERSION 5.00
Begin VB.Form frmMain2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ucMsgBox - Test Harness"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVersion 
      Caption         =   "Version"
      Height          =   375
      Left            =   120
      TabIndex        =   46
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdSeries 
      Caption         =   "Close"
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   39
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSeries 
      Caption         =   "MsgBox"
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   38
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Frame fmTab 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.OptionButton opTabs 
         Caption         =   "Welcome"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton opTabs 
         Caption         =   "Alignment"
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton opTabs 
         Caption         =   "Timer"
         Height          =   375
         Index           =   4
         Left            =   4200
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton opTabs 
         Caption         =   "Colors"
         Height          =   375
         Index           =   3
         Left            =   3240
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton opTabs 
         Caption         =   "Settings"
         Height          =   375
         Index           =   2
         Left            =   2280
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin prjucMsgBox.ucMsgBox ucMsgBox1 
      Left            =   1080
      Top             =   3480
      _ExtentX        =   661
      _ExtentY        =   661
      AlignPrompt     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8388608
      Icon            =   "frmMain2.frx":0000
      SelfClosing     =   -1  'True
      Theme           =   4
      TimerID         =   0
   End
   Begin VB.Frame fmProperties 
      Height          =   2415
      Index           =   0
      Left            =   120
      TabIndex        =   41
      Top             =   960
      Width           =   5055
      Begin VB.Label lblCaption 
         Alignment       =   1  'Right Justify
         Caption         =   "ucMsgBox"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3905
         TabIndex        =   45
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblWelcome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain2.frx":1CDA
         ForeColor       =   &H80000008&
         Height          =   1635
         Index           =   1
         Left            =   200
         TabIndex        =   44
         Top             =   240
         Width           =   4215
      End
      Begin VB.Image imXImage 
         Height          =   600
         Left            =   4280
         Picture         =   "frmMain2.frx":1E94
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label lblLink 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "click here"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3320
         MouseIcon       =   "frmMain2.frx":22F0
         MousePointer    =   99  'Custom
         TabIndex        =   42
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lblAuthorMessage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "To provide feedback on this control, please                 ...."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   200
         TabIndex        =   43
         Top             =   2040
         Width           =   4095
      End
   End
   Begin VB.Frame fmProperties 
      Caption         =   "Alignment Properties:"
      Height          =   2415
      Index           =   1
      Left            =   120
      TabIndex        =   28
      Top             =   960
      Width           =   5055
      Begin VB.PictureBox picContainer 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         ScaleHeight     =   375
         ScaleWidth      =   4575
         TabIndex        =   34
         Top             =   720
         Width           =   4575
         Begin VB.OptionButton opAlignment 
            Caption         =   "CenterOwner"
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   36
            Top             =   120
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton opAlignment 
            Caption         =   "CenterScreen"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   35
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.OptionButton opAlignPrompt 
         Caption         =   "Right"
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   33
         Top             =   1800
         Width           =   855
      End
      Begin VB.OptionButton opAlignPrompt 
         Caption         =   "Center"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   32
         Top             =   1800
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton opAlignPrompt 
         Caption         =   "Left"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   31
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblAlignment 
         Caption         =   "Alignment:"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblAlignPrompt 
         Caption         =   "Align Prompt:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1560
         Width           =   1095
      End
   End
   Begin VB.Frame fmProperties 
      Caption         =   "Settings Properties:"
      Height          =   2415
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   960
      Width           =   5055
      Begin VB.CheckBox chkCustomCaption 
         Caption         =   "Custom Caption:"
         Height          =   255
         Left            =   2760
         TabIndex        =   49
         Top             =   1540
         Width           =   1455
      End
      Begin VB.ComboBox cmbButtonID 
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtCustomCaption 
         Height          =   315
         Left            =   2760
         TabIndex        =   47
         Text            =   "Enter New Caption"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   325
         Left            =   240
         TabIndex        =   37
         Top             =   1800
         Width           =   1215
      End
      Begin prjucMsgBox.ucPickBox pbFont 
         Height          =   315
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         UseAutoForeColor=   0   'False
         Color           =   0
         DialogType      =   2
         Enabled         =   0   'False
         Filters         =   "Supported files|*.*|All Files (*.*)"
         FolderFlags     =   0
         Printer         =   "False"
         ToolTipText3    =   "Click Here to Locate File"
         ToolTipText4    =   "Click Here to Locate Printer"
      End
      Begin prjucMsgBox.ucPickBox pbIcon 
         Height          =   315
         Left            =   2760
         TabIndex        =   26
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         UseAutoForeColor=   0   'False
         Color           =   0
         DialogType      =   3
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.*|All Files (*.*)"
         FolderFlags     =   0
         Printer         =   "False"
         ToolTipText3    =   "Click Here to Locate File"
         ToolTipText4    =   "Click Here to Locate Printer"
      End
      Begin VB.Label lblButtonID 
         Caption         =   "Button:"
         Height          =   255
         Left            =   4320
         TabIndex        =   50
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblReset 
         Caption         =   "Properties:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblFont 
         Caption         =   "Font:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblIcon 
         Caption         =   "Icon:"
         Height          =   255
         Left            =   2760
         TabIndex        =   22
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame fmProperties 
      Caption         =   "Color Properties:"
      Height          =   2415
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   5055
      Begin VB.OptionButton opTheme 
         Caption         =   "Metallic"
         Height          =   255
         Index           =   4
         Left            =   4080
         TabIndex        =   52
         Top             =   1800
         Width           =   855
      End
      Begin VB.OptionButton opTheme 
         Caption         =   "HomeStead"
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   51
         Top             =   1800
         Width           =   1215
      End
      Begin prjucMsgBox.ucPickBox pbBackColor 
         Height          =   315
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         UseAutoForeColor=   0   'False
         Color           =   0
         Enabled         =   0   'False
         Filters         =   "Supported files|*.*|All Files (*.*)"
         FolderFlags     =   0
         Printer         =   "False"
         ToolTipText3    =   "Click Here to Locate File"
         ToolTipText4    =   "Click Here to Locate Printer"
      End
      Begin VB.OptionButton opTheme 
         Caption         =   "Blue"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   15
         Top             =   1800
         Width           =   855
      End
      Begin VB.OptionButton opTheme 
         Caption         =   "Auto"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton opTheme 
         Caption         =   "Classic"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   13
         Top             =   1800
         Width           =   855
      End
      Begin prjucMsgBox.ucPickBox pbForeColor 
         Height          =   315
         Left            =   2760
         TabIndex        =   20
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         UseAutoForeColor=   0   'False
         Color           =   0
         Enabled         =   0   'False
         Filters         =   "Supported files|*.*|All Files (*.*)"
         FolderFlags     =   0
         Printer         =   "False"
         ToolTipText3    =   "Click Here to Locate File"
         ToolTipText4    =   "Click Here to Locate Printer"
      End
      Begin VB.Label lblForeColor 
         Caption         =   "ForeColor:"
         Height          =   255
         Left            =   2760
         TabIndex        =   18
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblBackColor 
         Caption         =   "BackColor:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblTheme 
         Caption         =   "Theme:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   855
      End
   End
   Begin VB.Frame fmProperties 
      Caption         =   "Timer Properties:"
      Height          =   2415
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   5055
      Begin VB.TextBox txtTimerToken 
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Text            =   "%T"
         Top             =   1800
         Width           =   2055
      End
      Begin VB.ComboBox cbThreadPriority 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtDuration 
         Height          =   315
         Left            =   2760
         TabIndex        =   9
         Text            =   "10"
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox chkSelfClosing 
         Caption         =   "Self Closing"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.Label lblDuration 
         Caption         =   "Timer Duration (sec):"
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblThreadPriority 
         Caption         =   "Timer Thread Priority:"
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label lblTimerToken 
         Caption         =   "Timer String Token:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   2055
      End
   End
   Begin VB.Image imIcon 
      Height          =   720
      Left            =   240
      Picture         =   "frmMain2.frx":25FA
      Top             =   2520
      Width           =   720
   End
End
Attribute VB_Name = "frmMain2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

'   Shell to call Explorer ;-)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'   Link URL address which searches for our control submission on PCS
Const sLink As String = "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&?lngWId=1&grpCategories=&txtMaxNumberOfEntriesPerPage=10&optSort=Alphabetical&chkThoroughSearch=&blnTopCode=False&blnNewestCode=False&blnAuthorSearch=False&lngAuthorId=&strAuthorName=&blnResetAllVariables=&blnEditCode=False&mblnIsSuperAdminAccessOn=False&intFirstRecordOnPage=1&intLastRecordOnPage=10&intMaxNumberOfEntriesPerPage=10&intLastRecordInRecordset=499&chkCodeTypeZip=&chkCodeDifficulty=&chkCodeTypeText=&chkCodeTypeArticle=&chkCode3rdPartyReview=&txtCriteria=ucMsgBox"

Private bLoading As Boolean

'   Force Declarations
Option Explicit

Private Sub cbThreadPriority_Click()
    With Me
        Select Case cbThreadPriority.ListIndex
            Case 0  'Realtime
                .ucMsgBox1.ThreadPriority = mbRealTime
            Case 1  'High
                .ucMsgBox1.ThreadPriority = mbHigh
            Case 2  'Normal
                .ucMsgBox1.ThreadPriority = mbNormal
            Case 3  'Idle
                .ucMsgBox1.ThreadPriority = mbIdle
        End Select
    End With
End Sub

Private Sub chkCustomCaption_Click()
    With Me
        If (.chkCustomCaption.Value = vbUnchecked) Then
            .ucMsgBox1.CaptionType = mbDefault
        Else
            .ucMsgBox1.CaptionType = mbCustom
        End If
    End With
End Sub

Private Sub chkSelfClosing_Click()
    With Me
        .ucMsgBox1.SelfClosing = (.chkSelfClosing.Value = vbChecked)
    End With
End Sub

Private Sub cmbButtonID_Click()
    With Me
        If Not bLoading Then
            .txtCustomCaption.Text = .ucMsgBox1.Caption(.cmbButtonID.ListIndex)
            .ucMsgBox1.Caption(.cmbButtonID.ListIndex) = .txtCustomCaption.Text
        End If
    End With
End Sub

Private Sub cmdReset_Click()
    Dim StdFnt As StdFont
    With Me
       With .ucMsgBox1
            '   MsgBox Alignment relative to.....Owner
            .Alignment = mbCenterOwner
            '   Align Prompt to Center
            .AlignPrompt = mbCenter
            '   Set the BackColor
            .BackColor = vbButtonFace
            '   Set the SelfClosing Duration in sec
            .Duration = 10
            '   Create a Standard Font that the MsgBox would use...
            Set StdFnt = New StdFont
            '   Fill the Font structure
            With StdFnt
                .Bold = False
                .Italic = False
                .Name = "Tahoma"
                .Size = 8
                .Strikethrough = False
                .Underline = False
                .Weight = 400
            End With
            '   Now assign it
            Set .Font = StdFnt
            '   ForeColor for Buttons and Prompt
            .ForeColor = vbBlack
            '   Icon....could be "Nothing" if you want the defaults
            Set .Icon = imIcon.Picture
            '   Selfclosing style
            .SelfClosing = False 'True
            '   Theme....Auto Detect
            .Theme = mbAuto
            '   Time Thread Priority...set to Most Accurate
            .ThreadPriority = mbRealTime
            '   Set the Token to have replaced when the count down
            '   occurs on the MsgBox....could use anything for this...
            .TimerToken = "%T"
        End With
    End With
End Sub

Private Sub cmdSeries_Click(Index As Integer)
    With Me
        Select Case Index
            Case 0
                '    Show the MsgBoxs....
                '
                '   Start off simple
                MsgBox "This is a Test of a Simple Message, One Button, No Icon", vbOKOnly, "ucMsgBox"
                '   Now A bit more complex...similar to what an application might have
                MsgBox "This is a Test of a More Complex Message with Self Closure." & vbCrLf & "This Message Dialog will Close in %T Sec.", vbDefaultButton1 + vbYesNoCancel, "ucMsgBox"
                '   This is propably the Most Complicated....
                MsgBox "This is a Test of Yet a More Complex Message," & vbCrLf & "with Multiple Buttons, Default Buttons, and Self Closure." & vbCrLf & "This Message Dialog will Close in %T Sec.", vbMsgBoxHelpButton + vbDefaultButton1 + vbRetryCancel + vbCritical, "ucMsgBox"
                '   Store the Image so we can reset things...
                Set .imIcon.Picture = .ucMsgBox1.Icon
                '   Set the Icon to Null
                Set .ucMsgBox1.Icon = Nothing
                '   This Illustrates that the Number or Complexity of the Lines is not an issue ;-)
                MsgBox "This is a Test of Complex Prompts, Buttons, Defaults, and Self Closure." & vbCrLf & "This Prompt is Multiple Lines, Yet the Prompt Alignment is Maintained!!" & vbCrLf & "This Message Dialog will Close in %T Sec.", vbMsgBoxHelpButton + vbDefaultButton1 + vbQuestion + vbAbortRetryIgnore + vbCritical, "ucMsgBox"
                Set .ucMsgBox1.Icon = .imIcon.Picture
            Case 1
                '   Unload the Project...
                Unload Me
        End Select
    End With
End Sub

Private Sub cmdVersion_Click()
    With Me
        '   Get our version info
        MsgBox "ucMsgBox Control Version:" & vbCrLf & vbCrLf & .ucMsgBox1.Version(True) & vbCrLf & vbCrLf & "MsgBox will Close in %T Seconds...", vbInformation + vbOKOnly, "ucMsgBox"
    End With
End Sub

Private Sub Form_Load()
    With Me
        bLoading = True
        '   Set the Startup Default Values
        .opAlignment(1).Value = True
        .opAlignPrompt(1).Value = True
        With .pbFont
            .DialogType = ucFont
            .FontFlags = ShowFont_Default
        End With
        With .pbIcon
            .DialogType = ucOpen
            .MultiSelect = False
        End With
        With .pbBackColor
            .DialogType = ucColor
            .ColorFlags = ShowColor_Default
        End With
        With .pbForeColor
            .DialogType = ucColor
            .ColorFlags = ShowColor_Default
        End With
        With .cmbButtonID
            .AddItem 0
            .AddItem 1
            .AddItem 2
            .AddItem 3
            .ListIndex = 0
        End With
        .opTheme(0).Value = True
        .chkSelfClosing.Value = vbUnchecked
        .txtDuration.Text = "10"
        .txtTimerToken.Text = "%T"
        With .cbThreadPriority
            .AddItem "RealTime"
            .AddItem "High"
            .AddItem "Normal"
            .AddItem "Idle"
            .ListIndex = 0
        End With
        Call cmdReset_Click
        Call opTabs_Click(0)
        bLoading = False
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    With Me
        '   Turn off SelfClosing
        .ucMsgBox1.SelfClosing = False
        '   Call the MsgBox as normal
        If MsgBox("Close the Current Session?", vbQuestion + vbYesNoCancel, "ucMsgBox") <> vbYes Then
            '   We are canceling, so set back the SelfClosing property to
            '   what the user selected via the GUI
            .ucMsgBox1.SelfClosing = (.chkSelfClosing.Value = vbChecked)
            '   Cancel the form closure
            Cancel = True
        End If
    End With
End Sub

Private Sub lblAuthorMessage_Click()
    Dim OpenLink As String
    With Me
        '   Set the link color
        .lblLink.ForeColor = &HC000C0
        '   Launch the browser and follow the link
        OpenLink = ShellExecute(hWnd, "open", sLink, vbNull, vbNull, 1)
        Me.SetFocus
    End With
End Sub

Private Sub lblLink_Click()
    Dim OpenLink As String
    With Me
        '   Set the link color
        .lblLink.ForeColor = &HC000C0
        '   Launch the browser and follow the link
        OpenLink = ShellExecute(hWnd, "open", sLink, vbNull, vbNull, 1)
        Me.SetFocus
    End With
End Sub

Private Sub opAlignment_Click(Index As Integer)
    With Me
        '   Set the MsgBox Alignment based on indexes....
        .ucMsgBox1.Alignment = Index
    End With
End Sub

Private Sub opAlignPrompt_Click(Index As Integer)
    With Me
        '   Set the Promp Alignment property based on indexes
        .ucMsgBox1.AlignPrompt = Index
    End With
End Sub

Private Sub opTabs_Click(Index As Integer)
    Dim i As Long
    
    With Me
        '   Simple Tab Like effect with OptionButtons and Frames
        For i = .fmProperties.LBound To .fmProperties.UBound
            '   Make them all InVisible
            .fmProperties(i).Visible = False
        Next i
        '   Now set our selected one to visible
        .fmProperties(Index).Visible = True
    End With
End Sub

Private Sub opTheme_Click(Index As Integer)
    With Me
        '   Set the Theme by index
        .ucMsgBox1.Theme = Index
    End With
End Sub

Private Sub pbBackColor_Click()
    With Me
        '   Set the BackColor from the passed value
        .ucMsgBox1.BackColor = .pbBackColor.Color
    End With
End Sub

Private Sub pbFont_Click()
    With Me
        '   Set the Font from the passed value
        Set .ucMsgBox1.Font = .pbFont.Font
    End With
End Sub

Private Sub pbForeColor_Click()
    With Me
        '   Set the ForeColor from the passed value
        .ucMsgBox1.ForeColor = .pbForeColor.Color
    End With
End Sub

Private Sub pbIcon_Click()
    With Me
        If pbIcon.FileExists(pbIcon.Filename) Then
            '   Only set Icons which are located on disk
            Set .ucMsgBox1.Icon = LoadPicture(pbIcon.Filename)
        End If
    End With
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub txtCustomCaption_GotFocus()
    With Me.txtCustomCaption
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtCustomCaption_KeyUp(KeyCode As Integer, Shift As Integer)
    With Me
        If KeyCode = vbKeyReturn Then
            txtCustomCaption_LostFocus
        End If
    End With
End Sub

Private Sub txtCustomCaption_LostFocus()
    With Me
        .ucMsgBox1.Caption(.cmbButtonID.ListIndex) = .txtCustomCaption.Text
    End With
End Sub

Private Sub txtDuration_GotFocus()
    With Me
        '   Auto Select the Text
        With .txtDuration
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End With
End Sub

Private Sub txtDuration_KeyUp(KeyCode As Integer, Shift As Integer)
    With Me
        '   Forward our calls
        Call txtDuration_LostFocus
    End With
End Sub

Private Sub txtDuration_LostFocus()
    With Me
        '   Only use Numeric Values here!!!
        If IsNumeric(.txtDuration.Text) Then
            .ucMsgBox1.Duration = CLng(.txtDuration.Text)
        Else
            MsgBox "Error: Non-Numeric Value" & vbCrLf & "Please Try Again", vbExclamation + vbOKOnly, "ucMsgBox"
            .txtDuration.Text = .ucMsgBox1.Duration
        End If
    End With
End Sub

Private Sub txtTimerToken_Change()
    With Me
        '   Set the TimerToken
        .ucMsgBox1.TimerToken = .txtTimerToken.Text
    End With
End Sub

Private Sub txtTimerToken_GotFocus()
    With Me
        '   Auto Select the Text
        With .txtTimerToken
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End With
End Sub

Private Sub ucMsgBox1_Status(ByVal sStatus As String)
    '   Print our Status Messages
    Debug.Print sStatus
End Sub
