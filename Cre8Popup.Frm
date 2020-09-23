VERSION 5.00
Begin VB.Form Cre8Popup 
   BackColor       =   &H80000007&
   Caption         =   "Click inside the form..."
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   366
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   525
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Cre8Popup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'*************************************************************************
'*************************************************************************
'****                                                                 ****
'****    This is the PopUp Menu Tutorial.                             ****
'****                                                                 ****
'****    Written by:  Randy Giese                                     ****
'****    Company:     RandyGrams                                      ****
'****    Date:        April 09, 2007                                  ****
'****                                                                 ****
'****    I struggled with PopUp Menus because I was not able to       ****
'****    find very good documentation.  I found examples on PSC       ****
'****    and elsewhere but they seemed too limited, as well as        ****
'****    complex and difficult, to relate to my own coding.  So I     ****
'****    decided to dig into it and was able to come up with this     ****
'****    brief Tutorial.  I hope you find it helpful...               ****
'****                                                                 ****
'****    Randy Giese                                                  ****
'****                                                                 ****
'*************************************************************************
'*************************************************************************
'
Option Explicit
Private lngBoxNum                    As Long
'
'*************************************************************************
'****                                                                 ****
'****    These are the Constants for the 'AppendMenu' Function.       ****
'****                                                                 ****
'****    Note:                                                        ****
'****    I included all of the Menu Function Constants and the        ****
'****    TrackPopupMenu Constants for completeness.                   ****
'****                                                                 ****
'*************************************************************************
'
Private Const MF_APPEND             As Long = &H100&
Private Const MF_CHECKED            As Long = &H8&
Private Const MF_DISABLED           As Long = &H2&
Private Const MF_GRAYED             As Long = &H1&
Private Const MF_POPUP              As Long = &H10&
Private Const MF_SEPARATOR          As Long = &H800&
Private Const MF_STRING             As Long = &H0&
Private Const MF_UNCHECKED          As Long = &H0&
'
'*************************************************************************
'****                                                                 ****
'****    These are the Constants for the 'TrackPopupMenuEx'           ****
'****    Function.                                                    ****
'****                                                                 ****
'*************************************************************************
'
Private Const TPM_BOTTOMALIGN       As Long = &H20&
Private Const TPM_CENTERALIGN       As Long = &H4&
Private Const TPM_HORIZONTAL        As Long = &H0&
Private Const TPM_HORNEGANIMATION   As Long = &H800&
Private Const TPM_HORPOSANIMATION   As Long = &H400&
Private Const TPM_LEFTALIGN         As Long = &H0&
Private Const TPM_LEFTBUTTON        As Long = &H0&
Private Const TPM_NOANIMATION       As Long = &H4000&
Private Const TPM_NONOTIFY          As Long = &H80&
Private Const TPM_RECURSE           As Long = &H1&
Private Const TPM_RETURNCMD         As Long = &H100&
Private Const TPM_RIGHTALIGN        As Long = &H8&
'
'*************************************************************************
'****                                                                 ****
'****    I wanted to give a 'Heads-up' about the BUTTON               ****
'****    parameters.                                                  ****
'****                                                                 ****
'****    TPM_LEFTBUTTON                                               ****
'****    If this flag is set, the user can select menu items with     ****
'****    only the left mouse button.                                  ****
'****                                                                 ****
'****    TPM_RIGHTBUTTON                                              ****
'****    If this flag is set, the user can select menu items with     ****
'****    both the left and right mouse buttons.                       ****
'****                                                                 ****
'****    You might think that TPM_RIGHTBUTTON would select menu       ****
'****    items with the right mouse button but reread                 ****
'****    TPM_RIGHTBUTTON and you will see that this isn't the         ****
'****    case.                                                        ****
'****                                                                 ****
'****    TPM_RIGHTBUTTON allows both the left and right mouse         ****
'****    buttons to be used.  Just a warning...                       ****
'****                                                                 ****
'*************************************************************************
'
Private Const TPM_RIGHTBUTTON       As Long = &H2&
Private Const TPM_TOPALIGN          As Long = &H0&
Private Const TPM_VCENTERALIGN      As Long = &H10&
Private Const TPM_VERNEGANIMATION   As Long = &H2000&
Private Const TPM_VERPOSANIMATION   As Long = &H1000&
Private Const TPM_VERTICAL          As Long = &H40&
'
'*************************************************************************
'****                                                                 ****
'****    POINTAPI is used in the 'TrackPopupMenuEx' Function.         ****
'****                                                                 ****
'*************************************************************************
'
Private Type POINTAPI
    lngX                            As Long
    lngY                            As Long
End Type
'
'*************************************************************************
'****                                                                 ****
'****    These are the Menu Functions.                                ****
'****                                                                 ****
'*************************************************************************
'
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal HWnd As Long, ByVal lptpm As Any) As Long

Private Sub Form_Load()

'   lngBoxNum is used in the 'Color' Sub-Menu
    lngBoxNum = 1

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim lngTPM                          As Long
Dim pntXorY                         As POINTAPI
'
'*************************************************************************
'****                                                                 ****
'****    First I'll dimension each of the Menu variables.             ****
'****    The variable will hold the Menu address.                     ****
'****                                                                 ****
'*************************************************************************
'
Dim lngMainMenu                     As Long
Dim lngFirst_Sub_Menu               As Long
Dim lngSecond_Sub_Menu              As Long
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Then get the address of the Menu.  The variable contains     ****
'   ****    the address.                                                 ****
'   ****                                                                 ****
'   ****    Note:                                                        ****
'   ****    You need to 'CreatePopupMenu' for the Main-Menu and for      ****
'   ****    each Sub-Menu.                                               ****
'   ****                                                                 ****
'   ****    In my example I have my Main-Menu (lngMainMenu) and two      ****
'   ****    Sub-Menus (lngFirst_Sub_Menu) and (lngSecond_Sub_Menu).      ****
'   ****                                                                 ****
'   *************************************************************************
'
    lngMainMenu = CreatePopupMenu()
    lngFirst_Sub_Menu = CreatePopupMenu()
    lngSecond_Sub_Menu = CreatePopupMenu()
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    Save the current Cursor Position.                            ****
'   ****    This will be used to locate the Popup Menu in the            ****
'   ****    'TrackPopupMenuEx' statement.                                ****
'   ****                                                                 ****
'   *************************************************************************
'
    GetCursorPos pntXorY
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    This is the Main-Menu.                                       ****
'   ****                                                                 ****
'   *************************************************************************
'
    AppendMenu lngMainMenu, MF_STRING, 1, "This is the..."
    AppendMenu lngMainMenu, MF_CHECKED, 2, "Main Menu"
    AppendMenu lngMainMenu, MF_SEPARATOR, 3, ByVal 0&
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    The 'MF_POPUP' is the button that will Popup the             ****
'   ****    Sub-Menu.                                                    ****
'   ****                                                                 ****
'   ****    It references the Sub-Menu's identifier                      ****
'   ****    (lngFirst_Sub_Menu).                                         ****
'   ****                                                                 ****
'   *************************************************************************
'
    AppendMenu lngMainMenu, MF_POPUP, lngFirst_Sub_Menu, "More Menus"
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    This is the First Sub-Menu.                                  ****
'   ****                                                                 ****
'   *************************************************************************
'
    AppendMenu lngFirst_Sub_Menu, MF_STRING, 4, "This is an example of..."
    AppendMenu lngFirst_Sub_Menu, MF_SEPARATOR, 5, ByVal 0&
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    The 'MF_POPUP' is the button that will Popup the 'Colors'    ****
'   ****    Sub-Menu (Check-Boxes).                                      ****
'   ****                                                                 ****
'   ****    It references the Sub-Menu's identifier                      ****
'   ****    (lngSecond_Sub_Menu).                                        ****
'   ****                                                                 ****
'   *************************************************************************
'
    AppendMenu lngFirst_Sub_Menu, MF_POPUP, lngSecond_Sub_Menu, "A Sub Menu w/Checked items..."
    AppendMenu lngFirst_Sub_Menu, MF_SEPARATOR, 6, ByVal 0&
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    This is the 'Colors' Sub-Menu w/Check-Boxes.                 ****
'   ****                                                                 ****
'   ****    I replaced the cumbersome "Select Case" routine with this    ****
'   ****    much more compact version.  The  "Select Case" required a    ****
'   ****    line for each Check-Box in each "Case".  In my example,      ****
'   ****    there were 3 Check-Boxes, which meant 3 lines of code 3      ****
'   ****    times, plus the "Select Case" lines.  This isn't a           ****
'   ****    terrible thing, but what if someone wanted to have a         ****
'   ****    Check-Box for each letter of the alphabet?  26 letters,      ****
'   ****    26 times plus "Select Case" code would be over 600 lines     ****
'   ****    of code.  That wouldn't be good, so I came up with the       ****
'   ****    following formula:                                           ****
'   ****                                                                 ****
'   ****    Hex(Abs(lngBoxNum = 1) * 8)                                  ****
'   ****                                                                 ****
'   ****    I take the Check-Box that was selected (lngBoxNum) and do    ****
'   ****    a Boolean check to see if it matches the Check-Box number    ****
'   ****    (the '1' in the formula above).  The Boolean check will      ****
'   ****    return a '-1' if it's true and a '0' if it's false.          ****
'   ****                                                                 ****
'   ****    Next I take the Absolute value of the Boolean check.         ****
'   ****    That leaves me with a '0' or a '1'.                          ****
'   ****                                                                 ****
'   ****    Then I multiply that number by 8 (8 * 0 = 0, 8 * 1 = 8)      ****
'   ****    and convert it to Hex.                                       ****
'   ****                                                                 ****
'   ****    That leaves me with a Hex 8, which is the value of a         ****
'   ****    Checked box (MF_CHECKED) or a Hex 0, which is the value      ****
'   ****    of an unChecked box (MF_UNCHECKED).                          ****
'   ****                                                                 ****
'   ****    Using this formula, if someone wants to have 26              ****
'   ****    Check-Boxes, it will take 26 lines of code.  Much better     ****
'   ****    than 600+ code lines.                                        ****
'   ****                                                                 ****
'   *************************************************************************
'
    AppendMenu lngSecond_Sub_Menu, Hex(Abs(lngBoxNum = 1) * 8), 7, "Red"
    AppendMenu lngSecond_Sub_Menu, Hex(Abs(lngBoxNum = 2) * 8), 8, "Green"
    AppendMenu lngSecond_Sub_Menu, Hex(Abs(lngBoxNum = 3) * 8), 9, "Blue"
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    This is the code that was replaced above.                    ****
'   ****                                                                 ****
'   ****    Select Case lngBoxNum                                        ****
'   ****    Case 1                                                       ****
'   ****    AppendMenu lngSecond_Sub_Menu, MF_CHECKED, 7, "Red"          ****
'   ****    AppendMenu lngSecond_Sub_Menu, MF_UNCHECKED, 8, "Green"      ****
'   ****    AppendMenu lngSecond_Sub_Menu, MF_UNCHECKED, 9, "Blue"       ****
'   ****    Case 2                                                       ****
'   ****    AppendMenu lngSecond_Sub_Menu, MF_UNCHECKED, 7, "Red"        ****
'   ****    AppendMenu lngSecond_Sub_Menu, MF_CHECKED, 8, "Green"        ****
'   ****    AppendMenu lngSecond_Sub_Menu, MF_UNCHECKED, 9, "Blue"       ****
'   ****    Case 3                                                       ****
'   ****    AppendMenu lngSecond_Sub_Menu, MF_UNCHECKED, 7, "Red"        ****
'   ****    AppendMenu lngSecond_Sub_Menu, MF_UNCHECKED, 8, "Green"      ****
'   ****    AppendMenu lngSecond_Sub_Menu, MF_CHECKED, 9, "Blue"         ****
'   ****    End Select  '   lngBoxNum                                    ****
'   ****                                                                 ****
'   *************************************************************************
'

'
'   *************************************************************************
'   ****                                                                 ****
'   ****    This is the rest of the First Sub_Menu.                      ****
'   ****                                                                 ****
'   *************************************************************************
'
    AppendMenu lngFirst_Sub_Menu, MF_STRING, 10, "Open the Sub Menu above."
    AppendMenu lngFirst_Sub_Menu, MF_STRING, 11, "Then Click one of the colors."
    AppendMenu lngFirst_Sub_Menu, MF_STRING, 12, "The next time you open it,"
    AppendMenu lngFirst_Sub_Menu, MF_STRING, 13, "your new color will be Checked."
    AppendMenu lngFirst_Sub_Menu, MF_SEPARATOR, 14, ByVal 0&
    AppendMenu lngFirst_Sub_Menu, MF_STRING, 15, "Close"
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    This is the rest of the Main-Menu.                           ****
'   ****                                                                 ****
'   *************************************************************************
'
    AppendMenu lngMainMenu, MF_SEPARATOR, 16, ByVal 0&
    AppendMenu lngMainMenu, MF_STRING, 17, "The next one is"
    AppendMenu lngMainMenu, MF_GRAYED, 18, "Grayed out."
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    The TrackPopupMenuEx function displays a shortcut menu at    ****
'   ****    the specified location and tracks the selection of items     ****
'   ****    on the shortcut menu. The shortcut menu can appear           ****
'   ****    anywhere on the screen. (This definition is from the         ****
'   ****    API-Guide which was created by the KPD team).                ****
'   ****                                                                 ****
'   *************************************************************************
'
    lngTPM = TrackPopupMenuEx(lngMainMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, pntXorY.lngX, pntXorY.lngY, Me.HWnd, ByVal 0&)
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    The DestroyMenu function destroys the specified menu and     ****
'   ****    frees any memory that the menu occupies.                     ****
'   ****                                                                 ****
'   *************************************************************************
'
    DestroyMenu lngMainMenu
    DestroyMenu lngFirst_Sub_Menu
    DestroyMenu lngSecond_Sub_Menu
'
'   *************************************************************************
'   ****                                                                 ****
'   ****    This is where you control the Menu buttons.                  ****
'   ****    This Menu is merely for demonstration purposes so I only     ****
'   ****    included an "End" statement for the Quit button here.        ****
'   ****                                                                 ****
'   *************************************************************************
'
    Select Case lngTPM
        Case 7
            lngBoxNum = 1
        Case 8
            lngBoxNum = 2
        Case 9
            lngBoxNum = 3
        Case 15
            End
    End Select  '   lngTPM

End Sub
