Attribute VB_Name = "mod_kb_evnt"

Option Explicit


Public Const VK_LBUTTON As Long = &H1 'Left mouse button

Public Const VK_RBUTTON As Long = &H2 'Right mouse button

Public Const VK_CANCEL As Long = &H3 'Control-break processing

Public Const VK_MBUTTON As Long = &H4 'Middle mouse button (three-button mouse), NOT contiguous with L & RBUTTON

Public Const VK_XBUTTON1 As Long = &H5 'Windows 2000/XP: X1 mouse button, NOT contiguous with L & RBUTTON

Public Const VK_XBUTTON2 As Long = &H6 'Windows 2000/XP: X2 mouse button, NOT contiguous with L & RBUTTON


'&H7 unassigned

Public Const VK_BACK As Long = &H8 'BACKSPACE key

Public Const VK_TAB As Long = &H9 'TAB key


'&H0A to &H0B reserved

Public Const VK_CLEAR As Long = &HC 'CLEAR key

Public Const VK_RETURN As Long = &HD 'ENTER key


Public Const VK_SHIFT As Long = &H10 'SHIFT key

Public Const VK_CONTROL As Long = &H11 'CTRL key

Public Const VK_MENU As Long = &H12 'ALT key

Public Const VK_PAUSE As Long = &H13 'PAUSE key

Public Const VK_CAPITAL As Long = &H14 'CAPS LOCK key

Public Const VK_KANA As Long = &H15 'Input Method Editor (IME) Kana mode

Public Const VK_HANGEUL As Long = &H15 'IME Hanguel mode (maintained for compatibility; use VK_HANGUL), old name - should be here for compatibility

Public Const VK_HANGUL As Long = &H15 'IME Hangul mode

Public Const VK_JUNJA As Long = &H17 'IME Junja mode

Public Const VK_FINAL As Long = &H18 'IME final mode

Public Const VK_HANJA As Long = &H19 'IME Hanja mode

Public Const VK_KANJI As Long = &H19 'IME Kanji mode

Public Const VK_ESCAPE As Long = &H1B 'ESC key

 

Public Const VK_CONVERT As Long = &H1C 'IME convert

Public Const VK_NONCONVERT As Long = &H1D 'IME nonconvert

Public Const VK_ACCEPT As Long = &H1E 'IME accept

Public Const VK_MODECHANGE As Long = &H1F 'IME mode change request

 

Public Const VK_SPACE As Long = &H20 'SPACEBAR

Public Const VK_PRIOR As Long = &H21 'PAGE UP key

Public Const VK_NEXT As Long = &H22 'PAGE DOWN key

Public Const VK_END As Long = &H23 'END key

Public Const VK_HOME As Long = &H24 'HOME key

Public Const VK_LEFT As Long = &H25 'LEFT ARROW key

Public Const VK_UP As Long = &H26 'UP ARROW key

Public Const VK_RIGHT As Long = &H27 'RIGHT ARROW key

Public Const VK_DOWN As Long = &H28 'DOWN ARROW key

Public Const VK_SELECT As Long = &H29 'SELECT key

Public Const VK_PRINT As Long = &H2A 'PRINT key

Public Const VK_EXECUTE As Long = &H2B 'EXECUTE key

Public Const VK_SNAPSHOT As Long = &H2C 'PRINT SCREEN key

Public Const VK_INSERT As Long = &H2D 'INS key

Public Const VK_DELETE As Long = &H2E 'DEL key

Public Const VK_HELP As Long = &H2F 'HELP key

 

Public Const VK_0 As Long = &H30 '0 key

Public Const VK_1 As Long = &H31 '1 key

Public Const VK_2 As Long = &H32 '2 key

Public Const VK_3 As Long = &H33 '3 key

Public Const VK_4 As Long = &H34 '4 key

Public Const VK_5 As Long = &H35 '5 key

Public Const VK_6 As Long = &H36 '6 key

Public Const VK_7 As Long = &H37 '7 key

Public Const VK_8 As Long = &H38 '8 key

Public Const VK_9 As Long = &H39 '9 key

 

'&H40 unassigned


Public Const VK_A As Long = &H41 'A key

Public Const VK_B As Long = &H42 'B key

Public Const VK_C As Long = &H43 'C key

Public Const VK_D As Long = &H44 'D key

Public Const VK_E As Long = &H45 'E key

Public Const VK_F As Long = &H46 'F key

Public Const VK_G As Long = &H47 'G key

Public Const VK_H As Long = &H48 'H key

Public Const VK_I As Long = &H49 'I key

Public Const VK_J As Long = &H4A 'J key

Public Const VK_K As Long = &H4B 'K key

Public Const VK_L As Long = &H4C 'L key

Public Const VK_M As Long = &H4D 'M key

Public Const VK_N As Long = &H4E 'N key

Public Const VK_O As Long = &H4F 'O key

Public Const VK_P As Long = &H50 'P key

Public Const VK_Q As Long = &H51 'Q key

Public Const VK_R As Long = &H52 'R key

Public Const VK_S As Long = &H53 'S key

Public Const VK_T As Long = &H54 'T key

Public Const VK_U As Long = &H55 'U key

Public Const VK_V As Long = &H56 'V key

Public Const VK_W As Long = &H57 'W key

Public Const VK_X As Long = &H58 'X key

Public Const VK_Y As Long = &H59 'Y key

Public Const VK_Z As Long = &H5A 'Z key

 

Public Const VK_STARTKEY As Long = &H5B 'Windows Start Key (equivalent to Left Windows Key)

Public Const VK_LWIN As Long = &H5B 'Left Windows key (Microsoft Natural keyboard)

Public Const VK_RWIN As Long = &H5C 'Right Windows key (Natural keyboard)

Public Const VK_APPS As Long = &H5D 'Applications key (Natural keyboard)

Public Const VK_CONTEXTKEY As Long = &H5D 'Context-Menu Key (equivalent to Applications key)

 

'&H5E reserved

 

Public Const VK_SLEEP As Long = &H5F 'Computer Sleep key

 

Public Const VK_NUMPAD0 As Long = &H60 'Numeric keypad 0 key

Public Const VK_NUMPAD1 As Long = &H61 'Numeric keypad 1 key

Public Const VK_NUMPAD2 As Long = &H62 'Numeric keypad 2 key

Public Const VK_NUMPAD3 As Long = &H63 'Numeric keypad 3 key

Public Const VK_NUMPAD4 As Long = &H64 'Numeric keypad 4 key

Public Const VK_NUMPAD5 As Long = &H65 'Numeric keypad 5 key

Public Const VK_NUMPAD6 As Long = &H66 'Numeric keypad 6 key

Public Const VK_NUMPAD7 As Long = &H67 'Numeric keypad 7 key

Public Const VK_NUMPAD8 As Long = &H68 'Numeric keypad 8 key

Public Const VK_NUMPAD9 As Long = &H69 'Numeric keypad 9 key

Public Const VK_MULTIPLY As Long = &H6A 'Multiply key

Public Const VK_ADD As Long = &H6B 'Add key

Public Const VK_SEPARATOR As Long = &H6C 'Separator key

Public Const VK_SUBTRACT As Long = &H6D 'Subtract key

Public Const VK_DECIMAL As Long = &H6E 'Decimal key

Public Const VK_DIVIDE As Long = &H6F 'Divide key

Public Const VK_F1 As Long = &H70 'F1 key

Public Const VK_F2 As Long = &H71 'F2 key

Public Const VK_F3 As Long = &H72 'F3 key

Public Const VK_F4 As Long = &H73 'F4 key

Public Const VK_F5 As Long = &H74 'F5 key

Public Const VK_F6 As Long = &H75 'F6 key

Public Const VK_F7 As Long = &H76 'F7 key

Public Const VK_F8 As Long = &H77 'F8 key

Public Const VK_F9 As Long = &H78 'F9 key

Public Const VK_F10 As Long = &H79 'F10 key

Public Const VK_F11 As Long = &H7A 'F11 key

Public Const VK_F12 As Long = &H7B 'F12 key

Public Const VK_F13 As Long = &H7C 'F13 key

Public Const VK_F14 As Long = &H7D 'F14 key

Public Const VK_F15 As Long = &H7E 'F15 key

Public Const VK_F16 As Long = &H7F 'F16 key

Public Const VK_F17 As Long = &H80 'F17 key

Public Const VK_F18 As Long = &H81 'F18 key

Public Const VK_F19 As Long = &H82 'F19 key

Public Const VK_F20 As Long = &H83 'F20 key

Public Const VK_F21 As Long = &H84 'F21 key

Public Const VK_F22 As Long = &H85 'F22 key

Public Const VK_F23 As Long = &H86 'F23 key

Public Const VK_F24 As Long = &H87 'F24 key

 

'&H88 to &H8F unassigned

 

Public Const VK_NUMLOCK As Long = &H90 'NUM LOCK key

Public Const VK_SCROLL As Long = &H91 'SCROLL LOCK key

 

'NEC PC-9800 kbd definitions

Public Const VK_OEM_NEC_EQUAL As Long = &H92 '= key on numpad (OEM specific)

 

'Fujitsu/OASYS kbd definitions

Public Const VK_OEM_FJ_JISHO As Long = &H92 '"Dictionary" key (OEM specific)

Public Const VK_OEM_FJ_MASSHOU As Long = &H93 '"Unregister word" key (OEM specific)

Public Const VK_OEM_FJ_TOUROKU As Long = &H94 '"Register word" key (OEM specific)

Public Const VK_OEM_FJ_LOYA As Long = &H95 '"Left OYAYUBI" key (OEM specific)

Public Const VK_OEM_FJ_ROYA As Long = &H96 '"Right OYAYUBI" key (OEM specific)

 

'&H97 to &H9F unassigned

 

'VK_L* & VK_R* - left and right Alt, Ctrl and Shift virtual keys.

'Used only as parameters to GetAsyncKeyState() and GetKeyState().

'No other API or message will distinguish left and right keys in this way.

Public Const VK_LSHIFT As Long = &HA0 'Left SHIFT key

Public Const VK_RSHIFT As Long = &HA1 'Right SHIFT key

Public Const VK_LCONTROL As Long = &HA2 'Left CONTROL key

Public Const VK_RCONTROL As Long = &HA3 'Right CONTROL key

Public Const VK_LMENU As Long = &HA4 'Left MENU key

Public Const VK_RMENU As Long = &HA5 'Right MENU key

 

Public Const VK_BROWSER_BACK As Long = &HA6 'Windows 2000/XP: Browser Back key

Public Const VK_BROWSER_FORWARD As Long = &HA7 'Windows 2000/XP: Browser Forward key

Public Const VK_BROWSER_REFRESH As Long = &HA8 'Windows 2000/XP: Browser Refresh key

Public Const VK_BROWSER_STOP As Long = &HA9 'Windows 2000/XP: Browser Stop key

Public Const VK_BROWSER_SEARCH As Long = &HAA 'Windows 2000/XP: Browser Search key

Public Const VK_BROWSER_FAVORITES As Long = &HAB 'Windows 2000/XP: Browser Favorites key

Public Const VK_BROWSER_HOME As Long = &HAC 'Windows 2000/XP: Browser Start and Home key

 

Public Const VK_VOLUME_MUTE As Long = &HAD 'Windows 2000/XP: Volume Mute key

Public Const VK_VOLUME_DOWN As Long = &HAE 'Windows 2000/XP: Volume Down key

Public Const VK_VOLUME_UP As Long = &HAF 'Windows 2000/XP: Volume Up key

Public Const VK_MEDIA_NEXT_TRACK As Long = &HB0 'Windows 2000/XP: Next Track key

Public Const VK_MEDIA_PREV_TRACK As Long = &HB1 'Windows 2000/XP: Previous Track key

Public Const VK_MEDIA_STOP As Long = &HB2 'Windows 2000/XP: Stop Media key

Public Const VK_MEDIA_PLAY_PAUSE As Long = &HB3 'Windows 2000/XP: Play/Pause Media key

Public Const VK_LAUNCH_MAIL As Long = &HB4 'Windows 2000/XP: Start Mail key

Public Const VK_LAUNCH_MEDIA_SELECT As Long = &HB5 'Windows 2000/XP: Select Media key

Public Const VK_LAUNCH_APP1 As Long = &HB6 'Windows 2000/XP: Start Application 1 key

Public Const VK_LAUNCH_APP2 As Long = &HB7 'Windows 2000/XP: Start Application 2 key

'&HB8 to &HB9 reserved

Public Const VK_OEM_1 As Long = &HBA 'Windows 2000/XP: For the US standard keyboard, the ';:' key (Used for miscellaneous characters; it can vary by keyboard)

Public Const VK_OEM_PLUS As Long = &HBB 'Windows 2000/XP: For any country/region, the '+' key

Public Const VK_OEM_COMMA As Long = &HBC 'Windows 2000/XP: For any country/region, the ',' key

Public Const VK_OEM_MINUS As Long = &HBD 'Windows 2000/XP: For any country/region, the '-' key

Public Const VK_OEM_PERIOD As Long = &HBE 'Windows 2000/XP: For any country/region, the '.' key

Public Const VK_OEM_2 As Long = &HBF 'Windows 2000/XP: For the US standard keyboard, the '/?' key (Used for miscellaneous characters; it can vary by keyboard)

Public Const VK_OEM_3 As Long = &HC0 'Windows 2000/XP: For the US standard keyboard, the '`~' key (Used for miscellaneous characters; it can vary by keyboard)

'&HC1 to &HD7 reserved

'&HD8 to &HDA unassigned

Public Const VK_OEM_4 As Long = &HDB 'Windows 2000/XP: For the US standard keyboard, the '[{' key (Used for miscellaneous characters; it can vary by keyboard)

Public Const VK_OEM_5 As Long = &HDC 'Windows 2000/XP: For the US standard keyboard, the '\|' key (Used for miscellaneous characters; it can vary by keyboard)

Public Const VK_OEM_6 As Long = &HDD 'Windows 2000/XP: For the US standard keyboard, the ']}' key (Used for miscellaneous characters; it can vary by keyboard)

Public Const VK_OEM_7 As Long = &HDE 'Windows 2000/XP: For the US standard keyboard, the 'single-quote/double-quote' key (Used for miscellaneous characters; it can vary by keyboard)

Public Const VK_OEM_8 As Long = &HDF 'Used for miscellaneous characters; it can vary by keyboard.

 

Public Const VK_ICO_F17 As Long = &HE0 '&HE0 reserved
'Various extended or enhanced keyboards
Public Const VK_OEM_AX As Long = &HE1 ' AX key on Japanese AX kbd (OEM specific)
Public Const VK_ICO_F18 As Long = &HE1 'OEM specific
Public Const VK_OEM_102 As Long = &HE2 ' "<>" or "\|" on RT 102-key kbd.
Public Const VK_OEM102 As Long = &HE2 ' "<>" or "\|" on RT 102-key kbd.
Public Const VK_ICO_HELP As Long = &HE3 ' Help key on ICO (OEM specific)
Public Const VK_ICO_00 As Long = &HE4 ' 00 key on ICO (OEM specific)
 
Public Const VK_PROCESSKEY As Long = &HE5 'Windows 95/98/Me, Windows NT 4.0, Windows 2000/XP: IME PROCESS key
Public Const VK_ICO_CLEAR As Long = &HE6 'OEM specific

Public Const VK_PACKET As Long = &HE7 'Windows 2000/XP: Used to pass Unicode characters as if they were keystrokes. The VK_PACKET key is the low word of a 32-bit Virtual Key value used for non-keyboard input methods. For more information, see KEYBDINPUT, SendInput, WM_KEYDOWN, and WM_KEYUP

'&HE8 unassigned

 

'Nokia/Ericsson definitions (0xE9-F5 OEM specific)

Public Const VK_OEM_RESET As Long = &HE9

Public Const VK_OEM_JUMP As Long = &HEA

Public Const VK_OEM_PA1 As Long = &HEB

Public Const VK_OEM_PA2 As Long = &HEC

Public Const VK_OEM_PA3 As Long = &HED

Public Const VK_OEM_WSCTRL As Long = &HEE

Public Const VK_OEM_CUSEL As Long = &HEF

Public Const VK_OEM_ATTN As Long = &HF0

Public Const VK_OEM_FINISH As Long = &HF1

Public Const VK_OEM_FINNISH As Long = &HF1

Public Const VK_OEM_COPY As Long = &HF2

Public Const VK_OEM_AUTO As Long = &HF3
Public Const VK_OEM_ENLW As Long = &HF4

Public Const VK_OEM_BACKTAB As Long = &HF5

 

Public Const VK_ATTN As Long = &HF6 'Attn key

Public Const VK_CRSEL As Long = &HF7 'CrSel key

Public Const VK_EXSEL As Long = &HF8 'ExSel key

Public Const VK_EREOF As Long = &HF9 'Erase EOF key

Public Const VK_PLAY As Long = &HFA 'Play key

Public Const VK_ZOOM As Long = &HFB 'Zoom key
Public Const VK_NONAME As Long = &HFC 'Reserved
Public Const VK_PA1 As Long = &HFD 'PA1 key
Public Const VK_OEM_CLEAR As Long = &HFE 'Clear key
'&HFF reserved

