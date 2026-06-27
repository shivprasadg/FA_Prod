Attribute VB_Name = "WindowsSizingMetrics"
Option Compare Database
#If VBA7 Then ' Win API Declarations for 32 and 64 bit versions of Office 2010 and later

Declare PtrSafe Function GetSystemMetrics& Lib "user32" (ByVal nIndex&)
Const SM_CXSCREEN = 0        ' Width of screen
Const SM_CYSCREEN = 1        ' Height of screen
Const SM_CXFULLSCREEN = 16   ' Width of window client area
Const SM_CYFULLSCREEN = 17   ' Height of window client area
Const SM_CYMENU = 15         ' Height of menu
Const SM_CYCAPTION = 4       ' Height of caption or title
Const SM_CXFRAME = 32        ' Width of window frame
Const SM_CYFRAME = 33        ' Height of window frame
Const SM_CXHSCROLL = 21      ' Width of arrow bitmap on
                             '  horizontal scroll bar
Const SM_CYHSCROLL = 3       ' Height of arrow bitmap on
                             '  horizontal scroll bar
Const SM_CXVSCROLL = 2       ' Width of arrow bitmap on
                             '  vertical scroll bar
Const SM_CYVSCROLL = 20      ' Height of arrow bitmap on
                             '  vertical scroll bar
Const SM_CXSIZE = 30         ' Width of bitmaps in title bar
Const SM_CYSIZE = 31         ' Height of bitmaps in title bar
Const SM_CXCURSOR = 13       ' Width of cursor
Const SM_CYCURSOR = 14       ' Height of cursor
Const SM_CXBORDER = 5        ' Width of window frame that cannot
                             '  be sized
Const SM_CYBORDER = 6        ' Height of window frame that cannot
                             '  be sized
Const SM_CXDOUBLECLICK = 36  ' Width of rectangle around the
                             '  location of the first click. The
                             '  second click must occur in the
                             '  same rectangular location.
Const SM_CYDOUBLECLICK = 37  ' Height of rectangle around the
                             '  location of the first click. The
                             '  second click must occur in the
                             '  same rectangular location.
Const SM_CXDLGFRAME = 7      ' Width of dialog frame window
Const SM_CYDLGFRAME = 8      ' Height of dialog frame window
Const SM_CXICON = 11         ' Width of icon
Const SM_CYICON = 12         ' Height of icon
Const SM_CXICONSPACING = 38  ' Width of rectangles the system
                             ' uses to position tiled icons
Const SM_CYICONSPACING = 39  ' Height of rectangles the system
                             ' uses to position tiled icons
Const SM_CXMIN = 28          ' Minimum width of window
Const SM_CYMIN = 29          ' Minimum height of window
Const SM_CXMINTRACK = 34     ' Minimum tracking width of window
Const SM_CYMINTRACK = 35     ' Minimum tracking height of window
Const SM_CXHTHUMB = 10       ' Width of scroll box (thumb) on
                             '  horizontal scroll bar
Const SM_CYVTHUMB = 9        ' Width of scroll box (thumb) on
                             '  vertical scroll bar
Const SM_DBCSENABLED = 42    ' Returns a non-zero if the current
                             '  Windows version uses double-byte
                             '  characters, otherwise returns
                             '  zero
Const SM_DEBUG = 22          ' Returns non-zero if the Windows
                             '  version is a debugging version
Const SM_MENUDROPALIGNMENT = 40
                             ' Alignment of pop-up menus. If zero,
                             '  left side is aligned with
                             '  corresponding left side of menu-
                             '  bar item. If non-zero, left side
                             '  is aligned with right side of
                             '  corresponding menu bar item
Const SM_MOUSEPRESENT = 19   ' Non-zero if mouse hardware is
                             '  installed
Const SM_PENWINDOWS = 41     ' Handle of Pen Windows dynamic link
                             '  library if Pen Windows is
                             '  installed
Const SM_SWAPBUTTON = 23     ' Non-zero if the left and right
                             ' mouse buttons are swapped
#Else

Declare Function GetSystemMetrics& Lib "user32" (ByVal nIndex&)
Const SM_CXSCREEN = 0        ' Width of screen
Const SM_CYSCREEN = 1        ' Height of screen
Const SM_CXFULLSCREEN = 16   ' Width of window client area
Const SM_CYFULLSCREEN = 17   ' Height of window client area
Const SM_CYMENU = 15         ' Height of menu
Const SM_CYCAPTION = 4       ' Height of caption or title
Const SM_CXFRAME = 32        ' Width of window frame
Const SM_CYFRAME = 33        ' Height of window frame
Const SM_CXHSCROLL = 21      ' Width of arrow bitmap on
                             '  horizontal scroll bar
Const SM_CYHSCROLL = 3       ' Height of arrow bitmap on
                             '  horizontal scroll bar
Const SM_CXVSCROLL = 2       ' Width of arrow bitmap on
                             '  vertical scroll bar
Const SM_CYVSCROLL = 20      ' Height of arrow bitmap on
                             '  vertical scroll bar
Const SM_CXSIZE = 30         ' Width of bitmaps in title bar
Const SM_CYSIZE = 31         ' Height of bitmaps in title bar
Const SM_CXCURSOR = 13       ' Width of cursor
Const SM_CYCURSOR = 14       ' Height of cursor
Const SM_CXBORDER = 5        ' Width of window frame that cannot
                             '  be sized
Const SM_CYBORDER = 6        ' Height of window frame that cannot
                             '  be sized
Const SM_CXDOUBLECLICK = 36  ' Width of rectangle around the
                             '  location of the first click. The
                             '  second click must occur in the
                             '  same rectangular location.
Const SM_CYDOUBLECLICK = 37  ' Height of rectangle around the
                             '  location of the first click. The
                             '  second click must occur in the
                             '  same rectangular location.
Const SM_CXDLGFRAME = 7      ' Width of dialog frame window
Const SM_CYDLGFRAME = 8      ' Height of dialog frame window
Const SM_CXICON = 11         ' Width of icon
Const SM_CYICON = 12         ' Height of icon
Const SM_CXICONSPACING = 38  ' Width of rectangles the system
                             ' uses to position tiled icons
Const SM_CYICONSPACING = 39  ' Height of rectangles the system
                             ' uses to position tiled icons
Const SM_CXMIN = 28          ' Minimum width of window
Const SM_CYMIN = 29          ' Minimum height of window
Const SM_CXMINTRACK = 34     ' Minimum tracking width of window
Const SM_CYMINTRACK = 35     ' Minimum tracking height of window
Const SM_CXHTHUMB = 10       ' Width of scroll box (thumb) on
                             '  horizontal scroll bar
Const SM_CYVTHUMB = 9        ' Width of scroll box (thumb) on
                             '  vertical scroll bar
Const SM_DBCSENABLED = 42    ' Returns a non-zero if the current
                             '  Windows version uses double-byte
                             '  characters, otherwise returns
                             '  zero
Const SM_DEBUG = 22          ' Returns non-zero if the Windows
                             '  version is a debugging version
Const SM_MENUDROPALIGNMENT = 40
                             ' Alignment of pop-up menus. If zero,
                             '  left side is aligned with
                             '  corresponding left side of menu-
                             '  bar item. If non-zero, left side
                             '  is aligned with right side of
                             '  corresponding menu bar item
Const SM_MOUSEPRESENT = 19   ' Non-zero if mouse hardware is
                             '  installed
Const SM_PENWINDOWS = 41     ' Handle of Pen Windows dynamic link
                             '  library if Pen Windows is
                             '  installed
Const SM_SWAPBUTTON = 23     ' Non-zero if the left and right
                             ' mouse buttons are swapped
#End If
