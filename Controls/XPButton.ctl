VERSION 5.00
Begin VB.UserControl XPButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "XPButton.ctx":0000
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1620
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   1620
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "XPButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'XPButton Control 1.0
'====================
'
'(c) 2001, Roman Bobik (roman.bobik@aon.at, http://www.freevbcode.com/AuthorInfo.asp?AuthorID=7044)
'All methods and events are described inline.
'This button can be used and modified as you like.
'Please do not compile this control and sell it as it is.
'
'Dependencies:
'The control uses the "vbAccelerator VB6 Subclassing and Timer Assistant".
'This library is needed, so distribute it with your application.
'If it is not contained in this package please download it from
'www.vbaccelerator.com

Option Explicit

'Use subclassing to receive mouse-leave-"events"
Implements ISubclass

Private Const WM_MOUSELEAVE = &H2A3
Private Type TRACKMOUSEEVENTTYPE
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type
Private Const TME_LEAVE = &H2&

Private Declare Function TrackMouseEvent Lib "User32" (lpEventTrack As TRACKMOUSEEVENTTYPE) As Long

Private Declare Function SetCapture Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long

Private Declare Function DrawText Lib "User32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_SINGLELINE = &H20

Private Type SIZEL
    cx As Long
    cy As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Enum EXPButtonAlign
    xpbaCenter = 0
    xpbaLeftTop = 1
    xpbaRightBottom = 2
End Enum


Private mTextAlignH As EXPButtonAlign
Private mTextAlignV As EXPButtonAlign
Private mPictureAlignH As EXPButtonAlign
Private mPictureAlignV As EXPButtonAlign


Private bHovering As Boolean
Private bPressed As Boolean

Private mCaption As String
Private mEnabled As Boolean
Private mTextColor As OLE_COLOR
Private mHoverColor As OLE_COLOR

Dim m_IsPressed As Boolean
Dim m_NoDown As Boolean
Dim m_DrawNoHoverThinLine As Boolean
Dim m_Picture As StdPicture
Dim m_PictureHover As StdPicture
Dim m_PicturePressed As StdPicture

Dim m_NoRefresh As Boolean

Dim PictureSize As SIZEL
Dim PictureHoverSize As SIZEL
Dim PicturePressedSize As SIZEL

Const m_def_IsPressed = False
Const m_def_NoDown = False
Const m_def_DrawNoHoverThinLine = False

Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Tritt auf, wenn der Benutzer die Maus bewegt."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste losläßt, während ein Objekt den Fokus hat."
Event MouseEnter()
Event MouseLeave()
Event Click()
Event DblClick()

Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal Clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long

'Translates an OLE-Color (&H80xxxxxx) to the real color it is representing
Private Function GetTranslatedColor(ByVal col As Long) As Long
    TranslateColor col, 0, GetTranslatedColor
End Function

'Translated an XPButton-Align-Constant "EXPButtonAlign" to a Windows-API-DrawText-Align-Constant
Private Function GetTranslatedAlign(ByVal a As EXPButtonAlign, ByVal Horizontal As Boolean) As Long
    Select Case a
        Case EXPButtonAlign.xpbaCenter
            GetTranslatedAlign = &H1
        Case EXPButtonAlign.xpbaLeftTop
            GetTranslatedAlign = &H0
        Case EXPButtonAlign.xpbaRightBottom
            GetTranslatedAlign = &H2
    End Select
    If Not Horizontal Then
        GetTranslatedAlign = GetTranslatedAlign * 4
    End If
End Function

'Get/Set Caption. Accelerator may be specified by a leading "&"
Public Property Get Caption() As String
    Caption = mCaption
End Property
Public Property Let Caption(val As String)
    mCaption = val
    Dim zahl As Long
    Dim start As Long
    start = 1
anotherTry:
    zahl = InStr(start, val, "&")
    If zahl > 0 And zahl < Len(val) Then
        If Mid$(val, zahl + 1, 1) = "&" Then
            start = zahl + 2
            GoTo anotherTry
        End If
        UserControl.AccessKeys = Mid$(val, zahl + 1, 1)
    End If
        
    Refresh
    PropertyChanged "Caption"
End Property

'Get/Set Enabled-State.
Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property
Public Property Let Enabled(val As Boolean)
    UserControl.Enabled = val
    mEnabled = val
    Refresh
    PropertyChanged "Enabled"
End Property

'Get/Set Font-Properties
Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property
Public Property Set Font(fnt As StdFont)
    Set UserControl.Font = fnt
    Refresh
    PropertyChanged "Font"
End Property

'Get/Set text fore color
Public Property Get TextColor() As OLE_COLOR
    TextColor = mTextColor
End Property
Public Property Let TextColor(Color As OLE_COLOR)
    mTextColor = Color
    Refresh
    PropertyChanged "TextColor"
End Property

'Get/Set button hover color. Hover- and pressed-colors are determined dynamically.
'Note that the button will never be filled with this color
Public Property Get HoverColor() As OLE_COLOR
    HoverColor = mHoverColor
End Property
Public Property Let HoverColor(Color As OLE_COLOR)
    mHoverColor = Color
    Refresh
    PropertyChanged "HoverColor"
End Property

'Get/Set button background color. This is used for non-hovering and
'disabled state and to calculate a fitting hover-color.
Public Property Get BackgroundColor() As OLE_COLOR
    BackgroundColor = UserControl.BackColor
End Property
Public Property Let BackgroundColor(Color As OLE_COLOR)
    UserControl.BackColor = Color
    Refresh
    PropertyChanged "BackgroundColor"
End Property

'Must be implemented...
Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
'
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer6.EMsgResponse
'
End Property

'On Mouse-Leave
Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If hWnd = UserControl.hWnd And iMsg = WM_MOUSELEAVE Then
        bHovering = False
        Refresh
        RaiseEvent MouseLeave
    End If
End Function

'If Access-Key (Character with a leading "&") is pressed
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent Click
End Sub

Private Sub UserControl_Hide()
    If Not IsWin95 Then
        'no more listening for leave-message if control hides (only >win95)
        DetachMessage Me, UserControl.hWnd, WM_MOUSELEAVE
    End If
End Sub

Private Sub UserControl_Initialize()
    NoRefresh = True
End Sub

Private Sub UserControl_InitProperties()
    Set Font = UserControl.Ambient.Font
    Caption = UserControl.Ambient.DisplayName
    Enabled = True
    TextColor = vbButtonText
    HoverColor = vbHighlight
    BackgroundColor = &H8000000F
    Set m_Picture = Nothing
    Set m_PictureHover = Nothing
    Set m_PicturePressed = Nothing
    m_DrawNoHoverThinLine = m_def_DrawNoHoverThinLine
    m_NoDown = m_def_NoDown
    m_IsPressed = m_def_IsPressed
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Set UserControl.Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
    Caption = PropBag.ReadProperty("Caption", UserControl.Ambient.DisplayName)
    mEnabled = PropBag.ReadProperty("Enabled", True)
    mTextColor = PropBag.ReadProperty("TextColor", vbButtonText)
    mHoverColor = PropBag.ReadProperty("HoverColor", vbHighlight)
    BackgroundColor = PropBag.ReadProperty("BackgroundColor", vbButtonFace)
    m_DrawNoHoverThinLine = PropBag.ReadProperty("DrawNoHoverThinLine", m_def_DrawNoHoverThinLine)
    m_NoDown = PropBag.ReadProperty("NoDown", m_def_NoDown)
    m_IsPressed = PropBag.ReadProperty("IsPressed", m_def_IsPressed)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Set PictureHover = PropBag.ReadProperty("PictureHover", Nothing)
    Set PicturePressed = PropBag.ReadProperty("PicturePressed", Nothing)
    
    mTextAlignH = PropBag.ReadProperty("TextAlignH", xpbaCenter)
    mTextAlignV = PropBag.ReadProperty("TextAlignV", xpbaCenter)
    mPictureAlignH = PropBag.ReadProperty("PictureAlignH", xpbaCenter)
    mPictureAlignV = PropBag.ReadProperty("PictureAlignV", xpbaCenter)
End Sub

Private Sub UserControl_Show()
    If Not IsWin95 Then
        'start listen for leave-message if control shows (only >win95)
        AttachMessage Me, UserControl.hWnd, WM_MOUSELEAVE
    End If
    NoRefresh = False
    Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    PropBag.WriteProperty "Font", UserControl.Font, UserControl.Ambient.Font
    PropBag.WriteProperty "Caption", mCaption, UserControl.Ambient.DisplayName
    PropBag.WriteProperty "Enabled", mEnabled, True
    PropBag.WriteProperty "TextColor", mTextColor, vbButtonText
    PropBag.WriteProperty "HoverColor", mHoverColor, vbHighlight
    PropBag.WriteProperty "BackgroundColor", UserControl.BackColor, &H8000000F
    PropBag.WriteProperty "Picture", m_Picture, Nothing
    PropBag.WriteProperty "PictureHover", m_PictureHover, Nothing
    PropBag.WriteProperty "PicturePressed", m_PicturePressed, Nothing
    PropBag.WriteProperty "DrawNoHoverThinLine", m_DrawNoHoverThinLine, m_def_DrawNoHoverThinLine
    PropBag.WriteProperty "NoDown", m_NoDown, m_def_NoDown
    PropBag.WriteProperty "IsPressed", m_IsPressed, m_def_IsPressed
    PropBag.WriteProperty "TextAlignH", mTextAlignH, xpbaCenter
    PropBag.WriteProperty "TextAlignV", mTextAlignV, xpbaCenter
    PropBag.WriteProperty "PictureAlignH", mPictureAlignH, xpbaCenter
    PropBag.WriteProperty "PictureAlignV", mPictureAlignV, xpbaCenter
End Sub

'On double click raise doubleclick-event
Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'On Mouse-Up redraw button with pressed-state
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Not NoDown Then
        Dim bTemp As Boolean
        bTemp = bPressed
        bPressed = False
        Refresh
        If bTemp = True Then RaiseEvent Click
        If IsWin95 Then UserControl_MouseMove Button, Shift, X, Y
    End If
End Sub

'On Mouse-Move determine if mouse has entered control and redraw button if so
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Button <> 0 Then Exit Sub
    
    Dim temp As Boolean
    
    If IsWin95 Then
        'This part is needed for manually determining if mouse is over the control
        'only in win95
        DoEvents
        
        ReleaseCapture
         
        If (X < 0) Or (Y < 0) Or (X > ScaleWidth) Or (Y > ScaleHeight) Then
            temp = False
        Else
            temp = True
            SetCapture UserControl.hWnd
        End If
    Else
        'Say that we want to be notified if mouse is moved out of the control
        temp = True
        Dim ET As TRACKMOUSEEVENTTYPE
        'initialize structure
        ET.cbSize = Len(ET)
        ET.hwndTrack = UserControl.hWnd
        ET.dwFlags = TME_LEAVE
        'start the tracking
        TrackMouseEvent ET
    End If
    
    'Redraw the button if needed
    If bHovering <> temp Then
        Dim oldNoRefresh As Boolean
        oldNoRefresh = NoRefresh
        NoRefresh = True
        If temp Then RaiseEvent MouseEnter Else RaiseEvent MouseLeave
        bPressed = Button <> 0 And Not NoDown
        bHovering = temp
        If bHovering = False Then bPressed = False
        NoRefresh = oldNoRefresh
        Refresh
    End If
End Sub

'This function calculates an appropriate XP hover color for a background
'and selected color. If "pressed" is specified the color is darker.
'It depends on formulars which are near to the real ones used in Office XP
Private Function GetXPHoverColor(ByVal BackgroundColor As Long, ByVal SelColor As Long, ByVal Pressed As Boolean) As Long
    Dim r1 As Long, g1 As Long, b1 As Long, r2 As Long, g2 As Long, b2 As Long
    SelColor = GetTranslatedColor(SelColor)
    BackgroundColor = GetTranslatedColor(BackgroundColor)
    
    r1 = BackgroundColor Mod &H100
    g1 = (BackgroundColor \ &H100) Mod &H100
    b1 = BackgroundColor \ &H10000
    r2 = SelColor Mod &H100
    g2 = (SelColor \ &H100) Mod &H100
    b2 = SelColor \ &H10000
    
    If Pressed Then
        GetXPHoverColor = _
            ToBounds((r2 + r1) / 2 + (100 * (255 - r1) / 255), 0, 255) + _
            ToBounds((g2 + g1) / 2 + (100 * (255 - g1) / 255), 0, 255) * &H100 + _
            ToBounds((b2 + b1) / 2 + (100 * (255 - b1) / 255), 0, 255) * &H10000
    
    Else
        GetXPHoverColor = _
            ToBounds((r2 + r1) / 2 + (100 * (255 - r2) / 255) / 2, 0, 255) + _
            ToBounds((g2 + g1) / 2 + (100 * (255 - g2) / 255) / 2, 0, 255) * &H100 + _
            ToBounds((b2 + b1) / 2 + (100 * (255 - b2) / 255) / 2, 0, 255) * &H10000
    End If
End Function

'Refresh button if control is resized
Private Sub UserControl_Resize()
    Refresh
End Sub

'Get/Set main picture.
Public Property Get Picture() As StdPicture
Attribute Picture.VB_Description = "Gibt eine Grafik zurück, die in einem Steuerelement angezeigt werden soll, oder legt diese fest."
    Set Picture = m_Picture
End Property
Public Property Set Picture(ByVal New_Picture As StdPicture)
    Set m_Picture = New_Picture
    If New_Picture Is Nothing Then
        PictureSize.cx = 0
        PictureSize.cy = 0
    Else
        'determine picture size
        PictureSize.cx = ScaleX(New_Picture.Width, vbHimetric, vbPixels)
        PictureSize.cy = ScaleY(New_Picture.Height, vbHimetric, vbPixels)
    End If
    Refresh
    PropertyChanged "Picture"
End Property

'Get/Set the hover picture.
Public Property Get PictureHover() As StdPicture
    Set PictureHover = m_PictureHover
End Property
Public Property Set PictureHover(ByVal New_Picture As StdPicture)
    Set m_PictureHover = New_Picture
    If New_Picture Is Nothing Then
        PictureHoverSize.cx = 0
        PictureHoverSize.cy = 0
    Else
        'determine picture size
        PictureHoverSize.cx = ScaleX(New_Picture.Width, vbHimetric, vbPixels)
        PictureHoverSize.cy = ScaleY(New_Picture.Height, vbHimetric, vbPixels)
    End If
    Refresh
    PropertyChanged "PictureHover"
End Property

'Get/Set the pressed picture.
Public Property Get PicturePressed() As StdPicture
    Set PicturePressed = m_PicturePressed
End Property
Public Property Set PicturePressed(ByVal New_Picture As StdPicture)
    Set m_PicturePressed = New_Picture
    If New_Picture Is Nothing Then
        PicturePressedSize.cx = 0
        PicturePressedSize.cy = 0
    Else
        'determine picture size
        PicturePressedSize.cx = ScaleX(New_Picture.Width, vbHimetric, vbPixels)
        PicturePressedSize.cy = ScaleY(New_Picture.Height, vbHimetric, vbPixels)
    End If
    Refresh
    PropertyChanged "PicturePressed"
End Property

'Get/Set whether a border-line should be drawn if the button is not hovered
Public Property Get DrawNoHoverThinLine() As Boolean
    DrawNoHoverThinLine = m_DrawNoHoverThinLine
End Property
Public Property Let DrawNoHoverThinLine(ByVal New_DrawNoHoverThinLine As Boolean)
    m_DrawNoHoverThinLine = New_DrawNoHoverThinLine
    Refresh
    PropertyChanged "DrawNoHoverThinLine"
End Property

'On mouse-down set pressed-state and redraw
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Not m_NoDown Then bPressed = True
    If Not m_NoDown Then Refresh
End Sub

'Get/Set whether the button should be locked for press-down.
Public Property Get NoDown() As Boolean
    NoDown = m_NoDown
End Property
Public Property Let NoDown(ByVal New_NoDown As Boolean)
    m_NoDown = New_NoDown
    Refresh
    PropertyChanged "NoDown"
End Property

'Get/Set if the button should be displayed as pressed-down.
Public Property Get IsPressed() As Boolean
    IsPressed = m_IsPressed
End Property
Public Property Let IsPressed(ByVal New_IsPressed As Boolean)
    m_IsPressed = New_IsPressed
    Refresh
    PropertyChanged "IsPressed"
End Property

'Get/Set the horizontaly align of the caption.
Public Property Get TextAlignH() As EXPButtonAlign
    TextAlignH = mTextAlignH
End Property
Public Property Let TextAlignH(ByVal v As EXPButtonAlign)
    mTextAlignH = v
    Refresh
    PropertyChanged "TextAlignH"
End Property

'Get/Set the verticaly align of the caption.
Public Property Get TextAlignV() As EXPButtonAlign
    TextAlignV = mTextAlignV
End Property
Public Property Let TextAlignV(ByVal v As EXPButtonAlign)
    mTextAlignV = v
    Refresh
    PropertyChanged "TextAlignV"
End Property

'Get/Set the horizontaly align of the picture.
Public Property Get PictureAlignH() As EXPButtonAlign
    PictureAlignH = mPictureAlignH
End Property
Public Property Let PictureAlignH(ByVal v As EXPButtonAlign)
    mPictureAlignH = v
    Refresh
    PropertyChanged "PictureAlignH"
End Property

'Get/Set the verticaly align of the picture.
Public Property Get PictureAlignV() As EXPButtonAlign
    PictureAlignV = mPictureAlignV
End Property
Public Property Let PictureAlignV(ByVal v As EXPButtonAlign)
    mPictureAlignV = v
    Refresh
    PropertyChanged "PictureAlignV"
End Property

'Get/Set if refreshes should be ignored. Use this if you change many
'properties at once and don't want that the button is redrawed each
'property
Property Let NoRefresh(ByVal v As Boolean)
    m_NoRefresh = v
End Property
Property Get NoRefresh() As Boolean
    NoRefresh = m_NoRefresh
End Property

'This is where all the drawing of the button is done.
Public Function Refresh()
    If m_NoRefresh Then Exit Function

    On Error Resume Next
    Dim pct As StdPicture, pctsize As SIZEL
    
    'Determine the rectangle (in pixels) where text should be drawn
    Dim rRect As RECT
    With rRect
        .Left = 1
        .Top = 0
        .Right = ScaleWidth / Screen.TwipsPerPixelX - 2
        .Bottom = ScaleHeight / Screen.TwipsPerPixelY - 2
    End With
    
    'same rectangle as rRect but moved one pixel down-right
    Dim rRectPlus As RECT
    rRectPlus = rRect
    With rRectPlus
        .Left = .Left + 1
        .Top = .Top + 1
        .Right = .Right + 1
        .Bottom = .Bottom + 1
    End With
    
    'This determines which picture to use
    If bPressed Or (m_IsPressed And Not bHovering) Then
        If m_PicturePressed Is Nothing Then
            If bHovering Then
                If m_PictureHover Is Nothing Then
                    Set pct = m_Picture
                    pctsize = PictureSize
                Else
                    Set pct = m_PictureHover
                    pctsize = PictureHoverSize
                End If
            Else
                Set pct = m_Picture
                pctsize = PictureSize
            End If
        Else
            Set pct = m_PicturePressed
            pctsize = PicturePressedSize
        End If
    Else
        If bHovering Then
            If m_PictureHover Is Nothing Then
                Set pct = m_Picture
                pctsize = PictureSize
            Else
                Set pct = m_PictureHover
                pctsize = PictureHoverSize
            End If
        Else
            Set pct = m_Picture
            pctsize = PictureSize
        End If
    End If
    
    'Resize the temporary picture box which is used to fill the
    'transparent regions of an picture with an background-color
    If Not pct Is Nothing Then Picture1.Move 0, 0, pctsize.cx * Screen.TwipsPerPixelX, pctsize.cy * Screen.TwipsPerPixelY
    
    
    'Set the point where the upper left corner of the picture should be
    '(depending on the selected alignment)
    Dim pctTopLeft As SIZEL
    With pctTopLeft
        Select Case PictureAlignH
            Case EXPButtonAlign.xpbaCenter
                .cx = Round((ScaleWidth - PictureSize.cx * Screen.TwipsPerPixelX) / 2)
            Case EXPButtonAlign.xpbaLeftTop
                .cx = 1 * Screen.TwipsPerPixelX
            Case EXPButtonAlign.xpbaRightBottom
                .cx = ScaleWidth - 2 * Screen.TwipsPerPixelX - PictureSize.cx * Screen.TwipsPerPixelX
        End Select
                
        Select Case PictureAlignV
            Case EXPButtonAlign.xpbaCenter
                .cy = Round((ScaleHeight - PictureSize.cy * Screen.TwipsPerPixelY) / 2)
            Case EXPButtonAlign.xpbaLeftTop
                .cy = 1 * Screen.TwipsPerPixelY
            Case EXPButtonAlign.xpbaRightBottom
                .cy = ScaleHeight - 2 * Screen.TwipsPerPixelY - PictureSize.cy * Screen.TwipsPerPixelY
        End Select
    End With
    
    'If in design-mode: always draw in hover-state
    If Not UserControl.Ambient.UserMode Then bHovering = True
    
    Picture1.Picture = pct
    If mEnabled Then 'If enabled state should be drawn
        If bPressed Or (m_IsPressed And Not bHovering) Then 'if pressed state
            UserControl.Line (1, 1)-(UserControl.ScaleWidth - 2 * Screen.TwipsPerPixelX, UserControl.ScaleHeight - 2 * Screen.TwipsPerPixelY), GetXPHoverColor(UserControl.BackColor, mHoverColor, True), BF
            
            With rRect
                'Fill transparent regions of the picture with xp-background-color
                Picture1.BackColor = GetXPHoverColor(UserControl.BackColor, mHoverColor, True)
                If Not pct Is Nothing Then UserControl.PaintPicture Picture1.Image, _
                    pctTopLeft.cx + 1 * Screen.TwipsPerPixelX, pctTopLeft.cy + 1 * Screen.TwipsPerPixelY
            End With
            UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1 * Screen.TwipsPerPixelX, UserControl.ScaleHeight - 1 * Screen.TwipsPerPixelY), mHoverColor, B
            UserControl.ForeColor = mTextColor
            
            DrawText UserControl.hdc, mCaption, Len(mCaption), rRectPlus, _
                GetTranslatedAlign(mTextAlignH, True) Or GetTranslatedAlign(mTextAlignV, False) Or DT_SINGLELINE
        Else 'if non-pressed
            If bHovering Then 'if hovering, set background color to xp
                Picture1.BackColor = GetXPHoverColor(UserControl.BackColor, mHoverColor, False)
                UserControl.Line (1, 1)-(UserControl.ScaleWidth - 2 * Screen.TwipsPerPixelX, UserControl.ScaleHeight - 2 * Screen.TwipsPerPixelY), GetXPHoverColor(UserControl.BackColor, mHoverColor, False), BF
            Else
                Picture1.BackColor = UserControl.BackColor
                UserControl.Line (1, 1)-(UserControl.ScaleWidth - 2 * Screen.TwipsPerPixelX, UserControl.ScaleHeight - 2 * Screen.TwipsPerPixelY), UserControl.BackColor, BF
            End If
            
            UserControl.ForeColor = mTextColor
            If Not pct Is Nothing Then UserControl.PaintPicture Picture1.Image, _
                    pctTopLeft.cx, pctTopLeft.cy
            DrawText UserControl.hdc, mCaption, Len(mCaption), rRect, _
                GetTranslatedAlign(mTextAlignH, True) Or GetTranslatedAlign(mTextAlignV, False) Or DT_SINGLELINE
        
            If bHovering Then
                UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1 * Screen.TwipsPerPixelX, UserControl.ScaleHeight - 1 * Screen.TwipsPerPixelY), mHoverColor, B
            Else
                UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1 * Screen.TwipsPerPixelX, UserControl.ScaleHeight - 1 * Screen.TwipsPerPixelY), IIf(m_DrawNoHoverThinLine, &H80000010, UserControl.BackColor), B
            End If
        
        End If
    Else
        UserControl.Line (1 * Screen.TwipsPerPixelX, 1 * Screen.TwipsPerPixelY)-(UserControl.ScaleWidth - 2 * Screen.TwipsPerPixelX, UserControl.ScaleHeight - 2 * Screen.TwipsPerPixelY), UserControl.BackColor, BF
        
        'Draw the caption two times: once dark, once white to get a
        'disabled-text-state
        UserControl.ForeColor = vb3DHighlight
        DrawText UserControl.hdc, mCaption, Len(mCaption), rRectPlus, _
            GetTranslatedAlign(mTextAlignH, True) Or GetTranslatedAlign(mTextAlignV, False) Or DT_SINGLELINE
        
        UserControl.ForeColor = vb3DShadow
        DrawText UserControl.hdc, mCaption, Len(mCaption), rRect, _
            GetTranslatedAlign(mTextAlignH, True) Or GetTranslatedAlign(mTextAlignV, False) Or DT_SINGLELINE
    
    
        Picture1.BackColor = UserControl.BackColor
        If Not pct Is Nothing Then UserControl.PaintPicture Picture1.Image, _
                pctTopLeft.cx, pctTopLeft.cy
    
        If m_DrawNoHoverThinLine Then UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1 * Screen.TwipsPerPixelX, UserControl.ScaleHeight - 1 * Screen.TwipsPerPixelY), &H80000010, B
    
        UserControl.ForeColor = mTextColor
    End If
End Function
