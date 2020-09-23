VERSION 5.00
Begin VB.UserControl Button 
   ClientHeight    =   3045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   EditAtDesignTime=   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3045
   ScaleWidth      =   4350
   ToolboxBitmap   =   "BUTTON.ctx":0000
   Begin VB.Timer tmrToolTip 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   90
      Top             =   1770
   End
   Begin VB.PictureBox picDisp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Tag             =   "2"
      Top             =   105
      Width           =   240
   End
   Begin VB.PictureBox picOver 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1620
      Picture         =   "BUTTON.ctx":0312
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Tag             =   "2"
      Top             =   2010
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picDown 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2340
      Picture         =   "BUTTON.ctx":045C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Tag             =   "2"
      Top             =   2010
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picUp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   840
      Picture         =   "BUTTON.ctx":05A6
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Tag             =   "2"
      Top             =   2010
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblDisp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   660
      TabIndex        =   7
      Top             =   120
      Width           =   45
   End
   Begin VB.Label lblOver 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblOver"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1470
      TabIndex        =   6
      Top             =   1650
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblDown 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblDown"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2250
      TabIndex        =   5
      Top             =   1650
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblUp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblUp"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   780
      TabIndex        =   4
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Bouton paramètrable, sonore..."
'=============================
'
'Nom Du Projet: BioRythmes
'
'Auteur:Les Productions J.F.
'
'=============================
Option Explicit
Private Type TYPEPOINT
    x As Long
    y As Long
End Type
Private Type TYPERECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Enum BtnBackStyle
    Transparent = 0
    Opaque = 1
End Enum
Public Enum BtnAppearance
    Flat = 0
    HalfRaised = 1
    Raised = 2
End Enum
Public Enum BtnContentEffects
    bceYes = 0
    bceNo = 1
End Enum
Public Enum BtnSunkenEffects
    bseYes = 0
    bseNo = 1
End Enum
Public Enum picAutoFit
    pafYes = 0
    pafNo = 1
End Enum
Public Enum picDefaultArrange
    pdaLeft = 0
    pdaTop = 1
End Enum
Public Enum Sound
    None = 0
    Whoosh = 1
    Laser = 2
End Enum

Private Const SND_SYNC = &H1                      ' Jouer de façon synchrone, et ASYNC de façon asyncrone '
Private Const SND_NODEFAULT = &H2                 ' Ne pas utiliser le son par défaut.                    '
Private Const SND_MEMORY = &H4                    ' lpszSoundName pointe vers un fichier en mémoire.      '
Private SoundBuffer() As Byte

Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_FLAT = &H4000
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As TYPERECT, _
    ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, _
    lpPoint As TYPEPOINT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As TYPEPOINT) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) _
    As Long
Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Private Const HWND_TOP = 0
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
    lpRect As TYPERECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Dim rRed As Integer, rGreen As Integer, rBlue As Integer



Event Click()
Event ClickSunken()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Const defBackColor = &HC0C0C0
Const defBackColorOver = &HC0C0C0
Const defBackColorDown = &HC0C0C0
Const defCapColor = &H0
Const defCapColorOver = &H0
Const defCapColorDown = &H0
Const defPicAutoFit = 1
Const defPicDefaultArrange = 0
Const defBtnAppearance = 1
Const defBtnSunkenEffects = 0
Const defBtnContentEffects = 0
Const defPicCapSpacing = 75
Const defBtnWidth = 0
Const defBtnHeight = 0
Const defPicPosLeft = 0
Const defPicPosTop = 0
Const defCapPosLeft = 0
Const defCapPosTop = 0
Dim m_picAutoFit As picAutoFit
Dim m_BtnAppearance As BtnAppearance
Dim m_BtnSunkenEffects As BtnSunkenEffects
Dim m_BtnContentEffects As BtnContentEffects
Dim m_PicDefaultArrange As picDefaultArrange
Dim m_PicPosLeft As Long
Dim m_PicPosTop As Long
Dim m_CapPosLeft As Long
Dim m_CapPosTop As Long
Dim m_PicCapSpacing As Long
Dim m_CaptionDown As String
Dim m_BackColor As OLE_COLOR
Dim m_BackColorOver As OLE_COLOR
Dim m_BackColorDown As OLE_COLOR
Dim m_ToolTipBackColor As OLE_COLOR
Dim m_ToolTipForeColor As OLE_COLOR
Dim m_ToolTipText As String
Dim BtnAppearanceSet As Long
Dim IsSunken As Boolean
Dim OrigCaption As String
Dim ImageUpLoaded As Boolean
Dim ImageUpDownLoaded As Boolean
Dim EdgePerimeter As Integer
Dim ArePropertiesRead As Boolean
Dim DownWard As Boolean
Dim AncienneCouleur
Dim FichierSon As Sound
Dim Dedans As Boolean



Private Sub UserControl_InitProperties()
    lblDisp.Caption = Extender.Name
    m_BtnAppearance = defBtnAppearance
    m_BtnContentEffects = defBtnContentEffects
    m_BtnSunkenEffects = defBtnSunkenEffects
    m_PicDefaultArrange = defPicDefaultArrange
    m_PicCapSpacing = defPicCapSpacing
    m_picAutoFit = defPicAutoFit
    m_BackColor = defBackColor
    m_BackColorOver = defBackColorOver
    m_BackColorDown = defBackColorDown
    m_ToolTipBackColor = frmTooltip.lblDefToolTipColor.BackColor
    m_ToolTipForeColor = frmTooltip.lblDefToolTipColor.ForeColor
    m_ToolTipText = Extender.ToolTipText
    picUp.Picture = LoadPicture()
    picDown.Picture = LoadPicture()
    picOver.Picture = LoadPicture()
End Sub
Private Sub UserControl_Initialize()
    Dedans = False
    If UserControl.ScaleMode = 1 Then
        EdgePerimeter = 75
    ElseIf UserControl.ScaleMode = 3 Then
        EdgePerimeter = 5
    End If
    If UserControl.ScaleWidth < EdgePerimeter Then
        UserControl.ScaleWidth = EdgePerimeter + 1
    ElseIf UserControl.ScaleHeight < EdgePerimeter Then
        UserControl.ScaleHeight = EdgePerimeter + 1
    End If
    lblDisp.ForeColor = lblUp.ForeColor
    IsSunken = False
    ImageUpLoaded = (picUp.Picture > 0 And picDisp.Picture > 0)
    ImageUpDownLoaded = (picUp.Picture > 0 And picDown.Picture > 0)
    If picUp.Picture = 0 Then
        picDisp.Picture = LoadPicture()
        picDisp.Visible = False
    Else
        picDisp.Picture = picUp.Picture
        picDisp.BackColor = UserControl.BackColor
        picDisp.Visible = True
    End If
    CenterBtnContents
End Sub
Private Sub UserControl_Paint()
    PaintUserControl
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.BackStyle = PropBag.ReadProperty("BtnBackStyle", 1)
    UserControl.Width = PropBag.ReadProperty("BtnWidth", defBtnWidth)
    UserControl.Height = PropBag.ReadProperty("BtnHeight", defBtnHeight)
    m_BackColor = PropBag.ReadProperty("BackColor", defBackColor)
    m_BackColorOver = PropBag.ReadProperty("BackColorOver", defBackColorOver)
    m_BackColorDown = PropBag.ReadProperty("BackColorDown", defBackColorDown)
    m_picAutoFit = PropBag.ReadProperty("PicAutoFit", defPicAutoFit)
    m_PicDefaultArrange = PropBag.ReadProperty("PicDefaultArrange", defPicDefaultArrange)
    m_PicCapSpacing = PropBag.ReadProperty("PicCapSpacing", defPicCapSpacing)
    m_BtnAppearance = PropBag.ReadProperty("BtnAppearance", defBtnAppearance)
    m_BtnSunkenEffects = PropBag.ReadProperty("BtnSunkenEffects", defBtnSunkenEffects)
    m_BtnContentEffects = PropBag.ReadProperty("BtnContentEffects", defBtnContentEffects)
    Set lblDisp.Font = PropBag.ReadProperty("CaptionFont", Ambient.Font)
    lblDisp.Caption = PropBag.ReadProperty("Caption", Extender.Name)
    m_CaptionDown = PropBag.ReadProperty("CaptionDown", "")
    lblUp.ForeColor = PropBag.ReadProperty("CapColor", defCapColor)
    AncienneCouleur = PropBag.ReadProperty("CapColor", defCapColor)
    lblDisp.ForeColor = PropBag.ReadProperty("CapColor", defCapColor)
    lblOver.ForeColor = PropBag.ReadProperty("CapColorOver", defCapColorOver)
    lblDown.ForeColor = PropBag.ReadProperty("CapColorDown", defCapColorDown)
    Set picDisp.Picture = PropBag.ReadProperty("PictureDisp", Nothing)
    Set picUp.Picture = PropBag.ReadProperty("PictureUp", Nothing)
    Set picOver.Picture = PropBag.ReadProperty("PictureOver", Nothing)
    Set picDown.Picture = PropBag.ReadProperty("PictureDown", Nothing)
    picDisp.Left = PropBag.ReadProperty("PicPosLeft", defPicPosLeft)
    picDisp.Top = PropBag.ReadProperty("PicPosTop", defPicPosTop)
    lblDisp.Left = PropBag.ReadProperty("CapPosLeft", defCapPosLeft)
    lblDisp.Top = PropBag.ReadProperty("CapPosTop", defCapPosTop)
    m_ToolTipBackColor = PropBag.ReadProperty("ToolTipBackColor", Extender.ToolTipBackColor)
    m_ToolTipForeColor = PropBag.ReadProperty("ToolTipForeColor", Extender.ToolTipForeColor)
    m_ToolTipText = PropBag.ReadProperty("ToolTipText", Extender.ToolTipText)
    OrigCaption = lblDisp.Caption
    FichierSon = PropBag.ReadProperty("SoundToPlay", 0)
    If ArePropertiesRead = False Then
        ArePropertiesRead = True
        BtnAppearanceSet = m_BtnAppearance
        UserControl.BackColor = m_BackColor
        If UserControl.Enabled = False Then
            lblDisp.Enabled = False
            GreyPic
        Else
            lblDisp.Enabled = True
            Set picDisp.Picture = PropBag.ReadProperty("PictureDisp", Nothing)
        End If
        CenterBtnContents
    End If


End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("BtnBackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BtnWidth", UserControl.Width, defBtnWidth)
    Call PropBag.WriteProperty("BtnHeight", UserControl.Height, defBtnHeight)
    Call PropBag.WriteProperty("BackColor", m_BackColor, Ambient.BackColor)
    Call PropBag.WriteProperty("BackColorOver", m_BackColorOver, Ambient.BackColor)
    Call PropBag.WriteProperty("BackColorDown", m_BackColorDown, Ambient.BackColor)
    Call PropBag.WriteProperty("PicAutoFit", m_picAutoFit, defPicAutoFit)
    Call PropBag.WriteProperty("PicDefaultArrange", m_PicDefaultArrange, defPicDefaultArrange)
    Call PropBag.WriteProperty("PicCapSpacing", m_PicCapSpacing, defPicCapSpacing)
    Call PropBag.WriteProperty("BtnAppearance", m_BtnAppearance, defBtnAppearance)
    Call PropBag.WriteProperty("BtnSunkenEffects", m_BtnSunkenEffects, defBtnSunkenEffects)
    Call PropBag.WriteProperty("BtnContentEffects", m_BtnContentEffects, defBtnContentEffects)
    Call PropBag.WriteProperty("CaptionFont", lblDisp.Font, 0)
    Call PropBag.WriteProperty("Caption", lblDisp.Caption, Extender.Name)
    Call PropBag.WriteProperty("CaptionDown", m_CaptionDown, "")
    Call PropBag.WriteProperty("CapColor", lblDisp.ForeColor, Ambient.ForeColor)
    Call PropBag.WriteProperty("CapColor", lblUp.ForeColor, defCapColor)
    Call PropBag.WriteProperty("CapColorOver", lblOver.ForeColor, defCapColorOver)
    Call PropBag.WriteProperty("CapColorDown", lblDown.ForeColor, defCapColorDown)
    Call PropBag.WriteProperty("PictureDisp", picUp.Picture, Nothing)
    Call PropBag.WriteProperty("PictureUp", picUp.Picture, Nothing)
    Call PropBag.WriteProperty("PictureOver", picOver.Picture, Nothing)
    Call PropBag.WriteProperty("PictureDown", picDown.Picture, Nothing)
    Call PropBag.WriteProperty("PicPosLeft", m_PicPosLeft, defPicPosLeft)
    Call PropBag.WriteProperty("PicPosTop", m_PicPosTop, defPicPosTop)
    Call PropBag.WriteProperty("CapPosLeft", m_CapPosLeft, defCapPosLeft)
    Call PropBag.WriteProperty("CapPosTop", m_CapPosTop, defCapPosTop)
    Call PropBag.WriteProperty("ToolTipBackColor", m_ToolTipBackColor, Ambient.ForeColor)
    Call PropBag.WriteProperty("ToolTipForeColor", m_ToolTipForeColor, Ambient.ForeColor)
    Call PropBag.WriteProperty("ToolTipText", Extender.ToolTipText, m_ToolTipText)
    Call PropBag.WriteProperty("SoundToPlay", FichierSon, 0)
End Sub
Private Sub PaintUserControl()
    On Error Resume Next
    Dim typRect As TYPERECT
    Dim origScaleMode
    origScaleMode = UserControl.ScaleMode
    UserControl.ScaleMode = vbPixels
    UserControl.Cls
    With typRect
        .Left = UserControl.ScaleLeft
        .Top = UserControl.ScaleTop
        .Right = UserControl.ScaleWidth
        .Bottom = UserControl.ScaleHeight
    End With
    Select Case m_BtnAppearance
        Case 0
            DrawEdge hdc, typRect, EDGE_SUNKEN, BF_FLAT
        Case 1
            DrawEdge hdc, typRect, BDR_RAISEDINNER, BF_RECT
        Case 2
            DrawEdge hdc, typRect, EDGE_RAISED, BF_RECT
        Case 3
            DrawEdge hdc, typRect, EDGE_SUNKEN, BF_RECT
    End Select
    UserControl.ScaleMode = origScaleMode
    If UserControl.AutoRedraw Then
        UserControl.Refresh
    End If
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, _
    y As Single)
    UnallowToolTip
    If Button <> vbLeftButton Then
        Exit Sub
    End If
    If m_BtnSunkenEffects = 1 Then
        doSunkenAppearance
    Else
        toggleClickFunction
    End If
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub
Private Sub lblDisp_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseDown Button, Shift, x, y
End Sub
Private Sub picDisp_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseDown Button, Shift, x, y
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Dedans = False Then
        Dedans = True
        If SoundToPlay <> None Then
            BeginPlaySound (SoundToPlay)
        End If
    End If
    If GetCapture() <> UserControl.hwnd Then
        SetCapture UserControl.hwnd
    End If
    If x < 0 Or x > UserControl.ScaleWidth Or y < 0 Or y > UserControl.ScaleHeight Then
        ' doOrigAppearance '
        ReleaseCapture
        Dedans = False
    End If
    If Len(Extender.ToolTipText) > 0 Then
        If IsSunken = False Then
            ShowToolTip Extender.ToolTipText
        Else
            UnallowToolTip
        End If
    Else
        UnallowToolTip
    End If
    If m_BtnContentEffects = 0 Then
        If x < 0 Or x > UserControl.ScaleWidth Or y < 0 Or y > UserControl.ScaleHeight Then
            doMouseOverEffect False
        Else
            doMouseOverEffect True
            PaintUserControl
        End If
    End If
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub
Private Sub lblDisp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseMove Button, Shift, x, y
End Sub
Private Sub picDisp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseMove Button, Shift, x, y
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> vbLeftButton Then
        Exit Sub
    End If
    RaiseEvent MouseUp(Button, Shift, x, y)
    If m_BtnSunkenEffects = 0 Then
        If IsSunken Then
            RaiseEvent Click
        Else
            RaiseEvent ClickSunken
        End If
    Else
        RaiseEvent Click
    End If
    If m_BtnSunkenEffects = 1 Then
        doOrigAppearance
    End If
    doWorkAround
End Sub
Private Sub lblDisp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseUp Button, Shift, x, y
End Sub
Private Sub picDisp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseUp Button, Shift, x, y
End Sub
Private Sub doWorkAround()
    Dim typPoint As TYPEPOINT
    ClientToScreen UserControl.hwnd, typPoint
    GetCursorPos typPoint
    If DownWard = True Then
        typPoint.x = typPoint.x + 3
        typPoint.y = typPoint.y + 3
        DownWard = False
    Else
        typPoint.x = typPoint.x - 3
        typPoint.y = typPoint.y - 3
        DownWard = True
    End If
    SetCursorPos typPoint.x, typPoint.y
End Sub
Private Sub UserControl_Resize()
    On Error Resume Next
    If UserControl.ScaleMode = 1 Then
        If UserControl.Width > 4000 Then UserControl.Width = 4000
        If UserControl.Height > 2000 Then UserControl.Height = 2000
    ElseIf UserControl.ScaleMode = 3 Then
        If UserControl.Width > 300 Then UserControl.Width = 300
        If UserControl.Height > 150 Then UserControl.Height = 150
    End If
    PaintUserControl
    CenterBtnContents
End Sub
Private Sub tmrToolTip_Timer()
    On Error Resume Next
    If GetCapture() = UserControl.hwnd Then
        doToolTip tmrToolTip.Tag
    Else
        UnallowToolTip
    End If
End Sub
Private Sub UnallowToolTip()
    On Error Resume Next
    RemoveToolTip
    tmrToolTip.Enabled = False
End Sub
Private Sub RemoveToolTip()
    On Error Resume Next
    Unload frmTooltip
End Sub
Private Sub ShowToolTip(ByVal inText As String)
    On Error Resume Next
    If inText = "" Then
        tmrToolTip.Enabled = False
    Else
        tmrToolTip.Enabled = True
        tmrToolTip.Tag = inText
    End If
End Sub
Private Sub doToolTip(ByVal inText)
    On Error GoTo errHandler
    Dim x As Long, y As Long
    Dim adjW As Long, adjH As Long
    Dim textW As Long, textH As Long
    Dim typRect As TYPERECT
    Dim i As Integer
    GetWindowRect UserControl.hwnd, typRect
    x = (typRect.Left + (typRect.Right - typRect.Left) / 3) * Screen.TwipsPerPixelX
    y = (typRect.Bottom + 8) * Screen.TwipsPerPixelY
    adjW = 10 * Screen.TwipsPerPixelX
    adjH = 8 * Screen.TwipsPerPixelY
    i = frmTooltip.TextWidth(inText)
    Do While i > (Screen.Width * 80 / 100)
        inText = Left(inText, Len(inText) - 1)
        i = frmTooltip.TextWidth(inText)
    Loop
    textW = frmTooltip.TextWidth(inText) + adjW
    textH = frmTooltip.TextHeight(inText) + adjH
    If x < 0 Then
        x = 0
    ElseIf (x + textW) > Screen.Width Then
        x = Screen.Width - textW
    End If
    If (y + textH) > Screen.Height Then
        y = (typRect.Top - 2) * Screen.TwipsPerPixelY - textH
    End If
    With frmTooltip
        .BackColor = m_ToolTipBackColor
        .lblToolTipText.Width = textW
        .lblToolTipText.Height = textH
        .lblToolTipText.BackColor = m_ToolTipBackColor
        .lblToolTipText.ForeColor = m_ToolTipForeColor
        .lblToolTipText.Caption = inText
        .lblToolTipText.Refresh
        .Move x, y, textW, textH
        SetWindowPos .hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or _
            SWP_NOSIZE Or SWP_SHOWWINDOW
    End With
    Exit Sub
errHandler:
End Sub
Private Sub CenterBtnContents()
    Dim w1 As Long, h1 As Long
    Dim w2 As Long, h2 As Long
    Dim w3 As Long, h3 As Long
    Dim ctlW As Long, ctlH As Long
    Dim x As Long, y As Long
    Dim oldPicPosLeft As Long, oldPicPosTop As Long
    Dim oldCapPosLeft, oldCapPosTop As Long
    oldPicPosLeft = picDisp.Left
    oldPicPosTop = picDisp.Top
    oldCapPosLeft = lblDisp.Left
    oldCapPosTop = lblDisp.Top
    ImageUpLoaded = picDisp.Picture > 0 And picUp.Picture > 0
    ImageUpDownLoaded = (picUp.Picture > 0 And picDown.Picture > 0)
    w1 = 0: w2 = 0: w3 = 0: h1 = 0: h2 = 0: h3 = 0
    If ImageUpLoaded = True Then
        If UserControl.ScaleMode = 3 Then
            w1 = picDisp.ScaleWidth * Screen.TwipsPerPixelX
            h1 = picDisp.ScaleHeight * Screen.TwipsPerPixelY
        Else
            w1 = picDisp.ScaleWidth
            h1 = picDisp.ScaleHeight
        End If
        If lblDisp.Caption <> "" Then
            w2 = m_PicCapSpacing
            h2 = m_PicCapSpacing
        End If
        picDisp.Visible = True
        If picDisp.BackColor <> UserControl.BackColor Then
            picDisp.BackColor = UserControl.BackColor
        End If
        If picUp.BackColor <> UserControl.BackColor Then
            picUp.BackColor = UserControl.BackColor
        End If
        If m_BtnContentEffects = 0 Then
            If picOver.BackColor <> UserControl.BackColor Then
                picOver.BackColor = UserControl.BackColor
            End If
        Else
            If picOver.BackColor <> m_BackColorOver Then
                picOver.BackColor = m_BackColorOver
            End If
        End If
        If m_BtnContentEffects = 0 Then
            If picDown.BackColor <> UserControl.BackColor Then
                picDown.BackColor = UserControl.BackColor
            End If
        Else
            If picDown.BackColor <> m_BackColorDown Then
                picDown.BackColor = m_BackColorDown
            End If
        End If
    Else
        picDisp.Visible = False
        picDisp.Width = 0
        picDisp.Height = 0
    End If
    w3 = lblDisp.Width
    h3 = lblDisp.Height
    If UserControl.ScaleMode = 3 Then
        ctlW = UserControl.ScaleWidth * Screen.TwipsPerPixelX
        ctlH = UserControl.ScaleHeight * Screen.TwipsPerPixelY
    Else
        ctlW = UserControl.ScaleWidth
        ctlH = UserControl.ScaleHeight
    End If
    If m_PicDefaultArrange = 0 Then
        x = (ctlW - w1 - w2 - w3) / 2
        If x < 0 Then
            x = 0
        End If
        If ImageUpLoaded = True Then
            m_PicPosLeft = x
            picDisp.Left = m_PicPosLeft
            m_CapPosLeft = x + w1 + w2
            lblDisp.Left = m_CapPosLeft
        Else
            m_CapPosLeft = x
            lblDisp.Left = m_CapPosLeft
        End If
        If h1 > h3 Then
            y = (ctlH - h1) / 2
            If y < 0 Then
                y = 0
            End If
            m_PicPosTop = y
            picDisp.Top = m_PicPosTop
            m_CapPosTop = y + (h1 - h3) / 2
            lblDisp.Top = m_CapPosTop
        Else
            y = (ctlH - h3) / 2
            If y < 0 Then
                y = 0
            End If
            If ImageUpLoaded = True Then
                m_PicPosTop = y + (h3 - h1) / 2
                picDisp.Top = m_PicPosTop
            End If
            m_CapPosTop = y
            lblDisp.Top = m_CapPosTop
        End If
    Else
        If w1 > w3 Then
            x = (ctlW - w1) / 2
            If x < 0 Then
                x = 0
            End If
            m_PicPosLeft = x
            picDisp.Left = m_PicPosLeft
            m_CapPosLeft = x + (w1 - w3) / 2
            lblDisp.Left = m_CapPosLeft
        Else
            x = (ctlW - w3) / 2
            If x < 0 Then
                x = 0
            End If
            If ImageUpLoaded = True Then
                m_PicPosLeft = x + (w3 - w1) / 2
                picDisp.Left = m_PicPosLeft
            End If
            m_CapPosLeft = x
            lblDisp.Left = m_CapPosLeft
        End If
        y = (ctlH - h1 - h2 - h3) / 2
        If y < 0 Then
            y = 0
        End If
        If ImageUpLoaded = True Then
            m_PicPosTop = y
            picDisp.Top = m_PicPosTop
            m_CapPosTop = y + h1 + h2
            lblDisp.Top = m_CapPosTop
        Else
            m_CapPosTop = y
            lblDisp.Top = m_CapPosTop
        End If
    End If
    If picDisp.Left <> oldPicPosLeft Then
        PropertyChanged "PicPosLeft"
    End If
    If picDisp.Top <> oldPicPosTop Then
        PropertyChanged "PicPosTop"
    End If
    If lblDisp.Left <> oldCapPosLeft Then
        PropertyChanged "CapPosLeft"
    End If
    If lblDisp.Top <> oldCapPosTop Then
        PropertyChanged "CapPosTop"
    End If
    If UserControl.Enabled = False Then
        lblDisp.Enabled = False
        GreyPic
    Else
        lblDisp.Enabled = True
        Set picDisp.Picture = picDisp.Picture
    End If
End Sub
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled = New_Enabled
    If UserControl.Enabled = False Then
        lblDisp.Enabled = False
        GreyPic
    Else
        lblDisp.Enabled = True
        Set picDisp.Picture = picDisp.Picture
    End If
    PropertyChanged "Enabled"
End Property
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    On Error GoTo errHandler
    UserControl.BackColor = New_BackColor
    m_BackColor = New_BackColor
    picDisp.BackColor = m_BackColor
    PropertyChanged "BackColor"
    CenterBtnContents
errHandler:
End Property
Public Property Get BackColorOver() As OLE_COLOR
    BackColorOver = m_BackColorOver
End Property
Public Property Let BackColorOver(ByVal New_BackColorOver As OLE_COLOR)
    On Error GoTo errHandler
    m_BackColorOver = New_BackColorOver
    PropertyChanged "BackColorOver"
errHandler:
End Property
Public Property Get BackColorDown() As OLE_COLOR
    BackColorDown = m_BackColorDown
End Property
Public Property Let BackColorDown(ByVal New_BackColorDown As OLE_COLOR)
    On Error GoTo errHandler
    m_BackColorDown = New_BackColorDown
    PropertyChanged "BackColorDown"
errHandler:
End Property
Public Property Get BackStyle() As BtnBackStyle
    BackStyle = UserControl.BackStyle
End Property
Public Property Let BackStyle(ByVal New_BtnBackStyle As BtnBackStyle)
    UserControl.BackStyle = New_BtnBackStyle
    PropertyChanged "BtnBackStyle"
End Property
Public Property Get BtnAppearance() As BtnAppearance
    BtnAppearance = m_BtnAppearance
End Property
Public Property Let BtnAppearance(ByVal New_Appearance As BtnAppearance)
    On Error Resume Next
    If New_Appearance < 0 Or New_Appearance > 2 Then
        GoTo errHandler
    End If
    m_BtnAppearance = New_Appearance
    BtnAppearanceSet = New_Appearance
    PropertyChanged "BtnAppearance"
    PaintUserControl
    CenterBtnContents
    Exit Property
errHandler:
    Err.Raise Number:=383, Source:="Button.BtnAppearance", Description:="Index invalid"
End Property
Public Property Get BtnSunkenEffects() As BtnSunkenEffects
    BtnSunkenEffects = m_BtnSunkenEffects
End Property
Public Property Let BtnSunkenEffects(ByVal New_BtnFunction As BtnSunkenEffects)
    On Error Resume Next
    If New_BtnFunction < 0 Or New_BtnFunction > 1 Then
        GoTo errHandler
    End If
    m_BtnSunkenEffects = New_BtnFunction
    PropertyChanged "BtnSunkenEffects"
    PaintUserControl
    CenterBtnContents
    Exit Property
errHandler:
    Err.Raise Number:=383, Source:="Button.BtnSunkenEffects", Description:="Index invalid"
End Property
Public Property Get BtnContentEffects() As BtnContentEffects
    BtnContentEffects = m_BtnContentEffects
End Property
Public Property Let BtnContentEffects(ByVal New_Effect As BtnContentEffects)
    On Error Resume Next
    If New_Effect < 0 Or New_Effect > 1 Then
        GoTo errHandler
    End If
    m_BtnContentEffects = New_Effect
    PropertyChanged "BtnContentEffects"
    PaintUserControl
    CenterBtnContents
    Exit Property
errHandler:
    Err.Raise Number:=383, Source:="Button.BtnContentEffects", Description:="Index invalid"
End Property
Public Property Get Width() As Integer
    Width = UserControl.Width
End Property
Public Property Let Width(ByVal New_Width As Integer)
    UserControl.Width = New_Width
    PropertyChanged "BtnWidth"
End Property
Public Property Get Height() As Integer
    Height = UserControl.Height
End Property
Public Property Let Height(ByVal New_Height As Integer)
    UserControl.Height = New_Height
    PropertyChanged "BtnHeight"
End Property
Public Property Get Name() As String
    Name = Extender.Name
End Property
Public Property Let Name(ByVal New_Name As String)
    On Error Resume Next
    If Ambient.UserMode Then
        Err.Raise Number:=382, Source:="Button.Name", Description:="Index invalid"
        Exit Property
    End If
    Extender.Name = New_Name
    CenterBtnContents
    On Error GoTo 0
    Exit Property
errHandler:
End Property
Public Property Get Caption() As String
    Caption = lblDisp.Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
    lblDisp.Caption = New_Caption
    PropertyChanged "Caption"
    CenterBtnContents
End Property
Public Property Get CaptionDown() As String
    CaptionDown = m_CaptionDown
End Property
Public Property Let CaptionDown(ByVal New_Cap As String)
    m_CaptionDown = New_Cap
    PropertyChanged "CaptionDown"
End Property
Public Property Get CaptionFont() As Font
    Set CaptionFont = lblDisp.Font
End Property
Public Property Set CaptionFont(ByVal New_Font As Font)
    Set lblDisp.Font = New_Font
    PropertyChanged "CaptionFont"
    CenterBtnContents
End Property
Public Property Get CapColor() As OLE_COLOR
    CapColor = lblUp.ForeColor
End Property
Public Property Let CapColor(ByVal New_Color As OLE_COLOR)
    On Error Resume Next
    lblUp.ForeColor = New_Color
    lblDisp.ForeColor = New_Color
    PropertyChanged "CapColor"
End Property
Public Property Get CapColorOver() As OLE_COLOR
    CapColorOver = lblOver.ForeColor
End Property
Public Property Let CapColorOver(ByVal New_Color As OLE_COLOR)
    On Error Resume Next
    lblOver.ForeColor = New_Color
    PropertyChanged "CapColorOver"
End Property
Public Property Get CapColorDown() As OLE_COLOR
    CapColorDown = lblDown.ForeColor
End Property
Public Property Let CapColorDown(ByVal New_Color As OLE_COLOR)
    On Error Resume Next
    lblDown.ForeColor = New_Color
    PropertyChanged "CapColorDown"
End Property
Public Property Get picAutoFit() As picAutoFit
    picAutoFit = m_picAutoFit
End Property
Public Property Let picAutoFit(ByVal New_PicAutoFit As picAutoFit)
    On Error Resume Next
    If New_PicAutoFit < 0 Or New_PicAutoFit > 1 Then
        GoTo errHandler
    End If
    m_picAutoFit = New_PicAutoFit
    If Extender.picAutoFit = 0 Then
        If picUp.Picture > 0 Then
            picDisp.Left = 2 * Screen.TwipsPerPixelX
            picDisp.Top = 2 * Screen.TwipsPerPixelY
            UserControl.Width = picDisp.Width + 4 * Screen.TwipsPerPixelX
            UserControl.Height = picDisp.Height + 4 * Screen.TwipsPerPixelY
        End If
    End If
    PropertyChanged "picAutoFit"
    PaintUserControl
    CenterBtnContents
    Exit Property
errHandler:
    Err.Raise Number:=383, Source:="Button.picAutoFit", Description:="Index invalid"
End Property
Public Property Get picDefaultArrange() As picDefaultArrange
    picDefaultArrange = m_PicDefaultArrange
End Property
Public Property Let picDefaultArrange(ByVal New_DefaultArrange As picDefaultArrange)
    On Error Resume Next
    If New_DefaultArrange < 0 Or New_DefaultArrange > 1 Then
        GoTo errHandler
    End If
    m_PicDefaultArrange = New_DefaultArrange
    PropertyChanged "PicDefaultArrange"
    PaintUserControl
    CenterBtnContents
    Exit Property
errHandler:
    Err.Raise Number:=383, Source:="Button.PicDefaultArrange", Description:="Index invalid"
End Property
Public Property Get PicCapSpacing() As Long
    PicCapSpacing = m_PicCapSpacing
End Property
Public Property Let PicCapSpacing(ByVal vNewValue As Long)
    On Error Resume Next
    If vNewValue < 0 Then
        vNewValue = 0
    End If
    Dim i
    m_PicCapSpacing = vNewValue
    If UserControl.ScaleMode = 3 Then
        i = UserControl.ScaleWidth * Screen.TwipsPerPixelX
    Else
        i = UserControl.ScaleWidth
    End If
    If m_PicCapSpacing > (i - EdgePerimeter) Then
        m_PicCapSpacing = (i - EdgePerimeter)
    End If
    PropertyChanged "PicCapSpacing"
    CenterBtnContents
End Property
Public Property Get ToolTipBackColor() As OLE_COLOR
    ToolTipBackColor = m_ToolTipBackColor
End Property
Public Property Let ToolTipBackColor(ByVal New_Color As OLE_COLOR)
    On Error Resume Next
    m_ToolTipBackColor = New_Color
    PropertyChanged "ToolTipBackColor"
End Property
Public Property Get ToolTipForeColor() As OLE_COLOR
    ToolTipForeColor = m_ToolTipForeColor
End Property
Public Property Let ToolTipForeColor(ByVal New_Color As OLE_COLOR)
    On Error Resume Next
    m_ToolTipForeColor = New_Color
    PropertyChanged "ToolTipForeColor"
End Property
Public Property Get ToolTipText() As String
    ToolTipText = Extender.ToolTipText
End Property
Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    Extender.ToolTipText = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property
Public Property Get Picture() As Picture
    Set Picture = picUp.Picture
End Property
Public Property Set Picture(ByVal New_Picture As Picture)
    On Error Resume Next
    Set picUp.Picture = LoadPicture()
    Set picDisp.Picture = LoadPicture()
    Set picUp.Picture = New_Picture
    If picUp.Picture > 0 Then
        Set picDisp.Picture = picUp.Picture
        If Extender.picAutoFit = 0 Then
            picDisp.Left = 2 * Screen.TwipsPerPixelX
            picDisp.Top = 2 * Screen.TwipsPerPixelY
            UserControl.Width = picDisp.Width + 4 * Screen.TwipsPerPixelX
            UserControl.Height = picDisp.Height + 4 * Screen.TwipsPerPixelY
        End If
    End If
    If picDisp.Picture = 0 Then
        picDisp.Visible = False
        ImageUpLoaded = False
        ImageUpDownLoaded = False
    Else
        picDisp.Visible = True
        ImageUpLoaded = True
    End If
    PropertyChanged "PictureUp"
    PropertyChanged "PictureDisp"
    CenterBtnContents
End Property
Public Property Get PictureOver() As Picture
    Set PictureOver = picOver.Picture
End Property
Public Property Set PictureOver(ByVal New_Picture As Picture)
    On Error Resume Next
    Set picOver.Picture = New_Picture
    PropertyChanged "PictureOver"
End Property
Public Property Get PictureDown() As Picture
    Set PictureDown = picDown.Picture
End Property
Public Property Set PictureDown(ByVal New_Picture As Picture)
    On Error Resume Next
    Set picDown.Picture = New_Picture
    If picDown.Picture = 0 Then
        ImageUpDownLoaded = False
    Else
        If picUp.Picture > 0 Then
            ImageUpDownLoaded = True
        End If
    End If
    PropertyChanged "PictureDown"
End Property
Private Sub doMouseOverEffect(OnOff As Boolean)
    If m_BtnContentEffects = 1 Then
        Exit Sub
    End If
    If OnOff = False Then
        If m_BtnAppearance = BtnAppearanceSet Then
            If ImageUpLoaded = True Then
                picDisp.Picture = LoadPicture()
                picDisp.Picture = picUp.Image
            End If
            lblDisp.ForeColor = lblUp.ForeColor
        Else
            If BtnAppearanceSet = 0 Then
                If IsSunken = False Then
                    m_BtnAppearance = 0
                    PaintUserControl
                    If ImageUpLoaded = True Then
                        picDisp.Picture = LoadPicture()
                        picDisp.Picture = picUp.Image
                    End If
                    lblDisp.ForeColor = lblUp.ForeColor
                Else
                    If ImageUpDownLoaded Then
                        picDisp.Picture = LoadPicture()
                        picDisp.Picture = picDown.Image
                    End If
                    lblDisp.ForeColor = lblDown.ForeColor
                End If
            Else
                If ImageUpDownLoaded Then
                    picDisp.Picture = LoadPicture()
                    picDisp.Picture = picDown.Image
                End If
                lblDisp.ForeColor = lblDown.ForeColor
            End If
        End If
        If IsSunken = False Then
            If UserControl.BackColor <> m_BackColor Then
                UserControl.BackColor = m_BackColor
                picDisp.BackColor = UserControl.BackColor
            End If
        End If
    Else
        If m_BtnAppearance = BtnAppearanceSet Then
            If BtnAppearanceSet = 0 Then
                If IsSunken = False Then
                    m_BtnAppearance = 1
                    PaintUserControl
                End If
            End If
            If picOver.Picture > 0 Then
                picDisp.Picture = LoadPicture()
                picDisp.Picture = picOver.Image
            End If
            lblDisp.ForeColor = lblOver.ForeColor
            If UserControl.BackColor <> m_BackColorOver Then
                m_BackColor = UserControl.BackColor
                UserControl.BackColor = m_BackColorOver
                picDisp.BackColor = UserControl.BackColor
            End If
        End If
    End If
End Sub
Private Sub toggleClickFunction()
    IsSunken = Not IsSunken
    CenterBtnContents
    If IsSunken Then
        doSunkenAppearance
    Else
        doOrigAppearance
    End If
End Sub
Private Sub doSunkenAppearance()
    m_BtnAppearance = 3
    PaintUserControl
    IsSunken = True
    If m_BtnContentEffects = 0 Then
        If ImageUpDownLoaded Then
            picDisp.Picture = LoadPicture()
            picDisp.Picture = picDown.Picture
        End If
        lblDisp.ForeColor = lblDown.ForeColor
        If m_CaptionDown <> "" Then
            OrigCaption = lblDisp.Caption
            lblDisp.Caption = m_CaptionDown
        End If
        If UserControl.BackColor <> m_BackColorDown Then
            UserControl.BackColor = m_BackColorDown
            picDisp.BackColor = UserControl.BackColor
        End If
        CenterBtnContents
    End If
End Sub
Private Sub doOrigAppearance()
    m_BtnAppearance = BtnAppearanceSet
    PaintUserControl
    IsSunken = False
    If m_BtnContentEffects = 0 Then
        If ImageUpDownLoaded Then
            picDisp.Picture = LoadPicture()
            picDisp.Picture = picUp.Picture
        End If
        lblDisp.ForeColor = lblUp.ForeColor
        If m_CaptionDown <> "" Then
            lblDisp.Caption = OrigCaption
        End If
        If UserControl.BackColor <> m_BackColor Then
            UserControl.BackColor = m_BackColor
            picDisp.BackColor = UserControl.BackColor
        End If
        CenterBtnContents
    End If
End Sub
Sub GreyPic()
    Dim AveCol As Integer, a As Integer, Total As Long
    Dim x As Double
    Dim y As Double
    On Error Resume Next

    Total = (picDisp.Height * picDisp.Width)
    For y = 0 To picDisp.Height Step 15
        For x = 0 To picDisp.Width Step 15
            AveCol = 0
            a = 0
            RGBfromLONG (GetPixel(picDisp.hdc, x / 15, y / 15))
            AveCol = AveCol + rGreen: a = a + 1
            If AveCol <= 0 Then AveCol = 0
            AveCol = (AveCol / a)

            If (GetPixel(picDisp.hdc, x / 15, y / 15)) <> picDisp.BackColor Then

                SetPixel picDisp.hdc, x / 15, y / 15, RGB(AveCol, AveCol, AveCol)
            Else
                SetPixel picDisp.hdc, x / 15, y / 15, picDisp.BackColor
            End If

        Next x
        picDisp.Refresh
    Next y
    On Error GoTo 0
End Sub

Private Function RGBfromLONG(LongCol As Long)
    Dim Blue As Double, Green As Double, Red As Double
    On Error Resume Next
    Blue = Fix((LongCol / 256) / 256)
    Green = Fix((LongCol - ((Blue * 256) * 256)) / 256)
    Red = Fix(LongCol - ((Blue * 256) * 256) - (Green * 256))
    rRed = Red: rBlue = Blue: rGreen = Green
    On Error GoTo 0
End Function

Public Property Get SoundToPlay() As Sound
    SoundToPlay = FichierSon
End Property
Public Property Let SoundToPlay(ByVal New_SoundToPlay As Sound)
    FichierSon = New_SoundToPlay
    PropertyChanged "SoundToPlay"
End Property

Sub BeginPlaySound(ByVal ResourceId As Integer)
    SoundBuffer = LoadResData(ResourceId, "JF_Button_SOUND")
    sndPlaySound SoundBuffer(0), SND_SYNC Or SND_NODEFAULT Or SND_MEMORY
End Sub


