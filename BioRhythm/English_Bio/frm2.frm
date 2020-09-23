VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "BioRhythm"
   ClientHeight    =   4245
   ClientLeft      =   2550
   ClientTop       =   2340
   ClientWidth     =   7905
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frm2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   527
   Begin BioRythmes.Button Cimprimer 
      Height          =   585
      Left            =   7200
      TabIndex        =   18
      ToolTipText     =   "To print current biorhythm"
      Top             =   105
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1032
      BtnWidth        =   630
      BtnHeight       =   585
      BtnAppearance   =   0
      BtnSunkenEffects=   1
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      CapColor        =   0
      PictureDisp     =   "frm2.frx":A282
      PictureUp       =   "frm2.frx":A514
      PictureOver     =   "frm2.frx":A7A6
      PicPosLeft      =   52
      PicPosTop       =   52
      CapPosLeft      =   532
      CapPosTop       =   194
      ToolTipBackColor=   -2147483624
      ToolTipForeColor=   -2147483630
      SoundToPlay     =   1
   End
   Begin VB.HScrollBar Défilaa1 
      Height          =   225
      Left            =   1680
      Max             =   3000
      Min             =   1900
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3720
      Value           =   1900
      Width           =   435
   End
   Begin VB.HScrollBar Défilmm1 
      Height          =   225
      LargeChange     =   3
      Left            =   435
      Max             =   12
      Min             =   1
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3735
      Value           =   1
      Width           =   435
   End
   Begin VB.HScrollBar Défiljj1 
      Height          =   225
      LargeChange     =   3
      Left            =   1185
      Max             =   31
      Min             =   1
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3720
      Value           =   1
      Width           =   435
   End
   Begin VB.HScrollBar Défiljj2 
      Height          =   225
      LargeChange     =   3
      Left            =   6120
      Max             =   31
      Min             =   1
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3705
      Value           =   1
      Width           =   435
   End
   Begin VB.HScrollBar Défilmm2 
      Height          =   225
      LargeChange     =   3
      Left            =   5355
      Max             =   13
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3705
      Value           =   1
      Width           =   435
   End
   Begin VB.HScrollBar Défilaa2 
      Height          =   225
      Left            =   6615
      Max             =   3000
      Min             =   1900
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3705
      Value           =   1900
      Width           =   435
   End
   Begin VB.PictureBox Img3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Enabled         =   0   'False
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   4050
      Picture         =   "frm2.frx":AA38
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   3
      Top             =   3285
      Width           =   855
   End
   Begin VB.PictureBox Img2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Enabled         =   0   'False
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3195
      Picture         =   "frm2.frx":B0BE
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   2
      Top             =   3285
      Width           =   855
   End
   Begin VB.PictureBox Img1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Enabled         =   0   'False
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      Picture         =   "frm2.frx":B744
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   1
      Top             =   3285
      Width           =   855
   End
   Begin VB.PictureBox Img0 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      DrawStyle       =   2  'Dot
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3045
      Left            =   120
      ScaleHeight     =   199
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   462
      TabIndex        =   0
      Top             =   75
      Width           =   6990
   End
   Begin BioRythmes.Button Cmdquit 
      Height          =   585
      Left            =   7200
      TabIndex        =   19
      ToolTipText     =   "To leave application"
      Top             =   750
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1032
      BtnWidth        =   630
      BtnHeight       =   585
      BtnAppearance   =   0
      BtnSunkenEffects=   1
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      CapColor        =   0
      PictureDisp     =   "frm2.frx":BDCA
      PictureUp       =   "frm2.frx":C05C
      PictureOver     =   "frm2.frx":C2EE
      PicPosLeft      =   52
      PicPosTop       =   52
      CapPosLeft      =   532
      CapPosTop       =   194
      ToolTipBackColor=   -2147483624
      ToolTipForeColor=   -2147483630
      SoundToPlay     =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Date of the Biorhythm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   285
      Left            =   5070
      TabIndex        =   17
      Top             =   3165
      Width           =   1965
   End
   Begin VB.Label Label1 
      Caption         =   "Date of birth"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   165
      TabIndex        =   16
      Top             =   3180
      Width           =   2040
   End
   Begin VB.Label Etiaa1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1960"
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
      Left            =   1620
      TabIndex        =   15
      Top             =   3465
      Width           =   555
   End
   Begin VB.Label Etimm1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Septembre"
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
      Left            =   90
      TabIndex        =   14
      Top             =   3465
      Width           =   1095
   End
   Begin VB.Label Etijj1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01"
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
      Left            =   1185
      TabIndex        =   13
      Top             =   3465
      Width           =   435
   End
   Begin VB.Label Etijj2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01"
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
      Left            =   6120
      TabIndex        =   9
      Top             =   3450
      Width           =   435
   End
   Begin VB.Label Etimm2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Septembre"
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
      Left            =   5025
      TabIndex        =   8
      Top             =   3450
      Width           =   1095
   End
   Begin VB.Label Etiaa2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1960"
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
      Left            =   6555
      TabIndex        =   7
      Top             =   3450
      Width           =   555
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'=============================
'
'Nom Du Projet: BioRythmes
'
'Auteur:Les Productions J.F.
'
'=============================

Dim bb As Integer
Dim m001C As Integer
Dim m001E As Integer
Dim m0020 As Integer
Dim m0022 As Integer
Dim m0024 As Integer
Dim m0026 As Integer
Dim m002A As Integer
Dim m002C As Integer
Dim m002E As Integer
Dim Changement As Boolean
Dim m0036 As Single
Dim m003A As Single
Dim m003E As Single
Dim m0042 As Single
Dim m0046 As Single
Dim m004A As Single
Dim m004E As Single
Dim m0066(31) As Integer
Dim m007C(31) As Integer
Dim m0092(31) As Integer
Dim m00A8(31) As Integer
Dim m00BE(31) As Integer
Dim m00D4(31) As Integer
Dim m0102 As Integer

Sub CalculEtDessine()
    Dim l013A As Integer
    Dim l0140 As String * 2
    Dim l0142 As Integer
    Dim l0144 As Integer
    Img1.Cls
    Img2.Cls
    Img3.Cls
    Img0.DrawMode = 13
    Img1.DrawMode = 13
    Img2.DrawMode = 13
    Img3.DrawMode = 13
    m0026 = 85
    m0024 = 14
    VerifieNombreJourVSmaxDansLeMois
    No_DesMois(2) = 28
    If (Défilaa2.Value Mod 4) = 0 Then
        If Défilmm2.Value = 2 Then
            No_DesMois(2) = 29
        End If
    End If
    m0036 = CSng(DateSerial(Défilaa2.Value, Défilmm2.Value, Défiljj2.Value) - DateSerial(Défilaa1.Value, Défilmm1.Value, Défiljj1.Value))
    m0036 = m0036 - Défiljj2.Value - 1
    m001E = (1 * m0024) + 7
    m0020 = m001E + (No_DesMois(Défilmm2.Value) * m0024)
    Img0.Cls
    Img0.FontName = "MS Sans Serif"
    Img0.DrawStyle = 0
    Img0.FontSize = 14
    Img0.CurrentY = 0
    Img0.CurrentX = 190
    Img0.Print Trim(NomDesMois(Défilmm2.Value)); "  "; Format(Défilaa2.Value, "0000");
    Img0.FontName = "Small Fonts"
    Img0.FontBold = 0
    Img0.FontSize = 6.75
    Img0.CurrentX = 149
    Img0.CurrentY = 85
    Img0.Line (0, 100)-(471, 100), RGB(128, 128, 128)
    Img0.Line (m001E - 1, 0)-(m001E - 1, 199), RGB(128, 128, 128)
    Img0.Line (m0020 + 1, 0)-(m0020 + 1, 199), RGB(128, 128, 128)
    Img0.FontBold = False
    For l013A = 1 To No_DesMois(Défilmm2.Value)
        l0142 = m001E + ((l013A - 1) * 14)
        l0140 = NomDesJours(WeekDay(DateSerial(Défilaa2.Value, Défilmm2.Value, l013A)))
        Img0.CurrentX = l0142 + 4
        If l0140 = "M " Or l0140 = "M " Then Img0.CurrentX = l0142 + 1
        Img0.CurrentY = 85
        Img0.Print l0140
        Img0.CurrentX = l0142 + 5
        If l013A > 9 Then Img0.CurrentX = l0142 + 2
        Img0.CurrentY = 105
        Img0.Print Format(l013A, "0")
        If l0140 = "S " Then
            Img0.Line (l0142, 100)-(l0142 + 14, 100), RGB(0, 0, 128)
        Else
            If l0140 = "M " Or l0140 = "W " Or l0140 = "V " Then
                Img0.Line (l0142, 100)-(l0142 + 14, 100), RGB(128, 0, 0)
            End If
        End If
    Next
    Img0.DrawWidth = 1
    l013A = 0
    m0046 = (m0036 / 23) - Int(m0036 / 23)
    m004A = (m0036 / 28) - Int(m0036 / 28)
    m004E = (m0036 / 33) - Int(m0036 / 33)
    m0022 = 0
    m002A = 100 - (m0026 * Sin(6.2832 * m0046))
    m002C = 100 - (m0026 * Sin(6.2832 * m004A))
    m002E = 100 - (m0026 * Sin(6.2832 * m004E))
    l0144 = 1
    For m001C = 0 To 496 Step 7
        m003A = 100 - (m0026 * Sin(6.2832 * (m0046 + m001C / (m0024 * 23))))
        m003E = 100 - (m0026 * Sin(6.2832 * (m004A + m001C / (m0024 * 28))))
        m0042 = 100 - (m0026 * Sin(6.2832 * (m004E + m001C / (m0024 * 33))))
        Img0.Line (m0022, m002A)-(m001C, m003A), RGB(255, 0, 255)
        Img0.Line (m0022, m002C)-(m001C, m003E), RGB(0, 255, 255)
        Img0.Line (m0022, m002E)-(m001C, m0042), RGB(0, 255, 0)
        m0022 = m001C
        m002A = m003A
        m002C = m003E
        m002E = m0042
        If m001C > m001E And m001C < m0020 Then
            l0144 = l0144 + 1
            If l0144 = 2 Then
                l0144 = 0
                l013A = l013A + 1
                m0066(l013A) = 2 + (100 - (m0026 * Cos(6.2832 * (m0046 + m001C / (m0024 * 23))))) * 21 / m0026
                m0092(l013A) = 2 + (100 - (m0026 * Cos(6.2832 * (m004A + m001C / (m0024 * 28))))) * 21 / m0026
                m00BE(l013A) = 2 + (100 - (m0026 * Cos(6.2832 * (m004E + m001C / (m0024 * 33))))) * 21 / m0026
                m007C(l013A) = 2 + (m003A * 21 / m0026)
                m00A8(l013A) = 2 + (m003E * 21 / m0026)
                m00D4(l013A) = 2 + (m0042 * 21 / m0026)
            End If
        End If
    Next
    Changement = True
    Img1.Line (27, 6)-(27, 25), RGB(0, 0, 0)
    Img2.Line (27, 6)-(27, 25), RGB(0, 0, 0)
    Img3.Line (27, 6)-(27, 25), RGB(0, 0, 0)
    Img1.DrawMode = 7
    Img2.DrawMode = 7
    Img3.DrawMode = 7
    Img0.DrawMode = 7
    Img0.DrawStyle = 2
    bb = m001E + (Défiljj2.Value * m0024) - 7
    Img0.Line (bb, 0)-(bb, 199), RGB(255, 255, 255)
    Img1.Line (m0066(Défiljj2.Value), m007C(Défiljj2.Value))-(27, 27), RGB(255, 0, 255)
    Img2.Line (m0092(Défiljj2.Value), m00A8(Défiljj2.Value))-(27, 27), RGB(0, 255, 255)
    Img3.Line (m00BE(Défiljj2.Value), m00D4(Défiljj2.Value))-(27, 27), RGB(0, 255, 0)
End Sub
Sub DessineChangementJour()
    Img0.Line (bb, 0)-(bb, 199), RGB(255, 255, 255)
    Img1.Line (m0066(Défiljj2.Value), m007C(Défiljj2.Value))-(27, 27), RGB(255, 0, 255)
    Img2.Line (m0092(Défiljj2.Value), m00A8(Défiljj2.Value))-(27, 27), RGB(0, 255, 255)
    Img3.Line (m00BE(Défiljj2.Value), m00D4(Défiljj2.Value))-(27, 27), RGB(0, 255, 0)
    bb = m001E + (Défiljj2.Value * m0024) - 7
    Img0.Line (bb, 0)-(bb, 199), RGB(255, 255, 255)
    Img1.Line (m0066(Défiljj2.Value), m007C(Défiljj2.Value))-(27, 27), RGB(255, 0, 255)
    Img2.Line (m0092(Défiljj2.Value), m00A8(Défiljj2.Value))-(27, 27), RGB(0, 255, 255)
    Img3.Line (m00BE(Défiljj2.Value), m00D4(Défiljj2.Value))-(27, 27), RGB(0, 255, 0)
End Sub

Private Sub Cimprimer_Click()
    CalculEtDessine
    frmChoix.Show 1
End Sub

Private Sub Cmdquit_Click()
    Unload frmMain
End Sub

Sub Défilaa1_Change()
    Etiaa1.Caption = Str$(Défilaa1.Value)
    Rafraichir
    CalculEtDessine
End Sub

Sub Défilaa2_Change()
    Etiaa2.Caption = Str$(Défilaa2.Value)
    If Défilmm2.Value = 2 And Défiljj2.Value = 29 Then
        Défiljj2.Value = 28
        DessineChangementJour
    End If
    CalculEtDessine
End Sub

Sub Défiljj1_Change()
    Etijj1.Caption = Str$(Défiljj1.Value)
    Rafraichir
    CalculEtDessine
End Sub

Sub Défiljj2_Change()
    Etijj2.Caption = Str$(Défiljj2.Value)
    VerifieNombreJourVSmaxDansLeMois
    DessineChangementJour
    CalculEtDessine
End Sub

Sub Défilmm1_Change()
    Etimm1.Caption = NomDesMois(Défilmm1.Value)
    Rafraichir
    CalculEtDessine
End Sub

Sub Défilmm2_Change()
    If Défilmm2.Value = 13 Then
        If Défilaa2.Value < Défilaa2.Max Then
            Défilmm2.Value = 1
            Défilaa2.Value = Défilaa2.Value + 1
            Etiaa2.Caption = Str$(Défilaa2.Value)
        End If
    End If
    If Défilmm2.Value = 0 Then
        If Défilaa2.Value > Défilaa2.Min Then
            Défilmm2.Value = 12
            Défilaa2.Value = Défilaa2.Value - 1
            Etiaa2.Caption = Str$(Défilaa2.Value)
        End If
    End If
    Etimm2.Caption = NomDesMois(Défilmm2.Value)
    If Défiljj2.Value > No_DesMois(Défilmm2.Value) Then
        If (Défilaa2.Value Mod 4) = 0 And Défilmm2.Value = 2 Then
            Défiljj2.Value = 29
        Else
            Défiljj2.Value = No_DesMois(Défilmm2.Value)
        End If
        Etijj2.Caption = Str$(Défiljj2.Value)
        DessineChangementJour
    End If
    CalculEtDessine
End Sub

Sub Form_Load()

    Changement = False
    No_DesMois(0) = 31
    No_DesMois(1) = 31
    No_DesMois(2) = 28
    No_DesMois(3) = 31
    No_DesMois(4) = 30
    No_DesMois(5) = 31
    No_DesMois(6) = 30
    No_DesMois(7) = 31
    No_DesMois(8) = 31
    No_DesMois(9) = 30
    No_DesMois(10) = 31
    No_DesMois(11) = 30
    No_DesMois(12) = 31
    No_DesMois(13) = 31
    NomDesMois(0) = "  December"
    NomDesMois(1) = "  January"
    NomDesMois(2) = "  February"
    NomDesMois(3) = "  March"
    NomDesMois(4) = "  April"
    NomDesMois(5) = "  May"
    NomDesMois(6) = "  June"
    NomDesMois(7) = "  July"
    NomDesMois(8) = "  August"
    NomDesMois(9) = "  September"
    NomDesMois(10) = "  October"
    NomDesMois(11) = "  November"
    NomDesMois(12) = "  December"
    NomDesMois(13) = "  January"
    NomDesJours(1) = "S "
    NomDesJours(2) = "M "
    NomDesJours(3) = "Tu"
    NomDesJours(4) = "W "
    NomDesJours(5) = "T "
    NomDesJours(6) = "F "
    NomDesJours(7) = "Sa"
    m0102 = 0
    Défiljj1.Value = "02"
    Défilmm1.Value = "07"
    Défilaa1.Value = "1961"
    Etijj1.Caption = Str$(Défiljj1.Value)
    Etiaa1.Caption = Str$(Défilaa1.Value)
    Etimm1.Caption = NomDesMois(Défilmm1.Value)
    Défiljj2.Value = Day(Now)
    Défilmm2.Value = Month(Now)
    Défilaa2.Value = Year(Now)
    Etijj2.Caption = Str$(Défiljj2.Value)
    Etiaa2.Caption = Str$(Défilaa2.Value)
    Etimm2.Caption = NomDesMois(Défilmm2.Value)
    frmMain.Show
End Sub

Private Sub Form_Resize()
    If frmMain.WindowState = 0 Then
    End If
End Sub

Sub VerifieNombreJourVSmaxDansLeMois()
    If Défiljj1.Value > No_DesMois(Défilmm1.Value) Then
        If (Défilaa1.Value Mod 4) = 0 And Défilmm1.Value = 2 Then
            Défiljj1.Value = 29
        Else
            Défiljj1.Value = No_DesMois(Défilmm1.Value)
        End If
        Etijj1.Caption = Str$(Défiljj1.Value)
    End If
    If Défiljj2.Value > No_DesMois(Défilmm2.Value) Then
        If (Défilaa2.Value Mod 4) = 0 And Défilmm2.Value = 2 Then
            Défiljj2.Value = 29
        Else
            Défiljj2.Value = No_DesMois(Défilmm2.Value)
        End If
        Etijj2.Caption = Str$(Défiljj2.Value)
    End If
End Sub

Sub Rafraichir()
    If Changement = True Then
        Changement = False
        Img0.Cls
        Img1.Cls
        Img2.Cls
        Img3.Cls
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

