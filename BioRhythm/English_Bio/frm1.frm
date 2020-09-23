VERSION 5.00
Begin VB.Form frmChoix 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   990
   ClientLeft      =   5115
   ClientTop       =   4350
   ClientWidth     =   2535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   Icon            =   "frm1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   990
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   Begin BioRythmes.Button Cimpcoul 
      Height          =   570
      Left            =   75
      TabIndex        =   0
      ToolTipText     =   "To print in colour"
      Top             =   375
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1005
      BtnWidth        =   735
      BtnHeight       =   570
      BackColor       =   12632256
      BackColorOver   =   12632256
      BackColorDown   =   12632256
      BtnAppearance   =   0
      BtnSunkenEffects=   1
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      CapColor        =   0
      PictureDisp     =   "frm1.frx":000C
      PictureUp       =   "frm1.frx":0C5E
      PicPosLeft      =   90
      PicPosTop       =   45
      CapPosLeft      =   570
      CapPosTop       =   188
      ToolTipBackColor=   12648447
      ToolTipForeColor=   0
      SoundToPlay     =   1
   End
   Begin BioRythmes.Button Cimpnoir 
      Height          =   570
      Left            =   900
      TabIndex        =   1
      ToolTipText     =   "To print in black"
      Top             =   375
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1005
      BtnWidth        =   735
      BtnHeight       =   570
      BackColor       =   12632256
      BackColorOver   =   12632256
      BackColorDown   =   12632256
      BtnAppearance   =   0
      BtnSunkenEffects=   1
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      CapColor        =   0
      PictureDisp     =   "frm1.frx":18B0
      PictureUp       =   "frm1.frx":2502
      PicPosLeft      =   90
      PicPosTop       =   45
      CapPosLeft      =   570
      CapPosTop       =   188
      ToolTipBackColor=   12648447
      ToolTipForeColor=   0
      SoundToPlay     =   1
   End
   Begin BioRythmes.Button Cannuler 
      Height          =   570
      Left            =   1725
      TabIndex        =   2
      ToolTipText     =   "To cancel printing"
      Top             =   375
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1005
      BtnWidth        =   735
      BtnHeight       =   570
      BackColor       =   12632256
      BackColorOver   =   12632256
      BackColorDown   =   12632256
      BtnAppearance   =   0
      BtnSunkenEffects=   1
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      CapColor        =   0
      PictureDisp     =   "frm1.frx":3154
      PictureUp       =   "frm1.frx":3DA6
      PicPosLeft      =   90
      PicPosTop       =   45
      CapPosLeft      =   570
      CapPosTop       =   188
      ToolTipBackColor=   12648447
      ToolTipForeColor=   0
      SoundToPlay     =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Type of printing"
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   540
      TabIndex        =   3
      Top             =   60
      Width           =   1650
   End
End
Attribute VB_Name = "frmChoix"
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
Dim m0020 As Integer
Dim m0022 As Integer
Dim m0024 As Integer
Dim m0026 As Integer
Dim m0028 As Single
Dim m002C As Single
Dim m0030 As Single
Dim m0034 As Single
Dim m0038 As Single
Dim m003C As Single
Dim m0040 As Single
Dim JourNaissance As Integer
Dim MoisNaissance As Integer
Dim AnneeNaissance As Integer
Dim JourDeLaBio As Integer
Dim MoisDeLaBio As Integer
Dim AnneeDeLaBio As Integer
Dim TypeEncre As Integer
Const c008E = 450 ' &H1C2%
Const c0090 = 60 ' &H3C%
Const c0092 = 45 ' &H2D%
Private Sub Cannuler_Click()
  frmChoix.Hide
End Sub

Private Sub Cimpcoul_Click()
  TypeEncre = 1
  Imprime
  frmChoix.Hide
End Sub

Private Sub Cimpnoir_Click()
  TypeEncre = 0
  Imprime
  frmChoix.Hide
End Sub

Sub Imprime()
Dim CoordOrigineX As Integer
Dim CoordFinX As Integer
Dim CoordOrigineY As Integer
Dim CoordFinY As Integer
Dim aa As Integer
Dim LettreDuJour As String * 2
Dim EpaisseurDesLignes As Integer
Dim StyleDeLignePhy As Integer
Dim StyleDeLigneEmo As Integer
Dim StyleDeLigneInt As Integer
Dim CouleurDeLignePhy As Long
Dim CouleurDeLigneEmo As Long
Dim CouleurDeLigneInt As Long

On Error Resume Next
  frmChoix.MousePointer = 11
  JourNaissance = frmMain.Défiljj1.Value
  MoisNaissance = frmMain.Défilmm1.Value
  AnneeNaissance = frmMain.Défilaa1.Value
  JourDeLaBio = frmMain.Défiljj2.Value
  MoisDeLaBio = frmMain.Défilmm2.Value
  AnneeDeLaBio = frmMain.Défilaa2.Value
  No_DesMois(2) = 28
  If (AnneeDeLaBio Mod 4) = 0 Then
     If MoisDeLaBio = 2 Then 'Février
        No_DesMois(2) = 29
     End If
  End If
  m0028 = CSng(DateSerial(AnneeDeLaBio, MoisDeLaBio, JourDeLaBio) - DateSerial(AnneeNaissance, MoisNaissance, JourNaissance))
  m0028 = m0028 - JourDeLaBio
  Printer.DrawMode = 13 'Copy Pen
  Printer.ScaleMode = 3 'Pixel (plus petite unité de résolution du moniteur ou de l'imprimante).
  Printer.FontName = "Courier New"
  If TypeEncre = 1 Then
     EpaisseurDesLignes = 3
     StyleDeLignePhy = 0
     StyleDeLigneEmo = 0
     StyleDeLigneInt = 0
     CouleurDeLignePhy = RGB(255, 0, 0)
     CouleurDeLigneEmo = RGB(0, 0, 255)
     CouleurDeLigneInt = RGB(0, 255, 0)
  Else
     EpaisseurDesLignes = 1
     StyleDeLignePhy = 0
     StyleDeLigneEmo = 1
     StyleDeLigneInt = 2
     CouleurDeLignePhy = RGB(0, 0, 0)
     CouleurDeLigneEmo = RGB(0, 0, 0)
     CouleurDeLigneInt = RGB(0, 0, 0)
  End If
  
  Printer.FontSize = 24
  Printer.FontBold = True
  Printer.CurrentX = 910
  Printer.CurrentY = 74
  Printer.Print "Biorythms";
  Printer.FontSize = 12
  Printer.FontBold = True
  Printer.CurrentX = 160
  Printer.CurrentY = 280
  Printer.Print "Date of birth";
  Printer.CurrentX = 1990
  Printer.CurrentY = 280
  Printer.Print Format(AnneeNaissance, "0000"); "/"; Format(MoisNaissance, "00"); "/"; Format(JourNaissance, "00");
  Printer.FontBold = True
  Printer.FontSize = 16
  Printer.CurrentX = 960
  Printer.CurrentY = 410
  Printer.Print Trim(NomDesMois(MoisDeLaBio)); " "; Format(AnneeDeLaBio, "0000");
  Printer.FontBold = True
  Printer.FontSize = 10
  Printer.DrawWidth = EpaisseurDesLignes
  Printer.DrawStyle = StyleDeLignePhy
  Printer.CurrentX = 600
  Printer.CurrentY = 1510
  Printer.Print "Physical";
  Printer.Line (490, 1536)-(590, 1536), CouleurDeLignePhy
  Printer.DrawStyle = StyleDeLigneEmo
  Printer.CurrentX = 1110
  Printer.CurrentY = 1510
  Printer.Print "Emotional";
  Printer.Line (1000, 1536)-(1100, 1536), CouleurDeLigneEmo
  Printer.DrawStyle = StyleDeLigneInt
  Printer.CurrentX = 1660
  Printer.CurrentY = 1510
  Printer.Print "Intellectual";
  Printer.Line (1550, 1536)-(1650, 1536), CouleurDeLigneInt
    'Cadre ext. bas
  Printer.DrawStyle = 0
  Printer.DrawWidth = 4
  CoordOrigineX = 160
  CoordFinX = 2280
  CoordOrigineY = 395
  CoordFinY = 1575
  Printer.Line (CoordOrigineX, CoordOrigineY)-(CoordFinX, CoordOrigineY), RGB(0, 0, 0)
  Printer.Line (CoordOrigineX, CoordFinY)-(CoordFinX, CoordFinY), RGB(0, 0, 0)
  Printer.Line (CoordOrigineX, CoordOrigineY)-(CoordOrigineX, CoordFinY), RGB(0, 0, 0)
  Printer.Line (CoordFinX, CoordOrigineY)-(CoordFinX, CoordFinY), RGB(0, 0, 0)
    'Cadre int. bas
  CoordOrigineX = 260
  CoordFinX = 2180
  CoordOrigineY = 495
  CoordFinY = 1485
  Printer.Line (CoordOrigineX, CoordOrigineY)-(CoordFinX, CoordOrigineY), RGB(0, 0, 0)
  Printer.Line (CoordOrigineX, CoordFinY)-(CoordFinX, CoordFinY), RGB(0, 0, 0)
  Printer.Line (CoordOrigineX, CoordOrigineY)-(CoordOrigineX, CoordFinY), RGB(0, 0, 0)
  Printer.Line (CoordFinX, CoordOrigineY)-(CoordFinX, CoordFinY), RGB(0, 0, 0)
  Printer.Line (CoordOrigineX, CoordOrigineY + ((CoordFinY - CoordOrigineY) / 2))-(CoordFinX, CoordOrigineY + ((CoordFinY - CoordOrigineY) / 2)), RGB(0, 0, 0)
  
  CoordOrigineY = (22 * c0092)
  For aa = 2 To 32
    CoordOrigineX = 200 + (aa * 60)
    Printer.Line (CoordOrigineX, CoordOrigineY - 10)-(CoordOrigineX, CoordOrigineY + 10), RGB(0, 0, 0)
  Next
  Printer.DrawWidth = 1
  CoordOrigineY = 495
  CoordFinY = 1485
  For aa = 1 To 33
    CoordOrigineX = 200 + (aa * 60)
    Printer.Line (CoordOrigineX, CoordOrigineY)-(CoordOrigineX, CoordFinY), RGB(0, 0, 0)
  Next
  CoordOrigineX = 248
  CoordFinX = 2180
  For aa = 2 To 22
    CoordOrigineY = 450 + (aa * 45)
    Printer.Line (CoordOrigineX, CoordOrigineY)-(CoordFinX, CoordOrigineY), RGB(0, 0, 0)
  Next
  Printer.FontSize = 10
  Printer.FontBold = False
  CoordOrigineX = 208
  CoordOrigineY = 990 - (45 / 2)
  For aa = 0 To 10
    CoordFinY = CoordOrigineY + (aa * 45)
    CoordFinX = CoordOrigineX - 23
    If aa = 10 Then CoordFinX = CoordFinX - 23
    Printer.CurrentX = CoordFinX
    Printer.CurrentY = CoordFinY
    If aa > 0 Then
       Printer.Print Format(-aa, "0")
    End If
    CoordFinY = CoordOrigineY - (aa * 45)
    CoordFinX = CoordOrigineX
    If aa = 10 Then CoordFinX = CoordFinX - 23
    Printer.CurrentX = CoordFinX
    Printer.CurrentY = CoordFinY
    Printer.Print Format(aa, "0")
  Next
  Printer.FontSize = 9
  Printer.FontBold = True
  CoordOrigineY = 1071 - (45 / 2)
  For aa = 1 To No_DesMois(MoisDeLaBio)
    CoordOrigineX = 200 + ((aa + 1) * 60)
    LettreDuJour = NomDesJours(WeekDay(DateSerial(AnneeDeLaBio, MoisDeLaBio, aa)))
    Printer.CurrentX = CoordOrigineX - 11
    If LettreDuJour = "Tu" Or LettreDuJour = "Sa" Then Printer.CurrentX = CoordOrigineX - 22
    Printer.CurrentY = CoordOrigineY - 45
    Printer.Print LettreDuJour
    Printer.CurrentX = CoordOrigineX - 11
    If aa > 9 Then Printer.CurrentX = CoordOrigineX - 22
    Printer.CurrentY = CoordOrigineY
    Printer.Print Format(aa, "0")
  Next
  Printer.DrawWidth = EpaisseurDesLignes
  m0038 = (m0028 / 23) - Int(m0028 / 23)
  m003C = (m0028 / 28) - Int(m0028 / 28)
  m0040 = (m0028 / 33) - Int(m0028 / 33)
  m0020 = 260
  For bb = 260 To 2180 Step 10
    m002C = 990 - (c008E * Sin(6.2832 * (m0038 + (bb - 260) / (c0090 * 23))))
    m0030 = 990 - (c008E * Sin(6.2832 * (m003C + (bb - 260) / (c0090 * 28))))
    m0034 = 990 - (c008E * Sin(6.2832 * (m0040 + (bb - 260) / (c0090 * 33))))
    If bb > 260 Then
       Printer.DrawStyle = StyleDeLignePhy
       Printer.Line (m0020, m0022)-(bb, m002C), CouleurDeLignePhy
       Printer.DrawStyle = StyleDeLigneEmo
       Printer.Line (m0020, m0024)-(bb, m0030), CouleurDeLigneEmo
       Printer.DrawStyle = StyleDeLigneInt
       Printer.Line (m0020, m0026)-(bb, m0034), CouleurDeLigneInt
    End If
    m0020 = bb
    m0022 = m002C
    m0024 = m0030
    m0026 = m0034
  Next
  JourDeLaBio = frmMain.Défiljj2.Value
  MoisDeLaBio = frmMain.Défilmm2.Value + 1
  AnneeDeLaBio = frmMain.Défilaa2.Value
  If MoisDeLaBio = 13 Then
    AnneeDeLaBio = AnneeDeLaBio + 1
    MoisDeLaBio = 1
  End If
  No_DesMois(2) = 28
  If (AnneeDeLaBio Mod 4) = 0 Then
     If MoisDeLaBio = 2 Then
        No_DesMois(2) = 29
     End If
  End If
  m0028 = CSng(DateSerial(AnneeDeLaBio, MoisDeLaBio, JourDeLaBio) - DateSerial(AnneeNaissance, MoisNaissance, JourNaissance))
  m0028 = m0028 - JourDeLaBio
  Printer.FontBold = True
  Printer.FontSize = 16
  Printer.CurrentX = 960
  Printer.CurrentY = 1910
  Printer.Print Trim(NomDesMois(MoisDeLaBio)); " "; Format(AnneeDeLaBio, "0000");
  Printer.DrawWidth = EpaisseurDesLignes
  Printer.FontBold = True
  Printer.FontSize = 10
  Printer.DrawStyle = StyleDeLignePhy
  Printer.CurrentX = 600
  Printer.CurrentY = 3010
  Printer.Print "Physical";
  Printer.Line (490, 3036)-(590, 3036), CouleurDeLignePhy
  Printer.DrawStyle = StyleDeLigneEmo
  Printer.CurrentX = 1110
  Printer.CurrentY = 3010
  Printer.Print "Emotional";
  Printer.Line (1000, 3036)-(1100, 3036), CouleurDeLigneEmo
  Printer.DrawStyle = StyleDeLigneInt
  Printer.CurrentX = 1660
  Printer.CurrentY = 3010
  Printer.Print "Intellectual";
  Printer.Line (1550, 3036)-(1650, 3036), CouleurDeLigneInt
  'Cadre ext. bas
  Printer.DrawStyle = 0
  Printer.DrawWidth = 4
  CoordOrigineX = 160
  CoordFinX = 2280
  CoordOrigineY = 1895
  CoordFinY = 3075 '3085
  Printer.Line (CoordOrigineX, CoordOrigineY)-(CoordFinX, CoordOrigineY), RGB(0, 0, 0)
  Printer.Line (CoordOrigineX, CoordFinY)-(CoordFinX, CoordFinY), RGB(0, 0, 0)
  Printer.Line (CoordOrigineX, CoordOrigineY)-(CoordOrigineX, CoordFinY), RGB(0, 0, 0)
  Printer.Line (CoordFinX, CoordOrigineY)-(CoordFinX, CoordFinY), RGB(0, 0, 0)
  'Cadre int. bas
  CoordOrigineX = 260
  CoordFinX = 2180
  CoordOrigineY = 1995
  CoordFinY = 2985
  Printer.Line (CoordOrigineX, CoordOrigineY)-(CoordFinX, CoordOrigineY), RGB(0, 0, 0)
  Printer.Line (CoordOrigineX, CoordFinY)-(CoordFinX, CoordFinY), RGB(0, 0, 0)
  Printer.Line (CoordOrigineX, CoordOrigineY)-(CoordOrigineX, CoordFinY), RGB(0, 0, 0)
  Printer.Line (CoordFinX, CoordOrigineY)-(CoordFinX, CoordFinY), RGB(0, 0, 0)
  Printer.Line (CoordOrigineX, CoordOrigineY + ((CoordFinY - CoordOrigineY) / 2))-(CoordFinX, CoordOrigineY + ((CoordFinY - CoordOrigineY) / 2)), RGB(0, 0, 0)
  CoordOrigineY = (22 * c0092) + 1500
  For aa = 2 To 32
    CoordOrigineX = 200 + (aa * 60)
    Printer.Line (CoordOrigineX, CoordOrigineY - 10)-(CoordOrigineX, CoordOrigineY + 10), RGB(0, 0, 0)
  Next
  Printer.DrawWidth = 1
  CoordOrigineY = 1995
  CoordFinY = 2985
  For aa = 1 To 33
    CoordOrigineX = 200 + (aa * 60)
    Printer.Line (CoordOrigineX, CoordOrigineY)-(CoordOrigineX, CoordFinY), RGB(0, 0, 0)
  Next
  CoordOrigineX = 248
  CoordFinX = 2180
  For aa = 2 To 22
    CoordOrigineY = 1950 + (aa * 45)
    Printer.Line (CoordOrigineX, CoordOrigineY)-(CoordFinX, CoordOrigineY), RGB(0, 0, 0)
  Next
  Printer.FontSize = 10
  Printer.FontBold = False
  CoordOrigineX = 208
  CoordOrigineY = 2490 - (45 / 2)
  For aa = 0 To 10
    CoordFinY = CoordOrigineY + (aa * 45)
    CoordFinX = CoordOrigineX - 23
    If aa = 10 Then CoordFinX = CoordFinX - 23
    Printer.CurrentX = CoordFinX
    Printer.CurrentY = CoordFinY
    If aa > 0 Then
       Printer.Print Format(-aa, "0")
    End If
    CoordFinY = CoordOrigineY - (aa * 45)
    CoordFinX = CoordOrigineX
    If aa = 10 Then CoordFinX = CoordFinX - 23
    Printer.CurrentX = CoordFinX
    Printer.CurrentY = CoordFinY
    Printer.Print Format(aa, "0")
  Next
  Printer.FontSize = 9
  Printer.FontBold = True
  CoordOrigineY = 2571 - (45 / 2)
  For aa = 1 To No_DesMois(MoisDeLaBio)
    CoordOrigineX = 200 + ((aa + 1) * 60)
    LettreDuJour = NomDesJours(WeekDay(DateSerial(AnneeDeLaBio, MoisDeLaBio, aa)))
    Printer.CurrentX = CoordOrigineX - 11
    If LettreDuJour = "Tu" Or LettreDuJour = "Sa" Then Printer.CurrentX = CoordOrigineX - 22
    Printer.CurrentY = CoordOrigineY - 45
    Printer.Print LettreDuJour
    Printer.CurrentX = CoordOrigineX - 11
    If aa > 9 Then Printer.CurrentX = CoordOrigineX - 22
    Printer.CurrentY = CoordOrigineY
    Printer.Print Format(aa, "0")
  Next
  Printer.DrawWidth = EpaisseurDesLignes
  m0038 = (m0028 / 23) - Int(m0028 / 23)
  m003C = (m0028 / 28) - Int(m0028 / 28)
  m0040 = (m0028 / 33) - Int(m0028 / 33)
  m0020 = 260
  For bb = 260 To 2180 Step 10
    m002C = 2490 - (c008E * Sin(6.2832 * (m0038 + (bb - 260) / (c0090 * 23))))
    m0030 = 2490 - (c008E * Sin(6.2832 * (m003C + (bb - 260) / (c0090 * 28))))
    m0034 = 2490 - (c008E * Sin(6.2832 * (m0040 + (bb - 260) / (c0090 * 33))))
    If bb > 260 Then
       Printer.DrawStyle = StyleDeLignePhy
       Printer.Line (m0020, m0022)-(bb, m002C), CouleurDeLignePhy
       Printer.DrawStyle = StyleDeLigneEmo
       Printer.Line (m0020, m0024)-(bb, m0030), CouleurDeLigneEmo
       Printer.DrawStyle = StyleDeLigneInt
       Printer.Line (m0020, m0026)-(bb, m0034), CouleurDeLigneInt
    End If
    m0020 = bb
    m0022 = m002C
    m0024 = m0030
    m0026 = m0034
  Next
  Printer.EndDoc
  frmChoix.MousePointer = 0
End Sub

