VERSION 5.00
Begin VB.Form frmTooltip 
   Appearance      =   0  'Flat
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   315
   ClientLeft      =   2955
   ClientTop       =   3195
   ClientWidth     =   1635
   ControlBox      =   0   'False
   Icon            =   "ToolTip.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   315
   ScaleWidth      =   1635
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblToolTipText 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   675
   End
   Begin VB.Label lblDefToolTipColor 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   525
   End
End
Attribute VB_Name = "frmTooltip"
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

