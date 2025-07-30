VERSION 5.00
Object = "{F53BE214-7AC6-11D0-9B0E-006097A80EFD}#6.7#0"; "TMS_LABEL.ocx"
Object = "{B473387D-A75F-4A83-9879-4A8FE48EE80F}#1.6#0"; "TMS_TBARMENU.ocx"
Begin VB.Form FRMGB_PASSWORD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framePassword 
      BackColor       =   &H00FFFFFF&
      Height          =   3585
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   8505
      Begin VB.TextBox TXT_PASSWORD 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   4800
         PasswordChar    =   "*"
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1320
         Width           =   3285
      End
      Begin PRJFW_TBARMENU.TMS_TBARMENU cmdConferma 
         Height          =   495
         Left            =   2820
         TabIndex        =   1
         Top             =   2250
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   873
         Caption         =   "Conferma"
         IsMenuPopup     =   0   'False
      End
      Begin PRJFW_TBARMENU.TMS_TBARMENU cmdAnnulla 
         Height          =   495
         Left            =   4650
         TabIndex        =   2
         Top             =   2250
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   873
         Caption         =   "Annulla"
         IsMenuPopup     =   0   'False
      End
      Begin VB.Label lblTipoBlocco 
         Caption         =   "Attenzione, cliente bloccato per FIDO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   150
         TabIndex        =   4
         Top             =   420
         Width           =   7935
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL2 
         Height          =   300
         Left            =   150
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1350
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   529
         Caption         =   "Inserire password per sbloccare l'inserimento"
      End
   End
End
Attribute VB_Name = "FRMGB_PASSWORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Password           As String
Public PasswordCorretta   As Boolean
Public FormAperto             As Boolean

Private Sub cmdAnnulla_ButtonClick()
  FormAperto = False
  Unload Me
End Sub

Private Sub cmdConferma_ButtonClick()
  
  If Password <> TXT_PASSWORD.Text Then
    MsgBox "Password errata"
    PasswordCorretta = False
  Else
    PasswordCorretta = True
    FormAperto = False
    Unload Me
  End If
  
End Sub

Private Sub Form_Load()
  TXT_PASSWORD.Text = ""
  FormAperto = True
  PasswordCorretta = False
End Sub
