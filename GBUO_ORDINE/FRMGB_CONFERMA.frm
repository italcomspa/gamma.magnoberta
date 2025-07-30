VERSION 5.00
Object = "{B473387D-A75F-4A83-9879-4A8FE48EE80F}#1.6#0"; "TMS_TBARMENU.ocx"
Begin VB.Form FRMGB_CONFERMA 
   Caption         =   "Conferma"
   ClientHeight    =   1350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   2955
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framePassword 
      BackColor       =   &H00FFFFFF&
      Height          =   3585
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8505
      Begin PRJFW_TBARMENU.TMS_TBARMENU cmdFattura 
         Height          =   495
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   873
         Caption         =   "FATTURA"
         IsMenuPopup     =   0   'False
      End
      Begin PRJFW_TBARMENU.TMS_TBARMENU cmdDdt 
         Height          =   495
         Left            =   1470
         TabIndex        =   2
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   873
         Caption         =   "DDT"
         IsMenuPopup     =   0   'False
      End
      Begin PRJFW_TBARMENU.TMS_TBARMENU cmdAnnulla 
         Height          =   495
         Left            =   150
         TabIndex        =   3
         Top             =   780
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   873
         Caption         =   "Annulla"
         IsMenuPopup     =   0   'False
      End
   End
End
Attribute VB_Name = "FRMGB_CONFERMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AnnullaConferma As Boolean
Public TpDoc As String
Public AccontodaForm As Double




Private Sub cmdAnnulla_ButtonClick()
AnnullaConferma = True
Unload Me
End Sub

Private Sub cmdDdt_ButtonClick()
        
        
        TpDoc = "DDT"
        AnnullaConferma = False
        Unload Me
End Sub

Private Sub cmdFattura_ButtonClick()
        
        TpDoc = "FAT"
        Unload Me
        AnnullaConferma = False
End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    TpDoc = ""
    Set MyForm = New FRMGB_BANCO
    
    
    If CDbl(AccontodaForm) > 0 Then

        cmdFattura.Enabled = False


    End If
End Sub
