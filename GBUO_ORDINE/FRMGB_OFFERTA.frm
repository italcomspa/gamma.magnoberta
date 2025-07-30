VERSION 5.00
Object = "{5032AB27-52C8-11D2-A1C0-0060082875F9}#4.24#0"; "TMS_EDITM.ocx"
Object = "{0EF4EA3A-2617-11D2-A1C0-0060082875F9}#8.12#0"; "TMS_EDIT.ocx"
Object = "{F53BE214-7AC6-11D0-9B0E-006097A80EFD}#6.12#0"; "TMS_LABEL.ocx"
Object = "{5CC9FF70-1720-11D2-A1C0-0060082875F9}#3.6#0"; "TMS_GRIDNAV.ocx"
Object = "{C99C525C-61F9-11D2-AE21-00A0244C5B50}#3.16#0"; "TMS_EDITNUMM.ocx"
Object = "{0EF4E9DB-2617-11D2-A1C0-0060082875F9}#10.13#0"; "TMS_EDITNUM.ocx"
Object = "{F2DC983F-61F7-11D2-AE21-00A0244C5B50}#3.15#0"; "TMS_EDITDATEM.ocx"
Object = "{0EF4EAA6-2617-11D2-A1C0-0060082875F9}#4.14#0"; "TMS_COMBOBOX.ocx"
Object = "{C217CF55-DAD6-4868-A146-622ECD75BC60}#1.63#0"; "TMS_QGRID.ocx"
Object = "{840F600B-FE39-42F4-AE87-798701D999E2}#1.25#0"; "TMS_RESIZEFORM.ocx"
Object = "{B473387D-A75F-4A83-9879-4A8FE48EE80F}#1.8#0"; "TMS_TBARMENU.ocx"
Object = "{0EF4EAD5-2617-11D2-A1C0-0060082875F9}#5.13#0"; "TMS_CHECKBOX.ocx"
Object = "{9AE03505-25F7-11D2-A1C0-0060082875F9}#7.3#0"; "TMS_FRAME.ocx"
Object = "{EF28CC5E-FCE3-448A-AB46-AEA7C5A209AA}#1.4#0"; "TMS_SSTAB.ocx"
Object = "{53EEE555-1204-4E18-B5DB-A659E06A9EEB}#1.3#0"; "TMS_FLATBUTTON.ocx"
Object = "{1B331849-A95B-4C35-9982-01A4869D197E}#1.10#0"; "TMS_EDITPATH.ocx"
Object = "{0EF4E915-2617-11D2-A1C0-0060082875F9}#7.22#0"; "TMS_RICHTEXTBOX.ocx"
Object = "{0EF4EA13-2617-11D2-A1C0-0060082875F9}#7.13#0"; "TMS_EDITDATE.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FRMGB_OFFERTA 
   Caption         =   "Gestione Offerta"
   ClientHeight    =   10425
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   17865
   Icon            =   "FRMGB_OFFERTA.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10990.12
   ScaleMode       =   0  'User
   ScaleWidth      =   18105
   StartUpPosition =   3  'Windows Default
   Begin PRJ_SSTAB.TMS_SSTAB TabDocumenti 
      Height          =   13440
      Left            =   0
      TabIndex        =   2
      Top             =   -30
      Width           =   21885
      _ExtentX        =   38603
      _ExtentY        =   23707
      TabCount        =   2
      TabCaption(0)   =   "Filtri"
      TabContCtrlCnt(0)=   5
      Tab(0)ContCtrlCap(1)=   "cmdScarico"
      Tab(0)ContCtrlCap(2)=   "cmdModifica"
      Tab(0)ContCtrlCap(3)=   "TXT_ANNOINSERIMENTO"
      Tab(0)ContCtrlCap(4)=   "Label5"
      Tab(0)ContCtrlCap(5)=   "Label6"
      TabCaption(1)   =   "Offerta"
      TabContCtrlCnt(1)=   7
      Tab(1)ContCtrlCap(1)=   "FrameRiferimenti"
      Tab(1)ContCtrlCap(2)=   "TMS_FRAME20"
      Tab(1)ContCtrlCap(3)=   "frmOpticom"
      Tab(1)ContCtrlCap(4)=   "FrameGriglia"
      Tab(1)ContCtrlCap(5)=   "FrameCliente"
      Tab(1)ContCtrlCap(6)=   "Label4"
      Tab(1)ContCtrlCap(7)=   "lblTipoDocumento"
      ActiveTab       =   1
      PictureMaskColor=   4210816
      Begin PRJFW_TBARMENU.TMS_TBARMENU cmdScarico 
         Height          =   2550
         Left            =   -74970
         TabIndex        =   195
         Top             =   360
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   4498
         Caption         =   "&Nuova Offerta Cliente"
         IsMenuPopup     =   0   'False
      End
      Begin PRJFW_TBARMENU.TMS_TBARMENU cmdModifica 
         Height          =   2550
         Left            =   -72180
         TabIndex        =   194
         Top             =   360
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   4498
         Caption         =   "&Modifica Offerta Cliente"
         IsMenuPopup     =   0   'False
      End
      Begin PRJFW_FRAME.TMS_FRAME FrameRiferimenti 
         Height          =   1155
         Left            =   30
         Top             =   2880
         Width           =   13845
         _ExtentX        =   24421
         _ExtentY        =   2037
         Begin PRJFW_EDITM.TXT_EDITM TXT_GB06_CODCOMM 
            Height          =   300
            Left            =   1080
            TabIndex        =   28
            ToolTipText     =   "Articolo"
            Top             =   30
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   529
            IsGestione      =   -1  'True
            IsLookup        =   -1  'True
            DisplayFormat   =   "Maiuscolo"
            MaxChar         =   25
            Obbligatorio    =   -1  'True
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            DBField         =   "GB07_CODART_MG66"
            IsDecode        =   -1  'True
            Caption         =   "Articolo"
            NumRighe        =   0
            Object.Tag             =   "Articolo"
            MaxWidth        =   13
            IsInLingua      =   0   'False
            LinguaEntitaDes =   0
            LinguaIDProvenienzaExt=   ""
         End
         Begin PRJFW_EDITM.TXT_EDITM TXT_GB06_TIPOOFFERTA 
            Height          =   300
            Left            =   6840
            TabIndex        =   29
            ToolTipText     =   "Articolo"
            Top             =   60
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   529
            IsLookup        =   -1  'True
            DisplayFormat   =   "Maiuscolo"
            MaxChar         =   6
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            DBField         =   "GB07_CODART_MG66"
            IsDecode        =   -1  'True
            Caption         =   "Articolo"
            NumRighe        =   0
            Object.Tag             =   "Articolo"
            MaxWidth        =   6
            IsInLingua      =   0   'False
            LinguaEntitaDes =   0
            LinguaIDProvenienzaExt=   ""
         End
         Begin PRJFW_EDITM.TXT_EDITM TXT_GB06_TIPOAREA 
            Height          =   300
            Left            =   6840
            TabIndex        =   30
            ToolTipText     =   "Articolo"
            Top             =   720
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   529
            IsLookup        =   -1  'True
            DisplayFormat   =   "Maiuscolo"
            MaxChar         =   6
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            DBField         =   "GB07_CODART_MG66"
            IsDecode        =   -1  'True
            Caption         =   "Articolo"
            NumRighe        =   0
            Object.Tag             =   "Articolo"
            MaxWidth        =   6
            IsInLingua      =   0   'False
            LinguaEntitaDes =   0
            LinguaIDProvenienzaExt=   ""
         End
         Begin PRJFW_EDITM.TXT_EDITM TXT_GB06_AGENTE 
            Height          =   300
            Left            =   6840
            TabIndex        =   31
            ToolTipText     =   "Articolo"
            Top             =   390
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   529
            IsLookup        =   -1  'True
            DisplayFormat   =   "Maiuscolo"
            MaxChar         =   6
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            DBField         =   "GB07_CODART_MG66"
            IsDecode        =   -1  'True
            Caption         =   "Articolo"
            NumRighe        =   0
            Object.Tag             =   "Articolo"
            MaxWidth        =   6
            IsInLingua      =   0   'False
            LinguaEntitaDes =   0
            LinguaIDProvenienzaExt=   ""
         End
         Begin PRJFW_EDITDATEM.EditDateM TXT_GB06_DTCHIUSURA 
            Height          =   300
            Left            =   1080
            TabIndex        =   32
            Top             =   360
            Width           =   1455
            _ExtentX        =   2514
            _ExtentY        =   529
            IsCalendario    =   0   'False
            IsDbField       =   0   'False
         End
         Begin PRJFW_EDITM.TXT_EDITM TXT_GB06_CIG 
            Height          =   300
            Left            =   3270
            TabIndex        =   33
            ToolTipText     =   "Articolo"
            Top             =   360
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   529
            IsLookup        =   -1  'True
            DisplayFormat   =   "Maiuscolo"
            Numerico        =   0   'False
            Carattere       =   0   'False
            DBField         =   "GB07_CIG"
            IsDecode        =   -1  'True
            Caption         =   "Articolo"
            NumRighe        =   0
            Object.Tag             =   "Articolo"
            IsInLingua      =   0   'False
            LinguaEntitaDes =   0
            LinguaIDProvenienzaExt=   ""
         End
         Begin PRJFW_EDITM.TXT_EDITM TXT_GB06_CUP 
            Height          =   300
            Left            =   3270
            TabIndex        =   34
            ToolTipText     =   "Articolo"
            Top             =   690
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   529
            IsLookup        =   -1  'True
            DisplayFormat   =   "Maiuscolo"
            MaxChar         =   15
            Numerico        =   0   'False
            Carattere       =   0   'False
            DBField         =   "GB07_CUP"
            IsDecode        =   -1  'True
            Caption         =   "Articolo"
            NumRighe        =   0
            Object.Tag             =   "Articolo"
            MaxWidth        =   15
            IsInLingua      =   0   'False
            LinguaEntitaDes =   0
            LinguaIDProvenienzaExt=   ""
         End
         Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON TMS_FLATBUTTON2 
            Height          =   495
            Left            =   12180
            TabIndex        =   222
            Top             =   30
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            Caption         =   "Allinea da Web"
            ButtonBackColor =   255
            ButtonBorderColor=   255
            ButtonDisabledColor=   255
            ButtonForeColor =   0
            ButtonHilightBorderColor=   9473677
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TXT_GB06_COMMESSA_DEC 
            Height          =   300
            Left            =   3420
            TabIndex        =   49
            TabStop         =   0   'False
            Tag             =   "Descr. fornitore ENECO"
            Top             =   30
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   529
            Caption         =   ""
            Border          =   -1  'True
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TXT_GB06_TIPOOFFERTA_DEC 
            Height          =   300
            Left            =   8550
            TabIndex        =   48
            TabStop         =   0   'False
            Tag             =   "Descr. fornitore ENECO"
            Top             =   60
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   529
            Caption         =   ""
            Border          =   -1  'True
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TXT_GB06_TIPOAREA_DEC 
            Height          =   300
            Left            =   8550
            TabIndex        =   47
            TabStop         =   0   'False
            Tag             =   "Descr. fornitore ENECO"
            Top             =   720
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   529
            Caption         =   ""
            Border          =   -1  'True
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TXT_GB06_AGENTE_DEC 
            Height          =   300
            Left            =   8550
            TabIndex        =   46
            TabStop         =   0   'False
            Tag             =   "Descr. fornitore ENECO"
            Top             =   390
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   529
            Caption         =   ""
            Border          =   -1  'True
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_GB06_PERCCHIUSURA 
            Height          =   300
            Left            =   1080
            TabIndex        =   45
            Top             =   690
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   529
            DBField         =   "GB07_SCCORPO"
            Caption         =   "SCONTO RIGA"
            Object.Tag             =   "SCONTO RIGA"
            MaxWidth        =   4
            MaxChar         =   6
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   35
            Left            =   90
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   60
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            Caption         =   "Commessa"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   36
            Left            =   90
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   390
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            Caption         =   "Data Chiusura"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   37
            Left            =   90
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   720
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            Caption         =   "% Chiusura"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   38
            Left            =   2910
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   390
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   529
            Caption         =   "CIG"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   39
            Left            =   2880
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   720
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   529
            Caption         =   "CUP"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   40
            Left            =   5880
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   90
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            Caption         =   "Tipo Offerta"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   41
            Left            =   5880
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   420
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            Caption         =   "Agente"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   42
            Left            =   5880
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   750
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            Caption         =   "Tipo Area"
         End
         Begin PRJFW_EDIT.TxtEdit TXT_GB06_RESPONSABILE 
            Height          =   300
            Left            =   11100
            TabIndex        =   36
            Top             =   720
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   529
            MaxChar         =   72
            Numerico        =   0   'False
            Carattere       =   0   'False
            DBField         =   "GB06_RESPONSABILE"
            MaxWidth        =   20
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   50
            Left            =   10920
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   780
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            Caption         =   "Responsabile"
         End
      End
      Begin PRJFW_FRAME.TMS_FRAME TMS_FRAME2 
         Height          =   3615
         Index           =   0
         Left            =   13890
         Top             =   420
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   6376
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   1530
            Picture         =   "FRMGB_OFFERTA.frx":27A2
            ScaleHeight     =   225
            ScaleWidth      =   255
            TabIndex        =   177
            Top             =   90
            Width           =   255
         End
         Begin ComctlLib.ImageList ImageList 
            Left            =   360
            Top             =   2730
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   327682
         End
         Begin ComctlLib.ImageList ImgLstFornitore 
            Left            =   300
            Top             =   1890
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   327682
         End
         Begin ComctlLib.ImageList ImageList2 
            Left            =   270
            Top             =   990
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   327682
         End
         Begin ComctlLib.ImageList ImageList1 
            Left            =   300
            Top             =   150
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   327682
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   52
            Left            =   1920
            TabIndex        =   217
            TabStop         =   0   'False
            Top             =   2850
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   529
            Caption         =   "08 - ANNULLATA"
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   225
            Left            =   1860
            TabIndex        =   207
            Top             =   3240
            Width           =   1995
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   16
            Left            =   1920
            TabIndex        =   180
            TabStop         =   0   'False
            Top             =   2640
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   529
            Caption         =   "07 - ARCHIVIATA"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   15
            Left            =   1920
            TabIndex        =   181
            TabStop         =   0   'False
            Top             =   2430
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   529
            Caption         =   "06 - PERSA"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   14
            Left            =   1920
            TabIndex        =   182
            TabStop         =   0   'False
            Top             =   2220
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   529
            Caption         =   "05 - ORDINE"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   13
            Left            =   1920
            TabIndex        =   183
            TabStop         =   0   'False
            Top             =   2010
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   529
            Caption         =   "04 - REVISIONATA"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   12
            Left            =   1920
            TabIndex        =   184
            TabStop         =   0   'False
            Top             =   1800
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   529
            Caption         =   "03 - ATTIVA"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   11
            Left            =   1920
            TabIndex        =   185
            TabStop         =   0   'False
            Top             =   1590
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   529
            Caption         =   "02 - RILASCIATA"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   9
            Left            =   1920
            TabIndex        =   186
            TabStop         =   0   'False
            Top             =   1380
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   529
            Caption         =   "01 - COMPLETATA"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   18
            Left            =   1920
            TabIndex        =   189
            TabStop         =   0   'False
            Top             =   1170
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   529
            Caption         =   "00 - IN LAVORAZIONE"
         End
         Begin PRJFW_EDIT.TxtEdit TXT_GB06_PROPRIETARIO 
            Height          =   300
            Left            =   1680
            TabIndex        =   188
            Top             =   60
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   529
            MaxChar         =   100
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            MaxWidth        =   16
         End
         Begin PRJFW_EDIT.TxtEdit TXT_GB06_STATODOC 
            Height          =   300
            Left            =   1680
            TabIndex        =   187
            Top             =   390
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            MaxChar         =   2
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            MaxWidth        =   2
         End
         Begin VB.Shape Shape2 
            Height          =   285
            Index           =   0
            Left            =   2190
            Top             =   390
            Width           =   1665
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   17
            Left            =   1920
            TabIndex        =   179
            TabStop         =   0   'False
            Top             =   750
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   529
            Caption         =   "Legenda Stati"
         End
         Begin VB.Shape Shape1 
            Height          =   2475
            Index           =   0
            Left            =   1860
            Top             =   720
            Width           =   1995
         End
         Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON TMS_CODIFICA_STATO 
            Height          =   225
            Left            =   2220
            TabIndex        =   178
            Top             =   420
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   397
            ButtonBackColor =   12582912
            ButtonBorderColor=   0
            ButtonForeColor =   16777215
            ButtonHilightBorderColor=   9473677
         End
      End
      Begin VB.Frame frmOpticom 
         Height          =   2535
         Left            =   10620
         TabIndex        =   161
         Top             =   330
         Width           =   3255
         Begin VB.TextBox TXT_GB06_BUDGET1 
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """€"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Left            =   1380
            TabIndex        =   164
            Top             =   150
            Width           =   225
         End
         Begin VB.TextBox TXT_GB06_CONSUNTIVO1 
            BackColor       =   &H0000FFFF&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """€"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   2
            EndProperty
            Height          =   255
            Left            =   1380
            TabIndex        =   163
            Top             =   540
            Width           =   225
         End
         Begin VB.TextBox TXT_GB06_FORECAST1 
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """€"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   2
            EndProperty
            Height          =   285
            Left            =   1380
            TabIndex        =   162
            Top             =   900
            Width           =   225
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_GB06_PERCRIBGARA 
            Height          =   300
            Left            =   1440
            TabIndex        =   176
            Top             =   1290
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
            IsDbField       =   0   'False
            DBField         =   "GB07_SCCORPO"
            Caption         =   "SCONTO RIGA"
            Object.Tag             =   "SCONTO RIGA"
            MaxWidth        =   8
            MaxChar         =   6
            TipoFormato     =   6
         End
         Begin PRJFW_EDITDATE.TxtEditDate TXT_GB06_DTULTMOD 
            Height          =   300
            Left            =   1470
            TabIndex        =   175
            Top             =   1680
            Width           =   1680
            _ExtentX        =   2566
            _ExtentY        =   529
            IsCalendario    =   0   'False
            Enabled         =   0   'False
            IsDbField       =   0   'False
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   19
            Left            =   120
            TabIndex        =   174
            TabStop         =   0   'False
            Top             =   180
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            Caption         =   "Budget"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   20
            Left            =   120
            TabIndex        =   173
            TabStop         =   0   'False
            Top             =   540
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            Caption         =   "Consuntivo"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   21
            Left            =   120
            TabIndex        =   172
            TabStop         =   0   'False
            Top             =   930
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            Caption         =   "Forecast"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   23
            Left            =   120
            TabIndex        =   171
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   529
            Caption         =   "% Ribasso Gara"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   24
            Left            =   120
            TabIndex        =   170
            TabStop         =   0   'False
            Top             =   1710
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   529
            Caption         =   "Ultima Variazione"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   61
            Left            =   120
            TabIndex        =   169
            TabStop         =   0   'False
            Top             =   2040
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   529
            Caption         =   "Perc. Trasporti"
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_GB06_PERCTRASP 
            Height          =   300
            Left            =   1470
            TabIndex        =   168
            Top             =   2040
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   529
            IsDbField       =   0   'False
            DBField         =   "GB07_QTA"
            Caption         =   "Perc.Trasporto"
            Object.Tag             =   "Perc.Trasporto"
            MaxWidth        =   5
            MaxChar         =   14
            TipoFormato     =   3
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_GB06_CONSUNTIVO 
            Height          =   300
            Left            =   1440
            TabIndex        =   167
            Top             =   510
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
            DBField         =   "GB06_CONSUNTIVO"
            Caption         =   "Consuntivo"
            Object.Tag             =   "Consuntivo"
            MaxWidth        =   8
            MaxChar         =   17
            TipoFormato     =   3
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_GB06_FORECAST 
            Height          =   300
            Left            =   1440
            TabIndex        =   166
            Top             =   900
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
            DBField         =   "GB06_FORECAST"
            Caption         =   "Forecast"
            Object.Tag             =   "Forecast"
            MaxWidth        =   8
            MaxChar         =   17
            TipoFormato     =   3
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_GB06_BUDGET 
            Height          =   300
            Left            =   1440
            TabIndex        =   165
            Top             =   150
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
            DBField         =   "GB06_BUDGET"
            Caption         =   "Prezzo"
            Object.Tag             =   "Budget"
            MaxWidth        =   8
            MaxChar         =   17
            TipoFormato     =   3
         End
      End
      Begin PRJFW_FRAME.TMS_FRAME FrameGriglia 
         Height          =   6405
         Left            =   30
         Top             =   4050
         Width           =   17775
         _ExtentX        =   31353
         _ExtentY        =   11298
         Begin PRJ_SSTAB.TMS_SSTAB TMS_SSTAB1 
            Height          =   6075
            Left            =   30
            TabIndex        =   50
            Top             =   30
            Width           =   17655
            _ExtentX        =   31141
            _ExtentY        =   10716
            TabCount        =   9
            TabCaption(0)   =   "Inserimento Articoli"
            TabContCtrlCnt(0)=   7
            Tab(0)ContCtrlCap(1)=   "FrameGenera"
            Tab(0)ContCtrlCap(2)=   "TMS_RESIZEFORM"
            Tab(0)ContCtrlCap(3)=   "GridNavDocumenti"
            Tab(0)ContCtrlCap(4)=   "FrameDoc"
            Tab(0)ContCtrlCap(5)=   "FrameCorpo"
            Tab(0)ContCtrlCap(6)=   "QGridDocumenti"
            Tab(0)ContCtrlCap(7)=   "CMD_DELALL"
            TabCaption(1)   =   "Simulazione Prezzi"
            TabContCtrlCnt(1)=   2
            Tab(1)ContCtrlCap(1)=   "TMS_FRAME3"
            Tab(1)ContCtrlCap(2)=   "QGRID_SIMULAZIONE"
            TabCaption(2)   =   "Analisi Margini"
            TabVisible(2)   =   0   'False
            TabCaption(3)   =   "Documento su Gamma"
            TabContCtrlCnt(3)=   1
            Tab(3)ContCtrlCap(1)=   "TMS_GRIDDOC"
            TabCaption(4)   =   "Note Offerta"
            TabContCtrlCnt(4)=   3
            Tab(4)ContCtrlCap(1)=   "TMS_GRIDNAV_NOTE"
            Tab(4)ContCtrlCap(2)=   "QGRID_NOTE"
            Tab(4)ContCtrlCap(3)=   "TMS_FRAME6"
            TabCaption(5)   =   "Immagine"
            TabContCtrlCnt(5)=   2
            Tab(5)ContCtrlCap(1)=   "PictureArticoli"
            Tab(5)ContCtrlCap(2)=   "TMS_FLATBUTTON5"
            TabCaption(6)   =   "Nomi Gruppi Articoli"
            TabContCtrlCnt(6)=   7
            Tab(6)ContCtrlCap(1)=   "TMS_QGRIDGRUPPI"
            Tab(6)ContCtrlCap(2)=   "TMS_GRIDNAV_GRUPPI"
            Tab(6)ContCtrlCap(3)=   "TMS_LABEL1632"
            Tab(6)ContCtrlCap(4)=   "TMS_LABEL1631"
            Tab(6)ContCtrlCap(5)=   "Label2"
            Tab(6)ContCtrlCap(6)=   "TXT_GB01_DESCRIZIONE"
            Tab(6)ContCtrlCap(7)=   "TXT_GB01_PROG"
            TabCaption(7)   =   "Testi Offerta"
            TabContCtrlCnt(7)=   12
            Tab(7)ContCtrlCap(1)=   "TXT_GB06_TEXT6"
            Tab(7)ContCtrlCap(2)=   "TXT_GB06_TEXT5"
            Tab(7)ContCtrlCap(3)=   "TXT_GB06_TEXT4"
            Tab(7)ContCtrlCap(4)=   "TXT_GB06_TEXT2"
            Tab(7)ContCtrlCap(5)=   "TXT_GB06_TEXT1"
            Tab(7)ContCtrlCap(6)=   "TXT_GB06_TEXT3"
            Tab(7)ContCtrlCap(7)=   "CMD_REFRESH_TXT4"
            Tab(7)ContCtrlCap(8)=   "CMD_REFRESH_TXT6"
            Tab(7)ContCtrlCap(9)=   "CMD_REFRESH_TXT2"
            Tab(7)ContCtrlCap(10)=   "CMD_REFRESH_TXT3"
            Tab(7)ContCtrlCap(11)=   "CMD_REFRESH_TXT5"
            Tab(7)ContCtrlCap(12)=   "CMD_REFRESH_TXT1"
            TabCaption(8)   =   "Classificazione Articolo"
            TabVisible(8)   =   0   'False
            Begin PRJFW_FRAME.TMS_FRAME FrameGenera 
               Height          =   585
               Left            =   330
               Top             =   1830
               Width           =   17295
               _ExtentX        =   30506
               _ExtentY        =   1032
               Begin PRJFW_TBARMENU.TMS_TBARMENU cmdNuovoDoc 
                  Height          =   495
                  Left            =   30
                  TabIndex        =   196
                  Top             =   30
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   873
                  Caption         =   "&Nuovo doc."
                  IsMenuPopup     =   0   'False
               End
               Begin PRJFW_TBARMENU.TMS_TBARMENU cmdAnnulla 
                  Height          =   495
                  Left            =   1140
                  TabIndex        =   197
                  Top             =   30
                  Width           =   1065
                  _ExtentX        =   1879
                  _ExtentY        =   873
                  Caption         =   "&Indietro"
                  IsMenuPopup     =   0   'False
               End
               Begin PRJFW_TBARMENU.TMS_TBARMENU cmdNewVersion 
                  Height          =   495
                  Left            =   2220
                  TabIndex        =   198
                  Top             =   30
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   873
                  Caption         =   "&Nuova Versione"
                  IsMenuPopup     =   0   'False
               End
               Begin PRJFW_TBARMENU.TMS_TBARMENU cmdNewRevision 
                  Height          =   495
                  Left            =   3600
                  TabIndex        =   199
                  Top             =   30
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   873
                  Caption         =   "&Nuova Revisione"
                  IsMenuPopup     =   0   'False
               End
               Begin PRJFW_TBARMENU.TMS_TBARMENU CMD_DUPLICA 
                  Height          =   495
                  Left            =   4980
                  TabIndex        =   200
                  Top             =   30
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   873
                  Caption         =   "&Duplica"
                  IsMenuPopup     =   0   'False
               End
               Begin PRJFW_TBARMENU.TMS_TBARMENU CMD_ELIMINA 
                  Height          =   495
                  Left            =   6090
                  TabIndex        =   201
                  Top             =   30
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   873
                  Caption         =   "&Elimina"
                  IsMenuPopup     =   0   'False
               End
               Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON CMD_ANNULLA 
                  Height          =   495
                  Left            =   7530
                  TabIndex        =   216
                  Top             =   30
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   873
                  Caption         =   "Annulla Offerta"
                  ButtonBackColor =   255
                  ButtonBorderColor=   255
                  ButtonDisabledColor=   255
                  ButtonForeColor =   0
                  ButtonHilightBorderColor=   9473677
               End
               Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON CMD_SAVE 
                  Height          =   495
                  Left            =   12360
                  TabIndex        =   206
                  Top             =   30
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   873
                  Caption         =   "Salva Offerta"
                  ButtonBackColor =   65280
                  ButtonBorderColor=   65280
                  ButtonForeColor =   0
                  ButtonHilightBorderColor=   9473677
               End
               Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON CMD_PRINT 
                  Height          =   495
                  Left            =   13980
                  TabIndex        =   205
                  Top             =   30
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   873
                  Caption         =   "Stampa Offerta"
                  ButtonBackColor =   16777215
                  ButtonBorderColor=   16777215
                  ButtonDisabledColor=   16777215
                  ButtonForeColor =   4210752
                  ButtonHilightBorderColor=   9473677
               End
               Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON CMD_GENERAOFFERTA 
                  Height          =   495
                  Left            =   15660
                  TabIndex        =   204
                  Top             =   30
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   873
                  Caption         =   "Genera Ordine"
                  ButtonBackColor =   16711680
                  ButtonBorderColor=   16711680
                  ButtonDisabledColor=   16711680
                  ButtonForeColor =   16777215
                  ButtonHilightBorderColor=   9473677
               End
               Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON CMD_ANALISIMARGINI 
                  Height          =   495
                  Left            =   9120
                  TabIndex        =   203
                  Top             =   30
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   873
                  Caption         =   "Analisi Margini"
                  ButtonBorderColor=   0
                  ButtonForeColor =   16777215
                  ButtonHilightBorderColor=   9473677
               End
               Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON CDM_RICPERCPROVVMASSI 
                  Height          =   495
                  Left            =   10620
                  TabIndex        =   202
                  Top             =   30
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   873
                  Caption         =   "Ric. Tutte % Provv. "
                  ButtonBorderColor=   0
                  ButtonForeColor =   16777215
                  ButtonHilightBorderColor=   9473677
               End
            End
            Begin PRJ_RESIZEFORM.TMS_RESIZEFORM TMS_RESIZEFORM 
               Left            =   750
               Top             =   2730
               _ExtentX        =   847
               _ExtentY        =   847
            End
            Begin PRJFW_GRIDNAV.TMS_GRIDNAV GridNavDocumenti 
               Height          =   1845
               Left            =   0
               TabIndex        =   192
               Top             =   1800
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   3254
               Primo           =   0   'False
               Precedente      =   0   'False
               Successivo      =   0   'False
               Ultimo          =   0   'False
               Apri            =   0   'False
               Orizzontale     =   0   'False
            End
            Begin PRJFW_FRAME.TMS_FRAME FrameDoc 
               Height          =   1245
               Left            =   5310
               Top             =   3960
               Visible         =   0   'False
               Width           =   7425
               _ExtentX        =   13097
               _ExtentY        =   2196
               Begin VB.PictureBox Picture2 
                  BackColor       =   &H00EAE8D0&
                  Height          =   1095
                  Left            =   30
                  ScaleHeight     =   1035
                  ScaleWidth      =   7155
                  TabIndex        =   190
                  Top             =   90
                  Width           =   7215
                  Begin VB.Label Label1 
                     Alignment       =   2  'Center
                     Caption         =   "Scrittura documento in corso"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   525
                     Left            =   900
                     TabIndex        =   191
                     Top             =   240
                     Width           =   5385
                  End
               End
            End
            Begin TMS_QGRID.TMS_QGRIDWRAPPER TMS_GRIDDOC 
               Height          =   3015
               Left            =   -74970
               TabIndex        =   148
               Top             =   360
               Width           =   17565
               _ExtentX        =   30983
               _ExtentY        =   5318
               GridLoadMode    =   0
               DefaultRowHeight=   18
               BeginProperty PreviewFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PreviewFontColor=   0
            End
            Begin PRJFW_GRIDNAV.TMS_GRIDNAV TMS_GRIDNAV_NOTE 
               Height          =   1815
               Left            =   -57840
               TabIndex        =   147
               Top             =   1470
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   3201
               Primo           =   0   'False
               Precedente      =   0   'False
               Successivo      =   0   'False
               Ultimo          =   0   'False
               Apri            =   0   'False
               Orizzontale     =   0   'False
            End
            Begin VB.TextBox TXT_GB06_TEXT6 
               Appearance      =   0  'Flat
               Height          =   705
               Left            =   -66720
               MultiLine       =   -1  'True
               TabIndex        =   146
               Top             =   1170
               Width           =   8985
            End
            Begin VB.TextBox TXT_GB06_TEXT5 
               Appearance      =   0  'Flat
               Height          =   705
               Left            =   -74910
               MultiLine       =   -1  'True
               TabIndex        =   145
               Top             =   1170
               Width           =   7935
            End
            Begin VB.TextBox TXT_GB06_TEXT4 
               Appearance      =   0  'Flat
               Height          =   2925
               Left            =   -66720
               MultiLine       =   -1  'True
               TabIndex        =   144
               Top             =   1920
               Width           =   8985
            End
            Begin VB.TextBox TXT_GB06_TEXT2 
               Appearance      =   0  'Flat
               Height          =   705
               Left            =   -66720
               MultiLine       =   -1  'True
               TabIndex        =   143
               Top             =   420
               Width           =   8985
            End
            Begin VB.TextBox TXT_GB06_TEXT1 
               Appearance      =   0  'Flat
               Height          =   705
               Left            =   -74910
               MultiLine       =   -1  'True
               TabIndex        =   142
               Top             =   420
               Width           =   7935
            End
            Begin VB.TextBox TXT_GB06_TEXT3 
               Appearance      =   0  'Flat
               Height          =   2925
               Left            =   -74880
               MultiLine       =   -1  'True
               TabIndex        =   141
               Top             =   1920
               Width           =   7905
            End
            Begin TMS_QGRID.TMS_QGRIDWRAPPER TMS_QGRIDGRUPPI 
               Height          =   2895
               Left            =   -74970
               TabIndex        =   140
               Top             =   1050
               Width           =   10065
               _ExtentX        =   17754
               _ExtentY        =   5106
               GridLoadMode    =   0
               DefaultRowHeight=   18
               BeginProperty PreviewFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PreviewFontColor=   0
            End
            Begin PRJFW_GRIDNAV.TMS_GRIDNAV TMS_GRIDNAV_GRUPPI 
               Height          =   1815
               Left            =   -64890
               TabIndex        =   139
               Top             =   1200
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   3201
               Primo           =   0   'False
               Precedente      =   0   'False
               Successivo      =   0   'False
               Ultimo          =   0   'False
               Apri            =   0   'False
               Orizzontale     =   0   'False
            End
            Begin VB.PictureBox PictureArticoli 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               DataSource      =   "rstCorpoBanco(""GB07_IMG"")"
               ForeColor       =   &H80000008&
               Height          =   5385
               Left            =   -74970
               ScaleHeight     =   5355
               ScaleWidth      =   16965
               TabIndex        =   138
               Top             =   660
               Width           =   16995
            End
            Begin TMS_QGRID.TMS_QGRIDWRAPPER QGRID_NOTE 
               Height          =   3015
               Left            =   -74970
               TabIndex        =   137
               Top             =   1470
               Width           =   17055
               _ExtentX        =   30083
               _ExtentY        =   5318
               GridLoadMode    =   0
               DefaultRowHeight=   18
               BeginProperty PreviewFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PreviewFontColor=   0
            End
            Begin PRJFW_FRAME.TMS_FRAME TMS_FRAME6 
               Height          =   1095
               Left            =   -74970
               Top             =   360
               Width           =   17505
               _ExtentX        =   30877
               _ExtentY        =   1931
               Begin PRJFW_EDITDATEM.EditDateM TXT_GB08_DATA 
                  Height          =   300
                  Left            =   990
                  TabIndex        =   133
                  Top             =   60
                  Width           =   1425
                  _ExtentX        =   2514
                  _ExtentY        =   529
                  IsCalendario    =   0   'False
                  DBField         =   "GB08_DATA"
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
                  Height          =   300
                  Index           =   30
                  Left            =   90
                  TabIndex        =   136
                  TabStop         =   0   'False
                  Top             =   420
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   529
                  Caption         =   "Annotazione"
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
                  Height          =   300
                  Index           =   29
                  Left            =   90
                  TabIndex        =   135
                  TabStop         =   0   'False
                  Top             =   60
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   529
                  Caption         =   "Data Nota"
               End
               Begin PRJFW_EDIT.TxtEdit TXT_GB08_TESTONOTA 
                  Height          =   600
                  Left            =   990
                  TabIndex        =   134
                  Top             =   420
                  Width           =   16455
                  _ExtentX        =   29025
                  _ExtentY        =   529
                  DisplayFormat   =   "Qualsiasi"
                  MaxChar         =   1500
                  Numerico        =   0   'False
                  Carattere       =   0   'False
                  DBField         =   "GB08_TESTONOTA"
                  MaxWidth        =   135
               End
            End
            Begin PRJFW_FRAME.TMS_FRAME TMS_FRAME3 
               Height          =   1425
               Left            =   -74970
               Top             =   360
               Width           =   17625
               _ExtentX        =   31089
               _ExtentY        =   2514
               Begin PRJFW_EDITM.TXT_EDITM TXT_FAM 
                  Height          =   300
                  Left            =   7020
                  TabIndex        =   102
                  ToolTipText     =   "Articolo"
                  Top             =   60
                  Width           =   1425
                  _ExtentX        =   2514
                  _ExtentY        =   529
                  IsLookup        =   -1  'True
                  DisplayFormat   =   "Maiuscolo"
                  MaxChar         =   6
                  Numerico        =   0   'False
                  Carattere       =   0   'False
                  IsDbField       =   0   'False
                  DBField         =   "GB07_CODART_MG66"
                  IsDecode        =   -1  'True
                  Caption         =   "Articolo"
                  NumRighe        =   0
                  Object.Tag             =   "Articolo"
                  MaxWidth        =   6
                  IsInLingua      =   0   'False
                  LinguaEntitaDes =   0
                  LinguaIDProvenienzaExt=   ""
               End
               Begin PRJFW_EDITM.TXT_EDITM TXT_SFAM 
                  Height          =   300
                  Left            =   7020
                  TabIndex        =   103
                  ToolTipText     =   "Articolo"
                  Top             =   390
                  Width           =   1425
                  _ExtentX        =   2514
                  _ExtentY        =   529
                  IsLookup        =   -1  'True
                  DisplayFormat   =   "Maiuscolo"
                  MaxChar         =   6
                  Numerico        =   0   'False
                  Carattere       =   0   'False
                  IsDbField       =   0   'False
                  DBField         =   "GB07_CODART_MG66"
                  IsDecode        =   -1  'True
                  Caption         =   "Articolo"
                  NumRighe        =   0
                  Object.Tag             =   "Articolo"
                  MaxWidth        =   6
                  IsInLingua      =   0   'False
                  LinguaEntitaDes =   0
                  LinguaIDProvenienzaExt=   ""
               End
               Begin PRJFW_EDITM.TXT_EDITM TXT_GRUPPO 
                  Height          =   300
                  Left            =   12810
                  TabIndex        =   104
                  ToolTipText     =   "Articolo"
                  Top             =   60
                  Width           =   1425
                  _ExtentX        =   2514
                  _ExtentY        =   529
                  IsLookup        =   -1  'True
                  DisplayFormat   =   "Maiuscolo"
                  MaxChar         =   6
                  Numerico        =   0   'False
                  Carattere       =   0   'False
                  IsDbField       =   0   'False
                  DBField         =   "GB07_CODART_MG66"
                  IsDecode        =   -1  'True
                  Caption         =   "Articolo"
                  NumRighe        =   0
                  Object.Tag             =   "Articolo"
                  MaxWidth        =   6
                  IsInLingua      =   0   'False
                  LinguaEntitaDes =   0
                  LinguaIDProvenienzaExt=   ""
               End
               Begin PRJFW_EDITM.TXT_EDITM TXT_SGRUPPO 
                  Height          =   300
                  Left            =   12810
                  TabIndex        =   105
                  ToolTipText     =   "Articolo"
                  Top             =   390
                  Width           =   1425
                  _ExtentX        =   2514
                  _ExtentY        =   529
                  IsLookup        =   -1  'True
                  DisplayFormat   =   "Maiuscolo"
                  MaxChar         =   6
                  Numerico        =   0   'False
                  Carattere       =   0   'False
                  IsDbField       =   0   'False
                  DBField         =   "GB07_CODART_MG66"
                  IsDecode        =   -1  'True
                  Caption         =   "Articolo"
                  NumRighe        =   0
                  Object.Tag             =   "Articolo"
                  MaxWidth        =   6
                  IsInLingua      =   0   'False
                  LinguaEntitaDes =   0
                  LinguaIDProvenienzaExt=   ""
               End
               Begin PRJFW_EDITM.TXT_EDITM TXT_GRST1 
                  Height          =   300
                  Left            =   7020
                  TabIndex        =   106
                  ToolTipText     =   "Articolo"
                  Top             =   720
                  Width           =   1425
                  _ExtentX        =   2514
                  _ExtentY        =   529
                  IsLookup        =   -1  'True
                  DisplayFormat   =   "Maiuscolo"
                  MaxChar         =   6
                  Numerico        =   0   'False
                  Carattere       =   0   'False
                  IsDbField       =   0   'False
                  DBField         =   "GB07_CODART_MG66"
                  IsDecode        =   -1  'True
                  Caption         =   "Articolo"
                  NumRighe        =   0
                  Object.Tag             =   "Articolo"
                  MaxWidth        =   6
                  IsInLingua      =   0   'False
                  LinguaEntitaDes =   0
                  LinguaIDProvenienzaExt=   ""
               End
               Begin PRJFW_EDITM.TXT_EDITM TXT_GRST2 
                  Height          =   300
                  Left            =   7020
                  TabIndex        =   107
                  ToolTipText     =   "Articolo"
                  Top             =   1050
                  Width           =   1425
                  _ExtentX        =   2514
                  _ExtentY        =   529
                  IsLookup        =   -1  'True
                  DisplayFormat   =   "Maiuscolo"
                  MaxChar         =   6
                  Numerico        =   0   'False
                  Carattere       =   0   'False
                  IsDbField       =   0   'False
                  DBField         =   "GB07_CODART_MG66"
                  IsDecode        =   -1  'True
                  Caption         =   "Articolo"
                  NumRighe        =   0
                  Object.Tag             =   "Articolo"
                  MaxWidth        =   6
                  IsInLingua      =   0   'False
                  LinguaEntitaDes =   0
                  LinguaIDProvenienzaExt=   ""
               End
               Begin PRJFW_EDITM.TXT_EDITM TXT_GRST3 
                  Height          =   300
                  Left            =   12810
                  TabIndex        =   108
                  ToolTipText     =   "Articolo"
                  Top             =   720
                  Width           =   1425
                  _ExtentX        =   2514
                  _ExtentY        =   529
                  IsLookup        =   -1  'True
                  DisplayFormat   =   "Maiuscolo"
                  MaxChar         =   6
                  Numerico        =   0   'False
                  Carattere       =   0   'False
                  IsDbField       =   0   'False
                  DBField         =   "GB07_CODART_MG66"
                  IsDecode        =   -1  'True
                  Caption         =   "Articolo"
                  NumRighe        =   0
                  Object.Tag             =   "Articolo"
                  MaxWidth        =   6
                  IsInLingua      =   0   'False
                  LinguaEntitaDes =   0
                  LinguaIDProvenienzaExt=   ""
               End
               Begin PRJFW_EDITM.TXT_EDITM TXT_GRST4 
                  Height          =   300
                  Left            =   12810
                  TabIndex        =   109
                  ToolTipText     =   "Articolo"
                  Top             =   1050
                  Width           =   1425
                  _ExtentX        =   2514
                  _ExtentY        =   529
                  IsLookup        =   -1  'True
                  DisplayFormat   =   "Maiuscolo"
                  MaxChar         =   6
                  Numerico        =   0   'False
                  Carattere       =   0   'False
                  IsDbField       =   0   'False
                  DBField         =   "GB07_CODART_MG66"
                  IsDecode        =   -1  'True
                  Caption         =   "Articolo"
                  NumRighe        =   0
                  Object.Tag             =   "Articolo"
                  MaxWidth        =   6
                  IsInLingua      =   0   'False
                  LinguaEntitaDes =   0
                  LinguaIDProvenienzaExt=   ""
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
                  Height          =   300
                  Index           =   49
                  Left            =   11850
                  TabIndex        =   132
                  TabStop         =   0   'False
                  Top             =   1050
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   529
                  Caption         =   "Gr. Stat4"
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
                  Height          =   300
                  Index           =   48
                  Left            =   11850
                  TabIndex        =   131
                  TabStop         =   0   'False
                  Top             =   720
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   529
                  Caption         =   "Gr. Stat3"
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
                  Height          =   300
                  Index           =   47
                  Left            =   6000
                  TabIndex        =   130
                  TabStop         =   0   'False
                  Top             =   1080
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   529
                  Caption         =   "Gr. Stat2"
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
                  Height          =   300
                  Index           =   46
                  Left            =   6000
                  TabIndex        =   129
                  TabStop         =   0   'False
                  Top             =   750
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   529
                  Caption         =   "Gr. Stat1"
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
                  Height          =   300
                  Index           =   45
                  Left            =   6000
                  TabIndex        =   128
                  TabStop         =   0   'False
                  Top             =   420
                  Width           =   1140
                  _ExtentX        =   2011
                  _ExtentY        =   529
                  Caption         =   "Sotto Fam"
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
                  Height          =   300
                  Index           =   44
                  Left            =   90
                  TabIndex        =   127
                  TabStop         =   0   'False
                  Top             =   90
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   529
                  Caption         =   "Tipo Operazione"
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
                  Height          =   300
                  Index           =   34
                  Left            =   90
                  TabIndex        =   126
                  TabStop         =   0   'False
                  Top             =   420
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   529
                  Caption         =   "Valore"
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
                  Height          =   300
                  Index           =   33
                  Left            =   6000
                  TabIndex        =   125
                  TabStop         =   0   'False
                  Top             =   90
                  Width           =   1170
                  _ExtentX        =   2064
                  _ExtentY        =   529
                  Caption         =   "Famiglia"
               End
               Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON CMD_CONFERMA 
                  Height          =   495
                  Left            =   3540
                  TabIndex        =   124
                  Top             =   780
                  Width           =   2085
                  _ExtentX        =   3678
                  _ExtentY        =   873
                  Enabled         =   0   'False
                  Caption         =   "Conferma Simulazione"
                  ButtonBorderColor=   0
                  ButtonForeColor =   16777215
                  ButtonHilightBorderColor=   9473677
               End
               Begin PRJFW_CHECKBOX.TMS_CHECKBOX TMS_MANTIENI 
                  Height          =   300
                  Left            =   2490
                  TabIndex        =   123
                  Top             =   450
                  Width           =   3075
                  _ExtentX        =   5424
                  _ExtentY        =   529
                  IsDbField       =   0   'False
                  Caption         =   "Accoda variazioni"
               End
               Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON CMD_SIMULA 
                  Height          =   495
                  Left            =   1410
                  TabIndex        =   122
                  Top             =   780
                  Width           =   2085
                  _ExtentX        =   3678
                  _ExtentY        =   873
                  Caption         =   "Simula"
                  ButtonBorderColor=   0
                  ButtonForeColor =   16777215
                  ButtonHilightBorderColor=   9473677
               End
               Begin PRJFW_COMBOBOX.TMS_COMBO CMB_TIPOVARIAZIONE 
                  Height          =   315
                  Left            =   1260
                  TabIndex        =   121
                  Top             =   60
                  Width           =   4650
                  _ExtentX        =   8202
                  _ExtentY        =   556
                  MaxChar         =   45
                  IsDbField       =   0   'False
                  DbCol           =   0
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TXT_GRST4_DEC 
                  Height          =   300
                  Left            =   14520
                  TabIndex        =   120
                  TabStop         =   0   'False
                  Tag             =   "Descr. fornitore ENECO"
                  Top             =   1050
                  Width           =   3015
                  _ExtentX        =   5318
                  _ExtentY        =   529
                  Caption         =   ""
                  Border          =   -1  'True
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TXT_GRST3_DEC 
                  Height          =   300
                  Left            =   14520
                  TabIndex        =   119
                  TabStop         =   0   'False
                  Tag             =   "Descr. fornitore ENECO"
                  Top             =   720
                  Width           =   3015
                  _ExtentX        =   5318
                  _ExtentY        =   529
                  Caption         =   ""
                  Border          =   -1  'True
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TXT_GRST2_DEC 
                  Height          =   300
                  Left            =   8730
                  TabIndex        =   118
                  TabStop         =   0   'False
                  Tag             =   "Descr. fornitore ENECO"
                  Top             =   1050
                  Width           =   3015
                  _ExtentX        =   5318
                  _ExtentY        =   529
                  Caption         =   ""
                  Border          =   -1  'True
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TXT_GRST1_DEC 
                  Height          =   300
                  Left            =   8730
                  TabIndex        =   117
                  TabStop         =   0   'False
                  Tag             =   "Descr. fornitore ENECO"
                  Top             =   720
                  Width           =   3015
                  _ExtentX        =   5318
                  _ExtentY        =   529
                  Caption         =   ""
                  Border          =   -1  'True
               End
               Begin PRJFW_EDITNUM.TxtEditNum TXT_PRECENTUALE 
                  Height          =   300
                  Left            =   1260
                  TabIndex        =   116
                  Top             =   420
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   529
                  Obbligatorio    =   -1  'True
                  DBField         =   "GB07_PREZZO"
                  Caption         =   "Prezzo"
                  Object.Tag             =   "Prezzo"
                  MaxWidth        =   5
                  MaxChar         =   17
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TXT_GRUPPO_DEC 
                  Height          =   300
                  Left            =   14520
                  TabIndex        =   115
                  TabStop         =   0   'False
                  Tag             =   "Descr. fornitore ENECO"
                  Top             =   60
                  Width           =   3015
                  _ExtentX        =   5318
                  _ExtentY        =   529
                  Caption         =   ""
                  Border          =   -1  'True
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL54 
                  Height          =   300
                  Left            =   11850
                  TabIndex        =   114
                  TabStop         =   0   'False
                  Top             =   90
                  Width           =   1170
                  _ExtentX        =   2064
                  _ExtentY        =   529
                  Caption         =   "Gruppo"
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TXT_SGRUPPO_DEC 
                  Height          =   300
                  Left            =   14520
                  TabIndex        =   113
                  TabStop         =   0   'False
                  Tag             =   "Descr. fornitore ENECO"
                  Top             =   390
                  Width           =   3015
                  _ExtentX        =   5318
                  _ExtentY        =   529
                  Caption         =   ""
                  Border          =   -1  'True
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL52 
                  Height          =   300
                  Left            =   11850
                  TabIndex        =   112
                  TabStop         =   0   'False
                  Top             =   420
                  Width           =   1170
                  _ExtentX        =   2064
                  _ExtentY        =   529
                  Caption         =   "Sotto Gruppo"
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TXT_FAM_DEC 
                  Height          =   300
                  Left            =   8730
                  TabIndex        =   111
                  TabStop         =   0   'False
                  Tag             =   "Descr. fornitore ENECO"
                  Top             =   60
                  Width           =   3015
                  _ExtentX        =   5318
                  _ExtentY        =   529
                  Caption         =   ""
                  Border          =   -1  'True
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TXT_SFMA_DEC 
                  Height          =   300
                  Left            =   8730
                  TabIndex        =   110
                  TabStop         =   0   'False
                  Tag             =   "Descr. fornitore ENECO"
                  Top             =   390
                  Width           =   3015
                  _ExtentX        =   5318
                  _ExtentY        =   529
                  Caption         =   ""
                  Border          =   -1  'True
               End
            End
            Begin TMS_QGRID.TMS_QGRIDWRAPPER QGRID_SIMULAZIONE 
               Height          =   4455
               Left            =   -75000
               TabIndex        =   101
               Top             =   1800
               Width           =   17625
               _ExtentX        =   31089
               _ExtentY        =   7858
               GridLoadMode    =   0
               DefaultRowHeight=   18
               BeginProperty PreviewFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PreviewFontColor=   0
            End
            Begin PRJFW_FRAME.TMS_FRAME FrameCorpo 
               Height          =   1455
               Left            =   0
               Top             =   360
               Width           =   17625
               _ExtentX        =   31089
               _ExtentY        =   2566
               Begin PRJFW_RICHTEXTBOX.TmsRichTextBox TXT_DESART 
                  Height          =   300
                  Left            =   3600
                  TabIndex        =   220
                  Top             =   390
                  Width           =   10095
                  _ExtentX        =   17806
                  _ExtentY        =   529
                  MaxChar         =   1672
                  Numerico        =   0   'False
                  Carattere       =   0   'False
                  DBField         =   "GB07_DESCART"
                  NumRighe        =   5
                  IsExpand        =   -1  'True
                  MaxWidth        =   82
                  IsInLingua      =   0   'False
                  LinguaEntitaDes =   0
                  LinguaIDProvenienzaExt=   0
               End
               Begin PRJFW_EDITM.TXT_EDITM TXT_GB07_CODART_MG66 
                  Height          =   300
                  Left            =   1080
                  TabIndex        =   52
                  ToolTipText     =   "Articolo"
                  Top             =   390
                  Width           =   2505
                  _ExtentX        =   4419
                  _ExtentY        =   529
                  IsLookup        =   -1  'True
                  DisplayFormat   =   "Maiuscolo"
                  MaxChar         =   25
                  Numerico        =   0   'False
                  Carattere       =   0   'False
                  DBField         =   "GB07_CODART_MG66"
                  IsDecode        =   -1  'True
                  Caption         =   "Articolo"
                  NumRighe        =   0
                  Object.Tag             =   "Articolo"
                  MaxWidth        =   15
                  IsInLingua      =   0   'False
                  LinguaEntitaDes =   0
                  LinguaIDProvenienzaExt=   ""
               End
               Begin PRJFW_EDITPATH.TXT_EDITPATH TXT_GB07_IMAGEPATH 
                  DataField       =   "GB07_IMG"
                  Height          =   345
                  Left            =   8160
                  TabIndex        =   53
                  Top             =   30
                  Width           =   5505
                  _ExtentX        =   9737
                  _ExtentY        =   635
                  IsLookup        =   -1  'True
                  Carattere       =   0   'False
                  DBField         =   "GB07_PATHIMG"
                  NumRighe        =   0
               End
               Begin PRJFW_EDITNUMM.TxtEditNumM TXT_GB07_CLIFOR_CG44 
                  Height          =   300
                  Left            =   3600
                  TabIndex        =   54
                  Top             =   720
                  Width           =   2280
                  _ExtentX        =   4022
                  _ExtentY        =   529
                  IsLookup        =   -1  'True
                  IsCalculator    =   0   'False
                  MaxChar         =   15
                  DBField         =   "GB07_CLIFOR_CG44"
                  Caption         =   "Codice cliente"
                  Object.Tag             =   "Codice cliente"
                  IsDecode        =   -1  'True
               End
               Begin PRJFW_EDITNUMM.TxtEditNumM TXT_GB07_RAG 
                  Height          =   300
                  Left            =   3150
                  TabIndex        =   55
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   529
                  IsLookup        =   -1  'True
                  MaxChar         =   2
                  DBField         =   "GB07_RAG"
                  MaxWidth        =   2
               End
               Begin PRJFW_EDITNUMM.TxtEditNumM TXT_CLIFOR_CG44 
                  Height          =   300
                  Left            =   14970
                  TabIndex        =   56
                  Top             =   720
                  Width           =   2280
                  _ExtentX        =   4022
                  _ExtentY        =   529
                  IsLookup        =   -1  'True
                  IsCalculator    =   0   'False
                  MaxChar         =   15
                  DBField         =   "GB07_CLIFOR_CG44"
                  Caption         =   "Codice cliente"
                  Object.Tag             =   "Codice cliente"
                  IsDecode        =   -1  'True
               End
               Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON CMD_COSTO 
                  Height          =   285
                  Left            =   7320
                  TabIndex        =   221
                  Top             =   60
                  Width           =   225
                  _ExtentX        =   397
                  _ExtentY        =   503
                  Caption         =   "C"
                  ButtonBorderColor=   0
                  ButtonForeColor =   16777215
                  ButtonHilightBorderColor=   9473677
               End
               Begin PRJFW_EDITNUM.TxtEditNum TXT_SC1_FISSO 
                  Height          =   300
                  Left            =   15720
                  TabIndex        =   58
                  Top             =   390
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   529
                  DBField         =   "GB07_SC1"
                  Caption         =   "SCONTO RIGA"
                  Object.Tag             =   "SCONTO RIGA"
                  MaxWidth        =   4
                  MaxChar         =   6
                  TipoFormato     =   2
               End
               Begin PRJFW_EDITNUM.TxtEditNum TXT_GB07_IMPPROVV 
                  Height          =   300
                  Left            =   13560
                  TabIndex        =   57
                  Top             =   1050
                  Width           =   1500
                  _ExtentX        =   2646
                  _ExtentY        =   529
                  Enabled         =   0   'False
                  DBField         =   "GB07_IMPPROVV"
                  Caption         =   "Importo Provv"
                  Object.Tag             =   "Importo Provv"
                  MaxWidth        =   8
                  MaxChar         =   17
                  FormatMask      =   """€"" ###,###,###,##0.00"
                  TipoFormato     =   2
               End
               Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON CMD_REFRESH 
                  Height          =   285
                  Left            =   6900
                  TabIndex        =   76
                  Top             =   60
                  Width           =   405
                  _ExtentX        =   714
                  _ExtentY        =   503
                  Caption         =   "Refresh"
                  ButtonBorderColor=   0
                  ButtonForeColor =   16777215
                  ButtonHilightBorderColor=   9473677
               End
               Begin PRJFW_EDITNUM.TxtEditNum TXT_GB07_SC6 
                  Height          =   300
                  Left            =   14010
                  TabIndex        =   100
                  Top             =   1020
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   529
                  Object.Visible         =   0   'False
                  DBField         =   "GB07_SC6"
                  Caption         =   "SCONTO RIGA"
                  Object.Tag             =   "SCONTO RIGA"
                  MaxWidth        =   4
                  MaxChar         =   6
                  TipoFormato     =   2
               End
               Begin PRJFW_EDIT.TxtEdit TXT_GB07_ID 
                  Height          =   300
                  Left            =   7410
                  TabIndex        =   94
                  Top             =   60
                  Width           =   735
                  _ExtentX        =   1296
                  _ExtentY        =   529
                  MaxChar         =   15
                  Numerico        =   0   'False
                  Carattere       =   0   'False
                  DBField         =   "GB07_ID"
                  MaxWidth        =   4
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL1 
                  Height          =   300
                  Left            =   4380
                  TabIndex        =   92
                  TabStop         =   0   'False
                  Top             =   90
                  Width           =   1470
                  _ExtentX        =   2593
                  _ExtentY        =   529
                  Caption         =   "Costo (List. Acq. 1)"
               End
               Begin PRJFW_EDITNUM.TxtEditNum TXT_GB07_COSTO 
                  Height          =   300
                  Left            =   5700
                  TabIndex        =   91
                  Top             =   60
                  Width           =   1170
                  _ExtentX        =   2064
                  _ExtentY        =   529
                  DBField         =   "GB07_COSTO"
                  Caption         =   "Costo"
                  Object.Tag             =   "Costo"
                  MaxWidth        =   6
                  MaxChar         =   17
                  FormatMask      =   """€"" ###,###,###,##0.00"
                  TipoFormato     =   2
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
                  Height          =   300
                  Index           =   0
                  Left            =   10350
                  TabIndex        =   90
                  TabStop         =   0   'False
                  Top             =   1080
                  Width           =   930
                  _ExtentX        =   1640
                  _ExtentY        =   529
                  Caption         =   "Importo riga"
               End
               Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON CMD_RICPERCPROVV 
                  Height          =   285
                  Left            =   8100
                  TabIndex        =   89
                  Top             =   1050
                  Width           =   1515
                  _ExtentX        =   2672
                  _ExtentY        =   503
                  Caption         =   "Ric. % Provv."
                  ButtonBorderColor=   0
                  ButtonForeColor =   16777215
                  ButtonHilightBorderColor=   9473677
               End
               Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON TMS_FLATBUTTON1 
                  Height          =   255
                  Left            =   15570
                  TabIndex        =   88
                  Top             =   90
                  Width           =   1665
                  _ExtentX        =   2937
                  _ExtentY        =   450
                  Caption         =   "Carica Immagine"
                  ButtonBorderColor=   0
                  ButtonForeColor =   16777215
                  ButtonHilightBorderColor=   9473677
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
                  Height          =   300
                  Index           =   5
                  Left            =   12630
                  TabIndex        =   87
                  TabStop         =   0   'False
                  Top             =   1050
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   529
                  Caption         =   "Importo Provv."
               End
               Begin PRJFW_EDITNUM.TxtEditNum TXT_GB07_PERCPROVV 
                  Height          =   300
                  Left            =   9480
                  TabIndex        =   86
                  Top             =   1050
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   529
                  DBField         =   "GB07_PERCPROVV"
                  Caption         =   "SCONTO RIGA"
                  Object.Tag             =   "SCONTO RIGA"
                  MaxWidth        =   4
                  MaxChar         =   6
                  TipoFormato     =   2
               End
               Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON CMD_AGGIORNA 
                  Height          =   255
                  Left            =   15150
                  TabIndex        =   85
                  Top             =   1080
                  Width           =   2055
                  _ExtentX        =   3625
                  _ExtentY        =   450
                  Caption         =   "Aggiorna Fornitori"
                  ButtonBorderColor=   0
                  ButtonForeColor =   16777215
                  ButtonHilightBorderColor=   9473677
               End
               Begin PRJFW_CHECKBOX.TMS_CHECKBOX CHK_GB07_CHECK 
                  Height          =   300
                  Left            =   2250
                  TabIndex        =   84
                  Top             =   30
                  Width           =   255
                  _ExtentX        =   450
                  _ExtentY        =   529
                  Object.Visible         =   0   'False
                  DBField         =   "GB07_CHECK"
               End
               Begin PRJFW_EDITNUM.TxtEditNum TXT_GB07_SC5 
                  Height          =   300
                  Left            =   14010
                  TabIndex        =   83
                  Top             =   1020
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   529
                  Object.Visible         =   0   'False
                  DBField         =   "GB07_SC5"
                  Caption         =   "SCONTO RIGA"
                  Object.Tag             =   "SCONTO RIGA"
                  MaxWidth        =   4
                  MaxChar         =   6
                  TipoFormato     =   2
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
                  Height          =   300
                  Index           =   2
                  Left            =   13800
                  TabIndex        =   78
                  TabStop         =   0   'False
                  Top             =   720
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   529
                  Caption         =   "Fornitore/Posa"
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
                  Height          =   300
                  Index           =   1
                  Left            =   14850
                  TabIndex        =   77
                  TabStop         =   0   'False
                  Top             =   420
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   529
                  Caption         =   "Sc1/Sc2 Fisso"
               End
               Begin PRJFW_EDIT.TxtEdit TXT_GB07_SEQ 
                  Height          =   300
                  Left            =   2520
                  TabIndex        =   75
                  Top             =   30
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   529
                  MaxChar         =   3
                  Numerico        =   0   'False
                  Carattere       =   0   'False
                  DBField         =   "GB07_SEQ"
                  MaxWidth        =   3
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL39 
                  Height          =   300
                  Left            =   2940
                  TabIndex        =   74
                  TabStop         =   0   'False
                  Top             =   720
                  Width           =   810
                  _ExtentX        =   1429
                  _ExtentY        =   529
                  Caption         =   "Fornitore"
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TXT_CG16_RAGSOANAG_FOR 
                  Height          =   300
                  Left            =   5940
                  TabIndex        =   73
                  TabStop         =   0   'False
                  Tag             =   "Descr. fornitore ENECO"
                  Top             =   720
                  Width           =   7785
                  _ExtentX        =   13732
                  _ExtentY        =   529
                  Caption         =   ""
                  Border          =   -1  'True
               End
               Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON TMS_FLATBUTTON4 
                  Height          =   255
                  Left            =   13710
                  TabIndex        =   72
                  Top             =   90
                  Width           =   1725
                  _ExtentX        =   3043
                  _ExtentY        =   450
                  Caption         =   "Visualizza Immagine"
                  ButtonBorderColor=   0
                  ButtonForeColor =   16777215
                  ButtonHilightBorderColor=   9473677
               End
               Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON TMS_FLATBUTTON3 
                  Height          =   345
                  Left            =   10950
                  TabIndex        =   71
                  Top             =   2310
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   609
                  Caption         =   "Immagine"
                  ButtonBorderColor=   0
                  ButtonForeColor =   16777215
                  ButtonHilightBorderColor=   9473677
               End
               Begin PRJFW_EDITNUM.TxtEditNum TXT_GB07_QTA 
                  Height          =   300
                  Left            =   1080
                  TabIndex        =   70
                  Top             =   720
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   529
                  DBField         =   "GB07_QTA"
                  Caption         =   "Quantità"
                  Object.Tag             =   "Quantità"
                  MaxWidth        =   5
                  MaxChar         =   14
                  TipoFormato     =   3
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL10 
                  Height          =   300
                  Left            =   150
                  TabIndex        =   69
                  TabStop         =   0   'False
                  Top             =   420
                  Width           =   1380
                  _ExtentX        =   2434
                  _ExtentY        =   529
                  Caption         =   "Articolo"
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL8 
                  Height          =   300
                  Left            =   180
                  TabIndex        =   67
                  TabStop         =   0   'False
                  Top             =   1080
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   529
                  Caption         =   "Prezzo"
               End
               Begin PRJFW_EDITNUM.TxtEditNum TXT_GB07_PREZZO 
                  Height          =   300
                  Left            =   1080
                  TabIndex        =   66
                  Top             =   1050
                  Width           =   1170
                  _ExtentX        =   2064
                  _ExtentY        =   529
                  DBField         =   "GB07_PREZZO"
                  Caption         =   "Prezzo"
                  Object.Tag             =   "Prezzo"
                  MaxWidth        =   6
                  MaxChar         =   17
                  FormatMask      =   """€"" ###,###,###,##0.00"
                  TipoFormato     =   2
               End
               Begin PRJFW_EDITNUM.TxtEditNum TXT_GB07_IMPORTO 
                  Height          =   300
                  Left            =   11100
                  TabIndex        =   65
                  Top             =   1050
                  Width           =   1500
                  _ExtentX        =   2646
                  _ExtentY        =   529
                  Enabled         =   0   'False
                  DBField         =   "GB07_IMPORTO"
                  Caption         =   "Importo"
                  Object.Tag             =   "Importo"
                  MaxWidth        =   8
                  MaxChar         =   17
                  FormatMask      =   """€"" ###,###,###,##0.00"
                  TipoFormato     =   2
               End
               Begin PRJFW_EDITNUM.TxtEditNum TXT_SC2_FISSO 
                  Height          =   300
                  Left            =   16410
                  TabIndex        =   64
                  Top             =   390
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   529
                  DBField         =   "GB07_SC1"
                  Caption         =   "SCONTO RIGA"
                  Object.Tag             =   "SCONTO RIGA"
                  MaxWidth        =   4
                  MaxChar         =   6
                  TipoFormato     =   2
               End
               Begin PRJFW_EDITNUM.TxtEditNum TXT_GB07_SC1 
                  Height          =   300
                  Left            =   2880
                  TabIndex        =   63
                  Top             =   1050
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   529
                  DBField         =   "GB07_SC1"
                  Caption         =   "SCONTO RIGA"
                  Object.Tag             =   "SCONTO RIGA"
                  MaxWidth        =   4
                  MaxChar         =   6
                  TipoFormato     =   2
               End
               Begin PRJFW_EDITNUM.TxtEditNum TXT_GB07_SC2 
                  Height          =   300
                  Left            =   4350
                  TabIndex        =   62
                  Top             =   1050
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   529
                  DBField         =   "GB07_SC2"
                  Caption         =   "SCONTO RIGA"
                  Object.Tag             =   "SCONTO RIGA"
                  MaxWidth        =   4
                  MaxChar         =   6
                  TipoFormato     =   2
               End
               Begin PRJFW_EDITNUM.TxtEditNum TXT_GB07_SC3 
                  Height          =   300
                  Left            =   5820
                  TabIndex        =   61
                  Top             =   1050
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   529
                  DBField         =   "GB07_SC3"
                  Caption         =   "SCONTO RIGA"
                  Object.Tag             =   "SCONTO RIGA"
                  MaxWidth        =   4
                  MaxChar         =   6
                  TipoFormato     =   2
               End
               Begin PRJFW_EDITNUM.TxtEditNum TXT_GB07_SC4 
                  Height          =   300
                  Left            =   7170
                  TabIndex        =   60
                  Top             =   1050
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   529
                  DBField         =   "GB07_SC4"
                  Caption         =   "SCONTO RIGA"
                  Object.Tag             =   "SCONTO RIGA"
                  MaxWidth        =   4
                  MaxChar         =   6
                  TipoFormato     =   2
               End
               Begin PRJFW_EDIT.TxtEdit TXT_DES1ART 
                  Height          =   300
                  Left            =   3600
                  TabIndex        =   59
                  Top             =   390
                  Width           =   10095
                  _ExtentX        =   17806
                  _ExtentY        =   529
                  DisplayFormat   =   "Qualsiasi"
                  MaxChar         =   1600
                  Numerico        =   0   'False
                  Carattere       =   0   'False
                  DBField         =   "GB07_DESCART"
                  MaxWidth        =   82
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL12 
                  Height          =   300
                  Left            =   150
                  TabIndex        =   95
                  TabStop         =   0   'False
                  Top             =   750
                  Width           =   1290
                  _ExtentX        =   2275
                  _ExtentY        =   529
                  Caption         =   "Quantità"
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
                  Height          =   300
                  Index           =   3
                  Left            =   2400
                  TabIndex        =   79
                  TabStop         =   0   'False
                  Top             =   1080
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   529
                  Caption         =   "Sc.1%"
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
                  Height          =   300
                  Index           =   8
                  Left            =   3810
                  TabIndex        =   82
                  TabStop         =   0   'False
                  Top             =   1050
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   529
                  Caption         =   "Sc.2%"
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
                  Height          =   300
                  Index           =   7
                  Left            =   5340
                  TabIndex        =   81
                  TabStop         =   0   'False
                  Top             =   1050
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   529
                  Caption         =   "Sc.3%"
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
                  Height          =   300
                  Index           =   6
                  Left            =   6720
                  TabIndex        =   80
                  TabStop         =   0   'False
                  Top             =   1050
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   529
                  Caption         =   "Sc.4%"
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
                  Height          =   300
                  Index           =   4
                  Left            =   8100
                  TabIndex        =   93
                  TabStop         =   0   'False
                  Top             =   1080
                  Width           =   1410
                  _ExtentX        =   2487
                  _ExtentY        =   529
                  Caption         =   "Perc. Prov."
               End
               Begin PRJFW_EDITNUM.TxtEditNum TXT_GB07_TIPOCF_CG44 
                  Height          =   300
                  Left            =   7200
                  TabIndex        =   98
                  Top             =   60
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   529
                  Enabled         =   0   'False
                  Default         =   "1"
                  Object.Visible         =   0   'False
                  DBField         =   "GB07_TIPOCF_CG44"
                  MaxWidth        =   1
                  MaxChar         =   1
               End
               Begin PRJFW_EDITNUM.TxtEditNum TXT_GIAC 
                  Height          =   300
                  Left            =   6780
                  TabIndex        =   68
                  Top             =   360
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   529
                  Object.Visible         =   0   'False
                  DBField         =   "GB07_QTA"
                  Caption         =   "Quantità"
                  Object.Tag             =   "Quantità"
                  MaxWidth        =   5
                  MaxChar         =   14
               End
               Begin PRJFW_CHECKBOX.TMS_CHECKBOX TXT_GB07_FLPOSA 
                  Height          =   300
                  Left            =   13710
                  TabIndex        =   99
                  Top             =   390
                  Width           =   2445
                  _ExtentX        =   4313
                  _ExtentY        =   529
                  Default         =   "1"
                  DBField         =   "GB07_FLPOSA"
                  Caption         =   "Con Posa"
               End
               Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL67 
                  Height          =   300
                  Left            =   150
                  TabIndex        =   96
                  TabStop         =   0   'False
                  Top             =   60
                  Width           =   3600
                  _ExtentX        =   6350
                  _ExtentY        =   529
                  Caption         =   "Sequenza/Raggruppamento"
               End
               Begin PRJFW_EDIT.TxtEdit TXT_GB07_ALT 
                  Height          =   300
                  Left            =   2490
                  TabIndex        =   97
                  Top             =   30
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   529
                  MaxChar         =   3
                  Object.Visible         =   0   'False
                  Numerico        =   0   'False
                  Carattere       =   0   'False
                  DBField         =   "GB07_ALT"
                  MaxWidth        =   3
               End
            End
            Begin TMS_QGRID.TMS_QGRIDWRAPPER QGridDocumenti 
               Height          =   3615
               Left            =   360
               TabIndex        =   51
               Top             =   2430
               Width           =   17265
               _ExtentX        =   30454
               _ExtentY        =   6376
               GridLoadMode    =   0
               DefaultRowHeight=   18
               BeginProperty PreviewFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PreviewFontColor=   0
            End
            Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON CMD_DELALL 
               Height          =   1755
               Left            =   0
               TabIndex        =   223
               Top             =   3660
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   3096
               Caption         =   "D"
               ButtonBackColor =   255
               ButtonBorderColor=   255
               ButtonDisabledColor=   255
               ButtonForeColor =   0
               ButtonHilightBorderColor=   9473677
            End
            Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
               Height          =   300
               Index           =   32
               Left            =   -74880
               TabIndex        =   160
               TabStop         =   0   'False
               Top             =   690
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   529
               Caption         =   "Descrizione Gruppo"
            End
            Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
               Height          =   300
               Index           =   31
               Left            =   -74880
               TabIndex        =   159
               TabStop         =   0   'False
               Top             =   390
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   529
               Caption         =   "Codice Gruppo"
            End
            Begin VB.Label Label2 
               Caption         =   "Codice 99 per trasporto in calce all'offerta con descrizione personalizzata"
               Height          =   495
               Left            =   -64350
               TabIndex        =   158
               Top             =   1050
               Width           =   4305
            End
            Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON CMD_REFRESH_TXT4 
               Height          =   315
               Left            =   -57690
               TabIndex        =   157
               Top             =   4500
               Width           =   225
               _ExtentX        =   397
               _ExtentY        =   556
               Caption         =   "@"
               ButtonBorderColor=   0
               ButtonForeColor =   16777215
               ButtonHilightBorderColor=   9473677
            End
            Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON CMD_REFRESH_TXT6 
               Height          =   285
               Left            =   -57720
               TabIndex        =   156
               Top             =   1590
               Width           =   225
               _ExtentX        =   397
               _ExtentY        =   503
               Caption         =   "@"
               ButtonBorderColor=   0
               ButtonForeColor =   16777215
               ButtonHilightBorderColor=   9473677
            End
            Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON CMD_REFRESH_TXT2 
               Height          =   315
               Left            =   -57720
               TabIndex        =   155
               Top             =   780
               Width           =   225
               _ExtentX        =   397
               _ExtentY        =   556
               Caption         =   "@"
               ButtonBorderColor=   0
               ButtonForeColor =   16777215
               ButtonHilightBorderColor=   9473677
            End
            Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON CMD_REFRESH_TXT3 
               Height          =   285
               Left            =   -66930
               TabIndex        =   154
               Top             =   4560
               Width           =   225
               _ExtentX        =   397
               _ExtentY        =   503
               Caption         =   "@"
               ButtonBorderColor=   0
               ButtonForeColor =   16777215
               ButtonHilightBorderColor=   9473677
            End
            Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON CMD_REFRESH_TXT5 
               Height          =   285
               Left            =   -66960
               TabIndex        =   153
               Top             =   1590
               Width           =   225
               _ExtentX        =   397
               _ExtentY        =   503
               Caption         =   "@"
               ButtonBorderColor=   0
               ButtonForeColor =   16777215
               ButtonHilightBorderColor=   9473677
            End
            Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON CMD_REFRESH_TXT1 
               Height          =   285
               Left            =   -66960
               TabIndex        =   152
               Top             =   840
               Width           =   225
               _ExtentX        =   397
               _ExtentY        =   503
               Caption         =   "@"
               ButtonBorderColor=   0
               ButtonForeColor =   16777215
               ButtonHilightBorderColor=   9473677
            End
            Begin PRJFW_EDIT.TxtEdit TXT_GB01_DESCRIZIONE 
               Height          =   300
               Left            =   -73260
               TabIndex        =   151
               Top             =   690
               Width           =   8055
               _ExtentX        =   14208
               _ExtentY        =   529
               MaxChar         =   50
               Obbligatorio    =   -1  'True
               Numerico        =   0   'False
               Carattere       =   0   'False
               DBField         =   "GB01_DESCRIZIONE"
               MaxWidth        =   65
            End
            Begin PRJFW_EDITNUM.TxtEditNum TXT_GB01_PROG 
               Height          =   300
               Left            =   -73260
               TabIndex        =   150
               Top             =   360
               Width           =   510
               _ExtentX        =   900
               _ExtentY        =   529
               Obbligatorio    =   -1  'True
               DBField         =   "GB01_PROG"
               MaxWidth        =   2
               MaxChar         =   2
            End
            Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON TMS_FLATBUTTON5 
               Height          =   255
               Left            =   -74940
               TabIndex        =   149
               Top             =   360
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   450
               Caption         =   "<-- Indietro"
               ButtonBorderColor=   0
               ButtonForeColor =   16777215
               ButtonHilightBorderColor=   9473677
            End
         End
      End
      Begin PRJFW_FRAME.TMS_FRAME FrameCliente 
         Height          =   2445
         Left            =   30
         Top             =   420
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   4313
         Begin PRJFW_RICHTEXTBOX.TmsRichTextBox TXT_GB06_NOMEOFFERTA 
            Height          =   300
            Left            =   4920
            TabIndex        =   225
            Top             =   90
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   529
            MaxChar         =   150
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            DBField         =   "GB07_DESCART"
            NumRighe        =   5
            IsExpand        =   -1  'True
            MaxWidth        =   44
            IsInLingua      =   0   'False
            LinguaEntitaDes =   0
            LinguaIDProvenienzaExt=   0
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   930
            Picture         =   "FRMGB_OFFERTA.frx":2D2C
            ScaleHeight     =   285
            ScaleWidth      =   285
            TabIndex        =   4
            Top             =   750
            Width           =   285
         End
         Begin PRJFW_EDITNUMM.TxtEditNumM TXT_GB06_CLIFOR_CG44 
            Height          =   300
            Left            =   1080
            TabIndex        =   5
            Top             =   750
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   529
            IsLookup        =   -1  'True
            IsCalculator    =   0   'False
            MaxChar         =   15
            IsDbField       =   0   'False
            Caption         =   "Codice cliente"
            Object.Tag             =   "Codice cliente"
            MaxWidth        =   6
            IsDecode        =   -1  'True
         End
         Begin PRJFW_EDITM.TXT_EDITM TXT_GB06_CODPAG_CG62 
            Height          =   300
            Left            =   1080
            TabIndex        =   6
            ToolTipText     =   "Articolo"
            Top             =   2070
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   529
            IsLookup        =   -1  'True
            DisplayFormat   =   "Maiuscolo"
            MaxChar         =   6
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            DBField         =   "GB07_CODART_MG66"
            IsDecode        =   -1  'True
            Caption         =   "Articolo"
            NumRighe        =   0
            Object.Tag             =   "Articolo"
            MaxWidth        =   6
            IsInLingua      =   0   'False
            LinguaEntitaDes =   0
            LinguaIDProvenienzaExt=   ""
         End
         Begin PRJFW_EDITM.TXT_EDITM TXT_GB06_ID 
            Height          =   300
            Left            =   1080
            TabIndex        =   7
            ToolTipText     =   "Articolo"
            Top             =   90
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   529
            IsLookup        =   -1  'True
            DisplayFormat   =   "Maiuscolo"
            MaxChar         =   6
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            DBField         =   "GB07_CODART_MG66"
            IsDecode        =   -1  'True
            Caption         =   "Articolo"
            NumRighe        =   0
            Object.Tag             =   "Articolo"
            MaxWidth        =   6
            IsInLingua      =   0   'False
            LinguaEntitaDes =   0
            LinguaIDProvenienzaExt=   ""
         End
         Begin PRJFW_EDITM.TXT_EDITM TXT_CG16_RAGSOANAG 
            Height          =   300
            Left            =   2910
            TabIndex        =   8
            ToolTipText     =   "Articolo"
            Top             =   750
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   529
            IsLookup        =   0   'False
            DisplayFormat   =   "Maiuscolo"
            Enabled         =   0   'False
            MaxChar         =   25
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            DBField         =   "GB07_CODART_MG66"
            Caption         =   "Articolo"
            NumRighe        =   0
            Object.Tag             =   "Articolo"
            MaxWidth        =   22
            IsInLingua      =   0   'False
            LinguaEntitaDes =   0
            LinguaIDProvenienzaExt=   ""
         End
         Begin PRJFW_EDITM.TXT_EDITM TXT_CONTATTO 
            Height          =   300
            Left            =   1080
            TabIndex        =   9
            ToolTipText     =   "Articolo"
            Top             =   1740
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   529
            IsLookup        =   -1  'True
            DisplayFormat   =   "Maiuscolo"
            MaxChar         =   6
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            DBField         =   "GB07_CODART_MG66"
            IsDecode        =   -1  'True
            Caption         =   "Articolo"
            NumRighe        =   0
            Object.Tag             =   "Articolo"
            MaxWidth        =   6
            IsInLingua      =   0   'False
            LinguaEntitaDes =   0
            LinguaIDProvenienzaExt=   ""
         End
         Begin PRJFW_EDITM.TXT_EDITM TXT_GB06_CODDESTIN_MG22 
            Height          =   300
            Left            =   5970
            TabIndex        =   208
            Top             =   750
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   529
            IsGestione      =   -1  'True
            IsLookup        =   -1  'True
            DisplayFormat   =   "Maiuscolo"
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            DBField         =   "GB06_CODDESTIN_MG22"
            IsDecode        =   -1  'True
            NumRighe        =   0
            MaxWidth        =   6
            IsInLingua      =   0   'False
            LinguaEntitaDes =   0
            LinguaIDProvenienzaExt=   ""
         End
         Begin PRJFW_EDIT.TxtEdit TXT_GB06_NUMDOC 
            Height          =   300
            Left            =   2370
            TabIndex        =   15
            Top             =   90
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            MaxChar         =   50
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            MaxWidth        =   13
         End
         Begin PRJFW_EDIT.TxtEdit TXT_GB06_NREV 
            Height          =   300
            Left            =   4080
            TabIndex        =   18
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            Enabled         =   0   'False
            MaxChar         =   2
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            MaxWidth        =   2
         End
         Begin PRJFW_EDIT.TxtEdit TXT_GB06_NVERS 
            Height          =   300
            Left            =   4410
            TabIndex        =   26
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            Enabled         =   0   'False
            MaxChar         =   2
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            MaxWidth        =   2
         End
         Begin PRJFW_CHECKBOX.TMS_CHECKBOX TXT_FLPOSA 
            Height          =   300
            Left            =   8070
            TabIndex        =   224
            Top             =   2070
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   529
            Default         =   "1"
            DBField         =   "GB07_FLPOSA"
            Caption         =   "Forza Esclusione Posa"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   51
            Left            =   6180
            TabIndex        =   214
            TabStop         =   0   'False
            Top             =   480
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            Caption         =   "Destinazione"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TXT_MG22_DESTCITTA 
            Height          =   300
            Left            =   6990
            TabIndex        =   213
            TabStop         =   0   'False
            Tag             =   "Descr. fornitore ENECO"
            Top             =   1410
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   529
            Caption         =   ""
            Border          =   -1  'True
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TXT_MG22_DESTPROV 
            Height          =   300
            Left            =   9870
            TabIndex        =   212
            TabStop         =   0   'False
            Tag             =   "Descr. fornitore ENECO"
            Top             =   1410
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   529
            Caption         =   ""
            Border          =   -1  'True
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TXT_MG22_DESTCAP 
            Height          =   300
            Left            =   6150
            TabIndex        =   211
            TabStop         =   0   'False
            Tag             =   "Descr. fornitore ENECO"
            Top             =   1410
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   529
            Caption         =   ""
            Border          =   -1  'True
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TXT_MG22_DESTIND 
            Height          =   300
            Left            =   6150
            TabIndex        =   210
            TabStop         =   0   'False
            Tag             =   "Descr. fornitore ENECO"
            Top             =   1080
            Width           =   4365
            _ExtentX        =   7699
            _ExtentY        =   529
            Caption         =   ""
            Border          =   -1  'True
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TXT_MG22_DESTRAGSOC 
            Height          =   300
            Left            =   7440
            TabIndex        =   209
            TabStop         =   0   'False
            Tag             =   "Descr. fornitore ENECO"
            Top             =   750
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   529
            Caption         =   ""
            Border          =   -1  'True
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TXT_CG62_DESCPAG 
            Height          =   300
            Left            =   2790
            TabIndex        =   20
            TabStop         =   0   'False
            Tag             =   "Descr. fornitore ENECO"
            Top             =   2070
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   529
            Caption         =   ""
            Border          =   -1  'True
         End
         Begin PRJFW_EDIT.TxtEdit TXT_CODPAG 
            Height          =   300
            Left            =   3810
            TabIndex        =   27
            Top             =   2070
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            Enabled         =   0   'False
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TXT_CG16_RAGSOANAG1 
            Height          =   300
            Left            =   3360
            TabIndex        =   25
            TabStop         =   0   'False
            Tag             =   "Descr. fornitore ENECO"
            Top             =   750
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   529
            Caption         =   ""
            Border          =   -1  'True
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TXT_CG16_INDIRIZZO 
            Height          =   300
            Left            =   1260
            TabIndex        =   24
            TabStop         =   0   'False
            Tag             =   "Descr. fornitore ENECO"
            Top             =   1080
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   529
            Caption         =   ""
            Border          =   -1  'True
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TXT_CG16_CAP 
            Height          =   300
            Left            =   1260
            TabIndex        =   23
            TabStop         =   0   'False
            Tag             =   "Descr. fornitore ENECO"
            Top             =   1410
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   529
            Caption         =   ""
            Border          =   -1  'True
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TXT_CG16_CITTA 
            Height          =   300
            Left            =   2250
            TabIndex        =   22
            TabStop         =   0   'False
            Tag             =   "Descr. fornitore ENECO"
            Top             =   1410
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   529
            Caption         =   ""
            Border          =   -1  'True
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TXT_CG16_PROV 
            Height          =   300
            Left            =   5220
            TabIndex        =   21
            TabStop         =   0   'False
            Tag             =   "Descr. fornitore ENECO"
            Top             =   1410
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   529
            Caption         =   ""
            Border          =   -1  'True
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_MG19_LISTMAG 
            Height          =   300
            Left            =   5820
            TabIndex        =   19
            Top             =   2070
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   529
            Enabled         =   0   'False
            IsDbField       =   0   'False
            Caption         =   "Listini"
            Object.Tag             =   "Listini"
            MaxWidth        =   3
            MaxChar         =   14
         End
         Begin PRJFW_EDITDATE.TxtEditDate TXT_GB06_DTDOC 
            Height          =   300
            Left            =   1080
            TabIndex        =   17
            Top             =   420
            Width           =   1680
            _ExtentX        =   2566
            _ExtentY        =   529
            IsCalendario    =   0   'False
            IsDbField       =   0   'False
         End
         Begin PRJFW_EDIT.TxtEdit TXT_GB06_ALLACA 
            Height          =   300
            Left            =   2340
            TabIndex        =   16
            Top             =   1740
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   529
            MaxChar         =   100
            Numerico        =   0   'False
            Carattere       =   0   'False
            DBField         =   "GB06_ALLACA"
            MaxWidth        =   66
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   22
            Left            =   60
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   90
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            Caption         =   "N°/Rev/Vers."
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   25
            Left            =   60
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   450
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            Caption         =   "Data Offerta"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   26
            Left            =   60
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   780
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            Caption         =   "Cliente"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   27
            Left            =   60
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1740
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            Caption         =   "Alla C.A."
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
            Height          =   300
            Index           =   28
            Left            =   60
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   2040
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            Caption         =   "Cond. Pag."
         End
      End
      Begin PRJFW_EDITDATE.TxtEditDate TXT_ANNOINSERIMENTO 
         Height          =   300
         Left            =   -73440
         TabIndex        =   219
         Top             =   3240
         Width           =   1680
         _ExtentX        =   2566
         _ExtentY        =   529
         IsCalendario    =   0   'False
         IsDbField       =   0   'False
      End
      Begin VB.Label Label5 
         Caption         =   "Anno Listino:"
         Height          =   405
         Left            =   -74910
         TabIndex        =   218
         Top             =   3270
         Width           =   2595
      End
      Begin VB.Label Label6 
         Caption         =   "Ultimo Agg.05/05/2017 h. 12.15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74940
         TabIndex        =   215
         Top             =   2970
         Width           =   4095
      End
      Begin VB.Label Label4 
         Caption         =   "asdasdsadsa"
         Height          =   255
         Left            =   90
         TabIndex        =   193
         Top             =   9060
         Width           =   3735
      End
      Begin VB.Label lblTipoDocumento 
         Caption         =   "DOCUMENTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   -30
         TabIndex        =   3
         Top             =   0
         Width           =   17805
      End
   End
   Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
      Height          =   300
      Index           =   43
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      Caption         =   "Tipo Operazione"
   End
   Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
      Height          =   300
      Index           =   10
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   529
      Caption         =   "10 - COMPLETATA"
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   1350
      Picture         =   "FRMGB_OFFERTA.frx":32B6
      Top             =   7770
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "FRMGB_OFFERTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
 

' Reference to standard interface of class
Public ActiveInterface                As Cinterface
Public QtaRead As String
'Reference to extended interface of class
Public ActiveClass                    As CLSGB_OFFERTA
Public CurrentGridPosition             As Variant

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Generazione /cancellazione documenti
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private ClsRegDocBatch                    As MGBO_REGDOCUMENTI.CLSMG_REGDOCBATCH
'Public WithEvents Gcls_RegDoc             As MGBO_REGDOCUMENTI.CLSMG_REGDOC
Private ClasseInternaRegDoc               As CLSBO_REGDOC
Public TotRigheFile                       As Integer
Public TotRigheFileScartati                       As Integer
Public isLettoreMomoria                   As Boolean
Dim ContaRighe                            As Double
Public NumDocGenerato                     As Double
Public NumRegGenerato                     As String
Dim CodiceDocumento                       As String
Public isBarcode As Boolean

Public DEP As String
Public DEPCOLL As String

'Connection
Private Gstr_Connect                 As String
Private Gcon_Connect                 As ADODB.Connection
Private Gcls_Connect                 As CLSFW_SetConnect


'Virtual Frame Contratti
Public WithEvents FME_BANCO          As CLSFW_VIRTUALFRAME
Attribute FME_BANCO.VB_VarHelpID = -1
Public WithEvents FME_GRUPPI         As CLSFW_VIRTUALFRAME
Attribute FME_GRUPPI.VB_VarHelpID = -1
Public WithEvents FME_NOTE         As CLSFW_VIRTUALFRAME
Attribute FME_NOTE.VB_VarHelpID = -1


Private Pobj_SpecialLookUp              As COUO_QUERYANAGPDC.CLSCO_LOOKUP


'Recordset
Private Gcls_RecordS                 As CLSFW_Recordset
Private rstCorpoBanco                As ADODB.Recordset
Private rstScriveDoc                 As ADODB.Recordset

' Error management class
Private Gcls_Log                     As CLSFW_SrvLog

'Variabili globali
Private CodiceDitta                 As Variant
Private strSQL                      As String
Private ProgInsertTestata           As String
Private IDGB06                      As Double
Private TipoDocumento               As Integer
Private Pstr_old_codart             As String
Private Pbol_ReturnPressed          As Boolean
Private TipoDocumentoRecuperato     As Double
Private Errore                      As String

Private Pbol_BloccoFido             As Boolean
Private ValoreFido                  As Double
Private ResiduoFido                 As Double
Private ClienteNoFido               As Boolean
Private ClienteBloccato             As Boolean
Private Pbol_Generazione            As Boolean
Private Gcls_RecordPadre             As CLSFW_Recordset


'Classi per il controllo fido
'Dim objFido                         As CLS_CONTROLLOFIDO
Public Cls_ControlloRischio         As MGBO_CALCRISCHIO.CLSMG_CALCRISCHIO
Public Pcls_ProgArt                 As MGBO_PROGMAG.CLSMG_PROGART
Public Cls_GetProgMag               As MGBO_PROGMAG.CLSMG_GETPROGMAG
'Variabili per file parametri
Dim GB05_PWDFIDO       As String
Dim GB05_PWDFIDCARD    As String
Dim GB05_PWDSCRIGA     As String
Dim GB05_PWDSCPIEDE    As String
Dim GB05_MAXSCRIGA     As Double
Dim GB05_MAXSCPIEDE    As Double
Dim GB05_ANNOINSERIMENTO      As String


'Variabili per blocco articoli
Dim CODICE_DOCUMENTO    As String
Dim TIPO_DOCUMENTO      As Double
Dim SOTTOTIPO_DOCUMENTO As Double
Private Gcls_CommandBloccoStatiArt          As ADODB.Command


'Variabili per file
Dim strPathFile   As String
Dim strNomeFile   As String
Dim strNomeFileRen   As String
Dim numfile       As Integer
Dim numfileErr    As Integer
Private rsScontrino As ADODB.Recordset

'
' Gestione degli errori
'
Private lng_Stato                           As Long
Private str_StatoDesc                       As String
Private Gbol_DeadLock                       As Boolean

'Variabili per log
'Dim strNomeFile   As String
Dim decFlusso     As Integer
Dim decErrore     As Integer
Dim strNumReg     As String

'Classi standard
Public Cls_LookupMagazzino                  As MGBO_LOOKUPDECODE.CLSMG_LOOKUP
Public Cls_LookupCommon                     As COBO_LOOKUPDECODE.CLSCO_LOOKUP

Public Cls_DecodeMagazzino                  As MGBO_LOOKUPDECODE.CLSMG_DECODE
Public Cls_DecodeCommon                     As COBO_LOOKUPDECODE.CLSCO_DECODE

' Classi per i programmi di gestione
Public WithEvents Cls_ConnectMagazzino      As MGBO_LOOKUPDECODE.CLSMG_CONNECT
Attribute Cls_ConnectMagazzino.VB_VarHelpID = -1
Public WithEvents Cls_ConnectCommon         As COBO_LOOKUPDECODE.CLSCO_CONNECT
Attribute Cls_ConnectCommon.VB_VarHelpID = -1

Public Cls_LookupLotti                      As LTBO_LOOKUPDECODE.CLSLT_LOOKUP
Public Cls_DecodeLotti                      As LTBO_LOOKUPDECODE.CLSLT_DECODE
Public Cls_ConnectLotti                     As LTBO_LOOKUPDECODE.CLSLT_CONNECT

Private Cls_CalcPrezzi                      As MGBO_PRIORPRSC.CLSMG_LEGGOPRE

Public rstGRUPPI                As ADODB.Recordset
Public rstNOTE                As ADODB.Recordset
Public rstGridOrdini As ADODB.Recordset
Public rstGridMargine As ADODB.Recordset





' Semafor
Private FormIsActive                 As Boolean

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Public Event ErrorsOccurred(ErrorDetail As String, ShowMsgBox As Boolean, MessageStyle As VbMsgBoxStyle, MessageResult As VbMsgBoxResult)

Private Const VK_SHIFT = &H10

'Condizioni di pagamento
'501 contanti
'601 bancomat C01
'603 carta credito C02
'605 Assegni C03
'503 Contanti speciali CO4
'218 Bonifico bancario C05
'100 Rimessa diretta D03
'  da cliente
Private Const CODPAGCONTANTI = "C00"
Private Const CODPAGBANCOMAT = "C01"
Private Const CODPAGCARTACR = "C02"
Private Const CODPAGASSEGNI = "C03"
Private Const CODPAGCONTSPEC = "C04"
Private Const CODPAGBB = "C05"
Private Const CODPAGRD = "D03"
Dim CODPAGCLI As String
Dim ValContanti As Double

'Cliente bricomatt per scontrini
Private Const CODCLIBRICOMATT = 3
Private Const CODARTORDINE = ""

'Codice documento da generare
Private Const CODDOCCarico = "CL-ORDINE"
Private Const CODDOCCaricoTrasp = "CL-ORDINE"
Private Const CODDOCDDT = "CL-ORDINE"
Private Const CODDOCDDTTrasp = "CL-ORDINE"
Private Const CODDOCScontrini = "CL-ORDINE"
Private Const CODDOCScontriniTrasporti = "CL-ORDINE"
Private Const CODDOCNC = "CL-ORDINE"
Private Const CODDOCReso = "CL-ORDINE"
Private Const CODDOCDDTScarico = "CL-ORDINE"
Private Const DEPCENTRALE = "DM"
Private Const CODDOCOfferta = "CL-ORDINE"
Private Const CODDOCOrdine = "FO-ORDINE"
Private Const CODDOCOrdinePosa = "FO-INCPOSA"


'Variabili per file
'Private Const PARAM_DIRTEMP = "C:\TeamSystem Software\scambio\Temp"
'Private Const PARAM_DIRSCONTRINI = "C:\TeamSystem Software\scambio\Invio"
'Private Const PARAM_DIRCOPIA = "C:\TeamSystem Software\scambio\Invio"
'Private Const PARAM_NOMEFILE = "Scontrini.txt"

Private PARAM_DIRTEMP                         As String
Private PARAM_DIRSCONTRINI                    As String
Private PARAM_DIRCOPIA                        As String
Private PARAM_NOMEFILE                        As String
Private PARAM_EXESCONTR                        As String



'Report parametri
Private WithEvents PclsReport                 As FWBO_REPORT30.CLSFW_REPORTCD
Attribute PclsReport.VB_VarHelpID = -1
Private FiltroStampaChiusura                  As String
Private StampaSingola                         As Boolean







Private Sub cmd_Click()

End Sub

Public Function DirExists(ByVal Path As String) As Boolean
On Error Resume Next
Dim FileExists As Boolean
'Legge l'attributo e si assicura che si tratti di una directory
FileExists = GetAttr(Path) And vbDirectory
'Se avviene un errore la Function restituisce False
End Function

Private Sub CHK_FAM_AfterItem(Cancel As Boolean)
Call AnalisiMargini("MG66_FAM_MG53", "Famiglia")
End Sub

Private Sub CHK_GRUPPO_AfterItem(Cancel As Boolean)
Call AnalisiMargini("MG66_GRUPPO_MG55", "Gruppo")
End Sub

Private Sub CHK_SFAM_AfterItem(Cancel As Boolean)
Call AnalisiMargini("MG66_SFAM_MG54", "Sotto Famiglia")
End Sub

Private Sub CHK_SGRUPPO_AfterItem(Cancel As Boolean)
Call AnalisiMargini("MG66_SGRUPPO_MG56", "Sotto Gruppo")
End Sub





Private Sub CDM_RICPERCPROVVMASSI_Click()

Call CMD_SAVE_Click

If QGridDocumenti.DataSource Is Nothing Then
    Exit Sub
End If

If QGridDocumenti.DataSource.RecordCount = 0 Then
    Exit Sub
End If

QGridDocumenti.DataSource.MoveFirst
Do While Not QGridDocumenti.DataSource.EOF
      QGridDocumenti.DataSource("GB07_PERCPROVV").value = RecuperaProvvigione(QGridDocumenti.DataSource("GB07_PREZZO").value, QGridDocumenti.DataSource("GB07_IMPORTO").value, NVL(QGridDocumenti.DataSource("MG66_FAM_MG53").value, ""), NVL(QGridDocumenti.DataSource("MG66_SFAM_MG54").value, ""), TXT_GB06_AGENTE.Text, NVL(QGridDocumenti.DataSource("GB07_SC1").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC2").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC3").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC4").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC5").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC6").value, 0))
      QGridDocumenti.DataSource.MoveNext
Loop

Call CMD_SAVE_Click

End Sub

Private Sub CMD_AGGIORNA_Click()
Dim strUpdateFor As String
Dim nRighe As Integer
FME_BANCO.MoveFirst
nRighe = GetValFromQuery("SELECT COUNT(GB07_CHECK) AS tot From GB07_CORPODOC WHERE  GB07_ID_GB06 = " & IDGB06 & " AND GB07_CHECK = 1", 0, Gcon_Connect)

If MsgBox("Confermi la variazione del Fornitore " & TXT_CLIFOR_CG44.Text & " su " & CStr(nRighe) & " ?", vbYesNo) = vbYes Then

    strUpdateFor = "UPDATE GB07_CORPODOC SET GB07_TIPOCF_CG44 = 1 , GB07_CLIFOR_CG44 = " & TXT_CLIFOR_CG44.Text & _
    "WHERE  GB07_ID_GB06 = " & IDGB06 & " AND GB07_CHECK = 1"
Gcon_Connect.Execute strUpdateFor
Call ImpostaVirtualFrame
End If
End Sub

Private Sub CMD_ANALISIMARGINI_Click()
On Error GoTo ErrTrap
  Call CMD_SAVE_Click
  Set FRMGB_MARGINI.ActiveClass = ActiveClass
  Set FRMGB_MARGINI.ActiveInterface = ActiveInterface
  FRMGB_MARGINI.Caption = "Analisi Margini Offerta " & TXT_GB06_NOMEOFFERTA.Text
  FRMGB_MARGINI.IdOfferta = IDGB06
  FRMGB_MARGINI.ValPercTrasp = TXT_GB06_PERCTRASP.Text

  FRMGB_MARGINI.Show vbModal

  'QGRD_ORDINI.Refresh
  
  Me.Refresh

Exit_Handler:
  Exit Sub

ErrTrap:
  Select Case VisualizzaErrore("CMD_ANALISIMARGINI_Click")
    Case vbAbort
      GoTo Exit_Handler
    Case vbRetry
      Resume
    Case vbIgnore
      Resume Next
End Select
End Sub

Private Sub CMD_ANNULLA_Click()
If MsgBox("Confermi l'Annullamento dell'offerta " & TXT_GB06_NOMEOFFERTA.Text & " ?", vbYesNo) = vbNo Then
      Exit Sub
End If
 'Scrive dati in GB06
    strSQL = "UPDATE GB06_TESTADOC "
    strSQL = strSQL & " SET  GB06_STATODOC         = '08'"
    strSQL = strSQL & " WHERE GB06_ID   = " & IDGB06
    Gcon_Connect.Execute strSQL
    TXT_GB06_STATODOC.Text = "08"
Call CMD_SAVE_Click
End Sub

Private Sub CMD_CONFERMA_Click()
'scrivi note


Dim strInserisciNota As String
Dim strTesto As String
Dim TABLENAME As String
  
TABLENAME = "TMP_SIMULAZIONE_" & ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Codice

strTesto = "Variazione Offerta __ <Da completare>"
strInserisciNota = "INSERT INTO GB08_NOTEOFFERTA (GB08_DATA, GB08_TESTONOTA, GB08_ID_GB06, GB08_OPERATORE) "
strInserisciNota = strInserisciNota & " VALUES "
strInserisciNota = strInserisciNota & "          ('" & Month(Now()) & "/" & Day(Now()) & "/" & Year(Now()) & "', '" & strTesto & "', " & IDGB06 & ",'" & ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Codice & "' ) "
Gcon_Connect.Execute strInserisciNota

If MsgBox("Questa operazione renderà effettiva la simulazione " & vbCrLf & "creando una nuova versione dell'offerta. " & vbCrLf & " Confermi?", vbYesNo) = vbNo Then
      Exit Sub
End If

Call ConfermaSimulazione(TABLENAME)

TMS_SSTAB1.ActiveTab = 4

End Sub

Private Sub CMD_COSTO_Click()
Call RecuperaCosto
'Call CMD_REFRESH_Click
End Sub

Private Sub CMD_DELALL_Click()
Dim strDelete As String
Dim nRighe As Integer
FME_BANCO.MoveFirst
nRighe = GetValFromQuery("SELECT COUNT(GB07_CHECK) AS tot From GB07_CORPODOC WHERE  GB07_ID_GB06 = " & IDGB06 & " AND GB07_CHECK = 1", 0, Gcon_Connect)

If MsgBox("Confermi la cancellazione di " & CStr(nRighe) & " ?", vbYesNo) = vbYes Then

    strDelete = "DELETE GB07_CORPODOC " & _
    "WHERE  GB07_ID_GB06 = " & IDGB06 & " AND GB07_CHECK = 1"
Gcon_Connect.Execute strDelete
Call ImpostaVirtualFrame
End If
End Sub

Private Sub CMD_DUPLICA_ButtonClick()
Dim strNewVersione As String
Dim CurUserName As String
CurUserName = ActualComputerName & ActualUserName

If MsgBox("Confermi la Duplicazione dell'Offerta " & TXT_GB06_NOMEOFFERTA.Text & " Versione " & TXT_GB06_NVERS.Text & " Revisione " & TXT_GB06_NREV.Text & "?", vbYesNo) = vbNo Then
      Exit Sub
End If
TXT_GB06_CODCOMM.Text = ""
ProgInsertTestata = CurUserName & Year(Now) & String(2 - Len(Month(Now)), "0") & Month(Now) & String(2 - Len(Day(Now)), "0") & Day(Now) & String(2 - Len(Hour(Time)), "0") & Hour(Time) & String(2 - Len(Minute(Time)), "0") & Minute(Time) & String(2 - Len(Second(Time)), "0") & Second(Time) & Right(Format(Timer, "#0.00"), 2)





strNewVersione = " insert into GB06_TESTADOC"
strNewVersione = strNewVersione + "("
strNewVersione = strNewVersione + " GB06_PROG,"
strNewVersione = strNewVersione + " GB06_DITTA_CG18,"
strNewVersione = strNewVersione + " GB06_TIPODOC,"
strNewVersione = strNewVersione + " GB06_DATA,"
strNewVersione = strNewVersione + " GB06_GRUPPOCRE,"
strNewVersione = strNewVersione + " GB06_UTENTECRE,"
strNewVersione = strNewVersione + " GB06_TIPOCF_CG44,"
strNewVersione = strNewVersione + " GB06_CLIFOR_CG44,"
strNewVersione = strNewVersione + " GB06_BARCODE_LL01,"
strNewVersione = strNewVersione + " GB06_CODDESTIN_MG22,"
strNewVersione = strNewVersione + " GB06_CODPAG_CG62,"
strNewVersione = strNewVersione + " GB06_VETTORE_MG14,"
strNewVersione = strNewVersione + " GB06_NUMREG_CO99,"
strNewVersione = strNewVersione + " GB06_NUMDOC,"
strNewVersione = strNewVersione + " GB06_NREV,"
strNewVersione = strNewVersione + " GB06_NVERS,"
strNewVersione = strNewVersione + " GB06_DTDOC,"
strNewVersione = strNewVersione + " GB06_CODCOMM,"
strNewVersione = strNewVersione + " GB06_DTCHIUSURA,"
strNewVersione = strNewVersione + " GB06_PERCCHIUSURA,"
strNewVersione = strNewVersione + " GB06_RESPONSABILE,"
strNewVersione = strNewVersione + " GB06_PERCRIBGARA,"
strNewVersione = strNewVersione + " GB06_TIPOOFFERTA,"

strNewVersione = strNewVersione + " GB06_AGENTE ,"
strNewVersione = strNewVersione + " GB06_TIPOAREA,"
strNewVersione = strNewVersione + " GB06_STATODOC,"
strNewVersione = strNewVersione + " GB06_BUDGET,"
strNewVersione = strNewVersione + " GB06_FORECAST,"
strNewVersione = strNewVersione + " GB06_CONSUNTIVO,"
strNewVersione = strNewVersione + " GB06_NOMEOFFERTA,"
strNewVersione = strNewVersione + " GB06_DTULTMOD,"
strNewVersione = strNewVersione + " GB06_CIG,"
strNewVersione = strNewVersione + " GB06_CUP,"
strNewVersione = strNewVersione + " GB06_ALLACA,"
strNewVersione = strNewVersione + " GB06_TEXT1,"
strNewVersione = strNewVersione + " GB06_TEXT2,"
strNewVersione = strNewVersione + " GB06_TEXT3,"
strNewVersione = strNewVersione + " GB06_TEXT4,"
strNewVersione = strNewVersione + " GB06_TEXT5,"
strNewVersione = strNewVersione + " GB06_TEXT6"
strNewVersione = strNewVersione + ")"
strNewVersione = strNewVersione + " select"
strNewVersione = strNewVersione + " '" & ProgInsertTestata & "', GB06_DITTA_CG18, GB06_TIPODOC, GB06_DATA, GB06_GRUPPOCRE, GB06_UTENTECRE, GB06_TIPOCF_CG44,"
strNewVersione = strNewVersione + "                         GB06_CLIFOR_CG44, GB06_BARCODE_LL01, GB06_CODDESTIN_MG22, GB06_CODPAG_CG62, GB06_VETTORE_MG14, GB06_NUMREG_CO99, GB06_NUMDOC,"
strNewVersione = strNewVersione + "                         1 , 1, GB06_DTDOC, null, GB06_DTCHIUSURA, GB06_PERCCHIUSURA,GB06_RESPONSABILE, GB06_PERCRIBGARA, GB06_TIPOOFFERTA,"
strNewVersione = strNewVersione + "                          GB06_AGENTE , GB06_TIPOAREA, '00', GB06_BUDGET, GB06_FORECAST, GB06_CONSUNTIVO, GB06_NOMEOFFERTA, { fn NOW() }, GB06_CIG, GB06_CUP, GB06_ALLACA, GB06_TEXT1, GB06_TEXT2, GB06_TEXT3, GB06_TEXT4, GB06_TEXT5, GB06_TEXT6"
strNewVersione = strNewVersione + " from GB06_TESTADOC where GB06_ID = " & TXT_GB06_ID.Text
strNewVersione = strNewVersione + " group by GB06_PROG, GB06_DITTA_CG18, GB06_TIPODOC, GB06_DATA, GB06_GRUPPOCRE, GB06_UTENTECRE, GB06_TIPOCF_CG44,"
strNewVersione = strNewVersione + "                         GB06_CLIFOR_CG44, GB06_BARCODE_LL01, GB06_CODDESTIN_MG22, GB06_CODPAG_CG62, GB06_VETTORE_MG14, GB06_NUMREG_CO99, GB06_NUMDOC,"
strNewVersione = strNewVersione + "                          GB06_NVERS, GB06_DTDOC, GB06_CODCOMM, GB06_DTCHIUSURA, GB06_PERCCHIUSURA,GB06_RESPONSABILE, GB06_PERCRIBGARA, GB06_TIPOOFFERTA,"
strNewVersione = strNewVersione + "                          GB06_AGENTE , GB06_TIPOAREA, GB06_STATODOC, GB06_BUDGET, GB06_FORECAST, GB06_CONSUNTIVO, GB06_NOMEOFFERTA, GB06_DTULTMOD, GB06_CIG, GB06_CUP, GB06_ALLACA, GB06_TEXT1, GB06_TEXT2, GB06_TEXT3, GB06_TEXT4, GB06_TEXT5, GB06_TEXT6"
Gcon_Connect.Execute strNewVersione

strNewVersione = ""
strNewVersione = strNewVersione + "insert into GB07_CORPODOC "
strNewVersione = strNewVersione + " ( "
strNewVersione = strNewVersione + " GB07_ID_GB06, GB07_CODART_MG66, GB07_OPZIONE_MG5E , "
strNewVersione = strNewVersione + " GB07_QTA, GB07_PREZZO, GB07_SCCORPO, GB07_SC5, "
strNewVersione = strNewVersione + " GB07_SC4, GB07_SC3, GB07_SC2, GB07_SC1, "
strNewVersione = strNewVersione + " GB07_SC6, GB07_SCPIEDE, GB07_IMPORTO, GB07_DESCART, "
strNewVersione = strNewVersione + " GB07_MAGGIORAZIONE, GB07_SEQ, GB07_RAG, "
strNewVersione = strNewVersione + " GB07_ALT , GB07_FLPOSA, GB07_IMG, GB07_PATHIMG, GB07_TIPOCF_CG44, GB07_CLIFOR_CG44, GB07_COSTO, GB07_IMPPROVV,GB07_PERCPROVV "
strNewVersione = strNewVersione + " ) "
strNewVersione = strNewVersione + " select "
strNewVersione = strNewVersione + " (select max(GB06_ID) as ID from GB06_TESTADOC where GB06_PROG like '" & CurUserName & "%'), GB07_CODART_MG66, GB07_OPZIONE_MG5E, "
strNewVersione = strNewVersione + " GB07_QTA, GB07_PREZZO, GB07_SCCORPO, GB07_SC5, "
strNewVersione = strNewVersione + " GB07_SC4, GB07_SC3, GB07_SC2, GB07_SC1, "
strNewVersione = strNewVersione + " GB07_SC6, GB07_SCPIEDE, GB07_IMPORTO, GB07_DESCART, "
strNewVersione = strNewVersione + " GB07_MAGGIORAZIONE, GB07_SEQ, GB07_RAG, "
strNewVersione = strNewVersione + " GB07_ALT , GB07_FLPOSA, GB07_IMG, GB07_PATHIMG, GB07_TIPOCF_CG44, GB07_CLIFOR_CG44,  GB07_COSTO, GB07_IMPPROVV,GB07_PERCPROVV  "
strNewVersione = strNewVersione + " from GB07_CORPODOC where gb07_id_gb06 = " & TXT_GB06_ID.Text
Gcon_Connect.Execute strNewVersione

If MsgBox("Vuoi Gestire la nuova Offerta ?", vbYesNo) = vbYes Then
      TXT_GB06_ID.Text = GetValFromQuery("select max(GB06_ID) as ID from GB06_TESTADOC where GB06_PROG like '" & CurUserName & "%'", 0, Gcon_Connect)
      'Call TXT_GB06_ID_StartDecode(
      Select Case ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Username
        Case "ALESSANDRA"
          TXT_GB06_NUMDOC.Text = GetNumeroOfferta("AC")
        Case "BARBARA"
          TXT_GB06_NUMDOC.Text = GetNumeroOfferta("BP")
        Case "LAURA"
          TXT_GB06_NUMDOC.Text = GetNumeroOfferta("LD")
        Case "PIERO"
          TXT_GB06_NUMDOC.Text = GetNumeroOfferta("PB")
        Case "STEFANO"
          TXT_GB06_NUMDOC.Text = GetNumeroOfferta("SM")
        Case "SUSANNA"
          TXT_GB06_NUMDOC.Text = GetNumeroOfferta("SS")
        Case "VALERIA"
          TXT_GB06_NUMDOC.Text = GetNumeroOfferta("VV")
        Case "FEDRICA"
          TXT_GB06_NUMDOC.Text = GetNumeroOfferta("FB")
        Case Else
          TXT_GB06_NUMDOC.Text = GetNumeroOfferta("AB")
        
        End Select
      Call CMD_SAVE_Click
      TXT_GB06_ID.SetFocus
      
End If

End Sub

Private Sub CMD_ELIMINA_ButtonClick()
If MsgBox("Confermi l'eliminazione dell'offerta " & TXT_GB06_NOMEOFFERTA.Text & " ?", vbYesNo) = vbNo Then
      Exit Sub
End If
Gcon_Connect.Execute "delete GB07_CORPODOC where GB07_ID_GB06 = " & TXT_GB06_ID.Text
Gcon_Connect.Execute "delete GB06_TESTADOC  where GB06_id = " & TXT_GB06_ID.Text
Call cmdNuovoDoc_ButtonClick
End Sub

Private Sub CMD_GENERAOFFERTA_Click()
    Dim strFornitori
    Dim strCommessa As String
    Dim orsOrdiniF As ADODB.Recordset
   
    Dim contOrdF As Integer
    
    
    If NVL(TXT_GB06_CODCOMM.Text, "") = "" Then
        If MsgBox("Vuoi Creare una nuova Commessa?", vbYesNo) = vbYes Then
             'creazione nuova commessa
             
             strCommessa = " SELECT        isnull(MAX(DO11_NUMDOC),0) + 1 AS NEXTNUM " & _
                           " From DO11_DOCTESTATA " & _
                           " WHERE        (DO11_DITTA_CG18 = " & CodiceDitta & ") " & _
                           " AND (DO11_DOCUM_MG36 = '" & CODDOCOfferta & "') " & _
                           " AND (DO11_ANNODOC = " & Year(TXT_ANNOINSERIMENTO.Text) & ") AND (DO11_SEZDOC = '00') "
             
             TXT_GB06_CODCOMM.Text = GetCommessa(GetValFromQuery(strCommessa, 0, Gcon_Connect))
        End If
    
    End If
    
    
    
    Call CMD_SAVE_Click
    
    'controlli per creazione documenti
    
    If NVL(TXT_GB06_CODCOMM.Text, "") = "" Then
        If MsgBox("Commessa Mancate vuoi generare i documenti?", vbYesNo) = vbNo Then
            TXT_GB06_CODCOMM.SetFocus
            Exit Sub
        End If
    End If
    
    If NVL(TXT_GB06_NOMEOFFERTA.Text, "") = "" Then
        If MsgBox("Nome Offerta Mancante vuoi generare i documenti?", vbYesNo) = vbNo Then
            TXT_GB06_NOMEOFFERTA.SetFocus
            Exit Sub
        End If
    End If
    
    If NVL(TXT_GB06_BUDGET.Text, "") = "" Then
        If MsgBox("Budget Mancante vuoi generare i documenti?", vbYesNo) = vbNo Then
            TXT_GB06_BUDGET.SetFocus
            Exit Sub
        End If
    End If
    
    If NVL(TXT_GB06_CONSUNTIVO.Text, "") = "" Then
        If MsgBox("Totale Offerta Mancante vuoi generare i documenti?", vbYesNo) = vbNo Then
            TXT_GB06_CONSUNTIVO.SetFocus
            Exit Sub
        End If
    End If
    
    If NVL(TXT_GB06_CIG.Text, "") = "" Then
        If MsgBox("Cig Mancante vuoi generare i documenti?", vbYesNo) = vbNo Then
            TXT_GB06_CONSUNTIVO.SetFocus
            Exit Sub
        End If
    End If
    
    If NVL(TXT_GB06_AGENTE.Text, "") = "" Then
        If MsgBox("Agente Mancante vuoi generare i documenti?", vbYesNo) = vbNo Then
            TXT_GB06_CONSUNTIVO.SetFocus
            Exit Sub
        End If
    End If
    
     If NVL(TXT_GB06_RESPONSABILE.Text, "") = "" Then
        If MsgBox("Responsabile Mancante vuoi generare i documenti?", vbYesNo) = vbNo Then
            TXT_GB06_CONSUNTIVO.SetFocus
            Exit Sub
        End If
    End If
    
    ', ,
    
    Call OrdineCliente
   ''' Exit Sub
    'MsgBox "Ordine CLiente"
    

'''''''''    strFornitori = "    SELECT DISTINCT GB07_CORPODOC.GB07_TIPOCF_CG44, GB07_CORPODOC.GB07_CLIFOR_CG44, GB07_CORPODOC.GB07_ID_GB06, CG44_CLIFOR.CG44_DITTA_CG18, CG16_ANAGGEN.CG16_RAGSOANAG "
'''''''''    strFornitori = strFornitori & "    FROM            GB07_CORPODOC INNER JOIN"
'''''''''    strFornitori = strFornitori & "                             CG44_CLIFOR ON GB07_CORPODOC.GB07_TIPOCF_CG44 = CG44_CLIFOR.CG44_TIPOCF AND GB07_CORPODOC.GB07_CLIFOR_CG44 = CG44_CLIFOR.CG44_CLIFOR INNER JOIN"
'''''''''    strFornitori = strFornitori & "                             CG16_ANAGGEN ON CG44_CLIFOR.CG44_CODICE_CG16 = CG16_ANAGGEN.CG16_CODICE"
'''''''''    strFornitori = strFornitori & "    Where (GB07_CORPODOC.GB07_ID_GB06 = " & IDGB06 & ") And (CG44_CLIFOR.CG44_DITTA_CG18 = " & CodiceDitta & ")"
''''''''    strFornitori = " SELECT DISTINCT"
''''''''    strFornitori = strFornitori & "                     GB07_CORPODOC.GB07_TIPOCF_CG44, GB07_CORPODOC.GB07_CLIFOR_CG44, GB07_CORPODOC.GB07_ID_GB06, CG44_CLIFOR.CG44_DITTA_CG18, CG16_ANAGGEN.CG16_RAGSOANAG,"
''''''''    strFornitori = strFornitori & "                     MG66_ANAGRART.MG66_DITTA_CG18 , MG66_ANAGRART.MG66_GRUSTAT3_MG76"
''''''''    strFornitori = strFornitori & " FROM            GB07_CORPODOC INNER JOIN"
''''''''    strFornitori = strFornitori & "                     CG44_CLIFOR ON GB07_CORPODOC.GB07_TIPOCF_CG44 = CG44_CLIFOR.CG44_TIPOCF AND GB07_CORPODOC.GB07_CLIFOR_CG44 = CG44_CLIFOR.CG44_CLIFOR INNER JOIN"
''''''''    strFornitori = strFornitori & "                     CG16_ANAGGEN ON CG44_CLIFOR.CG44_CODICE_CG16 = CG16_ANAGGEN.CG16_CODICE LEFT OUTER JOIN"
''''''''    strFornitori = strFornitori & "                     MG66_ANAGRART ON GB07_CORPODOC.GB07_CODART_MG66 = MG66_ANAGRART.MG66_CODART"
''''''''    strFornitori = strFornitori & "    Where (GB07_CORPODOC.GB07_ID_GB06 = " & IDGB06 & ") And (CG44_CLIFOR.CG44_DITTA_CG18 = " & CodiceDitta & ") AND (MG66_ANAGRART.MG66_DITTA_CG18 = " & CodiceDitta & ") "
''''''''
    strFornitori = " SELECT DISTINCT "
    strFornitori = strFornitori & " GB07_CORPODOC.GB07_TIPOCF_CG44,"
    strFornitori = strFornitori & " GB07_CORPODOC.GB07_CLIFOR_CG44,"
    strFornitori = strFornitori & " GB07_CORPODOC.GB07_ID_GB06,"
    strFornitori = strFornitori & " CG44_CLIFOR.CG44_DITTA_CG18,"
    strFornitori = strFornitori & " CG16_ANAGGEN.CG16_RAGSOANAG,"
    strFornitori = strFornitori & " MG66_ANAGRART.MG66_DITTA_CG18,"
    strFornitori = strFornitori & " (CASE MG66_ANAGRART.MG66_GRUSTAT3_MG76"
    strFornitori = strFornitori & "  WHEN 'TRA' THEN 'GIO'"
    strFornitori = strFornitori & "  when 'LAV' then 'POS'"
'    strFornitori = strFornitori & "  when 'MAN' then 'POS'"
'    strFornitori = strFornitori & "  when 'LAV' then 'POS'"
    strFornitori = strFornitori & "  Else MG66_ANAGRART.MG66_GRUSTAT3_MG76"
    strFornitori = strFornitori & " END) AS MG66_GRUSTAT3_MG76"
    strFornitori = strFornitori & " From GB07_CORPODOC"
    strFornitori = strFornitori & " INNER JOIN CG44_CLIFOR"
    strFornitori = strFornitori & "  ON GB07_CORPODOC.GB07_TIPOCF_CG44 = CG44_CLIFOR.CG44_TIPOCF"
    strFornitori = strFornitori & "  AND GB07_CORPODOC.GB07_CLIFOR_CG44 = CG44_CLIFOR.CG44_CLIFOR"
    strFornitori = strFornitori & " INNER JOIN CG16_ANAGGEN"
    strFornitori = strFornitori & "  ON CG44_CLIFOR.CG44_CODICE_CG16 = CG16_ANAGGEN.CG16_CODICE"
    strFornitori = strFornitori & " LEFT OUTER JOIN MG66_ANAGRART"
    strFornitori = strFornitori & "  ON GB07_CORPODOC.GB07_CODART_MG66 = MG66_ANAGRART.MG66_CODART"
    strFornitori = strFornitori & " Where (GB07_CORPODOC.GB07_ID_GB06 = " & IDGB06 & ")"
    strFornitori = strFornitori & " AND (CG44_CLIFOR.CG44_DITTA_CG18 = " & CodiceDitta & ")"
    
    
    contOrdF = 0
    
    Set orsOrdiniF = New ADODB.Recordset

    With orsOrdiniF
    Set .ActiveConnection = Gcon_Connect
        .Source = strFornitori
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
    
    Do While Not orsOrdiniF.EOF
       contOrdF = contOrdF + 1
       If Trim(orsOrdiniF("MG66_GRUSTAT3_MG76")) = "POS" Then
            'CODDOCOrdine = CODDOCOrdinePosa
             Call OrdineFornitore(orsOrdiniF("GB07_CLIFOR_CG44"), orsOrdiniF("CG16_RAGSOANAG"), True)
       Else
             Call OrdineFornitore(orsOrdiniF("GB07_CLIFOR_CG44"), orsOrdiniF("CG16_RAGSOANAG"), False)
       End If
       orsOrdiniF.MoveNext
    Loop
    orsOrdiniF.Close
    
    Call LoadGrigliaOrdini
    TMS_SSTAB1.ActiveTab = 3
    
End Sub

Private Sub LoadGrigliaOrdini()
'carica griglia
Dim StringaSQL As String
     StringaSQL = "SELECT     * " & _
     "    FROM            GB01_ORDINI" & _
     "    Where (DO11_DITTA_CG18 = " & CodiceDitta & ") and GB09_ID_GB06 = " & IDGB06


    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridOrdini = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridOrdini
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With

    INITGRID_ORDINI
    TMS_GRIDDOC.BeginDataSourceSuspend

    Set TMS_GRIDDOC.DataSource = rstGridOrdini
    TMS_GRIDDOC.EndDataSourceSuspend
    TMS_GRIDDOC.Refresh
End Sub


Private Sub OrdineFornitore(CodFor As Double, Descrizione As String, isPosa As Boolean)
If MsgBox("Confermi la Generazione degli Ordini Fornitore " & Descrizione & " relativi all'offerta " & TXT_GB06_NOMEOFFERTA.Text & " Versione " & TXT_GB06_NVERS.Text & " Revisione " & TXT_GB06_NREV.Text & "?", vbYesNo) = vbNo Then
      Exit Sub
End If

  FrameDoc.Visible = True
  DoEvents
  Me.Enabled = False
  ContaRighe = 0
  
  Set rstScriveDoc = New ADODB.Recordset
  strSQL = "SELECT *  , GB07_CODART_MG66 AS DESART "
  strSQL = strSQL & "  FROM GB07_CORPODOC   "
  strSQL = strSQL & " WHERE GB07_ID_GB06 = " & IDGB06
  strSQL = strSQL & " AND GB07_TIPOCF_CG44 = 1 "
  strSQL = strSQL & " AND GB07_CLIFOR_CG44 = " & CodFor
  
  If isPosa Then
   ' strSQL = strSQL & " and GB07_CODART_MG66 like 'P\_%' escape '\'"
    strSQL = strSQL & " AND (GB07_CODART_MG66 LIKE 'MA%' or GB07_CODART_MG66 LIKE 'MA' or GB07_CODART_MG66 like 'Z\_%' escape '\' OR GB07_CODART_MG66 like 'P\_%' escape '\' OR GB07_CODART_MG66 LIKE 'LAVORI' OR GB07_CODART_MG66 LIKE 'ISP%' OR GB07_CODART_MG66 LIKE 'MTZ%')"
  Else
    strSQL = strSQL & " and GB07_CODART_MG66 not like 'P\_%' escape '\'"
  End If
  
  Set rstScriveDoc = Gcon_Connect.Execute(strSQL, , adCmdText)
  If Not rstScriveDoc.EOF Then

    Call ScriviDocumento(1, CodFor, isPosa)
  Else
    Pbol_Generazione = False
    MsgBox "Non ci sono righe da generare"
    Me.Enabled = True
    Exit Sub
  End If
  
  FrameDoc.Visible = False
  
  Set rstScriveDoc = Nothing
  
  FrameDoc.Visible = False
  DoEvents
  Me.Enabled = True
  
  If ContaRighe = 0 Then
    Pbol_Generazione = False
    MsgBox "Non ci sono righe da generare"
    Me.Enabled = True
    Exit Sub
  Else
  Call CMD_SAVE_Click
  
  
'  If MsgBox("Vuoi Aprire l'Ordine " & NumDocGenerato & "/" & Year(Now()) & " ?", vbYesNo) = vbNo Then
'      Exit Sub
'  End If
  'Call ApriDoc(CODDOCOfferta, NumRegGenerato)

  End If
End Sub



Private Sub OrdineCliente()
If MsgBox("Confermi la Trasformazione in Ordine dell'Offerta " & TXT_GB06_NOMEOFFERTA.Text & " Versione " & TXT_GB06_NVERS.Text & " Revisione " & TXT_GB06_NREV.Text & "?", vbYesNo) = vbNo Then
      Exit Sub
End If

  FrameDoc.Visible = True
  DoEvents
  Me.Enabled = False
  ContaRighe = 0
  Set rstScriveDoc = New ADODB.Recordset
  strSQL = "SELECT *  , GB07_CODART_MG66 AS DESART "
  strSQL = strSQL & "  FROM GB07_CORPODOC   "
  strSQL = strSQL & " WHERE GB07_ID_GB06 = " & IDGB06
  
  Set rstScriveDoc = Gcon_Connect.Execute(strSQL, , adCmdText)
  If Not rstScriveDoc.EOF Then

    Call ScriviDocumento(0, 0, False)
  Else
    Pbol_Generazione = False
    MsgBox "Non ci sono righe da generare"
    Me.Enabled = True
    Exit Sub
  End If
  
  FrameDoc.Visible = False
  
  Set rstScriveDoc = Nothing
  
  FrameDoc.Visible = False
  DoEvents
  Me.Enabled = True
  
  If ContaRighe = 0 Then
    Pbol_Generazione = False
    MsgBox "Non ci sono righe da generare"
    Me.Enabled = True
    Exit Sub
  Else
  Call CMD_SAVE_Click
  
  
'  If MsgBox("Vuoi Aprire l'Ordine " & NumDocGenerato & "/" & Year(Now()) & " ?", vbYesNo) = vbNo Then
'      Exit Sub
'  End If
  'Call ApriDoc(CODDOCOfferta, NumRegGenerato)

  End If
End Sub

Private Sub CMD_PRINT_Click()
On Error Resume Next
Dim SqlReport As String
Dim StampaArticoli As String
Dim StampaVarianti As String
Dim RsDocumento As ADODB.Recordset
Dim i As Long

Set PclsReport = Nothing
Set PclsReport = New FWBO_REPORT30.CLSFW_REPORTCD

PclsReport.UserObject = "GBUO_OFFERTA.CLSGB_OFFERTA" 'ActiveInterface.ClsGlobal.Gcls_VoceMenu.Classe
Set PclsReport.Connessione = Gcon_Connect
PclsReport.NomeServer = ActiveInterface.ClsGlobal.Gcls_GeConfig.ServerName
PclsReport.NomeDataBase = ActiveInterface.ClsGlobal.Gcls_GeConfig.DBName
PclsReport.NomeUtenteDb = ActiveInterface.ClsGlobal.Gcls_GeConfig.DBOwnerID
PclsReport.NomePasswordDb = ActiveInterface.ClsGlobal.Gcls_GeConfig.DBOwnerPwd
'PclsReport.TitoloReport = "1 - Stampa Offerta "
PclsReport.PercorsoReport = "C:\TeamSystem Software\Gamma Enterprise\Userfile\Rpt"
'
'PclsReport.NomeReport = "GBRP_OFFERTE.rpt"


PclsReport.ReportPersonalizzato = False
PclsReport.OrientamentoReport = tsPortrait
Set PclsReport.ActiveInterface = ActiveInterface
SqlReport = ""

PclsReport.AddReport "1 - Stampa Offerta Prezzi", "GBRP_OFFERTE.rpt", PclsReport.PercorsoReport, False, tsPortrait
PclsReport.AddReport "2 - Stampa Offerta Alternativa", "GBRP_OFFERTE_alternative.rpt", PclsReport.PercorsoReport, False, tsPortrait
PclsReport.AddReport "3 - Stampa Offerta No Prezzi", "GBRP_OFFERTE_NO_PREZZI.rpt", PclsReport.PercorsoReport, False, tsPortrait
PclsReport.AddReport "4 - Stampa Offerta Completa", "GBRP_OFFERTE_COMPLETA.rpt", PclsReport.PercorsoReport, False, tsPortrait
PclsReport.AddReport "5 - Stampa Offerta Capitolato", "GBRP_OFFERTE_CAPITOLATO.rpt", PclsReport.PercorsoReport, False, tsPortrait
PclsReport.AddReport "6 - Stampa Offerta PDF", "GBRP_OFFERTE_PDF.rpt", PclsReport.PercorsoReport, False, tsPortrait
PclsReport.AddReport "7 - Stampa Offerta Varia 2", "GBRP_OFFERTE_VARIA2.rpt", PclsReport.PercorsoReport, False, tsPortrait

PclsReport.OpenReport

If PclsReport.Stato <> tsOK Then
   Set PclsReport = Nothing
End If
End Sub


Private Sub CMD_REFRESH_Click()
  If Not QGridDocumenti.DataSource.EOF Then
  CurrentGridPosition = QGridDocumenti.DataSource(0).value
  Else
  CurrentGridPosition = 0
  End If
  Call ImpostaVirtualFrame
  Call RicalcolaImportoTotale
  RiposizionaCursore (CurrentGridPosition)
  End Sub

Private Sub RiposizionaCursore(ID As Variant)
    QGridDocumenti.DataSource.MoveFirst
    
    Do While Not QGridDocumenti.DataSource.EOF
        If QGridDocumenti.DataSource(0).value = ID Then Exit Do
        QGridDocumenti.DataSource.MoveNext
    Loop
End Sub

Private Sub CMD_REFRESH_TXT2_Click()
       
       TXT_GB06_TEXT2.Text = Replace(TXT_GB06_TEXT2.Text, "<Agente>", NVL(TXT_GB06_AGENTE_DEC.Text, "<Agente>"))
       TXT_GB06_TEXT2.Text = Replace(TXT_GB06_TEXT2.Text, "<Utente>", NVL(TXT_GB06_PROPRIETARIO.Text, "<Utente>"))
       TXT_GB06_TEXT2.Refresh
       Call CMD_SAVE_Click
End Sub

Private Sub CMD_REFRESH_TXT3_Click()
       TXT_GB06_TEXT3.Text = Replace(TXT_GB06_TEXT3.Text, "<Pagamento>", NVL(TXT_CG62_DESCPAG.Text, "<Pagamento>"))
       TXT_GB06_TEXT3.Refresh
       Call CMD_SAVE_Click
End Sub

Private Sub CMD_REFRESH_TXT5_Click()
  'TXT_GB06_TEXT5.Text = TextBox1.Text

       TXT_GB06_TEXT5.Text = Replace(TXT_GB06_TEXT5.Text, "<Agente>", NVL(TXT_GB06_AGENTE_DEC.Text, "<Agente>"))
       TXT_GB06_TEXT5.Text = Replace(TXT_GB06_TEXT5.Text, "<Utente>", NVL(TXT_GB06_PROPRIETARIO.Text, "<Utente>"))
       TXT_GB06_TEXT5.Refresh
       Call CMD_SAVE_Click

End Sub

Public Function myCheckFiletr() As Boolean
Dim CheckFiletr As Boolean
CheckFiletr = True

If TXT_FAM.Text <> 0 Or NVL(TXT_FAM.Text, 0) <> 0 Then
    TXT_FAM.SetFocus
    CheckFiletr = False
End If

If TXT_SFAM.Text <> 0 Or NVL(TXT_SFAM.Text, 0) <> 0 Then
    TXT_SFAM.SetFocus
    CheckFiletr = False
End If

If TXT_GRUPPO.Text <> 0 Or NVL(TXT_GRUPPO.Text, 0) <> 0 Then
    TXT_GRUPPO.SetFocus
    CheckFiletr = False
End If

If TXT_SGRUPPO.Text <> 0 Or NVL(TXT_SGRUPPO.Text, 0) <> 0 Then
    TXT_SGRUPPO.SetFocus
    myCheckFiletr = False
End If

If TXT_GRST1.Text <> 0 Or NVL(TXT_GRST1.Text, 0) <> 0 Then
    TXT_GRST1.SetFocus
    CheckFiletr = False
End If

If TXT_GRST2.Text <> 0 Or NVL(TXT_GRST2.Text, 0) <> 0 Then
    TXT_GRST2.SetFocus
    CheckFiletr = False
End If

If TXT_GRST3.Text <> 0 Or NVL(TXT_GRST3.Text, 0) <> 0 Then
    TXT_GRST3.SetFocus
    CheckFiletr = False
End If

If TXT_GRST4.Text <> 0 Or NVL(TXT_GRST4.Text, 0) <> 0 Then
    TXT_GRST4.SetFocus
    CheckFiletr = False
End If

myCheckFiletr = CheckFiletr
End Function

Private Sub CMD_RICPERCPROVV_Click()
Call RicalcolaImporto
TXT_GB07_PERCPROVV.Text = RecuperaProvvigione(QGridDocumenti.DataSource("GB07_PREZZO").value, QGridDocumenti.DataSource("GB07_IMPORTO").value, NVL(QGridDocumenti.DataSource("MG66_FAM_MG53").value, ""), NVL(QGridDocumenti.DataSource("MG66_SFAM_MG54").value, ""), TXT_GB06_AGENTE.Text, NVL(QGridDocumenti.DataSource("GB07_SC1").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC2").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC3").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC4").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC5").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC6").value, 0))

End Sub

Private Sub CMD_SIMULA_Click()
Dim CheckFilter As Boolean

'CheckFilter = False
If CMB_TIPOVARIAZIONE.Text <> 99 Then
    If NVL(TXT_PRECENTUALE.Text, "") = "" Then
    
        MsgBox "Inserire valore da applicare", vbCritical, "Gestione Offerta"
        TXT_PRECENTUALE.SetFocus
        Exit Sub
    End If
End If

CheckFilter = myCheckFiletr()

If CheckFilter = True Then

    If MsgBox("Non hai selezionato filtri." & vbCrLf & "Confermi di voler procedere, la variazione verrà applicata su tutti gli articoli!!", vbYesNo) = vbNo Then
          ' MsgBox "Esci"
          Exit Sub
         
    Else
   ' MsgBox "Elabora"
      Call ElaboraSimulazione(CheckFilter)
    End If
Else

'MsgBox "Elabora"
      Call ElaboraSimulazione(CheckFilter)

End If
CMD_CONFERMA.Enabled = True
End Sub


Private Sub RipristinaOriginale(TABLENAME As String)

Dim StringaSQL As String

'Dim TABLENAME As String
Dim rstSIMULAZIONE As ADODB.Recordset
        
        
StringaSQL = "IF OBJECT_ID('" & TABLENAME & "', 'U') IS NOT NULL " & _
             " DROP TABLE " & TABLENAME & ";"
Gcon_Connect.Execute StringaSQL
        
StringaSQL = " SELECT 0 as GB07_SEL, * , isnull(GB07_IMPORTO,0) AS GB07_IMPORTO_NEW, isnull(GB07_PREZZO,0) as GB07_PREZZO_NEW , isnull(GB07_SC1,0) as GB07_SC1_NEW, isnull(GB07_SC2,0) as GB07_SC2_NEW, isnull(GB07_SC3,0) as GB07_SC3_NEW, isnull(GB07_SC4,0) as GB07_SC4_NEW, isnull(GB07_SC5,0) as GB07_SC5_NEW, isnull(GB07_SC6,0) as GB07_SC6_NEW, "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_GRUSTAT1_MG74  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_GRUSTAT1_MG74,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_GRUSTAT2_MG75  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_GRUSTAT2_MG75,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_GRUSTAT3_MG76  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_GRUSTAT3_MG76,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_GRUSTAT4_MG77  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_GRUSTAT4_MG77,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_FAM_MG53  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_FAM_MG53,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_SFAM_MG54  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_SFAM_MG54,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_GRUPPO_MG55  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_GRUPPO_MG55,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_SGRUPPO_MG56  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_SGRUPPO_MG56  "
  
  StringaSQL = StringaSQL & " INTO " & TABLENAME & " FROM GB07_CORPODOC "
  
  StringaSQL = StringaSQL & " WHERE GB07_ID_GB06 = " & IDGB06
  StringaSQL = StringaSQL & " ORDER BY   GB07_RAG, GB07_SEQ "
  
  Gcon_Connect.Execute StringaSQL
  
  StringaSQL = "Select * from " & TABLENAME & " WHERE GB07_ID_GB06 = " & IDGB06
  
  Set rstSIMULAZIONE = Nothing
    
    Set rstSIMULAZIONE = New ADODB.Recordset
    With rstSIMULAZIONE
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
    
    If Not rstSIMULAZIONE.EOF Then
        Call INIT_SIMULAZIONE
        Set QGRID_SIMULAZIONE.DataSource = Nothing
        Set QGRID_SIMULAZIONE.DataSource = rstSIMULAZIONE
        QGRID_SIMULAZIONE.FullExpand
        Call RicalcolaImportoTotaleSimulazione(TABLENAME)
        Exit Sub
    End If
End Sub

Private Sub ElaboraSimulazione(CheckFilter As Boolean)
On Error GoTo ErrTrap:
Dim StringaSQL As String

Dim TABLENAME As String
Dim rstSIMULAZIONE As ADODB.Recordset
  
  TABLENAME = "TMP_SIMULAZIONE_" & ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Codice

If CMB_TIPOVARIAZIONE.Text = 99 Then
Call RipristinaOriginale(TABLENAME)
Exit Sub
End If

    
  If TMS_MANTIENI.Text = 0 Then
    StringaSQL = "IF OBJECT_ID('" & TABLENAME & "', 'U') IS NOT NULL " & _
                 " DROP TABLE " & TABLENAME & ";"
    Gcon_Connect.Execute StringaSQL
  End If
  If TMS_MANTIENI.Text = 0 Then
      StringaSQL = " SELECT * , isnull(GB07_IMPORTO,0) AS GB07_IMPORTO_NEW, isnull(GB07_PREZZO,0) as GB07_PREZZO_NEW , isnull(GB07_SC1,0) as GB07_SC1_NEW, isnull(GB07_SC2,0) as GB07_SC2_NEW, isnull(GB07_SC3,0) as GB07_SC3_NEW, isnull(GB07_SC4,0) as GB07_SC4_NEW, isnull(GB07_SC5,0) as GB07_SC5_NEW, isnull(GB07_SC6,0) as GB07_SC6_NEW, "

  Else
        StringaSQL = " SELECT * ,  "

  End If
  
  StringaSQL = StringaSQL & " (SELECT        MG66_GRUSTAT1_MG74  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_GRUSTAT1_MG74,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_GRUSTAT2_MG75  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_GRUSTAT2_MG75,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_GRUSTAT3_MG76  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_GRUSTAT3_MG76,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_GRUSTAT4_MG77  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_GRUSTAT4_MG77,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_FAM_MG53  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_FAM_MG53,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_SFAM_MG54  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_SFAM_MG54,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_GRUPPO_MG55  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_GRUPPO_MG55,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_SGRUPPO_MG56  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_SGRUPPO_MG56  "
  
  If TMS_MANTIENI.Text = 0 Then
    StringaSQL = StringaSQL & " INTO " & TABLENAME & " FROM GB07_CORPODOC "
  Else
    StringaSQL = StringaSQL & " FROM " & TABLENAME & "  "
  End If
  
  StringaSQL = StringaSQL & " WHERE GB07_ID_GB06 = " & IDGB06
  StringaSQL = StringaSQL & " ORDER BY   GB07_RAG, GB07_SEQ "
  
  Gcon_Connect.Execute StringaSQL
  
  ' INIZIO - applica variazioni
  
  StringaSQL = "UPDATE " & TABLENAME
  StringaSQL = StringaSQL & " SET "

  Select Case CMB_TIPOVARIAZIONE.Text
  Case 0
      StringaSQL = StringaSQL & " GB07_PREZZO_NEW = GB07_PREZZO - (GB07_PREZZO * " & SQLDouble(TXT_PRECENTUALE.Text) & " / 100) "
  Case 1
      StringaSQL = StringaSQL & " GB07_SC1_NEW = " & SQLDouble(TXT_PRECENTUALE.Text)
  Case 2
      StringaSQL = StringaSQL & " GB07_SC2_NEW = " & SQLDouble(TXT_PRECENTUALE.Text)
  Case 3
      StringaSQL = StringaSQL & " GB07_SC3_NEW = " & SQLDouble(TXT_PRECENTUALE.Text)
  Case 4
      StringaSQL = StringaSQL & " GB07_SC4_NEW = " & SQLDouble(TXT_PRECENTUALE.Text)
  Case 5
      StringaSQL = StringaSQL & " GB07_SC5_NEW = " & SQLDouble(TXT_PRECENTUALE.Text)
  Case 6
      StringaSQL = StringaSQL & " GB07_SC6_NEW = " & SQLDouble(TXT_PRECENTUALE.Text)
  Case 7
    StringaSQL = StringaSQL & " GB07_PREZZO_NEW = GB07_PREZZO + (GB07_PREZZO * " & SQLDouble(TXT_PRECENTUALE.Text) & " / 100) "
  Case 99
    'StringaSQL = StringaSQL & " GB07_PREZZO_NEW = GB07_PREZZO + (GB07_PREZZO * " & TXT_PRECENTUALE.Text & " / 100) "
    'pulisce tutte le variazioni
    'GoTo SaltaAggiornamento
  Case Else
    MsgBox "Tipo operazione non gestita"
    Exit Sub
  End Select
  
  StringaSQL = StringaSQL & " WHERE GB07_ID_GB06 = " & IDGB06

  
  If CheckFilter = True Then
  
      
    
  Else
        If TXT_FAM.Text <> 0 Or NVL(TXT_FAM.Text, 0) <> 0 Then
            StringaSQL = StringaSQL & " and MG66_FAM_MG53 = '" & TXT_FAM.Text & "'"
        End If
        
        If TXT_SFAM.Text <> 0 Or NVL(TXT_SFAM.Text, 0) <> 0 Then
           StringaSQL = StringaSQL & " and MG66_SFAM_MG54 = '" & TXT_SFAM.Text & "'"
        End If
        
        If TXT_GRUPPO.Text <> 0 Or NVL(TXT_GRUPPO.Text, 0) <> 0 Then
           StringaSQL = StringaSQL & " and MG66_GRUPPO_MG55 = '" & TXT_GRUPPO.Text & "'"
        End If
        
        If TXT_SGRUPPO.Text <> 0 Or NVL(TXT_SGRUPPO.Text, 0) <> 0 Then
            StringaSQL = StringaSQL & " and MG66_SGRUPPO_MG56 = '" & TXT_SGRUPPO.Text & "'"
        End If
        
        If TXT_GRST1.Text <> 0 Or NVL(TXT_GRST1.Text, 0) <> 0 Then
            StringaSQL = StringaSQL & " and MG66_GRUSTAT1_MG74 = '" & TXT_GRST1.Text & "'"
        End If
        
        If TXT_GRST2.Text <> 0 Or NVL(TXT_GRST2.Text, 0) <> 0 Then
            StringaSQL = StringaSQL & " and MG66_GRUSTAT2_MG75 = '" & TXT_GRST2.Text & "'"
        End If
        
        If TXT_GRST3.Text <> 0 Or NVL(TXT_GRST3.Text, 0) <> 0 Then
           StringaSQL = StringaSQL & " and MG66_GRUSTAT3_MG76 = '" & TXT_GRST3.Text & "'"
        End If
        
        If TXT_GRST4.Text <> 0 Or NVL(TXT_GRST4.Text, 0) <> 0 Then
           StringaSQL = StringaSQL & " and MG66_GRUSTAT4_MG77 = '" & TXT_GRST4.Text & "'"
        End If
  
  End If
  
  Gcon_Connect.Execute StringaSQL
  
  ' FINE   - applica variazioni
SaltaAggiornamento:
  'INIZIO - AGGIORNA TOTALI
  
  StringaSQL = " UPDATE " & TABLENAME & " SET GB07_IMPORTO_NEW = "
  StringaSQL = StringaSQL & "  ((    GB07_PREZZO_NEW - GB07_PREZZO_NEW * isnull(gb07_sc1_new,0) / 100) -"
  StringaSQL = StringaSQL & "  ((    GB07_PREZZO_NEW - GB07_PREZZO_NEW * isnull(gb07_sc1_new,0) / 100) * isnull(gb07_sc2_new,0) / 100) -"
  StringaSQL = StringaSQL & "  (((   GB07_PREZZO_NEW - GB07_PREZZO_NEW * isnull(gb07_sc1_new,0) / 100) * isnull(gb07_sc2_new,0) / 100) * isnull(gb07_sc3_new,0) / 100) -"
  StringaSQL = StringaSQL & "  ((((  GB07_PREZZO_NEW - GB07_PREZZO_NEW * isnull(gb07_sc1_new,0) / 100) * isnull(gb07_sc2_new,0) / 100) * isnull(gb07_sc3_new,0) / 100) * isnull(gb07_sc4_new,0) / 100 ) -"
  StringaSQL = StringaSQL & "  ((((( GB07_PREZZO_NEW - GB07_PREZZO_NEW * isnull(gb07_sc1_new,0) / 100) * isnull(gb07_sc2_new,0) / 100) * isnull(gb07_sc3_new,0) / 100) * isnull(gb07_sc4_new,0) / 100 ) * isnull(gb07_sc5_new,0) / 100 )-"
  StringaSQL = StringaSQL & "  ((((((GB07_PREZZO_NEW - GB07_PREZZO_NEW * isnull(gb07_sc1_new,0) / 100) * isnull(gb07_sc2_new,0) / 100) * isnull(gb07_sc3_new,0) / 100) * isnull(gb07_sc4_new,0) / 100 ) * isnull(gb07_sc5_new,0) / 100 ) * isnull(gb07_sc6_new,0) / 100 ))"
  StringaSQL = StringaSQL & "  WHERE GB07_ID_GB06 = " & IDGB06
  Gcon_Connect.Execute StringaSQL
  'FINE - AGGIORNA TOTALI
  
  
  StringaSQL = "Select * from " & TABLENAME & " WHERE GB07_ID_GB06 = " & IDGB06
  
  Set rstSIMULAZIONE = Nothing
    
    Set rstSIMULAZIONE = New ADODB.Recordset
    With rstSIMULAZIONE
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
    
    If Not rstSIMULAZIONE.EOF Then
        Call INIT_SIMULAZIONE
        Set QGRID_SIMULAZIONE.DataSource = Nothing
        Set QGRID_SIMULAZIONE.DataSource = rstSIMULAZIONE
        QGRID_SIMULAZIONE.FullExpand
        Call RicalcolaImportoTotaleSimulazione(TABLENAME)
        Exit Sub
    End If
    
    
    
    Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("Simulazione Prezzi")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select

End Sub



'Private Sub Command2_Click()
'Call RicalcolaCosto
'End Sub

Private Sub FME_BANCO_AfterAddNew(ByVal fenm_operationresult As FWBO_VIRTUALFRAME.EnumOperationResult)
'  Call RecuperImmagineHyperMedia(NVL(TXT_GB07_CODART_MG66.Text, ""))
'If NVL(TXT_FLPOSA.Text, 0) = "0" Then TXT_GB07_FLPOSA.Text = "0"

End Sub

Private Sub FME_GRUPPI_BeforeUpdate(fbol_Cancel As Boolean)
 On Error GoTo ErrTrap

   rstGRUPPI.Fields("GB07_ID_GB06").value = IDGB06
  


    Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("FME_GRUPPI_BeforeUpdate")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub

Private Sub FME_NOTE_BeforeUpdate(fbol_Cancel As Boolean)
 On Error GoTo ErrTrap

   rstNOTE.Fields("GB08_ID_GB06").value = IDGB06
   rstNOTE.Fields("GB08_OPERATORE").value = ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Descrizione
  


    Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("FME_NOTE_BeforeUpdate")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub

Private Sub PclsReport_BeforePrintReport(Cancel As Boolean, CrepReport As FWBO_REPORT30.CLSFW_OBJSTAMPA, CrepApp As Object)

    CrepReport.RecordSelectionFormula = " ( {GB06_TESTADOC.GB06_ID} = " & IDGB06 & ") "

End Sub

Private Sub CMD_SAVE_Click()
On Error GoTo SegnalaErrore
  Dim Ret As Boolean
    
''  Me.Enabled = False
''
''  If Pbol_Generazione Then
''    Me.Enabled = True
''    Exit Sub
''  Else
''    Pbol_Generazione = True
''  End If
''
''
''
''  ValContanti = 0
''
''  'Controllo inserimento contanti per scontrini
''  If TipoDocumento = 2 And TXT_GB06_CODPAG_CG62.Text = CODPAGCONTANTI Then
''
''        Pbol_Generazione = True
'''        Me.Enabled = True
''
''        'Exit Sub
''    Else
''
''    End If
''
''
''  'controlla cdc inseriti
  
 
  
  If FME_BANCO.Status = tsInsert And NVL(Trim(TXT_GB07_CODART_MG66.Text), "") = "" Then
    Pbol_Generazione = False
     
    
    
    
  '  FME_BANCO.CancelAddNew False, True
  End If
  
  Select Case TipoDocumento
  ' inserire se controllare l'inserimento dei riferimenti
  Case 10
'    If NVL(TXT_RIFERIMENTO.Text, "") = "" Or NVL(TXT_DATARIF.Text, "") = "" Or Not (TXT_DATARIF.IsValid) Then
'      Pbol_Generazione = False
'      Me.Enabled = True
'      MsgBox "Per il tipo documento da generare è obbligatorio inserire i dati di riferimento"
'      Exit Sub
'    End If
  End Select
  
  
'  If rstCorpoBanco.RecordCount = 1 And FME_BANCO.Status = tsInsert Then
'    Pbol_Generazione = False
'    MsgBox "Non ci sono record da creare"
'    Me.Enabled = True
'
'    Exit Sub
'  End If
  
  
    
 
  FrameDoc.Visible = True
  DoEvents
  Me.Enabled = False
  
  
 
  
  
'  If Pbol_BloccoFido Then
'    MsgBox "Attenzione, cliente fuori fido di " & CDbl(TXT_TOTALEDOC.Text) - ResiduoFido
'    'MsgBox "Attenzione blocco per fido. "
'    Exit Sub
'  End If
      
  'Scrive dati in GB06
  strSQL = "UPDATE GB06_TESTADOC "
  strSQL = strSQL & " SET  GB06_TIPOCF_CG44    = 0 "
  strSQL = strSQL & "    , GB06_CLIFOR_CG44    = " & NVL(TXT_GB06_CLIFOR_CG44.Text, "0")
  
  strSQL = strSQL & "    , GB06_CODPAG_CG62    = '" & NVL(TXT_GB06_CODPAG_CG62.Text, "") & "'"
  strSQL = strSQL & "    , GB06_VETTORE_MG14   = ''"
  strSQL = strSQL & "    , GB06_NUMDOC   = '" & NVL(TXT_GB06_NUMDOC.Text, "") & "'"
  strSQL = strSQL & "    , GB06_NREV   = " & NVL(TXT_GB06_NREV.Text, "1")
  strSQL = strSQL & "    , GB06_NVERS   = " & NVL(TXT_GB06_NVERS.Text, "1")
  strSQL = strSQL & "    , GB06_DTDOC   = " & SQLDate(NVL(TXT_GB06_DTDOC.Text, "01/01/2099"))
  strSQL = strSQL & "    , GB06_CODCOMM   = '" & NVL(TXT_GB06_CODCOMM.Text, "") & "'"
  strSQL = strSQL & "    , GB06_DTCHIUSURA   = " & SQLDate(NVL(TXT_GB06_DTCHIUSURA.Text, "01/01/2099"))
  strSQL = strSQL & "    , GB06_PERCCHIUSURA   = " & SQLDouble(NVL(TXT_GB06_PERCCHIUSURA.Text, "0"))
  strSQL = strSQL & "    , GB06_RESPONSABILE   = '" & (NVL(TXT_GB06_RESPONSABILE.Text, "")) & "'"
  strSQL = strSQL & "    , GB06_PERCRIBGARA   = " & SQLDouble(NVL(TXT_GB06_PERCRIBGARA.Text, "0"))
  strSQL = strSQL & "    , GB06_PERCTRASP   = " & SQLDouble(NVL(TXT_GB06_PERCTRASP.Text, "0"))
  strSQL = strSQL & "    , GB06_TIPOOFFERTA   = '" & NVL(TXT_GB06_TIPOOFFERTA.Text, "") & "'"
  strSQL = strSQL & "    , GB06_AGENTE   = '" & NVL(TXT_GB06_AGENTE.Text, "") & "'"
  strSQL = strSQL & "    , GB06_TIPOAREA   = '" & NVL(TXT_GB06_TIPOAREA.Text, "") & "'"
  strSQL = strSQL & "    , GB06_STATODOC   = '" & NVL(TXT_GB06_STATODOC.Text, "00") & "'"
  strSQL = strSQL & "    , GB06_BUDGET   = " & SQLDouble(NVL(TXT_GB06_BUDGET.Text, "0"))
  strSQL = strSQL & "    , GB06_FORECAST   = " & SQLDouble(NVL(TXT_GB06_FORECAST.Text, "0"))
  strSQL = strSQL & "    , GB06_CONSUNTIVO   = " & SQLDouble(NVL(TXT_GB06_CONSUNTIVO.Text, "0"))
  strSQL = strSQL & "    , GB06_NOMEOFFERTA   = '" & SQLString(NVL(TXT_GB06_NOMEOFFERTA.Text, "")) & "'"
  strSQL = strSQL & "    , GB06_CIG   = '" & NVL(TXT_GB06_CIG.Text, "") & "'"
  strSQL = strSQL & "    , GB06_CUP   = '" & NVL(TXT_GB06_CUP.Text, "") & "'"
  strSQL = strSQL & "    , GB06_DTULTMOD   = " & SQLDate(Now())
  strSQL = strSQL & "    , GB06_ALLACA   = '" & SQLString(NVL(TXT_GB06_ALLACA.Text, "")) & "'"
  strSQL = strSQL & "    , GB06_TEXT1   = '" & SQLString(NVL(TXT_GB06_TEXT1.Text, "")) & "'"
  strSQL = strSQL & "    , GB06_TEXT2   = '" & SQLString(NVL(TXT_GB06_TEXT2.Text, "")) & "'"
  strSQL = strSQL & "    , GB06_TEXT3   = '" & SQLString(NVL(TXT_GB06_TEXT3.Text, "")) & "'"
  strSQL = strSQL & "    , GB06_TEXT4   = '" & SQLString(NVL(TXT_GB06_TEXT4.Text, "")) & "'"
  strSQL = strSQL & "    , GB06_TEXT5   = '" & SQLString(NVL(TXT_GB06_TEXT5.Text, "")) & "'"
  strSQL = strSQL & "    , GB06_TEXT6   = '" & SQLString(NVL(TXT_GB06_TEXT6.Text, "")) & "'"
  strSQL = strSQL & "    , GB06_UTENTECRE   = '" & SQLString(NVL(TXT_GB06_PROPRIETARIO.Text, "")) & "'"
  strSQL = strSQL & "    , GB06_CODDESTIN_MG22   = '" & SQLString(NVL(TXT_GB06_CODDESTIN_MG22.Text, "")) & "'"
  
  
  
  strSQL = strSQL & " WHERE GB06_ID   = " & NVL(TXT_GB06_ID.Text, IDGB06)
  
  Gcon_Connect.Execute strSQL
  
  'fine salvataggio offerta
  FrameDoc.Visible = False
 ' Exit Sub
  
  ContaRighe = 0
  Set rstScriveDoc = New ADODB.Recordset
  strSQL = "SELECT *  , GB07_CODART_MG66 AS DESART "
  strSQL = strSQL & "  FROM GB07_CORPODOC   "
  strSQL = strSQL & " WHERE GB07_ID_GB06 = " & IDGB06
  
  Set rstScriveDoc = Gcon_Connect.Execute(strSQL, , adCmdText)
  
  If Not rstScriveDoc.EOF Then
       ContaRighe = rstScriveDoc.RecordCount
       
       FME_BANCO.MoveFirst
       FME_BANCO.Update
       
   ' If CheckCDC(IDGB06) Then
       
   '     Call ScriviDocumento
   ' Else
   '     Pbol_Generazione = False
   '     MsgBox "Mancano i Centri di Costo"
   '     Me.Enabled = True
        
   '     Exit Sub
    
   ' End If
    
  Else
    
    Pbol_Generazione = False
    'MsgBox "Non ci sono righe da generare"
    Me.Enabled = True
    
    'Exit Sub
  
  End If
  
  FrameDoc.Visible = False
  
  Set rstScriveDoc = Nothing
  
  If ContaRighe = 0 Then
    Pbol_Generazione = False
    'MsgBox "Non ci sono righe Caricate"
    Me.Enabled = True
    Exit Sub
  End If
  
  Call RicalcolaImportoTotale
  Call ImpostaVirtualFrame
  

  Me.Enabled = True
  Call TXT_GB06_ID_AfterItem(False)
  Exit Sub
SegnalaErrore:
  'Scrivi log errore
  Errore = "Attenzione! L'applicazione ha generato il seguente errore: " & Err.Number & " - " & Err.Description

'  ClasseInternaRegDoc = Nothing
  Select Case VisualizzaErrore("cmdGeneraFat_ButtonClick")
    Case vbAbort
        FrameDoc.Visible = False
        Me.Enabled = True
        
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
  End Select
End Sub

Private Sub cmdApriDoc_ButtonClick()
If NVL(NumRegGenerato, "") <> "" Then
        'Call StampaDocumento
        Call ApriDoc(CodiceDocumento, NumRegGenerato)
    End If
End Sub

Private Sub cmdGeneraFat1_ButtonClick()

End Sub

Private Sub cmdModifica_ButtonClick()
  
  TipoDocumento = 5
  lblTipoDocumento.Caption = "Creazione Offerta Cliente"
  Call ModificaTestata
  TXT_GB06_ID.SetFocus
End Sub

Private Sub cmdNewRevision_ButtonClick()
Dim strNewVersione As String
Dim CurUserName As String
CurUserName = ActualComputerName & ActualUserName

If MsgBox("Confermi la Creazione di una nuova Revisione dell'offerta " & TXT_GB06_NOMEOFFERTA.Text & " ?", vbYesNo) = vbNo Then
      Exit Sub
End If

Call CMD_SAVE_Click

ProgInsertTestata = CurUserName & Year(Now) & String(2 - Len(Month(Now)), "0") & Month(Now) & String(2 - Len(Day(Now)), "0") & Day(Now) & String(2 - Len(Hour(Time)), "0") & Hour(Time) & String(2 - Len(Minute(Time)), "0") & Minute(Time) & String(2 - Len(Second(Time)), "0") & Second(Time) & Right(Format(Timer, "#0.00"), 2)

strNewVersione = " insert into GB06_TESTADOC"
strNewVersione = strNewVersione + "("
strNewVersione = strNewVersione + " GB06_PROG,"
strNewVersione = strNewVersione + " GB06_DITTA_CG18,"
strNewVersione = strNewVersione + " GB06_TIPODOC,"
strNewVersione = strNewVersione + " GB06_DATA,"
strNewVersione = strNewVersione + " GB06_GRUPPOCRE,"
strNewVersione = strNewVersione + " GB06_UTENTECRE,"
strNewVersione = strNewVersione + " GB06_TIPOCF_CG44,"
strNewVersione = strNewVersione + " GB06_CLIFOR_CG44,"
strNewVersione = strNewVersione + " GB06_BARCODE_LL01,"
strNewVersione = strNewVersione + " GB06_CODDESTIN_MG22,"
strNewVersione = strNewVersione + " GB06_CODPAG_CG62,"
strNewVersione = strNewVersione + " GB06_VETTORE_MG14,"
strNewVersione = strNewVersione + " GB06_NUMREG_CO99,"
strNewVersione = strNewVersione + " GB06_NUMDOC,"
strNewVersione = strNewVersione + " GB06_NREV,"
strNewVersione = strNewVersione + " GB06_NVERS,"
strNewVersione = strNewVersione + " GB06_DTDOC,"
strNewVersione = strNewVersione + " GB06_CODCOMM,"
strNewVersione = strNewVersione + " GB06_DTCHIUSURA,"
strNewVersione = strNewVersione + " GB06_PERCCHIUSURA,"
strNewVersione = strNewVersione + " GB06_RESPONSABILE,"
strNewVersione = strNewVersione + " GB06_PERCRIBGARA,"
strNewVersione = strNewVersione + " GB06_TIPOOFFERTA,"

strNewVersione = strNewVersione + " GB06_AGENTE ,"
strNewVersione = strNewVersione + " GB06_TIPOAREA,"
strNewVersione = strNewVersione + " GB06_STATODOC,"
strNewVersione = strNewVersione + " GB06_BUDGET,"
strNewVersione = strNewVersione + " GB06_FORECAST,"
strNewVersione = strNewVersione + " GB06_CONSUNTIVO,"
strNewVersione = strNewVersione + " GB06_NOMEOFFERTA,"
strNewVersione = strNewVersione + " GB06_DTULTMOD,"
strNewVersione = strNewVersione + " GB06_CIG,"
strNewVersione = strNewVersione + " GB06_CUP,"
strNewVersione = strNewVersione + " GB06_ALLACA,"
strNewVersione = strNewVersione + " GB06_TEXT1,"
strNewVersione = strNewVersione + " GB06_TEXT2,"
strNewVersione = strNewVersione + " GB06_TEXT3,"
strNewVersione = strNewVersione + " GB06_TEXT4,"
strNewVersione = strNewVersione + " GB06_TEXT5,"
strNewVersione = strNewVersione + " GB06_TEXT6"
strNewVersione = strNewVersione + ")"
strNewVersione = strNewVersione + " select"
strNewVersione = strNewVersione + " '" & ProgInsertTestata & "', GB06_DITTA_CG18, GB06_TIPODOC, GB06_DATA, GB06_GRUPPOCRE, GB06_UTENTECRE, GB06_TIPOCF_CG44,"
strNewVersione = strNewVersione + "                         GB06_CLIFOR_CG44, GB06_BARCODE_LL01, GB06_CODDESTIN_MG22, GB06_CODPAG_CG62, GB06_VETTORE_MG14, GB06_NUMREG_CO99, GB06_NUMDOC,"
strNewVersione = strNewVersione + "                         max(GB06_NREV+1) as GB06_NREV , 1, GB06_DTDOC, GB06_CODCOMM, GB06_DTCHIUSURA, GB06_PERCCHIUSURA,GB06_RESPONSABILE, GB06_PERCRIBGARA, GB06_TIPOOFFERTA,"
strNewVersione = strNewVersione + "                          GB06_AGENTE , GB06_TIPOAREA, GB06_STATODOC, GB06_BUDGET, GB06_FORECAST, GB06_CONSUNTIVO, GB06_NOMEOFFERTA, { fn NOW() }, GB06_CIG, GB06_CUP, GB06_ALLACA, GB06_TEXT1, GB06_TEXT2, GB06_TEXT3, GB06_TEXT4, GB06_TEXT5, GB06_TEXT6"
strNewVersione = strNewVersione + " from GB06_TESTADOC where GB06_ID = " & TXT_GB06_ID.Text
strNewVersione = strNewVersione + " group by GB06_PROG, GB06_DITTA_CG18, GB06_TIPODOC, GB06_DATA, GB06_GRUPPOCRE, GB06_UTENTECRE, GB06_TIPOCF_CG44,"
strNewVersione = strNewVersione + "                         GB06_CLIFOR_CG44, GB06_BARCODE_LL01, GB06_CODDESTIN_MG22, GB06_CODPAG_CG62, GB06_VETTORE_MG14, GB06_NUMREG_CO99, GB06_NUMDOC,"
strNewVersione = strNewVersione + "                          GB06_NVERS, GB06_DTDOC, GB06_CODCOMM, GB06_DTCHIUSURA, GB06_PERCCHIUSURA, GB06_RESPONSABILE, GB06_PERCRIBGARA, GB06_TIPOOFFERTA,"
strNewVersione = strNewVersione + "                          GB06_AGENTE , GB06_TIPOAREA, GB06_STATODOC, GB06_BUDGET, GB06_FORECAST, GB06_CONSUNTIVO, GB06_NOMEOFFERTA, GB06_DTULTMOD, GB06_CIG, GB06_CUP, GB06_ALLACA, GB06_TEXT1, GB06_TEXT2, GB06_TEXT3, GB06_TEXT4, GB06_TEXT5, GB06_TEXT6"
Gcon_Connect.Execute strNewVersione

strNewVersione = ""
strNewVersione = strNewVersione + "insert into GB07_CORPODOC "
strNewVersione = strNewVersione + " ( "
strNewVersione = strNewVersione + " GB07_ID_GB06, GB07_CODART_MG66, GB07_OPZIONE_MG5E , "
strNewVersione = strNewVersione + " GB07_QTA, GB07_PREZZO, GB07_SCCORPO, GB07_SC5, "
strNewVersione = strNewVersione + " GB07_SC4, GB07_SC3, GB07_SC2, GB07_SC1, "
strNewVersione = strNewVersione + " GB07_SC6, GB07_SCPIEDE, GB07_IMPORTO, GB07_DESCART, "
strNewVersione = strNewVersione + " GB07_MAGGIORAZIONE, GB07_SEQ, GB07_RAG, "
strNewVersione = strNewVersione + " GB07_ALT , GB07_FLPOSA, GB07_IMG, GB07_PATHIMG, GB07_TIPOCF_CG44, GB07_CLIFOR_CG44, GB07_COSTO, GB07_IMPPROVV,GB07_PERCPROVV  "
strNewVersione = strNewVersione + " ) "
strNewVersione = strNewVersione + " select "
strNewVersione = strNewVersione + " (select max(GB06_ID) as ID from GB06_TESTADOC where GB06_PROG like '" & CurUserName & "%'), GB07_CODART_MG66, GB07_OPZIONE_MG5E, "
strNewVersione = strNewVersione + " GB07_QTA, GB07_PREZZO, GB07_SCCORPO, GB07_SC5, "
strNewVersione = strNewVersione + " GB07_SC4, GB07_SC3, GB07_SC2, GB07_SC1, "
strNewVersione = strNewVersione + " GB07_SC6, GB07_SCPIEDE, GB07_IMPORTO, GB07_DESCART, "
strNewVersione = strNewVersione + " GB07_MAGGIORAZIONE, GB07_SEQ, GB07_RAG, "
strNewVersione = strNewVersione + " GB07_ALT , GB07_FLPOSA, GB07_IMG, GB07_PATHIMG, GB07_TIPOCF_CG44, GB07_CLIFOR_CG44, GB07_COSTO, GB07_IMPPROVV,GB07_PERCPROVV  "
strNewVersione = strNewVersione + " from GB07_CORPODOC where gb07_id_gb06 = " & TXT_GB06_ID.Text
Gcon_Connect.Execute strNewVersione

Call VariaStato("04", TXT_GB06_ID.Text)


If MsgBox("Vuoi Aprire la nuova Revisione ?", vbYesNo) = vbYes Then
      TXT_GB06_ID.Text = GetValFromQuery("select max(GB06_ID) as ID from GB06_TESTADOC where GB06_PROG like '" & CurUserName & "%'", 0, Gcon_Connect)
      'Call TXT_GB06_ID_StartDecode
      TXT_GB06_ID.SetFocus
      
End If


End Sub


Private Sub VariaStato(NewStato As String, IdOfferta As Long)
    strSQL = "UPDATE GB06_TESTADOC "
    strSQL = strSQL & " SET  GB06_STATOdoc    = '" & NewStato & "'"
    strSQL = strSQL & " WHERE GB06_ID   = " & IdOfferta
    
    Gcon_Connect.Execute strSQL
End Sub

Private Sub ConfermaSimulazione(TABLENAME As String)
Dim strNewVersione As String
Dim CurUserName, strInserisciNota, strTesto As String
CurUserName = ActualComputerName & ActualUserName
  
ProgInsertTestata = CurUserName & Year(Now) & String(2 - Len(Month(Now)), "0") & Month(Now) & String(2 - Len(Day(Now)), "0") & Day(Now) & String(2 - Len(Hour(Time)), "0") & Hour(Time) & String(2 - Len(Minute(Time)), "0") & Minute(Time) & String(2 - Len(Second(Time)), "0") & Second(Time) & Right(Format(Timer, "#0.00"), 2)

If MsgBox("Confermi la Creazione di una nuova Versione " & TXT_GB06_NOMEOFFERTA.Text & " ?", vbYesNo) = vbNo Then
      Exit Sub
End If


strNewVersione = " insert into GB06_TESTADOC"
strNewVersione = strNewVersione + "("
strNewVersione = strNewVersione + " GB06_PROG,"
strNewVersione = strNewVersione + " GB06_DITTA_CG18,"
strNewVersione = strNewVersione + " GB06_TIPODOC,"
strNewVersione = strNewVersione + " GB06_DATA,"
strNewVersione = strNewVersione + " GB06_GRUPPOCRE,"
strNewVersione = strNewVersione + " GB06_UTENTECRE,"
strNewVersione = strNewVersione + " GB06_TIPOCF_CG44,"
strNewVersione = strNewVersione + " GB06_CLIFOR_CG44,"
strNewVersione = strNewVersione + " GB06_BARCODE_LL01,"
strNewVersione = strNewVersione + " GB06_CODDESTIN_MG22,"
strNewVersione = strNewVersione + " GB06_CODPAG_CG62,"
strNewVersione = strNewVersione + " GB06_VETTORE_MG14,"
strNewVersione = strNewVersione + " GB06_NUMREG_CO99,"
strNewVersione = strNewVersione + " GB06_NUMDOC,"
strNewVersione = strNewVersione + " GB06_NREV,"
strNewVersione = strNewVersione + " GB06_NVERS,"
strNewVersione = strNewVersione + " GB06_DTDOC,"
strNewVersione = strNewVersione + " GB06_CODCOMM,"
strNewVersione = strNewVersione + " GB06_DTCHIUSURA,"
strNewVersione = strNewVersione + " GB06_PERCCHIUSURA,"
strNewVersione = strNewVersione + " GB06_RESPONSABILE,"
strNewVersione = strNewVersione + " GB06_PERCRIBGARA,"
strNewVersione = strNewVersione + " GB06_TIPOOFFERTA,"
strNewVersione = strNewVersione + " GB06_AGENTE ,"
strNewVersione = strNewVersione + " GB06_TIPOAREA,"
strNewVersione = strNewVersione + " GB06_STATODOC,"
strNewVersione = strNewVersione + " GB06_BUDGET,"
strNewVersione = strNewVersione + " GB06_FORECAST,"
strNewVersione = strNewVersione + " GB06_CONSUNTIVO,"
strNewVersione = strNewVersione + " GB06_NOMEOFFERTA,"
strNewVersione = strNewVersione + " GB06_DTULTMOD,"
strNewVersione = strNewVersione + " GB06_CIG,"
strNewVersione = strNewVersione + " GB06_CUP,"
strNewVersione = strNewVersione + " GB06_ALLACA,"
strNewVersione = strNewVersione + " GB06_TEXT1,"
strNewVersione = strNewVersione + " GB06_TEXT2,"
strNewVersione = strNewVersione + " GB06_TEXT3,"
strNewVersione = strNewVersione + " GB06_TEXT4,"
strNewVersione = strNewVersione + " GB06_TEXT5,"
strNewVersione = strNewVersione + " GB06_TEXT6"
strNewVersione = strNewVersione + ")"
strNewVersione = strNewVersione + " select"
strNewVersione = strNewVersione + " '" & ProgInsertTestata & "', GB06_DITTA_CG18, GB06_TIPODOC, GB06_DATA, GB06_GRUPPOCRE, GB06_UTENTECRE, GB06_TIPOCF_CG44,"
strNewVersione = strNewVersione + "                         GB06_CLIFOR_CG44, GB06_BARCODE_LL01, GB06_CODDESTIN_MG22, GB06_CODPAG_CG62, GB06_VETTORE_MG14, GB06_NUMREG_CO99, GB06_NUMDOC,"
strNewVersione = strNewVersione + "                         GB06_NREV, max(GB06_NVERS+1) as GB06_NVERS, GB06_DTDOC, GB06_CODCOMM, GB06_DTCHIUSURA, GB06_PERCCHIUSURA,GB06_RESPONSABILE, GB06_PERCRIBGARA, GB06_TIPOOFFERTA,"
strNewVersione = strNewVersione + "                         GB06_AGENTE , GB06_TIPOAREA, GB06_STATODOC, GB06_BUDGET, GB06_FORECAST, GB06_CONSUNTIVO, GB06_NOMEOFFERTA, { fn NOW() }, GB06_CIG, GB06_CUP, GB06_ALLACA, GB06_TEXT1, GB06_TEXT2, GB06_TEXT3, GB06_TEXT4, GB06_TEXT5, GB06_TEXT6"
strNewVersione = strNewVersione + " from GB06_TESTADOC where GB06_ID = " & TXT_GB06_ID.Text
strNewVersione = strNewVersione + " group by GB06_PROG, GB06_DITTA_CG18, GB06_TIPODOC, GB06_DATA, GB06_GRUPPOCRE, GB06_UTENTECRE, GB06_TIPOCF_CG44,"
strNewVersione = strNewVersione + "   GB06_CLIFOR_CG44, GB06_BARCODE_LL01, GB06_CODDESTIN_MG22, GB06_CODPAG_CG62, GB06_VETTORE_MG14, GB06_NUMREG_CO99, GB06_NUMDOC,"
strNewVersione = strNewVersione + "  GB06_NREV, GB06_DTDOC, GB06_CODCOMM, GB06_DTCHIUSURA, GB06_PERCCHIUSURA, GB06_RESPONSABILE, GB06_PERCRIBGARA, GB06_TIPOOFFERTA,"
strNewVersione = strNewVersione + "  GB06_AGENTE , GB06_TIPOAREA, GB06_STATODOC, GB06_BUDGET, GB06_FORECAST, GB06_CONSUNTIVO, GB06_NOMEOFFERTA, GB06_DTULTMOD, GB06_CIG, GB06_CUP, GB06_ALLACA, GB06_TEXT1, GB06_TEXT2, GB06_TEXT3, GB06_TEXT4, GB06_TEXT5, GB06_TEXT6"
Gcon_Connect.Execute strNewVersione

strNewVersione = ""
strNewVersione = strNewVersione + "insert into GB07_CORPODOC "
strNewVersione = strNewVersione + " ( "
strNewVersione = strNewVersione + " GB07_ID_GB06, GB07_CODART_MG66, GB07_OPZIONE_MG5E , "
strNewVersione = strNewVersione + " GB07_QTA, GB07_PREZZO, GB07_SCCORPO, GB07_SC5, "
strNewVersione = strNewVersione + " GB07_SC4, GB07_SC3, GB07_SC2, GB07_SC1, "
strNewVersione = strNewVersione + " GB07_SC6, GB07_SCPIEDE, GB07_IMPORTO, GB07_DESCART, "
strNewVersione = strNewVersione + " GB07_MAGGIORAZIONE, GB07_SEQ, GB07_RAG, "
strNewVersione = strNewVersione + " GB07_ALT , GB07_FLPOSA, GB07_IMG, GB07_PATHIMG, GB07_TIPOCF_CG44, GB07_CLIFOR_CG44, GB07_COSTO, GB07_IMPPROVV,GB07_PERCPROVV  "
strNewVersione = strNewVersione + " ) "
strNewVersione = strNewVersione + " select "
strNewVersione = strNewVersione + " (select max(GB06_ID) as ID from GB06_TESTADOC where GB06_PROG like '" & CurUserName & "%'), GB07_CODART_MG66, GB07_OPZIONE_MG5E, "
strNewVersione = strNewVersione + " GB07_QTA, GB07_PREZZO_NEW, GB07_SCCORPO, GB07_SC5_NEW, "
strNewVersione = strNewVersione + " GB07_SC4_NEW, GB07_SC3_NEW, GB07_SC2_NEW, GB07_SC1_NEW, "
strNewVersione = strNewVersione + " GB07_SC6, GB07_SCPIEDE, GB07_IMPORTO_NEW, GB07_DESCART, "
strNewVersione = strNewVersione + " GB07_MAGGIORAZIONE, GB07_SEQ, GB07_RAG, "
strNewVersione = strNewVersione + " GB07_ALT , GB07_FLPOSA, GB07_IMG, GB07_PATHIMG, GB07_TIPOCF_CG44, GB07_CLIFOR_CG44, GB07_COSTO, GB07_IMPPROVV,GB07_PERCPROVV "
strNewVersione = strNewVersione + " from " & TABLENAME & " where gb07_id_gb06 = " & TXT_GB06_ID.Text
Gcon_Connect.Execute strNewVersione

strTesto = "Variazione Offerta __ <Da completare>"
strInserisciNota = "INSERT INTO GB08_NOTEOFFERTA (GB08_DATA, GB08_TESTONOTA, GB08_ID_GB06, GB08_OPERATORE) "
strInserisciNota = strInserisciNota & " VALUES "
strInserisciNota = strInserisciNota & "          ('" & Month(Now()) & "/" & Day(Now()) & "/" & Year(Now()) & "', '" & strTesto & "', " & GetValFromQuery("select max(GB06_ID) as ID from GB06_TESTADOC where GB06_PROG like '" & CurUserName & "%'", 0, Gcon_Connect) & ",'" & ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Codice & "' ) "
Gcon_Connect.Execute strInserisciNota
If MsgBox("Vuoi Aprire la nuova Versione ?", vbYesNo) = vbYes Then
      TXT_GB06_ID.Text = GetValFromQuery("select max(GB06_ID) as ID from GB06_TESTADOC where GB06_PROG like '" & CurUserName & "%'", 0, Gcon_Connect)
End If
End Sub


Private Sub cmdNewVersion_ButtonClick()

Dim strNewVersione As String
Dim CurUserName As String
CurUserName = ActualComputerName & ActualUserName
  
ProgInsertTestata = CurUserName & Year(Now) & String(2 - Len(Month(Now)), "0") & Month(Now) & String(2 - Len(Day(Now)), "0") & Day(Now) & String(2 - Len(Hour(Time)), "0") & Hour(Time) & String(2 - Len(Minute(Time)), "0") & Minute(Time) & String(2 - Len(Second(Time)), "0") & Second(Time) & Right(Format(Timer, "#0.00"), 2)

If MsgBox("Confermi la Creazione di una nuova Versione " & TXT_GB06_NOMEOFFERTA.Text & " ?", vbYesNo) = vbNo Then
      Exit Sub
End If


strNewVersione = " insert into GB06_TESTADOC"
strNewVersione = strNewVersione + "("
strNewVersione = strNewVersione + " GB06_PROG,"
strNewVersione = strNewVersione + " GB06_DITTA_CG18,"
strNewVersione = strNewVersione + " GB06_TIPODOC,"
strNewVersione = strNewVersione + " GB06_DATA,"
strNewVersione = strNewVersione + " GB06_GRUPPOCRE,"
strNewVersione = strNewVersione + " GB06_UTENTECRE,"
strNewVersione = strNewVersione + " GB06_TIPOCF_CG44,"
strNewVersione = strNewVersione + " GB06_CLIFOR_CG44,"
strNewVersione = strNewVersione + " GB06_BARCODE_LL01,"
strNewVersione = strNewVersione + " GB06_CODDESTIN_MG22,"
strNewVersione = strNewVersione + " GB06_CODPAG_CG62,"
strNewVersione = strNewVersione + " GB06_VETTORE_MG14,"
strNewVersione = strNewVersione + " GB06_NUMREG_CO99,"
strNewVersione = strNewVersione + " GB06_NUMDOC,"
strNewVersione = strNewVersione + " GB06_NREV,"
strNewVersione = strNewVersione + " GB06_NVERS,"
strNewVersione = strNewVersione + " GB06_DTDOC,"
strNewVersione = strNewVersione + " GB06_CODCOMM,"
strNewVersione = strNewVersione + " GB06_DTCHIUSURA,"
strNewVersione = strNewVersione + " GB06_PERCCHIUSURA,"
strNewVersione = strNewVersione + " GB06_RESPONSABILE,"
strNewVersione = strNewVersione + " GB06_PERCRIBGARA,"
strNewVersione = strNewVersione + " GB06_TIPOOFFERTA,"

strNewVersione = strNewVersione + " GB06_AGENTE ,"
strNewVersione = strNewVersione + " GB06_TIPOAREA,"
strNewVersione = strNewVersione + " GB06_STATODOC,"
strNewVersione = strNewVersione + " GB06_BUDGET,"
strNewVersione = strNewVersione + " GB06_FORECAST,"
strNewVersione = strNewVersione + " GB06_CONSUNTIVO,"
strNewVersione = strNewVersione + " GB06_NOMEOFFERTA,"
strNewVersione = strNewVersione + " GB06_DTULTMOD,"
strNewVersione = strNewVersione + " GB06_CIG,"
strNewVersione = strNewVersione + " GB06_CUP,"
strNewVersione = strNewVersione + " GB06_ALLACA,"
strNewVersione = strNewVersione + " GB06_TEXT1,"
strNewVersione = strNewVersione + " GB06_TEXT2,"
strNewVersione = strNewVersione + " GB06_TEXT3,"
strNewVersione = strNewVersione + " GB06_TEXT4,"
strNewVersione = strNewVersione + " GB06_TEXT5,"
strNewVersione = strNewVersione + " GB06_TEXT6"
strNewVersione = strNewVersione + ")"
strNewVersione = strNewVersione + " select"
strNewVersione = strNewVersione + " '" & ProgInsertTestata & "', GB06_DITTA_CG18, GB06_TIPODOC, GB06_DATA, GB06_GRUPPOCRE, GB06_UTENTECRE, GB06_TIPOCF_CG44,"
strNewVersione = strNewVersione + "                         GB06_CLIFOR_CG44, GB06_BARCODE_LL01, GB06_CODDESTIN_MG22, GB06_CODPAG_CG62, GB06_VETTORE_MG14, GB06_NUMREG_CO99, GB06_NUMDOC,"
strNewVersione = strNewVersione + "                         GB06_NREV, max(GB06_NVERS+1) as GB06_NVERS, GB06_DTDOC, GB06_CODCOMM, GB06_DTCHIUSURA, GB06_PERCCHIUSURA,GB06_RESPONSABILE, GB06_PERCRIBGARA, GB06_TIPOOFFERTA,"
strNewVersione = strNewVersione + "                          GB06_AGENTE , GB06_TIPOAREA, GB06_STATODOC, GB06_BUDGET, GB06_FORECAST, GB06_CONSUNTIVO, GB06_NOMEOFFERTA, { fn NOW() }, GB06_CIG, GB06_CUP, GB06_ALLACA, GB06_TEXT1, GB06_TEXT2, GB06_TEXT3, GB06_TEXT4, GB06_TEXT5, GB06_TEXT6"
strNewVersione = strNewVersione + " from GB06_TESTADOC where GB06_ID = " & TXT_GB06_ID.Text
strNewVersione = strNewVersione + " group by GB06_PROG, GB06_DITTA_CG18, GB06_TIPODOC, GB06_DATA, GB06_GRUPPOCRE, GB06_UTENTECRE, GB06_TIPOCF_CG44,"
strNewVersione = strNewVersione + "   GB06_CLIFOR_CG44, GB06_BARCODE_LL01, GB06_CODDESTIN_MG22, GB06_CODPAG_CG62, GB06_VETTORE_MG14, GB06_NUMREG_CO99, GB06_NUMDOC,"
strNewVersione = strNewVersione + "  GB06_NREV, GB06_DTDOC, GB06_CODCOMM, GB06_DTCHIUSURA, GB06_PERCCHIUSURA,GB06_RESPONSABILE, GB06_PERCRIBGARA, GB06_TIPOOFFERTA,"
strNewVersione = strNewVersione + "   GB06_AGENTE , GB06_TIPOAREA, GB06_STATODOC, GB06_BUDGET, GB06_FORECAST, GB06_CONSUNTIVO, GB06_NOMEOFFERTA, GB06_DTULTMOD, GB06_CIG, GB06_CUP, GB06_ALLACA, GB06_TEXT1, GB06_TEXT2, GB06_TEXT3, GB06_TEXT4, GB06_TEXT5, GB06_TEXT6"
Gcon_Connect.Execute strNewVersione

strNewVersione = ""
strNewVersione = strNewVersione + "insert into GB07_CORPODOC "
strNewVersione = strNewVersione + " ( "
strNewVersione = strNewVersione + " GB07_ID_GB06, GB07_CODART_MG66, GB07_OPZIONE_MG5E , "
strNewVersione = strNewVersione + " GB07_QTA, GB07_PREZZO, GB07_SCCORPO, GB07_SC5, "
strNewVersione = strNewVersione + " GB07_SC4, GB07_SC3, GB07_SC2, GB07_SC1, "
strNewVersione = strNewVersione + " GB07_SC6, GB07_SCPIEDE, GB07_IMPORTO, GB07_DESCART, "
strNewVersione = strNewVersione + " GB07_MAGGIORAZIONE, GB07_SEQ, GB07_RAG, "
strNewVersione = strNewVersione + " GB07_ALT , GB07_FLPOSA, GB07_IMG, GB07_PATHIMG, GB07_TIPOCF_CG44, GB07_CLIFOR_CG44,  GB07_COSTO, GB07_IMPPROVV,GB07_PERCPROVV  "
strNewVersione = strNewVersione + " ) "
strNewVersione = strNewVersione + " select "
strNewVersione = strNewVersione + " (select max(GB06_ID) as ID from GB06_TESTADOC where GB06_PROG like '" & CurUserName & "%'), GB07_CODART_MG66, GB07_OPZIONE_MG5E, "
strNewVersione = strNewVersione + " GB07_QTA, GB07_PREZZO, GB07_SCCORPO, GB07_SC5, "
strNewVersione = strNewVersione + " GB07_SC4, GB07_SC3, GB07_SC2, GB07_SC1, "
strNewVersione = strNewVersione + " GB07_SC6, GB07_SCPIEDE, GB07_IMPORTO, GB07_DESCART, "
strNewVersione = strNewVersione + " GB07_MAGGIORAZIONE, GB07_SEQ, GB07_RAG, "
strNewVersione = strNewVersione + " GB07_ALT , GB07_FLPOSA, GB07_IMG, GB07_PATHIMG, GB07_TIPOCF_CG44, GB07_CLIFOR_CG44, GB07_COSTO, GB07_IMPPROVV,GB07_PERCPROVV  "
strNewVersione = strNewVersione + " from GB07_CORPODOC where gb07_id_gb06 = " & TXT_GB06_ID.Text
Gcon_Connect.Execute strNewVersione








If MsgBox("Vuoi Aprire la nuova Versione ?", vbYesNo) = vbYes Then
      TXT_GB06_ID.Text = GetValFromQuery("select max(GB06_ID) as ID from GB06_TESTADOC where GB06_PROG like '" & CurUserName & "%'", 0, Gcon_Connect)
End If


End Sub


Function GetValFromQuery(strSQL As String, indNumCampo As Integer, ByRef rdbConn As ADODB.Connection)
    Dim strConnection As String
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    Dim cn As ADODB.Connection
    strConnection = rdbConn.ConnectionString
    Set cn = New ADODB.Connection
    cn.Open strConnection
   
    Set oRs = cn.Execute(strSQL)
    If oRs.EOF = False Then
        GetValFromQuery = NVL(Trim(oRs(indNumCampo)), "")
    Else
        GetValFromQuery = ""
    End If
    
    oRs.Close
    Set oRs = Nothing
    cn.Close
    Set cn = Nothing
    
End Function




Public Function CalcolaConsegna(articolo As String, Disponibilita As Double, Ordinato As Double) As Date

Dim LeadTime As Double
Dim strLeadTime As String
Dim DateConsegna As Date
Dim orsLeadTime As ADODB.Recordset

DateConsegna = Now


strLeadTime = " SELECT        MG67_TEMPOFIXPROD " & _
" From MG67_SCORTE " & _
" WHERE        (MG67_DITTA_CG18 = " & CodiceDitta & ") AND (MG67_CODART_MG66 = '" & Trim(articolo) & "') "

Set orsLeadTime = New ADODB.Recordset
orsLeadTime.Open strLeadTime, Gcon_Connect, adOpenKeyset, adLockReadOnly
  
If Not orsLeadTime.EOF And Not orsLeadTime.BOF Then
LeadTime = NVL(orsLeadTime("MG67_TEMPOFIXPROD"), 0)
End If




Select Case Disponibilita
    
    Case Is = 0
        DateConsegna = DateAdd("D", LeadTime, DateConsegna)
    Case Is > 0
        If Ordinato < Disponibilita Then
        DateConsegna = DateAdd("D", 7, DateConsegna)
        Else
        'duplica righe
       
        End If

Case Else

End Select




CalcolaConsegna = Format(DateConsegna, "dd/MM/yyyy")
End Function

Public Function CleanUpArray(Arr As Variant) As Variant
Dim aString()
Dim aString2
Dim x
Dim j As Integer

x = Null

For Each x In Arr
    If Not x & "" = "" Then
        ReDim Preserve aString(j + 1)
        aString(j) = x
        j = j + 1
    End If
Next
CleanUpArray = aString
'Debug.Print UBound(aString); " Records"
'For Each x In aString
'  Debug.Print x
'Next
End Function





Private Function ShowOpen(StartDir As String, DlgTitle As String, FileName As String) As String
  'Set the structure size
  OFName.lStructSize = Len(OFName)
  'Set the owner window
  OFName.hwndOwner = Me.hwnd
  'Set the application's instance
  OFName.hInstance = App.hInstance
  'Set the filet
  OFName.lpstrFilter = "Tutti i files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
  'Create a buffer
  OFName.lpstrFile = Space$(254)
  'Set the maximum number of chars
  OFName.nMaxFile = 255
  'Create a buffer
  OFName.lpstrFileTitle = Space$(254)
  'Set the maximum number of chars
  OFName.nMaxFileTitle = 255
  'Set the initial directory
  OFName.lpstrInitialDir = StartDir
  'Set the dialog title
  OFName.lpstrTitle = DlgTitle
  'no extra flags
  OFName.flags = 0
  
  'Show the 'Open File'-dialog
  If GetOpenFileName(OFName) Then
      ShowOpen = Trim$(OFName.lpstrFile)
  Else
      ShowOpen = ""
  End If
      
End Function

Private Sub cmdUltimiprezzi_Click()
Dim Opzione As String
'opzione = "RAL       6011      "
Set Cls_ConnectMagazzino.ActiveInterface = ActiveInterface
        
        Cls_ConnectMagazzino.Left = 10
        Cls_ConnectMagazzino.Top = 1000
        
        Call Cls_ConnectMagazzino.InterrogazioneUltimiPrezziRitroso(RTrimN(NVL(TXT_GB07_CODART_MG66.Text, "")), _
                                                                                     , _
                                                                                     CDecN(0) + 1, _
                                                                                     RTrimN(TXT_GB06_CLIFOR_CG44.Text), _
                                                                                     "00", _
                                                                                     "00", _
                                                                                     , , , , , _
                                                                                     "EURO", _
                                                                                     True, True, True, True, True)
                                                                
 
        ActiveInterface.IsActive = True
        Set Cls_ConnectMagazzino.ActiveInterface = Nothing
        Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
        'Call InitializeScript
End Sub

Private Sub Command1_Click()

'filePath = "c:\"
'    If Dir(filePath) <> "" Then
'        Kill filePath
'    End If
'    OutFile = FreeFile()
'    Open filePath For Output As OutFile
'    Print #OutFile, "#+TITLE:     Weekly Report"
'


End Sub

Private Sub FME_BANCO_AfterDelete(ByVal fenm_operationresult As FWBO_VIRTUALFRAME.EnumOperationResult)
 Call RicalcolaImportoTotale
End Sub

Private Sub FME_BANCO_AfterUpdate(ByVal fenm_operationresult As FWBO_VIRTUALFRAME.EnumOperationResult)
   Call RicalcolaImportoTotale
  
  'solamente per condizione pagamento diverso da contanti
  Select Case TXT_GB06_CODPAG_CG62.Text
  Case CODPAGCONTANTI, CODPAGBANCOMAT, CODPAGCARTACR, CODPAGASSEGNI ', CODPAGCONTSPEC
    Pbol_BloccoFido = False
    Exit Sub
  End Select
  
  'Solamente se documento diverso da scontrino
  If TipoDocumento = 2 Or TipoDocumento = 6 Then
    Pbol_BloccoFido = False
    Exit Sub
  End If
  
  Dim Ret As Boolean
  
'  Ret = ControlloFidoResiduo(0, 0)
  
'  If Not (ret) Then
'    If TipoDocumento = 5 Then
'      If ResiduoFido < 0 Then
'        If (ResiduoFido + CDbl(TXT_TOTALEDOC.Text)) < 0 Then
'            MsgBox "Attenzione, cliente fuori fido di " & Abs(Round(ResiduoFido + CDbl(TXT_TOTALEDOC.Text), 2))
'        End If
'      End If
'    Else
'      If CDbl(TXT_TOTALEDOC.Text) > ResiduoFido And Not (ClienteNoFido) Then
'        If ResiduoFido > 0 Then
'          MsgBox "Attenzione, cliente fuori fido di " & Round(CDbl(TXT_TOTALEDOC.Text) - ResiduoFido, 2)
'        Else
'          MsgBox "Attenzione, cliente fuori fido di " & Round(CDbl(TXT_TOTALEDOC.Text) + Abs(ResiduoFido), 2)
'        End If
'        'Abs(ResiduoFido) + Abs(CDbl(TXT_TOTALEDOC.Text))
'      End If
'    End If
'  End If
End Sub

Private Sub FME_BANCO_BeforeAddNew(fbol_Cancel As Boolean)
  TXT_GB07_CODART_MG66.SetFocus
End Sub

Private Sub FME_BANCO_BeforeUpdate(fbol_Cancel As Boolean)
    On Error Resume Next
    Dim Ret As Boolean
    
    'If TXT_GB07_CODART_MG66.IsValid Then
        rstCorpoBanco("GB07_ID_GB06") = IDGB06
        
        If Not (TXT_GB07_CODART_MG66.IsValid) Then
'        If isLettoreMomoria Then
'
'   '         TxtLog.Text = TxtLog.Text & vbCrLf & TXT_GB07_CODART_MG66.Text & " - Non Riconosciuto"
'            TotRigheFileScartati = TotRigheFileScartati + 1
'        Else
'
'            MsgBox "Articolo non valido"
'
'        End If
          fbol_Cancel = True
        End If
    ' End If
'    Ret = ControlloFidoResiduo(NVL(rstCorpoBanco("GB07_ID"), 0), CDbl(NVL(TXT_GB07_IMPORTO.Text, 0)))
    
End Sub

Private Sub Form_Initialize()
  On Error Resume Next
  
  FormIsActive = False
  
  Err.Clear
End Sub

Private Sub Form_Load()
  On Error GoTo ErrTrap
  
  'New Error management class
  Set Gcls_Log = New CLSFW_SrvLog
  
  'positioning and dimensioning form
  With Me
    If NVL(ActiveInterface.Left, 0) = 0 Then
       .Left = 100
    Else
       .Left = ActiveInterface.Left
    End If
    If NVL(ActiveInterface.Top, 0) = 0 Then
       .Top = 100
    Else
       .Top = ActiveInterface.Top
    End If
    .Width = 12210
    .Height = 7935
  End With

  
  Exit Sub
ErrTrap:
  Select Case VisualizzaErrore("Form_Load")
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
  End Select
End Sub

Private Sub Form_Activate()
  On Error GoTo ErrTrap
  
  SyncNavigator
  
  'Semafor
  If FormIsActive Then
      Exit Sub
  End If
  
  'Opening connection
  Gstr_Connect = ActiveInterface.ClsGlobal.Gcls_LibConnect.GetExtendedProperties
  Set Gcls_Connect = New CLSFW_SetConnect
  Set Gcon_Connect = Gcls_Connect.Gpr_GetConnect
  With Gcon_Connect
    .ConnectionString = Gstr_Connect
    .Open
  End With
  
  ' TeamSystem components use this connection for decode
  ' with this instruction I make a reference for them
  Set ActiveInterface.Connection = Gcon_Connect
    
  CodiceDitta = ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta

  Set Cls_LookupMagazzino = New MGBO_LOOKUPDECODE.CLSMG_LOOKUP
  Set Cls_LookupMagazzino.ActiveInterface = ActiveInterface
  Set Cls_LookupCommon = New COBO_LOOKUPDECODE.CLSCO_LOOKUP
  Set Cls_LookupCommon.ActiveInterface = ActiveInterface

  Set Cls_DecodeMagazzino = New MGBO_LOOKUPDECODE.CLSMG_DECODE
  Set Cls_DecodeMagazzino.ActiveInterface = ActiveInterface
  Set Cls_DecodeCommon = New COBO_LOOKUPDECODE.CLSCO_DECODE
  Set Cls_DecodeCommon.ActiveInterface = ActiveInterface

  Set Cls_ConnectMagazzino = New MGBO_LOOKUPDECODE.CLSMG_CONNECT
  Set Cls_ConnectCommon = New COBO_LOOKUPDECODE.CLSCO_CONNECT
  Set Cls_ConnectCommon = New COBO_LOOKUPDECODE.CLSCO_CONNECT
  
  Set Cls_LookupLotti = New LTBO_LOOKUPDECODE.CLSLT_LOOKUP
  Set Cls_LookupLotti.ActiveInterface = ActiveInterface
  Set Cls_ConnectLotti = New LTBO_LOOKUPDECODE.CLSLT_CONNECT
  
  'layout and script object initialization
  InizializzaScriptELayout
  
  If ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Username = "TeamSa" Then
  
  MsgBox "Utente TeamSa non abilitato a inserire offerte", vbCritical, "Gestione Offerte"
  Exit Sub
  End If
  'Lettura dei parametri generali
  If Not LeggiParametri Then
    MsgBox "Errore nel file parametri"
    Unload Me
  End If
  
  'Imposta virtualframe vuoto
'  Call ImpostaVirtualFrame
  TXT_ANNOINSERIMENTO.Text = CDate(GB05_ANNOINSERIMENTO)
  'Inizializza tutte le griglie
  Call InitQGrid
  QtaRead = 1
   
  'Abilito la gestione campi
  TXT_GB06_CLIFOR_CG44.IsGestione = True
 ' TXT_GB06_CODDESTIN_MG22.IsGestione = True
 ' TXT_GB06_VETTORE_MG14.IsGestione = True
  TXT_GB07_CODART_MG66.IsGestione = True
  TXT_DESART.IsGestione = True
  
  'Initialize Resize form
  Set TMS_RESIZEFORM.ActiveInterface = ActiveInterface
  TMS_RESIZEFORM.Initialize
  
  FormIsActive = True
  isLettoreMomoria = False
  
  cmdScarico.Enabled = True
  
  
    CMB_TIPOVARIAZIONE.AddItemData "Diminuzione prezzi percentuale", 0
    CMB_TIPOVARIAZIONE.AddItemData "Applica percentuale Sconto 1", 1
    CMB_TIPOVARIAZIONE.AddItemData "Applica percentuale Sconto 2", 2
    CMB_TIPOVARIAZIONE.AddItemData "Applica percentuale Sconto 3", 3
    CMB_TIPOVARIAZIONE.AddItemData "Applica percentuale Sconto 4", 4
    CMB_TIPOVARIAZIONE.AddItemData "Applica percentuale Sconto 5", 5
    CMB_TIPOVARIAZIONE.AddItemData "Applica percentuale Sconto 6", 6
    CMB_TIPOVARIAZIONE.AddItemData "Maggiorazione Prezzi precentuale", 7
    CMB_TIPOVARIAZIONE.AddItemData "Assegna Fornitore", 98
    CMB_TIPOVARIAZIONE.AddItemData "Ripristina Orginale", 99
    CMB_TIPOVARIAZIONE.AutoOpen = False
    CMB_TIPOVARIAZIONE.Text = 99
    
    
'    TMS_TIPOGRUPPO.AddItemData "Famiglia", 0
'    TMS_TIPOGRUPPO.AddItemData "Sotto Famiglia", 1
'    TMS_TIPOGRUPPO.AddItemData "Gruppo", 2
'    TMS_TIPOGRUPPO.AddItemData "Sotto Gruppo", 3
'    TMS_TIPOGRUPPO.AutoOpen = False
'    TMS_TIPOGRUPPO.Text = 0
  
  TabDocumenti.ActiveTab = 0
  cmdScarico.SetFocus
  
  TMS_SSTAB1.ActiveTab = 0
  
  Exit Sub
ErrTrap:
  Select Case VisualizzaErrore("Form_Activate")
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
  End Select
End Sub



'Leggi i parametri da tabella parametri
Function LeggiParametri() As Boolean
'
  
  On Error GoTo ErrTrap
  
  Dim rstParametri  As ADODB.Recordset
  
  strSQL = "SELECT  * FROM GB05_BLOCCHI "
  strSQL = strSQL & " WHERE GB05_DITTA_CG18 = " & CodiceDitta
  
  Set rstParametri = New ADODB.Recordset
  rstParametri.Open strSQL, Gcon_Connect, adOpenKeyset, adLockReadOnly
  
  If Not rstParametri.EOF And Not rstParametri.BOF Then
    LeggiParametri = True
    GB05_PWDFIDO = NVL(rstParametri("GB05_PWDFIDO"), "")
    GB05_ANNOINSERIMENTO = NVL(rstParametri("GB05_PWDFIDO"), "")
    GB05_PWDFIDCARD = NVL(rstParametri("GB05_PWDFIDCARD"), "")
    GB05_PWDSCRIGA = NVL(rstParametri("GB05_PWDSCRIGA"), "")
    GB05_PWDSCPIEDE = NVL(rstParametri("GB05_PWDSCPIEDE"), "")
    GB05_MAXSCRIGA = NVL(rstParametri("GB05_MAXSCRIGA"), "")
    GB05_MAXSCPIEDE = NVL(rstParametri("GB05_MAXSCPIEDE"), "")
    
    
    PARAM_DIRTEMP = NVL(rstParametri("GB05_DIRTEMP"), "")
    PARAM_DIRSCONTRINI = NVL(rstParametri("GB05_DIRSCONTRINI"), "")
    PARAM_DIRCOPIA = NVL(rstParametri("GB05_DIRCOPIA"), "")
    PARAM_NOMEFILE = NVL(rstParametri("GB05_NOMEFILE"), "")
    PARAM_EXESCONTR = NVL(rstParametri("GB05_EXESCONTR"), "")
    

  Else
    LeggiParametri = True
  End If
  
  rstParametri.Close
  Set rstParametri = Nothing


Exit Function
ErrTrap:
    Select Case VisualizzaErrore("LeggiParametri")
        Case vbAbort
            Exit Function
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select

  
End Function



Private Sub InizializzaScriptELayout()
  On Error GoTo ErrTrap
  'layout and script object initialization
  ActiveInterface.ActiveNavigator.ApplyPrsLayout
  ExecuteFormEvent "tsOpen"
  Set ActiveInterface.ClsGlobal.ApplicationObject = App

  Exit Sub
ErrTrap:
  Select Case VisualizzaErrore("InizializzaScriptELayout")
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
  End Select
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Dim UserObject As Variant
  On Error GoTo ErrTrap
  
'  Set LookUpMO = Nothing
'    Set DecodeMO = Nothing
  
'    Set PclsReport = Nothing
  
  'Destroy ActiveInterface reference
  ActiveInterface.ClsGlobal.RemoveCurrentInterface ActiveInterface
  
  'Destroy reference to layout and script class
  UserObject = ActiveInterface.ClsVoceMenu.Classe
  ActiveInterface.ActiveNavigator.ClsScript.TerminateByUserObject UserObject
  Set ActiveInterface.ActiveNavigator.ClsLayout = Nothing
  Set ActiveInterface.ActiveNavigator.ClsScript = Nothing
    
  Set Cls_LookupMagazzino.ActiveInterface = Nothing
  Set Cls_LookupMagazzino = Nothing
  Set Cls_LookupCommon.ActiveInterface = Nothing
  Set Cls_LookupCommon = Nothing
      
  Set Cls_DecodeMagazzino.ActiveInterface = Nothing
  Set Cls_DecodeMagazzino = Nothing
  Set Cls_DecodeCommon.ActiveInterface = Nothing
  Set Cls_DecodeCommon = Nothing
      
  Set Cls_ConnectMagazzino = Nothing
  Set Cls_ConnectCommon = Nothing
      
  If Not Cls_CalcPrezzi Is Nothing Then
      Set Cls_CalcPrezzi.ADOConnection = Nothing
      Set Cls_CalcPrezzi.ClsDittaCorrente = Nothing
      Set Cls_CalcPrezzi = Nothing
  End If
  
  If Not Cls_ControlloRischio Is Nothing Then
      Set Cls_ControlloRischio = Nothing
  End If
  
  If Not Gcls_CommandBloccoStatiArt Is Nothing Then
      Set Gcls_CommandBloccoStatiArt.ActiveConnection = Nothing
      Set Gcls_CommandBloccoStatiArt = Nothing
  End If
  
  
'  If Not objFido Is Nothing Then
'      Set objFido = Nothing
'  End If
  
  'Destroy reference to primary class
  Set ActiveClass = Nothing
  
  'Destroy working recorset
  Set rstCorpoBanco = Nothing
  Set Gcls_RecordS = Nothing
  
  'Destroy connection
  Set Gcls_Connect = Nothing
  Set Gcon_Connect = Nothing
  
  'Destroy ActiveInterface
  Set ActiveInterface.ClsGlobal.ActiveInterface = Nothing
  Set ActiveInterface.ClsGlobal.CallInterface = Nothing
  Set ActiveInterface.ActiveDll = Nothing
  Set ActiveInterface.ActiveNavigator.ActiveInterface = Nothing
  Set ActiveInterface.ActiveDll = Nothing
  Set ActiveInterface = Nothing
  
  'Destroy Error management class
  Set Gcls_Log = Nothing

  Exit Sub
ErrTrap:
  Select Case VisualizzaErrore("Form_UnLoad")
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error GoTo ErrTrap
  
  SyncNavigator
    
  'closing form event(for scripting)
  If Not Cancel Then
      ExecuteFormEvent ("tsClose")
      Cancel = ActiveInterface.ActiveNavigator.ClsScript.CancelEvent
  End If

  Exit Sub
ErrTrap:
  Select Case VisualizzaErrore("Form_QueryUnload")
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
  End Select
End Sub



Public Function VisualizzaErrore(ByVal SubOrFunctionName As String) As VbMsgBoxResult
  ' standard function for manageing errors
  If Gcls_Log Is Nothing Then
      VisualizzaErrore = vbAbort
      MsgBox "Classe Log non settata", vbOKOnly, Me.Caption
      Exit Function
  End If
  
  Set Gcls_Log.vbError = Err
  
  If Not (Gcon_Connect Is Nothing) Then
      Set Gcls_Log.ADOError = Gcon_Connect.Errors
  End If
  '
  'this method shows a standard error MsgBox
  '
  VisualizzaErrore = Gcls_Log.ShowError(App.Title, Me.Caption, SubOrFunctionName)
End Function


Public Function VisualizzaWarning(ByVal SubOrFunctionName As String, ByVal DettaglioWarning As String, Optional ByVal ShowWarningMode As VbMsgBoxStyle = vbCritical + vbOKOnly) As VbMsgBoxResult
    If SubOrFunctionName = "" Then
        VisualizzaWarning = MsgBox(DettaglioWarning, ShowWarningMode, "Emissione diretta documenti")
    Else
        VisualizzaWarning = MsgBox(DettaglioWarning & vbCr & "Funzione: " & SubOrFunctionName, ShowWarningMode, "Gestione documenti")
    End If
End Function


' scripting
Private Function ExecuteFormEvent(ByVal Mode As Variant)
  Dim ClsScript As FWUO_TMSDEVELOP.CLSFW_PRSVBSCRIPT

  On Error GoTo ErrTrap

  Select Case Mode
     Case "tsOpen"
        ActiveInterface.ActiveNavigator.InitializeScript
        Set ClsScript = ActiveInterface.ActiveNavigator.ClsScript
        If Not ClsScript Is Nothing Then
           ClsScript.ExecuteObjectEvent Me.Name, FWUO_TMSDEVELOP.tsForm, FWUO_TMSDEVELOP.tsOpenForm
        End If
     Case "tsClose"
        Set ClsScript = ActiveInterface.ActiveNavigator.ClsScript
        If Not ClsScript Is Nothing Then
           ClsScript.ExecuteObjectEvent Me.Name, FWUO_TMSDEVELOP.tsForm, FWUO_TMSDEVELOP.tsCloseForm
        End If
  End Select

  Exit Function
ErrTrap:
  Select Case VisualizzaErrore("ExecuteFormEvent")
    Case vbAbort
        Exit Function
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
  End Select
End Function

Private Sub Image2_Click()

End Sub

Private Sub MDIActiveX1_FormLoad()
'
'  'Variabili Long
'  Dim MyLeft      As Long
'  Dim MyTop       As Long
'
'  On Error GoTo ErrTrap
'
'  'positioning and dimensioning form
'  If NVL(ActiveInterface.Left, 0) = 0 Then
'     MyLeft = 100
'  Else
'     MyLeft = ActiveInterface.Left
'  End If
'  If NVL(ActiveInterface.Top, 0) = 0 Then
'     MyTop = 100
'  Else
'     MyTop = ActiveInterface.Top
'  End If
'  MDIActiveX1.Move MyLeft, MyTop, 9375, 18105
'  MDIActiveX1.WindowState = ActiveInterface.WindowState
'
'  Exit Sub
'ErrTrap:
'  Select Case VisualizzaErrore("MDIActiveX1_FormLoad")
'      Case vbAbort: Exit Sub
'      Case vbRetry: Resume
'      Case vbIgnore: Resume Next
'  End Select
'  Err.Clear
End Sub



Private Sub TMS_FLATBUTTON2_Click()

'Call CMD_SAVE_Click

If QGridDocumenti.DataSource Is Nothing Then
    Exit Sub
End If

If QGridDocumenti.DataSource.RecordCount = 0 Then
    Exit Sub
End If

QGridDocumenti.DataSource.MoveFirst
Do While Not QGridDocumenti.DataSource.EOF
      
   
    Call RecuperImmagineHyperMedia(NVL(TXT_GB07_CODART_MG66.Text, ""))
    TXT_GB07_CLIFOR_CG44.Text = RecuperaFornitorePref(NVL(TXT_GB07_CODART_MG66.Text, ""))
    'QGridDocumenti.DataSource("GB07_CLIFOR_CG44").value = "" ' = RecuperaFornitorePref(NVL(TXT_GB07_CODART_MG66.Text, ""))
    Call RecuperaCosto
    QGridDocumenti.DataSource("GB07_PERCPROVV").value = RecuperaProvvigione(QGridDocumenti.DataSource("GB07_PREZZO").value, QGridDocumenti.DataSource("GB07_IMPORTO").value, NVL(QGridDocumenti.DataSource("MG66_FAM_MG53").value, ""), NVL(QGridDocumenti.DataSource("MG66_SFAM_MG54").value, ""), TXT_GB06_AGENTE.Text, NVL(QGridDocumenti.DataSource("GB07_SC1").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC2").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC3").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC4").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC5").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC6").value, 0))
    FME_BANCO.Update True
    QGridDocumenti.DataSource.MoveNext
Loop

Call CMD_SAVE_Click
End Sub

'
Private Sub TMS_RESIZEFORM_AfterResize()
QGridDocumenti.Width = Me.Width - 1500
If Me.Height > 7000 Then
QGridDocumenti.Height = Me.Height - 7500
End If
Label3.Caption = CStr(Me.Width) & "x" & CStr(Me.Height)
End Sub

Private Sub TMS_RESIZEFORM_AfterAutoInitialize()

  With TMS_RESIZEFORM
'    .AddControl Picture1, tsAnchorTop And tsAnchorRight And tsAnchorleft And tsAnchorBottom
''    .AddControl GridNavDocumenti, tsAnchorTop
''   .AddControl TMS_SSTAB1, tsAnchorTop + tsAnchorRight + tsAnchorleft + tsAnchorBottom
    .AddControl FrameGriglia, tsAnchorTop Or tsAnchorleft Or tsAnchorBottom Or tsAnchorRight
    .AddControl TabDocumenti, tsAnchorTop Or tsAnchorleft Or tsAnchorBottom Or tsAnchorRight
 '   .AddControl FrameCorpo, tsAnchorTop Or tsAnchorleft Or tsAnchorRight Or tsAnchorBottom
    
'    .FrameResize TabDocumenti
'    .AddControl FrameCorpo, tsAnchorTop Or tsAnchorleft Or tsAnchorRight 'Or tsAnchorBottom
    .AddControl QGridDocumenti, tsAnchorTop Or tsAnchorRight Or tsAnchorleft Or tsAnchorBottom
    .AddControl TMS_SSTAB1, tsAnchorTop Or tsAnchorleft Or tsAnchorBottom Or tsAnchorRight
'   .AddControl FrameGenera, tsAnchorBottom Or tsAnchorTop Or tsAnchorRight Or tsAnchorleft
'   .AddControl Label5, tsAnchorBottom
   
 '  .AddControl PictureArticoli, tsAnchorTop Or tsAnchorleft Or tsAnchorBottom Or tsAnchorRight
   ''    .AddControl QGRID_SIMULAZIONE, tsAnchorTop + tsAnchorRight + tsAnchorleft + tsAnchorBottom
''    .AddControl TMS_GRIDDOC, tsAnchorTop + tsAnchorRight + tsAnchorleft + tsAnchorBottom
''    .AddControl frmOpticom, tsAnchorTop + tsAnchorRight + tsAnchorleft + tsAnchorBottom
'  .AddControl FrameGenera, tsAnchorleft
'
'
''
  End With

End Sub

Private Function SyncNavigator()
  On Error Resume Next
  Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
  If Not ActiveInterface.ActiveFrame Is Nothing Then
      ActiveInterface.ActiveNavigator.SetStatus (ActiveInterface.ActiveFrame.Status)
  End If
  ActiveInterface.ActiveNavigator.Refresh
  Err.Clear
End Function


Private Sub cmdCarico_ButtonClick()
  TipoDocumento = 1
  lblTipoDocumento.Caption = "Carico Merce da Deposito Centrale"
  Call InserisciTestata
End Sub

Private Sub cmdScontrino_ButtonClick()
  TipoDocumento = 2
  lblTipoDocumento.Caption = "SCONTRINO"
  Call InserisciTestata
End Sub

Private Sub cmdNotaCrFatt_ButtonClick()
  TipoDocumento = 3
  lblTipoDocumento.Caption = "NOTA CREDITO DA FATTURA"
  Call InserisciTestata
End Sub

Private Sub cmdResoDDT_ButtonClick()
  TipoDocumento = 4
  lblTipoDocumento.Caption = "RESO DA DDT"
  Call InserisciTestata
End Sub

Private Sub cmdScarico_ButtonClick()


  TipoDocumento = 5
  lblTipoDocumento.Caption = "Creazione Offerta Cliente"
  
'  DEPCENTRALE = TXT_DEP_DA.Text
'  DEPCENTRO = TXT_DEPCOLL.Text
  
  Call InserisciTestata
End Sub

Private Sub cmdReso_ButtonClick()
  TipoDocumento = 6
  lblTipoDocumento.Caption = "RESO MERCE DA CENTRO ESTERNO"
'  DEPCENTRALE = TXT_DEPCOLL.Text
'  DEPCENTRO = TXT_DEP_DA.Text
  Call InserisciTestata
End Sub


Function ReadIni() As Boolean
 On Error GoTo ErrTrap
'leggi file ini
    Dim varparm(30) As String
    Dim ris As Boolean
    Dim ind As Integer
    'Apro il file con i parametri per collegarmi al database
    Open "C:\gabe\inizializza.ini" For Input As 1
    'Carico i parametri in maniera sequenziale
    ind = 0
    Do While Not EOF(1)
        Input #1, varparm(ind)
        ind = ind + 1
    Loop
  
    Close #1
    ReadIni = True
ErrTrap:
 
ReadIni = False
End Function

Private Sub ModificaTestata()
On Error GoTo ErrTrap
  
  If ReadIni = True Then
  
  Else
  
  End If
  
  
  

  Dim Ret As Boolean
  FRMGB_PASSWORD.PasswordCorretta = False
  
  

  Call LeggiIDGB06

'Nascondo campi in base al tipo documento che devo gestire
  Ret = BloccoPerDocumento
  If Not (Ret) Then
    MsgBox Errore
    Exit Sub
  End If
  
'Apro il virtual frame
  Call DistruggiFramework
  Call ImpostaVirtualFrame
  
  Set QGridDocumenti.DataSource = Nothing
  Set QGridDocumenti.DataSource = rstCorpoBanco
  QGridDocumenti.FullExpand
  
  
 
  
'Apro folder gestione documento
  TabDocumenti.ActiveTab = 1
  TXT_GB06_STATODOC.Text = "00"
  Call StatoCampi
  
'  Call RicalcolaImporto
  
  
  Select Case TipoDocumento
  Case 1
    TXT_GB06_CLIFOR_CG44.SetFocus
  Case 2
    TXT_GB07_CODART_MG66.SetFocus
  Case 3
    TXT_GB06_CLIFOR_CG44.SetFocus
  Case 4
    TXT_GB06_CLIFOR_CG44.SetFocus
  Case 5
    TXT_GB06_CLIFOR_CG44.SetFocus
  Case 6
    TXT_GB07_CODART_MG66.SetFocus
  End Select
  
  Pbol_Generazione = False

  
  
  Exit Sub
ErrTrap:
  Select Case VisualizzaErrore("InserisciTestata")
      Case vbAbort: Exit Sub
      Case vbRetry: Resume
      Case vbIgnore: Resume Next
  End Select
  Err.Clear
End Sub

Private Sub InserisciTestata()

  On Error GoTo ErrTrap
  
  If ReadIni = True Then
  
  Else
  
  End If
  
  
  

  Dim Ret As Boolean
  FRMGB_PASSWORD.PasswordCorretta = False
  
  
'Genera riga in GB06
  Call ScriviGB06
  Call LeggiIDGB06

'Nascondo campi in base al tipo documento che devo gestire
  Ret = BloccoPerDocumento
  If Not (Ret) Then
    MsgBox Errore
    Exit Sub
  End If
  
'Apro il virtual frame
  Call DistruggiFramework
  Call ImpostaVirtualFrame
  Call CaricaTestiFissi
  
  Select Case ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Username
  Case "ALESSANDRA"
    TXT_GB06_NUMDOC.Text = GetNumeroOfferta("AC")
  Case "BARBARA"
    TXT_GB06_NUMDOC.Text = GetNumeroOfferta("BP")
  Case "LAURA"
    TXT_GB06_NUMDOC.Text = GetNumeroOfferta("LD")
  Case "PIERO"
    TXT_GB06_NUMDOC.Text = GetNumeroOfferta("PB")
  Case "STEFANO"
    TXT_GB06_NUMDOC.Text = GetNumeroOfferta("SM")
  Case "SUSANNA"
    TXT_GB06_NUMDOC.Text = GetNumeroOfferta("SS")
  Case "VALERIA"
    TXT_GB06_NUMDOC.Text = GetNumeroOfferta("VV")
  Case "FEDRICA"
    TXT_GB06_NUMDOC.Text = GetNumeroOfferta("FB")
  Case Else
    TXT_GB06_NUMDOC.Text = GetNumeroOfferta("AB")
  
  End Select
  
  
  
  Set QGridDocumenti.DataSource = Nothing
  Set QGridDocumenti.DataSource = rstCorpoBanco
  QGridDocumenti.FullExpand
  
  
  
  

  
'Apro folder gestione documento
  TabDocumenti.ActiveTab = 1
  TXT_GB06_RESPONSABILE.Text = NVL(TXT_GB06_PROPRIETARIO.Text, "")
  
'  Call RicalcolaImporto
  Call RicalcolaImportoTotale
  
  Select Case TipoDocumento
  Case 1
    TXT_GB06_NOMEOFFERTA.SetFocus
  Case 2
    TXT_GB06_NOMEOFFERTA.SetFocus
  Case 3
    TXT_GB06_NOMEOFFERTA.SetFocus
  Case 4
    TXT_GB06_NOMEOFFERTA.SetFocus
  Case 5
    TXT_GB06_NOMEOFFERTA.SetFocus
  Case 6
    TXT_GB06_NOMEOFFERTA.SetFocus
  End Select
  
  Pbol_Generazione = False

  
  
  Exit Sub
ErrTrap:
  Select Case VisualizzaErrore("InserisciTestata")
      Case vbAbort: Exit Sub
      Case vbRetry: Resume
      Case vbIgnore: Resume Next
  End Select
  Err.Clear
  

End Sub


Private Sub CaricaTestiFissi()
Dim str_sql1 As String
str_sql1 = " SELECT" & _
              "   GB05_TEXT1  " & _
              ",   GB05_TEXT2  " & _
              ",   GB05_TEXT3  " & _
              ",   GB05_TEXT4  " & _
              ",   GB05_TEXT5  " & _
              ",   GB05_TEXT6  " & _
              "    FROM        GB05_BLOCCHI "

  
  TXT_GB06_TEXT1.Text = GetValFromQuery(str_sql1, 0, Gcon_Connect)
  TXT_GB06_TEXT2.Text = GetValFromQuery(str_sql1, 1, Gcon_Connect)
  TXT_GB06_TEXT3.Text = GetValFromQuery(str_sql1, 2, Gcon_Connect)
  TXT_GB06_TEXT4.Text = GetValFromQuery(str_sql1, 3, Gcon_Connect)
  TXT_GB06_TEXT5.Text = GetValFromQuery(str_sql1, 4, Gcon_Connect)
  TXT_GB06_TEXT6.Text = GetValFromQuery(str_sql1, 5, Gcon_Connect)
End Sub


'Gestione dei blocchi
  '1 in apertura   blocco tutti i frame a parte la gestione del cliente
  '                sblocco i campi TXT_GB07_CODART_MG66, TXT_GB07_PREZZO,TXT_GB07_SCCORPO
  '2 Fattura/DDT
  '                nascondo FramePag (da visualizzare dopo aver inserito il cliente)
  '3 per Scontrino
  '                abilito i frame di inserimento FrameCorpo,FrameTotali, FrameGriglia, FramePag, FrameGenera
  '                imposto il cliente BRICOMATT
  '                nascondo il frame FrameDestVet
  '3 per nota credito/nc scontrino
  '                apro ricerca documenti per fattura emessa o scontrino
  '                inserisco tutti i record del documento selezionato nella GB07 e nei dati clienti/dest/vettori
  '                nascondo FramePag
  '                imposto a 20% lo sconto piede fisso
  '                blocco i campi TXT_GB07_CODART_MG66, TXT_GB07_PREZZO,TXT_GB07_SCCORPO
  '                blocco GridNavDocumenti
  '3 per reso da ddt
  '                apro ricerca documenti per DDT emesso
  '                inserisco tutti i record del documento selezionato nella GB07 e nei dati clienti/dest/vettori
  '                nascondo FramePag
  '                imposto a 20% lo sconto piede fisso
  '                blocco i campi TXT_GB07_CODART_MG66, TXT_GB07_PREZZO,TXT_GB07_SCCORPO
  '                blocco GridNavDocumenti

Private Function BloccoPerDocumento() As Boolean

  On Error GoTo ErrTrap

  BloccoPerDocumento = True

  '1 in apertura   blocco tutti i frame a parte la gestione del cliente
  '                sblocco i campi TXT_GB07_CODART_MG66, TXT_GB07_PREZZO,TXT_GB07_SCCORPO
  FrameCliente.Enable = False
 ' FrameDestVet.Visible = True
 ' FrameDestVet.Enable = False
  FrameCorpo.Enable = False
  
  FrameGriglia.Enable = False
  FrameRiferimenti.Enable = False

 ' FramePag.Enable = False
  FrameGenera.Enable = False
  
  TXT_GB07_CODART_MG66.Enabled = True
  
'  cmdContSpeciali.Visible = False
 ' cmdRimDiretta.Visible = False
'  cmdRimDiretta.Visible = True
  
  NumRegGenerato = ""
  
  Select Case TipoDocumento
  Case 1, 5
        '1 Fattura/DDT, 5 NOTA CREDITO/DDT RESO
        '                sblocco FrameCliente
        '                nascondo FramePag (da visualizzare dopo aver inserito il cliente)
        FrameCliente.Enable = True
       ' FramePag.Visible = False
        
       ' FrameDestVet.Enable = True
        FrameCorpo.Enable = True
        
        FrameGriglia.Enable = True
        FrameRiferimenti.Enable = True
       ' FramePag.Enable = True
        FrameGenera.Enable = True
'        cmdRimDiretta.Visible = True
'        If ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Gruppo = "BANCOTRAS" Then
'          TXT_GB06_VETTORE_MG14.Text = "29"
'        End If
        
  Case 2, 6
        '2 per Scontrino , 6 NC Scontrino
        '                abilito i frame di inserimento FrameCorpo,FrameTotali, FrameGriglia, FramePag, FrameGenera
        '                nascondo il frame FrameDestVet
        '                imposto il cliente BRICOMATT
        FrameCorpo.Enable = True
        
        FrameGriglia.Enable = True
        FrameRiferimenti.Enable = True
     '   FramePag.Enable = True
        FrameGenera.Enable = True
        
      '  FrameDestVet.Visible = False
'        cmdRimDiretta.Visible = False
        '
     
        TXT_GB06_CLIFOR_CG44.Text = CODCLIBRICOMATT
        Pbol_BloccoFido = False
        
  Case 3
        '3 per nota credito/nc scontrino
        '                apro ricerca documenti per fattura emessa o scontrino
        '                inserisco tutti i record del documento selezionato nella GB07 e nei dati clienti/dest/vettori
        '                nascondo FramePag e abilito gestione
        '                imposto a 20% lo sconto piede fisso
        '                blocco i campi TXT_GB07_CODART_MG66, TXT_GB07_PREZZO,TXT_GB07_SCCORPO
        '                blocco GridNavDocumenti
       ' TXT_RICERCANC.StartLookup
'        If NVL(TXT_RICERCANC.Text, "") <> "" Then
'          Call CaricaValoriDocOrigine(TXT_RICERCANC.Text)
'
'          If Pbol_BloccoFido Then
'            BloccoPerDocumento = False
'            Errore = "Cliente bloccato per fido"
'
'            Exit Function
'          End If
          
          FrameGriglia.Enable = True
          FrameRiferimenti.Enable = True
          FrameCorpo.Enable = True
          
          FrameGenera.Enable = True
      '    FramePag.Visible = False
          
          
          
          TXT_GB07_CODART_MG66.Enabled = False

          
'          GridNavDocumenti.Enabled = False
'        Else
'          BloccoPerDocumento = False
'          Errore = "Documento fattura/scontrino non caricato correttamente"
'        End If
        
  Case 4
        '4 per reso da ddt
        '                apro ricerca documenti per DDT emesso
        '                inserisco tutti i record del documento selezionato nella GB07 e nei dati clienti/dest/vettori
        '                nascondo FramePag
        '                imposto a 20% lo sconto piede fisso
        '                blocco i campi TXT_GB07_CODART_MG66, TXT_GB07_PREZZO,TXT_GB07_SCCORPO
        '                blocco GridNavDocumenti
  
'        TXT_RICERCARESI.StartLookup
'        If NVL(TXT_RICERCARESI.Text, "") <> "" Then
'          Call CaricaValoriDocOrigine(TXT_RICERCARESI.Text)
'
'          If Pbol_BloccoFido Then
'
'            BloccoPerDocumento = False
'            Errore = "Cliente bloccato per fido"
'
'            Exit Function
'          End If
'
'          FrameGriglia.Enable = True
'          FrameRiferimenti.Enable = True
'          FrameCorpo.Enable = True
'
'          FrameGenera.Enable = True
'          FramePag.Visible = False
'
'
'          TXT_GB07_CODART_MG66.Enabled = False
'
'
'        Else
'          BloccoPerDocumento = False
'          Errore = "Documento ddt non caricato correttamente"
'        End If
        
  End Select
  

Exit Function
ErrTrap:
    Select Case VisualizzaErrore("BloccoPerDocumento")
        Case vbAbort
            Exit Function
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select

  
End Function


'inserisco tutti i record del documento selezionato nella GB07 e nei dati clienti/dest/vettori
Private Sub CaricaValoriDocOrigine(Numreg As String)
  
  On Error GoTo ErrTrap
  
'Aggiorno GB06
  Dim MyRst       As ADODB.Recordset
  Dim Ret         As Boolean
  
  Pbol_BloccoFido = False
    
  strSQL = "SELECT * "
  strSQL = strSQL & "       FROM DO11_DOCTESTATA "
  strSQL = strSQL & " INNER JOIN DO14_DOCDATIACC "
  strSQL = strSQL & "         ON DO11_DITTA_CG18  = DO14_DITTA_CG18 "
  strSQL = strSQL & "        AND DO11_NUMREG_CO99 = DO14_NUMREG_CO99 "
  strSQL = strSQL & "      WHERE DO11_DITTA_CG18  = " & CodiceDitta
  strSQL = strSQL & "        AND DO11_NUMREG_CO99 = '" & Numreg & "'"
  
  Set MyRst = Gcon_Connect.Execute(strSQL, , adCmdText)
  If Not MyRst.EOF Then
  
    'Verifico fido del cliente
    Ret = CercoFidoCliente(MyRst("DO11_CLIFOR_CG44"))
    If Not Ret Then
      Set MyRst = Nothing
      Exit Sub
    End If
    
    strSQL = "UPDATE GB06_TESTADOC "
    strSQL = strSQL & " SET  GB06_TIPOCF_CG44    = 0 "
    strSQL = strSQL & "    , GB06_CLIFOR_CG44    = " & MyRst("DO11_CLIFOR_CG44")
    strSQL = strSQL & "    , GB06_CODDESTIN_MG22 = '" & MyRst("DO11_CLIFORDEST") & "'"
    strSQL = strSQL & "    , GB06_CODPAG_CG62    = '" & MyRst("DO11_CODPAG_CG62") & "'"
    strSQL = strSQL & "    , GB06_VETTORE_MG14   = '" & MyRst("DO14_VETTORE1_MG14") & "'"
    strSQL = strSQL & " WHERE GB06_ID   = " & IDGB06
    
    Gcon_Connect.Execute strSQL
    
    TXT_GB06_CLIFOR_CG44.Text = MyRst("DO11_CLIFOR_CG44")
  '  TXT_GB06_CODDESTIN_MG22.Text = MyRst("DO11_CLIFORDEST")
   ' TXT_GB06_VETTORE_MG14.Text = MyRst("DO14_VETTORE1_MG14")
    TXT_GB06_CODPAG_CG62.Text = CODPAGRD
  '  TXT_RIFERIMENTO.Text = MyRst("DO11_NUMDOC")
  '  TXT_DATARIF.Text = MyRst("DO11_DATADOC")
    
    TipoDocumentoRecuperato = MyRst("DO11_TIPODOC")
    
  End If
  
  strSQL = ""
  Set MyRst = Nothing
  
'Inserisco GB07
  strSQL = "INSERT INTO GB07_CORPODOC "
  strSQL = strSQL & " ( GB07_ID_GB06, GB07_CODART_MG66, GB07_QTA, GB07_PREZZO, GB07_SCCORPO, GB07_SCPIEDE, GB07_IMPORTO )"
  strSQL = strSQL & " SELECT   " & IDGB06
  strSQL = strSQL & "        , DO30_CODART_MG66, DO30_QTA1, DO30_PREZZO1, DO30_SCPER1, 20 "
  strSQL = strSQL & "        , ROUND((   ( DO30_PREZZO1 - DO30_PREZZO1 * DO30_SCPER1 / 100 ) "
  strSQL = strSQL & "                                   - ( DO30_PREZZO1 - DO30_PREZZO1 * DO30_SCPER1 / 100 ) * 20 / 100  "
  strSQL = strSQL & "                              ) * DO30_QTA1    "
  strSQL = strSQL & "                             ,2 )  "
  
  strSQL = strSQL & "   FROM  DO30_DOCCORPO "
  strSQL = strSQL & "  WHERE DO30_DITTA_CG18  = " & CodiceDitta
  strSQL = strSQL & "    AND DO30_NUMREG_CO99 = '" & Numreg & "'"
  strSQL = strSQL & "    AND DO30_INDTIPORIGA = 0 "

  Gcon_Connect.Execute strSQL
  
  
  On Error GoTo ErrTrap

Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("CaricaValoriDocOrigine")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select

  
End Sub

Private Function CercoFidoCliente(CodCli As Double) As Boolean

  On Error GoTo ErrTrap


  CercoFidoCliente = True
  Pbol_BloccoFido = False

  Dim MyRstFido   As ADODB.Recordset
'  Dim objDecode   As Object
'  Dim Ret         As Boolean
'
'  Set objDecode = CreateObject("LLBO_CONTROLLOFIDO.CLS_CONTROLLOFIDO")
'  Set objDecode.ActiveInterface = ActiveInterface
'  Set objDecode.ADOConnection = Gcon_Connect
'
'  Ret = objDecode.IsFuoriFido(CodiceDitta, CodCli, False)
'
'  Set objDecode.ActiveInterface = Nothing
'  Set objDecode = Nothing
  
  
'  Dim Ret         As Boolean
'  Dim objFido     As New CLS_CONTROLLOFIDO
'
'  If objFido.OpenConnection(ActiveInterface.ClsGlobal.Gcls_GeConfig.ServerName, ActiveInterface.ClsGlobal.Gcls_GeConfig.DBSa, ActiveInterface.ClsGlobal.Gcls_GeConfig.DBSaPwd, ActiveInterface.ClsGlobal.Gcls_GeConfig.DBName) Then
'    Ret = objFido.IsFuoriFido(CodiceDitta, CodCli, False)
'  End If
'
'  Set objFido = Nothing
  
  Dim Ret         As Boolean
  Ret = ElaboraSeFuoriFido(CodCli, False)
  
  If ClienteNoFido Then
    Pbol_BloccoFido = False
    CercoFidoCliente = True
        
    Exit Function
  End If
  
  'carico il valore del fido dalla tabella MG2D
  strSQL = "SELECT * "
  strSQL = strSQL & "       FROM MG2D_STORICORISCHIO "
  strSQL = strSQL & "      WHERE MG2D_DITTA_CG18  = " & CodiceDitta
  strSQL = strSQL & "        AND MG2D_TIPOCF_CG44 = 0"
  strSQL = strSQL & "        AND MG2D_CLIFOR_CG44 = " & CodCli
  
  Set MyRstFido = Gcon_Connect.Execute(strSQL, , adCmdText)
  If Not MyRstFido.EOF Then
    ValoreFido = MyRstFido("MG2D_FIDO")
    ResiduoFido = MyRstFido("MG2D_FIDO") - MyRstFido("MG2D_FIDOCALCOLATO")
  Else
    ValoreFido = 9999999
    ResiduoFido = 9999999
  End If
  MyRstFido.Close
  Set MyRstFido = Nothing
  
  'solamente per condizione pagamento diverso da contanti
  Select Case TXT_GB06_CODPAG_CG62.Text
  Case CODPAGCONTANTI, CODPAGBANCOMAT, CODPAGCARTACR, CODPAGASSEGNI ', CODPAGCONTSPEC
    Pbol_BloccoFido = False
    Exit Function
  End Select
  
  'Solamente se documento diverso da scontrino
  If TipoDocumento = 2 Or TipoDocumento = 6 Then
    Pbol_BloccoFido = False
    Exit Function
  End If
  
  If Ret Then
    Pbol_BloccoFido = True
    CercoFidoCliente = False
    MsgBox "Attenzione, cliente fuori fido di " & Round(Abs(ResiduoFido), 2)
  End If
  
'''  If Not FRMGB_PASSWORD.FormAperto Then
'''    If Ret And Not (FRMGB_PASSWORD.PasswordCorretta) Then
'''      'Apro mascera per password
'''      If TXT_GB06_CODPAG_CG62.Text = CODPAGCONTSPEC Then
'''        Pbol_BloccoFido = True
'''        MsgBox "Attenzione, cliente fuori fido di " & ResiduoFido
'''      Else
'''
'''        FRMGB_PASSWORD.lblTipoBlocco.Caption = "Attenzione, cliente fuori fido di " & ResiduoFido
'''
'''        FRMGB_PASSWORD.Password = GB05_PWDFIDO
'''        FRMGB_PASSWORD.Show vbModal
'''
'''        If FRMGB_PASSWORD.PasswordCorretta Then
'''          Pbol_BloccoFido = False
'''          CercoFidoCliente = True
'''        Else
'''          Pbol_BloccoFido = True
'''          CercoFidoCliente = False
'''        End If
'''      End If
'''    End If
'''  End If
  
  Set MyRstFido = Nothing
'  Set objDecode = Nothing


Exit Function
ErrTrap:
    Select Case VisualizzaErrore("CercoFidoCliente")
        Case vbAbort
            Exit Function
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select


End Function






Private Sub QGRID_SIMULAZIONE_CustomDrawCell(ByVal DataField As String, ByVal value As Variant, node As TMS_QGRID.cNODE, Column As TMS_QGRID.cCOLONNAGRIGLIA, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal Group As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.Font, FontColor As Long, Alignment As TMS_QGRID.enm_ALIGNMENT, ByVal Left As Single, ByVal Top As Single, ByVal Right As Single, ByVal Bottom As Single)
If InStr(DataField, "_NEW") <> 0 Then
 Color = "&H0000FF00"
End If
End Sub

Private Sub QGridDocumenti_DataFormatEXT(Text As String, ByVal FieldName As String, node As TMS_QGRID.cNODE)

    
  On Error GoTo ErrTrap
    
    Dim MyRst       As ADODB.Recordset
    Dim MySql       As String
    Dim MyId        As Variant
    
'    If VBA.Left(DataField, 6) <> "DECODE" Then Exit Sub
    
'    MyId = NVL(Node.ValueByColumnDataField(FieldName), "")
'    If MyId = "" Then Exit Sub
    
    MySql = ""
    
    Select Case FieldName

    
        Case "DESART"
          MySql = "SELECT  MG87_DESCART"
          MySql = MySql & " FROM MG66_ANAGRART INNER JOIN"
          MySql = MySql & " MG87_ARTDESC ON MG66_DITTA_CG18 = MG87_DITTA_CG18 "
          MySql = MySql & " AND MG66_CODART = MG87_CODART_MG66"
          MySql = MySql & " WHERE MG66_DITTA_CG18 = " & CodiceDitta
          MySql = MySql & " AND MG87_OPZIONE_MG5E = '' AND MG87_LINGUA_MG52 = ''"
          MySql = MySql & " AND MG66_CODART = '" & NVL(node.ValueByColumnDataField("GB07_CODART_MG66"), 0) & "'"

'        Case "RAGSOCFATT1"
'          MySql = " SELECT CG16_RAGSOANAG" & _
'                  " FROM CG44_CLIFOR INNER JOIN CG16_ANAGGEN ON CG44_CLIFOR.CG44_CODICE_CG16 = CG16_ANAGGEN.CG16_CODICE" & _
'                  " WHERE CG44_CLIFOR.CG44_DITTA_CG18 = " & CodiceDitta & _
'                  " AND CG44_CLIFOR.CG44_TIPOCF = 0 " & _
'                  " AND CG44_CLIFOR = " & NVL(Node.ValueByColumnDataField("LL01_CLIFORFATT1_CG44"), 0)
'        Case "RAGSOCFATT2"
'          MySql = " SELECT CG16_RAGSOANAG" & _
'                  " FROM CG44_CLIFOR INNER JOIN CG16_ANAGGEN ON CG44_CLIFOR.CG44_CODICE_CG16 = CG16_ANAGGEN.CG16_CODICE" & _
'                  " WHERE CG44_CLIFOR.CG44_DITTA_CG18 = " & CodiceDitta & _
'                  " AND CG44_CLIFOR.CG44_TIPOCF = 0 " & _
'                  " AND CG44_CLIFOR = " & NVL(Node.ValueByColumnDataField("LL01_CLIFORFATT2_CG44"), 0)
'        Case "RAGSOCFATT3"
'          MySql = " SELECT CG16_RAGSOANAG" & _
'                  " FROM CG44_CLIFOR INNER JOIN CG16_ANAGGEN ON CG44_CLIFOR.CG44_CODICE_CG16 = CG16_ANAGGEN.CG16_CODICE" & _
'                  " WHERE CG44_CLIFOR.CG44_DITTA_CG18 = " & CodiceDitta & _
'                  " AND CG44_CLIFOR.CG44_TIPOCF = 0 " & _
'                  " AND CG44_CLIFOR = " & NVL(Node.ValueByColumnDataField("LL01_CLIFORFATT3_CG44"), 0)
    
    End Select
    
    If MySql <> "" Then
          Set MyRst = Gcon_Connect.Execute(MySql, , adCmdText)
          If Not MyRst.EOF Then
              Text = NVL(MyRst.Fields(0).value, "")
          End If
          Set MyRst = Nothing
    End If

  
  Exit Sub
ErrTrap:
  Select Case VisualizzaErrore("QGridDocumenti_DataFormatEXT")
    Case vbAbort
      Exit Sub
    Case vbRetry
      Resume
    Case vbIgnore
      Resume Next
  End Select


End Sub

Private Sub TabDocumenti_Click()
  If TabDocumenti.ActiveTab = 0 Then


  Else
    Select Case TipoDocumento
    Case 1
      TXT_GB06_CLIFOR_CG44.SetFocus
    Case 2
      TXT_GB07_CODART_MG66.SetFocus
    Case 3
      TXT_GB06_CLIFOR_CG44.SetFocus
    Case 4
      TXT_GB06_CLIFOR_CG44.SetFocus
    Case 5
    TXT_GB06_CLIFOR_CG44.Enabled = True
      TXT_GB06_CLIFOR_CG44.SetFocus
    Case 6
      TXT_GB07_CODART_MG66.SetFocus
    End Select
  End If
End Sub




Private Sub TMS_FLATBUTTON1_Click()
Call InserisciAllegati
End Sub

Private Sub TMS_FLATBUTTON4_Click()
TMS_SSTAB1.ActiveTab = 5
'View Pic From DB
Dim OutFile() As Byte
Dim strSQL As String
Dim FileImgName As String
FileImgName = "Temp"
FileImgName = FileImgName & "_" & TXT_GB07_CODART_MG66.Text

Dim rsSearchResults As ADODB.Recordset

    strSQL = ""
    strSQL = "select GB07_img from GB07_CORPODOC where GB07_ID = " & rstCorpoBanco("GB07_ID")

    'Open a recordset to hold the search results.
    Set rsSearchResults = New ADODB.Recordset
    rsSearchResults.Open strSQL, Gcon_Connect, adOpenStatic, _
        adLockPessimistic

    If rsSearchResults.EOF Or IsNull(rsSearchResults("GB07_img").value) Then
        MsgBox "Immagine non Caricata"
        'Don't change rs, since no match was found we'll
        ' stay on whatever
        'record was previously selected.
    Else
        'Set the form's current recordset to hold only the
        ' search results.
        'Set rs = rsSearchResults

       If Dir(App.Path & "\temp.jpg") <> "" Then Kill _
            App.Path & "\temp.jpg"

        OutFile = rsSearchResults("GB07_img")
        'Write File
        Open App.Path & "\temp.jpg" For Binary Access Write _
            As #1
        Put #1, , OutFile
        Close #1

        PictureArticoli.Picture = LoadPicture(App.Path & "\temp.jpg")
    End If
 
    Set rsSearchResults = Nothing
        
   
End Sub



Private Sub TMS_FLATBUTTON5_Click()
PictureArticoli.Picture = LoadPicture("")
TMS_SSTAB1.ActiveTab = 0
End Sub



Private Sub TMS_FLATBUTTON7_Click()

End Sub

Private Sub TMS_GRIDDOC_ImageClick(ByVal ColumnCaption As String, ByVal FieldName As String, ByVal value As Variant)
If FieldName = "DO11_DITTA_CG18" Then
Call ApriDoc(TMS_GRIDDOC.DataSource("DO11_DOCUM_MG36"), TMS_GRIDDOC.DataSource("GB09_NUMREG_CO99"))
End If
End Sub

Private Sub TMS_RESIZEFORM1_VBRGetHandle(ControlHandle As Long, ContainerHandle As Long)

End Sub



Private Sub TMS_SSTAB1_TabSwitch(ByVal iLastActiveTab As Integer)

Select Case TMS_SSTAB1.ActiveTab
Case 2
    

Case 3
    Call LoadGrigliaOrdini
Case 4
    Call ImpostaVirtualFrame_Note
Case 6
    Call ImpostaVirtualFrame_Gruppi
Case 0
    Call ImpostaVirtualFrame
Case 1
    Call RipristinaOriginale("TMP_SIMULAZIONE_" & ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Codice)
End Select
 
End Sub

Private Sub AnalisiMargini(Filtro As String, Etichetta As String)
''carica griglia
'    Dim StringaSQL As String
'
'    StringaSQL = "SELECT     "
'    StringaSQL = StringaSQL & " ROW_NUMBER() OVER (ORDER BY GB07_ID_GB06) AS Id, "
'    StringaSQL = StringaSQL & " " & Filtro & " as Raggruppamento, "
'    StringaSQL = StringaSQL & " SUM(GB07_IMPORTO) AS Vendita,"
'    StringaSQL = StringaSQL & " SUM(PrezzoAcquistoTotale) AS Acquisto,"
'    StringaSQL = StringaSQL & " SUM(GB07_IMPORTO) - SUM(PrezzoAcquistoTotale) - SUM(GB07_IMPPROVV) as Margine,"
'    StringaSQL = StringaSQL & " SUM(GB07_IMPPROVV) as Provvigione"
'    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
'    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IDGB06 & ")"
'    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06, " & Filtro
'
'
'
'    Set Gcls_RecordPadre = New CLSFW_Recordset
'    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
'    With rstGridMargine
'        Set .ActiveConnection = Gcon_Connect
'        .Source = StringaSQL
'        .Open
'        .MarshalOptions = adMarshalModifiedOnly
'    End With
'
'    INITGRID_MARGINE (Etichetta)
'    TMS_MARGINE.BeginDataSourceSuspend
'
'    Set TMS_MARGINE.DataSource = rstGridMargine
'    TMS_MARGINE.EndDataSourceSuspend
'    TMS_MARGINE.Refresh
    
End Sub


'Private Sub TMS_TIPOGRUPPO_AfterChange(Cancel As Boolean)
'
'
'Select Case TMS_TIPOGRUPPO.Text
'    Case 0
'        Call AnalisiMargini("MG66_FAM_MG53", "Sotto Famiglia")
'    Case 1
'        Call AnalisiMargini("MG66_SFAM_MG54", "Sotto Famiglia")
'    Case 2
'        Call AnalisiMargini("MG66_GRUPPO_MG55", "Sotto Famiglia")
'    Case 3
'        Call AnalisiMargini("MG66_SGRUPPO_MG56", "Sotto Famiglia")
'End Select
'
'End Sub

Private Sub TXT_CG16_RAGSOANAG_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 0 And KeyCode = vbKeyReturn Then
        KeyCode = 0
        'Pint_TipoLookup = 1
        Call TXT_CG16_RAGSOANAG.StartLookup
       ' Pint_TipoLookup = 0
'        If Pint_Risposta = vbYes Then
''            Set FrmDocumenti.Cls_ConnectCommon.ActiveInterface = ActiveInterface
''            FrmDocumenti.Cls_ConnectCommon.CodiceAnagrafica = 0
''            FrmDocumenti.Cls_ConnectCommon.CallAnagraficaGenerale
''            ActiveInterface.IsActive = True
''            Set FrmDocumenti.Cls_ConnectCommon.ActiveInterface = Nothing
''            FrmDocumenti.Cls_ConnectCommon.TerminateConnect
''            Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
'            Call InitializeScript
'        Else
            DoEvents
            TXT_CG16_RAGSOANAG.SetTextFocus
'        End If
    End If
End Sub

Private Sub TXT_CG16_RAGSOANAG_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
 Dim Pvar_CodCliFor          As Variant
  Dim Pint_Risposta                   As Integer
    On Error Resume Next
    
    Cancel = True
    Pint_Risposta = vbNo
    
'    If FrmDocumenti.Cls_PersDoc.INDCLIFOR = tsIndCliForNoCliFor Then
'        Exit Sub
'    End If
    
    If Not (Pobj_SpecialLookUp Is Nothing) Then
        Set Pobj_SpecialLookUp = Nothing
    End If
        
    Set Pobj_SpecialLookUp = New COUO_QUERYANAGPDC.CLSCO_LOOKUP
    Set Pobj_SpecialLookUp.GConnect = Gcon_Connect
    Set Pobj_SpecialLookUp.ActiveInterface = ActiveInterface
    Pobj_SpecialLookUp.Sconnect = Gcon_Connect.ConnectionString
    Pobj_SpecialLookUp.DittaCorrente = CodiceDitta
        
'    Select Case FrmDocumenti.Cls_PersDoc.INDCLIFOR
'        Case tsIndCliForCliente
            Pobj_SpecialLookUp.Caption = "Clienti e anagrafiche non assegnate"
            Pobj_SpecialLookUp.TipoRicerca = ClientiPiùAnagrafica
'        Case tsIndCliForFornitore
'            Pobj_SpecialLookUp.Caption = "Fornitori e anagrafiche non assegnate"
'            Pobj_SpecialLookUp.TipoRicerca = FornitoriPiùAnagrafica
'        Case Else
'            Exit Sub
'    End Select
        
    Pobj_SpecialLookUp.Filtro = RTrimN(TXT_CG16_RAGSOANAG.Text)
   ' Pobj_SpecialLookUp.DataValidita = FME_DocTestata.Recordset("DO11_DATADOC").Value
    Pobj_SpecialLookUp.StartLookup
    TXT_GB06_CLIFOR_CG44.Text = Pobj_SpecialLookUp.CodiceCliente
    TXT_GB07_CODART_MG66.SetFocus
'    If (FrmDocumenti.Cls_PersDoc.INDCLIFOR = tsIndCliForCliente And RTrimN(Pobj_SpecialLookUp.CodiceCliente) <> "") Or _
'        (FrmDocumenti.Cls_PersDoc.INDCLIFOR = tsIndCliForFornitore And RTrimN(Pobj_SpecialLookUp.CodiceFornitore) <> "") Then
'        Select Case FrmDocumenti.Cls_PersDoc.INDCLIFOR
'            Case tsIndCliForCliente
'                Pvar_CodCliFor = Pobj_SpecialLookUp.CodiceCliente
'            Case tsIndCliForFornitore
'                Pvar_CodCliFor = Pobj_SpecialLookUp.CodiceFornitore
'        End Select
'        If FrmDocumenti.ValidaClienteFornitore(Pvar_CodCliFor) Then
'            DoEvents
'            If CDecN(FrmDocumenti.Gvar_CliForSelezionato) <> 0 Then
'                Pvar_CodCliFor = CDecN(FrmDocumenti.Gvar_CliForSelezionato)
'            End If
'            TXT_CG16_RAGSOANAG_DO11_CLIFOR_CG44.Enabled = False
'            Call FrmDocumenti.DisconnettiRecordsetsAiFrame
'            FrmDocumenti.Gbol_SyncFrames = False
'            FrmDocumenti.BoRegDoc.CambiaDocTestataCliFor Pvar_CodCliFor
'            FrmDocumenti.Gbol_SyncFrames = True
'            Call FrmDocumenti.RiconnettiRecordsetsAiFrame
'        End If
'    Else
'        If RTrimN(Pobj_SpecialLookUp.CodiceAnagrafica) <> "" Then
'            Set FrmDocumenti.Cls_ConnectCommon.ActiveInterface = ActiveInterface
'            FrmDocumenti.Cls_ConnectCommon.CodiceClifor = 0
'            FrmDocumenti.Cls_ConnectCommon.CodiceAnagrafica = Pobj_SpecialLookUp.CodiceAnagrafica
'            If Not (TXT_DO11_CLIFOR_CG44 Is Nothing) Then
'                Set FrmDocumenti.Cls_ConnectCommon.ConnectField = TXT_DO11_CLIFOR_CG44
'            End If
'            Select Case CDecN(ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsParMagaz.IndCallCliFor)
'                Case 0 'anagrafica clienti/fornitori
'                    Select Case FrmDocumenti.Cls_PersDoc.INDCLIFOR
'                        Case tsIndCliForCliente
'                            FrmDocumenti.Cls_ConnectCommon.CallAnagClienti
'                        Case tsIndCliForFornitore
'                            FrmDocumenti.Cls_ConnectCommon.CallAnagfornitori
'                    End Select
'                Case 1 'wizard creazione veloce
'                    Select Case FrmDocumenti.Cls_PersDoc.INDCLIFOR
'                        Case tsIndCliForCliente
'                            FrmDocumenti.Cls_ConnectCommon.WizardClienti
'                        Case tsIndCliForFornitore
'                            FrmDocumenti.Cls_ConnectCommon.WizardFornitori
'                    End Select
'            End Select
'            ActiveInterface.IsActive = True
'            Set FrmDocumenti.Cls_ConnectCommon.ActiveInterface = Nothing
'            FrmDocumenti.Cls_ConnectCommon.TerminateConnect
'            Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
'            Call InitializeScript
'        Else
'            Select Case Pint_TipoLookup
'                Case 0
'                    MsgBox "Nessuna anagrafica selezionata !!!", vbInformation + vbOKOnly, "Ricerca anagrafiche"
'                Case 1
'                    Pint_Risposta = MsgBox("Nessuna anagrafica selezionata !!!" & vbCr & "VUOI CREARLA ?", vbInformation + vbYesNo + vbDefaultButton2, "Ricerca anagrafiche")
'            End Select
'        End If
'    End If
    Set Pobj_SpecialLookUp.GConnect = Nothing
    Set Pobj_SpecialLookUp.ActiveInterface = Nothing
    Set Pobj_SpecialLookUp = Nothing
End Sub

Private Sub TXT_CLIFOR_CG44_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
On Error Resume Next

    Call Cls_LookupCommon.ClientiFornitori(1)
    str_SQL = Cls_LookupCommon.StringaSQL

    str_SQL = Replace(str_SQL, "cg44_tipocf = 1", " cg44_tipocf = 1 ")

    'str_SQL = " SELECT" & _
              "    CG16_RAGSOANAG," & _
              "    CG16_INDIRIZZO," & _
              "    CG16_CAP," & _
              "    CG16_CITTA," & _
              "    CG16_PROV, " & _
              "    CG44_CODPAG_CG62, " & _
              "    MG19_LISTMAG "
    'str_SQL = str_SQL & _
              " FROM" & _
              "    CG44_CLIFOR WITH (NOLOCK)" & _
              " INNER JOIN CG16_ANAGGEN WITH (NOLOCK) ON" & _
              "    CG16_CODICE = CG44_CODICE_CG16 " & _
              " INNER JOIN MG19_CLIFORVA WITH (NOLOCK) ON" & _
              "    MG19_DITTA_CG18  = CG44_DITTA_CG18  AND " & _
              "    MG19_TIPOCF_CG44 = CG44_TIPOCF AND " & _
              "    MG19_CLIFOR_CG44 = CG44_CLIFOR " & _
              " WHERE" & _
              "    CG44_DITTA_CG18 = " & CodiceDitta & " AND" & _
              "    CG44_TIPOCF = 0  AND" & _
              "    CG44_CLIFOR = " & CDecN(TXT_GB06_CLIFOR_CG44.Text) & " AND" & _
              "    CG16_RAGSOANAG NOT LIKE '*%'"


    Arr_Fields = Cls_LookupCommon.ArrayFields
    Str_Caption = Cls_LookupCommon.Titolo
    Str_Connect = Gstr_Connect
    TXT_CLIFOR_CG44.IDLookup = Cls_LookupCommon.IDLookup
End Sub

Private Sub TXT_CONTATTO_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
  If NVL(TXT_CONTATTO.Text, "") = "" Then
      Exit Sub
  End If

  Cancel = False

  ReDim Arr_Fields(0 To 0, 0 To 1)
  Arr_Fields(0, 1) = "ALLACA"
  Set Arr_Fields(0, 0) = TXT_GB06_ALLACA
  
  'Imposto la stringa SQL
  '
  
  str_SQL = " SELECT  'Alla C.A. ' + CO02_COGNOMENOME as ALLACA "
  str_SQL = str_SQL & " FROM  VTU05_CONTATTI "
  str_SQL = str_SQL & " WHERE  "
  str_SQL = str_SQL & "   CO02_ID = '" & TXT_CONTATTO.Text & "'"
  str_SQL = str_SQL & " and (CG44_CLIFOR = " & NVL(TXT_GB06_CLIFOR_CG44.Text, 0) & ") AND (CG44_DITTA_CG18 = " & CodiceDitta & ") AND (CG44_TIPOCF = 0) AND (CO03_TIPO_CO01 = '1') "
  
  
 
  Str_Connect = Gstr_Connect
  Err.Clear
End Sub

Private Sub TXT_CONTATTO_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
On Error Resume Next
    
    Dim Pst_Colonne(0 To 1, 0 To 1) As Variant
    
    Pst_Colonne(0, 0) = "Codice"
    Pst_Colonne(0, 1) = ""
    Pst_Colonne(1, 0) = "Descrizione"
    Pst_Colonne(1, 1) = ""
    

    
    str_SQL = " Select  CO02_ID, CO02_COGNOMENOME  FROM            VTU05_CONTATTI " & _
    " WHERE        (CG44_CLIFOR = " & NVL(TXT_GB06_CLIFOR_CG44.Text, 0) & ") AND (CG44_DITTA_CG18 = " & CodiceDitta & ") AND (CG44_TIPOCF = 0) AND (CO03_TIPO_CO01 = '1') "


    
    Str_Caption = "Elenco"
    Arr_Fields = Pst_Colonne
    Str_Connect = Gstr_Connect
    TXT_CONTATTO.IDLookup = "Lkp_Contatto"
End Sub

'Private Sub TXT_DEPCOLL_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
'On Error Resume Next
'    Call Cls_DecodeMagazzino.Deposito(RTrimN(TXT_DEPCOLL.Text))
'    str_SQL = Cls_DecodeMagazzino.StringaSQL
'    Arr_Fields = Cls_DecodeMagazzino.ArrayFields
'
'    Set Arr_Fields(0, 0) = TXT_DESC_DEP_COLL
'
'    Str_Connect = Gstr_Connect
'    If IsNull(Not TXT_DESC_DEP_COLL.Text) Then
'      '  Call cmdAnnulla_ButtonClick
'    End If
'End Sub

'Private Sub TXT_DEPCOLL_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
' On Error Resume Next
'
'  Call Cls_LookupMagazzino.Depositi
'  str_SQL = Cls_LookupMagazzino.StringaSQL
'  Arr_Fields = Cls_LookupMagazzino.ArrayFields
'  Str_Caption = Cls_LookupMagazzino.Titolo
'  Str_Connect = Gstr_Connect
'  TXT_DEPCOLL.IDLookup = Cls_LookupMagazzino.IDLookup
'End Sub



Private Sub TXT_GB07_CODLOTTO_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
  On Error Resume Next

    If RTrimN(FME_BANCO.Recordset("GB07_CODART_MG66").value) = "" Then
        Call Cls_LookupLotti.LottiSenzaArticolo
    Else
        Call Cls_LookupLotti.Lotti(RTrimN(FME_BANCO.Recordset("GB07_CODART_MG66").value), RTrimN(FME_BANCO.Recordset("GB07_OPZIONE_MG5E").value))
    End If
    str_SQL = Cls_LookupLotti.StringaSQL
    Arr_Fields = Cls_LookupLotti.ArrayFields
    Str_Caption = Cls_LookupLotti.Titolo
    Str_Connect = Gstr_Connect
    
End Sub
'
'Private Sub TXT_LL03_CONTODEST_PC03_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
'  'On Error Resume Next
'
'  Cancel = False
'
'  str_SQL = " SELECT    PC03_CONTO, PC03_DESCR"
'  str_SQL = str_SQL & " FROM PC03_ANAGPDC "
'  str_SQL = str_SQL & " WHERE   PC03_CODICE_PC01 = " & ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsParCOIN.CodicePDC
'
'  ReDim Arr_Fields(0 To 1, 0 To 1)
'  Arr_Fields(0, 0) = "Conto"
'  Arr_Fields(0, 1) = ""
'  Arr_Fields(1, 0) = "Descrizione"
'  Arr_Fields(1, 1) = ""
'
'  Str_Caption = "CDC"
'  Str_Connect = Gstr_Connect
'  TXT_LL03_CONTODEST_PC03.IDLookup = "lkp_CDC"
'
'  Err.Clear
'
'End Sub

'Private Sub TXT_LL03_CONTODEST_PC03_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
''
'  On Error Resume Next
'  '
'  'Imposto l'array dei campi
'  '
'  If NVL(TXT_GB07_CDC.Text, "") = "" Then
'      Exit Sub
'  End If
'
'  Cancel = False
'
'  ReDim Arr_Fields(0 To 0, 0 To 1)
'  Arr_Fields(0, 1) = "PC03_DESCR"
'  Set Arr_Fields(0, 0) = TXT_CDC_CAR_DEC
'
'  'Imposto la stringa SQL
'  '
'
'  str_SQL = " SELECT    PC03_DESCR"
'  str_SQL = str_SQL & " FROM PC03_ANAGPDC "
'  str_SQL = str_SQL & " WHERE PC03_CODICE_PC01 = 1 "
'  str_SQL = str_SQL & " AND PC03_CONTO = '" & TXT_LL03_CONTODEST_PC03.Text & "'"
'
'  Str_Connect = Gstr_Connect
'  Err.Clear
'
'End Sub





Private Sub TXT_FAM_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
'On Error Resume Next
'    Call Cls_DecodeMagazzino.Famiglia(TXT_FAM.Text)
'    str_SQL = Cls_DecodeMagazzino.StringaSQL
'    Arr_Fields = Cls_DecodeMagazzino.ArrayFields
'
'    Set Arr_Fields(0, 0) = TXT_FAM_DEC
'
'    Str_Connect = Gstr_Connect
If NVL(TXT_FAM.Text, "") = "" Then
      Exit Sub
  End If

  Cancel = False

  ReDim Arr_Fields(0 To 0, 0 To 1)
  Arr_Fields(0, 1) = "MG53_DESCRFAM"
  Set Arr_Fields(0, 0) = TXT_FAM_DEC
  
  'Imposto la stringa SQL
  '
  
  str_SQL = " SELECT  MG53_DESCRFAM "
  str_SQL = str_SQL & " FROM MG53_FAMIGLIE "
  str_SQL = str_SQL & " WHERE MG53_DITTA_CG18  = " & CodiceDitta
  str_SQL = str_SQL & "   AND MG53_CODFAM = '" & TXT_FAM.Text & "'"
  
  
  
 
  Str_Connect = Gstr_Connect
  Err.Clear
    
End Sub

Private Sub TXT_FAM_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
'  On Error Resume Next
'
'  Call Cls_LookupMagazzino.Famiglie
'  str_SQL = Cls_LookupMagazzino.StringaSQL
'  Arr_Fields = Cls_LookupMagazzino.ArrayFields
'  Str_Caption = Cls_LookupMagazzino.Titolo
'  Str_Connect = Gstr_Connect
'  TXT_FAM.IDLookup = Cls_LookupMagazzino.IDLookup
 On Error Resume Next

  Cancel = False
  str_SQL = " SELECT     MG53_CODFAM, MG53_DESCRFAM "
  str_SQL = str_SQL & " FROM MG53_FAMIGLIE "
  str_SQL = str_SQL & " WHERE MG53_DITTA_CG18  = " & CodiceDitta
  
  
  ReDim Arr_Fields(0 To 1, 0 To 1)
  Arr_Fields(0, 0) = "Codice"
  Arr_Fields(0, 1) = ""
  Arr_Fields(1, 0) = "Descrizione"
  Arr_Fields(1, 1) = ""

  Str_Caption = "Famiglie"
  Str_Connect = Gstr_Connect
  TXT_FAM = "lkp_Famiglie"
  
  Err.Clear
End Sub

Private Sub TXT_FLPOSA_AfterItem(Cancel As Boolean)
If TXT_FLPOSA.Text = "1" Then
TXT_GB07_FLPOSA.Text = "0"
TXT_GB07_FLPOSA.Default = "0"
Else
TXT_GB07_FLPOSA.Default = "1"
TXT_GB07_FLPOSA.Default = "1"
End If
End Sub

Private Sub TXT_GB06_AGENTE_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
If NVL(TXT_GB06_AGENTE.Text, "") = "" Then
      Exit Sub
  End If

  Cancel = False

  ReDim Arr_Fields(0 To 0, 0 To 1)
  Arr_Fields(0, 1) = "CG16_RAGSOANAG"
  Set Arr_Fields(0, 0) = TXT_GB06_AGENTE_DEC
  
  'Imposto la stringa SQL
  '
  
  str_SQL = " SELECT        CG16_ANAGGEN.CG16_RAGSOANAG"
  str_SQL = str_SQL + " FROM            MG17_AGENTI INNER JOIN"
  str_SQL = str_SQL + "                         CG16_ANAGGEN ON MG17_AGENTI.MG17_ANAGEN_CG16 = CG16_ANAGGEN.CG16_CODICE"
  str_SQL = str_SQL + " Where(MG17_AGENTI.MG17_DITTA_CG18 = " & CodiceDitta & ") and  MG17_AGENTI.MG17_AGENTE = '" & TXT_GB06_AGENTE.Text & "'"
    
  
 
  Str_Connect = Gstr_Connect
  Err.Clear
End Sub

Private Sub TXT_GB06_AGENTE_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)



 On Error Resume Next

  Cancel = False
  str_SQL = " SELECT        MG17_AGENTI.MG17_AGENTE, CG16_ANAGGEN.CG16_RAGSOANAG"
  str_SQL = str_SQL + " FROM            MG17_AGENTI INNER JOIN"
  str_SQL = str_SQL + "                         CG16_ANAGGEN ON MG17_AGENTI.MG17_ANAGEN_CG16 = CG16_ANAGGEN.CG16_CODICE"
  str_SQL = str_SQL + " Where(MG17_AGENTI.MG17_DITTA_CG18 = " & CodiceDitta & ")"
  
  ReDim Arr_Fields(0 To 1, 0 To 1)
  Arr_Fields(0, 0) = "Codice"
  Arr_Fields(0, 1) = ""
  Arr_Fields(1, 0) = "Descrizione"
  Arr_Fields(1, 1) = ""

  Str_Caption = "Agenti"
  Str_Connect = Gstr_Connect
  TXT_GB06_AGENTE = "lkp_Agenti"
  
  Err.Clear


End Sub

'Private Sub TXT_GB06_BUDGET_BeforeItem(Cancel As Boolean)
'TXT_GB06_BUDGET.BackColor "&H0000C000"
'End Sub

Private Sub TXT_GB06_CIG_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)

  On Error Resume Next

  Cancel = False
  str_SQL = " SELECT        TOP (200) CO1H_CIGCUP.CO1H_CIG"
  str_SQL = str_SQL & " FROM            CO1I_CFCIGCUP INNER JOIN  CO1H_CIGCUP ON CO1I_CFCIGCUP.CO1I_IDCIGCUP_CO1H = CO1H_CIGCUP.CO1H_ID "
  str_SQL = str_SQL & " WHERE  (CO1I_CFCIGCUP.CO1I_DITTA_CG18 = " & CodiceDitta & ") And (CO1I_CFCIGCUP.CO1I_TIPOCF_CG44 = 0) And (CO1I_CFCIGCUP.CO1I_CLIFOR_CG44 = " & NVL(TXT_GB06_CLIFOR_CG44.Text, 0) & ")"
  
  
  ReDim Arr_Fields(0 To 1, 0 To 1)
  Arr_Fields(0, 0) = "Codice"
  Arr_Fields(0, 1) = ""
  Arr_Fields(1, 0) = "Descrizione"
  Arr_Fields(1, 1) = ""

  Str_Caption = "CIG"
  Str_Connect = Gstr_Connect
  TXT_GB06_CIG = "lkp_CIG"
  
  Err.Clear


End Sub

Private Sub TXT_GB06_CODCOMM_KeyButtonMenuPress(Cancel As Boolean, ByVal Pstr_KeyButtonPress As Variant)
    On Error Resume Next
    
    If Pstr_KeyButtonPress = "Kgestione" Then
        If NVL(TXT_GB06_CODCOMM.Text, "") = 0 Then
            MsgBox "Codice Commessa non impostato !!!"
        Else
            Set Cls_ConnectMagazzino.ActiveInterface = ActiveInterface
            Call Cls_ConnectMagazzino.Commesse(RTrimN(TXT_GB06_CODCOMM.Text))
            ActiveInterface.IsActive = True
            Set Cls_ConnectMagazzino.ActiveInterface = Nothing
            Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
'            Call InitializeScript
        End If
    End If
End Sub

Private Sub TXT_GB06_CODCOMM_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
If NVL(TXT_GB06_CODCOMM.Text, "") = "" Then
      Exit Sub
  End If

  Cancel = False

  ReDim Arr_Fields(0 To 0, 0 To 1)
  Arr_Fields(0, 1) = "PD25_DESCR"
  Set Arr_Fields(0, 0) = TXT_GB06_COMMESSA_DEC
  
  'Imposto la stringa SQL
  '
  
  str_SQL = " SELECT  PD25_DESCR "
  str_SQL = str_SQL & " FROM PD25_COMMESSA "
  str_SQL = str_SQL & " WHERE PD25_DITTA_CG18  = " & CodiceDitta
  str_SQL = str_SQL & "   AND PD25_CODCOMMESSA = '" & TXT_GB06_CODCOMM.Text & "'"
  str_SQL = str_SQL & "   AND PD25_CODSOTCOMM  = 0 "
  
  
 
  Str_Connect = Gstr_Connect
  Err.Clear
End Sub

Private Sub TXT_GB06_CODCOMM_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
  On Error Resume Next

  Cancel = False
  str_SQL = " SELECT     PD25_CODCOMMESSA, PD25_DESCR "
  str_SQL = str_SQL & " FROM PD25_COMMESSA "
  str_SQL = str_SQL & " WHERE PD25_DITTA_CG18  = " & CodiceDitta
  str_SQL = str_SQL & "   AND PD25_CODSOTCOMM  = 0 "
  
  ReDim Arr_Fields(0 To 1, 0 To 1)
  Arr_Fields(0, 0) = "Codice"
  Arr_Fields(0, 1) = ""
  Arr_Fields(1, 0) = "Descrizione"
  Arr_Fields(1, 1) = ""

  Str_Caption = "Commessa"
  Str_Connect = Gstr_Connect
  TXT_GB06_CODCOMM = "lkp_Commessa"
  
  Err.Clear
  
End Sub

Private Sub TXT_GB06_CODDESTIN_MG22_KeyButtonMenuPress(Cancel As Boolean, ByVal Pstr_KeyButtonPress As Variant)
    On Error Resume Next
    
    If Pstr_KeyButtonPress = "Kgestione" Then
        If NVL(TXT_GB06_CLIFOR_CG44.Text, "") = 0 Then
            MsgBox "Codice cliente non impostato !!!"
        Else
            Set Cls_ConnectCommon.ActiveInterface = ActiveInterface
            Cls_ConnectCommon.CodiceDestMerceTipoCF = 0
            Cls_ConnectCommon.CodiceDestMerceClienteFornitore = CDecN(TXT_GB06_CLIFOR_CG44.Text)
            Cls_ConnectCommon.CodiceDestMerceDestinatario = RTrimN(TXT_GB06_CODDESTIN_MG22.Text)
                    
            Call Cls_ConnectCommon.CallDestinatariMerceClienti
                        
            ActiveInterface.IsActive = True
            Set Cls_ConnectCommon.ActiveInterface = Nothing
            Cls_ConnectCommon.TerminateConnect
            Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
'            Call InitializeScript
        End If
    End If
End Sub

Private Sub TXT_GB06_CODDESTIN_MG22_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
On Error Resume Next
    Dim Pvar_ArrCampi() As Variant
    Dim conta As Integer
    str_SQL = "SELECT * FROM MG22_CLIFORDEST WITH (NOLOCK) " & _
              "LEFT OUTER JOIN CG07_TABSTATIEST WITH (NOLOCK) ON " & _
              "CG07_CODICE = MG22_STATOEST_CG07 " & _
              "WHERE MG22_DITTA_CG18 = " & CodiceDitta & " AND MG22_TIPOCF_CG44 = 0 AND MG22_CLIFOR_CG44 = " & CDecN(TXT_GB06_CLIFOR_CG44.Text) & " AND " & _
              "MG22_CODDESTIN = '" & RTrimN(TXT_GB06_CODDESTIN_MG22.Text) & "'"
              
    conta = 5
    
    ReDim Preserve Pvar_ArrCampi(0 To conta - 1, 0 To 1)
    conta = 0
    Set Pvar_ArrCampi(conta, 0) = TXT_MG22_DESTRAGSOC
    Pvar_ArrCampi(conta, 1) = "MG22_DESTRAGSOC"
    conta = conta + 1
        
    Set Pvar_ArrCampi(conta, 0) = TXT_MG22_DESTIND
    Pvar_ArrCampi(conta, 1) = "MG22_DESTIND"
    conta = conta + 1
    
    Set Pvar_ArrCampi(conta, 0) = TXT_MG22_DESTCAP
    Pvar_ArrCampi(conta, 1) = "MG22_DESTCAPCHAR"
    conta = conta + 1
        
    Set Pvar_ArrCampi(conta, 0) = TXT_MG22_DESTPROV
    Pvar_ArrCampi(conta, 1) = "MG22_DESTPROV"
    conta = conta + 1
        
    Set Pvar_ArrCampi(conta, 0) = TXT_MG22_DESTCITTA
    Pvar_ArrCampi(conta, 1) = "MG22_DESTCITTA"
    conta = conta + 1

    Arr_Fields = Pvar_ArrCampi
    Str_Connect = Gstr_Connect
End Sub

Private Sub TXT_GB06_CODDESTIN_MG22_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    On Error Resume Next
    Call Cls_LookupMagazzino.Destinatari(0, CDecN(TXT_GB06_CLIFOR_CG44.Text))
    str_SQL = Cls_LookupMagazzino.StringaSQL
    Arr_Fields = Cls_LookupMagazzino.ArrayFields
    Str_Caption = Cls_LookupMagazzino.Titolo
    Str_Connect = Gstr_Connect
    TXT_GB06_CODDESTIN_MG22.IDLookup = Cls_LookupMagazzino.IDLookup
End Sub

Private Sub TXT_GB06_CODPAG_CG62_BeforeItem(Cancel As Boolean)
Call CMD_SAVE_Click
End Sub

Private Sub TXT_GB06_CUP_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
On Error Resume Next

  Cancel = False
  str_SQL = " SELECT        TOP (200) CO1H_CIGCUP.CO1H_CUP"
  str_SQL = str_SQL & " FROM            CO1I_CFCIGCUP INNER JOIN  CO1H_CIGCUP ON CO1I_CFCIGCUP.CO1I_IDCIGCUP_CO1H = CO1H_CIGCUP.CO1H_ID "
  str_SQL = str_SQL & " WHERE  (CO1I_CFCIGCUP.CO1I_DITTA_CG18 = " & CodiceDitta & ") And (CO1I_CFCIGCUP.CO1I_TIPOCF_CG44 = 0) And (CO1I_CFCIGCUP.CO1I_CLIFOR_CG44 = " & NVL(TXT_GB06_CLIFOR_CG44.Text, 0) & ")"
  
  
  ReDim Arr_Fields(0 To 1, 0 To 1)
  Arr_Fields(0, 0) = "Codice"
  Arr_Fields(0, 1) = ""
  Arr_Fields(1, 0) = "Descrizione"
  Arr_Fields(1, 1) = ""

  Str_Caption = "CIG"
  Str_Connect = Gstr_Connect
  TXT_GB06_CIG = "lkp_CIG"
  
  Err.Clear
End Sub

Private Sub TXT_GB06_DTDOC_AfterItem(Cancel As Boolean)
Call CMD_SAVE_Click
End Sub

Private Sub TXT_GB06_ID_AfterItem(Cancel As Boolean)
 Call StatoCampi
End Sub

Private Sub TXT_GB06_ID_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
Dim conta As Integer
  Dim Pvar_ArrCampi()
  Dim str_sql1 As String
  On Error Resume Next
        
  If NVL(TXT_GB06_NUMDOC.Text, "") = 0 Then
      Exit Sub
  End If


' ",   GB06_UTENTECRE  " & _
'              ",   GB06_CODCOMM  " & _
'              ",   GB06_PERCCHIUSURA  " & _
'              ",   GB06_BUDGET " & _
'              ",   GB06_TIPOAREA " & _
' ",   GB06_DATA  " & _

 str_SQL = " SELECT" & _
              "    GB06_NUMDOC  " & _
              ",   GB06_NREV  " & _
              ",   GB06_NVERS  " & _
              ",   GB06_NOMEOFFERTA  " & _
              ",   GB06_CLIFOR_CG44  " & _
              ",   GB06_UTENTECRE  " & _
              ",   GB06_CODCOMM  " & _
              ",   GB06_PERCCHIUSURA  " & _
              ",   GB06_BUDGET " & _
              ",   GB06_TIPOAREA " & _
              ",   GB06_TIPOOFFERTA " & _
              ",   GB06_STATODOC  " & _
              ",   GB06_CIG " & _
              ",   GB06_CUP " & _
              ",   GB06_ALLACA " & _
              ",   GB06_PERCTRASP " & _
              ",   GB06_CODDESTIN_MG22 " & _
              ",   GB06_RESPONSABILE " & _
              "    FROM        GB06_TESTADOC " & _
              "    WHERE " & _
              "    GB06_DITTA_CG18 = " & CodiceDitta & _
              "    AND GB06_ID = " & TXT_GB06_ID.Text & _
              "    ORder by GB06_ID "
        
        
        
  conta = 18

  ReDim Preserve Pvar_ArrCampi(0 To conta - 1, 0 To 1)
  
  conta = 0
  
  Set Pvar_ArrCampi(conta, 0) = TXT_GB06_NUMDOC
  Pvar_ArrCampi(conta, 1) = "GB06_NUMDOC"
  conta = conta + 1
  
  Set Pvar_ArrCampi(conta, 0) = TXT_GB06_NREV
  Pvar_ArrCampi(conta, 1) = "GB06_NREV"
  conta = conta + 1

  Set Pvar_ArrCampi(conta, 0) = TXT_GB06_NVERS
  Pvar_ArrCampi(conta, 1) = "GB06_NVERS"
  conta = conta + 1
  
  Set Pvar_ArrCampi(conta, 0) = TXT_GB06_NOMEOFFERTA
  Pvar_ArrCampi(conta, 1) = "GB06_NOMEOFFERTA"
  conta = conta + 1

  Set Pvar_ArrCampi(conta, 0) = TXT_GB06_CLIFOR_CG44
  Pvar_ArrCampi(conta, 1) = "GB06_CLIFOR_CG44"
  conta = conta + 1
  
'  Set Pvar_ArrCampi(conta, 0) = TXT_GB06_DTDOC
'  Pvar_ArrCampi(conta, 1) = "GB06_DTDOC"
'  conta = conta + 1
  
  Set Pvar_ArrCampi(conta, 0) = TXT_GB06_PROPRIETARIO
  Pvar_ArrCampi(conta, 1) = "GB06_UTENTECRE"
  conta = conta + 1

  Set Pvar_ArrCampi(conta, 0) = TXT_GB06_CODCOMM
  Pvar_ArrCampi(conta, 1) = "GB06_CODCOMM"
  conta = conta + 1

  Set Pvar_ArrCampi(conta, 0) = TXT_GB06_PERCCHIUSURA
  Pvar_ArrCampi(conta, 1) = "GB06_PERCCHIUSURA"
  conta = conta + 1

  Set Pvar_ArrCampi(conta, 0) = TXT_GB06_BUDGET
  Pvar_ArrCampi(conta, 1) = "GB06_BUDGET"
  conta = conta + 1

  Set Pvar_ArrCampi(conta, 0) = TXT_GB06_TIPOAREA
  Pvar_ArrCampi(conta, 1) = "GB06_TIPOAREA"
  conta = conta + 1
  
  Set Pvar_ArrCampi(conta, 0) = TXT_GB06_TIPOOFFERTA
  Pvar_ArrCampi(conta, 1) = "GB06_TIPOOFFERTA"
  conta = conta + 1
  
  Set Pvar_ArrCampi(conta, 0) = TXT_GB06_STATODOC
  Pvar_ArrCampi(conta, 1) = "GB06_STATODOC"
  conta = conta + 1
  
  Set Pvar_ArrCampi(conta, 0) = TXT_GB06_CIG
  Pvar_ArrCampi(conta, 1) = "GB06_CIG"
  conta = conta + 1
  
  Set Pvar_ArrCampi(conta, 0) = TXT_GB06_CUP
  Pvar_ArrCampi(conta, 1) = "GB06_CUP"
  conta = conta + 1
  
  Set Pvar_ArrCampi(conta, 0) = TXT_GB06_ALLACA
  Pvar_ArrCampi(conta, 1) = "GB06_ALLACA"
  conta = conta + 1
  
  Set Pvar_ArrCampi(conta, 0) = TXT_GB06_PERCTRASP
  Pvar_ArrCampi(conta, 1) = "GB06_PERCTRASP"
  conta = conta + 1
  
  Set Pvar_ArrCampi(conta, 0) = TXT_GB06_CODDESTIN_MG22
  Pvar_ArrCampi(conta, 1) = "GB06_CODDESTIN_MG22"
  conta = conta + 1
  
   Set Pvar_ArrCampi(conta, 0) = TXT_GB06_RESPONSABILE
  Pvar_ArrCampi(conta, 1) = "GB06_RESPONSABILE"
  conta = conta + 1
  
  
  Arr_Fields = Pvar_ArrCampi
  Str_Connect = Gstr_Connect

  
  ' aggiorno le date
  str_sql1 = " SELECT" & _
              "    GB06_DTCHIUSURA  " & _
              ",   GB06_DTULTMOD  " & _
              ",   GB06_DTDOC  " & _
              ",   GB06_TEXT1  " & _
              ",   GB06_TEXT2  " & _
              ",   GB06_TEXT3  " & _
              ",   GB06_TEXT4  " & _
              ",   GB06_TEXT5  " & _
              ",   GB06_TEXT6  " & _
              "    FROM        GB06_TESTADOC " & _
              "    WHERE " & _
              "    GB06_DITTA_CG18 = " & CodiceDitta & _
              "    AND GB06_ID = " & TXT_GB06_ID.Text & ""

  TXT_GB06_DTCHIUSURA.Text = GetValFromQuery(str_sql1, 0, Gcon_Connect)
  TXT_GB06_DTULTMOD.Text = Mid(GetValFromQuery(str_sql1, 1, Gcon_Connect), 1, 10)
  TXT_GB06_DTDOC.Text = GetValFromQuery(str_sql1, 2, Gcon_Connect)
   
  TXT_GB06_TEXT1.Text = GetValFromQuery(str_sql1, 3, Gcon_Connect)
  TXT_GB06_TEXT2.Text = GetValFromQuery(str_sql1, 4, Gcon_Connect)
  TXT_GB06_TEXT3.Text = GetValFromQuery(str_sql1, 5, Gcon_Connect)
  TXT_GB06_TEXT4.Text = GetValFromQuery(str_sql1, 6, Gcon_Connect)
  TXT_GB06_TEXT5.Text = GetValFromQuery(str_sql1, 7, Gcon_Connect)
  TXT_GB06_TEXT6.Text = GetValFromQuery(str_sql1, 8, Gcon_Connect)
 

  
  IDGB06 = TXT_GB06_ID.Text
  
  Call ImpostaVirtualFrame

  Call RicalcolaImportoTotale
 
  Exit Sub
End Sub

Private Sub TXT_GB06_ID_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
On Error Resume Next
    
    Dim Pst_Colonne(0 To 17, 0 To 1) As Variant
    
    Pst_Colonne(0, 0) = "ID Offerta"
    Pst_Colonne(0, 1) = ""
    Pst_Colonne(1, 0) = "Numero Offerta"
    Pst_Colonne(1, 1) = ""
    Pst_Colonne(2, 0) = "Revisione"
    Pst_Colonne(2, 1) = ""
    Pst_Colonne(3, 0) = "Versione"
    Pst_Colonne(3, 1) = ""
    Pst_Colonne(4, 0) = "Data documento"
    Pst_Colonne(4, 1) = ""
    Pst_Colonne(5, 0) = "Nome Offerta"
    Pst_Colonne(5, 1) = ""
    Pst_Colonne(6, 0) = "Ultima Modifica "
    Pst_Colonne(6, 1) = ""
    Pst_Colonne(7, 0) = "Operatore "
    Pst_Colonne(7, 1) = ""
    Pst_Colonne(8, 0) = "Stato "
    Pst_Colonne(8, 1) = ""
    Pst_Colonne(9, 0) = "Data Chiusura Prevista "
    Pst_Colonne(9, 1) = ""
    Pst_Colonne(10, 0) = "Rag.Sociale Cli. "
    Pst_Colonne(10, 1) = ""
    Pst_Colonne(11, 0) = "CAP "
    Pst_Colonne(11, 1) = ""
    Pst_Colonne(12, 0) = "Città "
    Pst_Colonne(12, 1) = ""
    Pst_Colonne(13, 0) = "Prov "
    Pst_Colonne(13, 1) = ""
    Pst_Colonne(14, 0) = "Cod. Agente "
    Pst_Colonne(14, 1) = ""
    Pst_Colonne(15, 0) = "Rag. Soc. Agente "
    Pst_Colonne(15, 1) = ""
    Pst_Colonne(16, 0) = "Stato Offerta "
    Pst_Colonne(16, 1) = ""
    Pst_Colonne(17, 0) = "Commessa "
    Pst_Colonne(17, 1) = ""
   
    
    
    str_SQL = " SELECT GB06_ID," & _
              "    GB06_NUMDOC," & _
              "    GB06_NREV," & _
              "    GB06_NVERS," & _
              "    GB06_DTDOC, " & _
              "    GB06_NOMEOFFERTA, " & _
              "    GB06_DTULTMOD, " & _
              "    GB06_UTENTECRE, " & _
              "    GB06_STATODOC , " & _
              "    GB06_DTCHIUSURA,  " & _
              "    CG16_ANAGGEN.CG16_RAGSOANAG, " & _
              "    CG16_ANAGGEN.CG16_CAP, " & _
              "    CG16_ANAGGEN.CG16_CITTA, " & _
              "    CG16_ANAGGEN.CG16_PROV, " & _
              "    GB06_AGENTE , " & _
              "    CG16_ANAGGEN_1.CG16_RAGSOANAG as RAGSOC_AGENTE ," & _
              "    case GB06_STATODOC when '08' then 'ANNULLATA' when '00' then 'IN LAVORAZIONE' when '01' then 'COMPLETATA' when '02' then 'RILASCIATA' when '03' then 'ATTIVA' when '04' then 'REVISIONATA' when '05' then 'ORDINE' when '06' then 'PERSA' when '07' then 'ARCHIVIATA' end as GB06_STATODOC_DEC , " & _
              "    GB06_CODCOMM " & _
              "    FROM            GB06_TESTADOC INNER JOIN " & _
              "           CG44_CLIFOR ON GB06_TESTADOC.GB06_DITTA_CG18 = CG44_CLIFOR.CG44_DITTA_CG18 AND GB06_TESTADOC.GB06_TIPOCF_CG44 = CG44_CLIFOR.CG44_TIPOCF AND " & _
              "           GB06_TESTADOC.GB06_CLIFOR_CG44 = CG44_CLIFOR.CG44_CLIFOR INNER JOIN " & _
              "           CG16_ANAGGEN ON CG44_CLIFOR.CG44_CODICE_CG16 = CG16_ANAGGEN.CG16_CODICE INNER JOIN " & _
              "           MG17_AGENTI ON GB06_TESTADOC.GB06_DITTA_CG18 = MG17_AGENTI.MG17_DITTA_CG18 AND GB06_TESTADOC.GB06_AGENTE = MG17_AGENTI.MG17_AGENTE INNER JOIN " & _
              "           CG16_ANAGGEN AS CG16_ANAGGEN_1 ON MG17_AGENTI.MG17_ANAGEN_CG16 = CG16_ANAGGEN_1.CG16_CODICE " & _
              "    WHERE GB06_DITTA_CG18 = " & CodiceDitta & " AND  (NOT (GB06_NUMDOC IS NULL)) Order by GB06_ID "


    
    Str_Caption = "Elenco Preventivi Caricati"
    Arr_Fields = Pst_Colonne
    Str_Connect = Gstr_Connect
    TXT_GB06_ID.IDLookup = "Lkp_Offerte"
End Sub

Private Sub TXT_GB06_NOMEOFFERTA_AfterItem(Cancel As Boolean)
Call CMD_SAVE_Click
'    If NVL(TXT_GB06_CODCOMM.Text, "") = "" Then
'        If MsgBox("Vuoi Creare una nuova Commessa?", vbYesNo) = vbYes Then
'             'creazione nuova commessa
'             TXT_GB06_CODCOMM.Text = GetCommessa(IDGB06)
'        End If
'
'    End If
    
End Sub

Private Sub TXT_GB06_NUMDOC_AfterItem(Cancel As Boolean)
'If MsgBox("Vuoi Creare una nuova Commessa?", vbYesNo) = vbYes Then
'     'creazione nuova commessa
'    Else
'      Exit Sub
'    End If
End Sub

Private Sub TXT_GB06_STATODOC_AfterItem(Cancel As Boolean)
    Select Case TXT_GB06_STATODOC.Text
    
'    00 – verde H0000C000
'01 – rosa H00C0C0FF
'02 – arancione H000080FF
'03 – rosso H000000FF
'04 – viola H00C000C0
'05  – giallo H0000FFFF
'06 – nero H00000000
'07 – blu H00C00000
    
    
        Case "00"
         TMS_CODIFICA_STATO.Caption = "In Lavorazione"
         TMS_CODIFICA_STATO.ButtonBackColor = "&H0000C000"
        Case "01"
         TMS_CODIFICA_STATO.Caption = "Completata"
         TMS_CODIFICA_STATO.ButtonBackColor = "&H00C0C0FF"
        Case "02"
         TMS_CODIFICA_STATO.Caption = "Rilasciata"
         TMS_CODIFICA_STATO.ButtonBackColor = "&H000080FF"
        Case "03"
         TMS_CODIFICA_STATO.Caption = "Attiva"
         TMS_CODIFICA_STATO.ButtonBackColor = "&H000000FF"
        Case "04"
         TMS_CODIFICA_STATO.Caption = "Revisionata"
         TMS_CODIFICA_STATO.ButtonBackColor = "&H00C000C0"
        Case "05"
         TMS_CODIFICA_STATO.Caption = "Ordine"
         TMS_CODIFICA_STATO.ButtonBackColor = "&H0000FFFF"
        Case "06"
         TMS_CODIFICA_STATO.Caption = "Persa"
         TMS_CODIFICA_STATO.ButtonBackColor = "&H00000000"
        Case "07"
         TMS_CODIFICA_STATO.Caption = "Archiviata"
         TMS_CODIFICA_STATO.ButtonBackColor = "&H0000C000"
        Case "08"
         TMS_CODIFICA_STATO.Caption = "Annullata"
         TMS_CODIFICA_STATO.ButtonBackColor = "&H000000FF"
'        Case "08"
'         TMS_CODIFICA_STATO.Caption = "Archiviata"
'         TMS_CODIFICA_STATO.ButtonBackColor = "&H0000C000"
        Case Else
         TMS_CODIFICA_STATO.Caption = "no stato"
         TMS_CODIFICA_STATO.ButtonBackColor = "&H00000000"
'    In lavorazione
'Attiva
'Archiviata
'Persa
'Ordine

    End Select
End Sub


Private Sub TXT_GB07_CLIFOR_CG44_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)

  Dim conta As Integer
  Dim Pvar_ArrCampi()
  On Error Resume Next

  If NVL(TXT_GB07_CLIFOR_CG44.Text, "") = 0 Then
      Exit Sub
  End If

  str_SQL = " SELECT" & _
            "    CG16_RAGSOANAG" & _
            " FROM" & _
            "    CG44_CLIFOR WITH (NOLOCK)" & _
            " INNER JOIN CG16_ANAGGEN WITH (NOLOCK) ON" & _
            "    CG16_CODICE = CG44_CODICE_CG16 " & _
            " INNER JOIN MG19_CLIFORVA WITH (NOLOCK) ON" & _
            "    MG19_DITTA_CG18  = CG44_DITTA_CG18  AND " & _
            "    MG19_TIPOCF_CG44 = CG44_TIPOCF AND " & _
            "    MG19_CLIFOR_CG44 = CG44_CLIFOR " & _
            " WHERE" & _
            "    CG44_DITTA_CG18 = " & CodiceDitta & " AND" & _
            "    CG44_TIPOCF = 1  AND" & _
            "    CG44_CLIFOR = " & CDecN(TXT_GB07_CLIFOR_CG44.Text)



  conta = 1

  ReDim Preserve Pvar_ArrCampi(0 To conta - 1, 0 To 1)

  conta = 0

  Set Pvar_ArrCampi(conta, 0) = TXT_CG16_RAGSOANAG_FOR
  Pvar_ArrCampi(conta, 1) = "CG16_RAGSOANAG"
  conta = conta + 1



  Arr_Fields = Pvar_ArrCampi
  Str_Connect = Gstr_Connect
  Exit Sub
End Sub

Private Sub TXT_GB07_CLIFOR_CG44_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
On Error Resume Next

    Call Cls_LookupCommon.ClientiFornitori(1)
    str_SQL = Cls_LookupCommon.StringaSQL

    str_SQL = Replace(str_SQL, "cg44_tipocf = 1", " cg44_tipocf = 1 ")

    'str_SQL = " SELECT" & _
              "    CG16_RAGSOANAG," & _
              "    CG16_INDIRIZZO," & _
              "    CG16_CAP," & _
              "    CG16_CITTA," & _
              "    CG16_PROV, " & _
              "    CG44_CODPAG_CG62, " & _
              "    MG19_LISTMAG "
    'str_SQL = str_SQL & _
              " FROM" & _
              "    CG44_CLIFOR WITH (NOLOCK)" & _
              " INNER JOIN CG16_ANAGGEN WITH (NOLOCK) ON" & _
              "    CG16_CODICE = CG44_CODICE_CG16 " & _
              " INNER JOIN MG19_CLIFORVA WITH (NOLOCK) ON" & _
              "    MG19_DITTA_CG18  = CG44_DITTA_CG18  AND " & _
              "    MG19_TIPOCF_CG44 = CG44_TIPOCF AND " & _
              "    MG19_CLIFOR_CG44 = CG44_CLIFOR " & _
              " WHERE" & _
              "    CG44_DITTA_CG18 = " & CodiceDitta & " AND" & _
              "    CG44_TIPOCF = 0  AND" & _
              "    CG44_CLIFOR = " & CDecN(TXT_GB06_CLIFOR_CG44.Text) & " AND" & _
              "    CG16_RAGSOANAG NOT LIKE '*%'"


    Arr_Fields = Cls_LookupCommon.ArrayFields
    Str_Caption = Cls_LookupCommon.Titolo
    Str_Connect = Gstr_Connect
    TXT_GB07_CLIFOR_CG44.IDLookup = Cls_LookupCommon.IDLookup
End Sub

Private Sub TXT_GB06_TIPOAREA_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
If NVL(TXT_GB06_TIPOAREA.Text, "") = "" Then
      Exit Sub
  End If

  Cancel = False

  ReDim Arr_Fields(0 To 0, 0 To 1)
  Arr_Fields(0, 1) = "GB02_DESCRIZIONE"
  Set Arr_Fields(0, 0) = TXT_GB06_TIPOAREA_DEC
  
  'Imposto la stringa SQL
  '
  
  str_SQL = " SELECT  GB02_DESCRIZIONE "
  str_SQL = str_SQL & " FROM  GB02_LOOKUPVALUE "
  str_SQL = str_SQL & " WHERE  "
  str_SQL = str_SQL & "   GB02_CODICE = '" & TXT_GB06_TIPOAREA.Text & "'"
  str_SQL = str_SQL & "   AND GB02_TIPOCAMPO  = 'TXT_GB06_TIPOAREA' "
  
  
 
  Str_Connect = Gstr_Connect
  Err.Clear
End Sub

Private Sub TXT_GB06_TIPOAREA_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
On Error Resume Next
    
    Dim Pst_Colonne(0 To 1, 0 To 1) As Variant
    
    Pst_Colonne(0, 0) = "Codice"
    Pst_Colonne(0, 1) = ""
    Pst_Colonne(1, 0) = "Descrizione"
    Pst_Colonne(1, 1) = ""
    

    
    str_SQL = " SELECT" & _
              "    GB02_CODICE, " & _
              "    GB02_DESCRIZIONE " & _
              " FROM        GB02_LOOKUPVALUE " & _
              "  WHERE " & _
              "    GB02_TIPOCAMPO = 'TXT_GB06_TIPOAREA'"


    
    Str_Caption = "Elenco"
    Arr_Fields = Pst_Colonne
    Str_Connect = Gstr_Connect
    TXT_GB06_TIPOOFFERTA.IDLookup = "Lkp_tipo"
End Sub

Private Sub TXT_GB06_TIPOOFFERTA_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
  If NVL(TXT_GB06_TIPOOFFERTA.Text, "") = "" Then
      Exit Sub
  End If

  Cancel = False

  ReDim Arr_Fields(0 To 0, 0 To 1)
  Arr_Fields(0, 1) = "GB02_DESCRIZIONE"
  Set Arr_Fields(0, 0) = TXT_GB06_TIPOOFFERTA_DEC
  
  'Imposto la stringa SQL
  '
  
  str_SQL = " SELECT  GB02_DESCRIZIONE "
  str_SQL = str_SQL & " FROM  GB02_LOOKUPVALUE "
  str_SQL = str_SQL & " WHERE  "
  str_SQL = str_SQL & "   GB02_CODICE = '" & TXT_GB06_TIPOOFFERTA.Text & "'"
  str_SQL = str_SQL & "   AND GB02_TIPOCAMPO  = 'TXT_GB06_TIPOOFFERTA' "
  
  
 
  Str_Connect = Gstr_Connect
  Err.Clear
End Sub

Private Sub TXT_GB06_TIPOOFFERTA_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
On Error Resume Next
    
    Dim Pst_Colonne(0 To 1, 0 To 1) As Variant
    
    Pst_Colonne(0, 0) = "Codice"
    Pst_Colonne(0, 1) = ""
    Pst_Colonne(1, 0) = "Descrizione"
    Pst_Colonne(1, 1) = ""
    

    
    str_SQL = " SELECT" & _
              "    GB02_CODICE, " & _
              "    GB02_DESCRIZIONE " & _
              " FROM        GB02_LOOKUPVALUE " & _
              "  WHERE " & _
              "    GB02_TIPOCAMPO = 'TXT_GB06_TIPOOFFERTA'"


    
    Str_Caption = "Elenco"
    Arr_Fields = Pst_Colonne
    Str_Connect = Gstr_Connect
    TXT_GB06_TIPOOFFERTA.IDLookup = "Lkp_tipo"
End Sub

Private Sub TXT_GB07_IMAGEPATH_Change()
'Call InserisciAllegati
End Sub

Private Sub TXT_GB07_PREZZO_AfterChange(Cancel As Boolean)
Call RicalcolaImporto
TXT_GB07_PERCPROVV.Text = RecuperaProvvigione(QGridDocumenti.DataSource("GB07_PREZZO").value, QGridDocumenti.DataSource("GB07_IMPORTO").value, NVL(QGridDocumenti.DataSource("MG66_FAM_MG53").value, ""), NVL(QGridDocumenti.DataSource("MG66_SFAM_MG54").value, ""), TXT_GB06_AGENTE.Text, NVL(QGridDocumenti.DataSource("GB07_SC1").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC2").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC3").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC4").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC5").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC6").value, 0))
End Sub

Private Sub TXT_GB07_QTA_AfterChange(Cancel As Boolean)
If NVL(TXT_GB07_CODART_MG66.Text, "") <> "" Then
Call RicalcolaImporto
End If
End Sub

Private Sub TXT_GB07_RAG_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)

On Error Resume Next
    
    Dim Pst_Colonne(0 To 1, 0 To 1) As Variant
    
    Pst_Colonne(0, 0) = "Codice"
    Pst_Colonne(0, 1) = ""
    Pst_Colonne(1, 0) = "Descrizione"
    Pst_Colonne(1, 1) = ""
    

    
    str_SQL = " SELECT" & _
              "    GB01_PROG, " & _
              "    GB01_DESCRIZIONE " & _
              " FROM        GB01_GRUPPIOFFERTA " & _
              "  WHERE " & _
              "    GB01_NUMDOC_GB06 = " & TXT_GB06_NUMDOC.Text


    
    Str_Caption = "Elenco Gruppi"
    Arr_Fields = Pst_Colonne
    Str_Connect = Gstr_Connect
    TXT_GB07_RAG.IDLookup = "Lkp_Gruppi"
End Sub

Private Sub TXT_GB07_SC1_AfterChange(Cancel As Boolean)
 Call RicalcolaImporto
 TXT_GB07_PERCPROVV.Text = RecuperaProvvigione(QGridDocumenti.DataSource("GB07_PREZZO").value, QGridDocumenti.DataSource("GB07_IMPORTO").value, NVL(QGridDocumenti.DataSource("MG66_FAM_MG53").value, ""), NVL(QGridDocumenti.DataSource("MG66_SFAM_MG54").value, ""), TXT_GB06_AGENTE.Text, NVL(QGridDocumenti.DataSource("GB07_SC1").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC2").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC3").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC4").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC5").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC6").value, 0))
End Sub

Private Sub TXT_GB07_SC2_AfterChange(Cancel As Boolean)
 Call RicalcolaImporto
 TXT_GB07_PERCPROVV.Text = RecuperaProvvigione(QGridDocumenti.DataSource("GB07_PREZZO").value, QGridDocumenti.DataSource("GB07_IMPORTO").value, NVL(QGridDocumenti.DataSource("MG66_FAM_MG53").value, ""), NVL(QGridDocumenti.DataSource("MG66_SFAM_MG54").value, ""), TXT_GB06_AGENTE.Text, NVL(QGridDocumenti.DataSource("GB07_SC1").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC2").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC3").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC4").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC5").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC6").value, 0))
End Sub

Private Sub TXT_GB07_SC3_AfterChange(Cancel As Boolean)
 Call RicalcolaImporto
 TXT_GB07_PERCPROVV.Text = RecuperaProvvigione(QGridDocumenti.DataSource("GB07_PREZZO").value, QGridDocumenti.DataSource("GB07_IMPORTO").value, NVL(QGridDocumenti.DataSource("MG66_FAM_MG53").value, ""), NVL(QGridDocumenti.DataSource("MG66_SFAM_MG54").value, ""), TXT_GB06_AGENTE.Text, NVL(QGridDocumenti.DataSource("GB07_SC1").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC2").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC3").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC4").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC5").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC6").value, 0))
End Sub

Private Sub TXT_GB07_SC4_AfterChange(Cancel As Boolean)
 Call RicalcolaImporto
 TXT_GB07_PERCPROVV.Text = RecuperaProvvigione(QGridDocumenti.DataSource("GB07_PREZZO").value, QGridDocumenti.DataSource("GB07_IMPORTO").value, QGridDocumenti.DataSource("MG66_FAM_MG53").value, NVL(QGridDocumenti.DataSource("MG66_SFAM_MG54").value, ""), TXT_GB06_AGENTE.Text, NVL(QGridDocumenti.DataSource("GB07_SC1").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC2").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC3").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC4").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC5").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC6").value, 0))
End Sub

Private Sub TXT_GB07_SC5_AfterChange(Cancel As Boolean)
 Call RicalcolaImporto
 TXT_GB07_PERCPROVV.Text = RecuperaProvvigione(QGridDocumenti.DataSource("GB07_PREZZO").value, QGridDocumenti.DataSource("GB07_IMPORTO").value, QGridDocumenti.DataSource("MG66_FAM_MG53").value, NVL(QGridDocumenti.DataSource("MG66_SFAM_MG54").value, ""), TXT_GB06_AGENTE.Text, NVL(QGridDocumenti.DataSource("GB07_SC1").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC2").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC3").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC4").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC5").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC6").value, 0))
End Sub

Private Sub TXT_GB07_SC6_AfterChange(Cancel As Boolean)
 Call RicalcolaImporto
 TXT_GB07_PERCPROVV.Text = RecuperaProvvigione(QGridDocumenti.DataSource("GB07_PREZZO").value, QGridDocumenti.DataSource("GB07_IMPORTO").value, QGridDocumenti.DataSource("MG66_FAM_MG53").value, NVL(QGridDocumenti.DataSource("MG66_SFAM_MG54").value, ""), TXT_GB06_AGENTE.Text, NVL(QGridDocumenti.DataSource("GB07_SC1").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC2").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC3").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC4").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC5").value, 0), NVL(QGridDocumenti.DataSource("GB07_SC6").value, 0))
End Sub

Private Sub TXT_PROPRIETARIO_BeforeItem(Cancel As Boolean)

End Sub



Private Sub TXT_IMAGEPATH_AfterItem(Cancel As Boolean)
'InserisciAllegati
End Sub
Private Sub InserisciAllegati()

Dim rst As New ADODB.Recordset
Dim strSQL As String
    
Dim NomeFile As String
Dim TipoFile As String
Dim sArray() As String

Dim Contatore As Integer


Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim mystream As ADODB.Stream
Set mystream = New ADODB.Stream
mystream.Type = adTypeBinary
On Error Resume Next
 strSQL = "SELECT * FROM GB07_CORPODOC WHERE GB07_ID = " & NVL(rstCorpoBanco("GB07_ID"), 0) 'TXT_GB07_ID.Text
                    
                    rs.Open strSQL, Gcon_Connect, adOpenStatic, adLockOptimistic
                    
                   

                    
                    mystream.Open
                    mystream.LoadFromFile Mid(TXT_GB07_IMAGEPATH.Text, 4, Len(TXT_GB07_IMAGEPATH.Text))

                    'Gestire il size del file
                    
                    If mystream.Size <= 30408704 Then ' 29Mb

                        
                        rs!GB07_IMG = mystream.Read
                       
    
                        rs.Update
                    Else
                       ' Call ScriviLog("La dimensione dell' allegato supera il limite massimo consentito per l'invio tramite PEC. " & NomeFile & OrdInElab)
                    End If
                    
                    
                    
                    mystream.Close
                    rs.Close
End Sub

Private Sub TXT_GRST1_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
On Error Resume Next
    Call Cls_DecodeMagazzino.GruppoStatistico1(TXT_GRST1.Text)
    str_SQL = Cls_DecodeMagazzino.StringaSQL
    Arr_Fields = Cls_DecodeMagazzino.ArrayFields
    
    Set Arr_Fields(0, 0) = TXT_GRST1_DEC
    
    Str_Connect = Gstr_Connect
End Sub

Private Sub TXT_GRST1_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
 On Error Resume Next
  
  Call Cls_LookupMagazzino.GruppiStatistici1
  str_SQL = Cls_LookupMagazzino.StringaSQL
  Arr_Fields = Cls_LookupMagazzino.ArrayFields
  Str_Caption = Cls_LookupMagazzino.Titolo
  Str_Connect = Gstr_Connect
  TXT_GRST1.IDLookup = Cls_LookupMagazzino.IDLookup
End Sub

Private Sub TXT_GRST2_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
On Error Resume Next
    Call Cls_DecodeMagazzino.GruppoStatistico2(TXT_GRST2.Text)
    str_SQL = Cls_DecodeMagazzino.StringaSQL
    Arr_Fields = Cls_DecodeMagazzino.ArrayFields
    
    Set Arr_Fields(0, 0) = TXT_GRST2_DEC
    
    Str_Connect = Gstr_Connect
End Sub

Private Sub TXT_GRST2_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
On Error Resume Next
  
  Call Cls_LookupMagazzino.GruppiStatistici2
  str_SQL = Cls_LookupMagazzino.StringaSQL
  Arr_Fields = Cls_LookupMagazzino.ArrayFields
  Str_Caption = Cls_LookupMagazzino.Titolo
  Str_Connect = Gstr_Connect
  TXT_GRST2.IDLookup = Cls_LookupMagazzino.IDLookup
End Sub

Private Sub TXT_GRST3_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
On Error Resume Next
    Call Cls_DecodeMagazzino.GruppoStatistico3(TXT_GRST3.Text)
    str_SQL = Cls_DecodeMagazzino.StringaSQL
    Arr_Fields = Cls_DecodeMagazzino.ArrayFields
    
    Set Arr_Fields(0, 0) = TXT_GRST3_DEC
    
    Str_Connect = Gstr_Connect
End Sub

Private Sub TXT_GRST3_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
On Error Resume Next
  
  Call Cls_LookupMagazzino.GruppiStatistici3
  str_SQL = Cls_LookupMagazzino.StringaSQL
  Arr_Fields = Cls_LookupMagazzino.ArrayFields
  Str_Caption = Cls_LookupMagazzino.Titolo
  Str_Connect = Gstr_Connect
  TXT_GRST3.IDLookup = Cls_LookupMagazzino.IDLookup
End Sub


Private Sub TXT_GRST4_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
On Error Resume Next
    Call Cls_DecodeMagazzino.GruppoStatistico4(TXT_GRST4.Text)
    str_SQL = Cls_DecodeMagazzino.StringaSQL
    Arr_Fields = Cls_DecodeMagazzino.ArrayFields
    
    Set Arr_Fields(0, 0) = TXT_GRST4_DEC
    
    Str_Connect = Gstr_Connect
End Sub

Private Sub TXT_GRST4_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
On Error Resume Next
  
  Call Cls_LookupMagazzino.GruppiStatistici4
  str_SQL = Cls_LookupMagazzino.StringaSQL
  Arr_Fields = Cls_LookupMagazzino.ArrayFields
  Str_Caption = Cls_LookupMagazzino.Titolo
  Str_Connect = Gstr_Connect
  TXT_GRST4.IDLookup = Cls_LookupMagazzino.IDLookup
End Sub

Private Sub TXT_GRUPPO_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
'On Error Resume Next
'    Call Cls_DecodeMagazzino.Gruppo(NVL(TXT_FAM.Text, ""), NVL(TXT_SFAM.Text, ""), NVL(TXT_GRUPPO.Text, ""))
'    str_SQL = Cls_DecodeMagazzino.StringaSQL
'    Arr_Fields = Cls_DecodeMagazzino.ArrayFields
'
'    Set Arr_Fields(0, 0) = TXT_GRUPPO_DEC
'
'    Str_Connect = Gstr_Connect
If NVL(TXT_GRUPPO.Text, "") = "" Then
      Exit Sub
  End If

  Cancel = False

  ReDim Arr_Fields(0 To 0, 0 To 1)
  Arr_Fields(0, 1) = "MG55_DESCRGRUPPO"
  Set Arr_Fields(0, 0) = TXT_GRUPPO_DEC
  
  'Imposto la stringa SQL
  '
  
  str_SQL = " SELECT  MG55_DESCRGRUPPO "
  str_SQL = str_SQL & " FROM MG55_GRUPPI "
  str_SQL = str_SQL & " WHERE MG55_DITTA_CG18  = " & CodiceDitta
  str_SQL = str_SQL & "   AND MG55_CODGRUPPO = '" & TXT_GRUPPO.Text & "'"
  
  
  
 
  Str_Connect = Gstr_Connect
  Err.Clear
End Sub

Private Sub TXT_GRUPPO_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
'  On Error Resume Next
'
'  Call Cls_LookupMagazzino.Gruppi(TXT_FAM.Text, TXT_SFAM.Text)
'  str_SQL = Cls_LookupMagazzino.StringaSQL
'  Arr_Fields = Cls_LookupMagazzino.ArrayFields
'  Str_Caption = Cls_LookupMagazzino.Titolo
'  Str_Connect = Gstr_Connect
'  TXT_GRUPPO.IDLookup = Cls_LookupMagazzino.IDLookup

On Error Resume Next

  Cancel = False
  str_SQL = " SELECT  DISTINCT   MG55_CODGRUPPO, MG55_DESCRGRUPPO "
  str_SQL = str_SQL & " FROM MG55_GRUPPI "
  str_SQL = str_SQL & " WHERE MG55_DITTA_CG18  = " & CodiceDitta
  
  
  ReDim Arr_Fields(0 To 1, 0 To 1)
  Arr_Fields(0, 0) = "Codice"
  Arr_Fields(0, 1) = ""
  Arr_Fields(1, 0) = "Descrizione"
  Arr_Fields(1, 1) = ""

  Str_Caption = "Gruppi"
  Str_Connect = Gstr_Connect
  TXT_GRUPPO = "lkp_Gruppi"
  
  Err.Clear
End Sub

Private Sub TXT_RICERCANC_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
'Ricerca documenti fattura o scontrini
    On Error Resume Next
    
    Dim Pst_Colonne(0 To 6, 0 To 1) As Variant
    
    Pst_Colonne(0, 0) = "Numero registrazione"
    Pst_Colonne(0, 1) = ""
    Pst_Colonne(1, 0) = "Tipo documento"
    Pst_Colonne(1, 1) = ""
    Pst_Colonne(2, 0) = "Numero documento"
    Pst_Colonne(2, 1) = ""
    Pst_Colonne(3, 0) = "Data documento"
    Pst_Colonne(3, 1) = ""
    Pst_Colonne(4, 0) = "Sezionale documento"
    Pst_Colonne(4, 1) = ""
    Pst_Colonne(5, 0) = "Codice Cliente"
    Pst_Colonne(5, 1) = ""
    Pst_Colonne(6, 0) = "Ragione sociale"
    Pst_Colonne(6, 1) = ""
    
    str_SQL = " SELECT" & _
              "    DO11_NUMREG_CO99," & _
              "    DO11_DOCUM_MG36," & _
              "    DO11_NUMDOC," & _
              "    DO11_DATADOC," & _
              "    DO11_SEZDOC," & _
              "    DO11_CLIFOR_CG44," & _
              "    CG16_RAGSOANAG "
    str_SQL = str_SQL & " FROM        DO11_DOCTESTATA "
    str_SQL = str_SQL & "  INNER JOIN CG44_CLIFOR "
    str_SQL = str_SQL & "          ON DO11_DOCTESTATA.DO11_DITTACF_CG44 = CG44_CLIFOR.CG44_DITTA_CG18 "
    str_SQL = str_SQL & "         AND DO11_DOCTESTATA.DO11_TIPOCF_CG44 = CG44_CLIFOR.CG44_TIPOCF "
    str_SQL = str_SQL & "         AND DO11_DOCTESTATA.DO11_CLIFOR_CG44 = CG44_CLIFOR.CG44_CLIFOR "
    str_SQL = str_SQL & "  INNER JOIN CG16_ANAGGEN "
    str_SQL = str_SQL & "          ON CG44_CLIFOR.CG44_CODICE_CG16 = CG16_ANAGGEN.CG16_CODICE"
    str_SQL = str_SQL & "  WHERE " & _
                        "    DO11_DITTA_CG18 = " & CodiceDitta & " AND" & _
                        "    DO11_TIPOCF_CG44 = 0 "
              
'Scontrino 11, 13
'Fatture 3, 4, 5
    str_SQL = str_SQL & " AND DO11_TIPODOC IN (3,4,5,11,13 )"
    
    Str_Caption = "Elenco documenti fatture e scontrini"
    Arr_Fields = Pst_Colonne
    Str_Connect = Gstr_Connect
   ' TXT_RICERCANC.IDLookup = "Lkp_Documenti"

End Sub

Private Sub TXT_RICERCARESI_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
'Ricerca documenti fattura o scontrini
    On Error Resume Next
    
    Dim Pst_Colonne(0 To 6, 0 To 1) As Variant
    
    Pst_Colonne(0, 0) = "Numero registrazione"
    Pst_Colonne(0, 1) = ""
    Pst_Colonne(1, 0) = "Tipo documento"
    Pst_Colonne(1, 1) = ""
    Pst_Colonne(2, 0) = "Numero documento"
    Pst_Colonne(2, 1) = ""
    Pst_Colonne(3, 0) = "Data documento"
    Pst_Colonne(3, 1) = ""
    Pst_Colonne(4, 0) = "Sezionale documento"
    Pst_Colonne(4, 1) = ""
    Pst_Colonne(5, 0) = "Codice Cliente"
    Pst_Colonne(5, 1) = ""
    Pst_Colonne(6, 0) = "Ragione sociale"
    Pst_Colonne(6, 1) = ""
    
    str_SQL = " SELECT" & _
              "    DO11_NUMREG_CO99," & _
              "    DO11_DOCUM_MG36," & _
              "    DO11_NUMDOC," & _
              "    DO11_DATADOC," & _
              "    DO11_SEZDOC," & _
              "    DO11_CLIFOR_CG44," & _
              "    CG16_RAGSOANAG "
    
    str_SQL = str_SQL & " FROM        DO11_DOCTESTATA "
    str_SQL = str_SQL & "  INNER JOIN CG44_CLIFOR "
    str_SQL = str_SQL & "          ON DO11_DOCTESTATA.DO11_DITTACF_CG44 = CG44_CLIFOR.CG44_DITTA_CG18 "
    str_SQL = str_SQL & "         AND DO11_DOCTESTATA.DO11_TIPOCF_CG44 = CG44_CLIFOR.CG44_TIPOCF "
    str_SQL = str_SQL & "         AND DO11_DOCTESTATA.DO11_CLIFOR_CG44 = CG44_CLIFOR.CG44_CLIFOR "
    str_SQL = str_SQL & "  INNER JOIN CG16_ANAGGEN "
    str_SQL = str_SQL & "          ON CG44_CLIFOR.CG44_CODICE_CG16 = CG16_ANAGGEN.CG16_CODICE"
    str_SQL = str_SQL & "  WHERE " & _
                        "    DO11_DITTA_CG18 = " & CodiceDitta & " AND" & _
                        "    DO11_TIPOCF_CG44 = 0 "

              
'DDT 1
    str_SQL = str_SQL & " AND DO11_TIPODOC IN (1 )"
    
    Str_Caption = "Elenco documenti ddt"
    Arr_Fields = Pst_Colonne
    Str_Connect = Gstr_Connect
   ' TXT_RICERCARESI.IDLookup = "Lkp_Documenti"

End Sub


Private Sub ScriviGB06()
'
  Dim DateInserimento As Date

  Dim CurUserName As String
  CurUserName = ActualComputerName & ActualUserName
  DateInserimento = TXT_ANNOINSERIMENTO.Text
  ProgInsertTestata = CurUserName & Year(Now) & String(2 - Len(Month(Now)), "0") & Month(Now) & String(2 - Len(Day(Now)), "0") & Day(Now) & String(2 - Len(Hour(Time)), "0") & Hour(Time) & String(2 - Len(Minute(Time)), "0") & Minute(Time) & String(2 - Len(Second(Time)), "0") & Second(Time) & Right(Format(Timer, "#0.00"), 2)
  
  strSQL = "INSERT INTO GB06_TESTADOC "
  strSQL = strSQL & " ( GB06_PROG, GB06_DITTA_CG18, GB06_TIPODOC, GB06_DATA, GB06_GRUPPOCRE, GB06_UTENTECRE, GB06_STATODOC "
  strSQL = strSQL & " ) VALUES ( "
  strSQL = strSQL & "'" & ProgInsertTestata & "',"
  strSQL = strSQL & CodiceDitta & ","
  strSQL = strSQL & TipoDocumento & ","
  strSQL = strSQL & " CONVERT(DateTime,'" & Format(DateInserimento, "dd/mm/yyyy") & "',103),"
  strSQL = strSQL & "  '" & ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Gruppo & "',"
  strSQL = strSQL & "  '" & ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Descrizione & "',"
  strSQL = strSQL & "  '00'"
  strSQL = strSQL & " )   "
  
  Call Gcon_Connect.Execute(strSQL, , adCmdText)
  
  
End Sub

Private Function ActualComputerName() As String
Dim strComputerName As String

  strComputerName = String(100, Chr$(0))
  GetComputerName strComputerName, 100
  strComputerName = Left$(strComputerName, InStr(strComputerName, Chr$(0)) - 1)
  ActualComputerName = strComputerName

End Function


Private Function ActualUserName()
Dim strUserName As String
  
  'Create a buffer
  strUserName = String(100, Chr$(0))
  'Get the username
  GetUserName strUserName, 100
  'strip the rest of the buffer
  strUserName = Left$(strUserName, InStr(strUserName, Chr$(0)) - 1)
  ActualUserName = strUserName

End Function


Private Function LeggiIDGB06()
'
  On Error GoTo ErrTrap
  
  Dim MyRst       As ADODB.Recordset
  
  strSQL = "SELECT GB06_ID "
  strSQL = strSQL & "  FROM GB06_TESTADOC "
  strSQL = strSQL & " WHERE GB06_PROG       = '" & ProgInsertTestata & "'"
  strSQL = strSQL & "   AND GB06_DITTA_CG18 = " & CodiceDitta
  strSQL = strSQL & "   AND GB06_TIPODOC    = " & TipoDocumento
  strSQL = strSQL & "   AND GB06_DATA       = CONVERT(DateTime,'" & Format(CDate(TXT_ANNOINSERIMENTO.Text), "dd/mm/yyyy") & "',103)"
  
  Set MyRst = Gcon_Connect.Execute(strSQL, , adCmdText)
  If Not MyRst.EOF Then
      IDGB06 = NVL(MyRst.Fields(0).value, "")
      TXT_GB06_ID.Text = CStr(IDGB06)
      TXT_GB06_NREV.Text = 1
      TXT_GB06_NVERS.Text = 1
  Else
      IDGB06 = 0
  End If
  
  strSQL = ""
  Set MyRst = Nothing

  
  Exit Function
ErrTrap:
  Select Case VisualizzaErrore("LeggiIDGB06")
    Case vbAbort
        Exit Function
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
  End Select

End Function


Private Function GetFamArt(CodArt As String)
'
  On Error GoTo ErrTrap
  
  Dim MyRst       As ADODB.Recordset
  
  strSQL = "SELECT       MG66_FAM_MG53 " & _
           " From MG66_ANAGRART " & _
           " WHERE (MG66_DITTA_CG18 = " & CodiceDitta & ") AND (MG66_CODART = '" & CodArt & "') "
  
  Set MyRst = Gcon_Connect.Execute(strSQL, , adCmdText)
  If Not MyRst.EOF Then
      GetFamArt = NVL(MyRst.Fields(0).value, "")
  Else
      GetFamArt = ""
  End If
  
  strSQL = ""
  Set MyRst = Nothing

  
  Exit Function
ErrTrap:
  Select Case VisualizzaErrore("LeggiIDGB06")
    Case vbAbort
        Exit Function
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
  End Select

End Function


Private Function GetPzConfArt(CodArt As String, Opzione As String) As Double
'
  On Error GoTo ErrTrap
  
  Dim MyRst       As ADODB.Recordset
  
  strSQL = " SELECT TOP (1) MG68_PZCONF " & _
           " From MG68_CONFART " & _
           " WHERE (MG68_CODART_MG66 = '" & CodArt & "') AND (MG68_DITTA_CG18 = " & CodiceDitta & ") AND (MG68_OPZIONE_MG5E = '" & Opzione & "')"
  
  Set MyRst = Gcon_Connect.Execute(strSQL, , adCmdText)
  If Not MyRst.EOF Then
      GetPzConfArt = NVL(MyRst.Fields(0).value, 1)
  Else
      GetPzConfArt = 1
  End If
  
  strSQL = ""
  Set MyRst = Nothing

  
  Exit Function
ErrTrap:
  Select Case VisualizzaErrore("LeggiIDGB06")
    Case vbAbort
        Exit Function
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
  End Select

End Function

Private Function DistruggiFramework() As Boolean
  On Error GoTo ErrTrap
    
    DistruggiFramework = False
    If FME_BANCO Is Nothing Then Exit Function
    
    If FME_BANCO.Status <> tsInsert Then
'      DistruggiFramework = (FME_BANCO.Update(True, False, True) = tsMethodCanceled)
    End If
    
    If Not rstCorpoBanco Is Nothing Then
        If rstCorpoBanco.State = adStateOpen Then rstCorpoBanco.Close
        Set rstCorpoBanco = Nothing
    End If
    FME_BANCO.Terminate
    Set FME_BANCO = Nothing
    Set GridNavDocumenti.ActiveDll = Nothing
    Set GridNavDocumenti.ActiveFrame = Nothing

  Exit Function
ErrTrap:
  Select Case VisualizzaErrore("DistruggiFramework")
    Case vbAbort
      Exit Function
    Case vbRetry
      Resume
    Case vbIgnore
      Resume Next
  End Select

End Function





Private Sub ImpostaVirtualFrame()
  On Error GoTo ErrTrap
  
  'Variabili String
  Dim StringaSQL      As String
  
  'Istanze Oggetto Recordset
  Dim rstSql          As New ADODB.Recordset
 ' MsgBox QGridDocumenti.DataSource(0)
'  StringaSQL = " SELECT * FROM GB07_CORPODOC WHERE 1 = 0"
 'StringaSQL = StringaSQL & " Inner Join MG66_ANAGRART ON GB07_CORPODOC.GB07_CODART_MG66 = MG66_ANAGRART.MG66_CODART AND MG66_ANAGRART.MG66_DITTA_CG18 = " & CodiceDitta
  StringaSQL = " SELECT * , CASE WHEN GB07_CLIFOR_CG44 IS NULL THEN 2 ELSE 1 END  AS IMGCODFOR,  CASE WHEN GB07_IMG IS NULL THEN 0 ELSE 1 END  AS FLGIMG, GB07_IMPORTO / GB07_QTA AS GB07_IMPORTOUNI,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_GRUSTAT1_MG74  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CORPODOC.GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_GRUSTAT1_MG74,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_GRUSTAT2_MG75  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CORPODOC.GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_GRUSTAT2_MG75,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_GRUSTAT3_MG76  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CORPODOC.GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_GRUSTAT3_MG76,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_GRUSTAT4_MG77  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CORPODOC.GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_GRUSTAT4_MG77,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_FAM_MG53  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CORPODOC.GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_FAM_MG53,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_SFAM_MG54  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CORPODOC.GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_SFAM_MG54,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_GRUPPO_MG55  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CORPODOC.GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_GRUPPO_MG55,  "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_SGRUPPO_MG56  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CORPODOC.GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_SGRUPPO_MG56 , "
  
  StringaSQL = StringaSQL & " (SELECT        MG66_UM1  "
  StringaSQL = StringaSQL & " From MG66_ANAGRART "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CORPODOC.GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) as MG66_UM1 , "

  
  StringaSQL = StringaSQL & " (SELECT       CO5L_ATTRIBUTIDEN.CO5L_ALF2   "
  StringaSQL = StringaSQL & " FROM            CO5L_ATTRIBUTIDEN INNER JOIN              MG66_ANAGRART ON CO5L_ATTRIBUTIDEN.CO5L_GUID = MG66_ANAGRART.MG66_GUID "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CORPODOC.GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) AS Area,  "

  StringaSQL = StringaSQL & " (SELECT       CO5L_ATTRIBUTIDEN.CO5L_ALF3   "
  StringaSQL = StringaSQL & " FROM            CO5L_ATTRIBUTIDEN INNER JOIN              MG66_ANAGRART ON CO5L_ATTRIBUTIDEN.CO5L_GUID = MG66_ANAGRART.MG66_GUID "
  StringaSQL = StringaSQL & " WHERE        (MG66_CODART = GB07_CORPODOC.GB07_CODART_MG66) AND (MG66_DITTA_CG18 = " & CodiceDitta & ")) AS HCL  "

  
  
  StringaSQL = StringaSQL & " FROM GB07_CORPODOC "
  StringaSQL = StringaSQL & " WHERE GB07_ID_GB06 = " & IDGB06
  StringaSQL = StringaSQL & " ORDER BY   GB07_RAG, GB07_SEQ "
  
   'New STANDARD class for managing records
  Set Gcls_RecordS = New CLSFW_Recordset
  
  'using Gpr_GetADORecord method we have a recorset with
  ' standard properties
  Set rstCorpoBanco = Gcls_RecordS.Gpr_GetADORecord
  
  'opening recordset
  With rstCorpoBanco
    Set .ActiveConnection = Gcon_Connect
    .Source = StringaSQL
    .Open
    .MarshalOptions = adMarshalModifiedOnly
    
  End With
  
  'New Virtual Frame
  Set FME_BANCO = New FWBO_VIRTUALFRAME.CLSFW_VIRTUALFRAME
 ''''' If rstCorpoBanco.RecordCount <= 0 Then Exit Sub
'  Set GRID_CONTRATTI.DataSource = rstContratti
  'rstCorpoBanco.AbsolutePosition
  'Virtual Frame Initialization
  'FME_BANCO.ValueBookmark = 3
  FME_BANCO.Initialize ActiveInterface, Gcon_Connect, rstCorpoBanco, StringaSQL, "GB07_ID", GridNavDocumenti, QGridDocumenti
  
  ' binding components
  FME_BANCO.AddControl TXT_GB07_CODART_MG66
  FME_BANCO.AddControl TXT_GB07_QTA
  FME_BANCO.AddControl TXT_DESART
  FME_BANCO.AddControl TXT_GB07_SC1
  FME_BANCO.AddControl TXT_GB07_SC2
  FME_BANCO.AddControl TXT_GB07_SC3
  FME_BANCO.AddControl TXT_GB07_SC4
  FME_BANCO.AddControl TXT_GB07_SC5
  FME_BANCO.AddControl TXT_GB07_SC6
  FME_BANCO.AddControl TXT_GB07_ALT
  FME_BANCO.AddControl TXT_GB07_RAG
  FME_BANCO.AddControl TXT_GB07_SEQ
  FME_BANCO.AddControl TXT_GB07_PREZZO
  FME_BANCO.AddControl TXT_GB07_COSTO
  FME_BANCO.AddControl TXT_GB07_IMPORTO
  FME_BANCO.AddControl TXT_GB07_IMAGEPATH
  FME_BANCO.AddControl TXT_GB07_FLPOSA
  FME_BANCO.AddControl TXT_GB07_CLIFOR_CG44
  FME_BANCO.AddControl TXT_GB07_TIPOCF_CG44
  FME_BANCO.AddControl CHK_GB07_CHECK
  FME_BANCO.AddControl TXT_GB07_PERCPROVV
  FME_BANCO.AddControl TXT_GB07_IMPPROVV

  
  
  
  
  ' InternavContratti initialization
  Set GridNavDocumenti.ActiveDll = ActiveInterface
  Set GridNavDocumenti.ActiveFrame = FME_BANCO
  
  FME_BANCO.MsgOnUpdate = False
  
  If rstCorpoBanco.RecordCount > 0 Then
    FME_BANCO.Status = tsModify
  Else
    FME_BANCO.Status = tsInsert
  End If
  
  
  ' InternavContratti button
  GridNavDocumenti.Indietro = False
  GridNavDocumenti.Avanti = False
  GridNavDocumenti.Nuovo = True
  GridNavDocumenti.Elimina = True
  GridNavDocumenti.Annulla = True
  GridNavDocumenti.Conferma = True
  GridNavDocumenti.Apri = False
 ' QGridDocumenti.DataSource.AbsolutePosition = adPosEOF
  Exit Sub
ErrTrap:
      Select Case VisualizzaErrore("ImpostaVirtualFrame")
    Case vbAbort
      Exit Sub
    Case vbRetry
      Resume
    Case vbIgnore
      Resume Next
  End Select
End Sub

Private Sub ImpostaVirtualFrame_Note()
Dim StringaSQL As String

    On Error GoTo ErrTrap
    'SELECT     TOP (200) GB0A_DITTA_CG18, GB0A_CODCESPITE_CS04, GB0A_CODLEASING, GB0A_CONTODEST_PC03, GB0A_NOTE
    'From GB0A_MULTILEASING
    'SQL string
    StringaSQL = "SELECT     * " & _
    " FROM            GB08_NOTEOFFERTA " & _
    " where GB08_ID_GB06 = " & IDGB06


    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstNOTE = Gcls_RecordPadre.Gpr_GetADORecord
    With rstNOTE
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With

    Set FME_NOTE = New FWBO_VIRTUALFRAME.CLSFW_VIRTUALFRAME
    'Virtual Frame Initialization
    FME_NOTE.Initialize ActiveInterface, Gcon_Connect, rstNOTE, StringaSQL, "GB08_ID", TMS_GRIDNAV_NOTE, QGRID_NOTE

    ' binding components
   FME_NOTE.AddControl TXT_GB08_DATA
   FME_NOTE.AddControl TXT_GB08_TESTONOTA


    ' Controller initialization
    Set TMS_GRIDNAV_NOTE.ActiveDll = ActiveInterface
    Set TMS_GRIDNAV_NOTE.ActiveFrame = FME_NOTE

    ' GRIDNAV button
    TMS_GRIDNAV_NOTE.Indietro = False
    TMS_GRIDNAV_NOTE.Avanti = False
    TMS_GRIDNAV_NOTE.Apri = False

    If rstNOTE.RecordCount > 0 Then
          FME_NOTE.Status = ActiveInterface.ProgramMode
    Else
          FME_NOTE.Status = tsInsert
    End If

    INITGRID_NOTE
    QGRID_NOTE.BeginDataSourceSuspend

    Set QGRID_NOTE.DataSource = rstNOTE
    QGRID_NOTE.EndDataSourceSuspend
    QGRID_NOTE.Refresh

        'TXT_GB05_NRATA.SetFocus


       ' TXT_GB05_NRATA.Enabled = False


    Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("ImpostaVirtualFrame")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub

Private Sub ImpostaVirtualFrame_Gruppi()
Dim StringaSQL As String

    On Error GoTo ErrTrap
    'SELECT     TOP (200) GB0A_DITTA_CG18, GB0A_CODCESPITE_CS04, GB0A_CODLEASING, GB0A_CONTODEST_PC03, GB0A_NOTE
    'From GB0A_MULTILEASING
    'SQL string
    StringaSQL = "SELECT     * " & _
    " FROM            GB01_GRUPPIOFFERTA " & _
    " where GB07_ID_GB06 = '" & NVL(IDGB06, 0) & "'"


    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGRUPPI = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGRUPPI
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With

    Set FME_GRUPPI = New FWBO_VIRTUALFRAME.CLSFW_VIRTUALFRAME
    'Virtual Frame Initialization
    FME_GRUPPI.Initialize ActiveInterface, Gcon_Connect, rstGRUPPI, StringaSQL, "", TMS_GRIDNAV_GRUPPI, TMS_QGRIDGRUPPI

    ' binding components
   FME_GRUPPI.AddControl TXT_GB01_PROG
   FME_GRUPPI.AddControl TXT_GB01_DESCRIZIONE


    ' Controller initialization
    Set TMS_GRIDNAV_GRUPPI.ActiveDll = ActiveInterface
    Set TMS_GRIDNAV_GRUPPI.ActiveFrame = FME_GRUPPI

    ' GRIDNAV button
    TMS_GRIDNAV_GRUPPI.Indietro = False
    TMS_GRIDNAV_GRUPPI.Avanti = False
    TMS_GRIDNAV_GRUPPI.Apri = False

    If rstGRUPPI.RecordCount > 0 Then
          FME_GRUPPI.Status = ActiveInterface.ProgramMode
    Else
          FME_GRUPPI.Status = tsInsert
    End If

    INITGRID_GRUPPI
    TMS_QGRIDGRUPPI.BeginDataSourceSuspend

    Set TMS_QGRIDGRUPPI.DataSource = rstGRUPPI
    TMS_QGRIDGRUPPI.EndDataSourceSuspend
    TMS_QGRIDGRUPPI.Refresh

        'TXT_GB05_NRATA.SetFocus


       ' TXT_GB05_NRATA.Enabled = False


    Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("ImpostaVirtualFrame")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub

Private Sub INITGRID_GRUPPI()

        Dim cInit As New TMS_QGRID.CLSFW_INITQGRID
        Set cInit.ActiveInterface = ActiveInterface
        Set cInit.RSColumns = Nothing
        cInit.KeyField = "GB01_id"
        cInit.CreateColumnsFromRS = False
        Set TMS_QGRIDGRUPPI.InitializationClass = cInit
        With TMS_QGRIDGRUPPI
                .CustomDrawCellEnabled = True
                '.INIT_ADDColumnIMAGE "GB05_CHIUSA", "Stato", ImgListRate, "", ONLYWITHIMAGE, 1080, True
                .INIT_ADDColumn "GB01_PROG", "Codice", gedTextEdit, 1514, True
                .INIT_ADDColumn "GB01_DESCRIZIONE", "Descrizione", gedTextEdit, 1514, True
                

                .InitializeSTART
                     '.MODCOL_ColType "GB05_CHIUSA", gedImageEdit
                
'                    .MODCOL_SummaryFooter "GB05_QCAPITALE", cstSum, Enable
'                    .MODCOL_SummaryFooter "GB05_QINTERESSI", cstSum, Enable
'                    .MODCOL_SummaryFooter "GB05_TOTRATA", cstSum, Enable
                    
                .InitializeEND
        End With

        Set cInit = Nothing
End Sub

Private Sub INITGRID_NOTE()

        Dim cInit As New TMS_QGRID.CLSFW_INITQGRID
        Set cInit.ActiveInterface = ActiveInterface
        Set cInit.RSColumns = Nothing
        cInit.KeyField = "GB08_id"
        cInit.CreateColumnsFromRS = False
        Set QGRID_NOTE.InitializationClass = cInit
        With QGRID_NOTE
                .CustomDrawCellEnabled = True
                '.INIT_ADDColumnIMAGE "GB05_CHIUSA", "Stato", ImgListRate, "", ONLYWITHIMAGE, 1080, True
                .INIT_ADDColumn "GB08_DATA", "Data Nota", gedTextEdit, 1514, True
                .INIT_ADDColumn "GB08_TESTONOTA", "Testo NOta", gedTextEdit, 15000, True
                .INIT_ADDColumn "GB08_OPERATORE", "Operatore", gedTextEdit, 2000, True
                

                .InitializeSTART
                     '.MODCOL_ColType "GB05_CHIUSA", gedImageEdit
                
'                    .MODCOL_SummaryFooter "GB05_QCAPITALE", cstSum, Enable
'                    .MODCOL_SummaryFooter "GB05_QINTERESSI", cstSum, Enable
'                    .MODCOL_SummaryFooter "GB05_TOTRATA", cstSum, Enable
                    
                .InitializeEND
        End With

        Set cInit = Nothing
End Sub

Private Sub INITGRID_ORDINI()

        Dim cInit As New TMS_QGRID.CLSFW_INITQGRID
        Set cInit.ActiveInterface = ActiveInterface
        Set cInit.RSColumns = Nothing
        cInit.KeyField = "GB09_id"
        cInit.CreateColumnsFromRS = False
        Set TMS_GRIDDOC.InitializationClass = cInit
        With TMS_GRIDDOC
                
                .DataFormatEnabled = True
                .CustomDrawCellEnabled = True
                '.INIT_ADDColumnIMAGE "GB05_CHIUSA", "Stato", ImgListRate, "", ONLYWITHIMAGE, 1080, True
                '.INIT_ADDColumn "GB09_ID_GB06", "Data Nota", gedTextEdit, 1514, True
                .INIT_ADDColumnIMAGE "DO11_DITTA_CG18", "Img", ImageList1, "", ONLYWITHIMAGE, 500
                .INIT_ADDColumn "GB09_NUMREG_CO99", "Numero Doc", gedTextEdit, 2000, False
                .INIT_ADDColumn "DO11_DOCUM_MG36", "Tipo Doc", gedTextEdit, 2000, True
                .INIT_ADDColumn "DO11_NUMDOC", "Num.Doc", gedDateEdit, 2000, True
                .INIT_ADDColumn "DO11_DATADOC", "Data", gedDateEdit, 2000, True
                .INIT_ADDColumn "DO11_SEZDOC", "Sez.Doc.", gedDateEdit, 2000, True
                .INIT_ADDColumn "CG16_RAGSOANAG", "Fornitore", gedDateEdit, 2000, True
                .INIT_ADDColumn "DO13_TOTIMPONIBILE", "Totale Imp.", gedDateEdit, 2000, True, , "###,###,###,##0.00"
                .INIT_ADDColumn "DO13_TOTDOCUMENTO", "Totale Doc.", gedDateEdit, 2000, True, , "###,###,###,##0.00"
                
                

                .InitializeSTART
                     '.MODCOL_ColType "GB05_CHIUSA", gedImageEdit GB09_DTCREAZIONE
                
'                    .MODCOL_SummaryFooter "GB05_QCAPITALE", cstSum, Enable
'                    .MODCOL_SummaryFooter "GB05_QINTERESSI", cstSum, Enable
'                    .MODCOL_SummaryFooter "GB05_TOTRATA", cstSum, Enable
                    
                .InitializeEND
        End With

        Set cInit = Nothing
End Sub

'Private Sub INITGRID_MARGINE(Eti As String)
'
'        Dim cInit As New TMS_QGRID.CLSFW_INITQGRID
'        Set cInit.ActiveInterface = ActiveInterface
'        Set cInit.RSColumns = Nothing
'        cInit.KeyField = "Id"
'        cInit.CreateColumnsFromRS = False
'        Set TMS_MARGINE.InitializationClass = cInit
'        With TMS_MARGINE
'
'                .DataFormatEnabled = True
'                .CustomDrawCellEnabled = True
'                '.INIT_ADDColumnIMAGE "GB05_CHIUSA", "Stato", ImgListRate, "", ONLYWITHIMAGE, 1080, True
'                '.INIT_ADDColumn "GB09_ID_GB06", "Data Nota", gedTextEdit, 1514, True
'                '.INIT_ADDColumnIMAGE "DO11_DITTA_CG18", "Img", ImageList1, "", ONLYWITHIMAGE, 500
'                .INIT_ADDColumn "Raggruppamento", "Raggruppamento", gedTextEdit, 2000, True
'                .INIT_ADDColumn "Vendita", "Vendita", gedCurrencyEdit, 2000, True
'                .INIT_ADDColumn "Acquisto", "Acquisto", gedCurrencyEdit, 2000, True
'                .INIT_ADDColumn "Margine", "Margine", gedCurrencyEdit, 2000, True
'                .INIT_ADDColumn "Provvigione", "Provvigione", gedCurrencyEdit, 2000, True
'
'
'
'
'                .InitializeSTART
'                     '.MODCOL_ColType "GB05_CHIUSA", gedImageEdit GB09_DTCREAZIONE
'
''                    .MODCOL_SummaryFooter "GB05_QCAPITALE", cstSum, Enable
''                    .MODCOL_SummaryFooter "GB05_QINTERESSI", cstSum, Enable
''                    .MODCOL_SummaryFooter "GB05_TOTRATA", cstSum, Enable
'
'                .InitializeEND
'        End With
'
'        Set cInit = Nothing
'End Sub

Private Sub ShowVirtualFrameDocumenti()
  
  '---------------------------------------------------------------------------
  ' Descr..: La presente funzione permette di reinizializzare il VirtualFrame
  '          in funzione dei valori in Input.
  ' Ritorni: //
  ' Note...: //
  '---------------------------------------------------------------------------
  
  'Variabili String
  Dim StringaSQL     As String
  
  On Error GoTo ErrShowVF

  '*********************************************************************
  ' Devono essere indicati i Campi Chiave che verranno utilizzati nella
  ' SELECT.
  '*********************************************************************
  
  StringaSQL = " SELECT *, GB07_CODART_MG66 AS DESART "
  StringaSQL = StringaSQL & "  FROM GB07_CORPODOC "
  StringaSQL = StringaSQL & " WHERE GB07_ID_GB06 = " & IDGB06
  StringaSQL = StringaSQL & " ORDER BY   GB07_CODART_MG66 "
  
  If rstCorpoBanco.State = adStateOpen Then
    rstCorpoBanco.Close
  End If
  
  rstCorpoBanco.Open StringaSQL
  FME_BANCO.ReOpen StringaSQL
  
  Set QGridDocumenti.DataSource = Nothing
  Set QGridDocumenti.DataSource = rstCorpoBanco
  QGridDocumenti.FullExpand
  
  Exit Sub
ErrShowVF:
  Select Case VisualizzaErrore("ShowVF")
    Case vbAbort
      Exit Sub
    Case vbRetry
      Resume
    Case vbIgnore
      Resume Next
  End Select
End Sub


Private Sub InitQGrid()
On Error GoTo Errore


'****************************************************************************************************
'Inizializzazione qgrid anagrafica
    Dim cInit       As New TMS_QGRID.CLSFW_INITQGRID

    Set cInit.ActiveInterface = ActiveInterface
    cInit.ConnectionString = ActiveInterface.Connection

    Set cInit.RSColumns = Nothing
    cInit.CreateColumnsFromRS = False
    
    cInit.KeyField = "GB07_ID"
    cInit.ShowFooter = False
    
    With QGridDocumenti
      Set .InitializationClass = cInit
      
'      .Title = "Documenti"

      .DataFormatEnabled = True
      .INIT_ADDColumnSELECTIONBOXEXT "GB07_CHECK", "Sel", gedTextEdit, 700, True
      .INIT_ADDColumnIMAGE "FLGIMG", "Img", ImageList, "", ONLYWITHIMAGE, 500
      .INIT_ADDColumn "GB07_ID", "Id", gedTextEdit, 700, False
      .INIT_ADDColumn "GB07_RAG", "Raggr.", gedTextEdit, 2000
      .INIT_ADDColumn "GB07_SEQ", "Seq.", gedTextEdit, 2000
      .INIT_ADDColumn "GB07_ALT", "Altern", gedTextEdit, 2000, False
      .INIT_ADDColumn "GB07_CODART_MG66", "Articolo", gedTextEdit, 2000
      .INIT_ADDColumn "GB07_DESCART", "Descrizione", gedTextEdit, 3000
      .INIT_ADDColumn "GB07_QTA", "Quantita", gedCurrencyEdit, 1000
      .INIT_ADDColumn "MG66_UM1", "Um.", gedTextEdit, 1000
      .INIT_ADDColumn "GB07_PREZZO", "Prezzo", gedCurrencyEdit, 1000, True, , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_SC1", "Sc1", gedCurrencyEdit, 1000, True, , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_SC2", "Sc2", gedCurrencyEdit, 1000, True, , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_SC3", "Sc3", gedCurrencyEdit, 1000, True, , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_SC4", "Sc4", gedCurrencyEdit, 1000, True, , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_SC5", "Sc5", gedCurrencyEdit, 1000, True, , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_SC6", "Sc6", gedCurrencyEdit, 1000, True, , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_IMPORTO", "Importo", gedCurrencyEdit, 1000, True, , "###,###,###,##0.0"
      .INIT_ADDColumn "GB07_IMPORTOUNI", "Importo Unitario", gedCurrencyEdit, 1000, True, , "€###,###,###,##0.0"
      .INIT_ADDColumn "GB07_PERCPROVV", "Perc. Provv.", gedCurrencyEdit, 1000, True, , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_IMPPROVV", "Importo Provv.", gedCurrencyEdit, 1000, True, , "###,###,###,##0.0"
      .INIT_ADDColumn "GB07_COSTO", "Costo", gedCurrencyEdit, 1000, True, , "###,###,###,##0.00"
      
      .INIT_ADDColumn "MG66_FAM_MG53", "Fam.", gedTextEdit, 2000
      .INIT_ADDColumn "MG66_SFAM_MG54", "S.Fam.", gedTextEdit, 2000
      .INIT_ADDColumn "MG66_GRUPPO_MG55", "Gruppo.", gedTextEdit, 2000
      .INIT_ADDColumn "MG66_SGRUPPO_MG56", "S.Gruppo", gedTextEdit, 2000
      .INIT_ADDColumn "MG66_GRUSTAT1_MG74", "Gr.St1", gedTextEdit, 2000
      .INIT_ADDColumn "MG66_GRUSTAT2_MG75", "Gr.St2.", gedTextEdit, 2000
      .INIT_ADDColumn "MG66_GRUSTAT3_MG76", "Gr.St3.", gedTextEdit, 2000
      .INIT_ADDColumn "MG66_GRUSTAT4_MG77", "Gr.St4.", gedTextEdit, 2000
      
      .INIT_ADDColumn "HCL", "HCL", gedTextEdit, 2000
      .INIT_ADDColumn "Area", "Area", gedTextEdit, 2000
      .INIT_ADDColumnIMAGE "IMGCODFOR", "Forn.", ImgLstFornitore, "", ONLYWITHIMAGE, 500
      .INIT_ADDColumn "GB07_CLIFOR_CG44", "Cod. Fornitore", gedTextEdit, 2000

      

      

      .InitializeSTART
      .InitializeEND

    End With
    
    
'    Call CaricaGriglia("QGridDocumenti", "Fatture")
'    Set QGridDocumenti.DataSource = Nothing
'
'    Call CaricaGriglia("QGridScontrino", "Scontrini")
'    Set QGridScontrino.DataSource = Nothing
'
'    Call CaricaGriglia("QGridNC", "Note credito")
'    Set QGridNC.DataSource = Nothing
'
'    Call CaricaGriglia("QGridReso", "Reso")
'    Set QGridReso.DataSource = Nothing

Exit Sub

Errore:
    Select Case VisualizzaErrore("InitQGrid")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbCancel
            Resume Next
    End Select

End Sub
Private Sub INIT_SIMULAZIONE()
On Error GoTo Errore


'****************************************************************************************************
'Inizializzazione qgrid anagrafica
    Dim cInit       As New TMS_QGRID.CLSFW_INITQGRID

    Set cInit.ActiveInterface = ActiveInterface
    cInit.ConnectionString = ActiveInterface.Connection

    Set cInit.RSColumns = Nothing
    cInit.CreateColumnsFromRS = False
    
    cInit.KeyField = "GB07_ID"
    cInit.ShowFooter = False
    
    With QGRID_SIMULAZIONE
      Set .InitializationClass = cInit
      .DataFormatEnabled = True
      .CustomDrawCellEnabled = True
'      .Title = "Documenti"

      
      
      .INIT_ADDColumnIMAGE "FLGIMG", "Img", ImageList, "", ONLYWITHIMAGE, 500
      .INIT_ADDColumn "GB07_ID", "Id", gedTextEdit, 700, False
      .INIT_ADDColumn "GB07_RAG", "Raggr.", gedTextEdit, 2000
      .INIT_ADDColumn "GB07_SEQ", "Seq.", gedTextEdit, 2000
      .INIT_ADDColumn "GB07_ALT", "Altern", gedTextEdit, 2000, False
      .INIT_ADDColumn "GB07_CODART_MG66", "Articolo", gedTextEdit, 2000
      .INIT_ADDColumn "GB07_DESCART", "Descrizione", gedTextEdit, 3000
      .INIT_ADDColumn "GB07_QTA", "Quantita", gedCurrencyEdit, 1000, , , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_PREZZO", "Prezzo", gedCurrencyEdit, 1000, , , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_PREZZO_NEW", "Prezzo Sim.", gedCurrencyEdit, 1000, , , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_SC1", "Sc1", gedCurrencyEdit, 1000, , , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_SC1_NEW", "Sc1 Sim.", gedCurrencyEdit, 1000, , , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_SC2", "Sc2", gedCurrencyEdit, 1000, , , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_SC2_NEW", "Sc2 Sim.", gedCurrencyEdit, 1000, , , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_SC3", "Sc3", gedCurrencyEdit, 1000, , , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_SC3_NEW", "Sc3 Sim.", gedCurrencyEdit, 1000, , , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_SC4", "Sc4", gedCurrencyEdit, 1000, , , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_SC4_NEW", "Sc4 Sim.", gedCurrencyEdit, 1000, , , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_SC5", "Sc5", gedCurrencyEdit, 1000, , , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_SC5_NEW", "Sc5 Sim.", gedCurrencyEdit, 1000, , , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_SC6", "Sc6", gedCurrencyEdit, 1000, , , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_SC6_NEW", "Sc6 Sim.", gedCurrencyEdit, 1000, , , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_IMPORTO", "Importo", gedCurrencyEdit, 1000, , , "###,###,###,##0.00"
      .INIT_ADDColumn "GB07_IMPORTO_NEW", "Importo Sim.", gedCurrencyEdit, 1000, , , "###,###,###,##0.00"
      .INIT_ADDColumn "MG66_FAM_MG53", "Fam.", gedTextEdit, 2000
      .INIT_ADDColumn "MG66_SFAM_MG54", "S.Fam.", gedTextEdit, 2000
      .INIT_ADDColumn "MG66_GRUPPO_MG55", "Gruppo.", gedTextEdit, 2000
      .INIT_ADDColumn "MG66_SGRUPPO_MG56", "S.Gruppo", gedTextEdit, 2000
      .INIT_ADDColumn "MG66_GRUSTAT1_MG74", "Gr.St1", gedTextEdit, 2000
      .INIT_ADDColumn "MG66_GRUSTAT2_MG75", "Gr.St2.", gedTextEdit, 2000
      .INIT_ADDColumn "MG66_GRUSTAT3_MG76", "Gr.St3.", gedTextEdit, 2000
      .INIT_ADDColumn "MG66_GRUSTAT4_MG77", "Gr.St4.", gedTextEdit, 2000
      
     
      

      

      .InitializeSTART
      .InitializeEND

    End With
    
    
Exit Sub

Errore:
    Select Case VisualizzaErrore("InitQGrid")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbCancel
            Resume Next
    End Select

End Sub





Private Sub TXT_GB06_CLIFOR_CG44_AfterItem(Cancel As Boolean)
  
  On Error GoTo ErrTrap
  
  If Not TXT_GB06_CLIFOR_CG44.IsValid Then
    Exit Sub
  End If
  
  Dim Ret As Boolean
  'Recupero fido del cliente se documento diverso da scontrino
  Ret = CercoFidoCliente(NVL(TXT_GB06_CLIFOR_CG44.Text, 0))
  If Not Ret Then
    'MsgBox "Cliente bloccato per fido"
'    If TXT_GB06_CLIFOR_CG44.Enabled = True Then
'      TXT_GB06_CLIFOR_CG44.SetFocus
'    End If
    
'    FrameGriglia.Enable = False
'    cmdGeneraFat.Enabled = False
    If NVL(TXT_GB06_CODPAG_CG62.Text, "") = CODPAGCONTSPEC Then
      
    End If
  Else
    FrameGriglia.Enable = True
    
    
  End If
  
  'Controllo cliente bloccato
  Dim Pstr_Sql  As String
  Dim MyRst     As ADODB.Recordset
  
  'INIZIO CONTROLLO SE ABILITATO
  Pstr_Sql = " SELECT MG19_INDCLIBLOC " & _
             " FROM MG19_CLIFORVA " & _
             " WHERE (MG19_DITTA_CG18 = " & CodiceDitta & ") AND (MG19_TIPOCF_CG44 = 0) " & _
             " AND (MG19_CLIFOR_CG44  = " & NVL(TXT_GB06_CLIFOR_CG44.Text, 0) & ") "
  
  ClienteBloccato = False
  Set MyRst = Gcon_Connect.Execute(Pstr_Sql, , adCmdText)
  If Not MyRst.EOF Then
      If Trim$(MyRst("MG19_INDCLIBLOC").value) >= 1 Then
        ClienteBloccato = True
        MsgBox "Attenzione cliente bloccato"
        TXT_GB06_CLIFOR_CG44.SetFocus
      End If
  End If
  MyRst.Close
  Set MyRst = Nothing
  Call CMD_SAVE_Click
  
  
  On Error GoTo ErrTrap
  
  Exit Sub
ErrTrap:
  Select Case VisualizzaErrore("TXT_GB06_CLIFOR_CG44_AfterItem")
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
  End Select

End Sub



Private Sub TXT_GB06_CLIFOR_CG44_CloseDecode(Arr_Fields As Variant)
  
  On Error GoTo ErrTrap
  
  If Not TXT_GB06_CLIFOR_CG44.IsValid Then
    Exit Sub
  End If
 ' TXT_GB06_CODDESTIN_MG22.Text = TXT_DECDEST.Text
  
  If IsNull(Trim(TXT_CODPAG.Text)) Then
  
    MsgBox "Attenzione Condizione di pagamento non caricata in anagrafica", vbCritical
  
  Else
  
    CODPAGCLI = Trim(TXT_CODPAG.Text)
  
  End If
    
  
  If NVL(TXT_GB06_CODPAG_CG62.Text, "") <> "" Then
    If TXT_GB06_CODPAG_CG62.Text <> CODPAGCLI Then
'      If MsgBox("Condizione di pagamento impostata diversa da quella del cliente. Vuoi aggiornarla?", vbYesNo) = vbYes Then
'        TXT_GB06_CODPAG_CG62.Text = CODPAGCLI
'      End If
    End If
  Else
    TXT_GB06_CODPAG_CG62.Text = CODPAGCLI
  End If
  
  Select Case TXT_GB06_CODPAG_CG62.Text
  Case CODPAGCONTANTI, CODPAGBANCOMAT, CODPAGCARTACR, CODPAGASSEGNI ', CODPAGCONTSPEC
'    FramePag.Visible = True
'    FramePag.Enable = True
  Case Else
'    If CHK_FORZAFATT.Text = 1 Then
'        FramePag.Visible = True
'        FramePag.Enable = True
'    End If
  End Select
  
  Select Case TXT_GB06_CODPAG_CG62.Text
  Case CODPAGCONTANTI, CODPAGBANCOMAT, CODPAGCARTACR, CODPAGASSEGNI
  Case Else
    If TipoDocumento = 5 Then
    End If
  End Select
  
Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("TXT_GB06_CLIFOR_CG44_CloseDecode")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select

  
End Sub

Private Sub TXT_GB06_CLIFOR_CG44_KeyButtonMenuPress(Cancel As Boolean, ByVal Pstr_KeyButtonPress As Variant)

  On Error Resume Next
  
  Select Case Pstr_KeyButtonPress
    Case "Kgestione"
      Set Cls_ConnectCommon.ActiveInterface = ActiveInterface
      Cls_ConnectCommon.CodiceDitta = CodiceDitta
      Cls_ConnectCommon.CodiceClifor = RTrimN(TXT_GB06_CLIFOR_CG44.Text)
      Select Case CDecN(ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsParMagaz.IndCallCliFor)
        Case 0 'anagrafica clienti/fornitori
          Call Cls_ConnectCommon.CallAnagClienti
        Case 1 'wizard creazione veloce
          Call Cls_ConnectCommon.WizardClienti
      End Select
      ActiveInterface.IsActive = True
      Set Cls_ConnectCommon.ActiveInterface = Nothing
      Cls_ConnectCommon.TerminateConnect
      Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
'      Call InitializeScript

  End Select
  
End Sub


Private Sub TXT_GB06_CLIFOR_CG44_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
  
  Dim conta As Integer
  Dim Pvar_ArrCampi()
  On Error Resume Next
        
  If NVL(TXT_GB06_CLIFOR_CG44.Text, "") = 0 Then
      Exit Sub
  End If
  
  str_SQL = " SELECT" & _
            "    CG16_RAGSOANAG," & _
            "    CG16_INDIRIZZO," & _
            "    CG16_CAP," & _
            "    CG16_CITTA," & _
            "    CG16_PROV, " & _
            "    CG44_CODPAG_CG62, " & _
            "    CG44_AGENTE_MG17, " & _
            "    MG19_LISTMAG  "
  str_SQL = str_SQL & _
            " FROM" & _
            "    CG44_CLIFOR WITH (NOLOCK)" & _
            " INNER JOIN CG16_ANAGGEN WITH (NOLOCK) ON" & _
            "    CG16_CODICE = CG44_CODICE_CG16 " & _
            " INNER JOIN MG19_CLIFORVA WITH (NOLOCK) ON" & _
            "    MG19_DITTA_CG18  = CG44_DITTA_CG18  AND " & _
            "    MG19_TIPOCF_CG44 = CG44_TIPOCF AND " & _
            "    MG19_CLIFOR_CG44 = CG44_CLIFOR " & _
            " WHERE" & _
            "    CG44_DITTA_CG18 = " & CodiceDitta & " AND" & _
            "    CG44_TIPOCF = 0  AND" & _
            "    CG44_CLIFOR = " & CDecN(TXT_GB06_CLIFOR_CG44.Text) & " AND" & _
            "    CG16_RAGSOANAG NOT LIKE '*%'"
        
        
        
  conta = 8

  ReDim Preserve Pvar_ArrCampi(0 To conta - 1, 0 To 1)
  
  conta = 0
  
  Set Pvar_ArrCampi(conta, 0) = TXT_CG16_RAGSOANAG
  Pvar_ArrCampi(conta, 1) = "CG16_RAGSOANAG"
  conta = conta + 1
  
  Set Pvar_ArrCampi(conta, 0) = TXT_CG16_INDIRIZZO
  Pvar_ArrCampi(conta, 1) = "CG16_INDIRIZZO"
  conta = conta + 1
      
  Set Pvar_ArrCampi(conta, 0) = TXT_CG16_CAP
  Pvar_ArrCampi(conta, 1) = "CG16_CAP"
  conta = conta + 1
      
  Set Pvar_ArrCampi(conta, 0) = TXT_CG16_CITTA
  Pvar_ArrCampi(conta, 1) = "CG16_CITTA"
  conta = conta + 1
      
  Set Pvar_ArrCampi(conta, 0) = TXT_CG16_PROV
  Pvar_ArrCampi(conta, 1) = "CG16_PROV"
  conta = conta + 1
  
  Set Pvar_ArrCampi(conta, 0) = TXT_CODPAG
  Pvar_ArrCampi(conta, 1) = "CG44_CODPAG_CG62"
  conta = conta + 1
  
  Set Pvar_ArrCampi(conta, 0) = TXT_GB06_AGENTE
  Pvar_ArrCampi(conta, 1) = "CG44_AGENTE_MG17"
  conta = conta + 1
  
  Set Pvar_ArrCampi(conta, 0) = TXT_MG19_LISTMAG
  Pvar_ArrCampi(conta, 1) = "MG19_LISTMAG"
  conta = conta + 1

  
'  Set Pvar_ArrCampi(conta, 0) = TXT_DECDEST
'  Pvar_ArrCampi(conta, 1) = "MG19_CODDESTPREV"
'  conta = conta + 1
  
  
  'MG19_MAGPERCOR
  
  Arr_Fields = Pvar_ArrCampi
  Str_Connect = Gstr_Connect
  Exit Sub
End Sub

Private Sub TXT_GB06_CLIFOR_CG44_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    On Error Resume Next
    
    Call Cls_LookupCommon.ClientiFornitori(0)
    str_SQL = Cls_LookupCommon.StringaSQL
                
    str_SQL = Replace(str_SQL, "cg44_tipocf = 0", " cg44_tipocf = 0 AND  CG16_RAGSOANAG NOT LIKE '*%'")
                
    'str_SQL = " SELECT" & _
              "    CG16_RAGSOANAG," & _
              "    CG16_INDIRIZZO," & _
              "    CG16_CAP," & _
              "    CG16_CITTA," & _
              "    CG16_PROV, " & _
              "    CG44_CODPAG_CG62, " & _
              "    MG19_LISTMAG "
    'str_SQL = str_SQL & _
              " FROM" & _
              "    CG44_CLIFOR WITH (NOLOCK)" & _
              " INNER JOIN CG16_ANAGGEN WITH (NOLOCK) ON" & _
              "    CG16_CODICE = CG44_CODICE_CG16 " & _
              " INNER JOIN MG19_CLIFORVA WITH (NOLOCK) ON" & _
              "    MG19_DITTA_CG18  = CG44_DITTA_CG18  AND " & _
              "    MG19_TIPOCF_CG44 = CG44_TIPOCF AND " & _
              "    MG19_CLIFOR_CG44 = CG44_CLIFOR " & _
              " WHERE" & _
              "    CG44_DITTA_CG18 = " & CodiceDitta & " AND" & _
              "    CG44_TIPOCF = 0  AND" & _
              "    CG44_CLIFOR = " & CDecN(TXT_GB06_CLIFOR_CG44.Text) & " AND" & _
              "    CG16_RAGSOANAG NOT LIKE '*%'"
            

    Arr_Fields = Cls_LookupCommon.ArrayFields
    Str_Caption = Cls_LookupCommon.Titolo
    Str_Connect = Gstr_Connect
    TXT_GB06_CLIFOR_CG44.IDLookup = Cls_LookupCommon.IDLookup
    
End Sub



'Private Sub TXT_GB06_CODDESTIN_MG22_KeyButtonMenuPress(Cancel As Boolean, ByVal Pstr_KeyButtonPress As Variant)
'    On Error Resume Next
'
'    If Pstr_KeyButtonPress = "Kgestione" Then
'        If NVL(TXT_GB06_CLIFOR_CG44.Text, "") = 0 Then
'            MsgBox "Codice cliente non impostato !!!"
'        Else
'            Set Cls_ConnectCommon.ActiveInterface = ActiveInterface
'            Cls_ConnectCommon.CodiceDestMerceTipoCF = 0
'            Cls_ConnectCommon.CodiceDestMerceClienteFornitore = CDecN(TXT_GB06_CLIFOR_CG44.Text)
'            Cls_ConnectCommon.CodiceDestMerceDestinatario = RTrimN(TXT_GB06_CODDESTIN_MG22.Text)
'
'            Call Cls_ConnectCommon.CallDestinatariMerceClienti
'
'            ActiveInterface.IsActive = True
'            Set Cls_ConnectCommon.ActiveInterface = Nothing
'            Cls_ConnectCommon.TerminateConnect
'            Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
''            Call InitializeScript
'        End If
'    End If
'End Sub

'Private Sub TXT_GB06_CODDESTIN_MG22_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
'On Error Resume Next
'    Dim Pvar_ArrCampi() As Variant
'    Dim conta As Integer
'    str_SQL = "SELECT * FROM MG22_CLIFORDEST WITH (NOLOCK) " & _
'              "LEFT OUTER JOIN CG07_TABSTATIEST WITH (NOLOCK) ON " & _
'              "CG07_CODICE = MG22_STATOEST_CG07 " & _
'              "WHERE MG22_DITTA_CG18 = " & CodiceDitta & " AND MG22_TIPOCF_CG44 = 0 AND MG22_CLIFOR_CG44 = " & CDecN(TXT_GB06_CLIFOR_CG44.Text) & " AND " & _
'              "MG22_CODDESTIN = '" & RTrimN(TXT_GB06_CODDESTIN_MG22.Text) & "'"
'
'    conta = 5
'
'    ReDim Preserve Pvar_ArrCampi(0 To conta - 1, 0 To 1)
'    conta = 0
'    Set Pvar_ArrCampi(conta, 0) = TXT_MG22_DESTRAGSOC
'    Pvar_ArrCampi(conta, 1) = "MG22_DESTRAGSOC"
'    conta = conta + 1
'
'    Set Pvar_ArrCampi(conta, 0) = TXT_MG22_DESTIND
'    Pvar_ArrCampi(conta, 1) = "MG22_DESTIND"
'    conta = conta + 1
'
'    Set Pvar_ArrCampi(conta, 0) = TXT_MG22_DESTCAP
'    Pvar_ArrCampi(conta, 1) = "MG22_DESTCAPCHAR"
'    conta = conta + 1
'
'    Set Pvar_ArrCampi(conta, 0) = TXT_MG22_DESTPROV
'    Pvar_ArrCampi(conta, 1) = "MG22_DESTPROV"
'    conta = conta + 1
'
'    Set Pvar_ArrCampi(conta, 0) = TXT_MG22_DESTCITTA
'    Pvar_ArrCampi(conta, 1) = "MG22_DESTCITTA"
'    conta = conta + 1
'
'    Arr_Fields = Pvar_ArrCampi
'    Str_Connect = Gstr_Connect
'End Sub

'Private Sub TXT_GB06_CODDESTIN_MG22_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
'    On Error Resume Next
'    Call Cls_LookupMagazzino.Destinatari(0, CDecN(TXT_GB06_CLIFOR_CG44.Text))
'    str_SQL = Cls_LookupMagazzino.StringaSQL
'    Arr_Fields = Cls_LookupMagazzino.ArrayFields
'    Str_Caption = Cls_LookupMagazzino.Titolo
'    Str_Connect = Gstr_Connect
'    TXT_GB06_CODDESTIN_MG22.IDLookup = Cls_LookupMagazzino.IDLookup
'End Sub

Private Sub TXT_GB06_CODPAG_CG62_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
    On Error Resume Next
    Call Cls_DecodeCommon.CondizioneDiPagamento(RTrimN(TXT_GB06_CODPAG_CG62.Text))
    str_SQL = Cls_DecodeCommon.StringaSQL
    Arr_Fields = Cls_DecodeCommon.ArrayFields
    
    Set Arr_Fields(0, 0) = TXT_CG62_DESCPAG
    
    Str_Connect = Gstr_Connect
End Sub


Private Sub TXT_GB06_CODPAG_CG62_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    On Error Resume Next
    Call Cls_LookupCommon.CondizioniDiPagamento
    str_SQL = Cls_LookupCommon.StringaSQL
    Arr_Fields = Cls_LookupCommon.ArrayFields
    Str_Caption = Cls_LookupCommon.Titolo
    Str_Connect = Gstr_Connect
    TXT_GB06_CODPAG_CG62.IDLookup = Cls_LookupCommon.IDLookup

End Sub

'Private Sub TXT_GB06_VETTORE_MG14_KeyButtonMenuPress(Cancel As Boolean, ByVal Pstr_KeyButtonPress As Variant)
'    On Error Resume Next
'
'    If Pstr_KeyButtonPress = "Kgestione" Then
'        Cls_ConnectMagazzino.Left = 100
'        Cls_ConnectMagazzino.Top = 1000
'        Set Cls_ConnectMagazzino.ActiveInterface = ActiveInterface
'        Set Cls_ConnectMagazzino.ConnectField = TXT_GB06_VETTORE_MG14
'        Call Cls_ConnectMagazzino.Vettori(RTrimN(TXT_GB06_VETTORE_MG14.Text))
'        ActiveInterface.IsActive = True
'        Set Cls_ConnectMagazzino.ActiveInterface = Nothing
'        Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
''        Call InitializeScript
'    End If
'End Sub

'Private Sub TXT_GB06_VETTORE_MG14_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
'
'
'    Dim Pvar_ArrCampi() As Variant
'    Dim conta As Integer
'
'    On Error Resume Next
'
'    str_SQL = " SELECT" & _
'              "    *" & _
'              " FROM" & _
'              "    MG14_VETTORI WITH (NOLOCK)" & _
'              " INNER JOIN CG16_ANAGGEN WITH (NOLOCK) ON" & _
'              "    MG14_CODICE_CG16 = CG16_CODICE" & _
'              " LEFT JOIN MG41_CLIFORVETT WITH (NOLOCK) ON" & _
'              "    MG41_DITTA_CG18 = " & CodiceDitta & " AND" & _
'              "    MG41_TIPOCF_CG44 = 0 AND" & _
'              "    MG41_CLIFOR_CG44 = " & CDecN(TXT_GB06_CLIFOR_CG44.Text) & " AND" & _
'              "    MG41_CODICE_MG14 = MG14_CODICE" & _
'              " WHERE" & _
'              "    MG14_CODICE = '" & RTrimN(TXT_GB06_VETTORE_MG14.Text) & "'"
'
'    conta = 5
'
'    ReDim Preserve Pvar_ArrCampi(0 To conta - 1, 0 To 1)
'    conta = 0
'    Set Pvar_ArrCampi(conta, 0) = TXT_CG16_RAGSOANAG_VETT
'    Pvar_ArrCampi(conta, 1) = "CG16_RAGSOANAG"
'    conta = conta + 1
'
'    Set Pvar_ArrCampi(conta, 0) = TXT_CG16_INDIRIZZO_VETT
'    Pvar_ArrCampi(conta, 1) = "CG16_INDIRIZZO"
'    conta = conta + 1
'
'    Set Pvar_ArrCampi(conta, 0) = TXT_CG16_CAP_VETT
'    Pvar_ArrCampi(conta, 1) = "CG16_CAP"
'    conta = conta + 1
'
'    Set Pvar_ArrCampi(conta, 0) = TXT_CG16_CITTA_VETT
'    Pvar_ArrCampi(conta, 1) = "CG16_CITTA"
'    conta = conta + 1
'
'    Set Pvar_ArrCampi(conta, 0) = TXT_CG16_PROV_VETT
'    Pvar_ArrCampi(conta, 1) = "CG16_PROV"
'    conta = conta + 1
'
'
'    Arr_Fields = Pvar_ArrCampi
'    Str_Connect = Gstr_Connect
'End Sub
'
'Private Sub TXT_GB06_VETTORE_MG14_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
'    On Error Resume Next
'
'    Call Cls_LookupMagazzino.Vettori
'    str_SQL = Cls_LookupMagazzino.StringaSQL
'    Arr_Fields = Cls_LookupMagazzino.ArrayFields
'    Str_Caption = Cls_LookupMagazzino.Titolo
'    Str_Connect = Gstr_Connect
'    TXT_GB06_VETTORE_MG14.IDLookup = Cls_LookupMagazzino.IDLookup
'End Sub

Private Sub TXT_GB07_CODART_MG66_KeyButtonMenuPress(Cancel As Boolean, ByVal Pstr_KeyButtonPress As Variant)
    On Error Resume Next
    
    Select Case Pstr_KeyButtonPress
        Case "Kgestione"
            Set Cls_ConnectMagazzino.ActiveInterface = ActiveInterface
            Cls_ConnectMagazzino.Left = 5
            Cls_ConnectMagazzino.Top = 1000
            Set Cls_ConnectMagazzino.ConnectField = Nothing
            Call Cls_ConnectMagazzino.ArticoloAnagrafica(RTrimN(TXT_GB07_CODART_MG66.Text))
            ActiveInterface.IsActive = True
            Set Cls_ConnectMagazzino.ActiveInterface = Nothing
            Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
    End Select
End Sub

Private Sub TXT_GB07_CODART_MG66_AfterChange(Cancel As Boolean)
Dim old_CodArt As String
Dim strSQL As String
  On Error GoTo ErrTrap
  
  If NVL(TXT_GB07_CODART_MG66.Text, "") = "" Or Not (TXT_GB07_CODART_MG66.IsValid) Then
    Exit Sub
  End If

  
'Imposto qta 1

  old_CodArt = TXT_GB07_CODART_MG66.Text
  TXT_GB07_QTA.Text = 1
  TXT_GB07_ALT.Text = 1
  TXT_GB07_RAG.Text = 1
  
  
  If ArtSoggettoaScontoAuto(old_CodArt) Then
    TXT_GB07_SC1.Text = NVL(TXT_SC1_FISSO.Text, 0)
    TXT_GB07_SC2.Text = NVL(TXT_SC2_FISSO.Text, 0)
  End If
  strSQL = "select isnull(max(gb07_seq),0) +10 from GB07_CORPODOC where GB07_ID_GB06 = " & IDGB06
  TXT_GB07_SEQ.Text = GetValFromQuery(strSQL, 0, Gcon_Connect)

'Ricalcolo prezzo
  Call RicalcolaPrezzo

'ricacolo importi
  Call RicalcolaImporto
  
 Call RecuperImmagineHyperMedia(TXT_GB07_CODART_MG66.Text)
 If InStr(1, TXT_GB07_CODART_MG66.Text, "P_", vbTextCompare) = 0 Then
     TXT_GB07_CLIFOR_CG44.Text = RecuperaFornitorePref(TXT_GB07_CODART_MG66.Text)
 End If
 
' If InStr(1, TXT_GB07_CODART_MG66.Text, "MA", vbTextCompare) = 1 Then
'     TXT_GB07_CLIFOR_CG44.Text = NVL(TXT_CLIFOR_CG44.Text, "")
' End If
'
' If InStr(1, TXT_GB07_CODART_MG66.Text, "LAVO", vbTextCompare) = 0 Then
'     TXT_GB07_CLIFOR_CG44.Text = NVL(TXT_CLIFOR_CG44.Text, "")
' End If
 
 Call RecuperaCosto
 
' If NVL(RecuperaFornitorePref(TXT_GB07_CODART_MG66.Text), "") = "" Then
'        TXT_GB07_CLIFOR_CG44.Text = NVL(TXT_CLIFOR_CG44.Text, "")
'    Else
'        TXT_GB07_CLIFOR_CG44.Text = NVL(RecuperaFornitorePref(TXT_GB07_CODART_MG66.Text), NVL(TXT_CLIFOR_CG44.Text, ""))
' End If
 
 If TXT_GB07_FLPOSA.Text = 1 And ArtSoggettoaPosaAuto(old_CodArt) Then
' isBarcode = False
    FME_BANCO.Status = tsInsert
    FME_BANCO.AddNew False, True, False
    TXT_GB07_CODART_MG66.Text = "P_" & old_CodArt
   
    Call TXT_GB07_CODART_MG66_Validate(False)
    If NVL(RecuperaFornitorePref(TXT_GB07_CODART_MG66.Text), "") = "" Then
    TXT_GB07_CLIFOR_CG44.Text = NVL(TXT_CLIFOR_CG44.Text, "")
    Else
    TXT_GB07_CLIFOR_CG44.Text = NVL(RecuperaFornitorePref(TXT_GB07_CODART_MG66.Text), NVL(TXT_CLIFOR_CG44.Text, ""))
    End If
    Call RecuperaCosto

 End If
 
 TXT_GB07_QTA.SetFocus

  
Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("TXT_GB07_CODART_MG66_AfterChange")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
  
End Sub

Public Function RecuperaProvvigione(Prezzo As Double, Importo As Double, Famiglia As String, SFamiglia As String, Agente As String, Sc1 As Double, Sc2 As Double, Sc3 As Double, Sc4 As Double, Sc5 As Double, Sc6 As Double) As Double
    Dim strConnection As String
    Dim oRs As ADODB.Recordset
    Dim PercSconto As Double
    Set oRs = New ADODB.Recordset
    Dim cn As ADODB.Connection
    strConnection = Gcon_Connect.ConnectionString
    Set cn = New ADODB.Connection
    cn.Open strConnection
    Dim result As String
    Dim CodiceTabellaProvv As String
    Dim strSQL As String
'50 + 10% = 55
'
'55 + 3% = 56,65

    Select Case Famiglia
    Case "SERV"
        If SFamiglia = "POSA" Then
            SFamiglia = "POS"
        End If
        CodiceTabellaProvv = Famiglia + "-" + SFamiglia
    Case Else
        CodiceTabellaProvv = Famiglia
    End Select
    If NVL(Prezzo, 0) > 0 Then
    PercSconto = ((CDbl(NVL(Round(Prezzo * TXT_GB07_QTA.Text, 2), 0)) - CDbl(Importo)) / CDbl(Round(Prezzo * TXT_GB07_QTA.Text, 2))) * 100
    Else
    PercSconto = 0
    End If
   
    strSQL = " SELECT   TOP (1) PA12_PROVPER "
    strSQL = strSQL & " From PA12_PROVSCSCAGL "
    strSQL = strSQL & " WHERE        (PA12_DITTA_CG18 = " & CodiceDitta & ") "
    strSQL = strSQL & " AND (PA12_FLGVENACQ = 0) AND (PA12_INDTIPOAGE = 1)"
    strSQL = strSQL & " AND (PA12_VALUTA_CG08 = 'EURO') "
    strSQL = strSQL & " AND (PA12_AGENTE_MG17 = '" & Agente & "') "
    strSQL = strSQL & " AND (PA12_CODPROVSC = '" & CodiceTabellaProvv & "') "
    strSQL = strSQL & " AND (PA12_INDTIPOAGE = 1) "
    strSQL = strSQL & " AND (" & Replace(PercSconto, ",", ".") & " <= PA12_APERCSCONTO)"
    strSQL = strSQL & " ORDER BY PA12_APERCSCONTO"
    Set oRs = cn.Execute(strSQL)
    
    If oRs.EOF Then
        result = 0
        TXT_GB07_IMPPROVV.Text = 0
    Else
        result = NVL(CStr(oRs("PA12_PROVPER")), 0)
        TXT_GB07_IMPPROVV.Text = (CDbl(TXT_GB07_IMPORTO.Text) / 100) * CDbl(result)
        
    End If
    
    oRs.Close
    Set oRs = Nothing
    cn.Close
    Set cn = Nothing
    
    RecuperaProvvigione = result
End Function



Public Function RecuperaFornitorePref(articolo As String) As String
    Dim strConnection As String
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    Dim cn As ADODB.Connection
    strConnection = Gcon_Connect.ConnectionString
    Set cn = New ADODB.Connection
    cn.Open strConnection
    Dim result As String
   
    Dim strSQL As String
   
    strSQL = " Select MG73_CLIFOR_CG44 FROM MG73_ARTCLIFOR " & _
             " WHERE        (MG73_DITTA_CG18 = " & CodiceDitta & ") AND (MG73_TIPOCF_CG44 = 1) AND (MG73_CODART_MG66 = '" & articolo & "') "
   
    Set oRs = cn.Execute(strSQL)
    
    If oRs.EOF Then
        If InStr(1, articolo, "TRASPORTO") = 1 Then
            result = ""
        Else
        
            result = NVL(TXT_CLIFOR_CG44.Text, "")
        End If
        
    Else
        result = NVL(CStr(oRs("MG73_CLIFOR_CG44")), "")
    End If
    
    oRs.Close
    Set oRs = Nothing
    cn.Close
    Set cn = Nothing

    RecuperaFornitorePref = result
End Function


Public Sub RecuperImmagineHyperMedia(articolo As String)
 Dim strConnection As String
 
 
 Dim OutFile() As Byte
Dim strSQL As String
Dim FileImgName As String
FileImgName = "Temp"
FileImgName = FileImgName & "_" & TXT_GB07_CODART_MG66.Text

Dim rsSearchResults As ADODB.Recordset

    strSQL = " SELECT  CG99_IMAGE " & _
             " FROM CG99_MULTIMEDIA INNER JOIN " & _
             " MG66_ANAGRART ON CG99_MULTIMEDIA.CG99_IDMEDIA = MG66_ANAGRART.MG66_IDMEDIA_CG99" & _
             " WHERE MG66_CODART = '" & articolo & "' AND MG66_DITTA_CG18 = " & CodiceDitta

    'Open a recordset to hold the search results.
    Set rsSearchResults = New ADODB.Recordset
    rsSearchResults.Open strSQL, Gcon_Connect, adOpenStatic, _
        adLockPessimistic

    If rsSearchResults.EOF Then
       ' MsgBox "Immagine non Caricata"
        'Don't change rs, since no match was found we'll
        ' stay on whatever
        'record was previously selected.
    Else
       strSQL = " update  GB07_CORPODOC set gb07_IMG = (SELECT       CG99_IMAGE " & _
        " FROM            CG99_MULTIMEDIA INNER JOIN " & _
        "                         MG66_ANAGRART ON CG99_MULTIMEDIA.CG99_IDMEDIA = MG66_ANAGRART.MG66_IDMEDIA_CG99 " & _
        " where MG66_CODART = '" & articolo & "') " & _
        " WHERE        (GB07_ID_GB06 = " & IDGB06 & ") AND (GB07_CODART_MG66 = '" & articolo & "') "
        Gcon_Connect.Execute strSQL
       If Dir(App.Path & "\temp.jpg") <> "" Then Kill _
            App.Path & "\temp.jpg"

        OutFile = rsSearchResults("CG99_IMAGE")
        'Write File
        Open App.Path & "\temp.jpg" For Binary Access Write _
            As #1
        Put #1, , OutFile
        Close #1
        TXT_GB07_IMAGEPATH.Text = App.Path & "\temp.jpg"
        PictureArticoli.Picture = LoadPicture(App.Path & "\temp.jpg")
    End If
 
    Set rsSearchResults = Nothing
'    Dim oRs As ADODB.Recordset
'    Set oRs = New ADODB.Recordset
'    Dim cn As ADODB.Connection
'    strConnection = Gcon_Connect.ConnectionString
'    Set cn = New ADODB.Connection
'    cn.Open strConnection
'    Dim result As String
'    Dim strUpdate As String
'
'    Dim strSQL As String
'
'    strSQL = " SELECT  CG99_IMAGE " & _
'             " FROM CG99_MULTIMEDIA INNER JOIN " & _
'             " MG66_ANAGRART ON CG99_MULTIMEDIA.CG99_IDMEDIA = MG66_ANAGRART.MG66_IDMEDIA_CG99" & _
'             " WHERE MG66_CODART = '" & articolo & "' AND MG66_DITTA_CG18 = " & CodiceDitta
'    Set oRs = cn.Execute(strSQL)
'
'    If oRs.EOF Then
'        result = ""
'    Else
''        strUpdate = "UPDATE GB07_CORPODOC set GB07_IMG = '" & oRs("CG99_IMAGE") & "'"
''        strUpdate = strUpdate & "WHERE GB07_ID_GB06 = " & IDGB06
''        strUpdate = strUpdate & "end GB07_CODART_MG66 = '" & articolo & "'"
''        cn.Execute (strUpdate)
'         rstCorpoBanco("GB07_IMG") = oRs("CG99_IMAGE")
'
'    End If
'
'    oRs.Close
'    Set oRs = Nothing
'    cn.Close
'    Set cn = Nothing

End Sub

Public Function ArtSoggettoaScontoAuto(articolo As String) As Boolean
    Dim strConnection As String
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    Dim cn As ADODB.Connection
    strConnection = Gcon_Connect.ConnectionString
    Set cn = New ADODB.Connection
    cn.Open strConnection
    Dim result As Boolean
   
    Set oRs = cn.Execute("Select * FROM  MG66_ANAGRART WHERE (MG66_DITTA_CG18 = " & CodiceDitta & ") AND (MG66_CODART = '" & articolo & "') ")
    If Trim(oRs("MG66_GRUSTAT3_MG76")) = "GIO" And Not Trim(oRs("MG66_SFAM_MG54")) = "ANTI" Then
       result = True
    Else
        If Trim(oRs("MG66_SFAM_MG54")) = "PDET" Then
            result = True
        Else
            result = False
        End If
    End If
    
    oRs.Close
    Set oRs = Nothing
    cn.Close
    Set cn = Nothing

    ArtSoggettoaScontoAuto = result
End Function


Public Function GetNumeroOfferta(Utente As String) As String
    Dim strConnection As String
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    Dim cn As ADODB.Connection
    strConnection = Gcon_Connect.ConnectionString
    Set cn = New ADODB.Connection
    cn.Open strConnection
    Dim result As Boolean
   
    Set oRs = cn.Execute("select max(cast(substring(gb06_numdoc,3,5) as int)) as NEWNUMDOC from GB06_TESTADOC where gb06_numdoc like '" & Utente & "%' AND (YEAR(GB06_DATA) = YEAR('" & TXT_ANNOINSERIMENTO.Text & "'))")
    If NVL(Trim(oRs("NEWNUMDOC")), 0) = 0 Then
       GetNumeroOfferta = Utente & " " & "0001"
    Else
        GetNumeroOfferta = Utente & " " & Format(oRs("NEWNUMDOC") + 1, String(3, "0"))
    End If
    
    oRs.Close
    Set oRs = Nothing
    cn.Close
    Set cn = Nothing

    
End Function

Public Function GetCommessa(ID As Double) As String

Dim strInsert As String
Dim strCodCommessa As String
strCodCommessa = "CM" & Year(TXT_ANNOINSERIMENTO.Text) & Format(CStr(ID), String(4, "0"))
On Error Resume Next

strInsert = "INSERT INTO PD25_COMMESSA "
strInsert = strInsert & "(PD25_DITTA_CG18,PD25_CODCOMMESSA,PD25_CODSOTCOMM,PD25_DESCR)"
strInsert = strInsert & "VALUES"
strInsert = strInsert & "(" & CodiceDitta & ",'" & strCodCommessa & "',0,'" & TXT_GB06_NOMEOFFERTA.Text & "')"

Gcon_Connect.Execute strInsert

GetCommessa = strCodCommessa
    
End Function

Public Function ArtSoggettoaPosaAuto(articolo As String) As Boolean
    Dim strConnection As String
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    Dim cn As ADODB.Connection
    strConnection = Gcon_Connect.ConnectionString
    Set cn = New ADODB.Connection
    cn.Open strConnection
    Dim result As Boolean
   
    Set oRs = cn.Execute("Select * FROM  MG66_ANAGRART WHERE (MG66_DITTA_CG18 = " & CodiceDitta & ") AND (MG66_CODART = '" & articolo & "') ")
    If Trim(oRs("MG66_GRUSTAT3_MG76")) = "GIO" And Trim(oRs("MG66_FAM_MG53")) <> "SERV" Then
      If GetValFromQuery("Select * FROM  MG66_ANAGRART WHERE (MG66_DITTA_CG18 = " & CodiceDitta & ") AND (MG66_CODART = 'P_" & articolo & "') ", 0, Gcon_Connect) <> "" Then
            result = True
         Else
            result = False
        End If
    Else
       result = False
    End If
    
    oRs.Close
    Set oRs = Nothing
    cn.Close
    Set cn = Nothing

    ArtSoggettoaPosaAuto = result
End Function

Private Sub TXT_GB07_CODART_MG66_AfterItem(Cancel As Boolean)
  On Error GoTo ErrTrap
  If NVL(TXT_GB07_CODART_MG66.Text, "") = "" Then
  
  Exit Sub
  
  End If
'  If Not (rstCorpoBanco.EOF And rstCorpoBanco.BOF) Then
  If rstCorpoBanco.RecordCount > 1 Then
    If NVL(TXT_GB07_CODART_MG66.Text, "") = "" Then
     
    End If
  End If
  
  If Cls_CalcPrezzi.Stato = 0 Then
    If isBarcode Then
      
     FME_BANCO.AddNew False, True, False
     TXT_GB07_CODART_MG66.SetFocus
     Call CMD_REFRESH_Click
    End If
  End If
Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("TXT_GB07_CODART_MG66_AfterItem")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub

Private Sub TXT_GB07_CODART_MG66_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
    On Error Resume Next

    Dim Pvar_ArrCampi(0 To 0, 0 To 1) As Variant
    
    str_SQL = " SELECT" & _
              "    MG87_DESCARTEST as MG87_DESCART" & _
              " FROM" & _
              "    MG66_ANAGRART WITH (NOLOCK)" & _
              " INNER JOIN MG87_ARTDESC WITH (NOLOCK) ON" & _
              "    MG66_DITTA_CG18 = MG87_DITTA_CG18 AND" & _
              "    MG66_CODART = MG87_CODART_MG66" & _
              " WHERE" & _
              "    MG87_DITTA_CG18 = " & CodiceDitta & " AND" & _
              "    MG87_CODART_MG66 = '" & RTrimN(TXT_GB07_CODART_MG66.Text) & "' AND" & _
              "    MG87_OPZIONE_MG5E = '' AND" & _
              "    MG87_LINGUA_MG52 = ''"
    Set Pvar_ArrCampi(0, 0) = TXT_DESART
    Pvar_ArrCampi(0, 1) = "MG87_DESCART"
    Arr_Fields = Pvar_ArrCampi
    Str_Connect = Gstr_Connect

End Sub

Private Sub TXT_GB07_CODART_MG66_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
  On Error Resume Next
  
  Call Cls_LookupMagazzino.ArticoliDiMagazzino
  str_SQL = Cls_LookupMagazzino.StringaSQL
  Arr_Fields = Cls_LookupMagazzino.ArrayFields
  Str_Caption = Cls_LookupMagazzino.Titolo
  Str_Connect = Gstr_Connect
  TXT_GB07_CODART_MG66.IDLookup = Cls_LookupMagazzino.IDLookup

End Sub


Public Function ControllaBlocchiArticolo(ByVal CodiceArticolo As String) As Boolean
    '
    ' Variabili locali
    '
    Dim Pobj_Parameter      As ADODB.Parameter
    Dim Pint_TipoBlocco     As Integer
    Dim Pstr_CodiceBlocco   As String
    Dim Pstr_DescrBlocco    As String
    '
    ' Trap degli errori
    '
    On Error GoTo Err_ControllaBlocchiArticolo
    '
    ' Istanzio l'oggetto Command
    '
    If Gcls_CommandBloccoStatiArt Is Nothing Then
        Set Gcls_CommandBloccoStatiArt = New ADODB.Command
        With Gcls_CommandBloccoStatiArt
            Set .ActiveConnection = Gcon_Connect
            .CommandText = "SPMG_GETBLOCCOART"
            
            .CommandType = adCmdStoredProc
            .NamedParameters = True
            Set Pobj_Parameter = New ADODB.Parameter
            With Pobj_Parameter
                .Name = "@RETURN_VALUE"
                .Type = adInteger
                .Precision = 10
                .NumericScale = 0
                .Direction = adParamReturnValue
            End With
            .Parameters.Append Pobj_Parameter
            Set Pobj_Parameter = Nothing
            Set Pobj_Parameter = New ADODB.Parameter
            With Pobj_Parameter
                .Name = "@DITTA"
                .Type = adNumeric
                .Precision = 5
                .NumericScale = 0
                .Direction = adParamInput
            End With
            .Parameters.Append Pobj_Parameter
            Set Pobj_Parameter = Nothing
            Set Pobj_Parameter = New ADODB.Parameter
            With Pobj_Parameter
                .Name = "@CODICE_ARTICOLO"
                .Type = adVarChar
                .Size = 25
                .Direction = adParamInput
            End With
            .Parameters.Append Pobj_Parameter
            Set Pobj_Parameter = Nothing
            Set Pobj_Parameter = New ADODB.Parameter
            With Pobj_Parameter
                .Name = "@CODICE_DOCUMENTO"
                .Type = adVarChar
                .Size = 14
                .Direction = adParamInput
            End With
            .Parameters.Append Pobj_Parameter
            Set Pobj_Parameter = Nothing
            Set Pobj_Parameter = New ADODB.Parameter
            With Pobj_Parameter
                .Name = "@TIPO_DOCUMENTO"
                .Type = adNumeric
                .Precision = 2
                .NumericScale = 0
                .Direction = adParamInput
            End With
            .Parameters.Append Pobj_Parameter
            Set Pobj_Parameter = Nothing
            Set Pobj_Parameter = New ADODB.Parameter
            With Pobj_Parameter
                .Name = "@SOTTOTIPO_DOCUMENTO"
                .Type = adNumeric
                .Precision = 2
                .NumericScale = 0
                .Direction = adParamInput
            End With
            .Parameters.Append Pobj_Parameter
            Set Pobj_Parameter = Nothing
            Set Pobj_Parameter = New ADODB.Parameter
            With Pobj_Parameter
                .Name = "@TIPO_BLOCCO"
                .Type = adNumeric
                .Precision = 2
                .NumericScale = 0
                .Direction = adParamInputOutput
            End With
            .Parameters.Append Pobj_Parameter
            Set Pobj_Parameter = Nothing
            Set Pobj_Parameter = New ADODB.Parameter
            With Pobj_Parameter
                .Name = "@CODICE_BLOCCO"
                .Type = adVarChar
                .Size = 3
                .Direction = adParamInputOutput
            End With
            .Parameters.Append Pobj_Parameter
            Set Pobj_Parameter = Nothing
            Set Pobj_Parameter = New ADODB.Parameter
            With Pobj_Parameter
                .Name = "@DESCRIZIONE_BLOCCO"
                .Type = adVarChar
                .Size = 30
                .Direction = adParamInputOutput
            End With
            .Parameters.Append Pobj_Parameter
            Set Pobj_Parameter = Nothing
            'Inizio CCS 4.0.0.:Passaggio indicatore c/f alla SP
            Set Pobj_Parameter = New ADODB.Parameter
            With Pobj_Parameter
                .Name = "@TIPOCF"
                .Type = adNumeric
                .Precision = 2
                .NumericScale = 0
                .Direction = adParamInput
            End With
            .Parameters.Append Pobj_Parameter
            Set Pobj_Parameter = Nothing
            'Fine CCS 4.0.0
            .Prepared = True
        End With
    End If
    '
    ' Invoco il recupero del blocco articolo/stati articolo
    '
    With Gcls_CommandBloccoStatiArt
        .Parameters("@DITTA").value = CodiceDitta
        .Parameters("@CODICE_ARTICOLO").value = CodiceArticolo
        
        Call ImpostaDatiDocumento
        
        .Parameters("@CODICE_DOCUMENTO").value = RTrimN(CODICE_DOCUMENTO)
        .Parameters("@TIPO_DOCUMENTO").value = CDecN(TIPO_DOCUMENTO)
        .Parameters("@SOTTOTIPO_DOCUMENTO").value = CDecN(SOTTOTIPO_DOCUMENTO)
        
        .Parameters("@TIPO_BLOCCO").value = Null
        .Parameters("@CODICE_BLOCCO").value = Null
        .Parameters("@DESCRIZIONE_BLOCCO").value = Null
        'Inizio CCS 4.0.0
        .Parameters("@TIPOCF").value = 0
        'Fine CCS 4.0.0
        .Execute , , adExecuteNoRecords
        If IsNull(.Parameters("@TIPO_BLOCCO").value) Then
            ControllaBlocchiArticolo = True
            Exit Function
        End If
        Pint_TipoBlocco = CDecN(.Parameters("@TIPO_BLOCCO").value)
        Pstr_CodiceBlocco = RTrimN(.Parameters("@CODICE_BLOCCO").value)
        Pstr_DescrBlocco = RTrimN(.Parameters("@DESCRIZIONE_BLOCCO").value)
    End With
    '
    ' Visualizzo il messaggio opportuno
    '
'    If Pint_TipoBlocco = 0 Then
'        If VisualizzaWarning("", Pstr_DescrBlocco & vbCr & vbCr & "VUOI FORZARE ?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
'            ControllaBlocchiArticolo = False
'        Else
'            ControllaBlocchiArticolo = True
'            DoEvents
'        End If
'    Else
        Call VisualizzaWarning("", Pstr_DescrBlocco, vbExclamation + vbOKOnly)
        ControllaBlocchiArticolo = False
'    End If
    '
    ' Esco
    '
    Exit Function
Err_ControllaBlocchiArticolo:
    Select Case VisualizzaErrore("ControllaBlocchiArticolo")
        Case vbIgnore
            Resume Next
        Case vbRetry
            Resume
    End Select
    ControllaBlocchiArticolo = False
End Function

Private Sub ImpostaDatiDocumento()
  
  Select Case TipoDocumento
  Case 1
   
        CODICE_DOCUMENTO = CODDOCCarico
        TIPO_DOCUMENTO = 5
        SOTTOTIPO_DOCUMENTO = 1

   
  Case 2
   Exit Sub
  Case 3
   Exit Sub
  Case 4
   Exit Sub
  Case 5
  
      CODICE_DOCUMENTO = CODDOCDDTScarico
      TIPO_DOCUMENTO = 1
      SOTTOTIPO_DOCUMENTO = 3
  
  
  Case 6
    CODICE_DOCUMENTO = CODDOCReso
    TIPO_DOCUMENTO = 6
    SOTTOTIPO_DOCUMENTO = 1
  End Select

End Sub

'




Friend Function ShowError(Optional ByVal NomeSub As String = "", Optional ByVal StatoErrore As Long = 0, Optional ByVal DettaglioMessaggio As String = "", Optional ByVal Modalita As VbMsgBoxStyle = vbAbortRetryIgnore + vbCritical + vbDefaultButton2) As VbMsgBoxResult
    Dim Str_Messaggio       As String
    Dim Bol_ShowError       As Boolean
    Dim MsgStyle            As VbMsgBoxStyle
    Dim MsgResult           As VbMsgBoxResult
    
    Str_Messaggio = ""
    If NomeSub <> "" Then
        Str_Messaggio = Str_Messaggio & "Si e' verificato il seguente errore nella routine: " & NomeSub
    End If
    If Err.Number <> 0 Then
        If Str_Messaggio <> "" Then
            Str_Messaggio = Str_Messaggio & vbCr & " " & vbCr
        End If
        Str_Messaggio = Str_Messaggio & "Errore VB: " & CStr(Err.Number) & " " & Err.Description & vbCr & _
                                        "Source: " & Err.Source
        Err.Clear
    End If
    If Not (Gcon_Connect Is Nothing) Then
        If Gcon_Connect.Errors.Count > 0 Then
            If Str_Messaggio <> "" Then
                Str_Messaggio = Str_Messaggio & vbCr & " " & vbCr
            End If
            Str_Messaggio = Str_Messaggio & "Errore ADO: " & Gcon_Connect.Errors(0).Description & vbCr & _
                                            "Source: " & Gcon_Connect.Errors(0).Source & vbCr & _
                                            "SQL State: " & Gcon_Connect.Errors(0).SQLState
            Gcon_Connect.Errors.Clear
        End If
    End If
    If DettaglioMessaggio <> "" Then
        If Str_Messaggio <> "" Then
            Str_Messaggio = Str_Messaggio & vbCr & " " & vbCr
        End If
        Str_Messaggio = Str_Messaggio & DettaglioMessaggio
    End If
        
    lng_Stato = StatoErrore
    str_StatoDesc = Str_Messaggio
    
    Bol_ShowError = True
    MsgStyle = Modalita
    RaiseEvent ErrorsOccurred(Str_Messaggio, Bol_ShowError, MsgStyle, MsgResult)
    
    If Bol_ShowError Then
        ShowError = MsgBox(Str_Messaggio, MsgStyle, "Errore !!!")
    Else
        ShowError = MsgResult
    End If
End Function

Private Sub TXT_GB07_CODART_MG66_BeforeItem(Cancel As Boolean)
    On Error Resume Next
    
    Pstr_old_codart = RTrimN(TXT_GB07_CODART_MG66.Text)
'    Pstr_old_variante = ""
'    Pstr_old_codartODL = Pstr_old_codart ' michele 29/04/05
    Pbol_ReturnPressed = False
End Sub

Private Sub TXT_GB07_CODART_MG66_Validate(Cancel As Boolean)

      Dim str_SQL                 As String
    Dim Prst_Recordset          As ADODB.Recordset
    Dim Pbol_ArticoloTrovato    As Boolean
    Dim Pvar_BarCode            As Variant
    Dim Pvar_Opzione            As Variant
    Dim Pvar_PzConf             As Variant
    Dim Pvar_ArtCli             As Variant
    Dim Pvar_ArtFor             As Variant
    Dim Pbol_MovLotti           As Boolean
    Dim Pvar_OldQta             As Variant
    Dim Pvar_OldQta2            As Variant
    Dim Pbol_ArticoloModificato As Boolean
    Dim Pint_Index              As Integer
    Dim Pbol_FlgLotti           As Boolean
    Dim Pbol_FlgSerialNumber    As Boolean
    Dim Pvar_CodLotto           As Variant
    Dim Pvar_SerialNumber       As Variant
    Dim Pvar_Scadenza           As Variant
    Dim Pvar_Qta                As Variant
    
    
    On Error Resume Next
    isBarcode = False
    If RTrimN(TXT_GB07_CODART_MG66.Text) = "" Then
      Exit Sub
    End If
    
'    If Not TXT_GB07_CODART_MG66.IsValid Then
'        Cancel = True
'        MsgBox "Articolo errato"
'        Pstr_old_codart = ""
'        TXT_GB07_CODART_MG66.SetTextFocus
'        Exit Sub
'    End If
    
    Pbol_ReturnPressed = False
    
    If RTrimN(TXT_GB07_CODART_MG66.Text) = Pstr_old_codart Then
        Exit Sub
    End If
    
    Pbol_ArticoloModificato = False
    
    Pstr_old_codart = RTrimN(TXT_GB07_CODART_MG66.Text)
    
    Pvar_BarCode = Null
    Pvar_Opzione = Null
    Pvar_PzConf = 0
    Pvar_ArtCli = Null
    Pvar_ArtFor = Null
    Pvar_CodLotto = Null
    Pvar_SerialNumber = Null
    Pvar_Scadenza = Null
    Pvar_Qta = 0
    
    str_SQL = " SELECT" & _
              "    MG66_CODART" & _
              " FROM" & _
              "    MG66_ANAGRART WITH (NOLOCK)" & _
              " WHERE" & _
              "    MG66_DITTA_CG18 = " & CodiceDitta & " AND" & _
              "    MG66_CODART = '" & Replace(Pstr_old_codart, "'", "''") & "'"
    Set Prst_Recordset = Gcon_Connect.Execute(str_SQL, , adCmdText)
    Pbol_ArticoloTrovato = Not (Prst_Recordset.EOF)
    Prst_Recordset.Close
    Set Prst_Recordset = Nothing
    
    If Not Pbol_ArticoloTrovato Then
        str_SQL = " SELECT" & _
                  "    ARTICOLO ," & _
                  "     MG66_OPZIONE_MG5E," & _
                  "     lotto," & _
                  "     qta," & _
                  "    scadenza " & _
                  " FROM" & _
                  "     GSW03_BARCODE WITH (NOLOCK)" & _
                  " WHERE" & _
                  "    magazzino = " & CodiceDitta & " AND" & _
                  "    barcode = '" & Replace(Pstr_old_codart, "'", "''") & "'"
        Set Prst_Recordset = Gcon_Connect.Execute(str_SQL, , adCmdText)
        If Not Prst_Recordset.EOF Then
            Pvar_BarCode = Pstr_old_codart
            Pstr_old_codart = RTrimN(Prst_Recordset("ARTICOLO").value)
            Pvar_Opzione = RTrimN(Prst_Recordset("MG65_OPZIONE_MG5E").value)
            Pvar_Qta = CDecN(Prst_Recordset("qta").value)
            Pvar_CodLotto = RTrimN(Prst_Recordset("lotto").value)
            Pvar_Scadenza = RTrimN(NVL(Prst_Recordset("scadenza").value, Null))
            Pbol_ArticoloTrovato = True
            isBarcode = True
        End If
        Prst_Recordset.Close
        Set Prst_Recordset = Nothing
    End If
  
    If Not Pbol_ArticoloTrovato Then
        Exit Sub
    End If

  
    TXT_GB07_CODART_MG66.Text = Pstr_old_codart
'    If Pvar_Qta = 0 Then Pvar_Qta = 1
'    TXT_GB07_QTA.Text = Pvar_Qta
    TXT_GB07_QTA.Text = QtaRead '* GetPzConfArt(TXT_GB07_CODART_MG66.Text, " ")
    
    If Not ControllaBlocchiArticolo(Pstr_old_codart) Then
        Cancel = True
        TXT_GB07_CODART_MG66.Text = ""
        Pstr_old_codart = ""
   
        TXT_GB07_CODART_MG66.SetTextFocus
        Exit Sub
    End If
    
   

End Sub

Private Sub TXT_GB07_QTA_AfterItem(Cancel As Boolean)
  
  If NVL(TXT_GB07_CODART_MG66.Text, "") = "" Then
   
  Else
    GridNavDocumenti.Nuovo = True
    GridNavDocumenti.Conferma = False
    GridNavDocumenti.Elimina = False
    GridNavDocumenti.Annulla = False
    GridNavDocumenti.SetButtonFocus "Nuovo"
  
    GridNavDocumenti.Nuovo = True
    GridNavDocumenti.Conferma = True
    GridNavDocumenti.Elimina = True
    GridNavDocumenti.Annulla = True
  
  End If
  
End Sub



Private Sub PulisciCampi()
  TXT_GB06_CLIFOR_CG44.Text = ""
 ' TXT_GB06_CODDESTIN_MG22.Text = ""
 ' TXT_GB06_VETTORE_MG14.Text = ""
  TXT_GB06_CODPAG_CG62.Text = ""
 ' TXT_SETTORE.Text = ""
 ' TXT_SERVIZIO.Text = ""
  TXT_GB06_BUDGET.Text = ""
  TXT_GB06_FORECAST.Text = ""
  TXT_GB06_CONSUNTIVO.Text = ""
  TXT_GB06_AGENTE.Text = ""
  TXT_GB06_BUDGET.Text = ""
  TXT_GB06_CLIFOR_CG44.Text = ""
  TXT_GB06_CODCOMM.Text = ""
  TXT_GB06_CODDESTIN_MG22.Text = ""
  TXT_GB06_CODPAG_CG62.Text = ""
  TXT_GB06_DTCHIUSURA.Text = ""
  TXT_GB06_DTDOC.Text = ""
  TXT_GB06_PERCCHIUSURA.Text = ""
  TXT_GB06_RESPONSABILE.Text = ""
  TXT_GB06_PERCTRASP.Text = ""
  TXT_GB06_PERCRIBGARA.Text = ""
  TXT_GB06_STATODOC.Text = ""
  TXT_GB06_TIPOAREA.Text = ""
  TXT_GB06_TIPOOFFERTA.Text = ""
'  TxtLog.Text = ""
'  TXT_RIFERIMENTO.Text = ""
'  TXT_DATARIF.Text = Null
  TXT_GB06_NREV.Text = ""
  TXT_GB06_NVERS.Text = ""
  TXT_GB06_NUMDOC.Text = ""
  TXT_GB06_NOMEOFFERTA.Text = ""
  TXT_GB06_ID.Text = ""
  TXT_GB06_PROPRIETARIO.Text = ""
End Sub


Private Sub StatoCampi()
On Error Resume Next
  If TXT_GB06_STATODOC.Text = "08" Or TXT_GB06_STATODOC.Text = "04" Or TXT_GB06_STATODOC.Text = "05" Or TXT_GB06_STATODOC.Text = "06" Then
    TXT_GB06_CLIFOR_CG44.Enabled = False
  '  TXT_GB06_CODDESTIN_MG22.Enabled = False
  '  TXT_GB06_VETTORE_MG14.Enabled = False
    TXT_GB06_CODPAG_CG62.Enabled = False
  '  TXT_SETTORE.Enabled = False
'    TXT_SERVIZIO.Enabled = False
    TXT_GB06_BUDGET.Enabled = False
    TXT_GB06_FORECAST.Enabled = False
    TXT_GB06_CONSUNTIVO.Enabled = False
    TXT_GB06_AGENTE.Enabled = False
    TXT_GB06_BUDGET.Enabled = False
    TXT_GB06_CLIFOR_CG44.Enabled = False
    TXT_GB06_CODCOMM.Enabled = False
'    TXT_GB06_CODDESTIN_MG22.Enabled = False
    TXT_GB06_CODPAG_CG62.Enabled = False
    TXT_GB06_DTCHIUSURA.Enabled = False
    TXT_GB06_DTDOC.Enabled = False
    TXT_GB06_PERCCHIUSURA.Enabled = False
    TXT_GB06_RESPONSABILE.Enabled = False
    TXT_GB06_PERCTRASP.Enabled = False
    TXT_GB06_PERCRIBGARA.Enabled = False
    TXT_GB06_STATODOC.Enabled = False
    TXT_GB06_TIPOAREA.Enabled = False
    TXT_GB06_TIPOOFFERTA.Enabled = False
'    TxtLog.Enabled = False
'    TXT_RIFERIMENTO.Enabled = False
'    TXT_DATARIF.Enabled = False
    TXT_GB06_NREV.Enabled = False
    TXT_GB06_NVERS.Enabled = False
    TXT_GB06_NUMDOC.Enabled = False
    TXT_GB06_NOMEOFFERTA.Enabled = False
    'TXT_GB06_ID.Enabled = False
    TXT_GB06_PROPRIETARIO.Enabled = False
    CMD_GENERAOFFERTA.Enabled = False
    cmdNewRevision.Enabled = False
    cmdNewVersion.Enabled = False
    GridNavDocumenti.EnableButtons False, False, False, False, False, False, False
    CMD_SAVE.Enabled = False
    CMD_ELIMINA.Enabled = True
    Call TXT_GB06_STATODOC_AfterItem(True)
    
    
  Else
    TXT_GB06_CLIFOR_CG44.Enabled = True
'    TXT_GB06_CODDESTIN_MG22.Enabled = True
'    TXT_GB06_VETTORE_MG14.Enabled = True
    TXT_GB06_CODPAG_CG62.Enabled = True
'    TXT_SETTORE.Enabled = True
'    TXT_SERVIZIO.Enabled = True
    TXT_GB06_BUDGET.Enabled = True
    TXT_GB06_FORECAST.Enabled = True
    TXT_GB06_CONSUNTIVO.Enabled = True
    TXT_GB06_AGENTE.Enabled = True
    TXT_GB06_BUDGET.Enabled = True
    TXT_GB06_CLIFOR_CG44.Enabled = True
    TXT_GB06_CODCOMM.Enabled = True
'    TXT_GB06_CODDESTIN_MG22.Enabled = True
    TXT_GB06_CODPAG_CG62.Enabled = True
    TXT_GB06_DTCHIUSURA.Enabled = True
    TXT_GB06_DTDOC.Enabled = True
    TXT_GB06_PERCCHIUSURA.Enabled = True
    TXT_GB06_RESPONSABILE.Enabled = True
    TXT_GB06_PERCTRASP.Enabled = True
    TXT_GB06_PERCRIBGARA.Enabled = True
    TXT_GB06_STATODOC.Enabled = True
    TXT_GB06_TIPOAREA.Enabled = True
    TXT_GB06_TIPOOFFERTA.Enabled = True
'    TxtLog.Enabled = True
'    TXT_RIFERIMENTO.Enabled = True
'    TXT_DATARIF.Enabled = True
    TXT_GB06_NREV.Enabled = True
    TXT_GB06_NVERS.Enabled = True
    TXT_GB06_NUMDOC.Enabled = True
    TXT_GB06_NOMEOFFERTA.Enabled = True
    TXT_GB06_ID.Enabled = True
    TXT_GB06_PROPRIETARIO.Enabled = True
    CMD_GENERAOFFERTA.Enabled = True
    cmdNewRevision.Enabled = True
    cmdNewVersion.Enabled = True
    GridNavDocumenti.EnableButtons True, True, True, True, True, True, True
    CMD_SAVE.Enabled = True
    CMD_ELIMINA.Enabled = True
    Call TXT_GB06_STATODOC_AfterItem(True)
End If
End Sub





Private Sub cmdAnnulla_ButtonClick()



On Error GoTo ErrTrap

 
  
  'pulisce campi e torna a mappa iniziale
  
  Call PulisciCampi
  Call StatoCampi
  
  
  
  FrameDoc.Visible = False

  Call DistruggiFramework
  
 
  
  TabDocumenti.ActiveTab = 0
  cmdScarico.SetFocus
  
  Exit Sub
ErrTrap:
  Select Case VisualizzaErrore("cmdAnnulla_ButtonClick")
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
  End Select
End Sub

Private Sub cmdNuovoDoc_ButtonClick()
On Error GoTo ErrTrap

 
  
 
  Call PulisciCampi
  
  
  
  TabDocumenti.ActiveTab = 0
  'Call InserisciTestata

  
  Exit Sub
ErrTrap:
  Select Case VisualizzaErrore("cmdNuovoDoc_ButtonClick")
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
  End Select
End Sub

Private Sub cmdStampa_ButtonClick()
  If NVL(NumRegGenerato, "") <> "" Then
    Call StampaDocumento
  End If


End Sub

Private Sub ScriviDocumento(TipoCf As Integer, CodFor As Double, isPosa As Boolean)

Dim strCodDestinazioneFornitore As String
Dim IdNEWDest As String
Dim CodiceDestForn As String

On Error GoTo SegnalaErrore
  Dim strRiferimento As String
  
  Dim Ret As Boolean
  Dim CondizionePagamento As String
  If TipoCf = 0 Then
  CondizionePagamento = NVL(TXT_GB06_CODPAG_CG62.Text, "")
  Else
  
  End If
  
  ContaRighe = 0
  Set ClasseInternaRegDoc = New CLSBO_REGDOC

  Set ClasseInternaRegDoc.Gcon_Connect = Gcon_Connect
  Set ClasseInternaRegDoc.ActiveInterface = ActiveInterface
  
  If TipoCf = 0 Then
  ClasseInternaRegDoc.InizializzaBODocumenti True, False, Trim(CODDOCOfferta)
  Else
  If isPosa Then
        ClasseInternaRegDoc.InizializzaBODocumenti True, False, Trim(CODDOCOrdinePosa)
    Else
        ClasseInternaRegDoc.InizializzaBODocumenti True, False, Trim(CODDOCOrdine)
  End If
  End If
  
  If ClasseInternaRegDoc.Errore <> tsOK Then
    MsgBox ClasseInternaRegDoc.ErrDescr
    ContaRighe = 0
    Exit Sub
  End If
  
  
  strRiferimento = "Rif. ns." & Trim(TXT_GB06_NUMDOC.Text) & " del " & CStr(NVL(TXT_GB06_DTDOC.Text, Now()))
  strRiferimento = strRiferimento & "Rev: " & CStr(TXT_GB06_NREV.Text) & " Versione " & CStr(TXT_GB06_NVERS.Text)
  strRiferimento = strRiferimento & "Ope: " & NVL(TXT_GB06_PROPRIETARIO.Text, "")
  strRiferimento = Mid(strRiferimento, 1, 72)
  If TipoCf = 0 Then
    ClasseInternaRegDoc.AddNewTestataDocumento TXT_GB06_CLIFOR_CG44.Text, NVL(TXT_GB06_AGENTE.Text, ""), , , _
                                    , , NVL(TXT_GB06_CODPAG_CG62.Text, ""), , _
                                   TXT_GB06_DTDOC.Text, _
                                   TXT_GB06_DTDOC.Text, , , , , , , , , , , NVL(TXT_GB06_CODDESTIN_MG22.Text, "") _
                                   , _
                                   , NVL(TXT_GB06_NOMEOFFERTA.Text, ""), , _
                                   , Trim(TXT_GB06_NUMDOC.Text) & "-" & CStr(TXT_GB06_NREV.Text) & "-" & CStr(TXT_GB06_NVERS.Text), IDGB06, strRiferimento, IDGB06, TXT_GB06_DTDOC.Text, Now(), TXT_GB06_CIG.Text, TXT_GB06_CUP.Text, TXT_GB06_CODCOMM.Text, , , NVL(TXT_GB06_RESPONSABILE.Text, "")
  Else
  'genera la destinazione fornitore
'    IdNEWDest = CStr(Now())
'    IdNEWDest = Replace(IdNEWDest, "/", "")
'    IdNEWDest = Replace(IdNEWDest, " ", "")
'    IdNEWDest = Replace(IdNEWDest, ":", "")
    IdNEWDest = Format(NVL(CodFor, ""), "00000000") & "-" & Format(NVL(TXT_GB06_CODDESTIN_MG22.Text, ""), "0000")
    strCodDestinazioneFornitore = " SELECT MG22_CODDESTIN  " & _
                               " From MG22_CLIFORDEST " & _
                               " WHERE MG22_DITTA_CG18 = " & CodiceDitta & _
                               " AND (MG22_TIPOCF_CG44 = 1) " & _
                               " AND (MG22_CODALF = '" & IdNEWDest & "') "
                               
    CodiceDestForn = NVL(GetValFromQuery(strCodDestinazioneFornitore, 0, Gcon_Connect), "")
    If CodiceDestForn = "" Then
        strCodDestinazioneFornitore = "INSERT INTO MG22_CLIFORDEST "
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & " ( "
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & " MG22_DITTA_CG18 ,MG22_TIPOCF_CG44,MG22_CLIFOR_CG44,MG22_CODDESTIN,MG22_DESTRAGSOC,MG22_DESTIND,MG22_DESTCAP,"
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & " MG22_DESTCITTA,MG22_DESTPROV,MG22_DESTTEL,MG22_DESTCELL,MG22_DESTEMAIL,MG22_DESTFAX,MG22_DESTNOTE,"
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & " MG22_ALIQRID_CG28,MG22_CODABI_CG12,MG22_CODCAB_CG13,MG22_FLGSTETIC,MG22_CODALF,MG22_IDMEDIA_CG99,"
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & " MG22_DESTRAGSOCEX,MG22_DESTINDEX,MG22_STATOEST_CG07,MG22_INDPREFSTDOC,MG22_DESTCAPCHAR,MG22_AGENTE_MG17,"
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & " MG22_MACROAREA_MG07,MG22_AREA_MG08,MG22_ZONA_MG09,MG22_VETT1_MG14,MG22_VETT2_MG14,MG22_CODLINGUA_MG52,"
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & " MG22_TIPODESTXDOC,MG22_CODICE_CG16,MG22_MACROCAT,MG22_CATEG,MG22_SOTTOCAT,MG22_RAGGRCF1,MG22_RAGGRCF2,MG22_RAGGRCF3,"
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & " MG22_GUID , MG22_DESTPEC, MG22_FLGAPPIVA, MG22_DATAVALIVA"
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & " ) "
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & "    SELECT      MG22_DITTA_CG18, 1, " & CodFor & ","
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & "  ( "
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & "  SELECT       isnull( MAX(cast(MG22_CODDESTIN as int)),0) + 1 AS newdest "
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & " From MG22_CLIFORDEST"
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & " Where (MG22_DITTA_CG18 = " & CodiceDitta & ") And (MG22_TIPOCF_CG44 = 1) And (MG22_CLIFOR_CG44 = " & CodFor & ")"
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & "  ) as FORDEST"
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & "  , MG22_DESTRAGSOC, MG22_DESTIND, MG22_DESTCAP, MG22_DESTCITTA, MG22_DESTPROV, MG22_DESTTEL,"
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & "   MG22_DESTCELL, MG22_DESTEMAIL, MG22_DESTFAX, MG22_DESTNOTE, MG22_ALIQRID_CG28, MG22_CODABI_CG12, MG22_CODCAB_CG13, MG22_FLGSTETIC, '" & IdNEWDest & "', MG22_IDMEDIA_CG99,"
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & "   MG22_DESTRAGSOCEX, MG22_DESTINDEX, MG22_STATOEST_CG07, MG22_INDPREFSTDOC, MG22_DESTCAPCHAR, MG22_AGENTE_MG17, MG22_MACROAREA_MG07, MG22_AREA_MG08,"
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & "   MG22_ZONA_MG09, MG22_VETT1_MG14, MG22_VETT2_MG14, MG22_CODLINGUA_MG52, MG22_TIPODESTXDOC, MG22_CODICE_CG16, MG22_MACROCAT, MG22_CATEG, MG22_SOTTOCAT,"
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & "   MG22_RAGGRCF1 , MG22_RAGGRCF2, MG22_RAGGRCF3, NEWID(), MG22_DESTPEC, MG22_FLGAPPIVA, MG22_DATAVALIVA"
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & " From MG22_CLIFORDEST"
        strCodDestinazioneFornitore = strCodDestinazioneFornitore & " WHERE        (MG22_DITTA_CG18 = " & CodiceDitta & ") AND (MG22_TIPOCF_CG44 = 0) AND (MG22_CLIFOR_CG44 = " & TXT_GB06_CLIFOR_CG44.Text & ") AND (MG22_CODDESTIN = '" & NVL(TXT_GB06_CODDESTIN_MG22.Text, "") & "')"
        Gcon_Connect.Execute strCodDestinazioneFornitore
    End If
    
    
  strCodDestinazioneFornitore = " SELECT MG22_CODDESTIN  " & _
                               " From MG22_CLIFORDEST " & _
                               " WHERE MG22_DITTA_CG18 = " & CodiceDitta & _
                               " AND (MG22_TIPOCF_CG44 = 1) " & _
                               " AND (MG22_CLIFOR_CG44 = " & CodFor & ") " & _
                               " AND (MG22_CODALF = '" & IdNEWDest & "') "
  
  
    ClasseInternaRegDoc.AddNewTestataDocumento CodFor, , , , _
                                    , , , , _
                                   TXT_GB06_DTDOC.Text, _
                                   TXT_GB06_DTDOC.Text, , , , , , , , , , , NVL(GetValFromQuery(strCodDestinazioneFornitore, 0, Gcon_Connect), "") _
                                   , _
                                   , NVL(TXT_GB06_NOMEOFFERTA.Text, ""), , _
                                   , Trim(TXT_GB06_NUMDOC.Text) & "-" & CStr(TXT_GB06_NREV.Text) & "-" & CStr(TXT_GB06_NVERS.Text), IDGB06, strRiferimento, IDGB06, TXT_GB06_DTDOC.Text, Now(), TXT_GB06_CIG.Text, TXT_GB06_CUP.Text, TXT_GB06_CODCOMM.Text, , , NVL(TXT_GB06_RESPONSABILE.Text, ""), Mid(NVL(TXT_GB06_AGENTE_DEC.Text, ""), 1, 72), NVL(TXT_GB06_CIG.Text, ""), NVL(TXT_GB06_CUP.Text, "")

  End If
 

  If ClasseInternaRegDoc.Errore <> tsOK Then
    MsgBox ClasseInternaRegDoc.ErrDescr
    ContaRighe = 0
    Exit Sub
  End If

  If TipoCf = 0 Then
    If Not (rstScriveDoc.EOF) Then
    Ret = ScriviCorpoGamma(0, TipoCf, CodFor)
    If Ret Then
          'ContaRighe = ContaRighe + 1
        Else
          ContaRighe = 0
          rstScriveDoc.MoveLast
        End If
    End If
  End If
  
  Do While Not (rstScriveDoc.EOF)
    'se la quantità è maggiore di 0 scrivo la riga
    If NVL(rstScriveDoc("GB07_QTA"), 0) <> 0 Then
      Ret = ScriviCorpoGamma(1, TipoCf, CodFor)
      If Ret Then
        ContaRighe = ContaRighe + 1
      Else
        ContaRighe = 0
        Exit Do
      End If
    End If
    rstScriveDoc.MoveNext
  Loop
  
  If ContaRighe > 0 Then
    'Salva modifiche
    ClasseInternaRegDoc.RegistraDocumento
    NumDocGenerato = ClasseInternaRegDoc.NumDocumento
    NumRegGenerato = ClasseInternaRegDoc.NumregDocumento
    ClasseInternaRegDoc.RilasciaBODocumenti
    
    'aggiorna cig e cup
    strSQL = "UPDATE DO11_DOCTESTATA "
    strSQL = strSQL & " SET  DO11_CIG    = '" & NVL(TXT_GB06_CIG.Text, "") & "'"
    strSQL = strSQL & "    , DO11_CUP         = '" & NVL(TXT_GB06_CUP.Text, "") & "'"
    strSQL = strSQL & " WHERE DO11_NUMREG_CO99   = '" & NumRegGenerato & "'"
    strSQL = strSQL & "   AND DO11_DITTA_CG18   = " & CodiceDitta
    Gcon_Connect.Execute strSQL
    
    strSQL = "UPDATE DO30_DOCCORPO "
    strSQL = strSQL & " SET  DO30_CIG    = '" & NVL(TXT_GB06_CIG.Text, "") & "'"
    strSQL = strSQL & "    , DO30_CUP         = '" & NVL(TXT_GB06_CUP.Text, "") & "'"
    strSQL = strSQL & " WHERE DO30_NUMREG_CO99   = '" & NumRegGenerato & "'"
    strSQL = strSQL & "   AND DO30_DITTA_CG18   = " & CodiceDitta
    Gcon_Connect.Execute strSQL
    
    
    FrameDoc.Visible = False
    
    'Scrive dati in GB06
    strSQL = "UPDATE Offers "
    strSQL = strSQL & " SET  Status    = 2 "
    'strSQL = strSQL & "    , GB06_STATODOC         = '05'"
    strSQL = strSQL & " WHERE Code   = '" & NVL(TXT_GB06_NUMDOC.Text, "9999999999999") & "'"
    Gcon_Connect.Execute strSQL
    
    strSQL = "INSERT INTO GB09_ORDINI (GB09_ID_GB06, GB09_NUMREG_CO99, GB09_DESCRIZIONE)"
    strSQL = strSQL & "VALUES"
    strSQL = strSQL & "                        (" & IDGB06 & ", '" & NVL(NumRegGenerato, "") & "', '')"
    
    Gcon_Connect.Execute strSQL
    'aggiorna
   
    TXT_GB06_STATODOC.Text = "05"
    Call TXT_GB06_STATODOC_AfterItem(False)
    
    
  End If
  
  
  Set ClasseInternaRegDoc = Nothing
  
  Exit Sub
SegnalaErrore:
  'Scrivi log errore
  Errore = "Attenzione! L'applicazione ha generato il seguente errore: " & Err.Number & " - " & Err.Description
'  ClasseInternaRegDoc = Nothing
  Select Case VisualizzaErrore("ScriviDocumento")
    Case vbAbort
        ContaRighe = 0
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
  End Select

End Sub

Private Sub ApriDoc(CodiceDoc As String, Numreg As Variant)

    On Error GoTo ErrTrap
    Dim ProgramInterface   As Cinterface
    Dim CalledProgram      As Object
    Dim a As Variant
    
    Set CalledProgram = CreateObject("MGUO_DOCUMENTI.CLSMG_DOCUMENTI")
    Set ProgramInterface = CalledProgram
    
    ProgramInterface.IsActive = True
    ProgramInterface.IsCalled = True
   
    
    Set ActiveInterface.ClsGlobal.CallInterface = ProgramInterface
'    Codice =
    If NVL(CodiceDoc, "") <> "" And NVL(Numreg, "") <> "" Then
        a = CalledProgram.ActualParameters
        a.Codice = CodiceDoc
        a.Numreg = Numreg
        CalledProgram.ActualParameters = a
        
    End If
    

    ActiveInterface.ClsGlobal.ExecDll False, "MGUO_DOCUMENTI.CLSMG_DOCUMENTI", False, tsModify, Normale, , , False, True
    
    Set ProgramInterface = Nothing
    Set CalledProgram = Nothing
    Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("CallGestioneDocumenti")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub

Function ScriviCorpoGamma(TipoRiga As Integer, TipoCf As Integer, CodFor As Double) As Boolean
  On Error GoTo SegnalaErrore
  
'  ClasseInternaRegDoc.AddNewRigaDocumento tsImportDocumentiArticoloMagazzino, _
'                          rstCorpoBanco("GB07_CODART_MG66"), , , _
'                          "00", , , , , , _
'                          NVL(rstCorpoBanco("GB07_PREZZO"), 0), , _
'                          NVL(rstCorpoBanco("GB07_QTA"), 0), , _
'                          NVL(rstCorpoBanco("GB07_SCCORPO"), 0), _
'                          NVL(rstCorpoBanco("GB07_SCPIEDE"), 0)
If TipoCf = 0 Then
    If TipoRiga = 1 Then
    
     ClasseInternaRegDoc.AddNewRigaDocumento tsImportDocumentiArticoloMagazzino, _
                              SQLString(rstScriveDoc("GB07_CODART_MG66")), , , _
                              , , , , , , Format(Round(NVL(rstScriveDoc("GB07_PREZZO"), 0), 1), "0.00") _
                              , , _
                              NVL(rstScriveDoc("GB07_QTA"), 0), , NVL(rstScriveDoc("GB07_SC1"), 0) _
                              , NVL(rstScriveDoc("GB07_SC2"), 0) _
                              , NVL(rstScriveDoc("GB07_SC3"), 0), NVL(rstScriveDoc("GB07_SC4"), 0), NVL(rstScriveDoc("GB07_SC5"), 0), NVL(rstScriveDoc("GB07_SC6"), 0), , SQLString(NVL(rstScriveDoc("GB07_DESCART"), "")), , , , , , , , , , , , , , , , , TXT_GB06_CODCOMM.Text, , , , , , , , , , "In produzione"


      ScriviCorpoGamma = True
      Else

      ClasseInternaRegDoc.AddNewRigaDocumento tsImportDocumentiRigaDescrittiva, _
                              , , , _
                              , , , , , , 0 _
                              , , _
                              1, , _
                              , _
                              , , , , , , SQLString(TXT_GB06_TEXT1.Text), , , , , , , , , , , , , , , , , TXT_GB06_CODCOMM.Text

      ScriviCorpoGamma = True
      
'      ClasseInternaRegDoc.AddNewRigaDocumento tsImportDocumentiRigaDescrittiva, _
'                              , , , _
'                              , , , , , , 0 _
'                              , , _
'                              0, , _
'                              , _
'                              , , , , , , SQLString(rstScriveDoc("GB07_CODART_MG66")) & " - " & SQLString(NVL(rstScriveDoc("GB07_DESCART"), ""))
'
'
'      ScriviCorpoGamma = True
'      Else
'
'      ClasseInternaRegDoc.AddNewRigaDocumento tsImportDocumentiArticoloManuale, _
'                              , , , _
'                              , , , , , , CDbl(TXT_GB06_CONSUNTIVO.Text) _
'                              , , _
'                              1, , _
'                              , _
'                              , , , , , , SQLString(TXT_GB06_TEXT1.Text), , , , , , , , , , , , , , , , , TXT_GB06_CODCOMM.Text
'
'      ScriviCorpoGamma = True
      
      End If
  Else
   'tsImportDocumentiArticoloMagazzino
   'GB07_costo
       ClasseInternaRegDoc.AddNewRigaDocumento tsImportDocumentiRigaDescrittiva, _
                              SQLString(rstScriveDoc("GB07_CODART_MG66")), , , _
                              , , , , , , _
                              Format(Round(rstScriveDoc("GB07_costo"), 1), "0.00"), , _
                              rstScriveDoc("GB07_QTA"), , 0 _
                              , 0 _
                              , 0, 0, 0, 0, , SQLString(NVL(rstScriveDoc("GB07_DESCART"), "")), , , , , , , , , , , , , , , , , TXT_GB06_CODCOMM.Text
      
      
      
      ScriviCorpoGamma = True
  End If
  
  Exit Function
SegnalaErrore:
  'Scrivi log errore
  Errore = "Attenzione! L'applicazione ha generato il seguente errore: " & Err.Number & " - " & Err.Description
  ScriviCorpoGamma = False


End Function



Private Function CallDocumenti(ByVal rCodDoc As String, rNumReg As String)
  On Error Resume Next
    
    '
    ' Disattivo la Dll chiamante / Attivo la Dll chiamata
    '
    ActiveInterface.IsActive = False
    Set Cls_ConnectMagazzino.ActiveInterface = ActiveInterface
    Cls_ConnectMagazzino.Left = 0
    Cls_ConnectMagazzino.Top = 0
    Cls_ConnectMagazzino.WindowModal = False
    Call Cls_ConnectMagazzino.Documenti(rCodDoc, rNumReg)
    
    
  Err.Clear
End Function

Private Sub StampaDocumento()
'
On Error GoTo ErrTrap
  
  Set ClasseInternaRegDoc = New CLSBO_REGDOC

  Set ClasseInternaRegDoc.Gcon_Connect = Gcon_Connect
  Set ClasseInternaRegDoc.ActiveInterface = ActiveInterface
  
  ClasseInternaRegDoc.InizializzaBODocumenti True, False, Trim(CodiceDocumento)
    
  Call ClasseInternaRegDoc.StampaDocumento(CodiceDocumento, NumRegGenerato)
  
  ClasseInternaRegDoc.RilasciaBODocumenti
  
  Set ClasseInternaRegDoc = Nothing
  
  Exit Sub
ErrTrap:
  Select Case VisualizzaErrore("StampaDocumento")
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
  End Select
End Sub


Private Sub ScriveFileTestoScontrino()

  On Error GoTo ErrTrap
  Dim Ret         As Boolean
  
  strSQL = "SELECT * "
  strSQL = strSQL & "   FROM DO30_DOCCORPO "
  strSQL = strSQL & "  WHERE DO30_DITTA_CG18  = " & CodiceDitta
  strSQL = strSQL & "    AND DO30_NUMREG_CO99 = '" & NumRegGenerato & "'"
  
  Set rsScontrino = Gcon_Connect.Execute(strSQL, , adCmdText)
  If (rsScontrino.EOF And rsScontrino.BOF) Then
    Set rsScontrino = Nothing
    Exit Sub
  End If
  
  Call CancellaFile(PARAM_DIRSCONTRINI & "\" & strNomeFile)
  strPathFile = PARAM_DIRTEMP
  strNomeFile = PARAM_NOMEFILE
  
  numfile = FreeFile
    
  Open PARAM_DIRSCONTRINI & "\" & strNomeFile For Output As #numfile
  
  Call ScrivoTestata
  
  Do While Not rsScontrino.EOF

    'Scrivo righe
    Call ScrivoRighe

    rsScontrino.MoveNext
  Loop
  
  Call ScrivoPiede
  Set rsScontrino = Nothing

  Close #numfile

  'Copia file per interdoc
  If strNomeFile <> "" Then
   ' Call CopiaFile(strPathFile & "\" & strNomeFile, PARAM_DIRSCONTRINI & "\" & strNomeFile)
    
    'strNomeFileRen = Year(Now) & String(2 - Len(Month(Now)), "0") & Month(Now) & String(2 - Len(Day(Now)), "0") & Day(Now) & String(2 - Len(Hour(Time)), "0") & Hour(Time) & String(2 - Len(Minute(Time)), "0") & Minute(Time) & String(2 - Len(Second(Time)), "0") & Second(Time) & Right(Format(Timer, "#0.00"), 2) & strNomeFile
    
   ' Call CopiaFile(strPathFile & "\" & strNomeFile, PARAM_DIRCOPIA & "\" & strNomeFileRen)
   ' Call CancellaFile(strPathFile & "\" & strNomeFile)
    Call Shell(PARAM_EXESCONTR, vbHide)
    'Call CancellaFile(PARAM_DIRSCONTRINI & "\" & strNomeFile)
  End If


Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("ScriveFileTestoScontrino")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select

End Sub

'Esporta testata
Private Sub ScrivoTestata()
On Error GoTo ErrTrap
  'Scrittura del file secondo tracciato di export
  Dim strExport As String

'  strExport = "CLEAR"
'  Print #numfile, strExport
'
'  strExport = "CHIAVE REG"
'  Print #numfile, strExport

   strExport = "1" & vbCrLf
   strExport = strExport & "9600,N,8,1" & vbCrLf
   strExport = strExport & "7" & vbCrLf
   strExport = strExport & "=K"
   Print #numfile, strExport
 

Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("ScrivoTestata")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub

'Esporta righe
Private Sub ScrivoRighe()
On Error GoTo ErrTrap
  'Scrittura del file secondo tracciato di export
  Dim strExport As String
  Dim Descrizione As String
  
'  If rsScontrino("DO30_QTA1") > 1 Then
'    strExport = "prmsg riga='"
'    strExport = strExport & Replace(Abs(rsScontrino("DO30_QTA1")), ",", ".") & " X " & Replace(rsScontrino("DO30_PREZZO1"), ",", ".")
'    strExport = strExport & "'"
'
'    Print #numfile, strExport
'
'  End If
'
'  If rsScontrino("DO30_QTA1") >= 0 Then
'    strExport = "VEND REP= "
'    Select Case rsScontrino("DO30_ALIVA_CG28")
'    Case "04"
'      strExport = strExport & "3"
'    Case "10"
'      strExport = strExport & "2"
'    Case "121"
'      strExport = strExport & "1"
'    Case "4"
'      strExport = strExport & "3"
'    Case Else
'      strExport = strExport & "1"
'    End Select
'    strExport = strExport & ",PREZZO="
'
'    strExport = strExport & String(8 - Len(Replace(rsScontrino("DO30_IMPORTOIVA"), ",", ".")), " ") & Replace(rsScontrino("DO30_IMPORTOIVA"), ",", ".")
'    strExport = strExport & ",DES='"
'    If Len(Replace(Replace(rsScontrino("DO30_DESCART"), ",", "."), "total", " ")) > 20 Then
'      strExport = strExport & Mid(Replace(Replace(rsScontrino("DO30_DESCART"), ",", "."), "total", " "), 1, 20)
'    Else
'      strExport = strExport & Replace(Replace(rsScontrino("DO30_DESCART"), ",", "."), "total", " ")
'    End If
'    strExport = strExport & "'"
'
'    Print #numfile, strExport
'
'  End If
'
'  If rsScontrino("DO30_QTA1") < 0 Then
'    strExport = "VEND REP= "
'    Select Case rsScontrino("DO30_ALIVA_CG28")
'    Case "04"
'      strExport = strExport & "3"
'    Case "10"
'      strExport = strExport & "2"
'    Case "121"
'      strExport = strExport & "1"
'    Case "4"
'      strExport = strExport & "3"
'    Case Else
'      strExport = strExport & "1"
'    End Select
'    strExport = strExport & ",PREZZO="
'
'    strExport = strExport & String(8 - Len(Replace(Abs(rsScontrino("DO30_IMPORTOIVA")), ",", ".")), " ") & Replace(Abs(rsScontrino("DO30_IMPORTOIVA")), ",", ".")
'    strExport = strExport & ",DES='"
'
'    Descrizione = Replace(Abs(rsScontrino("DO30_QTA1")), ",", ".") & " X " & Replace(rsScontrino("DO30_PREZZO1"), ",", ".") & " " & Replace(Replace(rsScontrino("DO30_DESCART"), ",", "."), "total", " ")
'
'    If Len(Descrizione) > 30 Then
'      strExport = strExport & Mid(Descrizione, 1, 30)
'    Else
'      strExport = strExport & Descrizione
'    End If
'    strExport = strExport & "' , reso"
    If rsScontrino("DO30_QTA1") > 1 Then
'    strExport = "prmsg riga='"
'    strExport = strExport & Replace(Abs(rsScontrino("DO30_QTA1")), ",", ".") & " X " & Replace(rsScontrino("DO30_PREZZO1"), ",", ".")
'    strExport = strExport & "'"
'
'    Print #numfile, strExport
  
  End If
  
  'famiglia articolo per il reparto
  
  Dim strRepArtcolo As String
  strRepArtcolo = "SELECT *"
  
  'fine famiglia articolo per il reparto
  
  If rsScontrino("DO30_QTA1") >= 0 Then
    strExport = "=R"
    Select Case rsScontrino("DO30_ALIVA_CG28")
    Case "04"
      strExport = strExport & "3"
    Case "10"
      strExport = strExport & "2"
    Case "121"
      strExport = strExport & "1"
    Case "4"
      strExport = strExport & "3"
    Case Else
      strExport = strExport & "1"
    End Select
    strExport = strExport & "/$"
    
    'strExport = strExport & Trim(String(8 - Len(Replace(rsScontrino("DO30_IMPORTOIVA"), ",", ".")), " ") & Replace(rsScontrino("DO30_IMPORTOIVA"), ",", "."))
    strExport = strExport & (rsScontrino("DO30_IMPORTOIVA") * 100) / rsScontrino("DO30_QTA1")
    strExport = strExport & "/*"
    strExport = strExport & rsScontrino("DO30_QTA1")
    strExport = strExport & "/("
    If Len(Replace(Replace(rsScontrino("DO30_DESCART"), ",", "."), "total", " ")) > 20 Then
      strExport = strExport & Mid(Replace(Replace(rsScontrino("DO30_DESCART"), ",", "."), "total", " "), 1, 20)
    Else
      strExport = strExport & Replace(Replace(rsScontrino("DO30_DESCART"), ",", "."), "total", " ")
    End If
    strExport = strExport & ")"
    
    Print #numfile, strExport
  
  End If
  
  If rsScontrino("DO30_QTA1") < 0 Then
    strExport = "VEND REP= "
    Select Case rsScontrino("DO30_ALIVA_CG28")
    Case "04"
      strExport = strExport & "3"
    Case "10"
      strExport = strExport & "2"
    Case "121"
      strExport = strExport & "1"
    Case "4"
      strExport = strExport & "3"
    Case Else
      strExport = strExport & "1"
    End Select
    strExport = strExport & ",PREZZO="
    
    strExport = strExport & String(8 - Len(Replace(Abs(rsScontrino("DO30_IMPORTOIVA")), ",", ".")), " ") & Replace(Abs(rsScontrino("DO30_IMPORTOIVA")), ",", ".")
    strExport = strExport & ",DES='"
    
    Descrizione = Replace(Abs(rsScontrino("DO30_QTA1")), ",", ".") & " X " & Replace(rsScontrino("DO30_PREZZO1"), ",", ".") & " " & Replace(Replace(rsScontrino("DO30_DESCART"), ",", "."), "total", " ")
    
    If Len(Descrizione) > 30 Then
      strExport = strExport & Mid(Descrizione, 1, 30)
    Else
      strExport = strExport & Descrizione
    End If
    strExport = strExport & "' , reso"
    Print #numfile, strExport
  
  End If
  

Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("ScrivoRighe")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub

'Esporta piede
Private Sub ScrivoPiede()
On Error GoTo ErrTrap
  'Scrittura del file secondo tracciato di export
  Dim strExport As String

'  If NVL(TXT_CONTANTI.Text, 0) > 0 And NVL(TXT_CONTANTI.Text, 0) > TXT_TOTALEDOC.Text Then
'    strExport = "CHIUS T=1,IMP="
'    strExport = strExport & String(8 - Len(Replace(TXT_CONTANTI.Text, ",", ".")), " ") & Replace(TXT_CONTANTI.Text, ",", ".")
    
'  If ValContanti > 0 Then
'    strExport = "CHIUS T=1,IMP="
'    strExport = strExport & String(8 - Len(Replace(ValContanti, ",", ".")), " ") & Replace(ValContanti, ",", ".")
'    Print #numfile, strExport
'
'  Else
'    strExport = "CHIUS T=1"
'    Print #numfile, strExport
'
'  End If

 strSQL = "SELECT * "
  strSQL = strSQL & "   FROM DO13_DOCTOTALI "
  strSQL = strSQL & "  WHERE DO13_DITTA_CG18  = " & CodiceDitta
  strSQL = strSQL & "    AND DO13_NUMREG_CO99 = '" & NumRegGenerato & "'"
  Dim rsScontrinoTot As ADODB.Recordset
  Set rsScontrinoTot = Gcon_Connect.Execute(strSQL, , adCmdText)
  If (rsScontrinoTot.EOF And rsScontrinoTot.BOF) Then
    Set rsScontrinoTot = Nothing
    Exit Sub
  End If

  If ValContanti > 0 Then
    'strExport = "CHIUS T=1,IMP="
   ' strExport = "=T1/$" & Trim(String(8 - Len(Replace(rsScontrinoTot("DO13_TOTAPAGARE"), ",", ".")), " ") & Replace(rsScontrinoTot("DO13_TOTAPAGARE"), ",", "."))
    strExport = "=T1/$" & rsScontrinoTot("DO13_TOTAPAGARE") * 100
    Print #numfile, strExport

  Else
    strExport = "=T1/$" & Trim(String(8 - Len(Replace(rsScontrinoTot("DO13_TOTAPAGARE"), ",", ".")), " ") & Replace(rsScontrinoTot("DO13_TOTAPAGARE"), ",", "."))
    strExport = "=T1/$" & rsScontrinoTot("DO13_TOTAPAGARE") * 100
    Print #numfile, strExport

  End If

'=T1/$2631
Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("ScrivoPiede")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub



Private Function ElaboraSeFuoriFido(ByVal CodiceCliente As Long, ByVal FlagVisualizzaScheda As Boolean) As Boolean


  'Gestione errori
  On Error GoTo Err_ControlloFido
  ElaboraSeFuoriFido = False
  
  If NVL(CodiceCliente, 0) = 0 Then
    Exit Function
  End If
  
  Dim Pstr_Sql  As String
  Dim MyRst     As ADODB.Recordset
  
  'INIZIO CONTROLLE SE ABILITATO LA GESTIONE FIDO

  Pstr_Sql = " SELECT MG19_INDGESFIDO, MG19_INDCLIBLOC " & _
             " FROM MG19_CLIFORVA " & _
             " WHERE (MG19_DITTA_CG18 = " & CodiceDitta & ") AND (MG19_TIPOCF_CG44 = 0) " & _
             " AND (MG19_CLIFOR_CG44  = " & CodiceCliente & ") "

  Set MyRst = Gcon_Connect.Execute(Pstr_Sql, , adCmdText)
  If Not MyRst.EOF Then
      If Trim$(NVL(MyRst("MG19_INDGESFIDO").value, 0)) <> 2 Then
        ClienteNoFido = True
        ElaboraSeFuoriFido = False
        MyRst.Close
        Set MyRst = Nothing
        
        Exit Function
      End If
  End If
  MyRst.Close
  Set MyRst = Nothing
    
  ClienteNoFido = False
    
  Set Cls_ControlloRischio = New MGBO_CALCRISCHIO.CLSMG_CALCRISCHIO

  Set Cls_ControlloRischio.ActiveConnection = Gcon_Connect
  Cls_ControlloRischio.CodiceDitta = CodiceDitta

  Cls_ControlloRischio.GruppoPDC = ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPDC
  Cls_ControlloRischio.MastroClienti = "0700000000" 'ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.MastroClienti
  Cls_ControlloRischio.MastroClienti = ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.MastroClienti
  
  Cls_ControlloRischio.MastroFornitori = "" 'ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.MastroFornitori

  Cls_ControlloRischio.TipoCf = Cliente   '0 Clienti - 1 Fornitori
  Cls_ControlloRischio.CodiceCF = CodiceCliente
  Cls_ControlloRischio.CodiceDocumento = ""

  Cls_ControlloRischio.TipoElaborazioneRischio = ElaborazioneRischioSingolaNoScheda
  Cls_ControlloRischio.ModalitaSegnalazione = Disattiva
  Cls_ControlloRischio.DataElaborazione = Date

  'Lancio il programma di elaborazione del rischio
  Cls_ControlloRischio.ElaborazioneRischio

  If Cls_ControlloRischio.IsFuoriFido Then
    ElaboraSeFuoriFido = True
    'If FlagVisualizzaScheda Then Cls_ControlloRischio.ApriVisualizzazioneRischio
  Else
    'da togliere
    'If FlagVisualizzaScheda Then Cls_ControlloRischio.ApriVisualizzazioneRischio
    ElaboraSeFuoriFido = False
  End If

  Set Cls_ControlloRischio = Nothing
  
Exit_Handler:
   Exit Function

Err_ControlloFido:
    Select Case VisualizzaErrore("ElaboraSeFuoriFido")
        Case vbAbort
            Exit Function
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select


End Function







Private Sub InvocaStampa()
  Dim strReport               As String
  On Error GoTo ErrTrap
  
  Set PclsReport = Nothing
  Set PclsReport = New FWBO_REPORT30.CLSFW_REPORTCD

  PclsReport.UserObject = ActiveInterface.ClsVoceMenu.Classe
  Set PclsReport.Connessione = Gcon_Connect
  PclsReport.NomeServer = ActiveInterface.ClsGlobal.Gcls_GeConfig.ServerName
  PclsReport.NomeDataBase = ActiveInterface.ClsGlobal.Gcls_GeConfig.DBName
  PclsReport.NomeUtenteDb = ActiveInterface.ClsGlobal.Gcls_GeConfig.DBOwnerID
  PclsReport.NomePasswordDb = ActiveInterface.ClsGlobal.Gcls_GeConfig.DBOwnerPwd
  
  PclsReport.AddReport "Stampa chiusura casse", "GBUO_CHIUSURACASSE.rpt", App.Path, False, tsPortrait
  
  Set PclsReport.ActiveInterface = ActiveInterface
  
'  PclsReport.ReportStringaSql = strReport

  PclsReport.RhwndParent = ActiveInterface.hwndParent
  PclsReport.OpenReport

  'If printing is not good I destroy report class
  If PclsReport.Stato <> tsOK Then
     Set PclsReport = Nothing
  End If

  Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("InvocaStampa")
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
  
  

End Sub


'Private Sub PclsReport_BeforePrintReport(Cancel As Boolean, CrepReport As FWBO_REPORT30.CLSFW_OBJSTAMPA, CrepApp As Object)
'
'On Error GoTo ErrTrap
'    Dim strFiltroChiusura As String
'
'    'Report parameter(s)
''    CrepReport.SetParameterValueByName "DITTA", CStr(ActiveInterface.ClsGlobal.Gcls_DittaCorrente.RagSoc)
'
'    ' PASSA AL REPORT SOLO LA PARTE DI STRINGA RELATIVA ALLA CHIUSURA CASSA
'    'CrepReport.RecordSelectionFormula = " ( {VBR_CHISURACASSA.GB08_ID} = " & FiltroStampaChiusura & ") "
'    CrepReport.RecordSelectionFormula = FiltroStampaChiusura
'
'    ' Apre tutti i SottoReport
'    CrepReport.OpenAllSubReports
'
'    '
'    'Controllo stampa pagina iniziale
'    '
'    If PclsReport.StampaPaginaIniziale = False Then
'       CrepReport.SectionSuppress "RH", True
'       Call PreparaLimiti(CrepReport, False)
'       Call PreparaNote(CrepReport, False)
'       Exit Sub
'    End If
'
'    '
'    'Controllo se posso stampare i limiti e le note
'    '
'    If PclsReport.StampaLimiti = False And PclsReport.StampaNote = False Then
'       CrepReport.SectionSuppress "RH", True
'       Call PreparaLimiti(CrepReport, False)
'       Call PreparaNote(CrepReport, False)
'       Exit Sub
'    End If
'    '
'    'Controllo se posso stampare i limiti
'    '
'    If PclsReport.StampaLimiti = True Then
'       Call PreparaLimiti(CrepReport, True)
'    Else
'       Call PreparaLimiti(CrepReport, False)
'    End If
'    '
'    'Controllo se posso stampare le note
'    '
'    If PclsReport.StampaNote = True Then
'       Call PreparaNote(CrepReport, True)
'    Else
'       Call PreparaNote(CrepReport, False)
'    End If
'
'    Exit Sub
'ErrTrap:
'    Cancel = True
'    Select Case VisualizzaErrore("PclsReport_BeforePrintReport")
'        Case vbAbort
'            Exit Sub
'        Case vbRetry
'            Resume
'        Case vbIgnore
'            Resume Next
'    End Select
'
'End Sub




Private Sub PreparaNote(CrepReport As FWBO_REPORT30.CLSFW_OBJSTAMPA, ByVal WkTipo As Boolean)
    On Error GoTo ErrPreparaNote
    
    If WkTipo = False Then
        'Devo mettere a blank tutti i parametri
'        CrepReport.SetParameterValueByName "ESEGUITODA", ""
'        CrepReport.SetParameterValueByName "DESTINATOA", ""
'        CrepReport.SetParameterValueByName "NOTE", ""
    Else
'        CrepReport.SetParameterValueByName "ESEGUITODA", PclsReport.EseguitoDa
'        CrepReport.SetParameterValueByName "DESTINATOA", PclsReport.DestinatoA
'        CrepReport.SetParameterValueByName "NOTE", PclsReport.NoteDaStampare
    End If


    Exit Sub
ErrPreparaNote:
    Select Case VisualizzaErrore("PreparaNote")
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
End Sub

Function RecuperaDocCollegato(CodiceDocumento As String)
' On Error GoTo ErrTrap
  
  Dim MyRst       As ADODB.Recordset
  
  strSQL = "SELECT     MG36_CODDOCUMCOL " & _
        " From MG36_DOCUMENTI " & _
        " WHERE     (MG36_CODDOCUM = '" & CodiceDocumento & "')"
  Set MyRst = Gcon_Connect.Execute(strSQL, , adCmdText)
  If Not MyRst.EOF Then
      RecuperaDocCollegato = NVL(MyRst.Fields(0).value, "")
  Else
      RecuperaDocCollegato = ""
  End If
  
  strSQL = ""
 ' Set MyRst = Nothing
End Function

Private Sub PreparaLimiti(CrepReport As FWBO_REPORT30.CLSFW_OBJSTAMPA, ByVal WkTipo As Boolean)
  
  CrepReport.SetParameterValueByName "StampaSingola", StampaSingola
  
  Exit Sub
ErrPreparaLimiti:
    Select Case VisualizzaErrore("PreparaLimiti")
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
End Sub





'
' Metodo di ricalcolo del prezzo dell'articolo di magazzino
'
Public Sub RicalcolaPrezzo()
    '
    ' Variabili locali
    '
    Dim Prst_Corpo              As ADODB.Recordset
    Dim Pstr_DescrErrore        As String
    Dim Pbol_SalvaPrScVariati   As Boolean
    Dim Pvar_PercIva            As Variant
    Dim Prst_DocCorpoOrd_Clone  As ADODB.Recordset
    Dim Pvar_DataConsegna       As Variant
    Dim Pbol_ScorporaIva        As Boolean
    Dim pvarPrezzoLordo         As Variant
    Dim pvarSconto1Perc         As Variant
    Dim pvarSconto2Perc         As Variant
    Dim pvarSconto3Perc         As Variant
    Dim pvarSconto4Perc         As Variant
    Dim pvarSconto5Perc         As Variant
    Dim pvarSconto6Perc         As Variant
    Dim pvarScontoImp           As Variant
    Dim pvarMagg1Perc           As Variant
    Dim pvarMagg2Perc           As Variant
    Dim pvarMaggImp             As Variant
    Dim pvarPrezzoNetto         As Variant
    Dim pvarPrezzoIvaInclusa    As Variant
    Dim pvarLog                 As Variant
    Dim pvarIndUMPre            As Variant
    Dim pvarIndDimeRif          As Variant
    Dim pvarCodPag              As Variant
    Dim MyIndLisAcqVen          As Integer
    
    '
    ' Trap errori
    '
    On Error GoTo Err_RicalcolaPrezzo
    '
    ' Il tipo di riga deve prevedere l'articolo
    '
    '
    ' Imposto i dati del prezzo e sconti sul recordset di corpo
    MyIndLisAcqVen = CDecN(2)
    
    '
    ' Listino da priorità prezzi e sconti
    '
    If Cls_CalcPrezzi Is Nothing Then
        Set Cls_CalcPrezzi = New MGBO_PRIORPRSC.CLSMG_LEGGOPRE
        Set Cls_CalcPrezzi.ADOConnection = Gcon_Connect
        Set Cls_CalcPrezzi.ClsDittaCorrente = ActiveInterface.ClsGlobal.Gcls_DittaCorrente
        Cls_CalcPrezzi.TipoPriorita = PrioritaPrezziSconti
    End If
    Set Cls_CalcPrezzi.CallingObject = Me
    Cls_CalcPrezzi.Ditta = CodiceDitta
    Cls_CalcPrezzi.TipoPriorita = PrioritaPrezziSconti
    
    Cls_CalcPrezzi.DataElaborazione = CDate(TXT_ANNOINSERIMENTO.Text)
            
  
    Cls_CalcPrezzi.InUsoADecorrere = InUso
    Cls_CalcPrezzi.CodiceArticolo = TXT_GB07_CODART_MG66.Text
    Cls_CalcPrezzi.CodiceOpzione = ""
    Cls_CalcPrezzi.TipoCf = CDecN(0)
    
    Cls_CalcPrezzi.CodiceCF = TXT_GB06_CLIFOR_CG44.Text
   ' Cls_CalcPrezzi.Destinatario = TXT_GB06_CODDESTIN_MG22.Text
    
    
    Cls_CalcPrezzi.NumeroListinoSelezionato = TXT_MG19_LISTMAG.Text
    
    Cls_CalcPrezzi.Valuta = "EURO"
    Cls_CalcPrezzi.Cambio = 1
   'Cls_CalcPrezzi.Deposito = TXT_DEP_DA.Text
    Cls_CalcPrezzi.Quantita = TXT_GB07_QTA.Text
'            Cls_CalcPrezzi.Quantita2 = TXT_GB07_QTA.Text
'            Cls_CalcPrezzi.QuantitaCF = Prst_Corpo("DO30_QTACF").Value
            
    Cls_CalcPrezzi.TipoQuantita = TipoQuantita1
   
    Cls_CalcPrezzi.RecuperaPrezzoProvvigioni
    
    
     
    
    If Cls_CalcPrezzi.Stato <> 0 Then
      Pstr_DescrErrore = "Errore nel recupero del prezzo" & vbCr
      Select Case Cls_CalcPrezzi.Stato
        Case 1
            Pstr_DescrErrore = Pstr_DescrErrore & "VALIDAZIONE DATI"
        Case 2
            Pstr_DescrErrore = Pstr_DescrErrore & "MANCA LA TABELLA PRIORITA' PREZZI E SCONTI"
        Case 3
            Pstr_DescrErrore = Pstr_DescrErrore & "ERRORE NEI PARAMETRI DI MAGAZZINO"
        Case 4
            Pstr_DescrErrore = Pstr_DescrErrore & "ARTICOLO MANCANTE"
        Case 5
            Pstr_DescrErrore = Pstr_DescrErrore & "CLIENTE/FORNITORE MANCANTE"
        Case 6
            Pstr_DescrErrore = Pstr_DescrErrore & "ERRORE DI CONVERSIONE DA EURO A VALUTA"
        Case 7
            Pstr_DescrErrore = Pstr_DescrErrore & "ERRORE NEL CALCOLO PREZZO BASE/VARIANTE/MOLTIPLICATORE"
        Case 8
            Pstr_DescrErrore = Pstr_DescrErrore & "ERRORE NEL CALCOLO DEL PREZZO NETTO"
        Case 99
            Pstr_DescrErrore = Pstr_DescrErrore & "ERRORE VB"
        Case Else
            Pstr_DescrErrore = Pstr_DescrErrore & "ERRORE GENERICO"
      End Select
      MsgBox Pstr_DescrErrore
    Else
      TXT_GB07_PREZZO.Text = Cls_CalcPrezzi.PrezzoLordo
      'MODIFICA PER GESTIRE SCONTO AGGIUNTIVO SULLA CONDIZIONE PAGAMENTO
      If Cls_CalcPrezzi.Sconto1 > 0 Then
      Select Case Trim(TXT_GB06_CODPAG_CG62.Text)
      Case "C00", "C01", "C02", "C03", "C04"
        
        
        If Trim(GetFamArt(Cls_CalcPrezzi.CodiceArticolo)) = "CSM" Then
            'TXT_SCCONDPAG.Text = 5
           ' TXT_SCLIST.Text = Cls_CalcPrezzi.Sconto1
           ' TXT_GB07_SCCORPO.Text = Cls_CalcPrezzi.Sconto1 ' + TXT_SCCONDPAG.Text
        Else
           ' TXT_GB07_SCCORPO.Text = Cls_CalcPrezzi.Sconto1
            'TXT_SCLIST.Text = Cls_CalcPrezzi.Sconto1
           ' TXT_SCCONDPAG.Text = 0
        End If
      Case Else
       ' TXT_GB07_SCCORPO.Text = Cls_CalcPrezzi.Sconto1
       '' TXT_SCLIST.Text = Cls_CalcPrezzi.Sconto1
       ' TXT_SCCONDPAG.Text = 0
      End Select
      Else
     ' TXT_GB07_SCCORPO.Text = Cls_CalcPrezzi.Sconto1
     ''' TXT_SCLIST.Text = Cls_CalcPrezzi.Sconto1
     ' TXT_SCCONDPAG.Text = 0
      End If
    End If
    Set Cls_CalcPrezzi.CallingObject = Nothing
    '
    ' Rilascio il riferimento al recordset
    '
    Set Prst_Corpo = Nothing
    '
    ' Esco
    '
    Exit Sub
Err_RicalcolaPrezzo:
    Select Case ShowError("RicalcolaPrezzo", 99)
        Case vbIgnore
            Resume Next
        Case vbRetry
            Resume
    End Select
End Sub


'
' Metodo di ricalcolo del prezzo dell'articolo di magazzino
'
Public Sub RecuperaCosto()
    '
    ' Variabili locali
    '
    Dim Prst_Corpo              As ADODB.Recordset
    Dim Pstr_DescrErrore        As String
    Dim Pbol_SalvaPrScVariati   As Boolean
    Dim Pvar_PercIva            As Variant
    Dim Prst_DocCorpoOrd_Clone  As ADODB.Recordset
    Dim Pvar_DataConsegna       As Variant
    Dim Pbol_ScorporaIva        As Boolean
    Dim pvarPrezzoLordo         As Variant
    Dim pvarSconto1Perc         As Variant
    Dim pvarSconto2Perc         As Variant
    Dim pvarSconto3Perc         As Variant
    Dim pvarSconto4Perc         As Variant
    Dim pvarSconto5Perc         As Variant
    Dim pvarSconto6Perc         As Variant
    Dim pvarScontoImp           As Variant
    Dim pvarMagg1Perc           As Variant
    Dim pvarMagg2Perc           As Variant
    Dim pvarMaggImp             As Variant
    Dim pvarPrezzoNetto         As Variant
    Dim pvarPrezzoIvaInclusa    As Variant
    Dim pvarLog                 As Variant
    Dim pvarIndUMPre            As Variant
    Dim pvarIndDimeRif          As Variant
    Dim pvarCodPag              As Variant
    Dim MyIndLisAcqVen          As Integer
    
    '
    ' Trap errori
    '
    On Error GoTo Err_RicalcolaPrezzo
    '
    ' Il tipo di riga deve prevedere l'articolo
    '
    '
    ' Imposto i dati del prezzo e sconti sul recordset di corpo
    MyIndLisAcqVen = 1
    
    '
    ' Listino da priorità prezzi e sconti
    '
    If Cls_CalcPrezzi Is Nothing Then
        Set Cls_CalcPrezzi = New MGBO_PRIORPRSC.CLSMG_LEGGOPRE
        Set Cls_CalcPrezzi.ADOConnection = Gcon_Connect
        Set Cls_CalcPrezzi.ClsDittaCorrente = ActiveInterface.ClsGlobal.Gcls_DittaCorrente
        Cls_CalcPrezzi.TipoPriorita = PrioritaPrezziSconti
    End If
    Set Cls_CalcPrezzi.CallingObject = Me
    Cls_CalcPrezzi.Ditta = CodiceDitta
    Cls_CalcPrezzi.TipoPriorita = PrioritaPrezziSconti
   ' Cls_CalcPrezzi.TipoCf = Fornitore
    Cls_CalcPrezzi.DataElaborazione = CDate(TXT_ANNOINSERIMENTO.Text)
   ' Cls_CalcPrezzi.TipoListinoElaborato = ArticoliClientiFornitori
  
    Cls_CalcPrezzi.InUsoADecorrere = InUso
    Cls_CalcPrezzi.CodiceArticolo = TXT_GB07_CODART_MG66.Text
    Cls_CalcPrezzi.CodiceOpzione = ""
    Cls_CalcPrezzi.TipoCf = 1
    
    If NVL(TXT_GB07_CLIFOR_CG44.Text, "") = "" Then

        Cls_CalcPrezzi.CodiceCF = 2

    Else
   
        Cls_CalcPrezzi.CodiceCF = TXT_GB07_CLIFOR_CG44.Text
    
    End If
    
    Cls_CalcPrezzi.NumeroListinoSelezionato = 1
    
    Cls_CalcPrezzi.Valuta = "EURO"
    Cls_CalcPrezzi.Cambio = 1
    Cls_CalcPrezzi.Quantita = TXT_GB07_QTA.Text
            
    Cls_CalcPrezzi.TipoQuantita = TipoQuantita1
   
    Cls_CalcPrezzi.RecuperaPrezzoProvvigioni
    
    
     
    
    If Cls_CalcPrezzi.Stato <> 0 Then
      Pstr_DescrErrore = "Errore nel recupero del prezzo" & vbCr
      Select Case Cls_CalcPrezzi.Stato
        Case 1
            Pstr_DescrErrore = Pstr_DescrErrore & "VALIDAZIONE DATI"
        Case 2
            Pstr_DescrErrore = Pstr_DescrErrore & "MANCA LA TABELLA PRIORITA' PREZZI E SCONTI"
        Case 3
            Pstr_DescrErrore = Pstr_DescrErrore & "ERRORE NEI PARAMETRI DI MAGAZZINO"
        Case 4
            Pstr_DescrErrore = Pstr_DescrErrore & "ARTICOLO MANCANTE"
        Case 5
            Pstr_DescrErrore = Pstr_DescrErrore & "CLIENTE/FORNITORE MANCANTE"
        Case 6
            Pstr_DescrErrore = Pstr_DescrErrore & "ERRORE DI CONVERSIONE DA EURO A VALUTA"
        Case 7
            Pstr_DescrErrore = Pstr_DescrErrore & "ERRORE NEL CALCOLO PREZZO BASE/VARIANTE/MOLTIPLICATORE"
        Case 8
            Pstr_DescrErrore = Pstr_DescrErrore & "ERRORE NEL CALCOLO DEL PREZZO NETTO"
        Case 99
            Pstr_DescrErrore = Pstr_DescrErrore & "ERRORE VB"
        Case Else
            Pstr_DescrErrore = Pstr_DescrErrore & "ERRORE GENERICO"
      End Select
      MsgBox Pstr_DescrErrore
    Else
      TXT_GB07_COSTO.Text = Cls_CalcPrezzi.PrezzoNetto
      'MODIFICA PER GESTIRE SCONTO AGGIUNTIVO SULLA CONDIZIONE PAGAMENTO
      If Cls_CalcPrezzi.Sconto1 > 0 Then
      Select Case Trim(TXT_GB06_CODPAG_CG62.Text)
      Case "C00", "C01", "C02", "C03", "C04"
        
        
        If Trim(GetFamArt(Cls_CalcPrezzi.CodiceArticolo)) = "CSM" Then
            'TXT_SCCONDPAG.Text = 5
           ' TXT_SCLIST.Text = Cls_CalcPrezzi.Sconto1
           ' TXT_GB07_SCCORPO.Text = Cls_CalcPrezzi.Sconto1 ' + TXT_SCCONDPAG.Text
        Else
           ' TXT_GB07_SCCORPO.Text = Cls_CalcPrezzi.Sconto1
            'TXT_SCLIST.Text = Cls_CalcPrezzi.Sconto1
           ' TXT_SCCONDPAG.Text = 0
        End If
      Case Else
       ' TXT_GB07_SCCORPO.Text = Cls_CalcPrezzi.Sconto1
       '' TXT_SCLIST.Text = Cls_CalcPrezzi.Sconto1
       ' TXT_SCCONDPAG.Text = 0
      End Select
      Else
     ' TXT_GB07_SCCORPO.Text = Cls_CalcPrezzi.Sconto1
     ''' TXT_SCLIST.Text = Cls_CalcPrezzi.Sconto1
     ' TXT_SCCONDPAG.Text = 0
      End If
    End If
    Set Cls_CalcPrezzi.CallingObject = Nothing
    '
    ' Rilascio il riferimento al recordset
    '
    Set Prst_Corpo = Nothing
    '
    ' Esco
    '
    Exit Sub
Err_RicalcolaPrezzo:
    Select Case ShowError("RicalcolaPrezzo", 99)
        Case vbIgnore
            Resume Next
        Case vbRetry
            Resume
    End Select
End Sub

Public Sub RicalcolaImporto()
    On Error GoTo Err_RicalcolaImporto
      
    Dim PrezzoSc1  As Double
    Dim PrezzoSc2  As Double
    Dim PrezzoSc3  As Double
    Dim PrezzoSc4  As Double
    Dim PrezzoSc5  As Double
    Dim PrezzoSc6  As Double
    
'    If NVL(TXT_SCONTOPIEDE.Text, 0) <> 0 Then
'      TXT_GB07_SCPIEDE.Text = TXT_SCONTOPIEDE.Text
'    End If
    
    'Calcolo importo riga
    PrezzoSc1 = TXT_GB07_PREZZO.Text - TXT_GB07_PREZZO.Text * TXT_GB07_SC1.Text / 100
    PrezzoSc2 = PrezzoSc1 - PrezzoSc1 * TXT_GB07_SC2.Text / 100
    PrezzoSc3 = PrezzoSc2 - PrezzoSc2 * TXT_GB07_SC3.Text / 100
    PrezzoSc4 = PrezzoSc3 - PrezzoSc3 * TXT_GB07_SC4.Text / 100
    PrezzoSc5 = PrezzoSc4 - PrezzoSc4 * TXT_GB07_SC5.Text / 100
    PrezzoSc6 = PrezzoSc5 - PrezzoSc5 * TXT_GB07_SC6.Text / 100
    
    TXT_GB07_IMPORTO.Text = Round(PrezzoSc6 * TXT_GB07_QTA.Text, 2) + CInt(0) / 100
    TXT_GB07_PERCPROVV.Text = RecuperaProvvigione(TXT_GB07_PREZZO.Text, TXT_GB07_IMPORTO.Text, RecuperaFamigliaArticolo(NVL(TXT_GB07_CODART_MG66.Text, "")), RecuperaSFamigliaArticolo(NVL(TXT_GB07_CODART_MG66.Text, "")), TXT_GB06_AGENTE.Text, NVL(TXT_GB07_SC1.Text, 0), NVL(TXT_GB07_SC2.Text, 0), NVL(TXT_GB07_SC3.Text, 0), NVL(TXT_GB07_SC4.Text, 0), 0, 0)

    'Salvo virtualframe
    FME_BANCO.Update True, False, False
    Call CMD_REFRESH_Click
    Exit Sub
Err_RicalcolaImporto:
    Select Case ShowError("RicalcolaImporto", 99)
        Case vbIgnore
            Resume Next
        Case vbRetry
            Resume
    End Select
End Sub

Public Function RecuperaFamigliaArticolo(CodArt As String) As String
Dim MyRst       As ADODB.Recordset
Dim strSQL As String
  strSQL = " SELECT" & _
              "    MG66_FAM_MG53 " & _
              " FROM" & _
              "    MG66_ANAGRART WITH (NOLOCK)" & _
              " WHERE" & _
              "    MG66_DITTA_CG18 = " & CodiceDitta & " AND" & _
              "    MG66_CODART = '" & Replace(CodArt, "'", "''") & "'"
  Set MyRst = Gcon_Connect.Execute(strSQL, , adCmdText)
  If Not MyRst.EOF Then
      RecuperaFamigliaArticolo = NVL(MyRst.Fields(0).value, "")
  Else
      RecuperaFamigliaArticolo = ""
  End If

End Function

Public Function RecuperaSFamigliaArticolo(CodArt As String) As String
Dim MyRst       As ADODB.Recordset
Dim strSQL As String
  strSQL = " SELECT" & _
              "    MG66_SFAM_MG54 " & _
              " FROM" & _
              "    MG66_ANAGRART WITH (NOLOCK)" & _
              " WHERE" & _
              "    MG66_DITTA_CG18 = " & CodiceDitta & " AND" & _
              "    MG66_CODART = '" & Replace(CodArt, "'", "''") & "'"
  Set MyRst = Gcon_Connect.Execute(strSQL, , adCmdText)
  If Not MyRst.EOF Then
      RecuperaSFamigliaArticolo = NVL(MyRst.Fields(0).value, "")
  Else
      RecuperaSFamigliaArticolo = ""
  End If

End Function


Public Sub RicalcolaImportoTotale()
    On Error GoTo Err_RicalcolaImportoTotale
    Dim ris As Double
    Dim StringaSQL As String
    Dim rst        As ADODB.Recordset
    Dim ImportoTot As Double
    
    'Ricalcolo totale documento
    StringaSQL = " SELECT SUM(GB07_IMPORTO) AS TOTALE "
    StringaSQL = StringaSQL & "  FROM GB07_CORPODOC "
    StringaSQL = StringaSQL & " WHERE GB07_ID_GB06 = " & NVL(IDGB06, 0)
    
    Set rst = ActiveInterface.Connection.Execute(StringaSQL)
    If rst.EOF And rst.BOF Then
      ImportoTot = 0
    Else
      ImportoTot = NVL(rst("TOTALE").value, 0)
    End If
    
    rst.Close
    Set rst = Nothing
    
    ' TXT_TOTALEDOC.Text = FormatNumber(ImportoTot, 2)
    TXT_GB06_CONSUNTIVO.Text = Format(FormatNumber(ImportoTot, 1), "0.00")
    TXT_GB06_FORECAST.Text = NVL(TXT_GB06_BUDGET.Text, 0) - NVL(TXT_GB06_CONSUNTIVO.Text, 0)
    'recuperare iva
    If TipoDocumento = 1 Then
  '  TXT_TOTALEDOCIVA.Text = FormatNumber(ImportoTot + (ImportoTot * 22 / 100), 2)
    Else
  '  TXT_TOTALEDOCIVA.Text = FormatNumber(ImportoTot, 2)
    End If
    'percentuale ribasso gara
    If NVL(TXT_GB06_BUDGET.Text, 0) > 0 Then
    ris = 0
    ris = CDbl(NVL(TXT_GB06_BUDGET.Text, 0)) - CDbl(NVL(TXT_GB06_CONSUNTIVO.Text, 0))
    ris = ris / CDbl(NVL(TXT_GB06_BUDGET.Text, 0))
    ris = ris * 100
    TXT_GB06_PERCRIBGARA.Text = CStr(ris)
    Else
    TXT_GB06_PERCRIBGARA.Text = 0
    End If
    
    Exit Sub
Err_RicalcolaImportoTotale:
    Select Case ShowError("RicalcolaImportoTotale", 99)
        Case vbIgnore
            Resume Next
        Case vbRetry
            Resume
    End Select
End Sub

Public Sub RicalcolaImportoTotaleSimulazione(TABLENAME As String)
    On Error GoTo Err_RicalcolaImportoTotaleSim
    
    Dim StringaSQL As String
    Dim rst        As ADODB.Recordset
    Dim ImportoTot As Double
    
    'Ricalcolo totale documento
    StringaSQL = " SELECT SUM(GB07_IMPORTO_NEW) AS TOTALE "
    StringaSQL = StringaSQL & "  FROM " & TABLENAME & ""
    StringaSQL = StringaSQL & " WHERE GB07_ID_GB06 = " & NVL(IDGB06, 0)
    
    Set rst = ActiveInterface.Connection.Execute(StringaSQL)
    If rst.EOF And rst.BOF Then
      ImportoTot = 0
    Else
      ImportoTot = NVL(rst("TOTALE").value, 0)
    End If
    
    rst.Close
    Set rst = Nothing
    
   ' TXT_TOTALEDOC.Text = FormatNumber(ImportoTot, 2)
   TXT_GB06_CONSUNTIVO.Text = Format(FormatNumber(ImportoTot, 1), "0.00")
   TXT_GB06_FORECAST.Text = NVL(TXT_GB06_BUDGET.Text, 0) - NVL(TXT_GB06_CONSUNTIVO.Text, 0)
'    'recuperare iva
'    If TipoDocumento = 1 Then
'    TXT_TOTALEDOCIVA.Text = FormatNumber(ImportoTot + (ImportoTot * 22 / 100), 2)
'    Else
'    TXT_TOTALEDOCIVA.Text = FormatNumber(ImportoTot, 2)
'    End If
    Exit Sub
Err_RicalcolaImportoTotaleSim:
    Select Case ShowError("RicalcolaImportoTotale", 99)
        Case vbIgnore
            Resume Next
        Case vbRetry
            Resume
    End Select
End Sub

Private Sub TXT_SFAM_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
'On Error Resume Next
'    Call Cls_DecodeMagazzino.SottoFamiglia(TXT_FAM.Text, TXT_SFAM.Text)
'    str_SQL = Cls_DecodeMagazzino.StringaSQL
'    Arr_Fields = Cls_DecodeMagazzino.ArrayFields
'
'    Set Arr_Fields(0, 0) = TXT_SFMA_DEC
'
'    Str_Connect = Gstr_Connect
If NVL(TXT_SFAM.Text, "") = "" Then
      Exit Sub
  End If

  Cancel = False

  ReDim Arr_Fields(0 To 0, 0 To 1)
  Arr_Fields(0, 1) = "MG54_DESCRSFAM"
  Set Arr_Fields(0, 0) = TXT_SFMA_DEC
  
  'Imposto la stringa SQL
  '
  
  str_SQL = " SELECT  MG54_DESCRSFAM "
  str_SQL = str_SQL & " FROM MG54_SOTTOFAM "
  str_SQL = str_SQL & " WHERE MG54_DITTA_CG18  = " & CodiceDitta
  str_SQL = str_SQL & "   AND MG54_CODSFAM = '" & TXT_SFAM.Text & "'"
  
  
  
 
  Str_Connect = Gstr_Connect
  Err.Clear

End Sub

Private Sub TXT_SFAM_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
'  On Error Resume Next
'
'  Call Cls_LookupMagazzino.SottoFamiglie(TXT_FAM.Text)
'  str_SQL = Cls_LookupMagazzino.StringaSQL
'  Arr_Fields = Cls_LookupMagazzino.ArrayFields
'  Str_Caption = Cls_LookupMagazzino.Titolo
'  Str_Connect = Gstr_Connect
'  TXT_SFAM.IDLookup = Cls_LookupMagazzino.IDLookup
'  On Error Resume Next

 On Error Resume Next

  Cancel = False
  str_SQL = " SELECT  distinct   MG54_CODSFAM, MG54_DESCRSFAM ,  MG54_CODFAM_MG53"
  str_SQL = str_SQL & " FROM MG54_SOTTOFAM "
  str_SQL = str_SQL & " WHERE MG54_DITTA_CG18  = " & CodiceDitta
  
  
  ReDim Arr_Fields(0 To 1, 0 To 1)
  Arr_Fields(0, 0) = "Codice"
  Arr_Fields(0, 1) = ""
  Arr_Fields(1, 0) = "Descrizione"
  Arr_Fields(1, 1) = ""


  Str_Caption = "Sotto Famiglie"
  Str_Connect = Gstr_Connect
  TXT_SFAM = "lkp_SFamiglie"
  
  Err.Clear

End Sub

Private Sub TXT_SGRUPPO_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
'On Error Resume Next
'    Call Cls_DecodeMagazzino.SottoGruppo(TXT_FAM.Text, TXT_SFAM.Text, TXT_GRUPPO.Text, TXT_SGRUPPO.Text)
'    str_SQL = Cls_DecodeMagazzino.StringaSQL
'    Arr_Fields = Cls_DecodeMagazzino.ArrayFields
'
'    Set Arr_Fields(0, 0) = TXT_SGRUPPO_DEC
'
'    Str_Connect = Gstr_Connect
If NVL(TXT_SGRUPPO.Text, "") = "" Then
      Exit Sub
  End If

  Cancel = False

  ReDim Arr_Fields(0 To 0, 0 To 1)
  Arr_Fields(0, 1) = "MG56_DESCRSGRUPPO"
  Set Arr_Fields(0, 0) = TXT_SGRUPPO_DEC
  
  'Imposto la stringa SQL
  '
  
  str_SQL = " SELECT  MG56_DESCRSGRUPPO "
  str_SQL = str_SQL & " FROM MG56_SOTTOGRUPPI "
  str_SQL = str_SQL & " WHERE MG56_DITTA_CG18  = " & CodiceDitta
  str_SQL = str_SQL & "   AND MG56_CODSGRUPPO = '" & TXT_SGRUPPO.Text & "'"
  
  
  
 
  Str_Connect = Gstr_Connect
  Err.Clear
End Sub

Private Sub TXT_SGRUPPO_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
'  On Error Resume Next
'
'  Call Cls_LookupMagazzino.SottoGruppi(TXT_FAM.Text, TXT_SFAM.Text, TXT_GRUPPO.Text)
'  str_SQL = Cls_LookupMagazzino.StringaSQL
'  Arr_Fields = Cls_LookupMagazzino.ArrayFields
'  Str_Caption = Cls_LookupMagazzino.Titolo
'  Str_Connect = Gstr_Connect
'  TXT_SGRUPPO.IDLookup = Cls_LookupMagazzino.IDLookup

 On Error Resume Next

  Cancel = False
  str_SQL = " SELECT  distinct   MG56_CODSGRUPPO, MG56_DESCRSGRUPPO "
  str_SQL = str_SQL & " FROM MG56_SOTTOGRUPPI "
  str_SQL = str_SQL & " WHERE MG56_DITTA_CG18  = " & CodiceDitta
  
  
  ReDim Arr_Fields(0 To 1, 0 To 1)
  Arr_Fields(0, 0) = "Codice"
  Arr_Fields(0, 1) = ""
  Arr_Fields(1, 0) = "Descrizione"
  Arr_Fields(1, 1) = ""

  Str_Caption = "Sotto Gruppi"
  Str_Connect = Gstr_Connect
  TXT_SGRUPPO = "lkp_SottoGruppi"
  
  Err.Clear
End Sub
