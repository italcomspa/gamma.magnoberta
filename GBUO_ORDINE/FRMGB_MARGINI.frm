VERSION 5.00
Object = "{F53BE214-7AC6-11D0-9B0E-006097A80EFD}#6.12#0"; "TMS_LABEL.ocx"
Object = "{0EF4E9DB-2617-11D2-A1C0-0060082875F9}#10.13#0"; "TMS_EDITNUM.ocx"
Object = "{9AE03505-25F7-11D2-A1C0-0060082875F9}#7.3#0"; "TMS_FRAME.ocx"
Begin VB.Form FRMGB_MARGINI 
   Caption         =   "Anali Margini"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin PRJFW_FRAME.TMS_FRAME TMS_FRAME2 
      Height          =   5505
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   9710
      Begin PRJFW_EDITNUM.TxtEditNum TXT_C_LISTINOG 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   3450
         TabIndex        =   57
         Top             =   390
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_C_LISTINOR 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   3450
         TabIndex        =   56
         Top             =   720
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_P_LISTINOG 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   4890
         TabIndex        =   55
         Top             =   390
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_P_LISTINOR 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   4890
         TabIndex        =   54
         Top             =   720
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_C_GIOCHI 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   3450
         TabIndex        =   53
         Top             =   1050
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_P_GIOCHI 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   4890
         TabIndex        =   52
         Top             =   1050
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_C_RICAMBI 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   3450
         TabIndex        =   51
         Top             =   1710
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_P_RICAMBI 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   4890
         TabIndex        =   50
         Top             =   1710
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_C_NLOCAL 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   3450
         TabIndex        =   49
         Top             =   2040
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_C_NGROUP 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   3450
         TabIndex        =   48
         Top             =   2370
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_P_NLOCAL 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   4890
         TabIndex        =   47
         Top             =   2040
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_P_NGROUP 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   4890
         TabIndex        =   46
         Top             =   2370
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_C_ANTITRAUMA 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   3450
         TabIndex        =   45
         Top             =   2700
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_P_ANTITRAUMA 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   4890
         TabIndex        =   44
         Top             =   2700
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_C_POSALAVORI 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   3450
         TabIndex        =   43
         Top             =   3030
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_P_POSALAVORI 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   4890
         TabIndex        =   42
         Top             =   3030
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_C_TPROLUDIC 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   3450
         TabIndex        =   41
         Top             =   3360
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_P_TPROLUDIC 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   4890
         TabIndex        =   40
         Top             =   3360
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_C_TNEGOCE 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   3450
         TabIndex        =   39
         Top             =   3690
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_C_TANTI 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   3450
         TabIndex        =   38
         Top             =   4020
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_P_TNEGOCE 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   4890
         TabIndex        =   37
         Top             =   3690
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_P_TANTI 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   4890
         TabIndex        =   36
         Top             =   4020
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_O_LISTINOG 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   2010
         TabIndex        =   35
         Top             =   390
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_O_LISTINOR 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   2010
         TabIndex        =   34
         Top             =   720
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_O_GIOCHI 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   2010
         TabIndex        =   33
         Top             =   1050
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_O_RICAMBI 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   2010
         TabIndex        =   32
         Top             =   1710
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_O_NLOCAL 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   2010
         TabIndex        =   31
         Top             =   2040
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_O_NGROUP 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   2010
         TabIndex        =   30
         Top             =   2370
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_O_ANTITRAUMA 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   2010
         TabIndex        =   29
         Top             =   2700
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_O_POSALAVORI 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   2010
         TabIndex        =   28
         Top             =   3030
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_O_TPROLUDIC 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   2010
         TabIndex        =   27
         Top             =   3360
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_O_TNEGOCE 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   2010
         TabIndex        =   26
         Top             =   3690
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_C_PROVV 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   3450
         TabIndex        =   25
         Top             =   4350
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_O_TOTALE 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   2010
         TabIndex        =   24
         Top             =   4740
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_C_TOTALE 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   3450
         TabIndex        =   23
         Top             =   4740
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_MARGINE 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   3450
         TabIndex        =   22
         Top             =   5070
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_MARGINE_PER 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   4890
         TabIndex        =   21
         Top             =   5070
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
         Height          =   300
         Index           =   50
         Left            =   420
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   420
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         Caption         =   "Listino Giochi"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
         Height          =   300
         Index           =   51
         Left            =   420
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         Caption         =   "Trasp. Negoce"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
         Height          =   300
         Index           =   52
         Left            =   420
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3390
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         Caption         =   "Trasporto Proludic"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
         Height          =   300
         Index           =   53
         Left            =   420
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3060
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         Caption         =   "Posa e Lavori"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
         Height          =   300
         Index           =   54
         Left            =   420
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         Caption         =   "Antitrauma"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
         Height          =   300
         Index           =   55
         Left            =   420
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         Caption         =   "Negoce Gruop"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
         Height          =   300
         Index           =   56
         Left            =   420
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2070
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         Caption         =   "Negoce Local"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
         Height          =   300
         Index           =   57
         Left            =   420
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1740
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         Caption         =   "Ricambi"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
         Height          =   300
         Index           =   58
         Left            =   420
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         Caption         =   "Giochi"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
         Height          =   300
         Index           =   59
         Left            =   420
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   750
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         Caption         =   "Listino Ricambi"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
         Height          =   300
         Index           =   62
         Left            =   420
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   4380
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         Caption         =   "Provvigioni"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TXT_O_TANTITRAUMA 
         Height          =   300
         Index           =   64
         Left            =   420
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   4050
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         Caption         =   "Trasp. Antitrauma"
      End
      Begin VB.Label Label3 
         Caption         =   "Margine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   0
         Left            =   420
         TabIndex        =   8
         Top             =   5070
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Totale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   390
         TabIndex        =   7
         Top             =   4770
         Width           =   825
      End
      Begin VB.Shape Shape3 
         Height          =   735
         Left            =   360
         Top             =   4680
         Width           =   5895
      End
      Begin VB.Label Label3 
         Caption         =   "OFFERTA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   6
         Top             =   60
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "COSTI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   3
         Left            =   3630
         TabIndex        =   5
         Top             =   60
         Width           =   825
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_O_TANTI 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   2010
         TabIndex        =   4
         Top             =   4020
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL16 
         Height          =   300
         Index           =   60
         Left            =   420
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1410
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         Caption         =   "Promo"
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_O_PROMO 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   2010
         TabIndex        =   2
         Top             =   1380
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_P_PROMO 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   4890
         TabIndex        =   1
         Top             =   1380
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_C_PROMO 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   2
         EndProperty
         Height          =   300
         Left            =   3450
         TabIndex        =   0
         Top             =   1380
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Enabled         =   0   'False
         Obbligatorio    =   -1  'True
         DBField         =   "GB00_IMPMAXICANONE"
         Caption         =   "Edit numerico"
         Object.Tag             =   "Edit numerico"
         MaxWidth        =   8
         FormatMask      =   """€"" ###,###,###,##0.00"
         CanRequired     =   0   'False
         TipoFormato     =   2
      End
   End
End
Attribute VB_Name = "FRMGB_MARGINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ActiveInterface               As Cinterface

Public ActiveClass                   As CLSGB_ORDINE
Private Gstr_Connect                 As String
Private Gcon_Connect                 As ADODB.Connection
Private Gcls_Connect                 As CLSFW_SetConnect

Private Gcls_Log                    As CLSFW_SrvLog

Private CodiceDitta                 As Variant

Private FormIsActive                As Boolean

Public IdOfferta As Double
Public ValPercTrasp As Double

Private Sub BTN_AGGMARGINI_ButtonClick()

End Sub

Private Sub Form_Activate()
    On Error GoTo ErrTrap

    

    If FormIsActive Then
        Exit Sub
    End If

    'Apertura connessione
    Gstr_Connect = ActiveInterface.ClsGlobal.Gcls_LibConnect.GetExtendedProperties
    Set Gcls_Connect = New CLSFW_SetConnect
    Set Gcon_Connect = Gcls_Connect.Gpr_GetConnect
    With Gcon_Connect
        .ConnectionString = Gstr_Connect
        .Open
    End With

   Set ActiveInterface.Connection = Gcon_Connect

    CodiceDitta = ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta

   Call AnalisiTuttiMargini
    
    FormIsActive = True

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

'Private Sub InizializzaScriptELayout()
'    On Error GoTo ErrTrap
'    'layout and script object initialization
'    ActiveInterface.ActiveNavigator.ApplyPrsLayout
'    ExecuteFormEvent "tsOpen"
'    Set ActiveInterface.ClsGlobal.ApplicationObject = App
'
'Exit Sub
'ErrTrap:
'    Select Case VisualizzaErrore("InizializzaScriptELayout")
'        Case vbAbort
'            Exit Sub
'        Case vbRetry
'            Resume
'        Case vbIgnore
'            Resume Next
'    End Select
'End Sub

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
Private Sub AnalisiTuttiMargini()
Dim CostiTotali As Double
Dim TotaleOfferta As Double
CostiTotali = 0
TotaleOfferta = 0

'Offerte
Call Margini_ListinoGiochi
Call Margini_ListinoRicambi
Call Margini_Giochi
Call Margini_Promo
Call Margini_Ricambi
Call Margini_NegoceLocal
Call Margini_NegoceGruop
Call Margini_Antitrauma
Call Margini_PosaLavori
Call Margini_TrasportoProludic
Call Margini_TrasportoNegoce
Call Margini_TrasportoAntitrauma


'TotaleOfferta = TotaleOfferta + NVL(TXT_O_LISTINOG.Text, 0)
TotaleOfferta = TotaleOfferta + NVL(TXT_O_GIOCHI.Text, 0)
TotaleOfferta = TotaleOfferta + NVL(TXT_O_ANTITRAUMA.Text, 0)
TotaleOfferta = TotaleOfferta + NVL(TXT_O_LISTINOR.Text, 0)
TotaleOfferta = TotaleOfferta + NVL(TXT_O_NGROUP.Text, 0)
TotaleOfferta = TotaleOfferta + NVL(TXT_O_NLOCAL.Text, 0)
TotaleOfferta = TotaleOfferta + NVL(TXT_O_POSALAVORI.Text, 0)
TotaleOfferta = TotaleOfferta + NVL(TXT_O_RICAMBI.Text, 0)
TotaleOfferta = TotaleOfferta + NVL(TXT_O_TANTI.Text, 0)

TotaleOfferta = TotaleOfferta + NVL(TXT_O_TNEGOCE.Text, 0)
TotaleOfferta = TotaleOfferta + NVL(TXT_O_PROMO.Text, 0)
TotaleOfferta = TotaleOfferta + NVL(TXT_O_TPROLUDIC.Text, 0)

TXT_O_TOTALE.Text = TotaleOfferta


'costi
Call Costi_Provvigioni
Call Costi_ListinoGiochi
Call Costi_ListinoRicambi
Call Costi_Giochi
Call Costi_Promo
Call Costi_Ricambi
Call Costi_NegoceLocal
Call Costi_NegoceGruop
Call Costi_Antitrauma
Call Costi_PosaLavori
Call Costi_TrasportoProludic
Call Costi_TrasportoNegoce
Call Costi_TrasportoAntitrauma

'CostiTotali = CostiTotali + NVL(TXT_C_LISTINOG.Text, 0)
CostiTotali = CostiTotali + NVL(TXT_C_GIOCHI.Text, 0)
CostiTotali = CostiTotali + NVL(TXT_C_ANTITRAUMA.Text, 0)
CostiTotali = CostiTotali + NVL(TXT_C_LISTINOR.Text, 0)
CostiTotali = CostiTotali + NVL(TXT_C_NGROUP.Text, 0)
CostiTotali = CostiTotali + NVL(TXT_C_NLOCAL.Text, 0)
CostiTotali = CostiTotali + NVL(TXT_C_POSALAVORI.Text, 0)
CostiTotali = CostiTotali + NVL(TXT_C_RICAMBI.Text, 0)
CostiTotali = CostiTotali + NVL(TXT_C_TANTI.Text, 0)

CostiTotali = CostiTotali + NVL(TXT_C_TNEGOCE.Text, 0)
CostiTotali = CostiTotali + NVL(TXT_C_PROMO.Text, 0)
CostiTotali = CostiTotali + NVL(TXT_C_TPROLUDIC.Text, 0)
CostiTotali = CostiTotali + NVL(TXT_C_PROVV.Text, 0)

TXT_C_TOTALE.Text = CostiTotali

If NVL(TXT_O_LISTINOG.Text, 0) > 0 Then
TXT_P_LISTINOG.Text = Round((TXT_O_LISTINOG.Text - TXT_C_LISTINOG.Text) / TXT_O_LISTINOG.Text * 100, 2)
End If
If NVL(TXT_O_GIOCHI.Text, 0) > 0 Then
TXT_P_GIOCHI.Text = Round((TXT_O_GIOCHI.Text - TXT_C_GIOCHI.Text) / TXT_O_GIOCHI.Text * 100, 2)
End If
If NVL(TXT_O_ANTITRAUMA.Text, 0) > 0 Then
TXT_P_ANTITRAUMA.Text = Round((TXT_O_ANTITRAUMA.Text - TXT_C_ANTITRAUMA.Text) / TXT_O_ANTITRAUMA.Text * 100, 2)
End If
If NVL(TXT_O_LISTINOR.Text, 0) > 0 Then
TXT_P_LISTINOR.Text = Round((TXT_O_LISTINOR.Text - TXT_C_LISTINOR.Text) / TXT_O_LISTINOR.Text * 100, 2)
End If
If NVL(TXT_O_NGROUP.Text, 0) > 0 Then
TXT_P_NGROUP.Text = Round((TXT_O_NGROUP.Text - TXT_C_NGROUP.Text) / TXT_O_NGROUP.Text * 100, 2)
End If
If NVL(TXT_O_NLOCAL.Text, 0) > 0 Then
TXT_P_NLOCAL.Text = Round((TXT_O_NLOCAL.Text - TXT_C_NLOCAL.Text) / TXT_O_NLOCAL.Text * 100, 2)
End If
If NVL(TXT_O_POSALAVORI.Text, 0) > 0 Then
TXT_P_POSALAVORI.Text = Round((TXT_O_POSALAVORI.Text - TXT_C_POSALAVORI.Text) / TXT_O_POSALAVORI.Text * 100, 2)
End If
If NVL(TXT_O_RICAMBI.Text, 0) > 0 Then
TXT_P_RICAMBI.Text = Round((TXT_O_RICAMBI.Text - TXT_C_RICAMBI.Text) / TXT_O_RICAMBI.Text * 100, 2)
End If
If NVL(TXT_O_TANTI.Text, 0) > 0 Then
TXT_P_TANTI.Text = Round((TXT_O_TANTI.Text - TXT_C_TANTI.Text) / TXT_O_TANTI.Text * 100, 2)
End If
If NVL(TXT_O_TNEGOCE.Text, 0) > 0 Then
TXT_P_TNEGOCE.Text = Round((TXT_O_TNEGOCE.Text - TXT_C_TNEGOCE.Text) / TXT_O_TNEGOCE.Text * 100, 2)
End If
If NVL(TXT_O_PROMO.Text, 0) > 0 Then
TXT_P_PROMO.Text = Round((TXT_O_PROMO.Text - TXT_C_PROMO.Text) / TXT_O_PROMO.Text * 100, 2)
End If
If NVL(TXT_O_TPROLUDIC.Text, 0) > 0 Then
TXT_P_TPROLUDIC.Text = Round((TXT_O_TPROLUDIC.Text - TXT_C_TPROLUDIC.Text) / TXT_O_TPROLUDIC.Text * 100, 2)
End If



TXT_MARGINE.Text = TotaleOfferta - CostiTotali
TXT_MARGINE_PER.Text = Round((TotaleOfferta - CostiTotali) / TotaleOfferta * 100, 2)

End Sub

Private Sub Margini_Provvigioni()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    StringaSQL = StringaSQL & " SUM(GB07_IMPPROVV) as A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06"



    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_C_PROVV.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Margini_ListinoGiochi()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(GB07_PREZZO) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (MG66_FAM_MG53 = 'PROL' or  MG66_FAM_MG53 = 'PROP')"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "

    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_O_LISTINOG.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Margini_ListinoRicambi()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(GB07_PREZZO) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (MG66_FAM_MG53 = 'PDET' )"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_O_LISTINOR.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Margini_Giochi()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(GB07_IMPORTO) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (MG66_FAM_MG53 = 'PROL' )"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_O_GIOCHI.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Margini_Promo()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(GB07_IMPORTO) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (MG66_FAM_MG53 = 'PROP' or  MG66_FAM_MG53 = 'NEGP')"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_O_PROMO.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Margini_Ricambi()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(GB07_IMPORTO) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (MG66_FAM_MG53 = 'PDET' )"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_O_RICAMBI.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Margini_NegoceLocal()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(GB07_IMPORTO) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (MG66_FAM_MG53 = 'NEGL' )"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_O_NLOCAL.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Margini_NegoceGruop()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(GB07_IMPORTO) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (MG66_FAM_MG53 = 'NEGG' )"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_O_NGROUP.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Margini_Antitrauma()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(GB07_IMPORTO) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (MG66_FAM_MG53 = 'SERV' and  MG66_SFAM_MG54 = 'ANTI')"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_O_ANTITRAUMA.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Margini_PosaLavori()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(GB07_IMPORTO) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (MG66_FAM_MG53 = 'SERV' and  MG66_SFAM_MG54 = 'POSA')"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_O_POSALAVORI.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Margini_TrasportoProludic()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(GB07_IMPORTO) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (GB07_CODART_MG66 = 'TRASPORTO GIOCHI' )"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_O_TPROLUDIC.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Margini_TrasportoNegoce()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(GB07_IMPORTO) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (GB07_CODART_MG66 = 'TRASPORTO ALTRO' )"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_O_TNEGOCE.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Margini_TrasportoAntitrauma()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(GB07_IMPORTO) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (GB07_CODART_MG66 = 'TRASPORTO GOMMA' )"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_O_TANTI.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

'costi

Private Sub Costi_Provvigioni()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    StringaSQL = StringaSQL & " SUM(GB07_IMPPROVV) as A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06"



    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_C_PROVV.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Costi_ListinoGiochi()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(PrezzoAcquistoTotale) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (MG66_FAM_MG53 = 'PROL' or  MG66_FAM_MG53 = 'PROP')"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_C_LISTINOG.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Costi_ListinoRicambi()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(PrezzoAcquistoTotale) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (MG66_FAM_MG53 = 'PDET' )"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_C_LISTINOR.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Costi_Giochi()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(PrezzoAcquistoTotale) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (MG66_FAM_MG53 = 'PROL' )"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_C_GIOCHI.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Costi_Promo()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(PrezzoAcquistoTotale) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (MG66_FAM_MG53 = 'PROP' or  MG66_FAM_MG53 = 'NEGP')"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_C_PROMO.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Costi_Ricambi()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(PrezzoAcquistoTotale) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (MG66_FAM_MG53 = 'PDET' )"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_C_RICAMBI.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Costi_NegoceLocal()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(PrezzoAcquistoTotale) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (MG66_FAM_MG53 = 'NEGL' )"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_C_NLOCAL.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Costi_NegoceGruop()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(PrezzoAcquistoTotale) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (MG66_FAM_MG53 = 'NEGG' )"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_C_NGROUP.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Costi_Antitrauma()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(PrezzoAcquistoTotale) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (MG66_FAM_MG53 = 'SERV' and  MG66_SFAM_MG54 = 'ANTI')"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_C_ANTITRAUMA.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Costi_PosaLavori()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(PrezzoAcquistoTotale) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (MG66_FAM_MG53 = 'SERV' and  MG66_SFAM_MG54 = 'POSA')"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_C_POSALAVORI.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Costi_TrasportoProludic()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(PrezzoAcquistoTotale) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (GB07_CODART_MG66 = 'TRASPORTO GIOCHI' )"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_C_TPROLUDIC.Text = CStr(CDbl(NVL(rstGridMargine("A"), 0)) + ((ValPercTrasp * CDbl((NVL(TXT_C_GIOCHI.Text, 0)) + CDbl(NVL(TXT_C_PROMO.Text, 0)) + CDbl(NVL(TXT_C_NGROUP.Text, 0)) + CDbl(NVL(TXT_C_RICAMBI.Text, 0)))) / 100))
   End If
   rstGridMargine.Close

End Sub

Private Sub Costi_TrasportoNegoce()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(PrezzoAcquistoTotale) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (GB07_CODART_MG66 = 'TRASPORTO ALTRO' )"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_C_TNEGOCE.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Costi_TrasportoAntitrauma()

 Dim StringaSQL As String
    
    StringaSQL = "SELECT     "
    
    StringaSQL = StringaSQL & " SUM(PrezzoAcquistoTotale) AS A"
    StringaSQL = StringaSQL & " From GBVW_MARGINIOFFERTA"
    StringaSQL = StringaSQL & " Where (GB07_ID_GB06 = " & IdOfferta & ")"
    StringaSQL = StringaSQL & " and  (GB07_CODART_MG66 = 'TRASPORTO GOMMA' )"
    StringaSQL = StringaSQL & " GROUP BY GB07_ID_GB06 "




    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstGridMargine = Gcls_RecordPadre.Gpr_GetADORecord
    With rstGridMargine
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With
   If Not rstGridMargine.EOF Then
   TXT_C_TANTI.Text = NVL(rstGridMargine("A"), 0)
   End If
   rstGridMargine.Close

End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap

    Set Gcls_Log = New CLSFW_SrvLog

'Posizionamento e dimensionamento form
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
' Impostare le dimensioni effettive della form
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  FormIsActive = False
End Sub

