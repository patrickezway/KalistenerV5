VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RichTx32.Ocx"
Begin VB.Form KA_Alerte 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3690
   ClientLeft      =   5520
   ClientTop       =   3495
   ClientWidth     =   9690
   DrawStyle       =   5  'Transparent
   Icon            =   "KA_Alerte.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   3
      Left            =   8760
      Picture         =   "KA_Alerte.frx":0BCA
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Afficher le résumé"
      Top             =   0
      Width           =   350
   End
   Begin VB.Frame frmStats 
      BackColor       =   &H80000018&
      Caption         =   "Documents :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1670
      Index           =   0
      Left            =   50
      TabIndex        =   0
      Top             =   300
      Width           =   3255
      Begin VB.PictureBox pctStats 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   3
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   2895
         TabIndex        =   28
         Top             =   1320
         Width           =   2895
         Begin VB.PictureBox pctIndic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   3
            Left            =   0
            Picture         =   "KA_Alerte.frx":0FBF
            ScaleHeight     =   150
            ScaleWidth      =   150
            TabIndex        =   30
            Top             =   20
            Width           =   150
         End
         Begin VB.PictureBox pctPlus 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   3
            Left            =   2520
            Picture         =   "KA_Alerte.frx":1351
            ScaleHeight     =   285
            ScaleWidth      =   300
            TabIndex        =   29
            ToolTipText     =   "Accès au résumé"
            Top             =   0
            Width           =   300
         End
         Begin VB.Label lblLib 
            BackStyle       =   0  'Transparent
            Caption         =   "Informations :"
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
            Index           =   3
            Left            =   240
            TabIndex        =   32
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label lblStats 
            BackStyle       =   0  'Transparent
            Caption         =   "0 / 0"
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
            Index           =   3
            Left            =   1560
            TabIndex        =   31
            Top             =   15
            Width           =   855
         End
      End
      Begin VB.PictureBox pctStats 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   2
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   2895
         TabIndex        =   23
         Top             =   960
         Width           =   2895
         Begin VB.PictureBox pctIndic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   2
            Left            =   0
            Picture         =   "KA_Alerte.frx":17A8
            ScaleHeight     =   150
            ScaleWidth      =   150
            TabIndex        =   25
            Top             =   0
            Width           =   150
         End
         Begin VB.PictureBox pctPlus 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   2
            Left            =   2520
            Picture         =   "KA_Alerte.frx":1B38
            ScaleHeight     =   285
            ScaleWidth      =   300
            TabIndex        =   24
            ToolTipText     =   "Accès au résumé"
            Top             =   0
            Width           =   300
         End
         Begin VB.Label lblLib 
            BackStyle       =   0  'Transparent
            Caption         =   "AR :"
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
            Index           =   2
            Left            =   240
            TabIndex        =   27
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lblStats 
            BackStyle       =   0  'Transparent
            Caption         =   "0 / 0"
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
            Index           =   2
            Left            =   1560
            TabIndex        =   26
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox pctStats 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   1
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   2895
         TabIndex        =   18
         Top             =   600
         Width           =   2895
         Begin VB.PictureBox pctIndic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   1
            Left            =   0
            Picture         =   "KA_Alerte.frx":1F8F
            ScaleHeight     =   150
            ScaleWidth      =   150
            TabIndex        =   20
            Top             =   20
            Width           =   150
         End
         Begin VB.PictureBox pctPlus 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   1
            Left            =   2520
            Picture         =   "KA_Alerte.frx":231F
            ScaleHeight     =   285
            ScaleWidth      =   300
            TabIndex        =   19
            ToolTipText     =   "Accès au résumé"
            Top             =   20
            Width           =   300
         End
         Begin VB.Label lblLib 
            BackStyle       =   0  'Transparent
            Caption         =   "Modifications :"
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
            Index           =   1
            Left            =   240
            TabIndex        =   22
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label lblStats 
            BackStyle       =   0  'Transparent
            Caption         =   "0 / 0"
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
            Index           =   1
            Left            =   1560
            TabIndex        =   21
            Top             =   15
            Width           =   855
         End
      End
      Begin VB.PictureBox pctStats 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   0
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   2895
         TabIndex        =   13
         Top             =   240
         Width           =   2895
         Begin VB.PictureBox pctPlus 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   0
            Left            =   2520
            Picture         =   "KA_Alerte.frx":2776
            ScaleHeight     =   285
            ScaleWidth      =   300
            TabIndex        =   17
            ToolTipText     =   "Accès au résumé"
            Top             =   0
            Width           =   300
         End
         Begin VB.PictureBox pctIndic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   0
            Left            =   0
            Picture         =   "KA_Alerte.frx":2BCD
            ScaleHeight     =   150
            ScaleWidth      =   150
            TabIndex        =   14
            Top             =   0
            Width           =   150
         End
         Begin VB.Label lblLib 
            BackStyle       =   0  'Transparent
            Caption         =   "Actions :"
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
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Top             =   0
            Width           =   975
         End
         Begin VB.Label lblStats 
            BackStyle       =   0  'Transparent
            Caption         =   "0 / 0"
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
            Index           =   0
            Left            =   1560
            TabIndex        =   15
            Top             =   0
            Width           =   855
         End
      End
   End
   Begin VB.Frame frmStats 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Divers :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1670
      Index           =   2
      Left            =   6195
      TabIndex        =   12
      Top             =   300
      Width           =   3375
      Begin VB.PictureBox pctStats 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   4
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   2895
         TabIndex        =   41
         Top             =   360
         Width           =   2895
         Begin VB.PictureBox pctIndic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   4
            Left            =   0
            Picture         =   "KA_Alerte.frx":2F5D
            ScaleHeight     =   210
            ScaleWidth      =   225
            TabIndex        =   47
            Top             =   0
            Width           =   220
         End
         Begin VB.PictureBox pctPlus 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   4
            Left            =   2520
            Picture         =   "KA_Alerte.frx":32F4
            ScaleHeight     =   285
            ScaleWidth      =   300
            TabIndex        =   42
            ToolTipText     =   "Accès au résumé"
            Top             =   0
            Width           =   300
         End
         Begin VB.Label lblStats 
            BackColor       =   &H00C0E0FF&
            Caption         =   "0 / 0"
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
            Index           =   4
            Left            =   1320
            TabIndex        =   44
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label lblLib 
            BackColor       =   &H00C0E0FF&
            Caption         =   "KaliMails :"
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
            Index           =   4
            Left            =   250
            TabIndex        =   43
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.PictureBox pctStats 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   5
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   2895
         TabIndex        =   40
         Top             =   840
         Width           =   2895
         Begin VB.Label lblStats 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Aucun classeur"
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
            Index           =   5
            Left            =   1320
            TabIndex        =   46
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label lblLib 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Classeurs :"
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
            Index           =   5
            Left            =   120
            TabIndex        =   45
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.PictureBox pctStats 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   6
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   2895
         TabIndex        =   36
         Top             =   1320
         Width           =   2895
         Begin VB.PictureBox pctIndic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   5
            Left            =   0
            Picture         =   "KA_Alerte.frx":374B
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   48
            Top             =   0
            Width           =   240
         End
         Begin VB.PictureBox pctPlus 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   6
            Left            =   2520
            Picture         =   "KA_Alerte.frx":3B3C
            ScaleHeight     =   285
            ScaleWidth      =   300
            TabIndex        =   37
            ToolTipText     =   "Accès au résumé"
            Top             =   0
            Width           =   300
         End
         Begin VB.Label lblStats 
            BackColor       =   &H00C0E0FF&
            Caption         =   "0 / 0"
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
            Index           =   6
            Left            =   1440
            TabIndex        =   39
            Top             =   0
            Width           =   975
         End
         Begin VB.Label lblLib 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Formulaires :"
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
            Index           =   6
            Left            =   360
            TabIndex        =   38
            Top             =   0
            Width           =   1215
         End
      End
   End
   Begin VB.Frame frmStats 
      BackColor       =   &H80000016&
      Caption         =   "Gestion des risques :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1670
      Index           =   1
      Left            =   3420
      TabIndex        =   11
      Top             =   300
      Width           =   2680
      Begin VB.PictureBox pctStats 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   7
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   2535
         TabIndex        =   33
         Top             =   720
         Width           =   2535
         Begin VB.PictureBox pctPlus 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   280
            Index           =   7
            Left            =   2160
            Picture         =   "KA_Alerte.frx":3F93
            ScaleHeight     =   285
            ScaleWidth      =   300
            TabIndex        =   52
            ToolTipText     =   "Accès au résumé"
            Top             =   0
            Width           =   300
         End
         Begin VB.Label lblStats 
            BackColor       =   &H80000016&
            Caption         =   "0 / 0"
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
            Index           =   7
            Left            =   1320
            TabIndex        =   35
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lblLib 
            BackColor       =   &H80000016&
            Caption         =   "Qualifications :"
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
            Index           =   7
            Left            =   0
            TabIndex        =   34
            Top             =   0
            Width           =   1335
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   -480
      ScaleHeight     =   300
      ScaleWidth      =   10095
      TabIndex        =   6
      Top             =   0
      Width           =   10095
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9720
         Picture         =   "KA_Alerte.frx":43EA
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   7
         ToolTipText     =   "Fermer"
         Top             =   0
         Width           =   300
      End
      Begin ComctlLib.ImageList imglistMask 
         Left            =   8640
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   14
         ImageHeight     =   14
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   2
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "KA_Alerte.frx":48B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "KA_Alerte.frx":4E98
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "KaliAlerte - V.1.00.02"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   45
         Width           =   6255
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   840
      Top             =   3960
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   4200
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   4320
   End
   Begin VB.Frame frmDet 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Détail des indicateurs"
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   125
      TabIndex        =   1
      Top             =   2040
      Width           =   9400
      Begin VB.Frame frmPatience 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   0
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   9375
         Begin ComctlLib.ProgressBar pgb 
            Height          =   495
            Left            =   1440
            TabIndex        =   54
            Top             =   600
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   873
            _Version        =   327682
            Appearance      =   1
         End
         Begin VB.Label lblPatience 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Chargement en cours ..."
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
            Left            =   2880
            TabIndex        =   55
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   240
         ScaleHeight     =   225
         ScaleWidth      =   255
         TabIndex        =   49
         Top             =   50
         Width           =   255
      End
      Begin VB.CommandButton cmd 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   2
         Left            =   7680
         Picture         =   "KA_Alerte.frx":547A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Accès au détail"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   650
      End
      Begin VB.CommandButton cmd 
         Caption         =   ">>"
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
         Index           =   1
         Left            =   9000
         TabIndex        =   4
         ToolTipText     =   "Suivant"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmd 
         Caption         =   "<<"
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
         Index           =   0
         Left            =   8640
         TabIndex        =   3
         ToolTipText     =   "Précédent"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin RichTextLib.RichTextBox rtxt 
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   1720
         _Version        =   393217
         BackColor       =   -2147483629
         ScrollBars      =   3
         TextRTF         =   $"KA_Alerte.frx":5874
      End
      Begin VB.Image imgNew 
         Height          =   240
         Left            =   7080
         Picture         =   "KA_Alerte.frx":58F6
         Top             =   120
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblNb 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3120
         TabIndex        =   51
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label lblDet 
         BackStyle       =   0  'Transparent
         Caption         =   "Résumé :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   10
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Dernière Mise à jour : "
         Height          =   200
         Left            =   7200
         TabIndex        =   9
         Top             =   1320
         Width           =   2175
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   3600
      Y2              =   0
   End
   Begin ComctlLib.ImageList ImageList 
      Left            =   3480
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   10
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "KA_Alerte.frx":59A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "KA_Alerte.frx":5E72
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "KA_Alerte.frx":633C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "KA_Alerte.frx":6806
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "KA_Alerte.frx":6CD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "KA_Alerte.frx":7222
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Index           =   2
      X1              =   9600
      X2              =   9600
      Y1              =   3600
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Index           =   3
      X1              =   0
      X2              =   9600
      Y1              =   1995
      Y2              =   1995
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Index           =   1
      X1              =   0
      X2              =   9600
      Y1              =   3650
      Y2              =   3650
   End
End
Attribute VB_Name = "KA_Alerte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CYCLE_RELECTURE = 102

'Indicateurs
Private Const INDIC_ACTION = 0
Private Const INDIC_MODIF = 1
Private Const INDIC_AR = 2
Private Const INDIC_INFO = 3
Private Const INDIC_KMAIL = 4
Private Const INDIC_CLASS = 5
Private Const INDIC_FORM = 6
Private Const INDIC_KGDR = 7

'Gestion des indicateurs
Private Const INDIC_NO = 0
Private Const INDIC_OK = 1

'BOUTONS
Private Const CMD_PREC = 0
Private Const CMD_SUIV = 1
Private Const CMD_DET = 2
Private Const CMD_UP = 3

'Images
Private Const IMG_BAS = 1
Private Const IMG_HAUT = 2

Private Const IMG_ACTION = 1
Private Const IMG_MODIF = 2
Private Const IMG_AR = 3
Private Const IMG_INFO = 4
Private Const IMG_KMAIL = 5
Private Const IMG_KFORM = 6

'Tableau contenant le détails des indicateurs
Private Type p_indic_detail
    stxt As String
    stype As String
    surl As String
    bvu As Boolean
End Type
Private p_indic() As p_indic_detail
Private p_nbindic As Long
Private g_position As Integer

Private g_nbtimer As Integer

'Tableau de gestion des indicateurs
Private g_bindic() As Boolean

Private g_mode_saisie As Boolean
Private g_left As Integer

Public taille

Public Sub AppelFrm()

    Me.Show

End Sub

Public Function P_affiche_new() As Integer
    
    p_Result = SetForegroundWindow(Me.hwnd)
    
    Me.Left = Screen.Width
    Me.Visible = True
    g_left = 100
    
continue:
    Me.Left = Me.Left - g_left
    g_left = g_left * 2
    
    If Me.Left + Me.Width <= Screen.Width Then
        Me.Left = Screen.Width - Me.Width
    Else
        Call SYS_Sleep(150)
        Me.Refresh
        GoTo continue
    End If

End Function

'***********************************************
'Charger le détail des actions à réaliser
'***********************************************
Private Function charger_actions() As Integer

    Dim sql As String, stitre As String, sref As String, slibvers As String
    Dim scycle As String, scomm As String, sdatep As String, sdate As String
    Dim i As Integer, icyordre As Integer
    Dim rs As rdoResultset
    
    'Gestion des actions
    sql = "SELECT D_Num, D_Titre, D_Ident, D_LibVers, DAC_CyOrdre, DAC_DatePrevue, DAC_Date, DAC_ActionVu " _
        & " FROM DocAction, Document" _
        & " where D_Num=DAC_DNum and dac_unum=" & p_numUtil _
        & " ORDER BY DAC_ActionVu"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        charger_actions = P_ERREUR
        Exit Function
    End If
    
    While Not rs.EOF
        'Nom de l'étpe
        If rs("DAC_CyOrdre").Value <> CYCLE_RELECTURE Then
            If Odbc_RecupVal("SELECT CY_Action " _
                           & "FROM Cycle " _
                           & "WHERE CY_Ordre=" & rs("DAC_CyOrdre").Value, scycle) = P_ERREUR Then
                charger_actions = P_ERREUR
                Exit Function
            End If
        Else
            scycle = "Relecture"
        End If
        
        'Message d'information
        If sdatep <> "" Then
            scomm = scycle & "(Pour le " & rs("DAC_DatePrevue").Value & ") " & rs("D_Ident").Value & " / " & rs("D_Titre").Value & " - version : " & rs("D_LibVers").Value
        Else
            scomm = scycle & "(Dem. le " & rs("DAC_Date").Value & ") " & rs("D_Ident").Value & " / " & rs("D_Titre").Value & " - version : " & rs("D_LibVers").Value
        End If
        
        'Ajouter dans le tableau
        ReDim Preserve p_indic(p_nbindic) As p_indic_detail
        p_indic(p_nbindic).stxt = scomm
        p_indic(p_nbindic).stype = "action"
        p_indic(p_nbindic).surl = "in=action&numd=" & rs("D_Num").Value
        p_indic(p_nbindic).bvu = rs("DAC_ActionVu").Value
                
        p_nbindic = p_nbindic + 1
        
        rs.MoveNext
    Wend
    rs.Close
    
    'Gestion des ??
    sql = "SELECT D_Num, D_Titre, D_Ident, D_LibVers, DAC_CyOrdre, DAC_DatePrevue, DAC_Date, DAC_ActionVu " _
        & " from DocAction, Document" _
        & " where D_Num=DAC_DNum  and dac_unummodif=" & p_numUtil _
        & " ORDER BY DAC_ActionVu"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        charger_actions = P_ERREUR
        Exit Function
    End If
    
    While Not rs.EOF
        'Nom de l'étpe
        If Odbc_RecupVal("SELECT CY_Action " _
                       & "FROM Cycle " _
                       & "WHERE CY_Ordre=" & rs("DAC_CyOrdre").Value, scycle) = P_ERREUR Then
            charger_actions = P_ERREUR
            Exit Function
        End If
        
        'Message d'information
        If sdatep <> "" Then
            scomm = scycle & "(Pour le " & rs("DAC_DatePrevue").Value & ") " & rs("D_Ident").Value & " / " & rs("D_Titre").Value & " - version : " & rs("D_LibVers").Value
        Else
            scomm = scycle & "(Dem. le " & rs("DAC_Date").Value & ") " & rs("D_Ident").Value & " / " & rs("D_Titre").Value & " - version : " & rs("D_LibVers").Value
        End If

        'Ajouter dans le tableau
        ReDim Preserve p_indic(p_nbindic) As p_indic_detail
        p_indic(p_nbindic).stxt = scomm
        p_indic(p_nbindic).stype = "action"
        p_indic(p_nbindic).surl = "in=action&numd=" & rs("D_Num").Value
        p_indic(p_nbindic).bvu = rs("DAC_ActionVu").Value
        
        p_nbindic = p_nbindic + 1
        
        rs.MoveNext
    Wend
    rs.Close
    
    'Gestion des revisions
    sql = "SELECT D_Num, D_Titre, D_Ident, D_LibVers, D_DateRevision, DRV_Vu" _
        & " from docareviser, document" _
        & " where drv_dnum=d_num and D_UnumResp=" & p_numUtil _
        & " ORDER BY DRV_Vu"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        charger_actions = P_ERREUR
        Exit Function
    End If
    
    While Not rs.EOF
        
        If CDate(rs("D_DateRevision").Value) < Date Then
            sdatep = ""
        Else
            sdatep = "OK"
        End If
        
        'Message d'information
        If sdatep <> "" Then
            scomm = "A réviser (Pour le " & rs("D_DateRevision").Value & ") " & rs("D_Ident").Value & " / " & rs("D_Titre").Value & " - version : " & rs("D_LibVers").Value
        Else
            scomm = "A réviser (Dem. le " & rs("D_DateRevision").Value & ") " & rs("D_Ident").Value & " / " & rs("D_Titre").Value & " - version : " & rs("D_LibVers").Value
        End If

        'Ajouter dans le tableau
        ReDim Preserve p_indic(p_nbindic) As p_indic_detail
        p_indic(p_nbindic).stxt = scomm
        p_indic(p_nbindic).stype = "action"
        p_indic(p_nbindic).surl = "in=action&numd=" & rs("D_Num").Value
        p_indic(p_nbindic).bvu = rs("DRV_Vu").Value
        
        p_nbindic = p_nbindic + 1
        
        rs.MoveNext
    Wend
    rs.Close
    
    charger_actions = P_OK

End Function

'***********************************************
'Charger le détail des AR à réaliser
'***********************************************
Private Function charger_ar() As Integer

    Dim sql As String, scomm As String
    Dim rs As rdoResultset
    Dim i As Integer
    
    'Gestion des actions
    sql = "select D_Titre, D_Ident, DD_LibVers, DD_DateDiffusion, D_Num, DD_DiffusionVu " _
        & " from DocDiffusion, Document" _
        & " where DD_UNum=" & p_numUtil _
        & " and DD_DateAR is Null" _
        & " and DD_DateDiffusion is not Null" _
        & " and DD_ARecuperer=false" _
        & " and D_Num=DD_DNum" _
        & " ORDER BY DD_DiffusionVu"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        charger_ar = P_ERREUR
        Exit Function
    End If
    
    While Not rs.EOF
        scomm = rs("D_Ident").Value & " / " & rs("D_Titre").Value & " - version : " & rs("DD_LibVers").Value & ", diffusée le " & rs("DD_DateDiffusion").Value
        
        'Ajouter dans le tableau
        ReDim Preserve p_indic(p_nbindic) As p_indic_detail
        p_indic(p_nbindic).stxt = scomm
        p_indic(p_nbindic).stype = "ar"
        p_indic(p_nbindic).surl = "in=ar&dc=" & rs("D_Num").Value & "-" & rs("DD_Libvers").Value
        p_indic(p_nbindic).bvu = rs("DD_DiffusionVu").Value
        
        p_nbindic = p_nbindic + 1

        rs.MoveNext
    Wend
    rs.Close
    
    charger_ar = P_OK

End Function

'***********************************************
'Récupération du nom de l'utilisateur identifié
'***********************************************
Private Function charger_identifant() As Integer

    Dim snom As String
    Dim sprenom As String
    
    If Odbc_RecupVal("SELECT U_Nom, U_Prenom " _
                    & "From Utilisateur " _
                    & "WHERE U_Num=" & p_numUtil, snom, sprenom) = P_ERREUR Then
        charger_identifant = P_ERREUR
        Exit Function
    End If
    
    Label1.Caption = "KaliAlerte - V1.00.02 - " & UCase(Left$(sprenom, 1) & "." & snom)

    charger_identifant = P_OK

End Function

'***********************************************
'Charger les indicateurs disponibles
'***********************************************
Private Function charger_Indicateur() As Integer

    Dim i As Integer

    For i = 0 To INDIC_KGDR
        
        ReDim Preserve g_bindic(i) As Boolean
        
        'Si l'indicateur est géré
        If STR_GetChamp(p_salerte, ";", i) = INDIC_OK Then
            g_bindic(i) = True
            pctStats(i).Visible = True
        Else
            g_bindic(i) = False
            pctStats(i).Visible = False
        End If
    
    Next i

    charger_Indicateur = P_OK

End Function

Private Function charger_kform() As Integer

    Dim sql As String
    Dim rs As rdoResultset
    Dim vtxt As String, titre As String
    Dim scomm As String
    
    sql = "SELECT * " _
        & "FROM Donnees_Encours " _
        & "WHERE DONEC_UNum=" & p_numUtil _
        & " ORDER BY DONEC_Vu, DONEC_Date Desc"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        charger_kform = P_ERREUR
        Exit Function
    End If
    
    While Not rs.EOF
    
        ' On récupère les informations de la déclaration
        If recup_detail(rs("DONEC_FORNum").Value, rs("DONEC_DONNum").Value, rs("DONEC_NumEtape").Value, vtxt, titre) = P_ERREUR Then
            charger_kform = P_ERREUR
            Exit Function
        End If
        
        scomm = titre & vbCrLf & vtxt
        
        'Ajouter dans le tableau
        ReDim Preserve p_indic(p_nbindic) As p_indic_detail
        p_indic(p_nbindic).stxt = scomm
        p_indic(p_nbindic).stype = "kform"
        p_indic(p_nbindic).surl = "in=action_form&V_numfor=" & rs("DONEC_FORNum").Value & "&V_numdon=" & rs("DONEC_DONNum").Value
        p_indic(p_nbindic).bvu = rs("DONEC_Vu").Value
        
        p_nbindic = p_nbindic + 1
        
        pgb.Value = pgb.Value + 1
        If pgb.Value = pgb.Max Then
            pgb.Value = 0
        End If
        
        rs.MoveNext
    Wend
    rs.Close

    charger_kform = P_OK
    
End Function

Private Function charger_kgdr() As Integer

    Dim titre As String
    Dim vtxt As Variant
    Dim i As Long
    Dim scomm As String
    Dim rs As rdoResultset, rs2 As rdoResultset
    Dim sql As String
    
    'Pour chaque domaine
    If Odbc_SelectV("SELECT DMR_Num " _
                  & "FROM DomaineRisk " _
                  & "ORDER BY DMR_Libelle", rs) = P_ERREUR Then
        charger_kgdr = P_ERREUR
        Exit Function
    End If
    
    While Not rs.EOF
        
        'Pour chaque déclaration
        sql = "SELECT GDR_FORNum, GDR_DONNum, GDR_Etat " _
            & "FROM GestionRisk_" & rs("DMR_Num").Value _
            & " WHERE GDR_Etat<2 " _
            & "ORDER BY GDR_Etat, GDR_DONDateEvt Desc"
        If Odbc_SelectV(sql, rs2) = P_ERREUR Then
            charger_kgdr = P_ERREUR
            Exit Function
        End If
        
        While Not rs2.EOF
            ' On récupère les informations de la déclaration
            If recup_detail(rs2("GDR_FORNum").Value, rs2("GDR_DONNum").Value, 0, vtxt, titre) = P_ERREUR Then
                charger_kgdr = P_ERREUR
                Exit Function
            End If
            
            scomm = titre & vbCrLf & vtxt
            
            'Ajouter dans le tableau
            ReDim Preserve p_indic(p_nbindic) As p_indic_detail
            p_indic(p_nbindic).stxt = scomm
            p_indic(p_nbindic).stype = "kgdr"
            p_indic(p_nbindic).surl = "in=saisir_form&V_numfor=" & rs2("GDR_FORNum").Value & "&V_numdon=" & rs2("GDR_DONNum").Value
            p_indic(p_nbindic).bvu = IIf(rs2("GDR_Etat").Value = 0, False, True)
            
            p_nbindic = p_nbindic + 1
            
            pgb.Value = pgb.Value + 1
            If pgb.Value = pgb.Max Then
                pgb.Value = 0
            End If
            
            rs2.MoveNext
        
        Wend
        rs2.Close
        rs.MoveNext
    Wend
    rs.Close
    
    charger_kgdr = P_OK
    
End Function

Private Function charger_kmail() As Integer

    Dim sql As String
    Dim rs As rdoResultset
    Dim scomm As String
    Dim i As Integer
    
    sql = "SELECT KMD_DateEnvoi, KM_Sujet, U_Nom, KM_Num, KMD_DateLu " _
        & "FROM KaliMail, KaliMailDetail, Utilisateur " _
        & "WHERE KMD_DestNum=U_Num " _
        & "AND KM_Num=KMD_KMNum " _
        & "AND KMD_SuppDest=0 " _
        & "AND KMD_DestNum=" & p_numUtil _
        & " ORDER BY KMD_DateLu Desc"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        charger_kmail = P_ERREUR
        Exit Function
    End If
    
    While Not rs.EOF
        
        scomm = "De : " & rs("U_Nom").Value & vbCrLf & "Le :" & rs("KMD_DateEnvoi").Value & vbCrLf & "Objet : " & rs("KM_Sujet").Value
        
        'Ajouter dans le tableau
        ReDim Preserve p_indic(p_nbindic) As p_indic_detail
        p_indic(p_nbindic).stxt = scomm
        p_indic(p_nbindic).stype = "kmail"
        p_indic(p_nbindic).surl = "in=kalimail&V_nummail=" & rs("KM_Num").Value & "&V_numdest=" & p_numUtil
        p_indic(p_nbindic).bvu = IIf(IsNull(rs("KMD_Datelu").Value), False, True)
        
        p_nbindic = p_nbindic + 1
        
        rs.MoveNext
    Wend
    rs.Close
    
    charger_kmail = P_OK
    
End Function

'*******************************************************************
' Charger les informations concernant les demandes de modifications
'*******************************************************************
Private Function charger_modif() As Integer

    Dim sql As String, scomm As String
    Dim rs As rdoResultset
    Dim i As Integer
    
    'Infos provenant d'un document
    sql = "select ddm_dnum, ddm_datedemande, ddm_modifvu, ddm_textedemande, d_ident, d_titre, d_libvers, u_nom " _
        & "from DocDemModif, Document, Utilisateur " _
        & "where DDM_DNum=D_Num AND DDM_UNumDest=" & p_numUtil _
        & " AND U_Num=DDM_UnumDest AND DDM_DNum>0 " _
        & "and (DDM_Etat=0 or (DDM_Etat=2 and DDM_DateFait is null)) " _
        & "ORDER BY DDM_ModifVu"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        charger_modif = P_ERREUR
        Exit Function
    End If
    
    While Not rs.EOF
        
        scomm = "Demande de modification : " & rs("D_Ident").Value & " / " & rs("D_Titre").Value & " - version : " & rs("D_LibVers").Value & " demande du " & rs("DDM_DateDemande").Value & " par " & rs("U_Nom").Value & " :" & vbLf & rs("DDM_TexteDemande").Value
        
        'Ajouter dans le tableau
        ReDim Preserve p_indic(p_nbindic) As p_indic_detail
        p_indic(p_nbindic).stxt = scomm
        p_indic(p_nbindic).stype = "modif"
        p_indic(p_nbindic).surl = "in=demande&numdem=0"
        p_indic(p_nbindic).bvu = rs("DDM_ModifVu").Value
        
        p_nbindic = p_nbindic + 1
        
        rs.MoveNext
    Wend
    rs.Close
    
    sql = "select DACR_Titre, DACR_Descr, U_Nom, DDM_ModifVu " _
        & "from DocDemModif, DocACreer, Utilisateur " _
        & "where DDM_DNum<0 AND DDM_DNum=-(DACR_Num) " _
        & " AND DDM_UNum=U_Num AND DDM_UnumDest=" & p_numUtil _
        & " and (DDM_Etat=0 or (DDM_Etat=2 and DDM_DateFait is null)) " _
        & "ORDER BY DDM_ModifVu"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        charger_modif = P_ERREUR
        Exit Function
    End If
    
    While Not rs.EOF
        
        scomm = "Demande de création : " & rs("DACR_Titre").Value & " / " & rs("DACR_Descr").Value & " - demande du " & rs("DDM_DateDemande").Value & " par " & rs("U_Nom").Value
        
        'Ajouter dans le tableau
        ReDim Preserve p_indic(p_nbindic) As p_indic_detail
        p_indic(p_nbindic).stxt = scomm
        p_indic(p_nbindic).stype = "modif"
        p_indic(p_nbindic).surl = "in=demande&numdem=0"
        p_indic(p_nbindic).bvu = rs("DDM_ModifVu").Value
        
        p_nbindic = p_nbindic + 1
        
        rs.MoveNext
    Wend
    rs.Close
    
    charger_modif = P_OK
    
End Function

'*******************************************************************
' Charger les informations concernant les informations
'*******************************************************************
Private Function charger_info() As Integer

    Dim sql As String, scomm As String
    Dim rs As rdoResultset
    Dim i As Integer
    
    'Infos provenant d'un document
    sql = "select D_Ident, D_Titre, DI_LibVers, DI_Date, DI_Commentaire, DI_Num, DI_InfoVu" _
        & " from DocInfo, Document" _
        & " where DI_UNum=" & p_numUtil _
        & " and D_Num=DI_DNum" _
        & "ORDER BY DI_InfoVu"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        charger_info = P_ERREUR
        Exit Function
    End If
    
    While Not rs.EOF
        
        scomm = rs("D_Ident").Value & " / " & rs("D_Titre").Value & " - version" & rs("DI_LibVers").Value & ", du " & rs("DI_Date").Value & " " & rs("DI_Commentaire").Value
        
        'Ajouter dans le tableau
        ReDim Preserve p_indic(p_nbindic) As p_indic_detail
        p_indic(p_nbindic).stxt = scomm
        p_indic(p_nbindic).stype = "info"
        p_indic(p_nbindic).surl = "in=info&numi=" & rs("DI_Num").Value
        p_indic(p_nbindic).bvu = rs("DI_InfoVu").Value
        
        p_nbindic = p_nbindic + 1
        
        rs.MoveNext
    Wend
    rs.Close
    
    charger_info = P_OK

End Function

'***********************************************
'Réévaluation des chaque indicateur
'***********************************************
Private Function evalue_action(ByRef r_bnew) As Integer

    Dim sql As String
    Dim rs As rdoResultset
 
    ' KALIDOC
    sql = "SELECT SUM(UA_NbAction) as NbAction, SUM(UA_NbModif) as NbModif, SUM(UA_NbDiff) as NbDiff, SUM(UA_NbInfo) as NbInfo " _
        & " FROM UtilADIM" _
        & " WHERE UA_DONum>0 AND UA_Unum=" & p_numUtil
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        evalue_action = P_ERREUR
        Exit Function
    End If
    
    ' Si actions
    While Not rs.EOF
        
        'ACTIONS
        If g_bindic(INDIC_ACTION) Then
            If verif_actions(rs("NbAction").Value) = P_ERREUR Then
                evalue_action = P_ERREUR
                Exit Function
            End If
        End If
        
        'MODIF
        If g_bindic(INDIC_MODIF) Then
            If verif_modifications(rs("NbModif").Value) = P_ERREUR Then
                evalue_action = P_ERREUR
                Exit Function
            End If
        End If
        
        'AR
        If g_bindic(INDIC_AR) Then
            If verif_AR(rs("NbDiff").Value) = P_ERREUR Then
                evalue_action = P_ERREUR
                Exit Function
            End If
        End If
        
        'INFOS
        If g_bindic(INDIC_INFO) Then
            If verif_infos(rs("NbInfo").Value) = P_ERREUR Then
                evalue_action = P_ERREUR
                Exit Function
            End If
        End If
        
        rs.MoveNext
    Wend
    rs.Close
    
    ' CLASSEURS
    If g_bindic(INDIC_CLASS) Then
        If verif_Classeurs() = P_ERREUR Then
            evalue_action = P_ERREUR
            Exit Function
        End If
    End If
    
    ' KALIMAIL
    If g_bindic(INDIC_KMAIL) Then
        If verif_KaliMail() = P_ERREUR Then
            evalue_action = P_ERREUR
            Exit Function
        End If
    End If
    
    ' KALIFORM
    If g_bindic(INDIC_FORM) Then
        If verif_KaliForm() = P_ERREUR Then
            evalue_action = P_ERREUR
            Exit Function
        End If
    End If
    
    ' KALIGDR
    If g_bindic(INDIC_KGDR) Then
        If verif_KGDR() = P_ERREUR Then
            evalue_action = P_ERREUR
            Exit Function
        End If
    End If
    
    g_mode_saisie = True
    
    evalue_action = P_OK

End Function

'***********************************************
'INITIALISATION
'***********************************************
Private Sub initialiser()

    If p_numUtil = 0 Then
        Call quitter
        Exit Sub
    End If

    If charger_identifant = P_ERREUR Then
        Call quitter
        Exit Sub
    End If

    'Chargement des indicateurs à vérifier
    If charger_Indicateur = P_ERREUR Then
        Call quitter
        Exit Sub
    End If
    
    Me.Height = frmStats(0).Top + frmStats(0).Height + 70
    Set cmd(CMD_UP).Picture = imglistMask.ListImages(IMG_BAS).Picture
    cmd(CMD_UP).ToolTipText = "Afficher le résumé"
    
    'Activer le timer
    Timer1.Enabled = True

End Sub

'*******************************************************************
' QUITTER
'*******************************************************************
Private Function quitter()

    g_left = 100
    
continue:
    Me.Left = Me.Left + g_left
    g_left = g_left * 2
    If Me.Left > Screen.Width Then
        Me.Left = Screen.Width - Me.Width
        Me.Visible = False
    Else
        Call SYS_Sleep(100)
        GoTo continue
    End If
    
End Function

Public Function recup_detail(ByVal v_numfor As Long, _
                              ByVal v_numdon As Long, _
                              ByVal v_numetape As Integer, _
                              ByRef r_rtxt As Variant, _
                              ByRef r_titre As String) As Integer
                              
    Dim sql As String, nometape As String, libvalid As String, libannul As String, sdatfin As String
    Dim nom As String, lib_action As String, label As String, lib As String, nomutil As String
    Dim rs As rdoResultset, rs_don As rdoResultset, rs_chp As rdoResultset
    Dim last_etape As Long, numetape As Long, ietape As Long
    Dim fmodif As Boolean
    Dim sval As Variant

    
    r_rtxt = ""

    ' Récupère la ligne de données
    sql = "select * from donnees_" & v_numfor & " where don_num=" & v_numdon
    If Odbc_SelectV(sql, rs_don) = P_ERREUR Then
        recup_detail = ""
        Exit Function
    End If
    If rs_don.EOF Then
        rs_don.Close
        recup_detail = P_ERREUR
        Exit Function
    End If

    ' Etapes effectuées
     If v_numetape = 0 Then
        sql = "select max(done_numetape) from donnees_etape" _
              & " where done_fornum=" & v_numfor _
              & " and done_donnum=" & v_numdon
        If Odbc_MinMax(sql, last_etape) = P_ERREUR Then
            recup_detail = ""
            Exit Function
        End If
        sql = "select max(fore_numetape) from formetape" _
              & " where fore_fornum=" & v_numfor
        If Odbc_MinMax(sql, numetape) = P_ERREUR Then
            recup_detail = ""
            Exit Function
        End If
    Else
        numetape = v_numetape - 1
    End If
    
    For ietape = 1 To numetape
        ' Etape + la personne associée
        ' Libellé de l'étape
        sql = "select fore_lib, fore_btnvalid, fore_btnannul from formetape" _
              & " where fore_fornum=" & v_numfor & " and fore_numetape=" & ietape
        Call Odbc_RecupVal(sql, nometape, libvalid, libannul)
        ' Valideur de l'étape
        sql = "select done_unum, done_date, u_nom, u_prenom from donnees_etape left join utilisateur" _
              & " on donnees_etape.done_unum=utilisateur.u_num" _
              & " where done_fornum=" & v_numfor _
              & " and done_donnum=" & v_numdon _
              & " and done_numetape=" & ietape
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            GoTo lab_etape_suiv
        End If
        If rs.EOF Then
            rs.Close
             GoTo lab_etape_suiv
        End If
        If rs("u_prenom").Value <> "" Then
            nom = Left$(rs("u_prenom").Value, 1) & ". "
        Else
            nom = ""
        End If
        nom = nom & rs("u_nom").Value
        If ietape = last_etape Then
            If rs_don("don_etat").Value = 1 Then
                lib_action = libvalid
            Else
                lib_action = libannul
            End If
        Else
            lib_action = libvalid
        End If
        
        r_titre = "      " & nometape & " (" & lib_action & ") : " & nom & " le " & Format(rs("done_date").Value, "dd/mm/yyyy")
        
        ' Champs de l'étape
        sql = "select * from formetapechp where forec_fornum=" & v_numfor _
              & " and forec_numetape=" & ietape _
              & " and forec_nom<>'don_nomutil'" _
              & " and forec_type<>'METHODE'" _
              & " order by forec_ordre"
        Call Odbc_SelectV(sql, rs_chp)
        While Not rs_chp.EOF
            If rs_chp("forec_type").Value <> "BUTTON" Then
                label = rs_chp("forec_label").Value
                If label = "" Then
                    label = rs_chp("forec_nom").Value
                End If
                sval = recup_valeur_champ(rs_don(rs_chp("forec_nom").Value).Value, _
                                          rs_chp("forec_type").Value)
                If sval <> "" Then
                    'Affichage dans le richTextBox
                    r_rtxt = r_rtxt & label & "|" & sval & vbLf
                End If
            End If
            rs_chp.MoveNext
        Wend
        rs_chp.Close
lab_etape_suiv:
    Next ietape
    
    ' Etape à faire
    If v_numetape > 0 Then
        sql = "select fore_lib from formetape where fore_fornum=" & v_numfor _
              & " and fore_numetape=" & v_numetape
        Call Odbc_RecupVal(sql, lib)
        
        r_rtxt = r_rtxt & lib & "à effectuer|" & vbLf

        nomutil = ""
        sql = "select donec_modif, donec_datefin, u_nom, u_prenom from donnees_encours, utilisateur" _
              & " where donec_donnum=" & v_numdon _
              & " and donec_fornum=" & v_numfor _
              & " and u_num=donec_unum"
        Call Odbc_SelectV(sql, rs)
        fmodif = False
        While Not rs.EOF
            If Not IsNull(rs("donec_datefin").Value) Then
                sdatfin = rs("donec_datefin").Value
            End If
            If nomutil <> "" Then
                nomutil = nomutil & " / "
            End If
            If rs("u_prenom").Value <> "" Then
                nomutil = nomutil & Left$(rs("u_prenom").Value, 1) & ". "
            End If
            nomutil = nomutil & rs("u_nom").Value
            fmodif = rs("donec_modif").Value
            rs.MoveNext
        Wend
        rs.Close
        If sdatfin <> "" Then
            r_rtxt = r_rtxt & " pour le |" & Format(sdatfin, "dd/mm/yyyy") & vbLf
        End If
        r_rtxt = r_rtxt & " par |" & nomutil & vbLf

        ' Il y a déjà des données enregistrées
        If fmodif Then
            ' Champs de l'étape
            sql = "select * from formetapechp where forec_fornum=" & v_numfor _
                  & " and forec_numetape=" & v_numetape _
                  & " and forec_nom<>'don_nomutil'" _
                  & " and forec_type<>'METHODE'" _
                  & " order by forec_ordre"
            Call Odbc_SelectV(sql, rs_chp)
            While Not rs_chp.EOF
                label = rs_chp("forec_label").Value
                If label = "" Then
                    label = rs_chp("forec_nom").Value
                End If
                sval = recup_valeur_champ(rs_don(rs_chp("forec_nom").Value).Value, _
                                          rs_chp("forec_type").Value)
                If sval <> "" Then
                    r_rtxt = r_rtxt & label & "|" & sval & vbLf
                End If
                rs_chp.MoveNext
            Wend
            rs_chp.Close
        End If
    End If
    rs_don.Close
    
    recup_detail = P_OK
    
End Function

'***************************************************
'#O Recuperation des valeurs des champs du formulaires
'#D Type : Private
'#D Paramètre : v_sval Pour le titre du champ
'#D > v_stypchp pour le type du champ
'#D Retourne : Les valeurs du champ
'***************************************************

Public Function recup_valeur_champ(ByVal v_sval As Variant, _
                                     ByVal v_stypchp As String) As Variant

    Dim s As String, sql As String
    Dim n As Integer, i As Integer
    Dim rs As rdoResultset
    Dim sval As Variant
    Dim num As Long
    
    If v_stypchp = "SELECT" Or v_stypchp = "RADIO" Or v_stypchp = "CHECK" Then
        If v_sval <> "" Then
            sval = ""
            n = STR_GetNbchamp(v_sval, ";")
            For i = 0 To n - 1
                num = Mid$(STR_GetChamp(v_sval, ";", i), 2)
                If num > 0 Then
                
                    sql = "select vc_lib, vc_libcourt from valchp where vc_num=" & num
                    
                    Call Odbc_SelectV(sql, rs)
                    If Not rs.EOF Then
                        If rs("vc_libcourt").Value <> "" Then
                            s = rs("vc_libcourt").Value
                        Else
                            s = rs("vc_lib").Value
                        End If
                    Else
                        s = "???"
                    End If

                    rs.Close
                    
                    If sval <> "" Then
                        sval = sval & " + "
                    End If
                    sval = sval & s
                End If
            Next i
            recup_valeur_champ = sval
        Else
            recup_valeur_champ = ""
        End If
        Exit Function
    ElseIf v_stypchp = "TEXTAREA" Then
        If Not IsNull(v_sval) Then
            recup_valeur_champ = Replace(v_sval, "$$", vbCrLf)
        End If
        Exit Function
    ElseIf v_stypchp = "TEXT" Then
        If InStr(1, v_sval, "#") > 0 Then
            recup_valeur_champ = STR_GetChamp(v_sval, "#", 1)
        Else
            recup_valeur_champ = v_sval
        End If
    Else
        recup_valeur_champ = v_sval
        Exit Function
    End If

End Function

'****************************************
'Verification de nouvelles actions
'****************************************
Private Function verif_actions(ByVal v_lnbact As Long) As Integer

    Dim sql As String
    Dim lnb As Long
    
    'Si nouvelles actions
    If v_lnbact > 0 Then
        sql = "SELECT Count(*) " _
            & "FROM DocAction " _
            & "WHERE DAC_Unum=" & p_numUtil & " AND DAC_ActionVu='f'"
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            verif_actions = P_ERREUR
            Exit Function
        End If
    End If

    'Vérifie si il y a une nouveautés
    If lnb > STR_GetChamp(lblStats(INDIC_ACTION).Caption, "/", 0) And g_mode_saisie Then
        If Not Me.Visible Then
            Call P_affiche_new
        End If
        
        Call pctPlus_Click(INDIC_ACTION)
        
    End If
    
    'Affichage
    If v_lnbact > 0 Then
        pctStats(INDIC_ACTION).Visible = True
        lblStats(INDIC_ACTION).Caption = lnb & " / " & v_lnbact
    Else
        pctStats(INDIC_ACTION).Visible = False
    End If
    
    verif_actions = P_OK

End Function


'****************************************
'Verification des AR
'****************************************
Private Function verif_AR(ByVal v_lnbar As Long) As Integer

    Dim sql As String
    Dim lnb As Long
    
    'Si nouvelles actions
    If v_lnbar > 0 Then
        sql = "SELECT Count(*) " _
            & "FROM DocDiffusion " _
            & "WHERE DD_Unum=" & p_numUtil & " AND DD_DiffusionVu='f' AND DD_ARecuperer=false"
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            verif_AR = P_ERREUR
            Exit Function
        End If
    End If

    'Vérifie si il y a une nouveautés
    If lnb > STR_GetChamp(lblStats(INDIC_AR).Caption, "/", 0) And g_mode_saisie Then
        If Not Me.Visible Then
            Call P_affiche_new
        End If
        Call pctPlus_Click(INDIC_AR)
    End If

    'Affichage
    If v_lnbar > 0 Then
        pctStats(INDIC_AR).Visible = True
        lblStats(INDIC_AR).Caption = lnb & " / " & v_lnbar
    Else
        pctStats(INDIC_AR).Visible = False
    End If
    
    verif_AR = P_OK

End Function

'****************************************
'Vérification des classeurs
'****************************************
Private Function verif_Classeurs() As Integer

    Dim sql As String
    Dim rs As rdoResultset
    Dim lnb As Long
    Dim bclasseur  As Boolean
    Dim rs2 As rdoResultset
    
    bclasseur = False
    
    sql = "SELECT CLSR_CLSNum, CLSR_CLSVNum " _
        & "FROM ClasseurReferent " _
        & "WHERE CLSR_UNum=" & p_numUtil
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        verif_Classeurs = P_ERREUR
        Exit Function
    End If

    If rs.EOF Then
        pctStats(INDIC_CLASS).Visible = False
        verif_Classeurs = P_OK
        Exit Function
    End If

    pctStats(INDIC_CLASS).Visible = True
    While Not rs.EOF
    
        If Odbc_SelectV("SELECT MAX(CLSV_Num) " _
                        & "FROM ClasseurVersion " _
                        & "WHERE CLSV_CLSNum=" & rs("CLSR_CLSNum").Value, rs2) = P_ERREUR Then
            rs.Close
            verif_Classeurs = P_ERREUR
            Exit Function
        End If
    
        If rs2(0).Value > rs("CLSR_CLSVNum").Value Then
            bclasseur = True
            GoTo affich
        End If
        
        rs2.Close
        rs.MoveNext
    Wend

affich:
    rs.Close
    
    If bclasseur Then
        lblStats(INDIC_CLASS).Caption = "Mise à jour"
    Else
        lblStats(INDIC_CLASS).Caption = "OK"
    End If

    verif_Classeurs = P_OK
    
End Function

'****************************************
'Verification des demandes d'informations
'****************************************
Private Function verif_infos(ByVal v_lnbinfo As Long) As Integer

    Dim sql As String
    Dim lnb As Long
    
    'Si nouvelles actions
    If v_lnbinfo > 0 Then
        sql = "SELECT Count(*) " _
            & "FROM DocInfo " _
            & "WHERE DI_Unum=" & p_numUtil & " AND DI_InfoVu='f'"
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            verif_infos = P_ERREUR
            Exit Function
        End If
    End If

    'Vérifie si il y a une nouveautés
    If lnb > STR_GetChamp(lblStats(INDIC_INFO).Caption, "/", 0) And g_mode_saisie Then
        If Not Me.Visible Then
            Call P_affiche_new
        End If
        Call pctPlus_Click(INDIC_INFO)
    End If
    
    'Affichage
    If v_lnbinfo > 0 Then
        lblStats(INDIC_INFO).Caption = lnb & " / " & v_lnbinfo
        pctStats(INDIC_INFO).Visible = True
    Else
        pctStats(INDIC_INFO).Visible = False
    End If
    
    verif_infos = P_OK

End Function

'********************************************
'Verification de la réception de formulaires
'********************************************
Private Function verif_KaliForm() As Integer

    Dim sql As String
    Dim lnbtot As Long, lnbnew As Long
    Dim rs As rdoResultset

    sql = "SELECT * FROM UtilAdim WHERE UA_DONum=0 AND UA_UNum=" & p_numUtil
    
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        verif_KaliForm = P_ERREUR
        Exit Function
    End If
    
    If Not rs.EOF Then
        lnbtot = rs("UA_NbForm").Value
                
        If rs("UA_NewForm").Value Then
            
            sql = "SELECT Count(*) FROM Donnees_Encours WHERE DONEC_Vu='f' AND DONEC_Unum=" & p_numUtil
            
            If Odbc_Count(sql, lnbnew) = P_ERREUR Then
                verif_KaliForm = P_ERREUR
                Exit Function
            End If
            
        End If
            
    End If
    rs.Close
    
    'Vérifie si il y a une nouveautés
    If lnbnew > STR_GetChamp(lblStats(INDIC_FORM).Caption, "/", 0) And g_mode_saisie Then
        If Not Me.Visible Then
            Call P_affiche_new
        End If
        pctPlus_Click (INDIC_FORM)
    End If
    
    If lnbtot > 0 Then
        lblStats(INDIC_FORM).Caption = lnbnew & " / " & lnbtot
        pctStats(INDIC_FORM).Visible = True
    Else
        pctStats(INDIC_FORM).Visible = False
    End If
    
    verif_KaliForm = P_OK

End Function

'****************************************
'Verification des KaliMails
'****************************************
Private Function verif_KaliMail() As Integer

    Dim sql As String
    Dim lnbnew As Long, lnbtot As Long
    
    'Nombre de KaliMail
    sql = "SELECT Count(*) " _
        & "FROM KaliMailDetail " _
        & "WHERE KMD_SuppDest=0 AND KMD_DestNum=" & p_numUtil
    If Odbc_Count(sql, lnbtot) = P_ERREUR Then
        verif_KaliMail = P_ERREUR
        Exit Function
    End If
    
    'Nombre de nouveaux mails
    sql = "SELECT Count(*) " _
        & "FROM KaliMailDetail " _
        & "WHERE KMD_SuppDest=0 AND KMD_DestNum=" & p_numUtil _
        & " AND KMD_DateLu is null"
    If Odbc_Count(sql, lnbnew) = P_ERREUR Then
        verif_KaliMail = P_ERREUR
        Exit Function
    End If
    
    'Vérifie si il y a une nouveautés
    If lnbnew > STR_GetChamp(lblStats(INDIC_KMAIL).Caption, "/", 0) And g_mode_saisie Then
        If Not Me.Visible Then
            Call P_affiche_new
        End If
        Call pctPlus_Click(INDIC_KMAIL)
    End If
    
    If lnbtot > 0 Then
        lblStats(INDIC_KMAIL).Caption = lnbnew & " / " & lnbtot
        pctStats(INDIC_KMAIL).Visible = True
    Else
        pctStats(INDIC_KMAIL).Visible = False
    End If
    
    verif_KaliMail = P_OK

End Function

'****************************************
'Vérification des déclarations de KGDR
'****************************************
Private Function verif_KGDR() As Integer

    Dim lnb As Long, lnbt As Long, lnbnew As Long, lnbtot As Long
    Dim sql As String
    Dim rs As rdoResultset

    'Vérification des droits
    sql = "SELECT Count(*) FROM FctOk_Util WHERE FU_Unum=" & p_numUtil & " AND FU_FCTNum=(SELECT FCT_Num FROM Fonction WHERE FCT_Code='GDR_QUALIF')"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        verif_KGDR = P_ERREUR
        Exit Function
    End If

    If lnb > 0 Then
    
        If Odbc_SelectV("SELECT DMR_Num FROM DomaineRisk", rs) = P_ERREUR Then
            verif_KGDR = P_ERREUR
            Exit Function
        End If
        
        'Pas de domaine
        If rs.EOF Then
            pctStats(INDIC_KGDR).Visible = False
            verif_KGDR = P_OK
            Exit Function
        End If
        
        While Not rs.EOF
        
            sql = "SELECT Count(*) FROM GestionRisk_" & rs("DMR_Num").Value & " WHERE GDR_Etat<2"
            If Odbc_Count(sql, lnbt) = P_ERREUR Then
                verif_KGDR = P_ERREUR
                Exit Function
            End If
        
            sql = "SELECT Count(*) FROM GestionRisk_" & rs("DMR_Num").Value & " WHERE GDR_Etat=0"
            If Odbc_Count(sql, lnb) = P_ERREUR Then
                verif_KGDR = P_ERREUR
                Exit Function
            End If
        
            lnbtot = lnbtot + lnbt
            lnbnew = lnbnew + lnb
        
            rs.MoveNext
        Wend
        rs.Close
   
       'Vérifie si il y a une nouveautés
        If lnbnew > STR_GetChamp(lblStats(INDIC_KGDR).Caption, "/", 0) And g_mode_saisie Then
            If Not Me.Visible Then
                Call P_affiche_new
            End If
            Call pctPlus_Click(INDIC_KGDR)
        End If
   
        lblStats(INDIC_KGDR).Caption = lnbnew & " / " & lnbtot
        pctStats(INDIC_KGDR).Visible = True
    Else
        pctStats(INDIC_KGDR).Visible = False
    End If

    verif_KGDR = P_OK

End Function

'*******************************************
'Verification des demandes de modifications
'*******************************************
Private Function verif_modifications(ByVal v_lnbmod As Long) As Integer

    Dim sql As String
    Dim lnb As Long
    
    'Si nouvelles actions
    If v_lnbmod > 0 Then
        sql = "SELECT Count(*) " _
            & "FROM DocDemModif " _
            & "WHERE DDM_Unum=" & p_numUtil & " AND DDM_ModifVu='f'"
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            verif_modifications = P_ERREUR
            Exit Function
        End If
    End If

    'Affichage
    If v_lnbmod > 0 Then
    
        'Vérifie si il y a une nouveautés
        If lnb > STR_GetChamp(lblStats(INDIC_MODIF).Caption, "/", 0) And g_mode_saisie Then
            If Not Me.Visible Then
                Call P_affiche_new
            End If
            Call pctPlus_Click(INDIC_MODIF)
        End If
    
        lblStats(INDIC_MODIF).Caption = lnb & " / " & v_lnbmod
        pctStats(INDIC_MODIF).Visible = True
    Else
        pctStats(INDIC_MODIF).Visible = False
    End If
    
    verif_modifications = P_OK

End Function

Private Sub cmd_Click(Index As Integer)
    
    Dim sshel As String
    Dim snumutil As String
    
    Dim nblig As Long
    Dim i As Long
    Dim lig As String
    
    'Bouton de navigation
    If Index < CMD_DET Then
        
        'Déplacer le cursor
        If Index = CMD_PREC Then
            g_position = g_position - 1
        Else
            g_position = g_position + 1
        End If
        
        'Affiche le texte
        rtxt.Text = ""
        nblig = STR_GetNbchamp(p_indic(g_position).stxt, vbLf)
        For i = 0 To nblig - 1
            lig = STR_GetChamp(p_indic(g_position).stxt, vbLf, i)
            
            rtxt.SelStart = Len(rtxt.Text)
            rtxt.SelColor = vbBlue
            
            If (p_indic(g_position).stype = "kgdr" Or p_indic(g_position).stype = "kform") And i > 0 Then
                rtxt.SelText = STR_GetChamp(lig, "|", 0) & " : "
            Else
                rtxt.SelText = STR_GetChamp(lig, "|", 0)
            End If
            
            rtxt.SelStart = Len(rtxt.Text)
            rtxt.SelColor = vbBlack
            rtxt.SelText = STR_GetChamp(lig, "|", 1)
            
            rtxt.SelStart = Len(rtxt.Text)
            rtxt.SelText = vbLf
        Next i
    
        rtxt.SelStart = 0
        
        If p_indic(g_position).bvu Then
            imgNew.Visible = False
        Else
            imgNew.Visible = True
        End If
    
        'MAJ des bouttons
        If g_position = UBound(p_indic) Then
            cmd(CMD_SUIV).Visible = False
        Else
            cmd(CMD_SUIV).Visible = True
        End If
        
        If g_position = 0 Then
            cmd(CMD_PREC).Visible = False
        Else
            cmd(CMD_PREC).Visible = True
        End If
    
        lblNb.Caption = g_position + 1 & " / " & p_nbindic
    
    'Afficher / Masquer la partie du résumé
    ElseIf Index = CMD_UP Then
        
        If Me.Height > frmDet.Height + frmStats(0).Height Then
            Me.Height = frmStats(0).Top + frmStats(0).Height + 70
            Set cmd(CMD_UP).Picture = imglistMask.ListImages(IMG_BAS).Picture
            cmd(CMD_UP).ToolTipText = "Afficher le résumé"
        Else
            Me.Height = frmDet.Top + frmDet.Height + 70
            Set cmd(CMD_UP).Picture = imglistMask.ListImages(IMG_HAUT).Picture
            cmd(CMD_UP).ToolTipText = "Masquer le résumé"
        End If
    
    'Accès au détail
    Else
    
        snumutil = STR_CrypterNombre(Format(p_numUtil, "#0000000"))
        sshel = "V_util=" & snumutil & "&" & p_indic(g_position).surl
        Shell "C:\Program Files\Internet Explorer\iexplore.exe " & p_cheminphp & "/pident.php?" & sshel, vbMaximizedFocus
    
    End If

End Sub

Private Sub Form_Load()

    g_mode_saisie = False

    Me.Top = Screen.Height - Me.Height - 1000
    Me.Left = Screen.Width - Me.Width
    
   ' Timer1.Enabled = True

    Call initialiser
    
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call quitter
    
End Sub

Private Sub pctPlus_Click(Index As Integer)

    'Affichage
    frmPatience.Visible = True
    frmPatience.ZOrder 0
    Me.Refresh

    'Réinitialisation des objets
    g_position = 0
    rtxt.Text = ""
    
    p_nbindic = 0
    Erase p_indic
    
    'Chargement des objets
    Select Case Index
        
        Case INDIC_ACTION
            If charger_actions = P_ERREUR Then
                Exit Sub
            End If
            lblDet.Caption = "Résumé des actions : "
            Set Picture4.Picture = ImageList.ListImages(IMG_ACTION).Picture
        Case INDIC_MODIF
            If charger_modif = P_ERREUR Then
                Exit Sub
            End If
            lblDet.Caption = "Résumé des Demandes :"
            Set Picture4.Picture = ImageList.ListImages(IMG_MODIF).Picture
        Case INDIC_AR
            If charger_ar = P_ERREUR Then
                Exit Sub
            End If
            lblDet.Caption = "Résumé des AR :"
            Set Picture4.Picture = ImageList.ListImages(IMG_AR).Picture
        Case INDIC_INFO
            If charger_info = P_ERREUR Then
                Exit Sub
            End If
            lblDet.Caption = "Résumé des informations :"
            Set Picture4.Picture = ImageList.ListImages(IMG_INFO).Picture
        Case INDIC_KMAIL
            If charger_kmail = P_ERREUR Then
                Exit Sub
            End If
            lblDet.Caption = "KaliMails :"
            Set Picture4.Picture = ImageList.ListImages(IMG_KMAIL).Picture
        Case INDIC_CLASS
            'Lance les classeurs
            
        Case INDIC_FORM
            If charger_kform = P_ERREUR Then
                Exit Sub
            End If
            lblDet.Caption = "Formulaires reçus :"
            Set Picture4.Picture = ImageList.ListImages(IMG_KFORM).Picture
        Case INDIC_KGDR
            If charger_kgdr = P_ERREUR Then
                Exit Sub
            End If
            lblDet.Caption = "Fiche d'évènement indésirable :"
            Set Picture4.Picture = Nothing
    End Select
    
    If p_nbindic > 0 Then
        lblNb.Caption = "1 / " & p_nbindic
        cmd(CMD_PREC).Visible = False
        cmd(CMD_SUIV).Visible = True
        cmd(CMD_DET).Enabled = True
        g_position = 1
        Call cmd_Click(CMD_PREC)
    Else
        lblNb.Caption = "0 / 0"
        cmd(CMD_PREC).Visible = False
        cmd(CMD_SUIV).Visible = False
        cmd(CMD_DET).Enabled = False
    End If
            
    Me.Height = frmDet.Top + frmDet.Height + 70
    Set cmd(CMD_UP).Picture = imglistMask.ListImages(IMG_HAUT).Picture
    cmd(CMD_UP).ToolTipText = "Masquer le résumé"
            
    frmPatience.Visible = False
            
End Sub

Private Sub Picture2_Click()

    Call quitter

End Sub

Private Sub Timer1_Timer()

    Dim bnew As Boolean
    Dim frm As Form

    If g_nbtimer = 0 Or g_nbtimer = 10 Then
        If evalue_action(bnew) = P_ERREUR Then
            Call MsgBox("Une erreur est survenue lors de l'évaluation des alertes.", vbExclamation + vbOKOnly, "KaliAlerte")
            Call quitter
            Exit Sub
        End If
        
        g_nbtimer = 1
        
        Label2.Caption = "Dernière Mise à jour : " & Time
    Else
        g_nbtimer = g_nbtimer + 1
    End If
    
    Timer1.Interval = 60000
   

End Sub

Private Sub Timer2_Timer()

    Dim i As Integer
    
    p_slstaction = ""
    Timer2.Enabled = False
    Timer1.Enabled = True
    g_position = 0

End Sub

Private Sub Timer3_Timer()

    If nid.hIcon = Me.Icon Then
    '    Shell_NotifyIcon NIM_DELETE, nid
        nid.hIcon = KA_PrmAlerte.Icon
        Shell_NotifyIcon NIM_MODIFY, nid
    Else
   '     Shell_NotifyIcon NIM_DELETE, nid
        nid.hIcon = Me.Icon
        Shell_NotifyIcon NIM_MODIFY, nid
    End If

End Sub

