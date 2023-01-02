VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form PrmGeneral 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Paramétrage du fichier d'importation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   7395
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   60
      Width           =   10935
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   10440
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Prochaine ligne"
         Top             =   2280
         Width           =   375
      End
      Begin VB.Frame FramePosit 
         BackColor       =   &H00C0C0C0&
         Height          =   1815
         Left            =   120
         TabIndex        =   44
         Top             =   2880
         Width           =   10695
         Begin VB.TextBox txtF 
            Height          =   285
            Index           =   8
            Left            =   9600
            TabIndex        =   71
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtD 
            Height          =   285
            Index           =   8
            Left            =   8880
            TabIndex        =   70
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtF 
            Height          =   285
            Index           =   7
            Left            =   9600
            TabIndex        =   69
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtD 
            Height          =   285
            Index           =   7
            Left            =   8880
            TabIndex        =   68
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtF 
            Height          =   285
            Index           =   6
            Left            =   9600
            TabIndex        =   67
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtD 
            Height          =   285
            Index           =   6
            Left            =   8880
            TabIndex        =   66
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtF 
            Height          =   285
            Index           =   5
            Left            =   5760
            TabIndex        =   65
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtD 
            Height          =   285
            Index           =   5
            Left            =   5040
            TabIndex        =   64
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtF 
            Height          =   285
            Index           =   4
            Left            =   5760
            TabIndex        =   63
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtD 
            Height          =   285
            Index           =   4
            Left            =   5040
            TabIndex        =   62
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtF 
            Height          =   285
            Index           =   3
            Left            =   5760
            TabIndex        =   61
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtD 
            Height          =   285
            Index           =   3
            Left            =   5040
            TabIndex        =   60
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtF 
            Height          =   285
            Index           =   2
            Left            =   1800
            TabIndex        =   59
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtD 
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   58
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtF 
            Height          =   285
            Index           =   1
            Left            =   1800
            TabIndex        =   57
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtD 
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   56
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtF 
            Height          =   285
            Index           =   0
            Left            =   1800
            TabIndex        =   55
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtD 
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   54
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nom de JF"
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
            Index           =   25
            Left            =   7560
            TabIndex        =   53
            Top             =   1215
            Width           =   1290
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Civilité"
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
            Index           =   24
            Left            =   3600
            TabIndex        =   52
            Top             =   1230
            Width           =   1335
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nom"
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
            Index           =   23
            Left            =   240
            TabIndex        =   51
            Top             =   285
            Width           =   615
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Prénom"
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
            Index           =   22
            Left            =   240
            TabIndex        =   50
            Top             =   735
            Width           =   855
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Matricule"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   21
            Left            =   240
            TabIndex        =   49
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Code Service"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   20
            Left            =   3600
            TabIndex        =   48
            Top             =   375
            Width           =   1335
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Libellé Service"
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
            Index           =   19
            Left            =   3600
            TabIndex        =   47
            Top             =   825
            Width           =   1335
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Code Poste"
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
            Index           =   18
            Left            =   7560
            TabIndex        =   46
            Top             =   345
            Width           =   1095
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Libellé Poste"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   17
            Left            =   7560
            TabIndex        =   45
            Top             =   825
            Width           =   1215
         End
      End
      Begin VB.Frame FrameSepar 
         BackColor       =   &H00C0C0C0&
         Height          =   1815
         Left            =   120
         TabIndex        =   25
         Top             =   2880
         Width           =   10695
         Begin VB.ComboBox combox 
            Height          =   315
            Index           =   6
            Left            =   8760
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   300
            Width           =   1695
         End
         Begin VB.ComboBox combox 
            Height          =   315
            Index           =   7
            Left            =   8760
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   780
            Width           =   1695
         End
         Begin VB.ComboBox combox 
            Height          =   315
            Index           =   1
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox combox 
            Height          =   315
            Index           =   2
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   690
            Width           =   1695
         End
         Begin VB.ComboBox combox 
            Height          =   315
            Index           =   3
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1170
            Width           =   1695
         End
         Begin VB.ComboBox combox 
            Height          =   315
            Index           =   4
            Left            =   4920
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   330
            Width           =   1695
         End
         Begin VB.ComboBox combox 
            Height          =   315
            Index           =   5
            Left            =   4920
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   780
            Width           =   1695
         End
         Begin VB.ComboBox combox 
            Height          =   315
            Index           =   8
            Left            =   4920
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1230
            Width           =   1695
         End
         Begin VB.ComboBox combox 
            Height          =   315
            Index           =   9
            Left            =   8790
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1215
            Width           =   1695
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Libellé Poste"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   7560
            TabIndex        =   43
            Top             =   825
            Width           =   1215
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Code Poste"
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
            Index           =   8
            Left            =   7560
            TabIndex        =   42
            Top             =   345
            Width           =   1095
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Libellé Service"
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
            Left            =   3600
            TabIndex        =   41
            Top             =   825
            Width           =   1335
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Code Service"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   3600
            TabIndex        =   40
            Top             =   375
            Width           =   1335
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Matricule"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   240
            TabIndex        =   39
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Prénom"
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
            Left            =   240
            TabIndex        =   38
            Top             =   735
            Width           =   855
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nom"
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
            TabIndex        =   37
            Top             =   285
            Width           =   615
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Civilité"
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
            Index           =   14
            Left            =   3600
            TabIndex        =   36
            Top             =   1230
            Width           =   1335
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nom de JF"
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
            Index           =   15
            Left            =   7560
            TabIndex        =   35
            Top             =   1215
            Width           =   1290
         End
      End
      Begin VB.OptionButton optEmpl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sur le serveur"
         Height          =   195
         Index           =   1
         Left            =   2850
         TabIndex        =   22
         Top             =   930
         Width           =   1305
      End
      Begin VB.OptionButton optEmpl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "En local"
         Height          =   195
         Index           =   0
         Left            =   1710
         TabIndex        =   21
         Top             =   930
         Width           =   1305
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   3
         Left            =   5265
         MaxLength       =   15
         TabIndex        =   4
         Top             =   7020
         Width           =   3165
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   2
         Left            =   810
         MaxLength       =   15
         TabIndex        =   3
         ToolTipText     =   $"PrmGeneral.frx":0000
         Top             =   7005
         Width           =   2805
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   10140
         Picture         =   "PrmGeneral.frx":0098
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   6240
         Width           =   375
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   10110
         Picture         =   "PrmGeneral.frx":04DF
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5280
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   1335
         Left            =   240
         TabIndex        =   14
         Top             =   5280
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   2355
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   8454143
         BackColorBkg    =   16777215
      End
      Begin VB.TextBox txt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   9030
         TabIndex        =   1
         Text            =   "#"
         ToolTipText     =   "indiquez un caractère séparateur ou TAB pour tabulation"
         Top             =   450
         Width           =   885
      End
      Begin VB.ComboBox combox 
         Height          =   315
         Index           =   0
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   4770
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Parcourir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   9780
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   1
         Left            =   3030
         TabIndex        =   2
         Text            =   "< Veuillez indiquer le chemin du fichier d'importation >"
         Top             =   1320
         Width           =   6615
      End
      Begin VB.Label lblLig 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   72
         Top             =   2280
         Visible         =   0   'False
         Width           =   10215
      End
      Begin VB.Label LblTest 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4320
         TabIndex        =   24
         Top             =   1920
         Visible         =   0   'False
         Width           =   6495
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Emplacement"
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
         Index           =   16
         Left            =   240
         TabIndex        =   23
         Top             =   930
         Width           =   1575
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Création de compte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Index           =   13
         Left            =   465
         TabIndex        =   20
         Top             =   6660
         Width           =   4935
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mot de passe"
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
         Index           =   12
         Left            =   3945
         TabIndex        =   19
         Top             =   7035
         Width           =   1290
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Code"
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
         Index           =   11
         Left            =   195
         TabIndex        =   18
         Top             =   7050
         Width           =   630
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Types d'informations supplémentaires"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   270
         TabIndex        =   17
         Top             =   4950
         Width           =   3735
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Type de fichier"
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
         TabIndex        =   13
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Caractère séparateur"
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
         Left            =   7110
         TabIndex        =   12
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Position des champs dans la ligne"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   1890
         Width           =   4935
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Chemin du fichier d'importation"
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
         Index           =   10
         Left            =   240
         TabIndex        =   10
         Top             =   1380
         Width           =   2655
      End
   End
   Begin VB.Frame frm 
      BackColor       =   &H00808080&
      Height          =   860
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   7440
      Width           =   10935
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "PrmGeneral.frx":0936
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   0
         Left            =   840
         Picture         =   "PrmGeneral.frx":0E92
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Enregistrer les modifications et quitter."
         Top             =   230
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   1
         Left            =   9480
         Picture         =   "PrmGeneral.frx":13FB
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Quitter sans enregistrer les modofication."
         Top             =   230
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "PrmGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Index des TXT
Private Const LBL_SEPARATEUR = 1
Private Const TXT_SEPARATEUR = 0
Private Const TXT_CHEMIN = 1
Private Const TXT_CODE = 2
Private Const TXT_MPASSE = 3

' Index des ComboBox
Private Const COMBO_TYPE_FICH = 0
Private Const COMBO_NOM = 1
Private Const COMBO_PRENOM = 2
Private Const COMBO_MATRICULE = 3
Private Const COMBO_CODE_SECTION = 4
Private Const COMBO_LIB_SECTION = 5
Private Const COMBO_CODE_POSTE = 6
Private Const COMBO_LIB_POSTE = 7
Private Const COMBO_CIVILITE = 8
Private Const COMBO_NJF = 9

' Index des Opt
Private Const OPT_EMPL_LOCAL = 0
Private Const OPT_EMPL_SERVEUR = 1

' Index des CMD
Private Const CMD_OK = 0
Private Const CMD_QUITTER = 1
Private Const CMD_PARCOURIR = 2
' Type d'infos supplemntaires
Private Const CMD_PLUS_TIS = 3
Private Const CMD_MOINS_TIS = 4
Private Const CMD_NEXT_LIGNE = 5

' les colonnes du grid TypeInfoSuppl
Private Const GRDC_TIS_NUM = 0
Private Const GRDC_TIS_LIBELLE = 1
Private Const GRDC_TIS_POS = 2
Private Const GRDC_TIS_LONG = 3

' Nombre de caractères de séparation
Private Const g_nbr_car_sep_max = 1

' L'enregistrement existe
Private g_enreg_existe As Boolean
Private g_faire_combo_click As Boolean
Private g_mode_saisie As Boolean
Private g_form_active As Boolean
Private g_quitter As Boolean
Private g_txt_par_defaut As String
Private g_ligne1 As String
Private g_ligne2 As String
Private g_ligne3 As String

Public Sub AppelFrm()

    Call FRM_ResizeForm(Me, Me.width, Me.Height)

    Me.Show 1

End Sub

Private Sub ajouter_typeinfosuppl()
' *******************************************
' Ajouter un type d'infos suppl
' *******************************************
    Dim sql As String
    Dim rs As rdoResultset
    Dim nbr_tis_a_afficher As Integer, I As Integer

    nbr_tis_a_afficher = 0
    
lab_affiche:
    Call CL_Init
    'Choix du TIS
    sql = "SELECT * FROM KB_TypeInfoSuppl WHERE KB_TisImport='t'" _
        & " ORDER BY KB_TisLibelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    'Call CL_AddLigne("<Nouveau>", 0, "", False)

    While Not rs.EOF
        ' est-ce qu'on a déjà ce TIS ?
        With grd
            For I = 1 To .Rows - 1
                If rs("KB_TisNum").Value = .TextMatrix(I, GRDC_TIS_NUM) Then
                    GoTo lab_suivant
                End If
            Next I
        End With
        Call CL_AddLigne(rs("KB_TisLibelle").Value, rs("KB_TisNum").Value, "", False)
        nbr_tis_a_afficher = nbr_tis_a_afficher + 1
lab_suivant:
        rs.MoveNext
    Wend
    rs.Close

    If nbr_tis_a_afficher = 0 Then
        MsgBox "Il n'y a pas d'autres types d'informations supplémentaires." _
            & vbCrLf & vbCrLf & "Vous pouvez en ajouter en passant dans Dictionnaire - Types d'informations supplémentaires.", _
            vbInformation + vbOKOnly, "Liste vide"
        GoTo lab_erreur
    End If

    Call FRM_ResizeForm(Me, 0, 0)

    Call CL_InitTitreHelp("Liste des types de coordonnée", p_chemin_appli + "\help\kalidoc.chm" & ";" & "dico_d_fonction.htm")
    Call CL_InitTaille(0, -15)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
'    Call CL_AddBouton("", p_chemin_appli + "\btnimprimer.gif", vbKeyI, vbKeyF3, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)

    ChoixListe.Show 1

    ' Quitter
    If CL_liste.retour = 1 Then
        GoTo lab_erreur
    End If

'    ' Imprimer
'    If CL_liste.retour = 1 Then
'        Call imprimer
'        GoTo lab_affiche
'    End If

    ' => CL_liste.retour = 0
    With grd
        .AddItem ""
        .Row = .Rows - 1
        .TextMatrix(.Rows - 1, GRDC_TIS_NUM) = CL_liste.lignes(CL_liste.pointeur).num
        .TextMatrix(.Rows - 1, GRDC_TIS_LIBELLE) = CL_liste.lignes(CL_liste.pointeur).texte
        .TextMatrix(.Rows - 1, GRDC_TIS_POS) = 0
        cmd(CMD_OK).Enabled = True
    End With

lab_erreur:
    Call FRM_ResizeForm(Me, Me.width, Me.Height)

End Sub

Private Sub enlever_typeinfosuppl()
' *******************************************
' Enlever le type d'infos suppl
' *******************************************
    With grd
        If .Rows = 1 Then Exit Sub
        If .Rows = 2 Then
            .Rows = 1
            Exit Sub
        End If
        .RemoveItem (.Row)
        cmd(CMD_OK).Enabled = True
    End With
End Sub

Private Sub enregistrer()
' ***********************************************
' Enregistrer les modifications une fois validées
' ***********************************************
    Dim tis_lstposinfoautre As String
    Dim fichsurserveur As Boolean
    Dim lbid As Long
    Dim letype As Integer
    
    If verifier_tout_champ = P_ERREUR Then Exit Sub

    ' recupérer la liste des TIS et de leurs positions
    tis_lstposinfoautre = get_tis_lstposinfoautre()

    If optEmpl(OPT_EMPL_LOCAL).Value = True Then
        fichsurserveur = False
    Else
        fichsurserveur = True
    End If
    
    If combox(COMBO_TYPE_FICH).ListIndex = 0 Then
        letype = 1
    Else
        letype = 2
    End If
    
    txt(TXT_CHEMIN).Text = Replace(txt(TXT_CHEMIN).Text, "\", "/")
    txt(TXT_CHEMIN).Text = Trim(txt(TXT_CHEMIN).Text)
    If g_enreg_existe Then
        If Odbc_Update("PrmGenB", "PGB_Num", _
                       "", _
                       "PGB_FichType", letype, _
                       "PGB_FichSep", txt(TXT_SEPARATEUR).Text, _
                       "PGB_FichLstPos", get_FichLstSep(), _
                       "PGB_Chemin", IIf(txt(TXT_CHEMIN).Text = g_txt_par_defaut, "", txt(TXT_CHEMIN).Text), _
                       "PGB_FichSurServeur", fichsurserveur, _
                       "PGB_Code", txt(TXT_CODE).Text, _
                       "PGB_mdp", UCase(txt(TXT_MPASSE).Text), _
                       "PGB_LstPosInfoAutre", tis_lstposinfoautre) = P_ERREUR Then
           Exit Sub
        End If
    Else
        Call Odbc_AddNew("PrmGenB", "PGB_Num", "pgb_seq", False, lbid, _
                       "PGB_FichType", letype, _
                       "PGB_FichSep", txt(TXT_SEPARATEUR).Text, _
                       "PGB_FichLstPos", get_FichLstSep(), _
                       "PGB_Chemin", IIf(txt(TXT_CHEMIN).Text = g_txt_par_defaut, "", txt(TXT_CHEMIN).Text), _
                       "PGB_FichSurServeur", fichsurserveur, _
                       "PGB_Code", txt(TXT_CODE).Text, _
                       "PGB_mdp", UCase(txt(TXT_MPASSE).Text), _
                       "PGB_LstPosInfoAutre", tis_lstposinfoautre)
    End If
    
    Call quitter(True)

End Sub

Private Function get_FichLstSep() As String
' Construire et retourner la liste des champs avec leurs positions dans le lignes du fichier d'importation
' exemple: "NOM=2;PRENOM=3;MATRICULE=0;CODE_SECTION=4;LIB_SECTION=5;CODE_FONCTION=6;LIB_FONCTION=7;LIB_CIVILITE=8"

    If combox(COMBO_TYPE_FICH).ListIndex = 0 Then   ' séparateur
        get_FichLstSep = "NOM=" & combox(COMBO_NOM).List(combox(COMBO_NOM).ListIndex) - 1 _
                       & ";PRENOM=" & combox(COMBO_PRENOM).List(combox(COMBO_PRENOM).ListIndex) - 1 _
                       & ";MATRICULE=" & combox(COMBO_MATRICULE).List(combox(COMBO_MATRICULE).ListIndex) - 1 _
                       & ";CODE_SECTION=" & combox(COMBO_CODE_SECTION).List(combox(COMBO_CODE_SECTION).ListIndex) - 1 _
                       & ";LIB_SECTION=" & combox(COMBO_LIB_SECTION).List(combox(COMBO_LIB_SECTION).ListIndex) - 1 _
                       & ";CODE_FONCTION=" & combox(COMBO_CODE_POSTE).List(combox(COMBO_CODE_POSTE).ListIndex) - 1 _
                       & ";LIB_FONCTION=" & combox(COMBO_LIB_POSTE).List(combox(COMBO_LIB_POSTE).ListIndex) - 1 _
                       & ";LIB_CIVILITE=" & combox(COMBO_CIVILITE).List(combox(COMBO_CIVILITE).ListIndex) - 1 _
                       & ";NJF=" & combox(COMBO_NJF).List(combox(COMBO_NJF).ListIndex) - 1 _
                       & ";"
    Else    ' positionnel
        get_FichLstSep = "NOM=" & txtD(0).Text & ":" & txtF(0).Text _
                       & ";PRENOM=" & txtD(1).Text & ":" & txtF(1).Text _
                       & ";MATRICULE=" & txtD(2).Text & ":" & txtF(2).Text _
                       & ";CODE_SECTION=" & txtD(3).Text & ":" & txtF(3).Text _
                       & ";LIB_SECTION=" & txtD(4).Text & ":" & txtF(4).Text _
                       & ";CODE_FONCTION=" & txtD(6).Text & ":" & txtF(6).Text _
                       & ";LIB_FONCTION=" & txtD(7).Text & ":" & txtF(7).Text _
                       & ";LIB_CIVILITE=" & txtD(5).Text & ":" & txtF(5).Text _
                       & ";NJF=" & txtD(8).Text & ":" & txtF(8).Text _
                       & ";"
    End If
End Function

Private Function get_tis_lstposinfoautre() As String

    Dim str As String
    Dim I As Integer

    str = ""

    With grd
        For I = 1 To .Rows - 1
            If combox(COMBO_TYPE_FICH).ListIndex = 0 Then
                str = str & "I" & .TextMatrix(I, GRDC_TIS_NUM) & ";" & .TextMatrix(I, GRDC_TIS_POS) & "|"
            Else
                str = str & "I" & .TextMatrix(I, GRDC_TIS_NUM) & ";" & .TextMatrix(I, GRDC_TIS_POS) & ";" & .TextMatrix(I, GRDC_TIS_LONG) & "|"
            End If
        Next I
    End With

    get_tis_lstposinfoautre = str

End Function

Private Sub initialiser()

    Dim liste(10) As String
    Dim I As Integer

    g_mode_saisie = False
    g_txt_par_defaut = txt(TXT_CHEMIN).Text

    g_faire_combo_click = False
    ' Initialisation des combobox
    liste(0) = "0"
    liste(1) = "1": liste(2) = "2": liste(3) = "3": liste(4) = "4": liste(5) = "5"
    liste(6) = "6": liste(7) = "7": liste(8) = "8": liste(9) = "9": liste(10) = "10"
    For I = 1 To 10
        combox(COMBO_NOM).AddItem Item:=liste(I), Index:=I - 1
        combox(COMBO_NOM).ItemData(I - 1) = I
        combox(COMBO_PRENOM).AddItem Item:=liste(I), Index:=I - 1
        combox(COMBO_PRENOM).ItemData(I - 1) = I
        combox(COMBO_MATRICULE).AddItem Item:=liste(I), Index:=I - 1
        combox(COMBO_MATRICULE).ItemData(I - 1) = I
        combox(COMBO_CODE_SECTION).AddItem Item:=liste(I), Index:=I - 1
        combox(COMBO_CODE_SECTION).ItemData(I - 1) = I
        combox(COMBO_LIB_SECTION).AddItem Item:=liste(I), Index:=I - 1
        combox(COMBO_LIB_SECTION).ItemData(I - 1) = I
        combox(COMBO_CODE_POSTE).AddItem Item:=liste(I), Index:=I - 1
        combox(COMBO_CODE_POSTE).ItemData(I - 1) = I
        combox(COMBO_LIB_POSTE).AddItem Item:=liste(I), Index:=I - 1
        combox(COMBO_LIB_POSTE).ItemData(I - 1) = I
        'combox(COMBO_CIVILITE).AddItem Item:=liste(I), Index:=I - 1
        'combox(COMBO_CIVILITE).ItemData(I - 1) = I
    Next I
    
    'paramétrage civilité et njf de 0 à 10 avec 0 état désactivé
    For I = 0 To 10
        combox(COMBO_CIVILITE).AddItem Item:=liste(I), Index:=I
        combox(COMBO_CIVILITE).ItemData(I) = I
        combox(COMBO_NJF).AddItem Item:=liste(I), Index:=I
        combox(COMBO_NJF).ItemData(I) = I
    Next I

    With grd
        .Rows = 1
        .FormatString = "kb_tisnum|Type d'information supplémentaire|Position|longueur"
        .col = GRDC_TIS_LIBELLE
        .CellFontBold = True
        .col = GRDC_TIS_POS
        .CellFontBold = True
        .col = GRDC_TIS_LONG
        .CellFontBold = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionByRow
        .ColWidth(GRDC_TIS_NUM) = 0
        .ColWidth(GRDC_TIS_LIBELLE) = (3 / 4) * .width - 100
        .ColWidth(GRDC_TIS_POS) = (1 / 8) * .width
        .ColAlignment(GRDC_TIS_LIBELLE) = flexAlignCenterCenter
        .ColAlignment(GRDC_TIS_POS) = flexAlignCenterCenter
        .ColWidth(GRDC_TIS_LONG) = (1 / 8) * .width
        .ColAlignment(GRDC_TIS_LONG) = flexAlignCenterCenter
    End With

    Call remplir_les_champs
    g_faire_combo_click = True
    g_mode_saisie = True

End Sub

Private Sub parcourir()
' ***********************************************
' Choisir le fichier d'importation dans le disuqe
' ***********************************************
    Dim chemin As String, fichier_importation As String
    Dim rs As rdoResultset

    ' Récupérer le chemin du fichier d'importation
    If Odbc_SelectV("SELECT PGB_Chemin FROM PrmGenB", rs) = P_ERREUR Then
        Exit Sub
    End If
    chemin = ""
    If Not rs.EOF Then
        chemin = rs("PGB_Chemin").Value
        rs.Close
    End If
    If chemin <> "" Then
        chemin = left$(chemin, InStrRev(chemin, "\"))
        If FICH_EstFichierOuRep(chemin) <> FICH_REP Then
            chemin = "c:\"
        End If
    Else
        chemin = "c:\"
    End If
    
    ' Récupérer le nom du fichier d'importation
    fichier_importation = Com_ChoixFichier.AppelFrm("Choix du fichier d'importation", "", _
                          chemin, P_EXTENSIONS_FICHIERS_IMPORTATION, False)
    If fichier_importation <> "" Then
        txt(TXT_CHEMIN) = fichier_importation
        cmd(CMD_OK).Enabled = True
    End If

End Sub

Private Function position_prise(ByVal v_pos As Integer) As Boolean

    Dim I As Integer

    For I = 1 To 8
        If combox(I).List(combox(I).ListIndex) = v_pos Then
            ' position prise
            position_prise = True
            Exit Function
        End If
    Next I
    ' pas trouvé
    position_prise = False

End Function

Private Sub quitter(ByVal v_force As Boolean)

    If v_force Then
        GoTo lab_exit
    Else
        If cmd(CMD_OK).Enabled Or Not g_quitter Then
            If MsgBox("Etes-vous sûr de vouloir quitter le paramétrage sans enregistrer les modifications ?", _
                       vbQuestion + vbYesNo) = vbYes Then
                GoTo lab_exit
            End If
        Else
            GoTo lab_exit
        End If
    End If

    Exit Sub

lab_exit:
    Unload Me

End Sub

Private Sub remplir_les_champs()

    Dim chemin_fichier As String, separateur As String
    Dim liste_des_champs As String, pgb_lstposinfoautre As String, sql As String
    Dim tis_num As Long
    Dim tis_long As Integer
    Dim type_fichier As Integer, nbr_tis As Integer, nbr As Integer
    Dim I As Integer, tis_pos As Integer
    Dim rs As rdoResultset
    Dim s As String
    Dim pos As Integer, fd As Integer
    Dim nomfich As String, sext As String
    
    sql = "SELECT * FROM PrmGenB"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    If Not rs.EOF Then
        g_enreg_existe = True
        chemin_fichier = rs("PGB_Chemin").Value
        type_fichier = rs("PGB_FichType").Value
        separateur = rs("PGB_FichSep").Value & ""
        liste_des_champs = rs("PGB_FichLstPos").Value
        txt(TXT_CODE).Text = rs("PGB_code").Value & ""
        txt(TXT_MPASSE).Text = rs("PGB_mdp").Value & ""
        optEmpl(OPT_EMPL_LOCAL).Value = Not rs("PGB_fichsurserveur").Value
        optEmpl(OPT_EMPL_SERVEUR).Value = rs("PGB_fichsurserveur").Value
    Else
        g_enreg_existe = False
        optEmpl(OPT_EMPL_LOCAL).Value = True
        optEmpl(OPT_EMPL_SERVEUR).Value = False
    End If
    rs.Close
    
    cmd(CMD_PARCOURIR).Visible = optEmpl(OPT_EMPL_LOCAL).Value
    
    ' ***************** Le type de fichier d'imporation *****************
    'combox(COMBO_TYPE_FICH).AddItem Item:=type_fichier
    combox(COMBO_TYPE_FICH).AddItem Item:="Fichier texte avec séparateur"
    combox(COMBO_TYPE_FICH).AddItem Item:="Fichier positionnel"
    Me.lblLig.Visible = False
    Me.LblTest.Visible = False
    cmd(CMD_NEXT_LIGNE).Visible = False
    cmd(CMD_NEXT_LIGNE).tag = 1
    cmd(CMD_NEXT_LIGNE).Caption = cmd(CMD_NEXT_LIGNE).tag
    If type_fichier = 1 Then
        combox(COMBO_TYPE_FICH).Text = "Fichier texte avec séparateur"
        Me.FrameSepar.Visible = True
        Me.FramePosit.Visible = False
        Me.txt(TXT_SEPARATEUR).Visible = True
        Me.lbl(LBL_SEPARATEUR).Visible = True
        With grd
            .ColWidth(GRDC_TIS_POS) = (1 / 4) * .width
            .ColWidth(GRDC_TIS_LONG) = 0
        End With
    Else
        combox(COMBO_TYPE_FICH).Text = "Fichier positionnel"
        Me.FrameSepar.Visible = False
        Me.FramePosit.Visible = True
        Me.txt(TXT_SEPARATEUR).Visible = False
        Me.lbl(LBL_SEPARATEUR).Visible = False
        With grd
            .ColWidth(GRDC_TIS_POS) = (1 / 8) * .width
            .ColWidth(GRDC_TIS_LONG) = (1 / 8) * .width
        End With
    End If
    ' ********************* Le séparateur de champs *********************
    txt(TXT_SEPARATEUR).Text = IIf(separateur <> "", separateur, "#")
    ' *************************** Les combobox ***************************
    Call remplir_pos(liste_des_champs)
    ' **************** Le chemin du fichier d'importation ****************
    If chemin_fichier <> "" Then
        If optEmpl(OPT_EMPL_LOCAL).Value Then
            If FICH_FichierExiste(chemin_fichier) Then
                txt(TXT_CHEMIN).Text = chemin_fichier
            Else
                Call MsgBox("L'ancien fichier d'importation: " & vbCrLf & vbCrLf & chemin_fichier _
                          & vbCrLf & vbCrLf & "est introuvable (local).", vbInformation + vbOKOnly, "Attention")
            End If
        Else
            chemin_fichier = Replace(chemin_fichier, "\", "/")
            If KF_FichierExiste(chemin_fichier) Then
                txt(TXT_CHEMIN).Text = chemin_fichier
            Else
                Call MsgBox("L'ancien fichier d'importation: " & vbCrLf & vbCrLf & chemin_fichier _
                          & vbCrLf & vbCrLf & "est introuvable (serveur).", vbInformation + vbOKOnly, "Attention")
            End If
        End If
    End If
    ' ********************************************************
    ' Récupérer la première ligne
    Call P_InitFichierImportation
    If p_est_sur_serveur Then
        pos = InStrRev(p_nom_fichier_importation, ".")
        sext = Mid$(p_nom_fichier_importation, pos)
        nomfich = p_chemin_appli & "\tmp\personnel_" & format(Date, "hhmmss") & sext
        If KF_GetFichier(p_nom_fichier_importation, nomfich) = P_ERREUR Then
            Call quitter(True)
            Exit Sub
        End If
    Else
        nomfich = p_nom_fichier_importation
    End If
    
    If FICH_OuvrirFichier(nomfich, FICH_LECTURE, fd) = P_ERREUR Then
        'Call quitter(True)
        Exit Sub
    End If
    I = 1
    While Not EOF(fd)
        Line Input #fd, s
        If I = 1 Then
            g_ligne1 = s
        ElseIf I = 2 Then
            g_ligne2 = s
        ElseIf I = 3 Then
            g_ligne3 = s
        Else
            GoTo LabSuite
        End If
        I = I + 1
    Wend
LabSuite:
    Close fd
    ' ********************************************************
    ' le tableau des types d'infos suppl
    sql = "SELECT PGB_LstPosInfoAutre FROM PrmGenB"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    If Not rs.EOF Then
        pgb_lstposinfoautre = rs("PGB_LstPosInfoAutre").Value & ""
    End If
    rs.Close
    ' pgb_lstposinfoautre = 'lx;pos1|ly;pos2|....'
    nbr_tis = STR_GetNbchamp(pgb_lstposinfoautre, "|")
    With grd
        For I = 0 To nbr_tis - 1
            tis_num = Mid$(STR_GetChamp(STR_GetChamp(pgb_lstposinfoautre, "|", I), ";", 0), 2)
            tis_pos = STR_GetChamp(STR_GetChamp(pgb_lstposinfoautre, "|", I), ";", 1)
            If type_fichier = 2 Then
                s = STR_GetChamp(STR_GetChamp(pgb_lstposinfoautre, "|", I), ";", 2)
                If IsNumeric(s) And val(s) > 0 Then
                    tis_long = s
                End If
            End If
            sql = "SELECT KB_TisLibelle FROM KB_TypeInfoSuppl WHERE KB_TisNum=" & tis_num _
                  & " AND KB_TisImport='t'"
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                Exit Sub
            End If
            If Not rs.EOF Then
                .AddItem ""
                .Row = .Rows - 1
                .col = GRDC_TIS_NUM
                .TextMatrix(.Rows - 1, GRDC_TIS_NUM) = tis_num
                .TextMatrix(.Rows - 1, GRDC_TIS_LIBELLE) = rs("KB_TisLibelle").Value
                .TextMatrix(.Rows - 1, GRDC_TIS_POS) = tis_pos
                If type_fichier = 2 Then
                    .TextMatrix(.Rows - 1, GRDC_TIS_LONG) = tis_long
                End If
                rs.Close
            End If
        Next I
    End With

    cmd(CMD_OK).Enabled = False
    g_quitter = True

End Sub

Private Sub remplir_pos(ByVal v_liste_des_champs As String)
' **************************************************************************************************
' Remplir les combobox des champs de données du fichier d'importation, exemple de v_liste_des_champs
'    "NOM=2;PRENOM=3;MATRICULE=0;CODE_SECTION=4;LIB_SECTION=5;CODE_FONCTION=6;LIB_FONCTION=7;LIB_CIVILITE=8"
' **************************************************************************************************
    Dim I As Integer
    
    On Error Resume Next ' pour eviter tout conflit inatendu
    If combox(COMBO_TYPE_FICH).ListIndex = 0 Then   ' séparateur
        combox(COMBO_NOM).Text = STR_GetChamp(STR_GetChamp(v_liste_des_champs, "NOM=", 1), ";", 0) + 1
        combox(COMBO_PRENOM).Text = STR_GetChamp(STR_GetChamp(v_liste_des_champs, "PRENOM=", 1), ";", 0) + 1
        combox(COMBO_MATRICULE).Text = STR_GetChamp(STR_GetChamp(v_liste_des_champs, "MATRICULE=", 1), ";", 0) + 1
        combox(COMBO_CODE_SECTION).Text = STR_GetChamp(STR_GetChamp(v_liste_des_champs, "CODE_SECTION=", 1), ";", 0) + 1
        combox(COMBO_LIB_SECTION).Text = STR_GetChamp(STR_GetChamp(v_liste_des_champs, "LIB_SECTION=", 1), ";", 0) + 1
        combox(COMBO_CODE_POSTE).Text = STR_GetChamp(STR_GetChamp(v_liste_des_champs, "CODE_FONCTION=", 1), ";", 0) + 1
        combox(COMBO_LIB_POSTE).Text = STR_GetChamp(STR_GetChamp(v_liste_des_champs, "LIB_FONCTION=", 1), ";", 0) + 1
        combox(COMBO_CIVILITE).Text = STR_GetChamp(STR_GetChamp(v_liste_des_champs, "LIB_CIVILITE=", 1), ";", 0) + 1
        combox(COMBO_NJF).Text = STR_GetChamp(STR_GetChamp(v_liste_des_champs, "NJF=", 1), ";", 0) + 1
    Else
        'For i = 0 To 10
            txtD(0).Text = STR_GetChamp(STR_GetChamp(STR_GetChamp(v_liste_des_champs, ";", 0), "=", 1), ":", 0)
            txtF(0).Text = STR_GetChamp(STR_GetChamp(STR_GetChamp(v_liste_des_champs, ";", 0), "=", 1), ":", 1)
            txtD(1).Text = STR_GetChamp(STR_GetChamp(STR_GetChamp(v_liste_des_champs, ";", 1), "=", 1), ":", 0)
            txtF(1).Text = STR_GetChamp(STR_GetChamp(STR_GetChamp(v_liste_des_champs, ";", 1), "=", 1), ":", 1)
            txtD(2).Text = STR_GetChamp(STR_GetChamp(STR_GetChamp(v_liste_des_champs, ";", 2), "=", 1), ":", 0)
            txtF(2).Text = STR_GetChamp(STR_GetChamp(STR_GetChamp(v_liste_des_champs, ";", 2), "=", 1), ":", 1)
            txtD(3).Text = STR_GetChamp(STR_GetChamp(STR_GetChamp(v_liste_des_champs, ";", 3), "=", 1), ":", 0)
            txtF(3).Text = STR_GetChamp(STR_GetChamp(STR_GetChamp(v_liste_des_champs, ";", 3), "=", 1), ":", 1)
            txtD(4).Text = STR_GetChamp(STR_GetChamp(STR_GetChamp(v_liste_des_champs, ";", 4), "=", 1), ":", 0)
            txtF(4).Text = STR_GetChamp(STR_GetChamp(STR_GetChamp(v_liste_des_champs, ";", 4), "=", 1), ":", 1)
            txtD(5).Text = STR_GetChamp(STR_GetChamp(STR_GetChamp(v_liste_des_champs, ";", 7), "=", 1), ":", 0)
            txtF(5).Text = STR_GetChamp(STR_GetChamp(STR_GetChamp(v_liste_des_champs, ";", 7), "=", 1), ":", 1)
            txtD(6).Text = STR_GetChamp(STR_GetChamp(STR_GetChamp(v_liste_des_champs, ";", 5), "=", 1), ":", 0)
            txtF(6).Text = STR_GetChamp(STR_GetChamp(STR_GetChamp(v_liste_des_champs, ";", 5), "=", 1), ":", 1)
            txtD(7).Text = STR_GetChamp(STR_GetChamp(STR_GetChamp(v_liste_des_champs, ";", 6), "=", 1), ":", 0)
            txtF(7).Text = STR_GetChamp(STR_GetChamp(STR_GetChamp(v_liste_des_champs, ";", 6), "=", 1), ":", 1)
            txtD(8).Text = STR_GetChamp(STR_GetChamp(STR_GetChamp(v_liste_des_champs, ";", 8), "=", 1), ":", 0)
            txtF(8).Text = STR_GetChamp(STR_GetChamp(STR_GetChamp(v_liste_des_champs, ";", 8), "=", 1), ":", 1)
        'Next i
    End If
End Sub

Private Function verifier_tout_champ() As Integer
' ***********************************************************************
' Verifier si tous les champs sont renseigner ET avec des valeurs valides
' ***********************************************************************
    Dim s As String, stest As String
    Dim I As Integer, j As Integer, posdeb As Integer

    ' ***************** Le caractère de séparation
    If txt(TXT_SEPARATEUR).Text = "" Then
        Call MsgBox("Le separateur des champs est une information obligatoire.", vbExclamation + vbOKOnly, "")
        txt(TXT_SEPARATEUR).SetFocus
        GoTo lab_erreur
    End If
    If Len(txt(TXT_SEPARATEUR).Text) > 1 And txt(TXT_SEPARATEUR).Text <> "TAB" Then
        Call MsgBox("Le séparateur est incorrect." & vbCrLf & "Caractère ascii ou TAB", vbExclamation + vbOKOnly, "")
        txt(TXT_SEPARATEUR).SetFocus
        GoTo lab_erreur
    End If
    
    ' ***************** Les positions des coordonnées dans les lignes du fichier d'importation
    If combox(COMBO_TYPE_FICH).ListIndex = 0 Then
        For I = 1 To 8
            If combox(I).List(combox(I).ListIndex) = "" Then
                Call MsgBox("La position de chaque champ dans le fichier d'importation est une information obligatoire.", _
                             vbExclamation + vbOKOnly, "")
                combox(I).SetFocus
                GoTo lab_erreur
            End If
            For j = I + 1 To 8
                If I = j Then
                    GoTo j_suivant
                End If
                If combox(I).List(combox(I).ListIndex) = combox(j).List(combox(j).ListIndex) Then
                    If I = COMBO_CODE_SECTION And j = COMBO_LIB_SECTION Then
                        GoTo j_suivant
                    End If
                    If I = COMBO_CODE_POSTE And j = COMBO_LIB_POSTE Then
                        GoTo j_suivant
                    End If
                    Call MsgBox("La position de chaque champ doit être différente de celle des autres !", _
                                 vbExclamation + vbOKOnly, "")
                    combox(I).SetFocus
                    GoTo lab_erreur
                End If
j_suivant:
            Next j
        Next I
    Else
        For I = 0 To 8
            'Debug.Print i & " " & txtD(i).Text & " " & txtF(i).Text
            If txtD(I).Text <> "" Then
                If IsNumeric(txtD(I).Text) Then
                    ' controler longueur
                    If txtF(I).Text <> "" And IsNumeric(txtF(I).Text) Then
                        ' OK
                    Else
                        MsgBox "Les champs Longueur doivent être renseignés et numériques"
                        txtF(I).SetFocus
                        GoTo lab_erreur
                    End If
                Else
                    MsgBox "Les champs Position doivent numériques"
                    txtD(I).SetFocus
                    GoTo lab_erreur
                End If
            Else
                If I <> 5 And I <> 8 Then ' civilité et NJF
                    MsgBox "Les champs Position doivent être renseignés"
                    txtD(I).SetFocus
                    GoTo lab_erreur
                End If
            End If
        Next I
    End If
    ' ***************** Le type de fichier **************************
    If combox(COMBO_TYPE_FICH).List(combox(COMBO_TYPE_FICH).ListIndex) = "" Then
        Call MsgBox("Le type de fichier d'importation est une information obligatoire.", _
                     vbExclamation + vbOKOnly, "")
        combox(COMBO_TYPE_FICH).SetFocus
        GoTo lab_erreur
    End If

    ' Format du code - <P> (prénom), <P1> (1e lettre du prénom), <N> (nom), <M> (matricule)
    If txt(TXT_CODE).Text = "" Then
        Call MsgBox("Le format de code est une information obligatoire.", _
                     vbExclamation + vbOKOnly, "")
        txt(TXT_CODE).SetFocus
        GoTo lab_erreur
    End If
    I = 1
    s = txt(TXT_CODE).Text
    posdeb = 0
    While I <= Len(s)
        If Mid$(s, I, 1) = "<" Then
            If posdeb > 0 Then
                GoTo lab_err_code
            End If
            posdeb = I
        ElseIf Mid$(s, I, 1) = ">" Then
            If posdeb = 0 Then
                GoTo lab_err_code
            End If
            stest = Mid$(s, posdeb + 1, I - posdeb - 1)
            If UCase(left$(stest, 1)) <> "P" And left$(stest, 1) <> "N" And left$(stest, 1) <> "J" And left$(stest, 1) <> "M" Then
                GoTo lab_err_code
            End If
            If left(stest, 2) = "P=" Then ' Code spécial : P=PS2;PC2
            ElseIf Len(stest) > 1 Then
                If Not IsNumeric(Mid$(stest, 2)) Then
                    GoTo lab_err_code
                End If
            End If
            posdeb = 0
        End If
        I = I + 1
    Wend
    
    ' Mot de passe
    If txt(TXT_MPASSE).Text = "" Then
        Call MsgBox("Le mot de passe est une information obligatoire.", _
                     vbExclamation + vbOKOnly, "")
        txt(TXT_MPASSE).SetFocus
        GoTo lab_erreur
    End If
    
    verifier_tout_champ = P_OK
    Exit Function

lab_err_code:
    Call MsgBox("Le format de code est incorrect." & vbCrLf & "<P>=Prénom, <P1>=1e lettre du prénom, <p1>=idem à P1 mais 2e lettre du prénom si prénom composé, <N>=Nom, <J>=Nom de jeune fille <M>=Matricule", _
                 vbExclamation + vbOKOnly, "")
    txt(TXT_CODE).SetFocus

lab_erreur:
    verifier_tout_champ = P_ERREUR

End Function

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case CMD_OK
            Call enregistrer
        Case CMD_QUITTER
            Call quitter(False)
        Case CMD_PARCOURIR
            Call parcourir
        Case CMD_PLUS_TIS
            Call ajouter_typeinfosuppl
        Case CMD_MOINS_TIS
            Call enlever_typeinfosuppl
        Case CMD_NEXT_LIGNE
            cmd(CMD_NEXT_LIGNE).tag = cmd(CMD_NEXT_LIGNE).tag + 1
            If cmd(CMD_NEXT_LIGNE).tag > 3 Then
                cmd(CMD_NEXT_LIGNE).tag = 1
            End If
            Call metLigneExemple
    End Select

End Sub

Private Sub metLigneExemple()
    Dim ligne As String
    
    Select Case cmd(CMD_NEXT_LIGNE).tag
    Case 1
        ligne = g_ligne1
    Case 2
        ligne = g_ligne2
    Case 3
        ligne = g_ligne3
    End Select
    cmd(CMD_NEXT_LIGNE).Caption = cmd(CMD_NEXT_LIGNE).tag
    Me.lblLig.Caption = ligne
End Sub

Private Sub combox_Click(Index As Integer)
    Dim sep As String
    Dim I As Integer
    Dim ind As Integer
    Dim s As String
    Dim ligne As String
    
    cmd(CMD_OK).Enabled = True
    If g_faire_combo_click And Index = COMBO_TYPE_FICH Then
        If combox(COMBO_TYPE_FICH).ListIndex = 0 Then
            Me.FrameSepar.Visible = True
            Me.FramePosit.Visible = False
            Me.txt(TXT_SEPARATEUR).Visible = True
            Me.lbl(LBL_SEPARATEUR).Visible = True
            With grd
                .ColWidth(GRDC_TIS_POS) = (1 / 4) * .width
                .ColWidth(GRDC_TIS_LONG) = 0
            End With
            grd.TextMatrix(0, GRDC_TIS_LONG) = ""
        Else
            Me.FrameSepar.Visible = False
            Me.FramePosit.Visible = True
            Me.txt(TXT_SEPARATEUR).Visible = False
            Me.lbl(LBL_SEPARATEUR).Visible = False
            With grd
                .ColWidth(GRDC_TIS_POS) = (1 / 8) * .width
                .ColWidth(GRDC_TIS_LONG) = (1 / 8) * .width
            End With
            grd.TextMatrix(0, GRDC_TIS_LONG) = "longueur"
        End If
    ElseIf g_faire_combo_click And Index >= 1 And Index <= 9 Then
        Me.LblTest.Visible = True
        Me.lblLig.Visible = True
        cmd(CMD_NEXT_LIGNE).Visible = True
        sep = txt(TXT_SEPARATEUR).Text
        If sep = "TAB" Then sep = Chr(9)

        Me.LblTest.Caption = ""
        ind = combox(Index).ListIndex
        ind = combox(Index).ItemData(ind)
        If ind >= 0 Then
            For I = 1 To 3
                If I = 1 Then
                    ligne = g_ligne1
                ElseIf I = 2 Then
                    ligne = g_ligne2
                ElseIf I = 3 Then
                    ligne = g_ligne3
                End If
                Me.LblTest.Caption = Me.LblTest.Caption & IIf(I > 1, " / ", "") & STR_GetChamp(ligne, sep, ind - 1)
                Call metLigneExemple
            Next I
        Else
            Me.LblTest.Visible = False
            Me.LblTest.Caption = ""
        End If
    Else
        Me.LblTest.Visible = False
        Me.lblLig.Visible = False
        cmd(CMD_NEXT_LIGNE).Visible = False
    End If

End Sub


Private Sub combox_GotFocus(Index As Integer)
    Call combox_Click(Index)
End Sub

Private Sub combox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn
            KeyCode = 0
            SendKeys "{TAB}"
        Case vbKeyF1
            KeyCode = 0
            If cmd(CMD_OK).Enabled = True Then Call enregistrer
        Case vbKeyEscape
            KeyCode = 0
            Call quitter(False)
    End Select

End Sub


Private Sub Form_Activate()

    If g_form_active Then Exit Sub

    g_form_active = True
    Call initialiser

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKeyF1
            Call enregistrer
    End Select

End Sub

Private Sub Form_Load()

    g_form_active = False

End Sub

Private Sub grd_Click()
    
    Dim valeur_retournee As String
    Dim I As Integer, mouse_row As Integer, mouse_col As Integer, _
        mon_left As Integer, mon_top As Integer
    Dim lib As String

    With grd
        If .Rows = 1 Then Exit Sub

        mouse_row = .MouseRow
        mouse_col = .MouseCol

        ' vérifications
        If mouse_row = 0 Then Exit Sub
        If mouse_col <> GRDC_TIS_POS And mouse_col <> GRDC_TIS_LONG Then Exit Sub
        ' demander la valeur
        mon_top = ((Me.Height - 2000) / 2) + Me.Top
        mon_left = ((Me.width - 5000) / 2) + Me.left
lab_autre_essai:
        If mouse_col = GRDC_TIS_POS Then
            lib = "position"
        ElseIf mouse_col = GRDC_TIS_LONG Then
            lib = "longueur"
        End If
        valeur_retournee = InputBox("Veuillez saisir la " & lib & " du champ (un entier) : " _
                           & vbCrLf & vbCrLf _
                           & .TextMatrix(mouse_row, GRDC_TIS_LIBELLE), _
                           "Position du champ", _
                           .TextMatrix(mouse_row, mouse_col), _
                           mon_left, mon_top)
        If StrPtr(valeur_retournee) = 0 Then
            ' on a annulé
        Else ' on a validé
            If IsNumeric(valeur_retournee) Then
                If mouse_col = GRDC_TIS_POS Then
                    If combox(COMBO_TYPE_FICH).ListIndex = 0 Then   ' seulement pour séparateur
                        If position_prise(Int(valeur_retournee)) Then
                            If MsgBox("Position déjà prise !" & vbCrLf & vbCrLf & _
                                      "Voulez-vous saisir une autre position SVP", _
                                      vbCritical + vbOKCancel, _
                                      "Position incorrecte") = vbOK Then
                                GoTo lab_autre_essai
                            Else ' annulation
                                Exit Sub
                            End If
                        Else
                            .col = mouse_col
                            .Row = mouse_row
                            .Text = Int(valeur_retournee)
                            cmd(CMD_OK).Enabled = True
                        End If
                    Else
                        .col = mouse_col
                        .Row = mouse_row
                        .Text = Int(valeur_retournee)
                        cmd(CMD_OK).Enabled = True
                    End If
                Else
                    .col = mouse_col
                    .Row = mouse_row
                    .Text = Int(valeur_retournee)
                    cmd(CMD_OK).Enabled = True
                End If
            Else
                If MsgBox("Veuillez saisir un entier SVP", vbCritical + vbOKCancel, _
                          "Valeur incorrecte") = vbOK Then
                    GoTo lab_autre_essai
                Else ' annulation
                    Exit Sub
                End If
            End If
        End If
    End With

End Sub

Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn
            KeyCode = 0
            SendKeys "{TAB}"
        Case vbKeyF1
            KeyCode = 0
            If cmd(CMD_OK).Enabled = True Then Call enregistrer
        Case vbKeyEscape
            KeyCode = 0
            Call quitter(False)
    End Select

End Sub

Private Sub optEmpl_Click(Index As Integer)

    If g_mode_saisie Then
        cmd(CMD_PARCOURIR).Visible = optEmpl(OPT_EMPL_LOCAL).Value
        Call txt_LostFocus(TXT_CHEMIN)
    End If
End Sub

Private Sub txt_Change(Index As Integer)

    g_quitter = False
    
    If Index = TXT_SEPARATEUR Then
    End If
    
    cmd(CMD_OK).Enabled = True

End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn
            KeyCode = 0
            SendKeys "{TAB}"
        Case vbKeyF1
            KeyCode = 0
            If cmd(CMD_OK).Enabled = True Then Call enregistrer
        Case vbKeyEscape
            KeyCode = 0
            Call quitter(False)
    End Select
    
End Sub

Private Sub txt_LostFocus(Index As Integer)

    If Index = TXT_CHEMIN Then
        txt(TXT_CHEMIN).Text = Replace(txt(TXT_CHEMIN).Text, "\", "/")
        If txt(TXT_CHEMIN).Text <> "" Then
            If optEmpl(OPT_EMPL_LOCAL).Value = True Then
                If Not FICH_FichierExiste(txt(TXT_CHEMIN).Text) Then
                    Call MsgBox("Le fichier " & txt(TXT_CHEMIN).Text & " est inaccessible( local).", vbInformation + vbOKOnly, "")
                    'txt(TXT_CHEMIN).SetFocus
                End If
            Else
                If Not KF_FichierExiste(txt(TXT_CHEMIN).Text) Then
                    Call MsgBox("Le fichier " & txt(TXT_CHEMIN).Text & " est inaccessible (serveur).", vbInformation + vbOKOnly, "")
                    'txt(TXT_CHEMIN).SetFocus
                End If
            End If
        End If
    End If
    
End Sub

Private Sub txtD_Change(Index As Integer)
    Call FaitExemple(Index)
End Sub

Private Sub txtD_Click(Index As Integer)
    Call FaitExemple(Index)
End Sub

Private Sub txtF_Change(Index As Integer)
    Call FaitExemple(Index)
End Sub

Private Function FaitExemple(Index As Integer)
    Dim I As Integer
    Dim ligne As String
    
    If g_mode_saisie = True Then
        Me.LblTest.Caption = ""
        For I = 1 To 3
            If I = 1 Then
                ligne = g_ligne1
            ElseIf I = 2 Then
                ligne = g_ligne2
            ElseIf I = 3 Then
                ligne = g_ligne3
            End If
            If IsNumeric(txtD(Index).Text) And IsNumeric(txtF(Index).Text) Then
                Me.LblTest.Caption = Me.LblTest.Caption & IIf(I > 1, " / ", "") & Mid(ligne, txtD(Index).Text, txtF(Index).Text)
            End If
        Next I
        Me.lblLig.Visible = True
        cmd(CMD_NEXT_LIGNE).Visible = True
        Call metLigneExemple
        Me.LblTest.Visible = True
        cmd(CMD_OK).Enabled = True
    End If
End Function

Private Sub txtF_Click(Index As Integer)
    Call FaitExemple(Index)
End Sub
