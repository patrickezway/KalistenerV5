VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form PrmApplication 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8085
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmPatience 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Chargement en cours ..."
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
      Height          =   1335
      Left            =   7440
      TabIndex        =   36
      Top             =   2880
      Visible         =   0   'False
      Width           =   7815
      Begin ComctlLib.ProgressBar pgb 
         Height          =   495
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   873
         _Version        =   327682
         Appearance      =   1
         Max             =   1000
      End
   End
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Application"
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
      Height          =   7485
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8085
      Begin VB.CommandButton cmd 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   6360
         Picture         =   "PrmApplication.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Supprimer le type de coordonnées en cours de cette zone."
         Top             =   1440
         Width           =   300
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Gérer l'envoi du fichier des modifications"
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
         Left            =   405
         TabIndex        =   34
         Top             =   1845
         Width           =   3915
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Height          =   450
         Index           =   10
         Left            =   7410
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Gérer les types d'informations supplémentaires"
         Top             =   75
         Width           =   435
      End
      Begin VB.Frame frm 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   6360
         Width           =   7815
         Begin VB.CommandButton cmd 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   7170
            Picture         =   "PrmApplication.frx":0447
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "Ajouter des types de coordonnées dans cette zone."
            Top             =   0
            Width           =   300
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   5
            Left            =   7170
            Picture         =   "PrmApplication.frx":089E
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "Supprimer le type de coordonnées en cours de cette zone."
            Top             =   700
            Width           =   300
         End
         Begin MSFlexGridLib.MSFlexGrid grdAppli 
            Height          =   1005
            Index           =   0
            Left            =   1320
            TabIndex        =   26
            ToolTipText     =   "Liste des zones modifiables par le responsable de cette application."
            Top             =   0
            Width           =   5835
            _ExtentX        =   10292
            _ExtentY        =   1773
            _Version        =   393216
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   8388608
            BackColorFixed  =   8454143
            ForeColorFixed  =   0
            BackColorSel    =   8388608
            ForeColorSel    =   16777215
            BackColorBkg    =   16777215
            GridColor       =   16777215
            GridColorFixed  =   16777215
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            GridLines       =   0
            GridLinesFixed  =   0
            ScrollBars      =   2
            SelectionMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Zones modifiables"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   5
            Left            =   0
            TabIndex        =   27
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame frm 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   4125
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   2145
         Width           =   7695
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Height          =   495
            Index           =   9
            Left            =   240
            Picture         =   "PrmApplication.frx":0CE5
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Appliquer la règle à toutes les fonctions"
            Top             =   1830
            Width           =   615
         End
         Begin VB.PictureBox pct 
            BackColor       =   &H0000FF00&
            DrawStyle       =   4  'Dash-Dot-Dot
            FillColor       =   &H00C000C0&
            FillStyle       =   6  'Cross
            FontTransparent =   0   'False
            Height          =   255
            Index           =   0
            Left            =   2160
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   31
            ToolTipText     =   "Cliquez ici pour afficher uniquement les profils concernés."
            Top             =   60
            Width           =   255
         End
         Begin VB.PictureBox pct 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   2
            Left            =   5610
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   30
            ToolTipText     =   "Cliquez ici pour afficher uniquement les profils non ancore concernés."
            Top             =   30
            Width           =   255
         End
         Begin VB.PictureBox pct 
            BackColor       =   &H000000FF&
            Height          =   255
            Index           =   1
            Left            =   3680
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   29
            ToolTipText     =   "Cliquez ici pour afficher uniquement les profils non concernés."
            Top             =   30
            Width           =   255
         End
         Begin VB.OptionButton opt 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Non concernés"
            Height          =   285
            Index           =   2
            Left            =   3720
            TabIndex        =   28
            ToolTipText     =   "Afficher uniquement les profils non concernés."
            Top             =   30
            Width           =   1455
         End
         Begin VB.OptionButton opt 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Concernés"
            Height          =   285
            Index           =   1
            Left            =   2220
            TabIndex        =   17
            ToolTipText     =   "Afficher uniquement les profils concernés."
            Top             =   30
            Width           =   1155
         End
         Begin VB.OptionButton opt 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Tous"
            Height          =   285
            Index           =   0
            Left            =   1230
            TabIndex        =   16
            ToolTipText     =   "Afficher tous les profils."
            Top             =   30
            Width           =   795
         End
         Begin VB.OptionButton opt 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Non &renseignés"
            Height          =   285
            Index           =   3
            Left            =   5640
            TabIndex        =   15
            ToolTipText     =   "Afficher les profils non encore renseignés."
            Top             =   30
            Width           =   1455
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   7
            Left            =   7170
            Picture         =   "PrmApplication.frx":12A0
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Supprimer le type de coordonnées en cours de cette zone."
            Top             =   3750
            Width           =   300
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   6
            Left            =   7170
            Picture         =   "PrmApplication.frx":16E7
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Ajouter des types de coordonnées dans cette zone."
            Top             =   3030
            Width           =   300
         End
         Begin VB.CommandButton cmd 
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
            Height          =   495
            Index           =   8
            Left            =   240
            Picture         =   "PrmApplication.frx":1B3E
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Rafraichir"
            Top             =   1110
            Visible         =   0   'False
            Width           =   615
         End
         Begin ComctlLib.TreeView tv 
            Height          =   2475
            Index           =   0
            Left            =   1320
            TabIndex        =   18
            ToolTipText     =   "Utilisez la barre d'espace ou le bouton droit de la souris afin de basculer l'état du service/poste."
            Top             =   390
            Width           =   5835
            _ExtentX        =   10292
            _ExtentY        =   4366
            _Version        =   327682
            Indentation     =   0
            LabelEdit       =   1
            Style           =   7
            ImageList       =   "imglst"
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid grdAppli 
            Height          =   1005
            Index           =   1
            Left            =   1320
            TabIndex        =   19
            ToolTipText     =   "Liste des zones pour lesquelles le responsable de cette application sera prévenu en cas de modification."
            Top             =   3030
            Width           =   5835
            _ExtentX        =   10292
            _ExtentY        =   1773
            _Version        =   393216
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   8388608
            BackColorFixed  =   8454143
            ForeColorFixed  =   0
            BackColorSel    =   8388608
            ForeColorSel    =   16777215
            BackColorBkg    =   16777215
            GridColor       =   4194304
            GridColorFixed  =   16777215
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            GridLines       =   0
            GridLinesFixed  =   0
            ScrollBars      =   2
            SelectionMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ComctlLib.TreeView tv 
            Height          =   2475
            Index           =   1
            Left            =   1320
            TabIndex        =   22
            ToolTipText     =   "Utilisez la barre d'espace ou le bouton droit de la souris afin de basculer l'état du service/poste."
            Top             =   390
            Width           =   5835
            _ExtentX        =   10292
            _ExtentY        =   4366
            _Version        =   327682
            Indentation     =   0
            LabelEdit       =   1
            Style           =   7
            ImageList       =   "imglst"
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lbl 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Profil"
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
            TabIndex        =   21
            Top             =   750
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Prévenir sur modification de..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   3
            Left            =   0
            TabIndex        =   20
            Top             =   3150
            Width           =   1125
         End
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   6360
         Picture         =   "PrmApplication.frx":2021
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Choisir la personne à prevenir."
         Top             =   1080
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   930
         MaxLength       =   15
         TabIndex        =   0
         Top             =   600
         Width           =   2115
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   4170
         MaxLength       =   50
         TabIndex        =   1
         Top             =   630
         Width           =   3735
      End
      Begin MSFlexGridLib.MSFlexGrid grdAppli 
         Height          =   645
         Index           =   2
         Left            =   1680
         TabIndex        =   38
         ToolTipText     =   "Liste des zones pour lesquelles le responsable de cette application sera prévenu en cas de modification."
         Top             =   1080
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   1138
         _Version        =   393216
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   8388608
         BackColorFixed  =   8454143
         ForeColorFixed  =   0
         BackColorSel    =   8388608
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         GridColor       =   4194304
         GridColorFixed  =   16777215
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ComctlLib.ImageList imglst 
         Left            =   7440
         Top             =   1080
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   21
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   14
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmApplication.frx":2478
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmApplication.frx":2AC2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmApplication.frx":310C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmApplication.frx":3756
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmApplication.frx":3DA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmApplication.frx":43EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmApplication.frx":4A34
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmApplication.frx":50C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmApplication.frx":5750
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmApplication.frx":5DDE
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmApplication.frx":6038
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmApplication.frx":6292
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmApplication.frx":64EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmApplication.frx":6ABE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
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
         Left            =   300
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
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
         Index           =   0
         Left            =   300
         TabIndex        =   9
         Top             =   660
         Width           =   615
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
         Index           =   1
         Left            =   3660
         TabIndex        =   8
         Top             =   660
         Width           =   615
      End
   End
   Begin VB.Frame frmFct 
      BackColor       =   &H00808080&
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   7350
      Width           =   8085
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
         Index           =   11
         Left            =   2640
         Picture         =   "PrmApplication.frx":7090
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   240
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "PrmApplication.frx":76B5
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
         Left            =   510
         Picture         =   "PrmApplication.frx":7C11
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Left            =   6990
         Picture         =   "PrmApplication.frx":817A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   230
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "PrmApplication.frx":8733
         Height          =   510
         Index           =   2
         Left            =   4830
         Picture         =   "PrmApplication.frx":8CC2
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   230
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "PrmApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Index des objets CMD
Private Const CMD_OK = 0
Private Const CMD_QUITTER = 1
Private Const CMD_DETRUIRE = 2
Private Const CMD_PLUS_RESP = 3
Private Const CMD_MOINS_RESP = 12
Private Const CMD_PLUS_ZONE_RENS = 4
Private Const CMD_MOINS_ZONE_RENS = 5
Private Const CMD_PLUS_ZONE_PREV = 6
Private Const CMD_MOINS_ZONE_PREV = 7
Private Const CMD_RAFRAICHIR = 8
Private Const CMD_RENSEIGNER = 9
Private Const CMD_TYPE_INFO_SUPPL = 10
Private Const CMD_IMPRIMER = 11

' Index des carrés colorés
Private Const PCT_SEL = 0
Private Const PCT_NOSEL = 1
Private Const PCT_NORENS = 2

' Index des images des POSTE, SERVICES et SITE
Private Const IMG_SRV = 1
Private Const IMG_POSTE = 2
Private Const IMG_SRV_SEL = 3
Private Const IMG_POSTE_SEL = 4
Private Const IMG_SRV_NOSEL = 5
Private Const IMG_POSTE_NOSEL = 6
Private Const IMG_SITE = 7
Private Const IMG_SITE_SEL = 8
Private Const IMG_SITE_NOSEL = 9
' Index des cases cochées
Private Const IMG_VERT_COCHE = 10
Private Const IMG_ROUGE_COCHE = 11
Private Const IMG_GRIS_COCHE = 12
Private Const IMG_INFO = 13
Private Const IMG_PAS_INFO = 14

' Index des FRAMES
Private Const FRM_PRINCIPALE = 0
Private Const FRM_PREVENIR = 1
Private Const FRM_MODIFIABLES = 2

' Index des libellés
Private Const LBL_INFORMER = 6

' Clef de la racine du TreeView
Private Const CLEF_ROOT = "L1"

' Index des GRID
Private Const GRD_RENSEIGNER = 0
Private Const GRD_PREVENIR = 1
Private Const GRD_RESP = 2

' Index des colonnes du GRD_RENSEIGNER
Private Const GRDR_NUM = 0
Private Const GRDR_CODE = 1
Private Const GRDR_LIBELLE = 2

' Index des colonnes du GRD_PREVENIR
Private Const GRDP_NUM = 0
Private Const GRDP_CODE = 1
Private Const GRDP_LIBELLE = 2
Private Const GRDP_TYPE = 3 ' Zonutil ou Infosuppl

' Index des colonnes du GRD_RESP
Private Const GRDRESP_NUM = 0
Private Const GRDRESP_NOM = 1

' Index des objets txt
Private Const TXT_CODE = 0
Private Const TXT_NOM = 1

' Index des boutons radio
Private Const OPT_TOUS = 0
Private Const OPT_CONC = 1
Private Const OPT_NON_CONC = 2
Private Const OPT_NON_RENS = 3

' Index des TREEVIEW
Private Const TV_PROFIL_TOUS = 0
Private Const TV_PROFIL_REDUIT = 1

' La lettre des zonesrens et zoneprev
Private Const LETTRE_ZONE = "Z"
Private Const LETTRE_INFO_SUPPL = "I"

' No appli en saisie (0 si nouveau)
Private g_numappli As Long

Private g_spc As String
Private g_spnc As String
' Le bouton radio en cours
Private g_mode_conc As Integer

' Indique si la forme a déjà été activée
Private g_form_active As Boolean

' Indique si la saisie est en-cours
Private g_mode_saisie As Boolean

' Stocke le texte avant modif pour gérer le changement
Private g_txt_avant As String
' le TreeView actuellement affiché
Private g_tv_actuel As Integer

Private Function afficher_appli(ByVal v_numappli As Long) As Integer
' ***********************************************************
' Afficher les coordonnées de l'application.
' Un traitement spécial pour KaliDoc, KaliBottin et KaliMail.
' ***********************************************************
    Dim sql As String, champ_ZoneRens As String, champ_ZonePrev As String, _
        le_type As String, le_num As String, nomutil As String, s As String
    Dim i As Integer, n As Integer
    Dim numutil As Long
    Dim rs As rdoResultset

    Call FRM_ResizeForm(Me, Me.width, Me.Height)

    g_mode_saisie = False
    
    tv(TV_PROFIL_TOUS).ZOrder 0
    tv(TV_PROFIL_REDUIT).ZOrder 1
    opt(OPT_TOUS).Value = True
    opt(OPT_CONC).Value = False
    opt(OPT_NON_CONC).Value = False
    opt(OPT_NON_RENS).Value = False
    grdAppli(GRD_RESP).Rows = 0
    
    g_mode_conc = OPT_TOUS
    pct(PCT_SEL).Picture = LoadPicture("")
    pct(PCT_NOSEL).Picture = LoadPicture("")
    pct(PCT_NORENS).Picture = LoadPicture("")

    If v_numappli > 0 Then ' ************** MODE MODIFICATION *************************************
        g_numappli = v_numappli
        ' BTN Type d'info Supple visible seulement lors d'une MAJ
        cmd(CMD_TYPE_INFO_SUPPL).Visible = True
        Call evalue_btninfosuppl
        
        sql = "SELECT * from Application" _
            & " WHERE APP_Num=" & v_numappli
        If Odbc_Select(sql, rs) = P_ERREUR Then
            afficher_appli = P_ERREUR
            Exit Function
        End If
        If rs("APP_Code").Value = "KALIDOC" Or rs("APP_Code").Value = "KALIMAIL" _
                Or rs("APP_Code").Value = "KALIBOTTIN" Then
            cmd(CMD_DETRUIRE).Visible = False
            txt(TXT_CODE).Enabled = False
            txt(TXT_NOM).Enabled = False
        Else
            cmd(CMD_DETRUIRE).Visible = True
            txt(TXT_CODE).Enabled = True
            txt(TXT_NOM).Enabled = True
        End If
        txt(TXT_CODE).Text = rs("APP_Code").Value
        txt(TXT_NOM).Text = rs("APP_Nom").Value
        If Not IsNull(rs("app_lstresp").Value) Then
            n = STR_GetNbchamp(rs("app_lstresp").Value, ";")
            For i = 0 To n - 1
                s = STR_GetChamp(rs("app_lstresp").Value, ";", i)
                grdAppli(GRD_RESP).AddItem ""
                numutil = Mid$(s, 2)
                grdAppli(GRD_RESP).TextMatrix(i, 0) = numutil
                Call P_RecupUtilNomP(numutil, nomutil)
                grdAppli(GRD_RESP).TextMatrix(i, 1) = nomutil
            Next i
        End If
        cmd(CMD_MOINS_RESP).Visible = IIf(grdAppli(GRD_RESP).Rows > 0, True, False)
        champ_ZoneRens = rs("APP_ZoneRens").Value & ""
        champ_ZonePrev = rs("APP_ZonePrev").Value & ""
        If afficher_structure(rs("APP_Profil_conc").Value & "", rs("APP_Profil_NonConc").Value & "") = P_ERREUR Then
            afficher_appli = P_ERREUR
            Exit Function
        End If
        ' Afficher le checkbox pour informer l'appli ou pas (sauf KaliDoc et KaliBottin)
        If rs("APP_Code").Value <> "KALIDOC" And rs("APP_Code").Value <> "KALIBOTTIN" Then
            'frm(FRM_MODIFIABLES).Top = 6360
            frm(FRM_MODIFIABLES).Visible = True
            frm(FRM_PREVENIR).Visible = rs("APP_Informer").Value
            cmd(CMD_TYPE_INFO_SUPPL).Visible = rs("APP_Informer").Value
            chk.Visible = True
            chk.Value = IIf(rs("APP_Informer").Value, 1, 0)
        ElseIf rs("APP_Code").Value = "KALIDOC" Then
            'frm(FRM_MODIFIABLES).Top = 3000
            frm(FRM_MODIFIABLES).Visible = True
            frm(FRM_PREVENIR).Visible = False
            cmd(CMD_TYPE_INFO_SUPPL).Visible = False
            chk.Visible = False
            chk.Value = 1
        Else ' rs("APP_Code").Value = "KALIBOTTIN"
            frm(FRM_MODIFIABLES).Visible = False
            frm(FRM_PREVENIR).Visible = False
            cmd(CMD_TYPE_INFO_SUPPL).Visible = False
            chk.Visible = False
            chk.Value = 1
        End If
        rs.Close

        ' On réinitialise les deux grids:
        grdAppli(GRD_RENSEIGNER).Rows = 0
        grdAppli(GRD_PREVENIR).Rows = 0

        ' Remplissage du GRD_RENSEIGNER
        ' *****************************
        i = 0
        While i < STR_GetNbchamp(champ_ZoneRens, ";")
            sql = "SELECT ZU_Num, ZU_Code, ZU_Libelle FROM ZoneUtil WHERE ZU_Num=" _
                    & Mid$(STR_GetChamp(champ_ZoneRens, ";", i), 2)
            If Odbc_Select(sql, rs) = P_ERREUR Then
                afficher_appli = P_ERREUR
                Exit Function
            End If
            grdAppli(GRD_RENSEIGNER).AddItem rs("ZU_Num").Value
            grdAppli(GRD_RENSEIGNER).TextMatrix(grdAppli(GRD_RENSEIGNER).Rows - 1, GRDR_CODE) = rs("ZU_Code").Value
            grdAppli(GRD_RENSEIGNER).TextMatrix(grdAppli(GRD_RENSEIGNER).Rows - 1, GRDR_LIBELLE) = rs("ZU_Libelle").Value
            grdAppli(GRD_RENSEIGNER).Row = grdAppli(GRD_RENSEIGNER).Rows - 1
            grdAppli(GRD_RENSEIGNER).col = GRDR_CODE
            grdAppli(GRD_RENSEIGNER).CellFontBold = True
            rs.Close
            i = i + 1
        Wend
        cmd(CMD_MOINS_ZONE_RENS).Visible = IIf(grdAppli(GRD_RENSEIGNER).Rows > 0, True, False)

        ' Remplissage du GRID_PREVENIR
        ' ****************************
        i = 0
        While i < STR_GetNbchamp(champ_ZonePrev, ";")
            le_type = left$(STR_GetChamp(champ_ZonePrev, ";", i), 1)
            le_num = Mid$(STR_GetChamp(champ_ZonePrev, ";", i), 2)
            If le_type = "Z" Then
                ' les ZonUtil
                sql = "SELECT ZU_Num, ZU_Code, ZU_Libelle FROM Zoneutil" & _
                      " WHERE ZU_Num=" & le_num
                If Odbc_Select(sql, rs) = P_ERREUR Then
                    afficher_appli = P_ERREUR
                    Exit Function
                End If
                grdAppli(GRD_PREVENIR).AddItem le_num
                grdAppli(GRD_PREVENIR).TextMatrix(grdAppli(GRD_PREVENIR).Rows - 1, GRDP_CODE) = rs("ZU_Code").Value
                grdAppli(GRD_PREVENIR).TextMatrix(grdAppli(GRD_PREVENIR).Rows - 1, GRDP_LIBELLE) = rs("ZU_Libelle").Value
            Else ' le_type = "I"
                ' les InfoSuppl
                sql = "SELECT KB_TisCode, KB_TisLibelle FROM KB_TypeInfoSuppl" & _
                      " WHERE KB_TisNum=" & le_num
                If Odbc_Select(sql, rs) = P_ERREUR Then
                    afficher_appli = P_ERREUR
                    Exit Function
                End If
                grdAppli(GRD_PREVENIR).AddItem le_num
                grdAppli(GRD_PREVENIR).TextMatrix(grdAppli(GRD_PREVENIR).Rows - 1, GRDP_CODE) = rs("KB_TisCode").Value
                grdAppli(GRD_PREVENIR).TextMatrix(grdAppli(GRD_PREVENIR).Rows - 1, GRDP_LIBELLE) = rs("KB_TisLibelle").Value
            End If
            grdAppli(GRD_PREVENIR).TextMatrix(grdAppli(GRD_PREVENIR).Rows - 1, GRDP_TYPE) = le_type
            grdAppli(GRD_PREVENIR).Row = grdAppli(GRD_PREVENIR).Rows - 1
            grdAppli(GRD_PREVENIR).col = GRDR_CODE
            grdAppli(GRD_PREVENIR).CellFontBold = True
            i = i + 1
        Wend
        cmd(CMD_MOINS_ZONE_PREV).Visible = IIf(grdAppli(GRD_PREVENIR).Rows > 0, True, False)
        
        ' Empêcher partiellement la modification/suppression des trois applications suivantes:
        If txt(TXT_CODE).Text = "KALIDOC" Or txt(TXT_CODE).Text = "KALIMAIL" _
                Or txt(TXT_CODE).Text = "KALIBOTTIN" Then
            cmd(CMD_DETRUIRE).Visible = False
            txt(TXT_CODE).Enabled = False
            txt(TXT_NOM).Enabled = False
            'txt(TXT_RESP).SetFocus
        Else
            txt(TXT_CODE).Enabled = True
            txt(TXT_NOM).Enabled = True
            txt(TXT_CODE).SetFocus
        End If

    Else ' ****************************** MODE CRÉATION *******************************************
        cmd(CMD_TYPE_INFO_SUPPL).Visible = False
        grdAppli(GRD_RENSEIGNER).Rows = 0
        grdAppli(GRD_PREVENIR).Rows = 0
        txt(TXT_CODE).Text = ""
        txt(TXT_NOM).Text = ""
        If afficher_structure("", "") = P_ERREUR Then
            afficher_appli = P_ERREUR
            Exit Function
        End If
        g_numappli = 0
        ' pour les réactiver si on repasse par le mode création après les trois applications précédantes
        frm(FRM_MODIFIABLES).Visible = False
        frm(FRM_PREVENIR).Visible = False
        cmd(CMD_DETRUIRE).Visible = False
        txt(TXT_CODE).Enabled = True
        txt(TXT_NOM).Enabled = True
        txt(TXT_CODE).SetFocus
        cmd(CMD_MOINS_ZONE_RENS).Visible = False
        cmd(CMD_MOINS_ZONE_PREV).Visible = False
        cmd(CMD_MOINS_RESP).Visible = False
    End If ' **************************************************************************************
    
    cmd(CMD_OK).Enabled = False

    Me.MousePointer = 0
    g_mode_saisie = True

    afficher_appli = P_OK

End Function

Private Function afficher_structure(ByVal v_spc As Variant, _
                                    ByVal v_spnc As Variant) As Integer

    Dim sql As String, s As String, s_sp As String, _
        app_profil_conc As String, app_profil_nonconc As String
    Dim img As Integer, i As Integer, nbc As Integer, n As Integer, _
        image_racine As Integer, image_service As Integer, image_poste As Integer
    Dim rs As rdoResultset
    Dim nd As Node, ndP As Node, ndF As Node, nd_tmp As Node

    tv(TV_PROFIL_TOUS).Nodes.Clear

    If g_mode_saisie Then
        sql = "SELECT APP_Profil_Conc, APP_Profil_NonConc FROM Application " _
            & " WHERE APP_Num=" & g_numappli
        If Odbc_RecupVal(sql, app_profil_conc, app_profil_nonconc) = P_ERREUR Then
            GoTo lab_erreur
        End If
    Else
        app_profil_conc = ""
        app_profil_nonconc = ""
    End If

    ' Valeurs par défaut
    image_racine = IMG_SITE
    image_service = IMG_SRV
    image_poste = IMG_POSTE
    ' Gérer le cas où tout est selectionnés/non selectionnés:
    If app_profil_conc = "T" Then
        image_racine = IMG_SITE_SEL
        image_service = IMG_SRV_SEL
        image_poste = IMG_POSTE_SEL
    ElseIf app_profil_nonconc = "T" Then
        image_racine = IMG_SITE_NOSEL
        image_service = IMG_SRV_NOSEL
        image_poste = IMG_POSTE_NOSEL
    End If

    ' Affichage du(des) site(s)
    sql = "SELECT L_Num, L_Code FROM Laboratoire"
    If Odbc_Select(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    While Not rs.EOF
        Set nd = tv(TV_PROFIL_TOUS).Nodes.Add(, , "L" & rs("L_Num").Value, _
                 rs("L_Code").Value, image_racine, image_racine)
        If rs("L_Num").Value = p_NumLabo Then
            nd.selected = True
            nd.Expanded = True
        End If
        nd.Expanded = True
        nd.Sorted = True
        rs.MoveNext
    Wend
    rs.Close

    ' Affichage des services
    sql = "SELECT SRV_Num, SRV_LNum, SRV_NumPere, SRV_Nom FROM Service" _
        & " ORDER BY SRV_LNum, SRV_NumPere"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    While Not rs.EOF
'If rs("SRV_Num").Value = 62 Then MsgBox "srv_num=10: "
        ' Recherche le noeud père

        ' Le service est rattaché directement au site
        If rs("SRV_NumPere").Value = 0 Then
            Set ndP = tv(TV_PROFIL_TOUS).Nodes("L" & rs("SRV_LNum").Value)
        ' Le service est rattaché à un autre service
        Else
            ' Le service a déjà été ajouté car père d'un autre service déjà traité
            If TV_NodeExiste(tv(TV_PROFIL_TOUS), "S" & rs("SRV_Num").Value, nd) = P_OUI Then
                GoTo lab_suivant
            End If
            If TV_NodeExiste(tv(TV_PROFIL_TOUS), "S" & rs("SRV_NumPere").Value, nd) = P_NON Then
                ' Le service est père n'est pas affiché -> on l'ajoute
                Call ajouter_service(0, rs("SRV_NumPere").Value, app_profil_conc, app_profil_nonconc)
            End If
            Set ndP = tv(TV_PROFIL_TOUS).Nodes("S" & rs("SRV_NumPere").Value)
        End If
        ' Ajout du service
        Set nd = tv(TV_PROFIL_TOUS).Nodes.Add(ndP, _
                               tvwChild, _
                               "S" & rs("SRV_Num").Value, _
                               rs("SRV_Nom").Value, _
                               image_service, _
                               image_service)
        nd.Sorted = True
lab_suivant:
        rs.MoveNext
    Wend
    rs.Close

    ' Affichage des postes
    sql = "SELECT PO_Num, PO_SRVNum, FT_Libelle" _
        & " FROM Poste, FctTrav" _
        & " WHERE FT_Num=PO_FTNum and po_actif=true" _
        & " ORDER BY PO_Num"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    While Not rs.EOF
        Set ndP = tv(TV_PROFIL_TOUS).Nodes("S" & rs("PO_SRVNum").Value)
        Set nd = tv(TV_PROFIL_TOUS).Nodes.Add(ndP, _
                               tvwChild, _
                               "P" & rs("PO_Num").Value, _
                               rs("FT_Libelle").Value, _
                               image_poste, _
                               image_poste)
        rs.MoveNext
    Wend
    rs.Close

    ' Met en évidence les services/postes concernés
    If v_spc <> "" Then
        nbc = STR_GetNbchamp(v_spc, "|")
        For i = 0 To nbc - 1
            ' si on a tout selectionné
            s = STR_GetChamp(v_spc, "|", i)
            n = STR_GetNbchamp(s, ";")
            s_sp = STR_GetChamp(s, ";", n - 1)
            If TV_NodeExiste(tv(TV_PROFIL_TOUS), s_sp, nd) = P_OUI Then
                If left$(s_sp, 1) = "S" Then
                    img = IMG_SRV_SEL
                ElseIf left$(s_sp, 1) = "P" Then
                    img = IMG_POSTE_SEL
                ElseIf left$(s_sp, 1) = "L" Then
                    img = IMG_SITE_SEL
                End If
                nd.image = img
                nd.SelectedImage = img
                Set ndF = nd
                While TV_ChildNextParent(ndF, nd)
                    If left$(ndF.key, 1) = "S" Then
                        ndF.image = IMG_SRV_SEL
                        ndF.SelectedImage = IMG_SRV_SEL
                    ElseIf left$(ndF.key, 1) = "P" Then
                        ndF.image = IMG_POSTE_SEL
                        ndF.SelectedImage = IMG_POSTE_SEL
                    Else
                        ndF.image = img
                        ndF.SelectedImage = img
                    End If
                Wend
            End If
        Next i
    End If
    ' Met en évidence les services/postes non concernés
    If v_spnc <> "" Then
        nbc = STR_GetNbchamp(v_spnc, "|")
        For i = 0 To nbc - 1
            ' si rien n'a été selectionné
            s = STR_GetChamp(v_spnc, "|", i)
            n = STR_GetNbchamp(s, ";")
            s_sp = STR_GetChamp(s, ";", n - 1)
            If TV_NodeExiste(tv(TV_PROFIL_TOUS), s_sp, nd) = P_OUI Then
                If left$(s_sp, 1) = "S" Then
                    img = IMG_SRV_NOSEL
                ElseIf left$(s_sp, 1) = "P" Then
                    img = IMG_POSTE_NOSEL
                ElseIf left$(s_sp, 1) = "L" Then
                    img = IMG_SITE_NOSEL
                End If
                nd.image = img
                nd.SelectedImage = img
                Set ndF = nd
                While TV_ChildNextParent(ndF, nd)
                    If left$(ndF.key, 1) = "S" Then
                        ndF.image = IMG_SRV_NOSEL
                        ndF.SelectedImage = IMG_SRV_NOSEL
                    ElseIf left$(ndF.key, 1) = "P" Then
                        ndF.image = IMG_POSTE_NOSEL
                        ndF.SelectedImage = IMG_POSTE_NOSEL
                    Else
                        ndF.image = img
                        ndF.SelectedImage = img
                    End If
                Wend
            End If
        Next i
    End If

    ' une double vérification des icones (cas particuliers)
    sql = "SELECT MAX(PO_Num) AS PO_Num, PO_SrvNum" & _
          " FROM Poste GROUP BY PO_SrvNum"
    If Odbc_Select(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    While Not rs.EOF
        If TV_NodeExiste(tv(TV_PROFIL_TOUS), "P" & rs("PO_Num").Value, nd_tmp) = P_OUI Then
            Call fix_couleur_parent(TV_PROFIL_TOUS, tv(TV_PROFIL_TOUS).Nodes("P" & rs("PO_Num").Value))
        End If
        rs.MoveNext
    Wend
    rs.Close

    afficher_structure = P_OK
    Exit Function

lab_erreur:
    afficher_structure = P_ERREUR

End Function

Private Function ajouter_non_rens(ByVal v_key As String, _
                                  ByRef r_nd As Node) As Integer

    Dim sql As String
    Dim num_srv_poste As Integer
    Dim rs As rdoResultset
    Dim nd As Node

    If left$(v_key, 1) = "S" Then
        sql = "SELECT SRV_NumPere FROM Service WHERE SRV_Num=" & Mid(v_key, 2)
    ElseIf left$(v_key, 1) = "P" Then
        sql = "SELECT PO_SRVNum FROM Poste WHERE PO_Num=" & Mid(v_key, 2)
    End If

    If Odbc_RecupVal(sql, num_srv_poste) = P_ERREUR Then
        ajouter_non_rens = P_ERREUR
        Exit Function
    End If

    If num_srv_poste = 0 Then ' Le service est rattaché directement au site
        Set r_nd = tv(TV_PROFIL_REDUIT).Nodes(CLEF_ROOT)
        Exit Function
    Else ' Le service est rattaché à un autre service
        If TV_NodeExiste(tv(TV_PROFIL_REDUIT), "S" & num_srv_poste, nd) = P_NON Then
            ' Le service père n'est pas affiché -> on l'ajoute
            Call ajouter_service(1, num_srv_poste, "", "")
        End If
        Set r_nd = tv(TV_PROFIL_REDUIT).Nodes("S" & num_srv_poste)
    End If

    ajouter_non_rens = P_OK
    
End Function

Private Function ajouter_service(ByVal v_indtv As Integer, _
                                 ByVal v_numsrv As Long, _
                                 ByVal v_app_profil_conc As String, _
                                 ByVal v_app_profil_nonconc As String) As Integer

    Dim sql As String
    Dim rs As rdoResultset
    Dim image_color As Integer
    Dim nd As Node, ndP As Node

    sql = "SELECT SRV_Num, SRV_LNum, SRV_NumPere, SRV_Nom FROM Service" _
        & " WHERE SRV_Num=" & v_numsrv
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        ajouter_service = P_ERREUR
        Exit Function
    End If
    If TV_NodeExiste(tv(v_indtv), "S" & rs("SRV_Num").Value, nd) = P_OUI Then
        ajouter_service = P_OK
        Exit Function
    End If
    If rs("SRV_NumPere").Value > 0 Then
        If TV_NodeExiste(tv(v_indtv), "S" & rs("SRV_NumPere").Value, nd) = P_NON Then
            Call ajouter_service(v_indtv, rs("SRV_NumPere").Value, v_app_profil_conc, v_app_profil_nonconc)
        End If
        Set ndP = tv(v_indtv).Nodes("S" & rs("SRV_NumPere").Value)
    Else
        Set ndP = tv(v_indtv).Nodes(CLEF_ROOT)
    End If

    If v_app_profil_conc = "T" Then
        image_color = IMG_SRV_SEL
    ElseIf v_app_profil_nonconc = "T" Then
        image_color = IMG_SRV_NOSEL
    Else
        If InStr(v_app_profil_conc, "S" & v_numsrv & ";") > 0 Then
            image_color = IMG_SRV_SEL
        ElseIf InStr(v_app_profil_nonconc, "S" & v_numsrv & ";") > 0 Then
            image_color = IMG_SRV_NOSEL
        Else
            image_color = IMG_SRV
        End If
    End If
'    Set nd = tv(v_indtv).Nodes.Add(ndp, _
'                           tvwChild, _
'                           "S" & rs("SRV_Num").Value, _
'                           rs("SRV_Nom").Value, _
'                           image_color, _
'                           image_color)
    Set nd = tv(v_indtv).Nodes.Add(ndP, _
                           tvwChild, _
                           "S" & rs("SRV_Num").Value, _
                           rs("SRV_Nom").Value, _
                           IMG_SRV, _
                           IMG_SRV)
    nd.Sorted = True

    ajouter_service = P_OK

End Function

Private Sub ajouter_type_coord(ByVal v_indgrd As Integer)
    ' Commun pour ENREGISTRER et pour PREVENIR

    Dim sql As String, sret As String, new_code As String, new_libelle As String
    Dim is_selected As Boolean
    Dim i As Integer, col_num As Integer
    Dim col_code As Integer, col_libelle As Integer
    Dim num As Long
    Dim rs As rdoResultset

    num = 0
    is_selected = False

    ' Choix du GRID
    If v_indgrd = GRD_RENSEIGNER Then
        col_num = GRDR_NUM
        col_code = GRDR_CODE
        col_libelle = GRDR_LIBELLE
    Else ' Index = CMD_PLUS_ZONE_PREV
        col_num = GRDP_NUM
        col_code = GRDP_CODE
        col_libelle = GRDP_LIBELLE
    End If

    Call CL_Init
    Call CL_InitMultiSelect(True, False) 'selection multiple=True, retourner la ligne courante=False
    Call CL_InitTitreHelp("Liste des types de coordonnée", p_chemin_appli + "\help\kalidoc.chm" & ";" & "dico_d_fonction.htm")
    Call CL_InitTaille(0, -15)

    ' Boucle SQL d'ajout dans la liste des choix
    sql = "SELECT ZU_Num, ZU_Code, ZU_Libelle FROM ZoneUtil"
    If Odbc_Select(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    While Not rs.EOF
        ' Recherche si la coordonées existe déjà dans le GRID
        With grdAppli(v_indgrd)
            For i = 0 To .Rows - 1
                ' elle existe déjà dans le GRID => alors la cocher
                If rs("ZU_Num").Value = .TextMatrix(i, col_num) _
                        And rs("ZU_Code").Value = .TextMatrix(i, col_code) Then
                    is_selected = True
                    Exit For
                Else ' ne pas cocher
                    is_selected = False
                End If
            Next i
        End With
        Call CL_AddLigne(rs("ZU_Code").Value, rs("ZU_Num").Value, LETTRE_ZONE & rs("ZU_Libelle").Value, is_selected)
        rs.MoveNext
    Wend
    rs.Close

    ' uniquement pour les zones modifiables
    If v_indgrd = GRD_PREVENIR Then
        ' Boucle SQL d'ajout des type d'info suppl dans la liste des choix
        sql = "SELECT KB_TisNum, KB_TisCode, KB_TisLibelle FROM KB_TypeInfoSuppl" & _
            " ORDER BY KB_TisCode"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Exit Sub
        End If
        While Not rs.EOF
            ' Recherche si le type d'info suppl existe déjà dans le GRID
            With grdAppli(v_indgrd)
                For i = 0 To .Rows - 1
                    ' elle existe déjà dans le GRID => alors la cocher
                    If rs("KB_TisNum").Value = .TextMatrix(i, col_num) _
                            And rs("KB_TisCode").Value = .TextMatrix(i, col_code) Then
                        is_selected = True
                        Exit For
                    Else ' ne pas cocher
                        is_selected = False
                    End If
                Next i
            End With
            Call CL_AddLigne(rs("KB_TisCode").Value, rs("KB_TisNum").Value, LETTRE_INFO_SUPPL & rs("KB_TisLibelle").Value, is_selected)
            rs.MoveNext
        Wend
        rs.Close
    End If

    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("&Ajouter un type de coordonnée", "", 0, 0, 2000)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)

lab_afficher: ' ajouter la coodonnée nouvellement créée
    ChoixListe.Show 1

    ' ******************************** QUITTER ********************************
    If CL_liste.retour = 2 Then
        Exit Sub
    End If
    ' ***************** CRÉER UN NOUVEAU TYPE DE COORDONNÉES ******************
    If CL_liste.retour = 1 Then
        sret = PrmTypeCoordonnees.AppelFrm(0)
        If sret <> "" Then ' on a créé une corrdonnée
            ' recupérer son numéro
            num = STR_GetChamp(sret, "|", 0)
            If Odbc_RecupVal("SELECT ZU_Code, ZU_Libelle FROM ZoneUtil WHERE ZU_Num=" & num, _
                              new_code, new_libelle) = P_ERREUR Then
                Exit Sub
            End If
            Call CL_AddLigne(new_code, num, new_libelle, True)
        End If
        ' Revenir au débt du choix tout en cochant la coordonnée créée
        GoTo lab_afficher
    End If
    ' ****************************** ENREGISTRER ******************************
    grdAppli(v_indgrd).Rows = 0
    If CL_liste.retour = 0 Then
        For i = 0 To UBound(CL_liste.lignes)
            If CL_liste.lignes(i).selected Then
                grdAppli(v_indgrd).AddItem (CL_liste.lignes(i).num & vbTab)
                grdAppli(v_indgrd).TextMatrix(grdAppli(v_indgrd).Rows - 1, col_code) = CL_liste.lignes(i).texte
                grdAppli(v_indgrd).TextMatrix(grdAppli(v_indgrd).Rows - 1, col_libelle) = Mid$(CL_liste.lignes(i).tag, 2)
                If v_indgrd = GRD_PREVENIR Then
                    grdAppli(v_indgrd).TextMatrix(grdAppli(v_indgrd).Rows - 1, GRDP_TYPE) = left$(CL_liste.lignes(i).tag, 1)
                End If
                If v_indgrd = GRD_RENSEIGNER Then
                    cmd(CMD_MOINS_ZONE_RENS).Visible = True
                Else
                    cmd(CMD_MOINS_ZONE_PREV).Visible = True
                End If
            End If
        Next i
        cmd(CMD_OK).Enabled = True
    End If
    ' *************************************************************************

End Sub

' Appliquer le choix à toutes les fonctions du nême nom
Private Function appliquer_a_toutes_fct()
    
    Dim sql As String, question As String
    Dim i As Integer, img As Integer, choix As Integer
    Dim po_num As Long, ft_num As Long
    Dim rs As rdoResultset
    Dim nd As Node

    With tv(g_tv_actuel)
        po_num = Mid$(.SelectedItem.key, 2)
        question = "Êtes-vous sûr de vouloir appliquer l'état '"
        Select Case .SelectedItem.image
            Case IMG_POSTE
                question = question & "Non renseigné"
            Case IMG_POSTE_SEL
                question = question & "Concerné"
            Case IMG_POSTE_NOSEL
                question = question & "Non concerné"
        End Select
        question = question & "' à toutes les fonctions:" & vbCrLf _
                    & tv(g_tv_actuel).SelectedItem.Text & " ?"
        If MsgBox(question, vbQuestion + vbYesNo) = vbNo Then
            appliquer_a_toutes_fct = P_OK
            Exit Function
        End If

        sql = "SELECT Po_FtNum FROM Poste WHERE Po_Num=" & po_num
        If Odbc_RecupVal(sql, ft_num) = P_ERREUR Then
            GoTo lab_erreur
        End If
        sql = "SELECT Po_Num FROM Poste WHERE Po_FtNum=" & ft_num
        If Odbc_Select(sql, rs) = P_ERREUR Then
            GoTo lab_erreur
        End If
        ' Parcourir le treeview pour maj les noeuds
        While Not rs.EOF
            img = .SelectedItem.image
            For i = 1 To tv(TV_PROFIL_TOUS).Nodes.Count
                Set nd = tv(TV_PROFIL_TOUS).Nodes(i)
                ' traiter uniquement les postes
                If left$(nd.key, 1) = "P" And Mid$(nd.key, 2) = rs("Po_Num").Value Then
                    nd.image = img
                    nd.SelectedImage = img
                    Call fix_couleur_parent(TV_PROFIL_TOUS, nd)
                End If
            Next i
            rs.MoveNext
        Wend
        rs.Close
    End With
    
    If g_tv_actuel <> TV_PROFIL_TOUS Then
        If opt(OPT_CONC).Value = True Then
            choix = OPT_CONC
        ElseIf opt(OPT_NON_CONC).Value = True Then
            choix = OPT_NON_CONC
        Else
            choix = OPT_NON_RENS
        End If
        Call init_tv_reduit(choix)
    End If
    
    appliquer_a_toutes_fct = P_OK
    Exit Function

lab_erreur:
    appliquer_a_toutes_fct = P_ERREUR
End Function

Private Sub basculer_etat_sp(ByVal v_index As Integer)

    Dim traiter As Boolean
    Dim image As Integer, img_p As Integer, img_s As Integer, img_st As Integer, i As Integer, _
        image_noeud_1 As Integer, image_noeud_2 As Integer
    Dim nd As Node, ndF As Node, nd2 As Node, nd_tout As Node

    If tv(v_index).Nodes.Count < 1 Then Exit Sub

    Set nd = tv(v_index).SelectedItem

    ' Changer l'image selon l'état précédent...
    Select Case nd.image
        Case IMG_SRV
            img_st = IMG_SITE
            img_s = IMG_SRV_SEL
            img_p = IMG_POSTE_SEL
        Case IMG_SRV_SEL
            img_st = IMG_SITE
            img_s = IMG_SRV_NOSEL
            img_p = IMG_POSTE_NOSEL
        Case IMG_SRV_NOSEL
            img_st = IMG_SITE
            img_s = IMG_SRV
            img_p = IMG_POSTE
        Case IMG_POSTE
            img_st = IMG_SITE
            img_p = IMG_POSTE_SEL
        Case IMG_POSTE_SEL
            img_st = IMG_SITE
            img_p = IMG_POSTE_NOSEL
        Case IMG_POSTE_NOSEL
            img_st = IMG_SITE
            img_p = IMG_POSTE
        Case IMG_SITE
            img_st = IMG_SITE_SEL
            img_s = IMG_SRV_SEL
            img_p = IMG_POSTE_SEL
        Case IMG_SITE_SEL
            img_st = IMG_SITE_NOSEL
            img_s = IMG_SRV_NOSEL
            img_p = IMG_POSTE_NOSEL
        Case IMG_SITE_NOSEL
            img_st = IMG_SITE
            img_s = IMG_SRV
            img_p = IMG_POSTE
    End Select

    ' ... changement effectif
    If left$(nd.key, 1) = "P" Then
        image = img_p
    ElseIf left$(nd.key, 1) = "S" Then
        image = img_s
    ElseIf left$(nd.key, 1) = "L" Then
        image = img_st
    End If
Set nd_tout = tv(TV_PROFIL_TOUS).Nodes(nd.key)
nd_tout.image = image
nd_tout.SelectedImage = image
    nd.image = image
    nd.SelectedImage = image

    Set ndF = nd

    ' Traiter les sous sections
    While TV_ChildNextParent(ndF, nd)
        traiter = False
        If left$(ndF.key, 1) = "P" Then
            traiter = True
            image = img_p
        ElseIf left$(ndF.key, 1) = "S" Then
            image = img_s
        ElseIf left$(nd.key, 1) = "L" Then
            image = img_st
        End If
        ndF.image = image
        ndF.SelectedImage = image
        If traiter And v_index <> TV_PROFIL_TOUS Then
            Set nd2 = tv(TV_PROFIL_TOUS).Nodes(ndF.key)
            nd2.image = image
            nd2.SelectedImage = image
            Call fix_couleur_parent(TV_PROFIL_TOUS, nd2)
        End If
    Wend

    cmd(CMD_OK).Enabled = True

End Sub

Private Sub build_sp(ByRef r_spc As Variant, _
                     ByRef r_spnc As Variant)

    Dim s As String
    Dim encore As Boolean
    Dim i As Integer
    Dim nd As Node

    If tv(TV_PROFIL_TOUS).Nodes(1).image = IMG_SITE_SEL Then
        r_spc = "T"
        r_spnc = ""
        Exit Sub
    ElseIf tv(TV_PROFIL_TOUS).Nodes(1).image = IMG_SITE_NOSEL Then
        r_spc = ""
        r_spnc = "T"
        Exit Sub
    End If

    r_spc = ""
    r_spnc = ""
    For i = 1 To tv(TV_PROFIL_TOUS).Nodes.Count
        Set nd = tv(TV_PROFIL_TOUS).Nodes(i)
        If nd.image = IMG_SRV_SEL Or nd.image = IMG_POSTE_SEL Then
            If nd.Parent.image <> IMG_SRV_SEL And nd.Parent.image <> IMG_POSTE_SEL Then
                s = nd.key + ";"
                Do
                    If TV_NodeParent(nd) Then
                        If left$(nd.key, 1) = "L" Then
                            encore = False
                        Else
                            s = nd.key + ";" + s
                            encore = True
                        End If
                    Else
                        encore = False
                    End If
                Loop Until encore = False
                If r_spc <> "" Then
                    r_spc = r_spc + "|"
                End If
                r_spc = r_spc + s
            End If
        ElseIf nd.image = IMG_SRV_NOSEL Or nd.image = IMG_POSTE_NOSEL Then
            If nd.Parent.image <> IMG_SRV_NOSEL And nd.Parent.image <> IMG_POSTE_NOSEL Then
                s = nd.key + ";"
                Do
                    If TV_NodeParent(nd) Then
                        If left$(nd.key, 1) = "L" Then
                            encore = False
                        Else
                            s = nd.key + ";" + s
                            encore = True
                        End If
                    Else
                        encore = False
                    End If
                Loop Until encore = False
                If r_spnc <> "" Then
                    r_spnc = r_spnc + "|"
                End If
                r_spnc = r_spnc + s
            End If
        End If
    Next i

    If r_spc <> "" Then r_spc = r_spc + "|"
    If r_spnc <> "" Then r_spnc = r_spnc + "|"

End Sub

Private Function build_zones(ByVal v_indgrd As Integer) As String
' Entree: l'Index du GRID
' Sortie: [APP_ZoneRens] ou [APP_ZonePrev] à inserer dans la table [Application]

    Dim i As Integer

    build_zones = ""

    With grdAppli(v_indgrd)
        For i = 0 To .Rows - 1
            If v_indgrd = GRD_PREVENIR Then
                build_zones = build_zones & .TextMatrix(i, GRDP_TYPE) & .TextMatrix(i, GRDR_NUM) & ";"
            Else
                build_zones = build_zones & LETTRE_ZONE & .TextMatrix(i, GRDR_NUM) & ";"
            End If
        Next i
    End With

End Function

Private Function choisir_appli() As Integer

    Dim sret As String, sql As String
    Dim n As Integer
    Dim nofct As Long
    Dim rs As rdoResultset

    Call FRM_ResizeForm(Me, 0, 0)

lab_affiche:
    Call CL_Init
    'Choix de l'appli
    sql = "SELECT * FROM Application" _
        & " ORDER BY APP_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        choisir_appli = P_ERREUR
        Exit Function
    End If
    Call CL_AddLigne("<Nouvelle>", 0, "", False)
    n = 1
    While Not rs.EOF
        Call CL_AddLigne(rs("APP_Nom").Value, rs("APP_Num").Value, "", False)
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close

    Call CL_InitTitreHelp("Liste des applications", "")
    Call CL_InitTaille(0, -15)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    ChoixListe.Show 1
    ' Quitter
    If CL_liste.retour = 1 Then
        choisir_appli = P_NON
        Exit Function
    End If

    If afficher_appli(CL_liste.lignes(CL_liste.pointeur).num) = P_ERREUR Then
        choisir_appli = P_ERREUR
        Exit Function
    End If

    choisir_appli = P_OUI

End Function

Private Sub choisir_resp()

    Dim sret As String, sql As String, nomutil As String
    Dim numutil As Long
    Dim frm As Form

    p_siz_tblu = -1
    Set frm = ChoixUtilisateur
    sret = ChoixUtilisateur.AppelFrm("Choix de la personne à prévenir", _
                                    "", _
                                    False, _
                                    False, _
                                    "")
    Set frm = Nothing
    If sret = "" Then
        Exit Sub
    End If
    numutil = p_tblu_sel(0)
    Call P_RecupUtilNomP(numutil, nomutil)
    grdAppli(GRD_RESP).AddItem numutil & vbTab & nomutil
    cmd(CMD_MOINS_RESP).Visible = True
    
    cmd(CMD_OK).Enabled = True

End Sub

Private Function enregistrer_appli() As Integer

    Dim str_rens_champs As String, str_prev_champs As String, _
        chemin_kw As String, old_code As String, new_code As String, lstresp As String
    Dim i As Integer
    Dim s_spc As Variant, s_spnc As Variant

    Call build_sp(s_spc, s_spnc)

    ' constructiondes champs [APP_ZoneRens] et [APP_ZonePrev]
    str_rens_champs = build_zones(GRD_RENSEIGNER)
    str_prev_champs = build_zones(GRD_PREVENIR)
    chemin_kw = p_CheminKW & "/kalibottin/mouvements/"
    new_code = UCase$(txt(TXT_CODE).Text)

    lstresp = ""
    For i = 0 To grdAppli(GRD_RESP).Rows - 1
        lstresp = lstresp & "U" & grdAppli(GRD_RESP).TextMatrix(i, GRDRESP_NUM) & ";"
    Next i
    
    If g_numappli = 0 Then ' ajouter un nouvel enregistrement dans la table [Application]
        If Odbc_AddNew("Application", _
                       "APP_Num", _
                       "APP_Seq", _
                       True, _
                       g_numappli, _
                       "APP_Code", new_code, _
                       "APP_Nom", txt(TXT_NOM).Text, _
                       "APP_lstResp", lstresp, _
                       "APP_Profil_Conc", s_spc, _
                       "APP_Profil_NonConc", s_spnc, _
                       "APP_ZoneRens", str_rens_champs, _
                       "APP_ZonePrev", str_prev_champs, _
                       "APP_Informer", (chk.Value = vbChecked)) = P_ERREUR Then
            GoTo lab_erreur
        End If
        ' créer le repértoire pour la gestion des mouvements s'il n'existe pas
        If Not KF_EstRepertoire(chemin_kw & "/" & new_code, False) Then
            Call KF_CreerRepertoire(chemin_kw & "/" & new_code)
        End If
    Else ' mettre à jour la table [Application] avec de nouvelles modifications
        ' récupérer l'ancien code avant de le modifier
        If Odbc_RecupVal("SELECT APP_Code FROM Application WHERE APP_Num=" & g_numappli, old_code) = P_ERREUR Then
            GoTo lab_erreur
        End If
        If Odbc_Update("Application", _
                       "APP_Num", _
                       "WHERE APP_Num=" & g_numappli, _
                       "APP_Code", new_code, _
                       "APP_Nom", txt(TXT_NOM).Text, _
                       "APP_lstResp", lstresp, _
                       "APP_Profil_Conc", s_spc, _
                       "APP_Profil_NonConc", s_spnc, _
                       "APP_ZoneRens", str_rens_champs, _
                       "APP_ZonePrev", str_prev_champs, _
                       "APP_Informer", (chk.Value = vbChecked)) = P_ERREUR Then
            GoTo lab_erreur
        End If
        ' Exclure KALIDOC et KALIBOTTIN de la création des repértoires
        If new_code <> "KALIDOC" And new_code <> "KALIBOTTIN" Then
            ' créer le repértoire pour la gestion des mouvements s'il n'existe pas
            If Not KF_EstRepertoire(chemin_kw & "/" & old_code, False) Then
                Call KF_CreerRepertoire(chemin_kw & "/" & old_code)
            ' sinon, renommer l'ancien repértoire si le code diffère
            ElseIf old_code <> new_code Then
                If KF_RenommerFichier(chemin_kw & "/" & old_code, chemin_kw & "/" & new_code) = P_ERREUR Then
                    GoTo lab_erreur
                End If
            End If
        End If
    End If

    enregistrer_appli = P_OK
    Exit Function

lab_erreur:
    enregistrer_appli = P_ERREUR

End Function


Private Function est_non_rens(ByRef v_noeud As Node) As Boolean

    If v_noeud.Children = 0 Then
        If v_noeud.SelectedImage = IMG_POSTE Or v_noeud.SelectedImage = IMG_SRV Or v_noeud.SelectedImage = IMG_SITE Then
            est_non_rens = True
            Exit Function
        End If
    Else
        Set v_noeud = v_noeud.Child
        est_non_rens = est_non_rens(v_noeud)
        Exit Function
    End If

    est_non_rens = False

End Function

Private Sub evalue_btninfosuppl()

    Dim sql As String, app_infosuppl As String
    
    sql = "SELECT App_InfoSuppl FROM Application WHERE App_Num=" & g_numappli
    If Odbc_RecupVal(sql, app_infosuppl) = P_ERREUR Then
        Exit Sub
    End If

    If app_infosuppl = "" Then
        Set cmd(CMD_TYPE_INFO_SUPPL).Picture = imglst.ListImages(IMG_PAS_INFO).Picture
        cmd(CMD_TYPE_INFO_SUPPL).ToolTipText = "Aucune information supplémentaire à transmettre à cette application"
    Else
        Set cmd(CMD_TYPE_INFO_SUPPL).Picture = imglst.ListImages(IMG_INFO).Picture
        cmd(CMD_TYPE_INFO_SUPPL).ToolTipText = "Il y a des informations supplémentaires à transmettre à cette application"
    End If
    
End Sub

Private Sub fix_couleur_parent(ByVal v_index As Integer, v_noeud As Node)

    Dim racine As Boolean
    Dim i As Integer, j As Integer, nbr As Integer, _
        couleur_pivot As Integer, image_parent As Integer
    Dim ndP As Node, ndF As Node, ndp_tout As Node, ndf_tout As Node

    racine = False
    Set ndP = v_noeud
    Set ndp_tout = tv(TV_PROFIL_TOUS).Nodes(v_noeud.key)
    couleur_pivot = ndP.SelectedImage
    nbr = STR_GetNbchamp(ndP.FullPath, "\")
    ' tv(TV_PROFIL_TOUS).Nodes(v_noeud.Key).Text
    For i = 1 To nbr - 1 ' <----- pour les parents
        If i = nbr - 1 Then ' pour la racine
            racine = True
            Set ndP = tv(TV_PROFIL_TOUS).Nodes(CLEF_ROOT).Root
            Set ndp_tout = tv(TV_PROFIL_TOUS).Nodes(CLEF_ROOT).Root
            Set ndF = ndP.Child
            Set ndf_tout = ndp_tout.Child
        Else
            Set ndP = ndP.Parent
            Set ndp_tout = tv(TV_PROFIL_TOUS).Nodes(ndP.key)
            Set ndF = ndP.Child
            Set ndf_tout = ndp_tout.Child
        End If

        If ndp_tout.Children = 0 Then Exit Sub
        For j = 1 To ndp_tout.Children
            Select Case couleur_pivot
                Case IMG_POSTE, IMG_SRV
                    Select Case ndf_tout.SelectedImage
                        Case IMG_POSTE_NOSEL, IMG_POSTE_SEL, IMG_SRV_NOSEL, IMG_SRV_SEL
                            Exit For
                    End Select
                Case IMG_POSTE_NOSEL, IMG_SRV_NOSEL
                    Select Case ndf_tout.SelectedImage
                        Case IMG_POSTE, IMG_POSTE_SEL, IMG_SRV, IMG_SRV_SEL
                            couleur_pivot = IMG_POSTE
                            Exit For
                    End Select
                Case IMG_POSTE_SEL, IMG_SRV_SEL
                    Select Case ndf_tout.SelectedImage
                        Case IMG_POSTE_NOSEL, IMG_POSTE, IMG_SRV_NOSEL, IMG_SRV
                            couleur_pivot = IMG_POSTE
                            Exit For
                    End Select
            End Select
            Set ndf_tout = ndf_tout.Next
        Next j

        ' Affecter la couleur au noeud
        If racine Then
            Select Case couleur_pivot
                Case IMG_POSTE_SEL, IMG_SRV_SEL
                    image_parent = IMG_SITE_SEL
                Case IMG_POSTE_NOSEL, IMG_SRV_NOSEL
                    image_parent = IMG_SITE_NOSEL
                Case IMG_POSTE, IMG_SRV
                    image_parent = IMG_SITE
            End Select
            tv(TV_PROFIL_TOUS).Nodes(ndP.key).image = image_parent
            tv(TV_PROFIL_TOUS).Nodes(ndP.key).SelectedImage = image_parent
            If v_index = 1 Then
                tv(TV_PROFIL_REDUIT).Nodes(ndP.key).image = image_parent
                tv(TV_PROFIL_REDUIT).Nodes(ndP.key).SelectedImage = image_parent
            End If
        Else
            Select Case couleur_pivot
                Case IMG_POSTE_SEL, IMG_SRV_SEL
                    image_parent = IMG_SRV_SEL
                Case IMG_POSTE_NOSEL, IMG_SRV_NOSEL
                    image_parent = IMG_SRV_NOSEL
                Case IMG_POSTE, IMG_SRV
                    image_parent = IMG_SRV
            End Select

            tv(TV_PROFIL_TOUS).Nodes(ndP.key).image = image_parent
            tv(TV_PROFIL_TOUS).Nodes(ndP.key).SelectedImage = image_parent
            If v_index = 1 Then
                If True Then
                    tv(TV_PROFIL_REDUIT).Nodes(ndP.key).image = image_parent
                    tv(TV_PROFIL_REDUIT).Nodes(ndP.key).SelectedImage = image_parent
                End If
            End If
            couleur_pivot = image_parent
        End If
    Next i

End Sub

Private Function get_condition_imp(ByVal v_profil_conc As String) As String
' construire la condition pour la selection des personnes
    Dim sql As String, spm_en_cours As String, le_type As String, num As String
    Dim nbr_i As Integer, i As Integer, j As Integer

    nbr_i = STR_GetNbchamp(v_profil_conc, "|")
    For i = 0 To nbr_i - 1
        'num = Mid$(STR_GetChamp(v_profil_conc, "|", STR_GetNbchamp(le_premier_spm, ";") - v_P_POSTE_ou_P_SERVICE), 2)
        spm_en_cours = STR_GetChamp(v_profil_conc, "|", i)
        le_type = left$(STR_GetChamp(spm_en_cours, ";", STR_GetNbchamp(spm_en_cours, ";") - 1), 1)
        num = Mid$(STR_GetChamp(spm_en_cours, ";", STR_GetNbchamp(spm_en_cours, ";") - 1), 2)
        If Len(sql) = 0 Then
            sql = " WHERE"
        Else
            sql = sql & " OR"
        End If
        sql = sql & " U_SPM LIKE '%" & le_type & num & ";%'"
    Next i

    get_condition_imp = sql

End Function

Private Function get_coordonnees_imp(ByVal v_unum As Long, ByVal v_tbl_zone As Variant, _
                                     ByVal v_tbl_infosuppl As Variant) As String
' cumuler les coordonnées de la personne
    Dim coord As String, valeur As String, sql As String
    Dim rs As rdoResultset
    Dim nbr As Long
    Dim i As Integer

    ' les coordonnées
    For i = 0 To UBound(v_tbl_zone) - 1
        sql = "SELECT COUNT(*) FROM UtilCoordonnee, ZoneUtil" & _
              " WHERE ZU_Num=UC_ZuNum AND UC_Type='U' AND UC_TypeNum=" & v_unum & _
              " AND ZU_Num=" & v_tbl_zone(i)
        If Odbc_Count(sql, nbr) = P_ERREUR Then GoTo lab_erreur
        If nbr > 0 Then
            sql = "SELECT UC_Valeur FROM UtilCoordonnee, ZoneUtil" & _
                  " WHERE ZU_Num=UC_ZuNum AND UC_Type='U' AND UC_TypeNum=" & v_unum & _
                  " AND ZU_Num=" & v_tbl_zone(i)
            If Odbc_RecupVal(sql, valeur) = P_ERREUR Then GoTo lab_erreur
        Else
            valeur = " "
        End If
        If Len(coord) = 0 Then
            coord = valeur
        Else
            coord = coord & vbTab & valeur
        End If
    Next i
    ' les info supple
    For i = 0 To UBound(v_tbl_infosuppl) - 1
        sql = "SELECT COUNT(*) FROM InfoSupplEntite" & _
              " WHERE ISE_TisNum=" & v_tbl_infosuppl(i) & " AND ISE_Type='U' AND ISE_TypeNum=" & v_unum
        If Odbc_Count(sql, nbr) = P_ERREUR Then GoTo lab_erreur
        If nbr > 0 Then
            sql = "SELECT ISE_Valeur FROM InfoSupplEntite" & _
                  " WHERE ISE_TisNum=" & v_tbl_infosuppl(i) & " AND ISE_Type='U' AND ISE_TypeNum=" & v_unum
            If Odbc_RecupVal(sql, valeur) = P_ERREUR Then GoTo lab_erreur
        Else
            valeur = " "
        End If
        If Len(coord) = 0 Then
            coord = valeur
        Else
            coord = coord & vbTab & valeur
        End If
    Next i

    get_coordonnees_imp = coord
    Exit Function

lab_erreur:
    Call quitter(True)

End Function

Private Function get_lib_zone(ByVal v_num_zone) As String
' retourner le nom de la zone
    Dim sql As String, le_code As String, le_type As String, le_num As String

    le_type = left$(v_num_zone, 1)
    le_num = Mid$(v_num_zone, 2)
    If le_type = "Z" Then ' ZoneUtil
        sql = "SELECT ZU_Code FROM ZoneUtil WHERE ZU_Num=" & le_num
    Else ' information supplémentaire
        sql = "SELECT KB_TisCode FROM KB_TypeInfoSuppl WHERE KB_TisNum=" & le_num
    End If
    If Odbc_RecupVal(sql, le_code) = P_ERREUR Then GoTo lab_erreur

    get_lib_zone = le_code
    Exit Function

lab_erreur:
    Call quitter(True)

End Function

' POUR L'INSTANT, n'est pas appelée
Private Sub imprimer()
' exporter la liste des concernées+coordonnées
' vers un fichiers tabulé à ouvrir avec Excel
' REMARQUES:
' - les modif sur les infosuppl ne sont pas envoyées aux responsables
'   il risque d'y avoir des doublons dans le tableau généré si
'   infosuppl existe dans les zonePrev + InfoSuppl
    Dim sql As String, mon_fichier As String, str_entete As String, _
        profil_conc As String, zone_rens As String, zone_prev As String, _
        info_suppl As String, str As String, zone_tmp As String, _
        le_num As String, le_type As String
    Dim generer_ok As Boolean
    Dim tbl_zone() As Long, tbl_infosupp() As Long, num_zone As Long, lnb As Long
    Dim fd As Integer, nbr As Integer, i As Integer, j As Integer
    Dim rs As rdoResultset

    generer_ok = False
    ReDim tbl_zone(0)
    ReDim tbl_infosupp(0)
    ' y'a-t-il des modif à enregistrer ?
    If cmd(CMD_OK).Enabled Then
        If MsgBox("Vous avez apporté des modifications à cette application." _
                  & vbCrLf & vbCrLf & "L'impression ne tiendra pas compte" _
                  & " des modifications non enregistrées !" & vbCrLf & vbCrLf _
                  & "Voulez-vous continuer ?", vbYesNo + vbQuestion, _
                  "Attention") = vbNo Then
            generer_ok = False
            GoTo lab_erreur
        End If
    End If

    ' faire patienter....
    frmPatience.left = (Me.width / 2) - (frmPatience.width / 2)
    frmPatience.Top = (Me.Height / 2) - (frmPatience.Height / 2)
    frmPatience.Visible = True
    pgb.Value = 0

    sql = "SELECT App_Profil_Conc, App_ZoneRens, App_ZonePrev, App_InfoSuppl" & _
          " FROM Application WHERE App_Num=" & g_numappli
    If Odbc_RecupVal(sql, profil_conc, zone_rens, zone_prev, info_suppl) = P_ERREUR Then
        generer_ok = False
        GoTo lab_erreur
    End If
    If profil_conc = "" Then
        Call MsgBox("Aucun profil concerné par cette application", vbInformation + vbOKOnly, "")
        Exit Sub
    End If
    
    ' Ouverture du fichier en ecriture
    mon_fichier = p_chemin_appli & "/tmp/" & "lstAppliImprimer_APP" & _
                  g_numappli & "_UNUM" & p_NumUtil & ".xls"
    ' supprimer s'il existe
    Call KF_EffacerFichier(mon_fichier, False)
    If FICH_OuvrirFichier(mon_fichier, FICH_ECRITURE, fd) = P_ERREUR Then generer_ok = False: GoTo lab_erreur
    ' l'entête du fichier:
    str_entete = "Nom:" & vbTab & "Prénom:" & vbTab & "Matricule:" & vbTab & _
                 "Service:" & vbTab & "Poste:"
    ' le tableau des coordonnées
    ' les zone à renseigner
    nbr = STR_GetNbchamp(zone_rens, ";")
    ' virer les ZonUtil non coordonnées => 'X'
    zone_tmp = ""
    For i = 0 To nbr - 1
        le_num = Mid$(STR_GetChamp(zone_rens, ";", i), 2)
        sql = "SELECT COUNT(*) FROM ZoneUtil" & _
              " WHERE ZU_Num=" & le_num & _
              " AND ZU_Type='C'"
        If Odbc_Count(sql, lnb) = P_ERREUR Then generer_ok = False: GoTo lab_erreur
        If lnb > 0 Then
            zone_tmp = zone_tmp & "Z" & le_num & ";"
        End If
    Next i
    zone_rens = zone_tmp
    nbr = STR_GetNbchamp(zone_rens, ";")
    For i = 0 To nbr - 1
        ReDim Preserve tbl_zone(i)
        tbl_zone(i) = Mid$(STR_GetChamp(zone_rens, ";", i), 2)
        str_entete = str_entete & vbTab & get_lib_zone(STR_GetChamp(zone_rens, ";", i)) & ":"
    Next i
    ' les zone à prevenir
    nbr = STR_GetNbchamp(zone_prev, ";")
    ' virer les ZonUtil non coordonnées => 'X'
    zone_tmp = ""
    For i = 0 To nbr - 1
        le_type = left$(STR_GetChamp(zone_prev, ";", i), 1)
        le_num = Mid$(STR_GetChamp(zone_prev, ";", i), 2)
        If le_type = LETTRE_ZONE Then
            sql = "SELECT COUNT(*) FROM ZoneUtil" & _
                  " WHERE ZU_Num=" & le_num & _
                  " AND ZU_Type='C'"
            If Odbc_Count(sql, lnb) = P_ERREUR Then generer_ok = False: GoTo lab_erreur
            If lnb > 0 Then
                zone_tmp = zone_tmp & le_type & le_num & ";"
            End If
        Else ' le_type = LETTRE_INFO_SUPPL
'            zone_tmp = zone_tmp & le_type & le_num & ";"
        End If
    Next i
    zone_prev = zone_tmp
    nbr = STR_GetNbchamp(zone_prev, ";")
    For i = 0 To nbr - 1
        le_type = left$(STR_GetChamp(zone_prev, ";", i), 1)
        num_zone = Mid$(STR_GetChamp(zone_prev, ";", i), 2)
        If le_type = LETTRE_ZONE Then
            ' exclure ce qui existe déjà dans la zones
            For j = 0 To UBound(tbl_zone)
                If num_zone = tbl_zone(j) Then GoTo lab_suivant_zone
            Next j
            ReDim Preserve tbl_zone(UBound(tbl_zone) + 1)
            tbl_zone(UBound(tbl_zone)) = num_zone
        End If
        str_entete = str_entete & vbTab & get_lib_zone(STR_GetChamp(zone_prev, ";", i)) & ":"
lab_suivant_zone:
    Next i
    ' le tableau de info suppl
    nbr = STR_GetNbchamp(info_suppl, ";")
    For i = 0 To nbr - 1
        ' Il reste à exclure ce qui existe déjà dans les infos suppl
        ' ceci une fois les "envoyer alerte sur modif infosuppl" implementée
        ReDim Preserve tbl_infosupp(i)
        tbl_infosupp(i) = Mid$(STR_GetChamp(info_suppl, ";", i), 2)
        str_entete = str_entete & vbTab & get_lib_zone(STR_GetChamp(info_suppl, ";", i)) & ":"
lab_suivant_infosuppl:
    Next i
    Print #fd, str_entete
    ' la liste des personnes concernées
    If profil_conc <> "" Then
        sql = "SELECT U_Num, U_Nom, U_Prenom, U_Matricule, U_Spm, U_Po_Princ" & _
              " FROM Utilisateur WHERE U_kb_actif=True"
        If profil_conc <> "T" Then
            sql = sql & get_condition_imp(profil_conc)
        End If
        sql = sql & " order by U_Nom, U_Prenom"
        If Odbc_Select(sql, rs) = P_ERREUR Then generer_ok = False: GoTo lab_erreur
        While Not rs.EOF
            If pgb.Value = pgb.Max Then
                pgb.Value = 0
            End If
            pgb.Value = pgb.Value + 1
            str = rs("U_Nom").Value & vbTab & rs("U_Prenom").Value & vbTab & rs("U_Matricule").Value & vbTab & _
                  P_get_lib_srv_poste(P_get_num_srv_poste(rs("U_Spm").Value, P_SERVICE), P_SERVICE) & vbTab & _
                  P_get_lib_srv_poste(P_get_num_srv_poste(rs("U_Spm").Value, P_POSTE), P_POSTE) & vbTab & _
                  get_coordonnees_imp(rs("U_Num").Value, tbl_zone, tbl_infosupp)
            Print #fd, str
            rs.MoveNext
        Wend
        rs.Close
    End If
    generer_ok = True

lab_erreur:
    frmPatience.Visible = False
    ' fermer le file handle
    Close #fd
    ' lancer le fichier avec Excel
    If generer_ok Then
        Call MsgBox("Fichier généré avec succès !", vbOKOnly + vbInformation, "Impression")
        ' A FAIRE - LN
        Call Shell(SYS_GetIni("DOC", "EXCEL", p_nomini) & " " & mon_fichier, vbNormalFocus)
    End If

End Sub

Private Sub init_grid(ByVal v_grid As Integer)

    Dim i As Integer

    For i = 0 To grdAppli(v_grid).Rows - 1
        grdAppli(v_grid).Row = i
        grdAppli(v_grid).col = GRDR_CODE
        grdAppli(v_grid).CellFontBold = True
    Next i
    grdAppli(v_grid).col = GRDR_NUM
    grdAppli(v_grid).SetFocus

End Sub

Private Sub init_tv_reduit(ByVal v_choix As Integer)

    Dim i As Integer, j As Integer, n As Integer, img As Integer, img2 As Integer
    Dim tbl_srv() As String, stext As String
    Dim a_afficher As Boolean
    Dim nd As Node, ndP As Node, ndr As Node, ndp2 As Node, nd2 As Node

    cmd(CMD_RAFRAICHIR).Enabled = False
    tv(TV_PROFIL_REDUIT).Nodes.Clear
    a_afficher = False ' uniquement pour les non renseigné(e)s

    ' Parcourir tous les noeuds
    For i = 1 To tv(TV_PROFIL_TOUS).Nodes.Count
        Set nd = tv(TV_PROFIL_TOUS).Nodes(i)
        ' Traiter uniquement les [P]OSTES et les [S]ERVICES
        If left$(nd.key, 1) = "P" Or left$(nd.key, 1) = "S" Then
            Select Case v_choix
                Case OPT_CONC
                    img = IMG_POSTE_SEL
                    img2 = IMG_SRV_SEL
                    stext = "Postes concernés"
                Case OPT_NON_CONC
                    img = IMG_POSTE_NOSEL
                    img2 = IMG_SRV_NOSEL
                    stext = "Postes non concernés"
                Case OPT_NON_RENS
                    img = IMG_POSTE
                    img2 = IMG_SRV
                    stext = "Postes non renseignés"
                Case OPT_TOUS
            End Select
            If nd.image = img Or nd.image = img2 Then
                ReDim tbl_srv(0) As String
                tbl_srv(0) = nd.key
                n = 0
                Set ndP = nd
                While TV_NodeParent(ndP)
                    n = n + 1
                    ReDim Preserve tbl_srv(n) As String
                    tbl_srv(n) = ndP.key
                Wend
                
                ' Afficher le noeud racine s'il existe des sous éléments
                If TV_NodeExiste(tv(TV_PROFIL_REDUIT), nd.Root.key, ndr) = P_NON Then
                    Set nd2 = tv(TV_PROFIL_REDUIT).Nodes.Add(, tvwChild, nd.Root.key, stext, IMG_SITE, IMG_SITE)
                    'Set nd2 = tv(TV_PROFIL_REDUIT).Nodes.Add(, tvwChild, nd.Root.key, stext, nd.Root.image, nd.Root.image)
                    nd2.Expanded = True
                Else
                    Set nd2 = ndr
                End If

                ' Attacher les bons éléments selectionnés selon l'option choisie (concernés, non concernés, ...)
                For j = UBound(tbl_srv()) - 1 To 0 Step -1
                    If TV_NodeExiste(tv(TV_PROFIL_REDUIT), tbl_srv(j), ndr) = P_NON Then
                        Set nd = tv(TV_PROFIL_TOUS).Nodes(tbl_srv(j))
                        ' Les éléments concernés
                        If v_choix = OPT_CONC Then
                            If nd.image = IMG_SRV_SEL Or nd.image = IMG_POSTE_SEL Or nd.image = IMG_SITE_SEL Then
                                Set nd2 = tv(TV_PROFIL_REDUIT).Nodes.Add(nd2, tvwChild, nd.key, nd.Text, nd.image, nd.image)
                            Else ' afin d'attacher les pères non séléctionnés
                                Select Case left$(nd.key, 1)
                                Case "P"
                                    Set nd2 = tv(TV_PROFIL_REDUIT).Nodes.Add(nd2, tvwChild, nd.key, nd.Text, IMG_POSTE, IMG_POSTE)
                                Case "S"
                                    Set nd2 = tv(TV_PROFIL_REDUIT).Nodes.Add(nd2, tvwChild, nd.key, nd.Text, IMG_SRV, IMG_SRV)
                                End Select
                            End If
                        ' Les éléments non concernés
                        ElseIf v_choix = OPT_NON_CONC Then
                            If nd.image = IMG_SRV_NOSEL Or nd.image = IMG_POSTE_NOSEL Or nd.image = IMG_SITE_NOSEL Then
                                Set nd2 = tv(TV_PROFIL_REDUIT).Nodes.Add(nd2, tvwChild, nd.key, nd.Text, nd.image, nd.image)
                            Else ' afin d'attacher les pères non séléctionnés
                                Select Case left$(nd.key, 1)
                                Case "P"
                                    Set nd2 = tv(TV_PROFIL_REDUIT).Nodes.Add(nd2, tvwChild, nd.key, nd.Text, IMG_POSTE, IMG_POSTE)
                                Case "S"
                                    Set nd2 = tv(TV_PROFIL_REDUIT).Nodes.Add(nd2, tvwChild, nd.key, nd.Text, IMG_SRV, IMG_SRV)
                                End Select
                            End If
                        ' Les éléments non renseignés
                        ElseIf v_choix = OPT_NON_RENS Then
                            If nd.image = IMG_SRV Or nd.image = IMG_POSTE Or nd.image = IMG_SITE Then
                                If nd.Children = 0 Then
                                    a_afficher = True
                                Else
                                    a_afficher = est_non_rens(nd)
                                End If
                                If a_afficher Then
                                    ' ajouter ses parents
                                    If ajouter_non_rens(nd.key, nd2) = P_ERREUR Then
                                        Call quitter(True)
                                        Exit Sub
                                    End If
                                    Set nd2 = tv(TV_PROFIL_REDUIT).Nodes.Add(nd2, tvwChild, nd.key, nd.Text, nd.image, nd.image) ' le noeud lui même
                                End If
                            Else ' afin d'attacher les pères non séléctionnés
                                Select Case left$(nd.key, 1)
                                Case "P"
                                    Set nd2 = tv(TV_PROFIL_REDUIT).Nodes.Add(nd2, tvwChild, nd.key, nd.Text, IMG_POSTE, IMG_POSTE)
                                Case "S"
                                    Set nd2 = tv(TV_PROFIL_REDUIT).Nodes.Add(nd2, tvwChild, nd.key, nd.Text, IMG_SRV, IMG_SRV)
                                End Select
                            End If
                        Else
                        End If
                    Else
                        Set nd2 = ndr
                    End If
                Next j
            End If
        End If
    Next i
    For i = 1 To tv(TV_PROFIL_REDUIT).Nodes.Count
        If tv(TV_PROFIL_REDUIT).Nodes(i).Children > 0 Then
            tv(TV_PROFIL_REDUIT).Nodes(i).Sorted = True
        End If
    Next i
    cmd(CMD_RAFRAICHIR).Enabled = True

End Sub

Private Sub initialiser()
    Call FRM_ResizeForm(Me, 0, 0)

    g_mode_saisie = False
    g_mode_conc = OPT_TOUS
    g_tv_actuel = TV_PROFIL_TOUS

    grdAppli(GRD_RESP).Rows = 0
    grdAppli(GRD_RESP).ColWidth(GRDRESP_NUM) = 0
    grdAppli(GRD_RESP).ColWidth(GRDRESP_NOM) = grdAppli(GRD_RESP).width
    
    grdAppli(GRD_RENSEIGNER).Rows = 0
    grdAppli(GRD_RENSEIGNER).ColWidth(GRDR_NUM) = 0
    grdAppli(GRD_RENSEIGNER).ColWidth(GRDR_CODE) = grdAppli(GRD_RENSEIGNER).width / 3
    grdAppli(GRD_RENSEIGNER).ColAlignment(GRDR_CODE) = 0
    grdAppli(GRD_RENSEIGNER).ColWidth(GRDR_LIBELLE) = grdAppli(GRD_RENSEIGNER).width - grdAppli(GRD_RENSEIGNER).ColWidth(GRDR_CODE)
    grdAppli(GRD_RENSEIGNER).ColAlignment(GRDR_LIBELLE) = 0

    grdAppli(GRD_PREVENIR).Rows = 0
    grdAppli(GRD_PREVENIR).ColWidth(GRDP_NUM) = 0
    grdAppli(GRD_PREVENIR).ColWidth(GRDP_CODE) = grdAppli(GRD_PREVENIR).width / 3
    grdAppli(GRD_PREVENIR).ColAlignment(GRDP_CODE) = 0
    grdAppli(GRD_PREVENIR).ColWidth(GRDP_LIBELLE) = grdAppli(GRD_PREVENIR).width - grdAppli(GRD_PREVENIR).ColWidth(GRDP_CODE)
    grdAppli(GRD_PREVENIR).ColAlignment(GRDP_LIBELLE) = 0

    If choisir_appli() <> P_OUI Then
        Call quitter(True)
        Exit Sub
    End If

End Sub

Private Sub prm_typeinfosuppl()
' gérer les types d'informations supplémentaires

    Dim sql As String, app_infosuppl As String
    Dim is_selected As Boolean
    Dim i As Integer
    Dim rs As rdoResultset

    Call CL_Init
    Call CL_InitMultiSelect(True, False)
    Call CL_InitTitreHelp("Liste des types d'informations supplémentaires", p_chemin_appli + "\help\kalidoc.chm" & ";" & "dico_d_fonctions.html")
    Call CL_InitTaille(0, -15)

    ' chercher les infos suppl de cette application
    sql = "SELECT App_InfoSuppl FROM Application WHERE App_Num=" & g_numappli
    If Odbc_RecupVal(sql, app_infosuppl) = P_ERREUR Then
        Exit Sub
    End If
    ' La liste de tous les TIS
    sql = "SELECT * FROM KB_TypeInfoSuppl ORDER BY KB_TisLibelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    If rs.EOF Then
        Call MsgBox("Aucun type d'information supplémentaire n'a été trouvé.", vbInformation + vbOKOnly, "")
        rs.Close
        Exit Sub
    End If
    ' cocher les TIS selectionnés
    While Not rs.EOF
        is_selected = False
        ' est-ce que ce TIS existe ?
        If InStr(app_infosuppl, "I" & rs("KB_TisNum").Value & ";") Then
            is_selected = True
        End If
        Call CL_AddLigne(rs("KB_TisLibelle").Value, rs("KB_TisNum").Value, rs("KB_TisCode").Value, is_selected)
        rs.MoveNext
    Wend
    rs.Close

    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)

    ChoixListe.Show 1

    ' ******************************** QUITTER ********************************
    If CL_liste.retour = 1 Then
        Exit Sub
    End If
    ' ******************************** ENREGISTRER ********************************
    app_infosuppl = ""
    For i = 0 To UBound(CL_liste.lignes)
        If CL_liste.lignes(i).selected Then
            app_infosuppl = app_infosuppl & "I" & CL_liste.lignes(i).num & ";"
        End If
    Next i
    cmd(CMD_OK).Enabled = True

    ' MAJ du champ App_InfoSuppl
    If Odbc_Update("Application", "App_Num", "WHERE App_Num=" & g_numappli, _
                    "App_InfoSuppl", app_infosuppl) = P_ERREUR Then
        Exit Sub
    End If

    Call evalue_btninfosuppl
    
End Sub

Private Function quitter(ByVal v_bforce As Boolean) As Boolean

    Dim reponse As Integer
    
    If v_bforce Then
        Unload Me
        quitter = True
        Exit Function
    End If
    
    If cmd(CMD_OK).Visible And cmd(CMD_OK).Enabled Then
        reponse = MsgBox("Des modifications ont été effectuées !" & vbLf & vbLf & "Confirmez-vous l'abandon ?", _
                          vbYesNo + vbDefaultButton2 + vbQuestion)
        If reponse = vbNo Then
            quitter = False
            Exit Function
        End If
    End If

    If choisir_appli() <> P_OUI Then
        quitter = True
        Unload Me
    End If

End Function

Private Sub supprimer()

    Dim old_code As String
    Dim reponse As Integer, cr As Integer
    Dim lnb As Long

    reponse = MsgBox("Confirmez-vous la suppression de cette application ?", vbQuestion + vbYesNo + vbDefaultButton2, "")
    If reponse = vbNo Then
        Exit Sub
    End If
    ' récupérer l'ancien code avant de le supprimer
    If Odbc_RecupVal("SELECT APP_Code FROM Application WHERE APP_Num=" & g_numappli, old_code) = P_ERREUR Then
        Exit Sub
    End If
    ' Maj table [Application]
    If Odbc_Delete("Application", "APP_Num", _
                    "WHERE APP_Num=" & g_numappli, _
                    lnb) = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If
    ' Maj table [UtilMouvement]
    ' supprimer les mouvements opérés par cette application
    Call Odbc_Cnx.Execute("DELETE FROM UtilMouvement WHERE UM_APPNum=" & g_numappli)
    ' supprimer le repértiore des fichiers des mouvements
    If old_code <> "" Then ' eviter de faire des bêtises
        Call KF_EffacerRepertoire(p_CheminKW & "/kalibottin/mouvements/" & old_code)
    End If

    If choisir_appli() <> P_OUI Then Call quitter(True)

End Sub

Private Sub supprimer_resp()

    If grdAppli(GRD_RESP).Rows = 1 Then
        grdAppli(GRD_RESP).Rows = 0
        cmd(CMD_MOINS_RESP).Visible = False
    Else
        grdAppli(GRD_RESP).RemoveItem (grdAppli(GRD_RESP).Row)
    End If
    
    grdAppli(GRD_RESP).SetFocus
    cmd(CMD_OK).Enabled = True

End Sub

Private Sub supprimer_zone_prev()

    If grdAppli(GRD_PREVENIR).Rows = 1 Then
        grdAppli(GRD_PREVENIR).Rows = 0
        cmd(CMD_MOINS_ZONE_PREV).Visible = False
    Else
        grdAppli(GRD_PREVENIR).RemoveItem (grdAppli(GRD_PREVENIR).Row)
    End If

    grdAppli(GRD_PREVENIR).SetFocus
    cmd(CMD_OK).Enabled = True

End Sub

Private Sub supprimer_zone_rens()

    If grdAppli(GRD_RENSEIGNER).Rows = 1 Then
        grdAppli(GRD_RENSEIGNER).Rows = 0
        cmd(CMD_MOINS_ZONE_RENS).Visible = False
    Else
        grdAppli(GRD_RENSEIGNER).RemoveItem (grdAppli(GRD_RENSEIGNER).Row)
    End If

    grdAppli(GRD_RENSEIGNER).SetFocus
    cmd(CMD_OK).Enabled = True

End Sub

Private Sub valider()

    Dim cr As Integer
    
    cr = verif_tous_chp()
    If cr = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If
    If cr = P_NON Then
        Exit Sub
    End If

    Me.MousePointer = 11

    cr = enregistrer_appli()
    Me.MousePointer = 0
    If cr = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If

    If choisir_appli() <> P_OUI Then Call quitter(True)

End Sub

Private Function verif_code() As Integer

    Dim sql As String
    Dim rs As rdoResultset

    If txt(TXT_CODE).Text = "" Then
        Call MsgBox("Le CODE de l'application est une rubrique obligatoire.", vbOKOnly + vbExclamation, "")
        txt(TXT_CODE).SetFocus
        GoTo lab_non
    Else ' txt(TXT_CODE).Text <> ""
        sql = "SELECT APP_Num FROM Application" _
            & " WHERE APP_Code=" & Odbc_String(txt(TXT_CODE).Text)
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            GoTo lab_non
        End If
        If Not rs.EOF Then
            If rs("APP_Num").Value <> g_numappli Then
                rs.Close
                Call MsgBox("Application déjà existante.", vbOKOnly + vbExclamation, "")
                GoTo lab_non
            End If
        End If
        rs.Close
    End If

    verif_code = P_OUI
    Exit Function

lab_non:
    verif_code = P_NON

End Function

Private Function verif_tous_chp() As Integer

    If verif_code() <> P_OUI Then
        txt(TXT_CODE).SetFocus
        GoTo lab_non
    End If

    If txt(TXT_NOM).Text = "" Then
        Call MsgBox("Le NOM de l'application est une rubrique obligatoire.", vbOKOnly + vbExclamation, "")
        txt(TXT_NOM).SetFocus
        GoTo lab_non
    End If

    If grdAppli(GRD_RESP).Rows = 0 Then
        Call MsgBox("La PERSONNE A PREVENIR est une rubrique obligatoire.", vbOKOnly + vbExclamation, "")
        grdAppli(GRD_RESP).SetFocus
        GoTo lab_non
    End If

    verif_tous_chp = P_OUI
    Exit Function

lab_non:
    verif_tous_chp = P_NON

End Function

Private Sub chk_Click()

    If Not g_mode_saisie Then Exit Sub
    
    If chk.Value = 1 Then
        frm(FRM_PREVENIR).Visible = True
        cmd(CMD_TYPE_INFO_SUPPL).Visible = True
    Else
        frm(FRM_PREVENIR).Visible = False
        cmd(CMD_TYPE_INFO_SUPPL).Visible = False
    End If
    
    cmd(CMD_OK).Enabled = True
    
End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case CMD_PLUS_RESP
            Call choisir_resp
        Case CMD_MOINS_RESP
            Call supprimer_resp
        Case CMD_OK
            Call valider
        Case CMD_DETRUIRE
            Call supprimer
        Case CMD_QUITTER
            Call quitter(False)
        Case CMD_PLUS_ZONE_RENS
            Call ajouter_type_coord(GRD_RENSEIGNER)
            ' pour remettre la police en gras et le setFocus
            Call init_grid(GRD_RENSEIGNER)
        Case CMD_PLUS_ZONE_PREV
            Call ajouter_type_coord(GRD_PREVENIR)
            ' pour remettre la police en gras et le setFocus
            Call init_grid(GRD_PREVENIR)
        Case CMD_MOINS_ZONE_RENS
            Call supprimer_zone_rens
        Case CMD_MOINS_ZONE_PREV
            Call supprimer_zone_prev
        Case CMD_RAFRAICHIR
            Call init_tv_reduit(g_mode_conc)
        Case CMD_RENSEIGNER
            If appliquer_a_toutes_fct() = P_ERREUR Then
                Call quitter(True)
            End If
        Case CMD_TYPE_INFO_SUPPL
            Call prm_typeinfosuppl
        Case CMD_IMPRIMER
            Call imprimer
    End Select

End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If Index = CMD_QUITTER Then g_mode_saisie = False

End Sub

Private Sub Form_Activate()

    If g_form_active Then Exit Sub

    g_form_active = True
    Call initialiser

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If (KeyCode = vbKeyE And Shift = vbAltMask) Or KeyCode = vbKeyF1 Then
        KeyCode = 0
        If cmd(CMD_OK).Enabled Then Call valider
    ElseIf (KeyCode = vbKeyS And Shift = vbAltMask) Or KeyCode = vbKeyF2 Then
        KeyCode = 0
        If cmd(CMD_DETRUIRE).Enabled Then
            Call supprimer
        End If
    ElseIf KeyCode = vbKeyH And Shift = vbAltMask Then
        KeyCode = 0
        Call HtmlHelp(0, p_chemin_appli + "\help\kalidoc.chm", HH_DISPLAY_TOPIC, "dico_d_fonction.htm")
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    ElseIf KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Call quitter(False)
    End If

End Sub

Private Sub Form_Load()

    g_form_active = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        If Not quitter(False) Then
            Cancel = True
        End If
    End If

End Sub

Private Sub grdAppli_Click(Index As Integer)

    With grdAppli(Index)
        .ToolTipText = ""
    End With

End Sub

Private Sub opt_Click(Index As Integer)

    g_mode_conc = Index
    If Index = OPT_TOUS Then
        g_tv_actuel = TV_PROFIL_TOUS
    Else
        g_tv_actuel = TV_PROFIL_REDUIT
    End If
    ' cacher le boutonRenseigner pour les applications
    cmd(CMD_RENSEIGNER).Enabled = False
    cmd(CMD_RENSEIGNER).ToolTipText = "Veuillez selectionner une fonction"

    If Not g_mode_saisie Then
        Exit Sub
    End If

    If Index <> OPT_TOUS Then ' 3 > Index > 0
        Select Case Index
        Case OPT_CONC
            pct(PCT_SEL).Picture = imglst.ListImages(IMG_VERT_COCHE).Picture
            pct(PCT_NOSEL).Picture = LoadPicture("")
            pct(PCT_NORENS).Picture = LoadPicture("")
        Case OPT_NON_CONC
            pct(PCT_SEL).Picture = LoadPicture("")
            pct(PCT_NOSEL).Picture = imglst.ListImages(IMG_ROUGE_COCHE).Picture
            pct(PCT_NORENS).Picture = LoadPicture("")
        Case OPT_NON_RENS
            pct(PCT_SEL).Picture = LoadPicture("")
            pct(PCT_NOSEL).Picture = LoadPicture("")
            pct(PCT_NORENS).Picture = imglst.ListImages(IMG_GRIS_COCHE).Picture
        End Select
        cmd(CMD_RAFRAICHIR).Visible = True
        tv(TV_PROFIL_TOUS).ZOrder 1
        tv(TV_PROFIL_REDUIT).ZOrder 0
        Call init_tv_reduit(Index)
    Else ' Index = OPT_TOUS
        cmd(CMD_RAFRAICHIR).Visible = False
        pct(PCT_SEL).Picture = LoadPicture("")
        pct(PCT_NOSEL).Picture = LoadPicture("")
        pct(PCT_NORENS).Picture = LoadPicture("")
        tv(TV_PROFIL_TOUS).ZOrder 0
        tv(TV_PROFIL_REDUIT).ZOrder 1
    End If

End Sub

Private Sub pct_Click(Index As Integer)

    g_tv_actuel = TV_PROFIL_REDUIT
    cmd(CMD_RAFRAICHIR).Enabled = False
    cmd(CMD_RENSEIGNER).Enabled = False
    cmd(CMD_RENSEIGNER).ToolTipText = "Veuillez sélectionner une fonction"
    Select Case Index
        Case PCT_SEL
            pct(PCT_SEL).Picture = imglst.ListImages(IMG_VERT_COCHE).Picture
            pct(PCT_NOSEL).Picture = LoadPicture("")
            pct(PCT_NORENS).Picture = LoadPicture("")
            opt(OPT_CONC).Value = True
        Case PCT_NOSEL
            pct(PCT_SEL).Picture = LoadPicture("")
            pct(PCT_NOSEL).Picture = imglst.ListImages(IMG_ROUGE_COCHE).Picture
            pct(PCT_NORENS).Picture = LoadPicture("")
            opt(OPT_NON_CONC).Value = True
        Case PCT_NORENS
            pct(PCT_SEL).Picture = LoadPicture("")
            pct(PCT_NOSEL).Picture = LoadPicture("")
            pct(PCT_NORENS).Picture = imglst.ListImages(IMG_GRIS_COCHE).Picture
            opt(OPT_NON_RENS).Value = True
    End Select
    cmd(CMD_RAFRAICHIR).Enabled = True

End Sub

Private Sub tv_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = vbKeySpace Then
        KeyAscii = 0
        Call basculer_etat_sp(Index)
        Call fix_couleur_parent(Index, tv(Index).SelectedItem)
    End If

End Sub

Private Sub tv_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If tv(Index).Nodes.Count = 0 Then Exit Sub

    If Button = MouseButtonConstants.vbLeftButton Then
        tv(TV_PROFIL_TOUS).ToolTipText = ""
    ElseIf Button = MouseButtonConstants.vbRightButton And (Index = TV_PROFIL_TOUS Or left$(tv(Index).SelectedItem.key, 1) <> "L") Then
        Call basculer_etat_sp(Index)
        Call fix_couleur_parent(Index, tv(Index).SelectedItem)
    End If

    ' le bouton Appliquer à toutes les fonctions
    With tv(Index)
        If tv(Index).SelectedItem.image = IMG_POSTE _
                Or .SelectedItem.image = IMG_POSTE_SEL _
                Or .SelectedItem.image = IMG_POSTE_NOSEL Then
'MsgBox tv(Index).SelectedItem
            cmd(CMD_RENSEIGNER).Enabled = True
            cmd(CMD_RENSEIGNER).ToolTipText = "Appliquer la règle à toutes les fonctions: " & .SelectedItem.Text
        Else
            cmd(CMD_RENSEIGNER).Enabled = False
            cmd(CMD_RENSEIGNER).ToolTipText = "Veuillez sélectionner une fonction"
        End If
    End With

End Sub

Private Sub txt_Change(Index As Integer)

    Dim str As String

    ' mettre le CODE en majiscules
    If Index = TXT_CODE Then
        txt(TXT_CODE).Text = UCase$(txt(TXT_CODE).Text)
        txt(TXT_CODE).SelStart = Len(txt(TXT_CODE).Text)
    End If
    ' activer la validation quand-t-on tape quelque chose
    cmd(CMD_OK).Enabled = True

End Sub

Private Sub txt_GotFocus(Index As Integer)

    g_txt_avant = txt(Index).Text

End Sub

Private Sub txt_LostFocus(Index As Integer)

    Dim cr As Integer

    If g_mode_saisie Then
        If txt(Index).Text <> g_txt_avant Then
            If Index = TXT_CODE Then
                cr = verif_code()
                If cr = P_ERREUR Then
                    Call quitter(True)
                    Exit Sub
                End If
                If cr = P_NON Then
                    txt(Index).Text = g_txt_avant
                    txt(Index).SetFocus
                    Exit Sub
                End If
            End If
            cmd(CMD_OK).Enabled = True
        End If
    End If

End Sub
