VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form PrmPersonne 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8175
   ClientLeft      =   1125
   ClientTop       =   1455
   ClientWidth     =   11160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8175
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "   Personne"
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
      Height          =   7530
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   11145
      Begin TabDlg.SSTab sst 
         Height          =   7155
         Left            =   45
         TabIndex        =   20
         Top             =   360
         Width           =   11100
         _ExtentX        =   19579
         _ExtentY        =   12621
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   12632256
         ForeColor       =   8388736
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "&Général"
         TabPicture(0)   =   "PrmPersonne.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lbl(8)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lbl(7)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lbl(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lbl(1)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lbl(3)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lbl(2)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "imglst"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lbl(5)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "grdCoord"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txt(5)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txt(1)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txt(2)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "txt(0)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "txt(3)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "cmd(7)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "cmd(6)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "txt(4)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "cmd(3)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "grdCoordLiees"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "cmd(8)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "cmd(9)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "frm_import"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "cmd(10)"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "chk(1)"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "chk(0)"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "cmd(11)"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).ControlCount=   26
         TabCaption(1)   =   "&Profil"
         TabPicture(1)   =   "PrmPersonne.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblLabo"
         Tab(1).Control(1)=   "ImageListS"
         Tab(1).Control(2)=   "lbl(4)"
         Tab(1).Control(3)=   "grdLabo"
         Tab(1).Control(4)=   "tvSect"
         Tab(1).Control(5)=   "cmd(4)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "cmd(5)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "frmKaliDoc"
         Tab(1).ControlCount=   8
         Begin VB.CommandButton cmd 
            Height          =   495
            Index           =   11
            Left            =   10425
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   4005
            Width           =   495
         End
         Begin VB.CheckBox chk 
            Alignment       =   1  'Right Justify
            Caption         =   "Actif"
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
            Index           =   0
            Left            =   210
            TabIndex        =   41
            Top             =   1410
            Width           =   885
         End
         Begin VB.CheckBox chk 
            Alignment       =   1  'Right Justify
            Caption         =   "Externe"
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
            Index           =   1
            Left            =   1845
            TabIndex        =   40
            Top             =   1410
            Width           =   1005
         End
         Begin VB.CommandButton cmd 
            Height          =   570
            Index           =   10
            Left            =   10230
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   1275
            Width           =   570
         End
         Begin VB.Frame frm_import 
            Caption         =   "Particularités p/r au fichier d'importation"
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
            Height          =   795
            Left            =   240
            TabIndex        =   36
            Top             =   2010
            Width           =   10695
            Begin VB.TextBox txt 
               ForeColor       =   &H00800000&
               Height          =   285
               Index           =   6
               Left            =   1260
               MaxLength       =   50
               TabIndex        =   5
               Top             =   315
               Width           =   2655
            End
            Begin VB.TextBox txt 
               ForeColor       =   &H00800000&
               Height          =   285
               Index           =   7
               Left            =   5655
               MaxLength       =   20
               TabIndex        =   6
               Top             =   315
               Width           =   2175
            End
            Begin VB.CheckBox chk 
               Alignment       =   1  'Right Justify
               Caption         =   "Externe au fichier Import"
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
               Left            =   8115
               TabIndex        =   7
               Top             =   330
               Width           =   2475
            End
            Begin VB.Label lbl 
               Caption         =   "Nom Import"
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
               Left            =   120
               TabIndex        =   38
               Top             =   315
               Width           =   1215
            End
            Begin VB.Label lbl 
               Caption         =   "Prénom Import"
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
               Left            =   4335
               TabIndex        =   37
               Top             =   315
               Width           =   1335
            End
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
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
            Index           =   9
            Left            =   10350
            Picture         =   "PrmPersonne.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   6720
            Width           =   310
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
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
            Index           =   8
            Left            =   10320
            Picture         =   "PrmPersonne.frx":047F
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   5700
            Width           =   310
         End
         Begin MSFlexGridLib.MSFlexGrid grdCoordLiees 
            Height          =   1335
            Left            =   300
            TabIndex        =   11
            Top             =   5700
            Width           =   10035
            _ExtentX        =   17701
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   1
            Cols            =   6
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
            GridColorFixed  =   4194304
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            GridLines       =   2
            ScrollBars      =   2
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Historique"
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
            Index           =   3
            Left            =   9960
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   480
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1036
               SubFormatType   =   0
            EndProperty
            Height          =   405
            Index           =   4
            Left            =   7875
            TabIndex        =   21
            Text            =   "caché"
            Top             =   5385
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00E0E0E0&
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
            Index           =   6
            Left            =   10350
            Picture         =   "PrmPersonne.frx":08D6
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Ajouter une coordonnée."
            Top             =   3120
            Width           =   310
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00E0E0E0&
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
            Index           =   7
            Left            =   10335
            Picture         =   "PrmPersonne.frx":0D2D
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Supprimer la coordonnée selectionnée pour cette personne."
            Top             =   5055
            Width           =   310
         End
         Begin VB.Frame frmKaliDoc 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   2265
            Left            =   -74790
            TabIndex        =   33
            Top             =   4560
            Width           =   8115
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   4560
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   3
            Top             =   1050
            Width           =   1575
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   900
            MaxLength       =   50
            TabIndex        =   0
            Top             =   540
            Width           =   3735
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   2
            Left            =   900
            MaxLength       =   15
            TabIndex        =   2
            Top             =   1020
            Width           =   1935
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   5
            Left            =   -67620
            Picture         =   "PrmPersonne.frx":1174
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "Supprimer le service"
            Top             =   4230
            UseMaskColor    =   -1  'True
            Width           =   315
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
            Index           =   4
            Left            =   -67620
            Picture         =   "PrmPersonne.frx":15BB
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "Accéder aux services"
            Top             =   1620
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   5820
            MaxLength       =   50
            TabIndex        =   1
            Top             =   540
            Width           =   3855
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   5
            Left            =   7440
            MaxLength       =   20
            TabIndex        =   4
            Top             =   1050
            Width           =   1995
         End
         Begin ComctlLib.TreeView tvSect 
            Height          =   2925
            Left            =   -73500
            TabIndex        =   18
            Top             =   1620
            Width           =   5865
            _ExtentX        =   10345
            _ExtentY        =   5159
            _Version        =   327682
            Indentation     =   0
            LabelEdit       =   1
            Style           =   1
            ImageList       =   "ImageListS"
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
         Begin MSFlexGridLib.MSFlexGrid grdLabo 
            Height          =   765
            Left            =   -73500
            TabIndex        =   17
            Top             =   690
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   1349
            _Version        =   393216
            Cols            =   1
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
            GridColorFixed  =   4194304
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            GridLines       =   2
            ScrollBars      =   2
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
         Begin MSFlexGridLib.MSFlexGrid grdCoord 
            Height          =   2235
            Left            =   300
            TabIndex        =   8
            Top             =   3135
            Width           =   10035
            _ExtentX        =   17701
            _ExtentY        =   3942
            _Version        =   393216
            Cols            =   7
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   8388608
            BackColorFixed  =   8454143
            ForeColorFixed  =   0
            BackColorSel    =   8388608
            ForeColorSel    =   16777215
            BackColorBkg    =   16777215
            GridColor       =   4194304
            GridColorFixed  =   4194304
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            GridLines       =   2
            ScrollBars      =   2
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
            Caption         =   "Coordonnées liées"
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
            Left            =   360
            TabIndex        =   35
            Top             =   5460
            Width           =   2535
         End
         Begin ComctlLib.ImageList imglst 
            Left            =   10410
            Top             =   3540
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   13
            ImageHeight     =   13
            MaskColor       =   12632256
            _Version        =   327682
            BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
               NumListImages   =   6
               BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PrmPersonne.frx":1A12
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PrmPersonne.frx":1D64
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PrmPersonne.frx":20B6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PrmPersonne.frx":2688
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PrmPersonne.frx":2C5A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PrmPersonne.frx":396C
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label lbl 
            Caption         =   "Coordonnées"
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
            Index           =   2
            Left            =   300
            TabIndex        =   34
            Top             =   2865
            Width           =   1215
         End
         Begin VB.Label lbl 
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
            Index           =   3
            Left            =   3240
            TabIndex        =   32
            Top             =   1065
            Width           =   1215
         End
         Begin VB.Label lbl 
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
            Left            =   210
            TabIndex        =   31
            Top             =   540
            Width           =   495
         End
         Begin VB.Label lbl 
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
            Left            =   210
            TabIndex        =   30
            Top             =   1050
            Width           =   555
         End
         Begin VB.Label lbl 
            Caption         =   "Postes"
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
            Left            =   -74640
            TabIndex        =   29
            Top             =   2400
            Width           =   945
         End
         Begin ComctlLib.ImageList ImageListS 
            Left            =   -67560
            Top             =   2460
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   24
            ImageHeight     =   21
            MaskColor       =   12632256
            _Version        =   327682
            BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
               NumListImages   =   4
               BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PrmPersonne.frx":467E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PrmPersonne.frx":4F50
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PrmPersonne.frx":5822
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PrmPersonne.frx":60F4
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label lblLabo 
            Caption         =   "Sites de travail"
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
            Left            =   -74670
            TabIndex        =   28
            Top             =   840
            Width           =   1035
         End
         Begin VB.Label lbl 
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
            Index           =   7
            Left            =   5040
            TabIndex        =   27
            Top             =   540
            Width           =   705
         End
         Begin VB.Label lbl 
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
            Height          =   255
            Index           =   8
            Left            =   6600
            TabIndex        =   26
            Top             =   1080
            Width           =   885
         End
      End
      Begin VB.Image img 
         Height          =   300
         Index           =   0
         Left            =   0
         Picture         =   "PrmPersonne.frx":69C6
         Top             =   0
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   765
      Left            =   0
      TabIndex        =   22
      Top             =   7410
      Width           =   11145
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "PrmPersonne.frx":6E25
         Height          =   510
         Index           =   2
         Left            =   4620
         Picture         =   "PrmPersonne.frx":73B4
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   190
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
         Left            =   8790
         Picture         =   "PrmPersonne.frx":7949
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   190
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "PrmPersonne.frx":7F02
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
         Left            =   480
         Picture         =   "PrmPersonne.frx":845E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   190
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
   Begin VB.Menu mnuPobPrinc 
      Caption         =   "mnuPobPrinc"
      Visible         =   0   'False
      Begin VB.Menu mnuPostePrinc 
         Caption         =   "&Poste principal"
      End
      Begin VB.Menu mnuPosteAttribuerNum 
         Caption         =   "&Attribuer coordonnée"
      End
   End
   Begin VB.Menu mnuFct 
      Caption         =   "mnuFct"
      Visible         =   0   'False
      Begin VB.Menu mnuResp 
         Caption         =   "&Responsable du service"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQitter 
         Caption         =   "&Quitter"
      End
   End
End
Attribute VB_Name = "PrmPersonne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Liste des fichiers graphiques supportés par VisualBasic:
' Graphics formats recognized by Visual Basic include:
' bitmap (.bmp) files, icon (.ico) files, cursor (.cur) files,
' run-length encoded (.rle) files, metafile (.wmf) files,
' enhanced metafiles (.emf), GIF (.gif) files, and JPEG (.jpg) files.
Private Const LISTE_FICHIERS_IMAGES = "*.gif;*.bmp;*.jpg;*.jpeg;"

' Images TreeView des postes/services
Private Const IMGT_SERVICE = 1
Private Const IMGT_SERVICE_RESP = 3
Private Const IMGT_POSTE = 2
Private Const IMGT_POSTE_SEL = 4
' Images MSFlexGrid des coordonnées
Private Const IMG_PASCOCHE = 1
Private Const IMG_COCHE = 2
Private Const IMG_INFO_ON = 3
Private Const IMG_INFO_OFF = 4
Private Const IMG_PHOTO_ON = 5
Private Const IMG_PHOTO_OFF = 6

' Nombre de ligne visibles dans le grdCoord (éditables + fixe)
Private Const NBRMAX_ROWS = 7
' Largeur du grdCoord par defaut
Private Const LARGEUR_GRID_PAR_DEFAUT = 10035
' cmd(PLUS_TYPE ou MOINS_TYPE).Left par defaut
Private Const LEFT_CMD_PAR_DEFAUT = 10350

' Indes des boutons
Private Const CMD_OK = 0
Private Const CMD_QUITTER = 1
Private Const CMD_SUPPRIMER = 2
Private Const CMD_HISTORIQUE = 3
Private Const CMD_ACCES_SPM = 4
Private Const CMD_MOINS_SPM = 5
Private Const CMD_PLUS_TYPE = 6
Private Const CMD_MOINS_TYPE = 7
Private Const CMD_PLUS_COORDLIEES = 8
Private Const CMD_MOINS_COORDLIEES = 9
Private Const CMD_PHOTO = 10
Private Const CMD_INFO_SUPPL = 11

' Index des TextBox
Private Const TXT_NOM = 0
Private Const TXT_PRENOM = 1
Private Const TXT_CODE = 2
Private Const TXT_MPASSE = 3
Private Const TXT_MATRICULE = 5
' TextBox caché
Private Const TXT_CACHE = 4
' TextBox Junon
Private Const TXT_NOM_JUNON = 6
Private Const TXT_PRENON_JUNON = 7

' Index des onglets
Private Const ONGLET_GENERAL = 0
Private Const ONGLET_PROFIL = 1
Private Const CHK_ACTIF = 0
Private Const CHK_EXTERNE = 1
Private Const CHK_EXTERNE_IMPORT = 2

'Index des colonnes du GrdCoord
Private Const GRDC_ZUNUM = 0
Private Const GRDC_TYPE = 1
Private Const GRDC_VALEUR = 2
Private Const GRDC_NIVEAU = 3
Private Const GRDC_PRINCIPAL = 4
Private Const GRDC_COMMENTAIRE = 5
Private Const GRDC_UCNUM = 6
'Index des colonnes du GrdCoordLiees
Private Const GRDCL_CODE = 0
Private Const GRDCL_TYPE = 1
Private Const GRDCL_VALEUR = 2
Private Const GRDCL_NIVEAU = 3
Private Const GRDCL_PRINCIPAL = 4
Private Const GRDCL_IDENTITE = 5

' La position de la forme: "ligne-colonne"
Private g_position_txt_cache As String

' Constantes des largeurs des colonnes du GridCoord
Private Const LARG_TYPE = 2500
Private Const LARG_VALEUR = 905
Private Const LARG_NIVEAU = 1850
Private Const LARG_PRINCIPAL = 1800
Private Const LARG_COMMENTAIRE = 1890

' Index des colonnes du GridLabo
Private Const GRDL_NUMLABO = 0
Private Const GRDL_ESTLABO = 1
' Colonnes visibles
Private Const GRDL_CODLABO = 2
Private Const GRDL_IMG_ESTLABO = 3
Private Const GRDL_LABOPRINC = 4

Private g_txt_avant As String, g_sprm As String, g_spm As String, _
        g_coordonnees_avant As String, g_coordonnees_apres As String
Private g_numutil As Long, g_numfct As Long
Private g_mode_saisie As Boolean
Private g_cbo_avant As Integer, g_button As Integer, g_form_width As Integer, g_form_height As Integer
Private g_form_active As Boolean
' Variables pour [UtilMouvement]
Private g_old_nom As String, g_old_prenom As String, g_old_matricule As String, _
        g_old_code As String, g_old_spm As String
Private g_old_active As Boolean
' Variable pour AppelFrm
Private g_changements_importants As Boolean
' Le numéro du poste principal
Private g_pobPrinc As Long
' Le tableau des types de coordonnées supprimées
Private g_coord_supprimees() As Long
Private g_nbr_coord_supp As Integer
Private g_nom_fich_photo As String

Private Function afficher_coordonnee()

    Dim sql As String, poste As String
    Dim I As Integer
    Dim rs As rdoResultset

    sql = "SELECT * FROM UtilCoordonnee, ZoneUtil" _
        & " WHERE UC_TypeNum=" & g_numutil _
        & " AND UC_Type ='U' AND ZU_Type='C'" _
        & " AND UC_ZUNum=ZU_Num" _
        & " ORDER BY ZU_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        afficher_coordonnee = P_ERREUR
        Exit Function
    End If

    I = 1

    With grdCoord
        ' Affichage des lignes du gridCoord
        While Not rs.EOF
            'Indicateur d'attribution de la coordonnee
            If rs("UC_Lstposte").Value <> "" Then
                poste = " *"
            Else
                poste = "  "
            End If
            .AddItem rs("ZU_Num").Value & vbTab & rs("ZU_Libelle") & poste & vbTab & rs("UC_Valeur") _
                    & vbTab & rs("UC_Niveau") & vbTab & vbTab & rs("UC_Comm") & vbTab & rs("UC_Num").Value
            '.Row = i
            .Row = .Rows - 1
            .col = GRDC_PRINCIPAL
            .CellPictureAlignment = 4
            If rs("UC_Principal").Value Then
                Set .CellPicture = imglst.ListImages(IMG_COCHE).Picture
            Else
                Set .CellPicture = imglst.ListImages(IMG_PASCOCHE).Picture
            End If
            rs.MoveNext
            I = I + 1
        Wend

        ' Redimensionner le grid est déplacer les boutons +/- si besoin
        If .Rows > NBRMAX_ROWS And .width = LARGEUR_GRID_PAR_DEFAUT Then
            '.width = .width + 255
            'cmd(CMD_PLUS_TYPE).left = LEFT_CMD_PAR_DEFAUT + 255
            'cmd(CMD_MOINS_TYPE).left = LEFT_CMD_PAR_DEFAUT + 255
            .ColWidth(GRDC_COMMENTAIRE) = 2280
        End If
        .Enabled = IIf(.Rows - 1 = 0, False, True)
    End With

    afficher_coordonnee = P_OK
    rs.Close

End Function

Private Function afficher_coordonneeLiees()

    Dim sql As String
    Dim rs As rdoResultset

    sql = "SELECT * FROM Coordonnee_Associee, UtilCoordonnee, ZoneUtil" _
        & " WHERE UC_ZUNum=ZU_Num" _
        & " AND UC_Num=CA_UCNum" _
        & " AND CA_UCTypeNum=" & g_numutil _
        & " AND CA_UCType ='U' AND ZU_Type='C'" _
        & " ORDER BY ZU_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then GoTo lab_erreur

    ' Affichage des lignes du gridCoord
    With grdCoordLiees
        .Rows = 0
        While Not rs.EOF
            If set_coordLiees(rs("UC_Num").Value, rs("CA_Principal").Value) = P_ERREUR Then GoTo lab_erreur
'            .AddItem rs("ZU_Num").Value & vbTab & rs("ZU_Libelle") & vbTab & rs("UC_Valeur") _
'                    & vbTab & rs("UC_Niveau") & vbTab & vbTab & rs("UC_Comm")
'            '.Row = i
'            .Row = .Rows - 1
'            .col = GRDCL_PRINCIPAL
'            If rs("CA_Principal").Value Then
'                Set .CellPicture = ImageListGenerale.ListImages(IMG_COCHE).Picture
'            Else
'                Set .CellPicture = LoadPicture("")
'            End If
            rs.MoveNext
        Wend

        ' Redimensionner le grid est les boutons +/- si besoin est
        If .Rows > NBRMAX_ROWS And .width = LARGEUR_GRID_PAR_DEFAUT Then
            '.width = .width + 255
            'cmd(CMD_PLUS_TYPE).left = LEFT_CMD_PAR_DEFAUT + 255
            'cmd(CMD_MOINS_TYPE).left = LEFT_CMD_PAR_DEFAUT + 255
            .ColWidth(GRDCL_IDENTITE) = 2310
        End If
        .Enabled = IIf(.Rows = 0, False, True)
    End With
    rs.Close

    afficher_coordonneeLiees = P_OK
    Exit Function

lab_erreur:
    afficher_coordonneeLiees = P_ERREUR

End Function

Private Sub afficher_historique()
' Afficher l'historique des mouvements pour cet utilisateur, depuis UtilMouvement

    Call HistoriqueMouvement.AppelFrm(g_numutil, "", False)

End Sub

Private Sub afficher_laboratoires(ByVal v_ssite As String, _
                                  ByVal v_lnumprinc As Long)

    Dim n As Integer, I As Integer, j As Integer
    Dim numl As Long

    n = STR_GetNbchamp(v_ssite, ";")
    For I = 0 To n - 1
        numl = Mid$(STR_GetChamp(v_ssite, ";", I), 2)
        For j = 0 To grdLabo.Rows - 1
            If grdLabo.TextMatrix(j, GRDL_NUMLABO) = numl Then
                grdLabo.TextMatrix(j, GRDL_ESTLABO) = True
                grdLabo.Row = j
                grdLabo.col = GRDL_IMG_ESTLABO
                Set grdLabo.CellPicture = imglst.ListImages(IMG_COCHE).Picture
                Exit For
            End If
        Next j
    Next I
    For I = 0 To grdLabo.Rows - 1
        If grdLabo.TextMatrix(I, GRDL_NUMLABO) = v_lnumprinc Then
            grdLabo.TextMatrix(I, GRDL_LABOPRINC) = "Princiapl"
            Exit For
        End If
    Next I

    If grdLabo.Rows > 0 Then
        grdLabo.Row = 0
        grdLabo.RowSel = 0
    End If

End Sub

Private Sub afficher_menu(ByVal v_mode As Boolean)
' ********************************************************
' Afficher le menu contextuel de mettre le poste principal
' v_mode: le poste est déjà principal ou non
' ********************************************************

    mnuPostePrinc.Enabled = v_mode
    mnuPostePrinc.Visible = True
    Call PopupMenu(mnuPobPrinc)

End Sub

Private Sub afficher_page(ByVal v_sens As Integer)

    If v_sens = 0 Then
        If sst.Tab > 0 Then
            sst.Tab = sst.Tab - 1
        Else
            sst.Tab = sst.Tabs - 1
        End If
    Else
        If sst.Tab < sst.Tabs - 1 Then
            sst.Tab = sst.Tab + 1
        Else
            sst.Tab = ONGLET_GENERAL
        End If
    End If

    Call init_focus

End Sub

Private Sub afficher_photo()

    Dim s As String, nomimg_loc As String, nomimg_srv As String, ext As String
    Dim I As Integer, n As Integer
    
    If g_nom_fich_photo <> "" Then
        nomimg_loc = g_nom_fich_photo
    Else
        s = p_CheminKW & "/kalibottin/photos/" & g_numutil
        n = STR_GetNbchamp(LISTE_FICHIERS_IMAGES, ";")
        For I = 0 To n - 1
            ext = Mid$(STR_GetChamp(LISTE_FICHIERS_IMAGES, ";", I), 2)
            If KF_FichierExiste(s & ext) Then
                nomimg_srv = s & ext
                Exit For
            End If
        Next I
        nomimg_loc = p_chemin_appli + "/tmp/" & g_numutil & ext
        If KF_GetFichier(nomimg_srv, nomimg_loc) = P_ERREUR Then
            Exit Sub
        End If
    End If
'    pctPhoto.Picture = LoadPicture(nomimg_loc)
'    pctPhoto.Visible = True
    
End Sub

Private Function afficher_services(ByVal v_spm As Variant, Optional ByVal v_numPobPrinc As Long) As Integer

    Dim s As String, s1 As String, lib As String
    Dim n As Integer, I As Integer, j As Integer, n2 As Integer
    Dim num As Long
    Dim sql As String, rs As rdoResultset
    Dim nd As Node

    n = STR_GetNbchamp(v_spm, "|")
    For I = 1 To n
        s = STR_GetChamp(v_spm, "|", I - 1)
        n2 = STR_GetNbchamp(s, ";")
        For j = 1 To n2
            s1 = STR_GetChamp(s, ";", j - 1)
            If TV_NodeExiste(tvSect, s1, nd) = P_OUI Then
                GoTo lab_sp_suiv
            End If
            num = CLng(Mid$(s1, 2))
            If left(s1, 1) = "S" Then ' Un service
                If P_RecupSrvNom(num, lib) = P_ERREUR Then
                    afficher_services = P_ERREUR
                    Exit Function
                End If
                If j = 1 Then
                    Set nd = tvSect.Nodes.Add(, tvwChild, "S" & num, lib, IMGT_SERVICE, IMGT_SERVICE)
                Else
                    Set nd = tvSect.Nodes.Add(nd, tvwChild, "S" & num, lib, IMGT_SERVICE, IMGT_SERVICE)
                End If
                nd.Expanded = True
            Else ' Un poste
                If P_RecupPosteNom(num, lib) = P_ERREUR Then
                    afficher_services = P_ERREUR
                    Exit Function
                End If
                ' mis par synchro ?
                If Odbc_TableExiste("kb_poste_secondaire") Then
                    sql = "select * from kb_poste_secondaire where psu_unum=" & g_numutil
                    If Odbc_SelectV(sql, rs) = P_OK Then
                        While Not rs.EOF
                            ' MsgBox rsHisto("psu_poste")
                            If InStr(rs("psu_poste"), "P" & num & ";") > 0 Then
                                lib = lib & " (-> Synchro)"
                            End If
                            rs.MoveNext
                        Wend
                    End If
                End If
                If num = v_numPobPrinc Then
                    ' on ne passe ici qu'une seule fois = un seul poste principal
                    Call tvSect.Nodes.Add(nd, tvwChild, "P" & num, lib, IMGT_POSTE_SEL, IMGT_POSTE_SEL)
                Else
                    Call tvSect.Nodes.Add(nd, tvwChild, "P" & num, lib, IMGT_POSTE, IMGT_POSTE)
                End If
            End If
lab_sp_suiv:
        Next j
    Next I

    If tvSect.Nodes.Count = 0 Then
        tvSect.BorderStyle = ccNone
        cmd(CMD_MOINS_SPM).Visible = False
    End If

    afficher_services = P_OK

End Function

Private Function afficher_utilisateur() As Integer

    Dim sql As String, s As String
    Dim pos As Integer, nbr_mouvements As Integer
    Dim rs As rdoResultset

    g_mode_saisie = False

    Call FRM_ResizeForm(Me, 0, 0)

    grdLabo.Rows = 0
    Call init_grD_Site

    tvSect.Nodes.Clear

    If g_numutil > 0 Then ' Utilisateur existant
        sql = "SELECT * FROM Utilisateur WHERE U_kb_actif=True AND U_Num=" & g_numutil
        If Odbc_Select(sql, rs) = P_ERREUR Then
            GoTo lab_err
        End If
        ' Infos pour UtilMouvement
        g_old_nom = rs("U_Nom").Value
        g_old_prenom = rs("U_Prenom").Value
        g_old_matricule = rs("U_Matricule").Value
        g_old_active = rs("U_Actif").Value
        g_old_spm = rs("U_SPM").Value
        ' Onglet Général
        txt(TXT_NOM).Text = rs("U_Nom").Value
        txt(TXT_PRENOM).Text = rs("U_Prenom").Value
        chk(CHK_ACTIF).Value = IIf(rs("U_Actif").Value, 1, 0)
        chk(CHK_EXTERNE).Value = IIf(rs("U_Externe").Value, 1, 0)
        If chk(CHK_EXTERNE).Value = 1 Then
            chk(CHK_EXTERNE_IMPORT).Value = 1
            chk(CHK_EXTERNE_IMPORT).Enabled = False
        Else
            chk(CHK_EXTERNE_IMPORT).Value = IIf(rs("U_ExterneFich").Value, 1, 0)
        End If
        txt(TXT_MATRICULE).Text = rs("U_Matricule").Value
        ' Nom + prénom Junon
        txt(TXT_NOM_JUNON).Text = rs("U_NomJunon").Value & ""
        txt(TXT_PRENON_JUNON).Text = rs("U_PrenomJunon").Value & ""
        ' Onglet Poste
        Call afficher_laboratoires(rs("U_Labo").Value, rs("U_LNumPrinc").Value)
        g_pobPrinc = rs("U_Po_Princ").Value
        If afficher_services(rs("U_SPM").Value, rs("U_Po_Princ").Value) = P_ERREUR Then
            GoTo lab_err
        End If
        ' Ne pas pouvoir supprimer si la personne à un poste KaliDoc
        'cmd(CMD_SUPPRIMER).Enabled = rs("U_SPM").Value = ""
        cmd(CMD_SUPPRIMER).Visible = True
        rs.Close
        ' Code et Mot de passe
        sql = "SELECT UAPP_Code, UAPP_MotPasse, UAPP_TypeCrypt FROM UtilAppli" _
            & " WHERE UAPP_APPNum=" & p_appli_kalidoc _
            & " AND UAPP_UNum=" & g_numutil
        If Odbc_Select(sql, rs) = P_ERREUR Then
            GoTo lab_err
        End If
        g_old_code = rs("UAPP_Code").Value
        txt(TXT_CODE).Text = rs("UAPP_Code").Value
        If rs("UAPP_TypeCrypt").Value = "kalidoc" Or rs("UAPP_TypeCrypt").Value = "" Then
            txt(TXT_MPASSE).Text = STR_Decrypter(rs("UAPP_MotPasse").Value)
        ElseIf rs("UAPP_TypeCrypt").Value = "kalidocnew" Then
            txt(TXT_MPASSE).Text = STR_Decrypter_New(rs("UAPP_MotPasse").Value)
        ElseIf rs("UAPP_TypeCrypt").Value = "md5" Then
            txt(TXT_MPASSE).Text = rs("UAPP_MotPasse").Value
        End If
        rs.Close
        If afficher_coordonnee() = P_ERREUR Then
            GoTo lab_err
        End If
        If afficher_coordonneeLiees() = P_ERREUR Then
            GoTo lab_err
        End If
        cmd(CMD_OK).Enabled = False
        ' récupérer les anciennes coordonnées
        g_coordonnees_avant = get_str_coordonnees()
        ' La photo
        Call getPhotoUtilisateur
        Call evalue_btninfosuppl
    Else ' Nouvel utilisateur
        ' Remplir la form
        txt(TXT_NOM).Text = ""
        txt(TXT_PRENOM).Text = ""
        txt(TXT_CODE).Text = ""
        txt(TXT_MPASSE).Text = ""
        chk(CHK_ACTIF).Value = 1
        chk(CHK_EXTERNE).Value = 0
        chk(CHK_EXTERNE_IMPORT).Value = 0
        txt(TXT_MATRICULE).Text = ""
        ' Nom + prénom Junon
        txt(TXT_NOM_JUNON).Text = ""
        txt(TXT_PRENON_JUNON).Text = ""
        tvSect.BorderStyle = ccNone
        cmd(CMD_MOINS_SPM).Visible = False
        cmd(CMD_OK).Enabled = False
        cmd(CMD_SUPPRIMER).Visible = False
        pos = InStr(g_sprm, "NOM=")
        If pos > 0 Then
            txt(TXT_NOM).Text = Mid$(g_sprm, pos + 4)
        End If
        pos = InStr(g_sprm, "POSTE=")
        If pos > 0 Then
            If build_services(Mid$(g_sprm, pos + 6), s) = P_ERREUR Then
                GoTo lab_err
            End If
            If afficher_services(s) = P_ERREUR Then
                GoTo lab_err
            End If
        End If
        grdCoordLiees.Rows = 0
        ' La photo
        cmd(CMD_PHOTO).Visible = False
    End If

    g_mode_saisie = True

    sst.Tab = ONGLET_GENERAL

    If grdCoord.Rows > 1 Then
        grdCoord.Row = 1
        grdCoord.col = GRDC_VALEUR
    End If

    ' Désactiver le bouton CMD_HISTORIQUE s'il n'y a pas de mouvements à afficher
    If Odbc_RecupVal("SELECT COUNT(*) FROM UtilMouvement WHERE UM_UNum=" & g_numutil, nbr_mouvements) = P_ERREUR Then
        GoTo lab_err
    End If
    If nbr_mouvements = 0 Then
        cmd(CMD_HISTORIQUE).Enabled = False
    End If
    

    Call FRM_ResizeForm(Me, Me.width, Me.Height)

    txt(TXT_NOM).SetFocus
    
    afficher_utilisateur = P_OK
    Exit Function

lab_err:
    afficher_utilisateur = P_ERREUR

End Function

Private Sub ajouter_coordLiee()
    Dim selection As String, sret As String, entite As String, sql As String, _
        titre_entite As String, nom As String, prenom As String, choix As String
    Dim I As Long, lng As Long, nbr_ligne As Long
    Dim uc_selected As Boolean
    Dim rs As rdoResultset
    Dim frm As Form

    With grdCoordLiees
        Call CL_Init
        Call CL_InitMultiSelect(False, False) ' (selection multiple=True, retourner la ligne courante=False)
        Call CL_InitTitreHelp("Liste des entités possibles", p_chemin_appli + "\help\kalidoc.chm" & ";" & "dico_d_fonction.htm")
        Call CL_InitTaille(0, -15)
    
        Call CL_AddLigne("Personne", 0, "U", False)
        Call CL_AddLigne("Poste ou Pièce", 1, "P", False)
    
        Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
        Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    
        ChoixListe.Show 1
    
        If CL_liste.retour = 1 Then ' --------------- QUITTER
            Exit Sub
        End If
    
        ' selection
        If CL_liste.retour = 0 Then ' --------------- SELECTIONNER
            If CL_liste.pointeur = 0 Then ' personne
                ' selectionner depuis la fonction recherche
lab_reselection_personne:
                Set frm = ChoixUtilisateur
                choix = ChoixUtilisateur.AppelFrm("Choisir une personne", _
                                                    "", _
                                                    False, _
                                                    False, _
                                                    "", _
                                                    g_numutil)
                Set frm = Nothing
                If choix = "" Then ' pas de personne selectionnée
                    Exit Sub
                End If
                lng = p_tblu_sel(0)
                entite = "U"
                ' pour les utilisateur, on passe par
            Else ' CL_liste.pointeur = 1 => poste ou pièce
lab_reselection_poste_piece:
                Set frm = PrmService
                sret = PrmService.AppelFrm("Choix du poste/pièce", "S", False, "1;", "PC", False)
                Set frm = Nothing
                If sret = "" Then Exit Sub
                entite = STR_GetChamp(sret, ";", STR_GetNbchamp(sret, ";") - 1)
                lng = Mid$(entite, 2)
                entite = Mid$(entite, 1, 1)
            End If
            sql = "SELECT * FROM UtilCoordonnee WHERE UC_Type='" & entite & "' AND UC_TypeNum=" & lng
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                Exit Sub
            End If
            If rs.EOF Then ' pas de coordonnées pour cette entité
                If MsgBox("Cette entitié ne dispose pas de coordonnées." & vbCrLf & _
                           "Voulez-vous réessayer avec une autre ?", vbYesNo + vbQuestion, "Pas de coordonnées") = vbYes Then
                    If entite = "U" Then
                        GoTo lab_reselection_personne
                    Else
                        GoTo lab_reselection_poste_piece
                    End If
                Else ' on sort
                    Exit Sub
                End If
            End If
            rs.Close
            ' choix multiple : CL_InitMultiSelect
            Select Case entite
                Case "U":
                    If Odbc_RecupVal("SELECT U_Nom, U_Prenom FROM Utilisateur WHERE U_kb_actif=True AND U_Num=" & lng, _
                                    nom, prenom) = P_ERREUR Then
                        Exit Sub
                    End If
                    titre_entite = nom & " " & prenom
                Case "P":
                    titre_entite = "le poste: " & P_get_lib_srv_poste(lng, P_POSTE)
                Case "C":
                    titre_entite = "la pièce: " & P_get_nom_piece(lng)
            End Select
            ' Afficher la liste multiselect des coordonnées de l'entité
            Call CL_Init
            Call CL_InitMultiSelect(True, False) 'selection multiple=True, retourner la ligne courante=False
            Call CL_InitTitreHelp("Liste des coordonnées pour " & titre_entite, p_chemin_appli + "\help\kalidoc.chm" & ";" & "dico_d_fonction.htm")
            Call CL_InitTaille(0, -15)
            ' Boucle d'ajout dans la liste des coordonnées de l'entité selectionnée
            sql = "SELECT * FROM ZoneUtil, UtilCoordonnee" & _
                  " WHERE UC_ZUNum=ZU_Num AND UC_Type='" & entite & "' AND UC_TypeNum=" & lng
            If Odbc_Select(sql, rs) = P_ERREUR Then
                Exit Sub
            End If
            nbr_ligne = rs.RowCount
            While Not rs.EOF ' on est sur d'avoir au moin une coordonnée !
                uc_selected = coordLiee_existe(rs("UC_Num").Value)
                Call CL_AddLigne(rs("ZU_Libelle").Value & " :" & vbTab _
                                & rs("UC_Valeur").Value, rs("UC_Num").Value, _
                                rs("ZU_Code").Value, uc_selected)
                rs.MoveNext
            Wend
            rs.Close
            Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
            Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
            ChoixListe.Show 1
            ' La réponse:
            If CL_liste.retour = 1 Then ' QUITTER
                Exit Sub
            End If
            If CL_liste.retour = 0 Then ' ENREGISTRER
                For I = 0 To nbr_ligne - 1 ' on commence par ajouter
                    If CL_liste.lignes(I).selected Then
                        If Not coordLiee_existe(CL_liste.lignes(I).num) Then
                            Call set_coordLiees(CL_liste.lignes(I).num, False)
                        End If
                    End If
                Next I
                For I = 0 To nbr_ligne - 1 ' ..puis par supprimer
                    If Not CL_liste.lignes(I).selected Then ' essayer de la supprimer du tableau
                        Call enlever_coordLiees(CL_liste.lignes(I).num)
                    End If
                Next I
                cmd(CMD_OK).Enabled = True
            End If
        End If
        .Enabled = True
    End With

End Sub

Private Sub ajouter_type_coord()

    Dim sql As String, sret As String
    Dim insere_milieu As Boolean
    Dim I As Integer, col_num As Integer, nbrmax As Integer, j As Integer
    Dim num As Long
    Dim rs As rdoResultset

    I = 0
    j = 0
    num = 0
    nbrmax = 0
    insere_milieu = False

Lab_Debut: ' réinitialiser l'affichage
    Call CL_Init
    Call CL_InitMultiSelect(False, False) ' (selection multiple=True, retourner la ligne courante=False)
    Call CL_InitTitreHelp("Liste des types de coordonnée", p_chemin_appli + "\help\kalidoc.chm" & ";" & "dico_d_fonction.htm")
    Call CL_InitTaille(0, -15)

    ' boucle SQL d'ajout dans la liste des choix
    sql = "SELECT * FROM ZoneUtil WHERE ZU_Type='C'"
    If Odbc_Select(sql, rs) = P_ERREUR Then
    End If
    While Not rs.EOF
        ' Compter si NOMBRE MAX est atteint
        nbrmax = 0
        With grdCoord
            For j = 1 To .Rows - 1
                If .TextMatrix(j, GRDC_ZUNUM) = rs("ZU_Num") Then
                    nbrmax = nbrmax + 1
                End If
            Next j
        End With
        ' AJOUT DANS LA LISTE DES CHOIX
        If nbrmax < rs("ZU_NBREMax") Then
            Call CL_AddLigne(rs("ZU_Code").Value, rs("ZU_Num").Value, rs("ZU_Libelle").Value, False)
        End If
        rs.MoveNext
    Wend
    num = 0
    rs.Close

    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("&Créer un type de coordonnées", "", vbKeyU, vbKeyF3, 1500)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)

    ChoixListe.Show 1

    If CL_liste.retour = 2 Then ' --------------- QUITTER
        Exit Sub
    End If

    If CL_liste.retour = 1 Then ' --------------- AJOUTER
        sret = PrmTypeCoordonnees.AppelFrm(0)
        If sret <> "" Then
            num = STR_GetChamp(sret, "|", 0)
        End If
        GoTo Lab_Debut
    End If

    ' Enregistrer
    If CL_liste.retour = 0 Then ' --------------- ENREGISTRER
        With grdCoord
            For j = 1 To .Rows - 1
                If UCase$(CL_liste.lignes(CL_liste.pointeur).tag) < UCase$(.TextMatrix(j, GRDC_TYPE)) Then
                    .AddItem (CL_liste.lignes(I).num & vbTab), j
                    .Row = j
                    .col = GRDC_PRINCIPAL
                    .CellPictureAlignment = 4
                    Set .CellPicture = imglst.ListImages(IMG_PASCOCHE).Picture
                    insere_milieu = True
                    Exit For
                End If
            Next j
            If Not insere_milieu Then
                .AddItem (CL_liste.lignes(I).num & vbTab)
                .Row = .Rows - 1
                .col = GRDC_PRINCIPAL
                .CellPictureAlignment = 4
                Set .CellPicture = imglst.ListImages(IMG_PASCOCHE).Picture
            End If
            For I = 1 To .Rows - 2
                If .TextMatrix(I, GRDC_ZUNUM) = CL_liste.lignes(CL_liste.pointeur).num Then
                    .Row = I
                    .col = GRDC_PRINCIPAL
                    GoTo coord_existant:
                End If
            Next I
            .Row = j
            .col = GRDC_PRINCIPAL
            Set .CellPicture = imglst.ListImages(IMG_COCHE).Picture
            If Not insere_milieu Then ' on insert à la fin du grid
                .Row = .Rows - 1
            Else ' on insert dans la ligne J
                .Row = j
            End If
coord_existant:     ' et est déjà coché
            .TextMatrix(j, GRDC_ZUNUM) = CL_liste.lignes(CL_liste.pointeur).num
            .TextMatrix(j, GRDC_TYPE) = CL_liste.lignes(CL_liste.pointeur).tag
            cmd(CMD_OK).Enabled = True
            If .Rows > NBRMAX_ROWS Then ' gérer la bare de défilement vertical
                If .width = LARGEUR_GRID_PAR_DEFAUT Then
                    '.width = .width + 255
                    'cmd(CMD_PLUS_TYPE).left = LEFT_CMD_PAR_DEFAUT + 255
                    'cmd(CMD_MOINS_TYPE).left = LEFT_CMD_PAR_DEFAUT + 255
                    .ColWidth(GRDC_COMMENTAIRE) = 2310
                Else
                    .ScrollBars = flexScrollBarVertical
                End If
            End If

            .Row = j
            .col = GRDC_VALEUR
            txt(TXT_CACHE).Text = ""
            g_position_txt_cache = ""
            .Enabled = True
        End With
    End If

End Sub

Public Function AppelFrm(ByVal v_numutil As Long, _
                    ByVal v_sprm As String) As Boolean
' *******************************************************************
' Peut retourner une code s'il y a eu des changements importants
' dans les infocrmations d'une personne:
' * Nom prenom, matricule
' * Passage de actif <-> inactif, externe(fich) <-> non externe(fich)
' * Suppression d'un poste dans la liste U_SPM
' *******************************************************************

    g_numutil = v_numutil
    g_sprm = v_sprm

    Me.Show 1

    AppelFrm = g_changements_importants

End Function

Private Function supprimer_utilisateur() As Integer
    Dim ssite As String, sfct As String
    Dim num_util As Long, numlabop As Long, lng As Long
    Dim spm As Variant
    Dim frm As Form
    
    If MsgBox("Confimez vous la suppression de " & txt(TXT_NOM).Text & " " & txt(TXT_PRENOM).Text & " de la gestion de KaliBottin ? " & vbCrLf & _
              "Attention !!! Cet utilisateur n'apparaitra plus dans KaliBottin ", vbYesNo + vbQuestion, "Confirmer suppression") = vbYes Then
        If g_numutil <> 0 Then
            num_util = g_numutil
            Call changements_importants(spm, g_old_spm)
            ' Désactiver utilisateur de Kalibottin
            If Odbc_Update("Utilisateur", _
                           "U_Num", _
                           "WHERE U_Num=" & g_numutil, _
                           "U_Nom", UCase$(txt(TXT_NOM).Text), _
                           "U_Prenom", txt(TXT_PRENOM).Text, _
                           "U_NomJunon", UCase$(txt(TXT_NOM_JUNON).Text), _
                           "U_PrenomJunon", txt(TXT_PRENON_JUNON).Text, _
                           "U_Actif", IIf(chk(CHK_ACTIF).Value = 1, True, False), _
                           "U_kb_actif", False, _
                           "U_Externe", IIf(chk(CHK_EXTERNE).Value = 1, True, False), _
                           "U_ExterneFich", IIf(chk(CHK_EXTERNE_IMPORT).Value = 1, True, False), _
                           "U_Matricule", txt(TXT_MATRICULE).Text, _
                           "U_SPM", spm, _
                           "U_FctTrav", sfct, _
                           "U_Labo", ssite, _
                           "U_LNumPrinc", numlabop, _
                           "U_Po_Princ", g_pobPrinc) = P_ERREUR Then
                GoTo err_supp
            End If
        Else
            supprimer_utilisateur = P_NON
            Exit Function
        End If
    Else ' on sort
        Exit Function
    End If
    
lab_commit:
    
    supprimer_utilisateur = P_OUI
    Exit Function
    
err_supp:
    supprimer_utilisateur = P_ERREUR
    
End Function




Private Sub basculer_colonne_principal(ByVal v_row As Integer, ByVal v_col As Integer)

    Dim I As Integer, type_en_cours As Integer, nbr As Integer, ma_row As Integer

    With grdCoord
        ' ne pas tenter de basculer si on est sur la ligne fixe
        If v_row = 0 Then Exit Sub

        type_en_cours = .TextMatrix(v_row, GRDC_ZUNUM)
        If .CellPicture = imglst.ListImages(IMG_COCHE).Picture Then
        ' la cellule elle est cochée
            .Row = v_row
            .col = v_col
            Set .CellPicture = imglst.ListImages(IMG_PASCOCHE).Picture ' on décoche cette ligne
            nbr = 0
            For I = 1 To .Rows - 1 ' on chrche le nombre de lignes du même type
                If .TextMatrix(I, GRDC_ZUNUM) = type_en_cours And I <> v_row Then
                    ma_row = I
                    nbr = nbr + 1
                End If
            Next I
            If nbr = 1 Then ' s'il y en a deux, on coche la deuxième
                .Row = ma_row
                .col = GRDC_PRINCIPAL
                Set .CellPicture = imglst.ListImages(IMG_COCHE).Picture
            End If
            cmd(CMD_OK).Enabled = True
        Else ' elle n'est pas cochée
            ' vérifier s'il n'y pas d'autre principale pour le même type de coordonnée
            For I = 1 To .Rows - 1
                If .TextMatrix(I, GRDC_ZUNUM) = type_en_cours And I <> v_row Then
                    .col = GRDC_PRINCIPAL
                    .Row = I
                    If .CellPicture = imglst.ListImages(IMG_COCHE).Picture Then
                        Set .CellPicture = imglst.ListImages(IMG_PASCOCHE).Picture
                    End If
                End If
            Next I
            .Row = v_row
            .col = v_col
            Set .CellPicture = imglst.ListImages(IMG_COCHE).Picture
            cmd(CMD_OK).Enabled = True
        End If
    End With

End Sub

Private Sub basculer_postePrincipal(ByVal v_index As Integer)
' *************************************************
' Mettre le poste selectionné comme poste principal
' *************************************************
    Dim num_poste As Long
    Dim I As Integer
    Dim nd As Node

    For I = 1 To tvSect.Nodes.Count
        Set nd = tvSect.Nodes(I)
        If left$(nd.key, 1) = "P" Then
            num_poste = Mid$(nd.key, 2)
            If num_poste = g_pobPrinc Then
                nd.image = IMGT_POSTE
                nd.SelectedImage = IMGT_POSTE
                GoTo lab_ok
            End If
        End If
    Next I
lab_ok:
    tvSect.Nodes(v_index).image = IMGT_POSTE_SEL
    tvSect.Nodes(v_index).SelectedImage = IMGT_POSTE_SEL
    g_pobPrinc = Mid$(tvSect.Nodes(v_index).key, 2)
    cmd(CMD_OK).Enabled = True

End Sub
Private Sub attribuer_coordonnee()
    
    Dim num_poste As Long, num_poste_selected As Long
    Dim num_service As Long
    Dim I As Integer
    Dim nd As Node
    
    num_poste_selected = Mid$(tvSect.Nodes(tvSect.SelectedItem.Index).key, 2)
    
    For I = 1 To tvSect.Nodes.Count
        Set nd = tvSect.Nodes(I)
        'prendre le dernier service avant les postes
        If left$(nd.key, 1) = "S" Then
            num_service = Mid$(nd.key, 2)
        End If
        'le poste correspond t-il à la sélection ?
        If left$(nd.key, 1) = "P" Then
            num_poste = Mid$(nd.key, 2)
            If num_poste = num_poste_selected Then
                GoTo lab_ok
            End If
        End If
    Next I
lab_ok:
    Call afficher_liste_coordonnee_util(num_service, num_poste_selected)
    
End Sub
Private Sub afficher_liste_coordonnee_util(ByVal v_numsrv As Long, ByVal v_numposte As Long)
    
    Dim n As Integer, nbr_ligne As Integer, I As Integer, j As Integer
    Dim rs As rdoResultset
    Dim sql As String, nom As String, prenom As String, titre_entite As String
    Dim couple As String, lst_newposte As String, test As String
    Dim checked As Boolean
    
    titre_entite = ""
    sql = "SELECT U_Nom, U_Prenom FROM Utilisateur WHERE " _
        & " U_kb_actif=True AND U_Num=" & g_numutil
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    If Not rs.EOF Then
        titre_entite = rs("U_Nom") & " " & rs("U_Prenom")
    End If
    couple = "S" & v_numsrv & ";P" & v_numposte & ";|"
    sql = "SELECT ZU_Code, UC_Lstposte, UC_Valeur, UC_Num FROM UtilCoordonnee, ZoneUtil " _
        & " WHERE UC_TypeNum=" & g_numutil _
        & " AND UC_Type='U' AND ZU_Type='C'" _
        & " AND UC_ZUNum=ZU_Num" _
        & " ORDER BY ZU_Code"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    If rs.EOF Then
        MsgBox " Aucune coordonnee trouvée pour cet utilisateur "
        rs.Close
        Exit Sub
    End If

    Call CL_Init
    Call CL_InitMultiSelect(True, False)
    Call CL_InitTitreHelp("Liste des coordonnées pour " & titre_entite, p_chemin_appli + "\help\kalidoc.chm" & ";" & "dico_d_fonction.htm")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    nbr_ligne = rs.RowCount
    While Not rs.EOF
        'Cocher les cases si le couple S[numservice];P[numposte];| existe
        If InStr(rs("UC_Lstposte").Value, couple) > 0 And rs("UC_Lstposte").Value <> "" Then
           checked = True
        Else
           checked = False
        End If
        Call CL_AddLigne(rs("UC_Valeur").Value, rs("UC_Num").Value, rs("UC_Lstposte").Value, checked)
        rs.MoveNext
    Wend
    Call CL_InitTaille(0, -20)
    ChoixListe.Show 1
    If CL_liste.retour = 0 Then
        For I = 0 To nbr_ligne - 1
            lst_newposte = ""
            If CL_liste.lignes(I).selected Then
                If IsNull(CL_liste.lignes(I).tag) Then
                    Call update_liste_poste(CL_liste.lignes(I).num, couple)
                Else
                    'La chaine n'existe pas encore ajouter au couple
                    If InStr(CL_liste.lignes(I).tag, couple) = 0 Then
                        lst_newposte = CL_liste.lignes(I).tag + couple
                        Call update_liste_poste(CL_liste.lignes(I).num, lst_newposte)
                    End If
                End If
            Else
                If Not IsNull(CL_liste.lignes(I).tag) Then
                    test = CL_liste.lignes(I).tag
                    'Si pas cocher | et existe on le supprime
                    If InStr(test, couple) > 0 Then
                        'supprimer couple de la chaine
                        n = STR_GetNbchamp(test, "|")
                        For j = 0 To n - 1
                            If STR_GetChamp(test, "|", j) <> left(couple, Len(couple) - 1) Then
                                lst_newposte = lst_newposte + STR_GetChamp(test, "|", j) & "|"
                            End If
                        Next j
                        Call update_liste_poste(CL_liste.lignes(I).num, lst_newposte)
                    End If
                End If
            End If
        Next I
    ElseIf CL_liste.retour = 1 Then
        Exit Sub
    End If
    
    
End Sub
Private Sub update_liste_poste(ByVal v_ucnum As Long, ByVal v_couple As String)
    
    Dim rs As rdoResultset
    Dim sql As String
    
    If Odbc_Update("UtilCoordonnee", "UC_Num", "WHERE UC_Num=" & v_ucnum, _
                                "UC_Lstposte", v_couple) = P_ERREUR Then
        MsgBox " Erreur inscription coordonnee "
    End If
    
End Sub

Private Function build_services(ByVal v_numposte As Long, _
                               ByRef r_sp As String) As Integer

    Dim s As String, sql As String
    Dim numsrv As Long

    s = "P" & v_numposte & ";"
    sql = "SELECT PO_SRVNum FROM Poste WHERE PO_Num=" & v_numposte
    If Odbc_RecupVal(sql, numsrv) = P_ERREUR Then
        build_services = P_ERREUR
        Exit Function
    End If
    s = "S" & numsrv & ";" & s
    Do
        sql = "SELECT SRV_NumPere FROM Service WHERE SRV_Num=" & numsrv
        If Odbc_RecupVal(sql, numsrv) = P_ERREUR Then
            build_services = P_ERREUR
            Exit Function
        End If
        If numsrv > 0 Then
            s = "S" & numsrv & ";" & s
        End If
    Loop Until numsrv = 0
    
    r_sp = s
    
    build_services = P_OK
    
End Function

Private Sub build_SPM_Fct(ByRef r_spm As Variant, _
                          ByRef r_sfct As String)

    Dim s As String, sp As String, sql As String
    Dim encore As Boolean
    Dim I As Integer, j As Integer, n As Integer
    Dim numfct As Long, num As Long
    Dim nd As Node, ndP As Node

    r_spm = ""
    r_sfct = ""

    For I = 1 To tvSect.Nodes.Count
        Set nd = tvSect.Nodes(I)
        If left$(nd.key, 1) = "P" Then
            num = Mid$(nd.key, 2)
            sql = "SELECT PO_FTNum FROM Poste" _
                & " WHERE PO_Num=" & num
            Call Odbc_RecupVal(sql, numfct)
            If InStr(r_sfct, "F" & numfct & ";") = 0 Then
                r_sfct = r_sfct & "F" & numfct & ";"
            End If
            sp = nd.key & ";"
            Do
                Set ndP = nd.Parent
                encore = True
                On Error GoTo lab_no_prev
                s = ndP.key
                On Error GoTo 0
                If encore Then
                    sp = sp & ndP.key & ";"
                    Set nd = ndP
                End If
            Loop Until Not encore
            n = STR_GetNbchamp(sp, ";")
            For j = n - 1 To 0 Step -1
                r_spm = r_spm + STR_GetChamp(sp, ";", j) & ";"
            Next j
            r_spm = r_spm + "|"
        End If
    Next I

    r_spm = IIf(r_spm = "", r_spm, IIf(Right$(r_spm, 1) = "|", r_spm, r_spm + "|"))
    Exit Sub

lab_no_prev:
    encore = False
    Resume Next

End Sub

Private Sub changements_importants(ByVal v_new_spm As String, ByVal v_old_spm As String)
' *********************************************************************
' Verifier qu'on a effectué des changements importants sur la personnes
' *********************************************************************
    Dim spm_en_cours As String
    Dim num_poste_en_cours As Long
    Dim old_externe As Boolean, old_exeterneFich As Boolean, poste_trouve As Boolean
    Dim I As Integer, j As Integer, nbr_avant As Integer, nbr_apres As Integer

    If Odbc_RecupVal("SELECT U_Externe, U_ExterneFich FROM Utilisateur WHERE U_kb_actif=True AND U_Num=" & g_numutil, _
                     old_externe, old_exeterneFich) = P_ERREUR Then
        Exit Sub
    End If
    ' rendre la personne externe ?
    If old_externe <> chk(CHK_EXTERNE).Value Then GoTo lab_changements
    ' rendre la personne externe au fichier d'importation ?
    If old_exeterneFich <> chk(CHK_EXTERNE_IMPORT).Value Then GoTo lab_changements
    ' a-t-on supprimé des postes pour cette personne ?
    nbr_avant = STR_GetNbchamp(v_old_spm, "|")
    nbr_apres = STR_GetNbchamp(v_new_spm, "|")
    For I = 0 To nbr_avant - 1
        spm_en_cours = STR_GetChamp(v_old_spm, "|", I)
        num_poste_en_cours = Mid$(STR_GetChamp(spm_en_cours, ";", STR_GetNbchamp(spm_en_cours, ";") - 1), 2)
        poste_trouve = False
        For j = 0 To nbr_apres - 1
            spm_en_cours = STR_GetChamp(v_new_spm, "|", j)
            If num_poste_en_cours = Mid$(STR_GetChamp(spm_en_cours, ";", STR_GetNbchamp(spm_en_cours, ";") - 1), 2) Then
                poste_trouve = True
            End If
        Next j
        If Not poste_trouve Then
            GoTo lab_changements
        End If
    Next I

    Exit Sub

lab_changements:
    g_changements_importants = True

End Sub

Private Sub choisir_fonction()

    Dim n As Integer
    Dim rs As rdoResultset

    If Odbc_SelectV("SELECT FT_Num, FT_Libelle FROM FctTrav ORDER BY FT_Libelle", rs) = P_ERREUR Then
        Exit Sub
    End If
    If rs.EOF Then
        rs.Close
        Exit Sub
    End If

    Call CL_Init
    Call CL_InitTitreHelp("Choix d'une fonction", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    n = -1
    If g_numfct > 0 Then
        Call CL_AddLigne("Toutes les fonctions", 0, "", False)
        n = 0
    End If
    While Not rs.EOF
        n = n + 1
        Call CL_AddLigne(rs("FT_Libelle").Value, rs("FT_Num").Value, "", False)
        rs.MoveNext
    Wend
    Call CL_InitTaille(0, -20)
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        Exit Sub
    End If

    g_numfct = CL_liste.lignes(CL_liste.pointeur).num

End Sub

Private Function coordLiee_existe(ByVal v_ucnum As Long) As Boolean
' Vérifier si la coordonnée liée v_ucnum est déjà dans le grid
    Dim I As Integer

    With grdCoordLiees
        For I = 0 To .Rows - 1
            If .TextMatrix(I, GRDCL_CODE) = v_ucnum Then
                coordLiee_existe = True
                Exit Function
            End If
        Next I
    End With
    ' on n'a rien trouvé
    coordLiee_existe = False

End Function

Sub copier_photo()
    
    Dim ext As String, nomfich As String
    
    ' Effacer l'ancienne
    nomfich = p_CheminKW & "/kalibottin/photos/" & g_numutil
    Call KF_EffacerFichier(nomfich & ".gif", False)
    Call KF_EffacerFichier(nomfich & ".bmp", False)
    Call KF_EffacerFichier(nomfich & ".jpg", False)
    Call KF_EffacerFichier(nomfich & ".jpeg", False)

    ' Copier la nouvelle image
    If g_nom_fich_photo <> "" Then
        ext = Mid$(g_nom_fich_photo, InStrRev(g_nom_fich_photo, "."))
        Call KF_PutFichier(nomfich & ext, g_nom_fich_photo)
    End If
    
End Sub

Private Sub deplacer_txt(ByVal v_direction As String, ByVal v_row As Integer, ByVal v_col As Integer)
' ********************************************************************
' Sert à déplacer le TextBox(TXT_CACHE) selon les flèches de direction
' ********************************************************************
    If v_col = GRDC_NIVEAU Then ' COLCOLCOL
        txt(TXT_CACHE).MaxLength = 1
    Else
        txt(TXT_CACHE).MaxLength = 0
    End If
    With grdCoord
        Select Case v_direction
        Case "HAUT" '   ########################### VERS LE HAUT ##########################################"
            If v_row > 1 Then ' on n'est pas au TOP du grid
                .Row = .Row - 1
                Call positionner_txt(.CellLeft + .left, .CellTop + .Top, .CellWidth, .CellHeight, .TextMatrix(.Row, .col))
            End If
        Case "BAS" '    ########################### VERS LE BAS ##########################################"
            If v_row < .Rows - 1 Then ' on n'est pas au BOTTOM du grid
                .Row = .Row + 1
                Call positionner_txt(.CellLeft + .left, .CellTop + .Top, .CellWidth, .CellHeight, .TextMatrix(.Row, .col))
            End If
        Case "DROITE" ' ########################## VERS LA DROITE ########################################"
            'If v_col <> .Cols - 1 Then ' on n'est pas sur la derniere colonne
            If v_col <> GRDC_COMMENTAIRE Then ' on n'est pas sur la derniere colonne
                .col = .col + 1
                If .col <> GRDC_PRINCIPAL Then
                    Call positionner_txt(.CellLeft + .left, .CellTop + .Top, .CellWidth, .CellHeight, .TextMatrix(.Row, .col))
                Else
                    .SetFocus
                End If
            'ElseIf v_col = .Cols - 1 Then  ' on est sur la derniere colonne
            ElseIf v_col = GRDC_COMMENTAIRE Then  ' on est sur la derniere colonne
                If v_row < .Rows - 1 Then ' on n'est pas sur la derniere ligne
                    .Row = .Row + 1
                    .col = GRDC_VALEUR
                    Call positionner_txt(.CellLeft + .left, .CellTop + .Top, .CellWidth, .CellHeight, .TextMatrix(.Row, .col))
                ElseIf v_row = .Rows - 1 Then ' on est sur la derniere ligne
                    .Row = 1
                    .col = GRDC_VALEUR
                    Call positionner_txt(.CellLeft + .left, .CellTop + .Top, .CellWidth, .CellHeight, .TextMatrix(.Row, .col))
                End If
            End If
        Case "GAUCHE" ' ########################### VERS LA GAUCHE ########################################"
            If v_col > GRDC_VALEUR Then
                .col = .col - 1
                If .col <> GRDC_PRINCIPAL Then
                    Call positionner_txt(.CellLeft + .left, .CellTop + .Top, .CellWidth, .CellHeight, .TextMatrix(.Row, .col))
                Else
                    .SetFocus
                End If
            Else
                If v_row > 1 Then
                    .Row = .Row - 1
                    .col = .Cols - 1
                    .SetFocus
                End If
            End If
        End Select
    End With

End Sub

Private Function enlever_coordLiees(ByVal v_ucnum As Long) As Boolean
    Dim I As Integer

    With grdCoordLiees
        For I = 0 To .Rows - 1
            .Row = I
            If .TextMatrix(I, GRDCL_CODE) = v_ucnum Then
                Call supprimer_coordLiee
                enlever_coordLiees = True
                Exit Function
            End If
        Next I
    End With

    enlever_coordLiees = True

End Function

Private Function enregistrer_coordonnees(ByVal v_numutil As Long) As Integer
' *******************************************************************
' Enregistrer les coordonnées de la personne et ses coordonnées liées
' *******************************************************************
    Dim uc_principal_val As Boolean
    Dim I As Integer
    Dim lng As Long

    ' Supprimer les anciennes coordonnées SUPPRIMEES
    If g_numutil <> 0 Then
        For I = 0 To g_nbr_coord_supp - 1
            If Odbc_Delete("UtilCoordonnee", _
                           "UC_Num", _
                           "WHERE UC_Num=" & g_coord_supprimees(I), _
                           lng) = P_ERREUR Then
                GoTo lab_erreur
            End If
            ' pour les autres entités
            If Odbc_Delete("Coordonnee_Associee", _
                           "CA_Num", _
                           "WHERE CA_UCNum=" & g_coord_supprimees(I), _
                           lng) = P_ERREUR Then
                GoTo lab_erreur
            End If
'            If Odbc_Delete("UtilCoordonnee", _
'                           "WHERE UC_TypeNum=" & g_numutil & " AND UC_Type='U'", _
'                           lng) = P_ERREUR Then
'                GoTo lab_erreur
'            End If
        Next I
    End If
    With grdCoord
        For I = 1 To .Rows - 1
            .Row = I
            .col = GRDC_PRINCIPAL
            ' La coordonnée est-elle principale ?
            If .CellPicture = imglst.ListImages(IMG_COCHE).Picture Then uc_principal_val = True
            If .TextMatrix(I, GRDC_UCNUM) <> "" Then ' une mise à jour
                If Odbc_Update("UtilCoordonnee", "UC_Num", "WHERE UC_Num=" & .TextMatrix(I, GRDC_UCNUM), _
                                "UC_Valeur", .TextMatrix(I, GRDC_VALEUR), _
                                "UC_Comm", .TextMatrix(I, GRDC_COMMENTAIRE), _
                                "UC_Niveau", .TextMatrix(I, GRDC_NIVEAU), _
                                "UC_Principal", uc_principal_val, _
                                "UC_LstPoste", "") = P_ERREUR Then
                    GoTo lab_erreur
                End If
            Else ' une nouvelle coordonnée à ajouter
                If Odbc_AddNew("UtilCoordonnee", "UC_Num", "UC_Seq", False, lng, _
                               "UC_Type", "U", _
                               "UC_TypeNum", v_numutil, _
                               "UC_ZUNum", .TextMatrix(I, GRDC_ZUNUM), _
                               "UC_Valeur", .TextMatrix(I, GRDC_VALEUR), _
                               "UC_Comm", .TextMatrix(I, GRDC_COMMENTAIRE), _
                               "UC_Niveau", .TextMatrix(I, GRDC_NIVEAU), _
                               "UC_Principal", uc_principal_val, _
                               "UC_LstPoste", "") = P_ERREUR Then
                    GoTo lab_erreur
                End If
            End If
            uc_principal_val = False
        Next I
    End With

    ' Supprimer les anciennes coordonnées associées
    If g_numutil <> 0 Then
        If Odbc_Delete("Coordonnee_Associee", _
                       "CA_Num", _
                       "WHERE CA_UCTypeNum=" & g_numutil & " AND CA_UCType='U'", _
                       lng) = P_ERREUR Then
            GoTo lab_erreur
        End If
    End If
    With grdCoordLiees
        For I = 0 To .Rows - 1
            .Row = I
            .col = GRDCL_PRINCIPAL
            ' La coordonnée est-elle principale ?
            If .CellPicture = imglst.ListImages(IMG_COCHE).Picture Then uc_principal_val = True
            If Odbc_AddNew("Coordonnee_Associee", "CA_Num", "CA_Seq", False, lng, _
                           "CA_UCType", "U", _
                           "CA_UCTypeNum", v_numutil, _
                           "CA_UCNum", .TextMatrix(I, GRDCL_CODE), _
                           "CA_Principal", uc_principal_val) = P_ERREUR Then
                GoTo lab_erreur
            End If
            uc_principal_val = False
        Next I
    End With

    enregistrer_coordonnees = P_OK
    Exit Function

lab_erreur:
    enregistrer_coordonnees = P_ERREUR

End Function

Private Sub evalue_btninfosuppl()

    Dim sql As String
    Dim lnb As Long
    
    sql = "SELECT count(*) FROM InfoSupplEntite WHERE ISE_TypeNum=" & g_numutil & " AND ISE_Type='U'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        Exit Sub
    End If

    If lnb = 0 Then
        Set cmd(CMD_INFO_SUPPL).Picture = imglst.ListImages(IMG_INFO_OFF).Picture
        cmd(CMD_INFO_SUPPL).ToolTipText = "Cette personne ne possède aucune information supplémentaire"
    Else
        Set cmd(CMD_INFO_SUPPL).Picture = imglst.ListImages(IMG_INFO_ON).Picture
        cmd(CMD_INFO_SUPPL).ToolTipText = "Cette personne possède des informations supplémentaires"
    End If
    
End Sub

Private Function get_str_coordonnees() As String
' Retourner la chaine contenant les coordonnees principales et leurs valeurs trouvées
' dans UtilCoordonnee sous la forme "x=XX;y=YY;....n=NN;" où les miniscules sont les ZoneUtil.ZU_Num

    Dim sql As String
    Dim rs As rdoResultset

    sql = "SELECT ZU_Code, UC_Valeur FROM UtilCoordonnee, ZoneUtil " _
        & " WHERE UC_TypeNum=" & g_numutil _
        & " AND UC_Type='U' AND ZU_Type='C'" _
        & " AND UC_ZUNum=ZU_Num" _
        & " AND UC_Principal=True" _
        & " ORDER BY ZU_Code"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Function
    End If
    While Not rs.EOF
        get_str_coordonnees = get_str_coordonnees & rs("ZU_Code") & "=" & rs("UC_Valeur") & ";"
        rs.MoveNext
    Wend
    rs.Close

End Function

Private Sub getPhotoUtilisateur()
    
    Dim ext As String
    Dim n As Integer, I As Integer
    
    n = STR_GetNbchamp(LISTE_FICHIERS_IMAGES, ";")
    For I = 0 To n - 1
        ext = Mid$(STR_GetChamp(LISTE_FICHIERS_IMAGES, ";", I), 2)
        If KF_FichierExiste(p_CheminKW & "/kalibottin/photos/" & g_numutil & ".*") Then
            Set cmd(CMD_PHOTO).Picture = imglst.ListImages(IMG_PHOTO_ON).Picture
            cmd(CMD_PHOTO).ToolTipText = "Il y a une photo associée - Clic gauche pour changer le fichier, Clic droit pour le supprimer "
            Exit Sub
        End If
    Next I
        
    Set cmd(CMD_PHOTO).Picture = imglst.ListImages(IMG_PHOTO_OFF).Picture
    cmd(CMD_PHOTO).ToolTipText = "Pas de photo - Clic gauche pour choisir le fichier"

End Sub

Private Sub init_focus()

    txt(TXT_CACHE).Visible = False
    Select Case sst.Tab
    Case ONGLET_GENERAL
        txt(TXT_NOM).SetFocus
    Case ONGLET_PROFIL
        If grdLabo.Visible Then
            grdLabo.SetFocus
        Else
            tvSect.SetFocus
        End If
    End Select

End Sub

Private Function init_grD_Site() As Integer

    Dim sql As String
    Dim I As Integer
    Dim rs As rdoResultset

    If p_NbLabo > 1 Then
        sql = "SELECT * FROM Laboratoire" _
            & " ORDER BY L_Code"
        If Odbc_Select(sql, rs) = P_ERREUR Then
            init_grD_Site = P_ERREUR
            Exit Function
        End If
        I = 0
        While Not rs.EOF
            grdLabo.AddItem rs("L_Num").Value & vbTab _
                            & False & vbTab _
                            & rs("L_Code").Value
            I = I + 1
            rs.MoveNext
        Wend
        rs.Close
    Else
        sql = "SELECT * FROM Laboratoire"
        If Odbc_Select(sql, rs) = P_ERREUR Then
            init_grD_Site = P_ERREUR
            Exit Function
        End If
        grdLabo.AddItem rs("L_Num").Value & vbTab _
                        & True & vbTab _
                        & rs("L_Code").Value & vbTab _
                        & "" & vbTab _
                        & "Principal"
        lblLabo.Visible = False
        grdLabo.Visible = False
    End If

    init_grD_Site = P_OK

End Function

Private Sub initialiser()

    Dim col As Integer, I As Integer

    g_nom_fich_photo = ""
    ' initialiser le tableau des coordonnées
    Erase g_coord_supprimees
    g_nbr_coord_supp = 0

    cmd(CMD_MOINS_TYPE).Enabled = False
    cmd(CMD_MOINS_COORDLIEES).Enabled = False
    If g_numutil <> 0 Then cmd(CMD_HISTORIQUE).Visible = True
    g_position_txt_cache = ""
    grdCoord.ScrollTrack = True

    With grdCoord
        txt(TXT_CACHE).left = .CellLeft
        txt(TXT_CACHE).Top = .CellTop

        Me.txt(TXT_NOM).SetFocus

        .Rows = 1
        .FormatString = "|Type de coordonnée|Valeur|Niveau de confidentialité|Principal dans son type|Commentaire"
        .RowHeight(0) = 750
        .ColWidth(GRDC_ZUNUM) = 0
        .ColWidth(GRDC_TYPE) = 2250
        .ColWidth(GRDC_VALEUR) = 2300
        .ColWidth(GRDC_NIVEAU) = 1500
        .ColWidth(GRDC_PRINCIPAL) = 1650
        .ColWidth(GRDC_COMMENTAIRE) = 2580
        .ColWidth(GRDC_UCNUM) = 0
        .Row = 0
        For I = 0 To .Cols - 1
            .col = I
            .CellFontBold = True
            .ColAlignment(I) = 4
        Next I
        .col = GRDC_ZUNUM
    End With

    With grdCoordLiees
        '.SelectionMode = flexSelectionByRow
        .Rows = 1
        .RowHeight(0) = 750
        .ColWidth(GRDCL_CODE) = 0
        .ColWidth(GRDCL_TYPE) = 2250
        .ColWidth(GRDCL_VALEUR) = 2300
        .ColWidth(GRDCL_NIVEAU) = 1500
        .ColWidth(GRDCL_PRINCIPAL) = 1650
        .ColWidth(GRDCL_IDENTITE) = 2330
        .Row = 0
        For I = 0 To .Cols - 1
            .col = I
            .CellFontBold = True
            .ColAlignment(I) = 4
        Next I
        .col = GRDCL_CODE
    End With
    
    With grdLabo
        .Cols = 5
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 2800
        .ColWidth(3) = 500
        .ColWidth(4) = 1000
    End With

    g_coordonnees_avant = ""
    g_coordonnees_apres = ""
    g_numfct = 0
    g_spm = ""
    g_pobPrinc = 0

    ' le btn "Informations Supplémentaires"
    If g_numutil > 0 Then
        cmd(CMD_INFO_SUPPL).Visible = True
    Else
        cmd(CMD_INFO_SUPPL).Visible = False
    End If

    If afficher_utilisateur() = P_ERREUR Then
        Unload Me
        Exit Sub
    End If

    cmd(CMD_OK).Enabled = False

End Sub

Private Sub inverser_etat_labo()

    g_mode_saisie = False

    If grdLabo.TextMatrix(grdLabo.Row, GRDL_ESTLABO) = True Then
        grdLabo.TextMatrix(grdLabo.Row, GRDL_ESTLABO) = False
        grdLabo.col = GRDL_IMG_ESTLABO
        Set grdLabo.CellPicture = CM_LoadPicture("")
    Else
        grdLabo.TextMatrix(grdLabo.Row, GRDL_ESTLABO) = True
        grdLabo.col = GRDL_IMG_ESTLABO
        Set grdLabo.CellPicture = imglst.ListImages(IMG_COCHE).Picture
    End If
    cmd(CMD_OK).Enabled = True

    g_mode_saisie = True

End Sub

Private Sub inverser_laboprinc()

    Dim I As Integer

    g_mode_saisie = False

    If grdLabo.TextMatrix(grdLabo.Row, GRDL_LABOPRINC) = "" Then
        For I = 0 To grdLabo.Rows - 1
            grdLabo.TextMatrix(I, GRDL_LABOPRINC) = ""
        Next I
        grdLabo.TextMatrix(grdLabo.Row, GRDL_LABOPRINC) = "Principal"
        If grdLabo.TextMatrix(grdLabo.Row, GRDL_ESTLABO) = False Then
            grdLabo.TextMatrix(grdLabo.Row, GRDL_ESTLABO) = True
            grdLabo.col = GRDL_IMG_ESTLABO
            Set grdLabo.CellPicture = imglst.ListImages(IMG_COCHE).Picture
        End If
    Else
        grdLabo.TextMatrix(grdLabo.Row, GRDL_LABOPRINC) = ""
    End If

    grdLabo.col = GRDL_LABOPRINC
    grdLabo.ColSel = GRDL_LABOPRINC
    cmd(CMD_OK).Enabled = True

    g_mode_saisie = True

End Sub

Private Function labo_concerne(ByVal v_dlabo As String, _
                               ByVal v_ulabo As String) As Boolean

    Dim s As String, s_u As String, s_d As String
    Dim I As Integer, j As Integer, n As Integer, m As Integer

    n = STR_GetNbchamp(v_ulabo, ";")
    m = STR_GetNbchamp(v_dlabo, ";")
    For I = 0 To n - 1
        s_u = STR_GetChamp(v_ulabo, ";", I) & ";"
        For j = 0 To m - 1
            s_d = STR_GetChamp(v_dlabo, ";", j) & ";"
            If s_u = s_d Then
                labo_concerne = True
                Exit Function
            End If
        Next j
    Next I
    labo_concerne = False

End Function

Private Sub positionner_txt(ByVal v_txt_left As Long, ByVal v_txt_top As Long, _
                            ByVal v_txt_width As Long, ByVal v_txt_height As Long, _
                            ByVal v_txt_text As String)

    ' Les colonnes PRINCIPAL et COMMENTAIRE ne sont pas éditables directement
    If grdCoord.col = GRDC_PRINCIPAL Or grdCoord.col = GRDC_COMMENTAIRE Then
        txt(TXT_CACHE).Visible = False
        grdCoord.SetFocus
        Exit Sub
    End If

    With txt(TXT_CACHE)
        .left = v_txt_left
        .Top = v_txt_top
        .width = v_txt_width
        .Height = v_txt_height
        .Text = v_txt_text
        .SelLength = Len(.Text)
        .ZOrder 0
        .Visible = True
        .SetFocus
    End With

End Sub

Private Sub prm_service()
' ********************************************
' Ajout d'un nouveau poste pour cette personne
' ********************************************
    Dim s As String, s1 As String
    Dim lib As String, sret As String, ssite As String
    Dim au_moins_un As Boolean
    Dim I As Integer, j As Integer, nbch As Integer, n As Integer
    Dim numlabo As Long, num As Long
    Dim spm As Variant
    Dim nd As Node
    Dim frm As Form

    Call CL_Init
    Call build_SPM_Fct(spm, s)
    nbch = STR_GetNbchamp(spm, "|")
    n = 0
    For I = 1 To nbch
        s = STR_GetChamp(spm, "|", I - 1)
        ReDim Preserve CL_liste.lignes(n)
        CL_liste.lignes(n).texte = s
        CL_liste.lignes(n).fmodif = True
        n = n + 1
    Next I

    ssite = ""
    au_moins_un = False
    If p_NbLabo > 1 Then
        For I = 0 To grdLabo.Rows - 1
            If grdLabo.TextMatrix(I, GRDL_ESTLABO) = True Then
                au_moins_un = True
                ssite = ssite & grdLabo.TextMatrix(I, GRDL_NUMLABO) & ";"
            End If
        Next I
        If Not au_moins_un Then
            Call MsgBox("Aucun site n'est indiqué pour cette personne.", vbExclamation + vbOKOnly, "")
            Exit Sub
        End If
    Else
        ssite = ssite + "1;"
    End If

    Set frm = PrmService
    sret = PrmService.AppelFrm("Choix des postes", "S", True, ssite, "P", False)
    Set frm = Nothing
    p_NumLabo = numlabo
    If sret = "" Then
        Exit Sub
    End If

    cmd(CMD_OK).Enabled = True

    tvSect.Nodes.Clear
    n = CLng(Mid$(sret, 2))
    If n = 0 Then
        Exit Sub
    End If
    For I = 0 To n - 1
        s = CL_liste.lignes(I).texte
        nbch = STR_GetNbchamp(s, ";")
        For j = 1 To nbch
            s1 = STR_GetChamp(s, ";", j - 1)
            If TV_NodeExiste(tvSect, s1, nd) = P_NON Then
                num = Mid$(s1, 2)
                If left$(s1, 1) = "S" Then
                    If P_RecupSrvNom(num, lib) = P_ERREUR Then
                        Exit Sub
                    End If
                    If j = 1 Then
                        Set nd = tvSect.Nodes.Add(, tvwChild, s1, lib, IMGT_SERVICE, IMGT_SERVICE)
                    Else
                        Set nd = tvSect.Nodes.Add(nd, tvwChild, s1, lib, IMGT_SERVICE, IMGT_SERVICE)
                    End If
                Else
                    If P_RecupPosteNom(num, lib) = P_ERREUR Then
                        Exit Sub
                    End If
                    If n = 1 Then 'on selectionné un seul poste => principal par défaut !
                        g_pobPrinc = num
                    End If
                    If num = g_pobPrinc Then
                        ' on ne passe ici qu'une seule fois = un seul poste principal
                        Set nd = tvSect.Nodes.Add(nd, tvwChild, s1, lib, IMGT_POSTE_SEL, IMGT_POSTE_SEL)
                    Else
                        Set nd = tvSect.Nodes.Add(nd, tvwChild, s1, lib, IMGT_POSTE, IMGT_POSTE)
                    End If
                End If
                nd.Expanded = True
            End If
        Next j
    Next I

    tvSect.SetFocus
    If tvSect.Nodes.Count > 0 Then
        tvSect.BorderStyle = ccFixedSingle
        cmd(CMD_MOINS_SPM).Visible = True
        Set tvSect.SelectedItem = tvSect.Nodes(1)
        SendKeys "{PGDN}"
        SendKeys "{HOME}"
        DoEvents
    End If

End Sub

Private Sub quitter(ByVal v_bforce As Boolean)

    Dim reponse As Integer

    If v_bforce Then
        g_changements_importants = False
        Unload Me
        Exit Sub
    End If

    If cmd(CMD_OK).Enabled Then
        reponse = MsgBox("Des modifications ont été effectuées !" & vbLf & vbLf & "Confirmez-vous l'abandon ?", _
                          vbYesNo + vbDefaultButton2 + vbQuestion)
        If reponse = vbNo Then Exit Sub
    End If

    g_changements_importants = False
    Unload Me

End Sub

Private Function remplir_mvt_coord() As Integer
' Remplir les mouvements des coordonnées dans [UtilMouvement]

    Dim str_avant As String, str_apres As String, zu_code_avant As String, zu_code_apres As String, _
        uc_valeur_avant As String, uc_valeur_apres As String
    Dim code_disparu As Boolean
    Dim nbr_avant As Integer, nbr_apres As Integer, I As Integer, j As Integer

    ' *************************** MODIFICATIONS / SUPPRESSIONS ***************************
    ' parcourrir les anciennes coordonnées
    nbr_avant = STR_GetNbchamp(g_coordonnees_avant, ";")
    For I = 0 To nbr_avant - 1
        str_avant = STR_GetChamp(g_coordonnees_avant, ";", I)
        zu_code_avant = STR_GetChamp(str_avant, "=", 0)
        uc_valeur_avant = STR_GetChamp(str_avant, "=", 1)
        ' parcourrir les nouvelles coordonnées
        nbr_apres = STR_GetNbchamp(g_coordonnees_apres, ";")
        For j = 0 To nbr_apres - 1
            code_disparu = True
            str_apres = STR_GetChamp(g_coordonnees_apres, ";", j)
            zu_code_apres = STR_GetChamp(str_apres, "=", 0)
            ' le même type de coordonnée ?
            If zu_code_avant = zu_code_apres Then
                code_disparu = False
                uc_valeur_apres = STR_GetChamp(str_apres, "=", 1)
                If uc_valeur_avant <> uc_valeur_apres Then ' ------ MODIFICATION
                    If P_InsertIntoUtilmouvement(g_numutil, "M", zu_code_avant & "=" & uc_valeur_avant & ";" & uc_valeur_apres & ";", p_appli_kalibottin) = P_ERREUR Then
                        GoTo lab_erreur
                    End If
                End If
                GoTo lab_i_suivant
            End If
        Next j
        If code_disparu Then ' ------------------------------------ SUPPRESSION
            ' le type de coordonnée a disparu !
            If P_InsertIntoUtilmouvement(g_numutil, "M", zu_code_avant & "=" & uc_valeur_avant & ";" & ";", p_appli_kalibottin) = P_ERREUR Then
                GoTo lab_erreur
            End If
        End If
lab_i_suivant:
    Next I
    ' ************************************** AJOUTS **************************************
    ' parcourrir les nouvelles coordonnées
    For j = 0 To nbr_apres - 1
        str_apres = STR_GetChamp(g_coordonnees_apres, ";", j)
        zu_code_apres = STR_GetChamp(str_apres, "=", 0)
        uc_valeur_apres = STR_GetChamp(str_apres, "=", 1)
        ' parcourrir les anciennes coordonnées
        For I = 0 To nbr_avant - 1
            str_avant = STR_GetChamp(g_coordonnees_avant, ";", I)
            zu_code_avant = STR_GetChamp(str_avant, "=", 0)
            ' le type de coordonnée existe-il déjà ?
            If zu_code_apres = zu_code_avant Then
                GoTo lab_j_suivant
            End If
        Next I
        ' un nouveau type de coordonnée
        If P_InsertIntoUtilmouvement(g_numutil, "M", zu_code_apres & "=" & ";" & uc_valeur_apres & ";", p_appli_kalibottin) = P_ERREUR Then
            GoTo lab_erreur
        End If
lab_j_suivant:
    Next j

    remplir_mvt_coord = P_OK
    Exit Function

lab_erreur:
    remplir_mvt_coord = P_ERREUR

End Function

Private Function remplir_utilmouvement(ByVal v_spm As Variant, ByVal v_numutil As Long) As Integer
'*******************************************************************
' Appelée depuis valider() afin de renseigner la table UtilMouvement
'*******************************************************************
    Dim rendre_non_importe As Boolean

    rendre_non_importe = False
    If g_numutil = 0 Then ' **************** nouvel utilisateur *****************************************************
        If P_InsertIntoUtilmouvement(v_numutil, "C", "", p_appli_kalibottin) = P_ERREUR Then
            GoTo lab_erreur
        End If
    Else ' ********************************* utilisateur existant ***************************************************
        If g_old_nom <> UCase$(txt(TXT_NOM).Text) Then ' ------------------- NOM
            If P_InsertIntoUtilmouvement(g_numutil, "M", "NOM=" & g_old_nom & ";" & txt(TXT_NOM).Text & ";", p_appli_kalibottin) = P_ERREUR Then
                GoTo lab_erreur
            End If
'            rendre_non_importe = True
        End If
        If UCase$(g_old_prenom) <> UCase$(txt(TXT_PRENOM).Text) Then ' ----- PRENOM
            If P_InsertIntoUtilmouvement(g_numutil, "M", "PRENOM=" & g_old_prenom & ";" & txt(TXT_PRENOM).Text & ";", p_appli_kalibottin) = P_ERREUR Then
                GoTo lab_erreur
            End If
'            rendre_non_importe = True
        End If
        If g_old_matricule <> txt(TXT_MATRICULE).Text Then ' --------------- MATRICULE
            If P_InsertIntoUtilmouvement(g_numutil, "M", "MATRICULE=" & g_old_matricule & ";" & txt(TXT_MATRICULE).Text & ";", p_appli_kalibottin) = P_ERREUR Then
                GoTo lab_erreur
            End If
            rendre_non_importe = True
        End If
        If g_old_code <> txt(TXT_CODE).Text Then ' ------------------------- CODE
            If P_InsertIntoUtilmouvement(g_numutil, "M", "CODE=" & g_old_code & ";" & txt(TXT_CODE).Text & ";", p_appli_kalibottin) = P_ERREUR Then
                GoTo lab_erreur
            End If
        End If
        If g_old_spm <> v_spm Then ' --------------------------------------- POSTE
            If P_InsertIntoUtilmouvement(g_numutil, "M", P_get_poste_modif(g_old_spm, v_spm), p_appli_kalibottin) = P_ERREUR Then
                GoTo lab_erreur
            End If
            ' en ce qui concerne (g_changements_importants): voir changements_importants()
            ' qui est appelée depuis valider() juste avant d'enregistrer les nouvelles coordonnées
        End If
        If chk(CHK_ACTIF).Value = 0 Then ' --------------------------------- INACTIVE
            If g_old_active Then
                If P_InsertIntoUtilmouvement(g_numutil, "I", "", p_appli_kalibottin) = P_ERREUR Then
                    GoTo lab_erreur
                End If
                rendre_non_importe = True
                Call Shell(p_chemin_appli & "\Lance.exe " & p_chemin_appli & ";KaliDoc;" & p_nom_fichier_ini_kalidoc & ";CONNEXION=" & p_NumUtil & ";DESACTIVER_PERS=" & g_numutil)
                Call SYS_Sleep(1000)
            End If
        Else ' chk(CHK_ACTIF).Value = 1 ------------------------------------ ACTIVE
            If Not g_old_active Then
                If P_InsertIntoUtilmouvement(g_numutil, "A", "", p_appli_kalibottin) = P_ERREUR Then
                    GoTo lab_erreur
                End If
                rendre_non_importe = True
                Call Shell(p_chemin_appli & "\Lance.exe " & p_chemin_appli & ";KaliDoc;" & p_nom_fichier_ini_kalidoc & ";CONNEXION=" & p_NumUtil & ";ACTIVER_PERS=" & g_numutil)
                Call SYS_Sleep(1000)
            End If
        End If
        If rendre_non_importe Then
            If Odbc_Update("Utilisateur", "U_Num", _
                           "WHERE U_kb_actif=True AND U_Num=" & g_numutil, _
                           "U_Importe", False) = P_ERREUR Then
                GoTo lab_erreur
            End If
            g_changements_importants = True
        End If
        g_coordonnees_apres = get_str_coordonnees() ' ---------------------- COORDONNEES PRINCIPALES
        If g_coordonnees_avant <> g_coordonnees_apres Then
            If remplir_mvt_coord = P_ERREUR Then GoTo lab_erreur
        End If
    End If

    remplir_utilmouvement = P_OK
    Exit Function

lab_erreur:
    remplir_utilmouvement = P_ERREUR

End Function

Private Function set_coordLiees(ByVal v_ucnum As Long, ByVal v_principal As Boolean) As Integer
    Dim sql As String, uc_num As String, zu_libelle As String, _
        uc_valeur As String, uc_niveau As String, uc_type As String, _
        u_nom As String, u_prenom As String, libelle_entite As String
    Dim uc_typenum As Long

    sql = "SELECT UC_Num, ZU_Libelle, UC_Valeur, UC_Niveau, UC_Type, UC_TypeNum FROM ZoneUtil, UtilCoordonnee" & _
        " WHERE UC_ZUNum=ZU_Num AND UC_Num=" & v_ucnum
    If Odbc_RecupVal(sql, uc_num, zu_libelle, uc_valeur, uc_niveau, uc_type, uc_typenum) = P_ERREUR Then
        GoTo lab_erreur
    End If
    With grdCoordLiees
        .AddItem (v_ucnum)
        .TextMatrix(.Rows - 1, GRDCL_CODE) = uc_num
        .TextMatrix(.Rows - 1, GRDCL_TYPE) = zu_libelle
        .TextMatrix(.Rows - 1, GRDCL_VALEUR) = uc_valeur
        .TextMatrix(.Rows - 1, GRDCL_NIVEAU) = uc_niveau
        .Row = .Rows - 1
        .col = GRDCL_PRINCIPAL
        .CellAlignment = 4
        If v_principal Then
            Set .CellPicture = imglst.ListImages(IMG_COCHE).Picture
        Else
            Set .CellPicture = imglst.ListImages(IMG_PASCOCHE).Picture
        End If
        .col = GRDCL_IDENTITE
        If uc_type = "U" Then ' une personne
            If Odbc_RecupVal("SELECT U_Nom, U_Prenom FROM Utilisateur WHERE U_kb_actif=True AND U_Num=" & uc_typenum, _
                            u_nom, u_prenom) = P_ERREUR Then
                GoTo lab_erreur
            End If
            libelle_entite = u_nom & " " & u_prenom
        ElseIf uc_type = "P" Then ' poste
            libelle_entite = P_get_lib_srv_poste(uc_typenum, P_POSTE)
        Else ' pièce
            libelle_entite = P_get_nom_piece(uc_typenum)
        End If
        .TextMatrix(.Row, .col) = libelle_entite
    End With

    set_coordLiees = P_OK
    Exit Function

lab_erreur:
    set_coordLiees = P_ERREUR

End Function

Private Sub supprimer_poste()

    Dim encore As Boolean
    Dim nd As Node, ndP As Node

    If tvSect.Nodes.Count = 0 Then Exit Sub

    On Error GoTo err_tv
    Set nd = tvSect.SelectedItem
    On Error GoTo 0
    If InStr(nd.Text, "(-> Synchro)") > 0 Then
        MsgBox "Ce poste secondaire est issu d'une synchronisation automatique" & Chr(13) & Chr(10) & "Il risque donc de ré-apparaitre lors de la prochaine synchronisation", vbExclamation
    End If
    ' Prévoir la suppression du poste principal
    If nd.image = IMGT_POSTE_SEL Then ' on a selectionné le poste principal
        g_pobPrinc = 0
    End If

    Do
        encore = True
        Set ndP = nd
        ' A-t-on un service père, si oui le récuperer dans ndp
        If TV_NodeParent(ndP) Then
            If ndP.Children > 1 Then
            ' Le service père a des postes/services => n'enlever que le noeud selectionné
                encore = False
            Else ' remonter au noeud parent
                Set nd = ndP
            End If
        Else ' on a selectionné un poste => l'enlever
            encore = False
        End If
    Loop Until Not encore

    tvSect.Nodes.Remove (nd.Index)

    tvSect.Refresh
    cmd(CMD_OK).Enabled = True
    If tvSect.Nodes.Count = 0 Then
        cmd(CMD_MOINS_SPM).Visible = False
    End If

    Exit Sub

err_tv:
    MsgBox "Vous devez sélectionner l'élément à supprimer", vbOKOnly, ""
    On Error GoTo 0

End Sub

Private Sub supprimer_type_coord()

    Dim I As Integer, row_en_cours As Integer
    Dim num_en_cours As Long

    With grdCoord
        row_en_cours = .Row
        If .TextMatrix(row_en_cours, GRDC_UCNUM) <> "" Then
        ' ne pas prendre en compte les coordonnées nouvellement ajoutées
            ReDim Preserve g_coord_supprimees(g_nbr_coord_supp)
            g_coord_supprimees(g_nbr_coord_supp) = .TextMatrix(row_en_cours, GRDC_UCNUM)
            g_nbr_coord_supp = g_nbr_coord_supp + 1
        End If
        If .Rows = 2 Then ' On a une seule ligne + ligne fixe
            .Rows = 1     ' On ne laisse que la ligne fixe
        ElseIf .Row > 0 Then ' On a plusieurs lignes
            num_en_cours = .TextMatrix(.Row, GRDC_ZUNUM)
            .col = GRDC_PRINCIPAL
            If .CellPicture = imglst.ListImages(IMG_COCHE).Picture Then
                For I = 1 To .Rows - 1
                    .Row = I
                    .col = GRDC_PRINCIPAL
                    If .TextMatrix(I, GRDC_ZUNUM) = num_en_cours And I <> row_en_cours Then
                        Set .CellPicture = imglst.ListImages(IMG_COCHE).Picture
                        Exit For
                    End If
                Next I
            End If
            Call .RemoveItem(row_en_cours)
            ' Remettre la taille du grid par défaut, et la disposition des boutons
            If .Rows <= NBRMAX_ROWS And .width = LARGEUR_GRID_PAR_DEFAUT + 255 Then
                '.width = .width - 255
                'cmd(CMD_PLUS_TYPE).left = LEFT_CMD_PAR_DEFAUT
                'cmd(CMD_MOINS_TYPE).left = LEFT_CMD_PAR_DEFAUT
                .ColWidth(GRDC_COMMENTAIRE) = 2580
            End If
        End If
        ' il ne reste plus de coordonnées dans le tableau
        If .Rows - 1 = 0 Then .Enabled = False
    End With

    g_position_txt_cache = ""
    cmd(CMD_OK).Enabled = True

End Sub

Private Sub supprimer_coordLiee()
    Dim row_en_cours As Integer

    With grdCoordLiees
        row_en_cours = .Row
        If .Rows = 1 Then
            .Rows = 0
        Else
            Call .RemoveItem(row_en_cours)
        End If
        If .Rows = 0 Then .Enabled = False
    End With
    cmd(CMD_OK).Enabled = True

End Sub

Private Sub parcourir()

    Dim nom_fich As String
    
    ' La recherche du fichier image
    nom_fich = Com_ChoixFichier.AppelFrm("Sélectionnez la photo de cette personne", "", "", _
                    LISTE_FICHIERS_IMAGES, False)
    ' On n'active le bouton Enregistrer que si on a bien choisi une image
    If nom_fich <> "" Then
        cmd(CMD_PHOTO).Picture = imglst.ListImages(IMG_PHOTO_ON).Picture
        cmd(CMD_PHOTO).ToolTipText = "Il y a une photo associée - Clic gauche pour changer le fichier, Clic droit pour le supprimer "
        g_nom_fich_photo = nom_fich
        cmd(CMD_OK).Enabled = True
    End If

End Sub

Private Sub type_info_suppl()
' gérer les types d'info suppl pour cette personne
    Dim frm As Form

    Set frm = PrmPersonne
    PrmTypeInfSupPersonne.AppelFrm (g_numutil)
    Set frm = Nothing

End Sub

Private Function valider() As Integer
' **********************************************************
' Enregistrer toutes les modifications si elles sont valides
' **********************************************************
    Dim ssite As String, sfct As String
    Dim num_util As Long, numlabop As Long, lng As Long
    Dim spm As Variant
    Dim frm As Form

    ' ************************ VÉRIFICATIONS ************************
    If verifier_tous_champs(numlabop, ssite, sfct, spm) = P_NON Then
        valider = P_NON
        Exit Function
    End If
    If g_numutil <> 0 Then
        If chk(CHK_ACTIF) = 0 Then
            If verifier_inactive() = P_ERREUR Then
                valider = P_NON
                Exit Function
            End If
        End If
    End If

    ' *********************** ENREGISTREMENTS ***********************
    If Odbc_BeginTrans() = P_ERREUR Then
        valider = P_ERREUR
        Exit Function
    End If

    If g_numutil = 0 Then ' --- Ajouter la nouvelle personne ---
        If Odbc_AddNew("Utilisateur", "U_Num", "U_Seq", True, num_util, _
                       "U_Nom", UCase$(txt(TXT_NOM).Text), _
                       "U_Prenom", txt(TXT_PRENOM).Text, _
                       "U_NomJunon", txt(TXT_NOM_JUNON).Text, _
                       "U_PrenomJunon", txt(TXT_PRENON_JUNON).Text, _
                       "U_Prefixe", "", _
                       "U_Actif", IIf(chk(CHK_ACTIF).Value = 1, True, False), _
                       "U_Externe", IIf(chk(CHK_EXTERNE).Value = 1, True, False), _
                       "U_ExterneFich", IIf(chk(CHK_EXTERNE_IMPORT).Value = 1, True, False), _
                       "U_Importe", False, "U_AR", True, _
                       "U_SPM", spm, "U_FctTrav", sfct, _
                       "U_Labo", ssite, "U_LNumPrinc", numlabop, _
                       "U_Matricule", txt(TXT_MATRICULE).Text, _
                       "U_DONumLast", 0, _
                       "U_LstDocs", "", _
                       "U_CATPNum", 0, _
                       "U_DateDebEmbauche", Null, _
                       "U_DateFinEmbauche", Null, _
                       "U_CTRAVNum", 0, "U_NbHeures", 0, _
                       "U_BaseHeures", 0, _
                       "U_POTNumNext", 0, "U_LNumNext", 0, "U_NoSemNext", 0, _
                       "U_Po_Princ", g_pobPrinc, _
                       "U_kw_mailauth", True, _
                       "U_kb_actif", True, _
                       "U_Fictif", False) = P_ERREUR Then
            GoTo err_enreg
        End If
        ' Code + Mot de Passe
        If Odbc_AddNew("UtilAppli", _
                       "UAPP_Num", _
                       "UAPP_Seq", _
                       False, _
                       lng, _
                       "UAPP_APPNum", p_appli_kalidoc, _
                       "UAPP_UNum", num_util, _
                       "UAPP_Code", UCase$(txt(TXT_CODE).Text), _
                       "UAPP_TypeCrypt", p_Mode_Auth_UtilAppli, _
                       "UAPP_MotPasse", STR_Crypter_New(UCase(txt(TXT_MPASSE).Text))) = P_ERREUR Then
            GoTo err_enreg
        End If
        ' enregistrer les lignes du gridCoord
        If enregistrer_coordonnees(num_util) = P_ERREUR Then
            GoTo err_enreg
        End If
    Else ' --- Mettre à jour un utilisateur existant ---
        num_util = g_numutil
        Call changements_importants(spm, g_old_spm)
        ' Enregistrer la partie uitlisateur
        If Odbc_Update("Utilisateur", _
                       "U_Num", _
                       "WHERE U_Num=" & g_numutil, _
                       "U_Nom", UCase$(txt(TXT_NOM).Text), _
                       "U_Prenom", txt(TXT_PRENOM).Text, _
                       "U_NomJunon", UCase$(txt(TXT_NOM_JUNON).Text), _
                       "U_PrenomJunon", txt(TXT_PRENON_JUNON).Text, _
                       "U_Actif", IIf(chk(CHK_ACTIF).Value = 1, True, False), _
                       "U_Externe", IIf(chk(CHK_EXTERNE).Value = 1, True, False), _
                       "U_ExterneFich", IIf(chk(CHK_EXTERNE_IMPORT).Value = 1, True, False), _
                       "U_Matricule", txt(TXT_MATRICULE).Text, _
                       "U_SPM", spm, _
                       "U_FctTrav", sfct, _
                       "U_Labo", ssite, _
                       "U_LNumPrinc", numlabop, _
                       "U_Po_Princ", g_pobPrinc) = P_ERREUR Then
            GoTo err_enreg
        End If
        ' Code + Mot de Passe
        If Odbc_Update("UtilAppli", _
                       "UAPP_Num", _
                       "WHERE UAPP_UNum=" & g_numutil & " AND UAPP_APPNum=" & p_appli_kalidoc, _
                       "UAPP_Code", UCase$(txt(TXT_CODE).Text), _
                       "UAPP_MotPasse", STR_Crypter_New(UCase(txt(TXT_MPASSE).Text)), _
                       "UAPP_TypeCrypt", p_Mode_Auth_UtilAppli) = P_ERREUR Then
            GoTo err_enreg
        End If
        ' enregistrer les lignes du gridCoord
        If enregistrer_coordonnees(g_numutil) = P_ERREUR Then
            GoTo err_enreg
        End If
    End If

    ' la photo
    Call copier_photo
    
lab_commit:
    If Odbc_CommitTrans() = P_ERREUR Then
        valider = P_ERREUR
        Exit Function
    End If

    ' Renseigner la table UtilMouvements (après le Commit() afin de récupérer les nouvelles coordonnées)
    If remplir_utilmouvement(spm, num_util) = P_ERREUR Then
        GoTo err_enreg
    End If

    If g_numutil = 0 Then
        Call Shell(p_chemin_appli & "\Lance.exe " & p_chemin_appli & ";KaliDoc;" & p_nom_fichier_ini_kalidoc & ";CONNEXION=" & p_NumUtil & ";CREERPERS=" & num_util)
        SYS_Sleep (1000)
        'Set frm = KS_PrmPersonne
        'Call KS_PrmPersonne.gerer_nouvel_utilisateur(num_util)
        'Set frm = Nothing
    End If
    
    valider = P_OUI
    Exit Function

err_enreg:
    Call Odbc_RollbackTrans
    valider = P_ERREUR

End Function

Private Function valider_matricule() As Integer
' Verifier si le matricule renseigné n'existe pas dèjà dans la base

    Dim sql As String
    Dim rs As rdoResultset

    sql = "SELECT U_Num FROM Utilisateur" _
        & " WHERE U_kb_actif=True AND U_Matricule='" & txt(TXT_MATRICULE).Text & "'" _
        & " AND U_Num<>" & g_numutil
    If Odbc_SelectV(sql, rs) = P_ERREUR Then ' le matridule n'extiste pas
        valider_matricule = P_OK
        Exit Function
    End If
    If rs.EOF Then
        valider_matricule = P_OK
    Else
        valider_matricule = P_NON
    End If
    rs.Close

End Function

Private Sub valider_txt(ByVal v_row As Integer, ByVal v_col As Integer)
' *******************************************************
' Mettre le texte du TextBox dans la cellule du gridCoord
' Donner les coordonnées de la cellule suivante
' *******************************************************

    With grdCoord
        ' activer le bouton CMD_VALIDER
        If .TextMatrix(v_row, v_col) <> txt(TXT_CACHE).Text Then
            cmd(CMD_OK).Enabled = True
        End If
        .TextMatrix(.Row, .col) = txt(TXT_CACHE).Text
        ' On positionne le TextBox(TXT_CACHE) dans la bonne cellule (la suivante)
        Call deplacer_txt("DROITE", .Row, .col)

    End With

End Sub

Private Function verifier_champ(ByVal v_indtxt As Integer) As Integer

    Dim code As String, s As String

    Select Case v_indtxt
    Case TXT_CODE
        code = UCase(txt(v_indtxt).Text)
        If code = "ROOT" Then '       CODE RESERVÉ
            Call MsgBox("Code d'accès : " & code & " réservé.", vbOKOnly + vbExclamation, "")
            GoTo lab_non
        End If
        If Not P_ValiderCode(g_numutil, code, s) Then ' CODE DÉJÀ ATTRIBUÉ
            If p_traitement_background Then
                p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> Code d'accès : " & code & " est déjà attribué"
                If p_traitement_background_semiauto Then
                    MsgBox p_mess_fait_background
                End If
            Else
                p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> Code d'accès : " & code & " est déjà attribué"
                Call MsgBox("Code d'accès : " & code & " est déjà attribué.", vbOKOnly + vbExclamation, "")
            End If
            GoTo lab_non
        End If
    Case TXT_MATRICULE
        If txt(TXT_MATRICULE).Text <> "" Then
            ' Vérifier si le matricule n'existe pas dans la base
            If valider_matricule() = P_NON Then
                Call MsgBox("Le matricule " & txt(TXT_MATRICULE).Text _
                          & " fait référence à une autre personne dans le dictionnaire." & vbCrLf & vbCrLf _
                          & "Veuillez rectifier cette information.", _
                             vbExclamation + vbOKOnly, "Matricule redondant")
                GoTo lab_non
            End If
        End If
    End Select

lab_oui:
    verifier_champ = P_OUI
    Exit Function

lab_non:
    verifier_champ = P_NON

End Function

Private Function verifier_inactive() As Integer
' Vérifier si la personne qu'on veut désactiver n'est pas responsable d'une application

    Dim sql As String, u_nom As String, u_prenom As String, u_matricule As String

    sql = "SELECT U_Nom, U_Prenom, U_Matricule FROM Utilisateur WHERE U_kb_actif=True AND U_Num=" & g_numutil
    If Odbc_RecupVal(sql, u_nom, u_prenom, u_matricule) = P_ERREUR Then
        verifier_inactive = P_ERREUR
        Exit Function
    End If
    If P_RemplacerResponsableAppli(g_numutil, u_nom, u_prenom, u_matricule) = P_ERREUR Then
        verifier_inactive = P_ERREUR
        Exit Function
    End If

    verifier_inactive = P_OK

End Function

Private Function verifier_pobPrinc() As Integer
' ************************************************
' Vérifier qu'on a bien designé un poste principal
' ************************************************
    Dim I As Integer

    ' Vérifier si l'ancien poste principal existe toujour (même si on l'a supprimé ensuite on l'a ajouté !)
    For I = 1 To tvSect.Nodes.Count
        If left$(tvSect.Nodes(I).key, 1) = "P" Then
            If Mid$(tvSect.Nodes(I).key, 2) = g_pobPrinc Then
                GoTo lab_ok
            End If
        End If
    Next I
    ' On n'a pas de poste principal
    If tvSect.Nodes.Count = 2 Then ' 2 => un seul poste => principal
        g_pobPrinc = Mid$(tvSect.Nodes(tvSect.Nodes.Count).key, 2)
        GoTo lab_ok
    Else ' tvSect.Nodes.Count > 2
        Call MsgBox("Vous devez indiquer un poste principal pour " _
                    & txt(TXT_NOM).Text & " " & txt(TXT_PRENOM).Text & ".", vbOKOnly + vbExclamation, "Attention !")
        GoTo lab_non
    End If

lab_ok:
    verifier_pobPrinc = P_OK
    Exit Function

lab_non:
    verifier_pobPrinc = P_NON

End Function

Private Function verifier_tous_champs(ByRef r_numlabop As Long, _
                                      ByRef r_ssite As String, _
                                      ByRef r_sfct As String, _
                                      ByRef r_spm As Variant) As Integer

    Dim iligne As Integer
    Dim s As String

    With grdCoord ' vérifier la validité du NIVEAU de confidentialité et la VALEUR du coordonnée
        For iligne = 1 To .Rows - 1
            If .TextMatrix(iligne, GRDC_VALEUR) = "" Then ' la VALEUR du coordonnée est vide
                Call MsgBox("La VALEUR du coordonnée doit être renseignée.", vbCritical + vbOKOnly, "Erreur de validation")
                .col = GRDC_VALEUR
                GoTo lab_erreur
            End If
            If .TextMatrix(iligne, GRDC_NIVEAU) = "" Then ' Le NIVEAU
                Call MsgBox("Le NIVEAU de confidentialité est obligatoire." & "Il doit être compris entre 0 (niveau bas) et 9 (niveau haut).", _
                            vbCritical + vbOKOnly, "Erreur de validation")
                .col = GRDC_NIVEAU
                GoTo lab_erreur
            ElseIf Not STR_EstEntierPos(.TextMatrix(iligne, GRDC_NIVEAU)) Then ' NIVEAU de confidentialité est invalide
                Call MsgBox("Le NIVEAU de confidentialité doit être un entier positif.", vbCritical + vbOKOnly, "Erreur de validation")
                .TextMatrix(iligne, GRDC_NIVEAU) = ""
                .col = GRDC_NIVEAU
                GoTo lab_erreur
            ElseIf .TextMatrix(iligne, GRDC_NIVEAU) > 9 Then ' Le NIVEAU doit € [0; 9]
                Call MsgBox("Le NIVEAU de confidentialité doit être compris entre 0 (niveau bas) et 9 (niveau haut).", _
                            vbCritical + vbOKOnly, "Erreur de validation")
                .TextMatrix(iligne, GRDC_NIVEAU) = ""
                .col = GRDC_NIVEAU
                GoTo lab_erreur
            End If
        Next iligne
    End With
    ' ------- vérifier que le code n'est pas vide  --------
    If txt(TXT_CODE).Text = "" Then
        MsgBox "Le CODE de la personne est une rubrique obligatoire.", vbOKOnly + vbExclamation, ""
        sst.Tab = ONGLET_GENERAL
        txt(TXT_CODE).SetFocus
        GoTo lab_erreur
    Else ' ------- vérifier que le code est unique --------
        If Not P_ValiderCode(g_numutil, txt(TXT_CODE).Text, s) Then
            Call MsgBox("Le code entré: " & UCase$(txt(TXT_CODE).Text) & " est déjà utilisé. Veuillez en choisir un autre.", _
                        vbExclamation + vbOKOnly, "Attention")
            GoTo lab_erreur
        End If
    End If
    ' Vérifier si le matricule n'existe pas dans la base
    If verifier_champ(TXT_MATRICULE) = P_NON Then
        sst.Tab = ONGLET_GENERAL
        txt(TXT_CODE).SetFocus
        GoTo lab_erreur
    End If

    ' Construction U_SPM et U_FctTRav
    If tvSect.Nodes.Count = 0 Then
        Call MsgBox("Veuillez indiquer le poste affectée à la personne.", vbOKOnly + vbExclamation, "")
        sst.Tab = ONGLET_PROFIL
        tvSect.SetFocus
        GoTo lab_erreur
    End If
    ' Vérifier le poste principal
    If verifier_pobPrinc <> P_OK Then
        sst.Tab = ONGLET_PROFIL
        tvSect.SetFocus
        GoTo lab_erreur
    End If
    Call build_SPM_Fct(r_spm, r_sfct)

    ' Construction U_Labo
    r_ssite = ""
    r_numlabop = 0
    For iligne = 0 To grdLabo.Rows - 1
        If grdLabo.TextMatrix(iligne, GRDL_ESTLABO) = True Then
            r_ssite = r_ssite + "L" + grdLabo.TextMatrix(iligne, GRDL_NUMLABO) + ";"
        End If
        If grdLabo.TextMatrix(iligne, GRDL_LABOPRINC) <> "" Then
            r_numlabop = grdLabo.TextMatrix(iligne, GRDL_NUMLABO)
        End If
    Next iligne
    If r_numlabop = 0 Then
        MsgBox "Indiquez le site principal.", vbOKOnly + vbExclamation, ""
        grdLabo.SetFocus
        GoTo lab_erreur
    End If

    verifier_tous_champs = P_OUI
    Exit Function

lab_erreur:
    verifier_tous_champs = P_NON

End Function

Private Sub chk_Click(Index As Integer)

    If g_mode_saisie Then
        If Index = CHK_EXTERNE Then
            If chk(Index).Value = 1 Then
                chk(CHK_EXTERNE_IMPORT).Value = 1
                chk(CHK_EXTERNE_IMPORT).Enabled = False
            Else
                chk(CHK_EXTERNE_IMPORT).Enabled = True
            End If
        End If
        cmd(CMD_OK).Enabled = True
    End If

End Sub

Private Sub chk_GotFocus(Index As Integer)

    txt(TXT_CACHE).Visible = False

End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case CMD_HISTORIQUE
            Call afficher_historique
        Case CMD_PLUS_TYPE
            Call ajouter_type_coord
        Case CMD_MOINS_TYPE ' [s'il est active]
            Call supprimer_type_coord
            cmd(CMD_MOINS_TYPE).Enabled = False
        Case CMD_OK
            If valider() <> P_NON Then
                g_changements_importants = True ' pour rafraichir la page appelante
                Unload Me
                Exit Sub
            End If
        Case CMD_QUITTER
            Call quitter(False)
        Case CMD_ACCES_SPM
            Call prm_service
        Case CMD_MOINS_SPM
            Call supprimer_poste
        Case CMD_PLUS_COORDLIEES
            Call ajouter_coordLiee
        Case CMD_MOINS_COORDLIEES
            Call supprimer_coordLiee
            cmd(CMD_MOINS_COORDLIEES).Enabled = False
        Case CMD_PHOTO
            Call parcourir
        Case CMD_INFO_SUPPL ' type info suppl + valeurs
            Call type_info_suppl
        Case CMD_SUPPRIMER
            If supprimer_utilisateur() <> P_NON Then
                g_changements_importants = True ' pour rafraichir la page appelante
                Unload Me
                Exit Sub
            End If
    End Select

End Sub

Private Sub cmd_GotFocus(Index As Integer)

    txt(TXT_CACHE).Visible = False

End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim rep As Integer
    
    If Index = CMD_QUITTER Then
        g_mode_saisie = False
    ElseIf Index = CMD_PHOTO Then
        If Button = vbRightButton Then
            If cmd(CMD_PHOTO).Picture = imglst.ListImages(IMG_PHOTO_ON).Picture Then
                rep = MsgBox("Confirmez-vous la suppression du fichier ?", vbYesNo + vbQuestion, "")
                If rep = vbYes Then
                    g_nom_fich_photo = ""
                    Set cmd(CMD_PHOTO).Picture = imglst.ListImages(IMG_PHOTO_OFF).Picture
                    cmd(CMD_PHOTO).ToolTipText = "Pas de photo - Clic gauche pour choisir le fichier"
                    cmd(CMD_OK).Enabled = True
                End If
            End If
        End If
    End If

End Sub

Private Sub Form_Activate()

    If g_form_active Then Exit Sub

    g_form_active = True
    Call initialiser

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If (KeyCode = vbKeyE And Shift = vbAltMask) Or (KeyCode = vbKeyF1) Then
        KeyCode = 0
        If cmd(CMD_OK).Enabled Then
            If valider() <> P_NON Then
                Unload Me
                Exit Sub
            End If
        End If
    ElseIf (KeyCode = vbKeyH And Shift = vbAltMask) Then
        KeyCode = 0
        Call HtmlHelp(0, p_chemin_appli + "\help\kalidoc.chm", HH_DISPLAY_TOPIC, "dico_f_utilisateur.htm")
    ElseIf (KeyCode = vbKeyPageUp) Then
        Call afficher_page(0)
    ElseIf (KeyCode = vbKeyPageDown) Then
        Call afficher_page(1)
    ElseIf (KeyCode = vbKeyEscape) Then
        If txt(TXT_CACHE).Visible Or grdCoord.tag = "focus_oui" Then
            With grdCoord
                ' le grdCoord a le focus, alors on ne quitte pas
                txt(TXT_CACHE).Visible = False
                .tag = "focus_non"
            End With
            KeyCode = 0
        Else ' le grdCoord n'a pas le focus, on peut quitter
            KeyCode = 0
            Call quitter(False)
        End If

    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If grdCoord.tag = "focus_non" Then ' le grdCoord n'as pas le focus
            KeyAscii = 0
            SendKeys "{TAB}"
        End If
    End If

End Sub

Private Sub Form_Load()

    g_form_active = False

    g_form_width = Me.width
    g_form_height = Me.Height

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Call quitter(False)
    End If

End Sub

Private Sub grdcoord_Click()
    
    Dim mouse_row As Integer

    With grdCoord
        mouse_row = .MouseRow
        If mouse_row = 0 Then
               .Row = 0
            .col = 0
        End If
    End With

End Sub

Private Sub grdCoord_GotFocus()

    With grdCoord
        If .Rows - 1 = 0 Then Exit Sub
        .tag = "focus_oui"
        cmd(CMD_MOINS_TYPE).Enabled = True
        If .Row = 0 Then Exit Sub
        If .col = GRDC_PRINCIPAL Then
            txt(TXT_CACHE).Visible = False
        End If
    End With

End Sub

Private Sub grdCoord_KeyPress(KeyAscii As Integer)
    Dim frm As Form
    Dim ucnum As Long
    With grdCoord
        Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
            Call deplacer_txt("DROITE", .Row, .col)
        Case vbKeySpace
            KeyAscii = 0
            If .col = GRDC_PRINCIPAL Then
                Call basculer_colonne_principal(.Row, .col)
            ElseIf .col = GRDC_COMMENTAIRE Then
                If .TextMatrix(.Row, GRDC_UCNUM) <> "" Then
                    ucnum = .TextMatrix(.Row, GRDC_UCNUM)
                Else
                    ucnum = 0
                End If
                Set frm = SaisieCommentaire
                If SaisieCommentaire.AppelFrm(grdCoord, ucnum) Then
                    cmd(CMD_OK).Enabled = True
                End If
                Set frm = Nothing
            Else
                g_position_txt_cache = .Row & ";" & .col
                If .col = GRDC_NIVEAU Then ' COLCOLCOL
                    txt(TXT_CACHE).MaxLength = 1
                Else
                    txt(TXT_CACHE).MaxLength = 0
                End If
                Call positionner_txt(.CellLeft + .left, .CellTop + .Top, .CellWidth, .CellHeight, .TextMatrix(.Row, .col))
            End If
        Case vbKeyEscape
            KeyAscii = 0
            If .col = GRDC_COMMENTAIRE Then
            End If
'        Case Else
'            KeyAscii = 0
'            If .col = GRDC_PRINCIPAL Then
'                g_position_txt_cache = ""
'                Exit Sub
'            ElseIf .col = GRDC_COMMENTAIRE Then
'                Set frm = SaisieCommentaire
'                if SaisieCommentaire.AppelFrm then
'                    cmd(CMD_OK).Enabled = True
'                End If
'                Set frm = Nothing
'                Exit Sub
'            End If
' =================================================
'        Case vbKeyRight
'            KeyAscii = 0
'            Call deplacer_txt("DROITE", .Row, .col)
'        Case vbKeyLeft
'            KeyAscii = 0
'            Call deplacer_txt("GAUCHE", .Row, .col)
' =================================================
        End Select
    End With

End Sub

Private Sub grdCoord_LostFocus()

    grdCoord.tag = "focus_non"
    cmd(CMD_MOINS_TYPE).Enabled = True

End Sub

Private Sub grdCoord_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim old_row As Integer, old_col As Integer, mouse_row As Integer, mouse_col As Integer
    Dim frm As Form
    Dim ucnum As Long
    If Button = MouseButtonConstants.vbRightButton Then Exit Sub
    With grdCoord
        mouse_row = .MouseRow
        mouse_col = .MouseCol
        If .TextMatrix(mouse_row, GRDC_UCNUM) <> "" Then
            ucnum = .TextMatrix(mouse_row, GRDC_UCNUM)
        Else
            ucnum = 0
        End If
        .tag = "focus_oui"
        Select Case mouse_col
            ' -------------------------------------------------------
            Case GRDC_PRINCIPAL
                Call basculer_colonne_principal(mouse_row, mouse_col)
                txt(TXT_CACHE).Visible = False
                g_position_txt_cache = mouse_row & ";" & mouse_col
            ' -------------------------------------------------------
            Case Else
                If mouse_row > 0 And mouse_row < .Rows Then
                    If mouse_col > 1 Then
                        If mouse_col = GRDC_COMMENTAIRE Then
                            Set frm = SaisieCommentaire
                            If SaisieCommentaire.AppelFrm(grdCoord, ucnum) Then
                                cmd(CMD_OK).Enabled = True
                            End If
                            Set frm = Nothing
                            Exit Sub
                        Else ' ARRET
                            If mouse_col = GRDC_NIVEAU Then ' COLCOLCOL
                                txt(TXT_CACHE).MaxLength = 1
                            Else
                                txt(TXT_CACHE).MaxLength = 0
                            End If
                            Call positionner_txt(.CellLeft + .left, .CellTop + .Top, .CellWidth, _
                                                 .CellHeight, .TextMatrix(mouse_row, mouse_col))
                        End If
                    End If
                Else
                    txt(TXT_CACHE).Visible = False
                End If
        End Select
    End With

End Sub

Private Sub grdCoord_RowColChange()

    With grdCoord
        If .Row = 0 Then Exit Sub
        If .col < GRDC_VALEUR Then
            .col = GRDC_VALEUR
        End If
        If .col = GRDC_VALEUR Or .col = GRDC_NIVEAU Then
            If .col = GRDC_NIVEAU Then ' COLCOLCOL
                txt(TXT_CACHE).MaxLength = 1
            Else
                txt(TXT_CACHE).MaxLength = 0
            End If
            Call positionner_txt(.CellLeft + .left, .CellTop + .Top, .CellWidth, .CellHeight, .TextMatrix(.Row, .col))
        End If
    End With

End Sub

Private Sub grdCoord_Scroll()

    txt(TXT_CACHE).Visible = False

End Sub

Private Sub grdCoordLiees_MouseUp_(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim old_row As Integer, old_col As Integer, mouse_row As Integer, mouse_col As Integer
    Dim frm As Form
    Dim ucnum As Long
    
Exit Sub
    If Button = MouseButtonConstants.vbRightButton Then Exit Sub
    With grdCoordLiees
        mouse_row = .MouseRow
        mouse_col = .MouseCol
        If .TextMatrix(mouse_row, GRDC_UCNUM) <> "" Then
            ucnum = .TextMatrix(mouse_row, GRDC_UCNUM)
        Else
            ucnum = 0
        End If
        Select Case mouse_col
            ' -------------------------------------------------------
            Case GRDCL_PRINCIPAL
                Call basculer_colonne_principal(mouse_row, mouse_col)
                txt(TXT_CACHE).Visible = False
                g_position_txt_cache = mouse_row & ";" & mouse_col
            ' -------------------------------------------------------
            Case Else
                If mouse_row > 0 And mouse_row < .Rows Then
                    If mouse_col > 1 Then
                        If mouse_col = GRDC_COMMENTAIRE Then
                            Set frm = SaisieCommentaire
                            If SaisieCommentaire.AppelFrm(grdCoord, ucnum) Then
                                cmd(CMD_OK).Enabled = True
                            End If
                            Set frm = Nothing
                            Exit Sub
                        Else ' ARRET
                            If mouse_col = GRDC_NIVEAU Then ' COLCOLCOL
                                txt(TXT_CACHE).MaxLength = 1
                            Else
                                txt(TXT_CACHE).MaxLength = 0
                            End If
                            Call positionner_txt(.CellLeft + .left, .CellTop + .Top, .CellWidth, _
                                                 .CellHeight, .TextMatrix(mouse_row, mouse_col))
                        End If
                    End If
                Else
                    txt(TXT_CACHE).Visible = False
                End If
        End Select
    End With

End Sub

Private Sub grdCoordLiees_Click()
'MsgBox "click"
End Sub

Private Sub grdCoordLiees_GotFocus()

    With grdCoordLiees
        If .Rows = 0 Then Exit Sub
'MsgBox "gotfocus"
        cmd(CMD_MOINS_COORDLIEES).Enabled = True
    End With

End Sub

Private Sub grdCoordLiees_KeyPress(KeyAscii As Integer)

    With grdCoordLiees
        Select Case KeyAscii
            Case vbKeySpace
                KeyAscii = 0
                If .col = GRDCL_PRINCIPAL Then
                    If .CellPicture = imglst.ListImages(IMG_COCHE).Picture Then
                        Set .CellPicture = imglst.ListImages(IMG_PASCOCHE).Picture
                    Else
                        Set .CellPicture = imglst.ListImages(IMG_COCHE).Picture
                    End If
                End If
                cmd(CMD_OK).Enabled = True
        End Select
    End With

End Sub

Private Sub grdCoordLiees_LostFocus()
'MsgBox "lostfocus"
    cmd(CMD_MOINS_COORDLIEES).Enabled = True

End Sub

Private Sub grdCoordLiees_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then Exit Sub

    With grdCoordLiees
        Select Case .MouseCol
           Case GRDCL_PRINCIPAL
                If .CellPicture = imglst.ListImages(IMG_COCHE).Picture Then
                    Set .CellPicture = imglst.ListImages(IMG_PASCOCHE).Picture
                Else
                    Set .CellPicture = imglst.ListImages(IMG_COCHE).Picture
                End If
                cmd(CMD_OK).Enabled = True
        End Select
    End With

End Sub

Private Sub grdLabo_DblClick()

    If grdLabo.col = GRDL_IMG_ESTLABO Then
        Call inverser_etat_labo
    ElseIf grdLabo.col = GRDL_LABOPRINC Then
        Call inverser_laboprinc
    End If

End Sub

Private Sub grdLabo_GotFocus()

    If Not g_mode_saisie Then
        Exit Sub
    End If

    grdLabo.col = GRDL_CODLABO
    grdLabo.ColSel = GRDL_CODLABO

End Sub

Private Sub grdLabo_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeySpace Then
        KeyCode = 0
        If grdLabo.col = GRDL_IMG_ESTLABO Then
            Call inverser_etat_labo
        ElseIf grdLabo.col = GRDL_LABOPRINC Then
            Call inverser_laboprinc
        End If
    End If

End Sub

Private Sub grdLabo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub mnuPosteAttribuerNum_Click()
    
    Call attribuer_coordonnee
    
End Sub

Private Sub mnuPostePrinc_Click()

    Call basculer_postePrincipal(tvSect.SelectedItem.Index)

End Sub

'Private Sub pctPhoto_Click()

'    pctPhoto.Visible = False
    
'End Sub

Private Sub sst_Click(PreviousTab As Integer)

    If Not g_mode_saisie Then
        Exit Sub
    End If

    Call init_focus

End Sub

Private Sub tvSect_BeforeLabelEdit(Cancel As Integer)

    Cancel = True

End Sub

Private Sub tvSect_Click()

    Dim nd As Node

    If tvSect.Nodes.Count = 0 Then Exit Sub ' pour ne pas provoquer d'erreurs
    Set nd = tvSect.Nodes(tvSect.SelectedItem.Index)
    If g_button = MouseButtonConstants.vbRightButton Then
        ' Gérer le poste principal
        If left$(nd.key, 1) = "P" Then
            Call afficher_menu(nd.image <> IMGT_POSTE_SEL)
        End If
    End If

End Sub


Private Sub tvSect_GotFocus()

    If Not g_mode_saisie Then
        Exit Sub
    End If

    If tvSect.Nodes.Count > 0 Then
        Set tvSect.SelectedItem = tvSect.Nodes(1)
    Else
        tvSect.BorderStyle = ccFixedSingle
    End If

End Sub

Private Sub tvSect_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        KeyCode = 0
        Call prm_service
    ElseIf KeyCode = vbKeyDelete Then
        KeyCode = 0
        Call supprimer_poste
    End If

End Sub

Private Sub tvSect_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        sst.Tab = 2
    End If

End Sub

Private Sub tvSect_LostFocus()

    If tvSect.Nodes.Count = 0 Then
        tvSect.BorderStyle = ccNone
    End If

End Sub

Private Sub tvSect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    g_button = Button

End Sub

Private Sub txt_Change(Index As Integer)

    cmd(CMD_OK).Enabled = True

End Sub

Private Sub txt_GotFocus(Index As Integer)

    If Not g_mode_saisie Then
        Exit Sub
    End If

    If Index <> TXT_CACHE Then
        txt(TXT_CACHE).Visible = False
        grdCoord.tag = "focus_non"
        cmd(CMD_MOINS_TYPE).Enabled = False
    Else
        grdCoord.tag = "focus_oui"
        cmd(CMD_MOINS_TYPE).Enabled = True
    End If

    g_txt_avant = txt(Index).Text

End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    Dim I As Integer

    If Index = TXT_CACHE Then
        With grdCoord
            Select Case KeyCode
                Case vbKeyUp
                    KeyCode = 0
                    Call deplacer_txt("HAUT", .Row, .col)
                Case vbKeyDown
                    KeyCode = 0
                    Call deplacer_txt("BAS", .Row, .col)
                Case vbKeyRight
                    If txt(Index).SelLength > 0 Or Len(txt(Index).Text) = 0 Then
                        KeyCode = 0
                        Call deplacer_txt("DROITE", .Row, .col)
                    End If
                Case vbKeyLeft
                    If txt(Index).SelLength > 0 Or Len(txt(Index).Text) = 0 Then
                        KeyCode = 0
                        Call deplacer_txt("GAUCHE", .Row, .col)
                    End If
            End Select
        End With
    End If

End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)

    If Index = TXT_CACHE Then
        With grdCoord
            Select Case KeyAscii
            Case vbKeyReturn
                KeyAscii = 0
                Call valider_txt(.Row, .col)
                g_position_txt_cache = .Row & ";" & .col
            Case vbKeyEscape
                KeyAscii = 0
                If txt(TXT_CACHE).Visible Then
                    txt(TXT_CACHE).Text = ""
                    txt(TXT_CACHE).Visible = False
                    .tag = "focus_non"
                End If
            End Select
        End With
    End If
    
End Sub

Private Sub txt_LostFocus(Index As Integer)

    Dim I As Integer

    If Index = TXT_CACHE Then
            grdCoord.tag = "focus_non"
            txt(TXT_CACHE).Visible = False
    End If

    If g_mode_saisie Then
        If txt(Index).Text <> g_txt_avant Then
            If verifier_champ(Index) = P_NON Then
                txt(Index).Text = ""
                txt(Index).SetFocus
                Exit Sub
            End If
            cmd(CMD_OK).Enabled = True
        End If
    End If

End Sub
