VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PrmAppTypeInfoSuppl 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Type d'informations supplémentaires"
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
      Height          =   3615
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Command2"
         Height          =   375
         Index           =   3
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2880
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Command1"
         Height          =   375
         Index           =   2
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   2775
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   4895
         _Version        =   393216
      End
   End
   Begin VB.Frame frmFct 
      BackColor       =   &H00808080&
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   -240
      TabIndex        =   0
      Top             =   3480
      Width           =   6045
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
         Left            =   4560
         Picture         =   "PrmAppTypeInfoSuppl.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   230
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "PrmAppTypeInfoSuppl.frx":05B9
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
         Left            =   960
         Picture         =   "PrmAppTypeInfoSuppl.frx":0B15
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   230
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "PrmAppTypeInfoSuppl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

