VERSION 5.00
Begin VB.Form AnalyseFichier 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11895
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PréImport"
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
      Height          =   7695
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11895
   End
   Begin VB.Frame frm 
      BackColor       =   &H00808080&
      ForeColor       =   &H00000080&
      Height          =   855
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   7580
      Width           =   11895
   End
   Begin VB.Menu mnuMenuContextuel 
      Caption         =   "menu contextuel"
      Visible         =   0   'False
      Begin VB.Menu mnuActualiser 
         Caption         =   "Actualiser le tableau"
      End
      Begin VB.Menu mnuActionP 
         Caption         =   "Associer"
         Begin VB.Menu mnuAction 
            Caption         =   "une action à faire"
         End
         Begin VB.Menu mnurechercher 
            Caption         =   "Rechercher dans KaliWeb"
         End
      End
   End
End
Attribute VB_Name = "AnalyseFichier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

