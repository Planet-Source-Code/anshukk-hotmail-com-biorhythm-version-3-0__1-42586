VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "Splash.frx":5312
   ScaleHeight     =   4950
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Jai Shree Ram"
      Height          =   255
      Left            =   6120
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Programmer Anshuk Kumar      Contact: Anshukk@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   4200
      Width           =   6615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 3.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   0
      Top             =   3600
      Width           =   2895
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   240
      Picture         =   "Splash.frx":77B1C
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   4980
      Left            =   0
      Picture         =   "Splash.frx":7F2BC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7380
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Unload Me
Load MainFrm
MainFrm.Show
End Sub
