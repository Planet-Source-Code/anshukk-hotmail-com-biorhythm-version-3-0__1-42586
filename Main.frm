VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Techno Rex BioRhythm Version 3.0"
   ClientHeight    =   7755
   ClientLeft      =   -885
   ClientTop       =   330
   ClientWidth     =   7920
   ForeColor       =   &H00000000&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   7920
   Begin VB.Frame Frame2 
      Caption         =   "Add/ Remove Names"
      Height          =   7755
      Left            =   7920
      TabIndex        =   49
      Top             =   120
      Width           =   3015
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   56
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         Left            =   120
         TabIndex        =   55
         Top             =   2040
         Width           =   885
      End
      Begin VB.ComboBox Combo9 
         Height          =   315
         Left            =   1080
         TabIndex        =   54
         Top             =   2040
         Width           =   885
      End
      Begin VB.ComboBox Combo10 
         Height          =   315
         Left            =   2040
         TabIndex        =   53
         Top             =   2040
         Width           =   885
      End
      Begin VB.ListBox List1 
         Height          =   3765
         ItemData        =   "Main.frx":030A
         Left            =   240
         List            =   "Main.frx":030C
         TabIndex        =   52
         Top             =   3480
         Width           =   2655
      End
      Begin VB.CommandButton cmdAddName 
         Caption         =   "Add Name"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdRemoveName 
         Caption         =   "Remove Name"
         Height          =   255
         Left            =   1560
         TabIndex        =   50
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label32 
         Caption         =   "Default"
         Height          =   255
         Left            =   360
         TabIndex        =   57
         Top             =   7320
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label26 
         Caption         =   "Date Of Birth:"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label27 
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label28 
         Caption         =   "Month"
         Height          =   255
         Left            =   1080
         TabIndex        =   61
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Year"
         Height          =   255
         Left            =   2040
         TabIndex        =   60
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label31 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Niagara Solid"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   59
         Top             =   3400
         Width           =   255
      End
      Begin VB.Label Label31 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Niagara Solid"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   58
         Top             =   7200
         Width           =   255
      End
   End
   Begin VB.ComboBox Combo11 
      Height          =   315
      Left            =   1680
      TabIndex        =   47
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdAddRemoveShow 
      Caption         =   "Add /Remove Names      >>"
      Height          =   375
      Left            =   5520
      TabIndex        =   46
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Understanding Biorhythm"
      Height          =   7755
      Left            =   7920
      TabIndex        =   39
      Top             =   120
      Width           =   3015
      Begin VB.TextBox Text1 
         Height          =   6135
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         Text            =   "Main.frx":030E
         Top             =   1200
         Width           =   2775
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   120
         TabIndex        =   41
         Text            =   "What is Biorhythm?"
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label23 
         Caption         =   "Select Question:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdShow 
      Appearance      =   0  'Flat
      Caption         =   "Understanding Biorhythm >>"
      Height          =   375
      Left            =   5520
      TabIndex        =   38
      Top             =   6480
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   0
      ScaleHeight     =   4665
      ScaleWidth      =   7785
      TabIndex        =   37
      Top             =   1440
      Width           =   7815
      Begin VB.Label Label24 
         BackColor       =   &H00000000&
         Caption         =   "0%"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   7440
         TabIndex        =   45
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "-100%"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   7320
         TabIndex        =   44
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "+100%"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   7320
         TabIndex        =   43
         Top             =   0
         Width           =   615
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   30
         X1              =   7440
         X2              =   7440
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   7800
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   0
         X1              =   240
         X2              =   240
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   1
         X1              =   480
         X2              =   480
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   2
         X1              =   720
         X2              =   720
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   3
         X1              =   960
         X2              =   960
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   4
         X1              =   1200
         X2              =   1200
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   5
         X1              =   1440
         X2              =   1440
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   6
         X1              =   1680
         X2              =   1680
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   7
         X1              =   1920
         X2              =   1920
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   8
         X1              =   2160
         X2              =   2160
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   9
         X1              =   2400
         X2              =   2400
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   10
         X1              =   2640
         X2              =   2640
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   11
         X1              =   2880
         X2              =   2880
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   12
         X1              =   3120
         X2              =   3120
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   13
         X1              =   3360
         X2              =   3360
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   14
         X1              =   3600
         X2              =   3600
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   15
         X1              =   3840
         X2              =   3840
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   16
         X1              =   4080
         X2              =   4080
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   17
         X1              =   4320
         X2              =   4320
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   18
         X1              =   4560
         X2              =   4560
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   19
         X1              =   4800
         X2              =   4800
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   20
         X1              =   5040
         X2              =   5040
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   21
         X1              =   5280
         X2              =   5280
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   22
         X1              =   5520
         X2              =   5520
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   23
         X1              =   5760
         X2              =   5760
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   24
         X1              =   6000
         X2              =   6000
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   25
         X1              =   6240
         X2              =   6240
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   26
         X1              =   6480
         X2              =   6480
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   27
         X1              =   6720
         X2              =   6720
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   28
         X1              =   6960
         X2              =   6960
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   29
         X1              =   7200
         X2              =   7200
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808000&
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   4335
         Left            =   3840
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdDivineCycles 
      Caption         =   "Divine Cycles"
      Height          =   375
      Left            =   3480
      TabIndex        =   29
      Top             =   6240
      Width           =   1635
   End
   Begin VB.CommandButton cmdSecondaryCycles 
      Caption         =   "Secondary Cycles"
      Height          =   375
      Left            =   1800
      TabIndex        =   30
      Top             =   6240
      Width           =   1635
   End
   Begin VB.CommandButton cmdPrimaryCycles 
      Caption         =   "Primary Cycles"
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   6240
      Width           =   1635
   End
   Begin VB.CommandButton cmdTodayDate 
      Caption         =   "<< Today"
      Height          =   315
      Left            =   4560
      TabIndex        =   32
      Top             =   1080
      Width           =   975
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Text            =   "1"
      Top             =   1080
      Width           =   885
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   3600
      TabIndex        =   4
      Text            =   "2002"
      Top             =   1080
      Width           =   885
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   2640
      TabIndex        =   1
      Text            =   "1"
      Top             =   1080
      Width           =   885
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   3600
      TabIndex        =   0
      Text            =   "1981"
      Top             =   720
      Width           =   885
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Main.frx":04BF
      Left            =   2640
      List            =   "Main.frx":04C1
      TabIndex        =   2
      Text            =   "11"
      Top             =   720
      Width           =   885
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Main.frx":04C3
      Left            =   1680
      List            =   "Main.frx":04C5
      TabIndex        =   3
      Text            =   "7"
      Top             =   720
      Width           =   885
   End
   Begin VB.Label Label30 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label22 
      Caption         =   "Spiritual Cycle"
      Height          =   255
      Left            =   3840
      TabIndex        =   36
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   35
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label21 
      Caption         =   "Date and Time"
      Height          =   255
      Left            =   5640
      TabIndex        =   34
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label20 
      Height          =   255
      Left            =   5640
      TabIndex        =   33
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label19 
      Caption         =   "Self Awareness Cycle"
      Height          =   255
      Left            =   3840
      TabIndex        =   28
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   3480
      TabIndex        =   27
      Top             =   7440
      Width           =   255
   End
   Begin VB.Label Label17 
      Caption         =   "Aesthetic  Cycle"
      Height          =   255
      Left            =   3840
      TabIndex        =   26
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label16 
      BackColor       =   &H00808000&
      Height          =   255
      Left            =   3480
      TabIndex        =   25
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label15 
      Caption         =   "Intutive Cycle"
      Height          =   255
      Left            =   3840
      TabIndex        =   24
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Left            =   3480
      TabIndex        =   23
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label13 
      Caption         =   "Year"
      Height          =   255
      Left            =   3600
      TabIndex        =   22
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Month"
      Height          =   255
      Left            =   2640
      TabIndex        =   21
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "Date"
      Height          =   255
      Left            =   1680
      TabIndex        =   20
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Mastery Cycle"
      Height          =   255
      Left            =   2160
      TabIndex        =   19
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF00&
      Height          =   255
      Left            =   1800
      TabIndex        =   18
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "Wisdom Cycle"
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF00FF&
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "Passion Cycle"
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Intellectual Cycle"
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   13
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Emotional Cycle"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   11
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Physical Cycle"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   9
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Date Of Birth:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Condition On Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim days As Integer
Dim spiritualCycleDays, physicalCycleDays As Integer, emotionalCycleDays As Integer, intellectualCycleDays As Integer, intutiveCycleDays As Integer, aestheticCycleDays As Integer, selfAwarenessCycleDays As Integer
Dim lastButtonPressed As Integer
Dim TopY As Double
Dim dateOfBirth As Date, dateOfCondition As Date
Dim bShow As Single

Private Sub cmdAddName_Click()
Dim CntY As Integer
Dim CntX As Integer
Dim strName As String
'CHECK VALIDITY OF DATE
If Not (IsDate(Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text) And IsDate(Combo4.Text & "/" & Combo5.Text & "/" & Combo6.Text)) Then
    MsgBox "The Entered date is not valid."
    Exit Sub
End If
'CHECKS FOR BLANK NAMES AND ADDS NAME TO LIST
If LTrim(RTrim(Text2.Text)) <> "" Then
List1.AddItem Text2.Text & "|" & Combo8.Text & "/" & Combo9.Text & "/" & Combo10.Text
'REFRESH THE NAME COMBOBOX
Combo11.Clear
For CntX = 0 To List1.ListCount - 1
    For CntY = 1 To Len(List1.list(CntX))
        If Not ("|" = Mid((List1.list(CntX)), CntY, 1)) Then
                strName = strName & Mid((List1.list(CntX)), CntY, 1)
        Else
                Combo11.AddItem strName
                strName = ""
                Exit For
                
        End If
    Next
Next
End If
End Sub

Private Sub cmdAddRemoveShow_Click()
'SHOWS AND HIDES THE ADD/ REMOVE NAME FRAME
cmdShow.Enabled = True
If bShow = 0 Then
    cmdAddRemoveShow.Caption = "Add /Remove Names      <<"
        If cmdShow.Caption = "Understanding Biorhythm >>" Then
            Width = 11160
            bShow = 2
        Else
            bShow = 1
        End If
ElseIf bShow = 1 Then
    cmdAddRemoveShow.Caption = "Add /Remove Names      >>"
        If cmdShow.Caption = "Understanding Biorhythm <<" Then
            Width = 11160
            bShow = 2
        Else
            Width = 8010
            bShow = 0
        End If
Else
          If cmdAddRemoveShow.Caption = "Add /Remove Names      <<" Then
          cmdAddRemoveShow.Caption = "Add /Remove Names      >>"
          Width = 8010
          bShow = 0
          Else
          cmdAddRemoveShow.Caption = "Add /Remove Names      <<"
          Width = 11160
          bShow = 1
          End If

End If

Call OptimizeFormPosition

If cmdAddRemoveShow.Caption = "Add /Remove Names      <<" Then
Frame2.Visible = True
Else
Frame2.Visible = False
End If

If cmdAddRemoveShow.Caption = "Add /Remove Names      <<" And cmdShow.Caption = "Understanding Biorhythm <<" Then
cmdShow.Enabled = False
End If

End Sub

Private Sub cmdRemoveName_Click()
Dim CntY As Integer
Dim CntX As Integer

'REMOVES NAME FROM LIST
If List1.ListIndex = -1 Then
MsgBox "Please Select Name to delete."
Else
List1.RemoveItem List1.ListIndex
End If

'REFRESH NAMES IN COMBOBOX
Combo11.Clear
For CntX = 0 To List1.ListCount - 1
    For CntY = 1 To Len(List1.list(CntX))
        If Not ("|" = Mid((List1.list(CntX)), CntY, 1)) Then
                strName = strName & Mid((List1.list(CntX)), CntY, 1)
        Else
                Combo11.AddItem strName
                strName = ""
                Exit For
                
        End If
    Next
Next
End Sub

Private Sub cmdShow_Click()
'SHOWS AND HIDES THE UNDERSTANDING BIORHYTHM FRAME
cmdAddRemoveShow.Enabled = True


If bShow = 0 Then
    cmdShow.Caption = "Understanding Biorhythm <<"
        If cmdAddRemoveShow.Caption = "Add /Remove Names      >>" Then
            Width = 11160
            bShow = 2
        Else
            bShow = 1
        End If
ElseIf bShow = 1 Then
        cmdShow.Caption = "Understanding Biorhythm >>"
        If cmdAddRemoveShow.Caption = "Add /Remove Names      <<" Then
                Width = 11160
                bShow = 2
            Else
                Width = 8010
                bShow = 0
            End If
Else
          If cmdShow.Caption = "Understanding Biorhythm >>" Then
          cmdShow.Caption = "Understanding Biorhythm <<"
          Width = 11160
          bShow = 1
          Else
          cmdShow.Caption = "Understanding Biorhythm >>"
          Width = 8010
          bShow = 0
          End If

End If
Call OptimizeFormPosition

If cmdShow.Caption = "Understanding Biorhythm >>" Then
Frame2.Visible = True
Else
Frame2.Visible = False
End If
If cmdAddRemoveShow.Caption = "Add /Remove Names      <<" And cmdShow.Caption = "Understanding Biorhythm <<" Then
cmdAddRemoveShow.Enabled = False
End If
End Sub





Private Sub Combo1_Click()
Form_Paint
End Sub
Private Sub Combo2_Click()
Form_Paint
End Sub
Private Sub Combo3_Click()
Form_Paint
End Sub
Private Sub Combo4_Click()
Form_Paint
End Sub
Private Sub Combo5_Click()
Form_Paint
End Sub
Private Sub Combo6_Click()
Form_Paint
End Sub


Private Sub Combo11_Click()
Dim strDate As String, strName As String
Dim syntaxFlag As Boolean, DoNotEnter As Boolean
Dim CntX As Integer, CntY As Integer

For CntX = 0 To List1.ListCount - 1
syntaxFlag = True
DoNotEnter = False

strName = ""
    For CntY = 1 To Len(List1.list(CntX))
        If "|" <> Mid((List1.list(CntX)), CntY, 1) And Not DoNotEnter Then
                strName = strName & Mid((List1.list(CntX)), CntY, 1)
        Else
                DoNotEnter = True
                If Combo11.Text = strName Then
                    If ("|" = Mid((List1.list(CntX)), CntY, 1)) Then
                           CntY = CntY + 1
                           syntaxFlag = False
                            
                    End If
                    If Not (syntaxFlag) Then
                            strDate = strDate & Mid((List1.list(CntX)), CntY, 1)
                    End If
                Else
                Exit For
                End If
                If Len(List1.list(CntX)) = CntY Then
                Combo1.Text = Day(CDate(strDate))
                Combo2.Text = Month(CDate(strDate))
                Combo3.Text = Year(CDate(strDate))
                End If
        End If

    Next
    
Next

Form_Paint

End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
MsgBox "No entery allowed"
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
MsgBox "No entery allowed"
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
MsgBox "No entery allowed"
End Sub
Private Sub Combo3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
MsgBox "No entery allowed"
End Sub
Private Sub Combo5_KeyPress(KeyAscii As Integer)
KeyAscii = 0
MsgBox "No entery allowed"
End Sub
Private Sub Combo6_KeyPress(KeyAscii As Integer)
KeyAscii = 0
MsgBox "No entery allowed"
End Sub

Private Sub Combo8_KeyPress(KeyAscii As Integer)
KeyAscii = 0
MsgBox "No entery allowed"
End Sub
Private Sub Combo9_KeyPress(KeyAscii As Integer)
KeyAscii = 0
MsgBox "No entery allowed"
End Sub
Private Sub Combo10_KeyPress(KeyAscii As Integer)
KeyAscii = 0
MsgBox "No entery allowed"
End Sub
Private Sub cmdPrimaryCycles_Click()
Dim ndays As Double
Dim b, r, g As Integer
Dim l, m, n As Double


'CHECKS FOR VALIDITY OF DATE

If Not (IsDate(Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text) And IsDate(Combo4.Text & "/" & Combo5.Text & "/" & Combo6.Text)) Then
    MsgBox "The Entered date is not valid."
    Exit Sub
End If

Picture1.Cls
Label20.Visible = True
lastButtonPressed = 1


dateOfBirth = Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text
dateOfCondition = Combo4.Text & "/" & Combo5.Text & "/" & Combo6.Text

'FINDS THE AGE OF PERSON IN DAYS ON CONDITION DATE

ndays = DateDiff("d", dateOfBirth, dateOfCondition)

'FINDS THE EXACT POSITION OF EACH CYCLE
l = (ndays / physicalCycleDays - Int(ndays / physicalCycleDays)) * physicalCycleDays
m = (ndays / emotionalCycleDays - Int(ndays / emotionalCycleDays)) * emotionalCycleDays
n = (ndays / intellectualCycleDays - Int(ndays / intellectualCycleDays)) * intellectualCycleDays
 
b = (16 - l) * days
r = (16 - m) * days
g = (16 - n) * days

'DRAWS WAVES FROM CENTER TO FORWARD
Call MakeCurve(days + b, TopY, (days * physicalCycleDays) + b, RGB(0, 0, 255))
Call MakeCurve(days + r, TopY, (days * emotionalCycleDays) + r, RGB(255, 0, 0))
Call MakeCurve(days + g, TopY, (days * intellectualCycleDays) + g, RGB(0, 255, 0))

'DRAWS WAVES FROM CENTER TO BACKWARD
Call MakeReverseCurve(days + b, TopY, (days * physicalCycleDays) + b, RGB(0, 0, 255))
Call MakeReverseCurve(days + r, TopY, (days * emotionalCycleDays) + r, RGB(255, 0, 0))
Call MakeReverseCurve(days + g, TopY, (days * intellectualCycleDays) + g, RGB(0, 255, 0))


End Sub
'MAKE CURVES FORWARDS

 Sub MakeCurve(X As Double, Y As Double, d As Double, c As Long)
 Dim T
 Picture1.ForeColor = c
 Picture1.PSet (X, Y)
 For T = X To days * 31
    Picture1.Line -(T, Y - (2160 * Sin((3.14) * ((T - X) / ((d - X) / 2)))))
 Next

End Sub
'MAKES CURVES BACKWARDS
Sub MakeReverseCurve(X As Double, Y As Double, d As Double, c As Long)
Dim T
Picture1.ForeColor = c
Picture1.PSet (days - d + X, Y)
For T = (days - d + X) To X
    Picture1.Line -(T, Y - (2160 * Sin((3.14) * ((T - X) / ((d - X) / 2)))))
Next
End Sub

'MAKES SECONDARY CURVE
Sub MakeSecCurve(x1 As Double, y1 As Double, d1 As Double, c1 As Long, x2 As Double, y2 As Double, d2 As Double)
Dim T, X
 Picture1.ForeColor = c1
 Picture1.PSet (0, ((y1 - (2160 * Sin((3.14) * ((T - x1) / ((d1 - x1) / 2))))) + (y2 - (2160 * Sin((3.14) * ((T - x2) / ((d2 - x2) / 2)))))) / 2)
 For T = X To days * 31
    Picture1.Line -(T, ((y1 - (2160 * Sin((3.14) * ((T - x1) / ((d1 - x1) / 2))))) + (y2 - (2160 * Sin((3.14) * ((T - x2) / ((d2 - x2) / 2)))))) / 2)
 Next
End Sub


' CHECKS WHETHER YEAR IS LEAP YEAR OR NOT
Function bIsLeapYear(ByVal inYear As Integer) As Boolean
    bIsLeapYear = ((inYear Mod 4 = 0) _
               And (inYear Mod 100 <> 0) _
                Or (inYear Mod 400 = 0))
End Function



Private Sub Combo7_click()
'PUTS HELP TEXT IN TEXT BOX  WHEN COMBO7 IS CLICKED
Select Case Combo7.Text
Case "What is Biorhythm?"
Text1.Text = "Biorhythm study and use is considered a ""pseudo science"" in the United State however it is widely accepted and utilized throughout Europe and much of the rest of the world.              " & _
             "Biorhythms are inherent cycles which regulate your metabolism, coordination, emotions, memory, and more.  As your biorhythm cycles rise and fall, so does your ability to perform physical activities, deal with stress, and make sound decisions."
Case "Meaning of Physical Cycle"
Text1.Text = "The physical cycle is 23 days long and is the dominant cycle in men. It regulates hand-eye coordination, strength, endurance, sex drive, initiative, metabolic rate, resistance to, and recovery from illness. Surgery should be avoided on physical transition days and during negative physical cycles. "
Case "Meaning of Emotional Cycle"
Text1.Text = "The emotional cycle is 28 days long andis the dominant cycle in women. It regulates emotions, feelings, mood, sensitivity, sensation, sexuality, fantasy, temperament, nerves, reactions, affections and creativity."
Case "Meaning of Intellectual Cycle"
Text1.Text = "The intellectual cycle is 33 days long and regulates intelligence, logic, mental reaction, alertness, sense of direction, decision-making, judgment, power of deduction, memory, and ambition."

Case "Meaning of Passion Cycle"
Text1.Text = "Passion cycle is the composite of the Physical and Emotional cycles. Passion encompasses your motivation to act, and the drive that allows you to continue a difficult pursuit. This cycle also tracks sexuality in its purest form."
Case "Meaning of Mastery Cycle"
Text1.Text = "Mastery Cycle is the composite of the Intellectual and Physical cycles. Mastery encompasses your ability to succeed at tasks and to obtain what you desire. This cycle also tracks athletic ability and the focus required to learn physical skills."
Case "Meaning of Wisdom Cycle"
Text1.Text = "Wisdom Cycle is  the composite of the Emotional and Intellectual cycles. Wisdom encompasses your understanding of the world, your role in it, and the things that are truly important to your life. This cycle also tracks the presence of mind that you need to make crucial decisions."

Case "Meaning of Intutive Cycle"
Text1.Text = "Intutive Cycle is of 38 days and shows your intution or sixth sense."

Case "Meaning of Aesthetic Cycle"
Text1.Text = "Aesthetic cycle is 43 days long and describes interest in the beautiful and the harmonious, Self-Awareness - 48 days, it expresses ability to percept own personality and individuality."
Case "Meaning of Self Awareness Cycle"
Text1.Text = "Self Awareness Cycle is of 48 days and describes your knowledge of inner self."
Case "Meaning of Spiritual Cycle"
Text1.Text = "Spiritual Cycle is 53 days long and  describes inner stability and relaxed attitude. "

Case "Negative and Positive 100%"
Text1.Text = "The numbers from +100% (maximum) to -100% (minimum) indicate where the rhythms are on a particular day. In general, a rhythm at 0% is thought to have no real impact on your life, whereas a rhythm at +100% (a high) would give you an edge in that area, and a rhythm at -100% (a low) would make life more difficult in that area. There is no particular meaning to a day on which your rhythms are all high or all low, except the obvious benefits or hindrances that these rare extremes are thought to have on your life."
Case "Balance is the key"
Text1.Text = "Understanding your positive cycles may assist you in planning surgery, physical outings, sporting events, exams, and job interviews. " & _
            "Understanding your negative cycles and your particular reaction to them…may help you avoid accidents, hurtful situations, unnecessary grief and misfortune."
Case "Advice"
Text1.Text = "Humans are wonderful being. For them every thing is possible. If you start strongly beliving this it will start manifesting(since every human has been given the power of  'Manifestation of repeated thoughts') and during your negative cycles you  may harm youself. I CANT DO WELL BECAUSE I AM GOING THROUGH A NEGATIVE PHASE  are words of a loser. My advice is to take its help only when you are confused."
End Select
End Sub
'RESTRICT THE ENTERY IN UNDERSTANDING BIORHYTHM SELECT
'QUESTION COMBOBOX
Private Sub Combo7_KeyPress(KeyAscii As Integer)
KeyAscii = 0
MsgBox "No entery allowed"
End Sub
'PUTS CURRENT DATE IN THE COMBOBOX
Private Sub cmdTodayDate_Click()
Combo5.Text = Month(Date)
Combo6.Text = Year(Date)
Combo4.Text = Day(Date)
End Sub

Private Sub cmdSecondaryCycles_Click()
Dim ndays As Double
Dim b As Integer, r As Integer, g As Integer
Dim l As Double, m As Double, n As Double

'CHECKS FOR VALIDITY OF DATE

If Not (IsDate(Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text) And IsDate(Combo4.Text & "/" & Combo5.Text & "/" & Combo6.Text)) Then
MsgBox "The Entered date is not valid."
Exit Sub
End If

Picture1.Cls
lastButtonPressed = 2
Label20.Visible = True


dateOfBirth = Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text
dateOfCondition = Combo4.Text & "/" & Combo5.Text & "/" & Combo6.Text

'FINDS THE AGE OF PERSON IN DAYS ON CONDITION DATE

ndays = DateDiff("d", dateOfBirth, dateOfCondition)
'FINDS THE EXACT POSITION OF EACH CYCLE

l = (ndays / physicalCycleDays - Int(ndays / physicalCycleDays)) * physicalCycleDays
m = (ndays / emotionalCycleDays - Int(ndays / emotionalCycleDays)) * emotionalCycleDays
n = (ndays / intellectualCycleDays - Int(ndays / intellectualCycleDays)) * intellectualCycleDays
 
b = (16 - l) * days
r = (16 - m) * days
g = (16 - n) * days

'MAKES SECONDARY WAVE WHICH IS COMBINATION OF TWO PRIMARY WAVES
Call MakeSecCurve(days + b, TopY, (days * physicalCycleDays) + b, RGB(255, 255, 0), days + r, TopY, (days * emotionalCycleDays) + r)
Call MakeSecCurve(days + g, TopY, (days * intellectualCycleDays) + g, RGB(255, 0, 255), days + r, TopY, (days * emotionalCycleDays) + r)
Call MakeSecCurve(days + g, TopY, (days * intellectualCycleDays) + g, RGB(0, 255, 255), days + b, TopY, (days * physicalCycleDays) + b)

End Sub
Private Sub cmdDivineCycles_Click()
Dim ndays As Double
Dim b As Integer, r As Integer, g As Integer, w As Integer
Dim l As Double, m As Double, n As Double, o As Double

'CHECKS FOR VALIDITY OF DATE
If Not (IsDate(Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text) And IsDate(Combo4.Text & "/" & Combo5.Text & "/" & Combo6.Text)) Then
    MsgBox "The Entered date is not valid."
    Exit Sub
End If

Picture1.Cls
lastButtonPressed = 3
Label20.Visible = True

dateOfBirth = Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text
dateOfCondition = Combo4.Text & "/" & Combo5.Text & "/" & Combo6.Text


'FINDS THE AGE OF PERSON IN DAYS ON CONDITION DATE
ndays = DateDiff("d", dateOfBirth, dateOfCondition)

'FINDS THE EXACT POSITION OF EACH CYCLE
l = ((ndays / intutiveCycleDays - Int(ndays / intutiveCycleDays)) * intutiveCycleDays)
m = ((ndays / aestheticCycleDays - Int(ndays / aestheticCycleDays)) * aestheticCycleDays)
n = ((ndays / selfAwarenessCycleDays - Int(ndays / selfAwarenessCycleDays)) * selfAwarenessCycleDays)
o = ((ndays / spiritualCycleDays - Int(ndays / spiritualCycleDays)) * spiritualCycleDays)

b = (16 - l) * days
r = (16 - m) * days
g = (16 - n) * days
w = (16 - o) * days

'DRAWS WAVES FROM CENTER TO BACKWARD AND CENTER TO FORWARD
Call MakeCurve(days + b, TopY, (days * intutiveCycleDays) + b, RGB(255, 200, 100))
Call MakeReverseCurve(days + b, TopY, (days * intutiveCycleDays) + b, RGB(255, 200, 100))

Call MakeCurve(days + r, TopY, (days * aestheticCycleDays) + r, RGB(100, 200, 255))
Call MakeReverseCurve(days + r, TopY, (days * aestheticCycleDays) + r, RGB(100, 200, 255))

Call MakeCurve(days + g, TopY, (days * selfAwarenessCycleDays) + g, RGB(200, 255, 100))
Call MakeReverseCurve(days + g, TopY, (days * selfAwarenessCycleDays) + g, RGB(200, 255, 100))

Call MakeCurve(days + w, TopY, (days * spiritualCycleDays) + w, RGB(255, 255, 255))
Call MakeReverseCurve(days + w, TopY, (days * spiritualCycleDays) + w, RGB(255, 255, 255))

End Sub


Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
Dim CntX As Integer, CntY As Integer
Dim strName As String, strDate As String
Dim syntaxFlag As Boolean, DoNotEnter As Boolean
Call OptimizeFormPosition
'SETS THE NUMBER OF DAYS FOR THE VARIOUS CYCLES

physicalCycleDays = 23
emotionalCycleDays = 28
intellectualCycleDays = 33
intutiveCycleDays = 38
aestheticCycleDays = 43
selfAwarenessCycleDays = 48
spiritualCycleDays = 53

'SETS VALUES OF DAYS AND POSITION OF THE WAVE
days = 240
TopY = 2280

'PUTS YEARS IN THE COMBOBOX
For CntX = 1900 To 2010
Combo3.AddItem CntX
Combo6.AddItem CntX
Combo10.AddItem CntX
Next

'PUTS MONTHS IN COMBOBOX

For CntX = 1 To 12
Combo2.AddItem CntX
Combo5.AddItem CntX
Combo9.AddItem CntX
Next

'PUTS DATES IN COMBOBOX

For CntX = 1 To 31
Combo1.AddItem CntX
Combo4.AddItem CntX
Combo8.AddItem CntX
Next

'PUTS TODAY'S DATE IN CONDITION DATE COMBOBOX

Combo5.Text = Month(Date)
Combo6.Text = Year(Date)
Combo4.Text = Day(Date)

'SETS COLOR FOR LEGEND
Label14.BackColor = RGB(255, 200, 100)
Label16.BackColor = RGB(100, 200, 255)
Label18.BackColor = RGB(200, 255, 100)


'PUTING QUESTIONS IN UNDERSTANDING BIORHYTHM AND OTHER SETTING
bShow = 0

Combo7.AddItem "What is Biorhythm?"
Combo7.AddItem "Meaning of Physical Cycle"
Combo7.AddItem "Meaning of Emotional Cycle"
Combo7.AddItem "Meaning of Intellectual Cycle"
Combo7.AddItem "Meaning of Passion Cycle"
Combo7.AddItem "Meaning of Mastery Cycle"
Combo7.AddItem "Meaning of Wisdom Cycle"
Combo7.AddItem "Meaning of Intutive Cycle"
Combo7.AddItem "Meaning of Aesthetic Cycle"
Combo7.AddItem "Meaning of Self Awareness Cycle"
Combo7.AddItem "Meaning of Spiritual Cycle"
Combo7.AddItem "Negative and Positive 100%"
Combo7.AddItem "Balance is the key"
Combo7.AddItem "Advice"

'LOADING SAVED NAMES AND DOB'S
Call List_Load(List1, App.Path & "\DOB_BRHTHM3.DAT")

'LOAD NAMES IN COMBOBOX

For CntX = 0 To List1.ListCount - 1
    For CntY = 1 To Len(List1.list(CntX))
        If Not ("|" = Mid((List1.list(CntX)), CntY, 1)) Then
                strName = strName & Mid((List1.list(CntX)), CntY, 1)
        Else
                Combo11.AddItem strName
                strName = ""
                Exit For
               
        End If
    Next
Next

'LOADS DEFAULT NAME IN NAME TEXTBOX
If List1.ListCount <> 0 Then
CntX = 0
syntaxFlag = True
DoNotEnter = False

strName = ""
    For CntY = 1 To Len(List1.list(CntX))
        If "|" <> Mid((List1.list(CntX)), CntY, 1) And Not DoNotEnter Then
                strName = strName & Mid((List1.list(CntX)), CntY, 1)
        Else
                DoNotEnter = True
                    If ("|" = Mid((List1.list(CntX)), CntY, 1)) Then
                           CntY = CntY + 1
                           syntaxFlag = False
                            
                    End If
                    If Not (syntaxFlag) Then
                            strDate = strDate & Mid((List1.list(CntX)), CntY, 1)
                    End If
                If Len(List1.list(CntX)) = CntY Then
                Combo1.Text = Day(CDate(strDate))
                Combo2.Text = Month(CDate(strDate))
                Combo3.Text = Year(CDate(strDate))
                Combo11.Text = strName
                End If
        End If
    Next
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Call List_Save(List1, App.Path & "\DOB_BRHTHM3.DAT")
End Sub





Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'SHOWS DATE AND TIME WHEN MOUSE MOVES OVER THE GRAPH
Label20.Caption = dateOfCondition - (16 - (X / days))
End Sub

Private Sub Form_Paint()
' RESET OR REPAINTS THE GRAPH IF THE FORM HAS BEEN MINIMIZED
' OR PART OR ALL OF AN OBJECT IS EXPOSED AFTER BEING MOVED
' OR ENLARGED,
' OR AFTER A WINDOW THAT WAS COVERING THE OBJECT HAS BEEN MOVED
If lastButtonPressed = 1 Then
cmdPrimaryCycles_Click
Exit Sub
End If

If lastButtonPressed = 2 Then
cmdSecondaryCycles_Click
Exit Sub
End If

If lastButtonPressed = 3 Then
cmdDivineCycles_Click
Exit Sub
End If

End Sub
'POSITIONS THE FORM TO THE CENTER
Private Sub OptimizeFormPosition()
Me.Left = (Screen.Width / 2) - (Me.Width / 2)
Me.Top = ((Screen.Height - 322) / 2) - (Me.Height / 2)

End Sub
Public Sub List_Load(thelist As ListBox, FileName As String)
'LOADS A FILE TO A LIST BOX
    On Error Resume Next
    Dim TheContents As String
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Input As fFile
    Do
        Line Input #fFile, TheContents$
        If TheContents$ = "" Then
        Else
        Call List_Add(thelist, TheContents$)
        End If
    Loop Until EOF(fFile)
    Close fFile
End Sub

Public Sub List_Save(thelist As ListBox, FileName As String)
    'SAVE A LISTBOX AS FILENAME
    On Error Resume Next
    Dim Save As Long
    Dim fFile As Integer
    fFile = FreeFile
    Open FileName For Output As fFile
    For Save = 0 To thelist.ListCount - 1
        Print #fFile, List1.list(Save)
    Next Save
    Close fFile
End Sub
Public Sub List_Add(list As ListBox, txt As String)
On Error Resume Next
    List1.AddItem txt
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("|") Then
KeyAscii = 0
MsgBox "The charachter '|' is a sysntax indentifier."
End If
End Sub
