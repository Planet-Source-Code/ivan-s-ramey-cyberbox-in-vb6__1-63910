VERSION 5.00
Begin VB.Form frmCyberBox 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "CyberBox"
   ClientHeight    =   7440
   ClientLeft      =   0
   ClientTop       =   450
   ClientWidth     =   9855
   Icon            =   "frmCyberBox.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCyberBox.frx":23E2
   ScaleHeight     =   496
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   657
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox CBFrame 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   240
      Picture         =   "frmCyberBox.frx":F10E4
      ScaleHeight     =   150
      ScaleWidth      =   150
      TabIndex        =   44
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox Block 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   465
      Index           =   14
      Left            =   240
      Picture         =   "frmCyberBox.frx":1DFDE6
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   43
      Top             =   1200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Block 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   465
      Index           =   13
      Left            =   240
      Picture         =   "frmCyberBox.frx":1E09C8
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   42
      Top             =   1200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Block 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   465
      Index           =   12
      Left            =   240
      Picture         =   "frmCyberBox.frx":1E15AA
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   41
      Top             =   1200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Block 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   465
      Index           =   11
      Left            =   240
      Picture         =   "frmCyberBox.frx":1E218C
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   40
      Top             =   1200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Block 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   465
      Index           =   10
      Left            =   240
      Picture         =   "frmCyberBox.frx":1E2D6E
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   39
      Top             =   1200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Block 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   465
      Index           =   9
      Left            =   240
      Picture         =   "frmCyberBox.frx":1E3950
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   38
      Top             =   1200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Block 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   465
      Index           =   8
      Left            =   240
      Picture         =   "frmCyberBox.frx":1E4532
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   37
      Top             =   1200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Block 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   465
      Index           =   7
      Left            =   240
      Picture         =   "frmCyberBox.frx":1E5114
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   36
      Top             =   1200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Block 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   465
      Index           =   6
      Left            =   240
      Picture         =   "frmCyberBox.frx":1E5CF6
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   35
      Top             =   1200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Block 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   465
      Index           =   5
      Left            =   240
      Picture         =   "frmCyberBox.frx":1E68D8
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   34
      Top             =   1200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Block 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   465
      Index           =   4
      Left            =   240
      Picture         =   "frmCyberBox.frx":1E74BA
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   33
      Top             =   1200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Block 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   465
      Index           =   3
      Left            =   240
      Picture         =   "frmCyberBox.frx":1E809C
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   32
      Top             =   1200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Block 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   465
      Index           =   2
      Left            =   240
      Picture         =   "frmCyberBox.frx":1E8C7E
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   31
      Top             =   1200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Block 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   465
      Index           =   1
      Left            =   240
      Picture         =   "frmCyberBox.frx":1E9860
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   30
      Top             =   1200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Block 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   465
      Index           =   0
      Left            =   240
      Picture         =   "frmCyberBox.frx":1EA442
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   29
      Top             =   1200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Rnum 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   4
      Left            =   240
      Picture         =   "frmCyberBox.frx":1EB024
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   28
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox RetryNum 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3390
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   27
      Top             =   555
      Width           =   210
   End
   Begin VB.PictureBox Rnum 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   3
      Left            =   240
      Picture         =   "frmCyberBox.frx":1EB2CE
      ScaleHeight     =   210
      ScaleWidth      =   180
      TabIndex        =   26
      Top             =   600
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox Rnum 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   2
      Left            =   240
      Picture         =   "frmCyberBox.frx":1EB508
      ScaleHeight     =   210
      ScaleWidth      =   180
      TabIndex        =   25
      Top             =   600
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox Rnum 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   1
      Left            =   240
      Picture         =   "frmCyberBox.frx":1EB742
      ScaleHeight     =   210
      ScaleWidth      =   180
      TabIndex        =   24
      Top             =   600
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox Rnum 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   0
      Left            =   240
      Picture         =   "frmCyberBox.frx":1EB97C
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   23
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox RoomNum 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   3570
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   22
      Top             =   6790
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox rnpic 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   16
      Left            =   720
      Picture         =   "frmCyberBox.frx":1EBC26
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox rnpic 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   15
      Left            =   720
      Picture         =   "frmCyberBox.frx":1EDE94
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox rnpic 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   14
      Left            =   720
      Picture         =   "frmCyberBox.frx":1F0DAA
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox rnpic 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   13
      Left            =   720
      Picture         =   "frmCyberBox.frx":1F315C
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox rnpic 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   12
      Left            =   720
      Picture         =   "frmCyberBox.frx":1F55EE
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox rnpic 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   11
      Left            =   720
      Picture         =   "frmCyberBox.frx":1F8504
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox rnpic 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   10
      Left            =   720
      Picture         =   "frmCyberBox.frx":1FAE66
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox rnpic 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   9
      Left            =   720
      Picture         =   "frmCyberBox.frx":1FD218
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   14
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox rnpic 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   8
      Left            =   720
      Picture         =   "frmCyberBox.frx":1FF55E
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox rnpic 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   7
      Left            =   720
      Picture         =   "frmCyberBox.frx":202160
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox rnpic 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   6
      Left            =   720
      Picture         =   "frmCyberBox.frx":204B66
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox rnpic 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   5
      Left            =   720
      Picture         =   "frmCyberBox.frx":206C78
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox rnpic 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   4
      Left            =   720
      Picture         =   "frmCyberBox.frx":20916E
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox rnpic 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   3
      Left            =   720
      Picture         =   "frmCyberBox.frx":20B448
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox rnpic 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   2
      Left            =   720
      Picture         =   "frmCyberBox.frx":20DFDA
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox rnpic 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   1
      Left            =   720
      Picture         =   "frmCyberBox.frx":20FD38
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox rnpic 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   720
      Picture         =   "frmCyberBox.frx":211BAA
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox About 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   5250
      Left            =   990
      ScaleHeight     =   5250
      ScaleWidth      =   7875
      TabIndex        =   4
      Top             =   1290
      Visible         =   0   'False
      Width           =   7875
   End
   Begin VB.PictureBox picInstr 
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   240
      LinkTimeout     =   0
      Picture         =   "frmCyberBox.frx":213A1C
      ScaleHeight     =   150
      ScaleWidth      =   150
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picAbo 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   240
      Picture         =   "frmCyberBox.frx":29A50E
      ScaleHeight     =   150
      ScaleWidth      =   150
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox Field 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   5190
      Left            =   1020
      ScaleHeight     =   346
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   521
      TabIndex        =   0
      Top             =   1320
      Width           =   7815
      Begin VB.Image PersonShape 
         Height          =   435
         Left            =   3690
         Picture         =   "frmCyberBox.frx":321000
         Top             =   4740
         Width           =   435
      End
   End
   Begin VB.Label GOver 
      BackColor       =   &H00000000&
      Caption         =   "Game Over"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   5190
      Left            =   1020
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   7815
   End
End
Attribute VB_Name = "frmCyberBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#################################################
'# Doug Beeferman's  C Y B E R B O X             #
'#                                               #
'# VB implementation:                            #
'#  David Yang                                   #
'#  Ivan S. Ramey                                #
'#  Paul Bahlawan                                #
'#################################################

Option Explicit
Private Enum Blocks
Blank
UpDown
LeftRight
AllMove
Blocker
PushUp
PushDown
PushLeft
PushRight
ZapUp
ZapDown
ZapLeft
ZapRight
BlankSelector
FullSelector
End Enum

Private Tiles(0 To 14)                   As IPictureDisp
Private Person                           As PersonLocation
Private WorldArray(1 To 15, 1 To 10)     As Integer
Private retries                          As Integer
Private moves                            As Integer
Private level                            As Integer
Private mode                             As Integer
Private win                              As Boolean
Private lvlData(1 To 17)                 As String

Private Type PersonLocation
    x                                    As Integer
    y                                    As Integer
End Type

Private Sub Form_Load()
Dim Counter As Long
    For Counter = 0 To 14
        Set Tiles(Counter) = Block(CStr(Counter)).Picture
    Next Counter
    TitleScreen
End Sub

Private Sub Levels()
lvlData(1) = "AABBAAABBAAAAAAEAAAAAAAEAAAEAAAAAAAEAAAAEAAAAAAAACACAEACAEACABDAAAAAAAACABADABAAABDAAAAAAAACAEACAEACEAAAAAAAACAAAAAEAAAAAAAEAAAEAAAAAAAEAAAAABBAAAABBA"
lvlData(2) = "AAAABBBBBAAAAAADEAAAAACECAEAEEEECAAEABAAAECAAEAADEAACAAAEABAAACAAAEEAAAABDAAEEAAAACBDAEEECEACAAAAAAAEECAAEAECEAAAEAAAABAAAAEAAECECAAEEAAAAAAAABBAAAAAA"
lvlData(3) = "EAAAAAAAAAAABAEAEAAAABDBAEAEAEEEEAAEADBAAAEAAEADBAAEECAEAAAEEAEAAAEEAAAAEAEAAAEAAAEAEEEAAEABDBDBBAAAABDAEAEIEEEAEAAAABAAAAEAECCEDDAEAEACCAAAAAAAAAAAAE"
lvlData(4) = "AAAAAABBAAAAAEEEEHEAAAEAAAABAAAAEACEEHEAAEAAABBBAAAAEAEAEHEADAEAEAAAEACAEAAEACEADAAECAECEEDEAEACECAADAAEAAECACDECEAAEAACDAAAAEAAACGCAEEADAACABABAAAAAA"
lvlData(5) = "AAEEEAAAEAAEAAABBDBAAAEEECEEEAAAAEGCAAEAAABBGCAAEAAAAEEDEAEAEABAAAEEEAAEABAAEEEAAAEABAAEEEABAEAEEEEAAAAEAAAAAAACAEEEECDACAEAAAACCACEAAAAEDCAAAABAAEAAA"
lvlData(6) = "AAAAAIEAAAAAAAACEAAAAAAAACEAAAECAAACEAAAACAAABAAAAACAAAEECEAGGCAEEAAEAACAAAEEAEAACAAAEAAEEACAAAEAAEAACAAAEAAEAACAAAEEABAACEAAEEEDAABDBBBBBCFAHAAAEAAAA"
lvlData(7) = "CAAEAAAEAACAAEIAAEAAGCAABAAAAACAAAAAACAACADDDDDCAACAICAAADAACEIAAAACAAADBACFACAAAAAEAAACAAEEEAAAACAAAAAAAAACAAAAAEAAAAAAAAEAAAEEAEAAEAAEAAAAAAAEEAAAAA"
lvlData(8) = "AAEAAJAAAAAEAAEIEEEAAEAKABAAEAAAEAEAEEEAAAAAAAEAAAAEAAAAELLLEABDBCFAAAAJCEECEAAAEABDBCFEEEAEAAACFAAAAAAAAAAAAAAAAAAAAAAAAEEEEAEEAAAEEEEAEEAAAAAAAAAAAA"
lvlData(9) = "ABAJAAJAMAACEAAEALALACEEEAJAJAACKAELAMALACAAEAKAKAEDAEAAAKAEGCMELEEEEEGCAAAAAJAAGDLEMEEEEEAAAEAAAAAAAAAEEEAAAACAEAAECACAGCAAAACADAAAAAAECGCAAAAAAEAAAE"
lvlData(10) = "AAAAAAJAAAEIAAAEAEAAABDBBBCBBAAAEAAADAEAAAAAAACAEAAAAEEECEAAEAEAIECEAAACACBAAJAAEAEAEEEEAAAAEAAEIEAAEAEAAEIAAAAEAAAEIAAAAEAAEADAAEAEAAEADAAEAEAAAADAAE"
lvlData(11) = "AAAAAAAAAAAAAAAAAAAAAAEEEEAAAAAEAAAAEAAAEADACAAAAAEACCDCAAECECCCCCAEGCGCCCCCCAEAECCDDDAEGCEDACCCAAECECACCCAAAAEAAAAAAAAAAEAAAAEAAAAAEEEEAAAAAAAAAAAAAA"
lvlData(12) = "AAEAAAAAAAAAJAAAAAAAEAEALLLAAAGCBBAAAEAAAAACAAAEAAOAACDDDAOOCAACAAAEAACEACEAAECAAAECKDDEAAAAECEAAAOOGBDHEAEEAAAAEAECEAAAAAABACFAAAAAAAAAJAAAAAAAAAJAAA"
lvlData(13) = "AAAEAAEAAAAAAEAAAAAAAAAEAEEAAAAAAEAAEAAAAAAEAAEAAAAAEAAAEDAAOAEAEAEAEAOAEAECELEAOEAAEAAAAEAEAAEAEEAAAEAEAAAAAAAELEAAEEAEAAAAELAAAEAAAAAAAACFAAAAEAAAAE"
lvlData(14) = "AAJAKAJAKAAMALAMALAMMAKAKAKAJAAMALALAMALMAJAJAKAJAALAMALAMALMAJAJAKAJAALALALALAAMAKAKAKAJAALAKAMAMAMLAJALAJAJAAMALAJAMAMLAKAMAJAKAALALAMAMAMAAJAJAKAJA"
lvlData(15) = "AAAAAEAEAADDEABNBEAAAAEABNBAAAAAEEAAAEAAAEAEEEEEAECEAEAAAEAACEEAAAAAAAGBBBONNEAACEEEEAAAAACAAAEAAEAAGONNEEEEAECAAADAAAAACAAAAEEAAAGONNEEAAAACAAAAEAAAA"
lvlData(16) = "AEAJAEAAAAAEAALAAACAAABONAAACFAANEAAAADAAABONAAACFAEAEAAAAAAEAAAEEEELEAEEAAAKAAALAABONOBAEAEEAEAEMEEAEAAEAEAEAAEACECECKAAEENENENEAAAACACACAAAAEAEAEAEA"
lvlData(17) = "AAAAAAAAAAAAAAAACCAABBAAAAAACCAABBBAACCABBAAAAAACCAAAAAACCAAABBBAAAAAAGDDDFABBBAABBBAAAAAAAAAAAABBBBEEEEAAADAAAAAAEAAADAAAAAEABBBBEEEEAAAAAAAAAAAAAAAA"
End Sub

Private Sub TitleScreen()
    Field.Visible = False
    GOver.Visible = False
    RetryNum.Visible = False
    RoomNum.Visible = False
    mode = 0
End Sub

Private Sub Game()
    Me.Picture = CBFrame.Picture
    RetryNum.Visible = True
    RoomNum.Visible = True
    Field.Visible = True
    mode = 1
    level = 1
    retries = 5
    moves = 0
    Levels
    Retry
End Sub

Private Sub NextLevel()
Dim x As Long
Dim y As Long
    level = level + 1
    If level = 18 Then
        win = True
        GameOver
    Else
        win = False
        RoomNum.Picture = rnpic(CStr(level - 1)).Picture
        For x = 1 To 15
            For y = 1 To 10
                WorldArray(x, y) = Asc(Mid$(lvlData(level), (x - 1) * 10 + y, 1)) - 65
            Next y
        Next x
        Person.x = 8
        Person.y = 10
        ReDraw
    End If
End Sub

Private Sub ReDraw()
Dim x As Integer
Dim y As Integer
    Me.Cls
    'draw tiles
    For x = 1 To 15
        For y = 1 To 10
            Field.PaintPicture Tiles(WorldArray(x, y)), (x - 1) * 35, (y - 1) * 35
        Next y
    Next x
    'draw (position) player
    PersonShape.Top = (Person.y - 1) * 35 + 1
    PersonShape.Left = (Person.x - 1) * 35 + 1
End Sub

Private Sub Retry()
    retries = retries - 1
    Select Case retries
    Case 0 To 4
        RetryNum.Picture = Rnum(CStr(retries)).Picture
        level = level - 1
        NextLevel
    Case -1
        win = False
        GameOver
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case mode
    Case 0 'Title Screen mode
        Select Case KeyCode
        Case Else
            Game
        End Select
    Case 1 'Game mode
        Select Case KeyCode
        Case vbKeyLeft, vbKey4, vbKeyNumpad4
            If About.Visible = True Then
                About.Visible = False
            Else
                moves = moves + 1
                PersonPushTile Person.x, Person.y, -1, 0
            End If
        Case vbKeyRight, vbKey6, vbKeyNumpad6
            If About.Visible = True Then
                About.Visible = False
            Else
                moves = moves + 1
                PersonPushTile Person.x, Person.y, 1, 0
            End If
        Case vbKeyUp, vbKey8, vbKeyNumpad8
            If About.Visible = True Then
                About.Visible = False
            Else
                moves = moves + 1
                PersonPushTile Person.x, Person.y, 0, -1
            End If
        Case vbKeyDown, vbKey2, vbKeyNumpad2
                If About.Visible = True Then
                About.Visible = False
            Else
                moves = moves + 1
                PersonPushTile Person.x, Person.y, 0, 1
            End If
        Case vbKeyEscape
            If About.Visible = True Then
                About.Visible = False
            Else
                GameOver
            End If
        Case vbKeyR
            If About.Visible = True Then
                About.Visible = False
            Else
                Retry
            End If
        Case vbKeyL ' cheat !!!
            If About.Visible = True Then
                About.Visible = False
            Else
                NextLevel
            End If
        Case vbKeyA
            If About.Visible = True Then
                About.Visible = False
            Else
                About.Picture = picAbo.Picture
                About.Visible = True
            End If
        Case vbKeyI
            If About.Visible = True Then
                About.Visible = False
            Else
                About.Picture = picInstr.Picture
                About.Visible = True
            End If
        Case Else
            If About.Visible = True Then
                About.Visible = False
            End If
        End Select
        CheckPushers
        ReDraw
    Case 2 'Game over mode
        Select Case KeyCode
        Case vbKeyEscape
            Unload Me
            End
        End Select
    End Select
End Sub

Private Sub PersonPushTile(ByVal x As Long, ByVal y As Long, DirectionX As Long, DirectionY As Long)
Dim NewX As Long
Dim NewY As Long
    NewX = x + DirectionX
    NewY = y + DirectionY
    Person.x = x
    Person.y = y
    'Check if the person is in the play area...
    If NewX > 0 And NewY > 0 And NewX < 16 And NewY < 11 Then
        If WorldArray(NewX, NewY) = Blank Then
            'if it is a blank, move there
            Person.x = NewX
            Person.y = NewY
        ElseIf DirectionX <> 0 Then 'NOT WORLDARRAY(NEWX,...
            'movement left/right, but not into a blank
            Select Case WorldArray(NewX, NewY)
            Case LeftRight, AllMove
                'Check if it can push (with the other tiles)
                If InteractTiles(NewX, NewY, DirectionX, DirectionY) Then
                    'Can push
                    Person.x = NewX
                    Person.y = NewY
                End If
            Case BlankSelector
                'Cannot push, but can walk through
                Person.x = NewX
                Person.y = NewY
            Case ZapRight, ZapLeft
                'use exclusive to check - one has to be true and one false
                If DirectionX > -1 Xor WorldArray(NewX, NewY) = ZapLeft Then
                    'Only jump into an empty space
                    If WorldArray(NewX + DirectionX, NewY) = Blank Then
                        Person.x = NewX + DirectionX
                        Person.y = NewY
                    End If
                End If
            End Select
        Else 'Y direction...
            'Up/down
            Select Case WorldArray(NewX, NewY)
            Case UpDown, AllMove
                'Check if it can push (with the other tiles)
                If InteractTiles(NewX, NewY, DirectionX, DirectionY) Then
                    'Can push
                    Person.x = NewX
                    Person.y = NewY
                End If
            Case BlankSelector
                'Cannot push, but can walk through
                Person.x = NewX
                Person.y = NewY
            Case ZapUp, ZapDown
                'use exclusive to check - one has to be true and one false
                If DirectionY > -1 Xor WorldArray(NewX, NewY) = ZapUp Then
                    If WorldArray(NewX, NewY + DirectionY) = Blank Then
                        Person.x = NewX
                        Person.y = NewY + DirectionY
                    End If
                End If
            End Select
        End If
    ElseIf NewX = 8 And NewY = 0 Then
        NextLevel
    End If
End Sub

Private Function InteractTiles(ByVal x As Long, ByVal y As Long, DirectionX As Long, DirectionY As Long) As Boolean
Dim NewX As Long
Dim NewY As Long
    NewX = x + DirectionX
    NewY = y + DirectionY
    InteractTiles = False
    If NewX > 0 And NewY > 0 And NewX < 16 And NewY < 11 Then
        If DirectionX <> 0 Then
            Select Case WorldArray(NewX, NewY)
            Case LeftRight, AllMove, BlankSelector, FullSelector
                InteractTiles = InteractTiles(NewX, NewY, DirectionX, DirectionY)
            Case Blank
                InteractTiles = True
            End Select
        ElseIf DirectionY <> 0 Then 'NOT DIRECTIONX...
            Select Case WorldArray(NewX, NewY)
            Case UpDown, AllMove, BlankSelector, FullSelector
                InteractTiles = InteractTiles(NewX, NewY, DirectionX, DirectionY)
            Case Blank
                InteractTiles = True
            End Select
        End If
    End If
    If NewX = Person.x And NewY = Person.y Then
        InteractTiles = False
    End If
        If x = Person.x And y = Person.y Then
        InteractTiles = False
    End If
    If InteractTiles Then
        WorldArray(NewX, NewY) = WorldArray(x, y)
        WorldArray(x, y) = Blank
    End If
End Function

Private Sub CheckPushers()
Dim x As Long
Dim y As Long
Dim more As Boolean
    For x = 1 To 15
        For y = 1 To 10
            Select Case WorldArray(x, y)
            Case PushLeft
                If FullyInteractTiles(x, y, -1, 0) Then more = True
            Case PushRight
                If FullyInteractTiles(x, y, 1, 0) Then more = True
            Case PushUp
                If FullyInteractTiles(x, y, 0, -1) Then more = True
            Case PushDown
                If FullyInteractTiles(x, y, 0, 1) Then more = True
            End Select
        Next y
    Next x
    If more Then CheckPushers
End Sub
 
Private Function FullyInteractTiles(ByVal x As Long, ByVal y As Long, DirectionX As Long, DirectionY As Long) As Boolean
'Same as interact when pushing, but pushes all the way
    If InteractTiles(x, y, DirectionX, DirectionY) Then
        WorldArray(x, y) = Blank
        FullyInteractTiles x + DirectionX, y + DirectionY, DirectionX, DirectionY
        FullyInteractTiles = True
    End If
End Function

Private Sub GameOver()
Dim msg As String
    mode = 2
    Field.Visible = False
    If level = 18 Then
        msg = "Congratulations!"
    Else
        msg = "Game Over."
    End If
    level = level - 1
    If retries < 0 Then
        retries = 0
    End If
    GOver.Caption = vbNewLine & vbNewLine & _
     msg & vbNewLine & vbNewLine
If level >= 10 Then
     GOver.Caption = GOver.Caption & _
     "Rooms completed...  " & level & "x200 = " & level * 200 & vbNewLine
Else
     GOver.Caption = GOver.Caption & _
     "Rooms completed...   " & level & "x200 = " & level * 200 & vbNewLine
End If
     GOver.Caption = GOver.Caption & _
     "Unused attempts...   " & retries & "x50  = " & retries * 50 & vbNewLine & _
     "Moves made...               -" & moves & vbNewLine & _
     "                             ------" & vbNewLine & _
     "Total score...               " & level * 200 + retries * 50 - moves & vbNewLine & vbNewLine
If win = True Then
     GOver.Caption = GOver.Caption & "The original DOS CyberBox by Doug Beeferman" & vbNewLine & _
     "ported to Windows in Visual Basic 6.0" & vbNewLine & _
     "by Ivan S. Ramey, David Yang, and Paul Bahlawan." & vbNewLine & vbNewLine & _
     "Press escape key to exit this program."
Else
     GOver.Caption = GOver.Caption & "Press escape key to exit this program."
End If
    GOver.Visible = True
End Sub

'A good portion of this code done by David Yang.  Thanks!
'Also thanks to Paul Bahlawan for fixing the selector problem.
