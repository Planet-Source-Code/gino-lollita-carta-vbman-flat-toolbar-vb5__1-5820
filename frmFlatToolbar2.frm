VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmFlatToolbar 
   Caption         =   "Prova di coolbar (toolbar flat)"
   ClientHeight    =   2160
   ClientLeft      =   1170
   ClientTop       =   1545
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   5115
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   11
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Birra"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Scarpa"
            Object.Tag             =   ""
            ImageIndex      =   2
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Boxer"
            Object.Tag             =   ""
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Zucca"
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Teschio"
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   6
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Sigaretta"
            Object.Tag             =   ""
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Sapone"
            Object.Tag             =   ""
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Carta igienica"
            Object.Tag             =   ""
            ImageIndex      =   8
            Style           =   2
         EndProperty
      EndProperty
      MouseIcon       =   "frmFlatToolbar2.frx":0000
   End
   Begin VB.Line Line4 
      X1              =   2280
      X2              =   2280
      Y1              =   720
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   240
      Y1              =   720
      Y2              =   960
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Pulsanti a doppio stato: Checked è Unchecked"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin VB.Line Line2 
      X1              =   4080
      X2              =   4080
      Y1              =   720
      Y2              =   1320
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Gruppo di pulsanti: (è possibile selezionare un solo pulsante per volta)"
      Height          =   615
      Left            =   3240
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   960
      X2              =   960
      Y1              =   720
      Y2              =   1320
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Pulsante MixedState"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Pulsante Normale"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFlatToolbar2.frx":001C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFlatToolbar2.frx":0336
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFlatToolbar2.frx":0650
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFlatToolbar2.frx":096A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFlatToolbar2.frx":0C84
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFlatToolbar2.frx":0F9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFlatToolbar2.frx":12B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFlatToolbar2.frx":15D2
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFlatToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Non richiede la presenza di cmctl32.dll aggiornata
' Find Window Api Function - Used to Find ToolBar
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
                        (ByVal hWnd1 As Long, _
                        ByVal hWnd2 As Long, _
                        ByVal lpsz1 As String, _
                        ByVal lpsz2 As String) As Long
                                                                                                                                                                  
' Send Message Api Function - Used to Get and Set Toolbar Styles
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                        (ByVal hwnd As Long, _
                        ByVal wMsg As Long, _
                        ByVal wParam As Integer, _
                        ByVal lParam As Any) As Long

' General Constant
Private Const WM_USER = &H400

'Toolbar Const
Private Const TBSTYLE_TRANSPARENT = &H8000
Private Const TBSTYLE_FLAT = &H800
Private Const TB_SETSTYLE = (WM_USER + 56)
Private Const TB_GETSTYLE = (WM_USER + 57)
Private Const CCS_NODIVIDER = &H40

Private Sub Form_Load()
   Dim lStyle As Long                               'Variable to Hold Style
   Dim lReturn As Long                            ' Variable to hold Return Values
   Dim dlgToolBarHandle As Long             'Handle of the Toolbar, contained withing dialog box

    ' Find the Toolbar Handle
   dlgToolBarHandle = FindWindowEx(Toolbar1.hwnd, 0&, "ToolbarWindow32", vbNullString)

   ' Get the Toolbar Style
   lStyle = SendMessage(dlgToolBarHandle, TB_GETSTYLE, 0&, 0&)
   
   ' Change Style bits - Add Flat Style
   lStyle = lStyle Or TBSTYLE_FLAT Or TBSTYLE_TRANSPARENT Or CCS_NODIVIDER
   
   ' Set New Style
   lReturn = SendMessage(dlgToolBarHandle, TB_SETSTYLE, 0, lStyle)

End Sub


