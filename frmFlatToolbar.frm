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
            Caption         =   ""
            Key             =   ""
            Description     =   ""
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
            Caption         =   ""
            Key             =   ""
            Description     =   ""
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
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Boxer"
            Object.Tag             =   ""
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Zucca"
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Teschio"
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   ""
            Object.Tag             =   ""
            ImageIndex      =   6
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Sigaretta"
            Object.Tag             =   ""
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Sapone"
            Object.Tag             =   ""
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Carta igienica"
            Object.Tag             =   ""
            ImageIndex      =   8
            Style           =   2
         EndProperty
      EndProperty
      MouseIcon       =   "frmFlatToolbar.frx":0000
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
            Picture         =   "frmFlatToolbar.frx":001C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFlatToolbar.frx":08F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFlatToolbar.frx":11D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFlatToolbar.frx":1AB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFlatToolbar.frx":238C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFlatToolbar.frx":2C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFlatToolbar.frx":3544
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFlatToolbar.frx":3E20
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


'-------------------------------------------------------------------------------
'
' FUNCTION    : SetFlatToolbar(Toolbar)
'
' AUTHOR      : Roal Zanazzi
'
' DESCRIPTION : Questa funzione  attiva lo stile "flat" per il controllo Toolbar
'               dei common controls di Windows 95.
'               Vengono utilizzate alcune funzioni delle API WIN32 per
'               modificare lo stile della finestra della toolbar.
'
' PARAMETERS  : - aToolbar: Toolbar     Il controllo toolbar da modificare.
'
' NOTE        : Funziona solo se sul sistema e' installata la COMCTL32.DLL
'               aggiornata (distribuita con MS Internet Explorer 3 o successivi
'               (questa limitazione e' relativa solo a Windows 95 e NT 4.0,
'               nelle versioni successive la DLL dovrebbe essere gia' aggiornata)
'
' EXAMPLE     : SetFlatToolbar myToolbar
'
'-------------------------------------------------------------------------------

' API constants.
Private Const GW_CHILD = 5
Private Const GWL_STYLE = (-16)
' API functions.
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Sub SetFlatToolbar(aToolbar As Toolbar)
    ' Costante per lo stile "flat" della toolbar.
    Const TB_STYLE_FLAT = &H848
    ' Handle della finestra della toolbar.
    Dim hwnd As Long
    ' Questa deferenziazione alla finestra child e' necessaria perche' il
    ' controllo toolbar di VB5 e' in realta' il contenitore per la toolbar
    ' vera e propria!
    hwnd = GetWindow(aToolbar.hwnd, GW_CHILD)
    ' Modifica lo stile della toolbar per avere il look "flat".
    Call SetWindowLong(hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or TB_STYLE_FLAT)
End Sub


' Esempio di chiamata della routine.
Private Sub Form_Load()
  SetFlatToolbar Toolbar1
End Sub
