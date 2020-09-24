VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About RegView"
   ClientHeight    =   2508
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   3768
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2508
   ScaleWidth      =   3768
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   372
      Left            =   1134
      TabIndex        =   2
      Top             =   2088
      Width           =   1416
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Udi Azulay - Udi@Datascope.co.il"
      Height          =   192
      Left            =   144
      TabIndex        =   3
      Top             =   1836
      Width           =   2436
   End
   Begin VB.Label DateLbl 
      AutoSize        =   -1  'True
      Caption         =   "Made By Control^Zed  - (37254786)  6/2000"
      Height          =   192
      Left            =   144
      TabIndex        =   1
      Top             =   1584
      Width           =   3048
   End
   Begin VB.Label VerLbl 
      AutoSize        =   -1  'True
      Caption         =   "Registry Viewer Version : "
      Height          =   192
      Left            =   144
      TabIndex        =   0
      Top             =   1332
      Width           =   1824
   End
   Begin VB.Image Logo 
      Height          =   984
      Left            =   954
      Picture         =   "FrmAbout.frx":0000
      Top             =   108
      Width           =   1860
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    VerLbl.Caption = "Registry Viewer Version : " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Resize()
    Logo.Move (Me.ScaleWidth - Logo.Width) / 2, Me.ScaleTop + 50
    VerLbl.Move 100, Logo.Top + Logo.Height + 250
    DateLbl.Move 100, VerLbl.Top + VerLbl.Height + 50
    Label1.Move 100, DateLbl.Top + DateLbl.Height + 50
    Command1.Move (Me.ScaleWidth - Command1.Width) / 2
End Sub

Private Sub Logo_DblClick()
Dim Msg As String
    Msg = "This Program Fully Made By Udi Azulay."
    MsgBox Msg, vbOKOnly, "Hello"
End Sub
