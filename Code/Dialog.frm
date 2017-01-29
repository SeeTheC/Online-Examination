VERSION 5.00
Begin VB.Form errormsg 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "! ERROR"
   ClientHeight    =   2265
   ClientLeft      =   5070
   ClientTop       =   5010
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox msg 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Dialog.frx":0000
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      Picture         =   "Dialog.frx":0006
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   0
      Picture         =   "Dialog.frx":04DD
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "errormsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()

    Unload Me
End Sub
