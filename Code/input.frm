VERSION 5.00
Begin VB.Form inputbox 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "input"
   ClientHeight    =   1515
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      MaskColor       =   &H00FFFFFF&
      Picture         =   "input.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter subject  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1395
      Left            =   -480
      Picture         =   "input.frx":14A4
      Top             =   -120
      Width           =   2025
   End
End
Attribute VB_Name = "inputbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i
Option Explicit


Public Function inputresult() As Integer

  inputresult = i

End Function




Public Sub OKButton_Click()
   
   
End Sub

