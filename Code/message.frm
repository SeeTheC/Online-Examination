VERSION 5.00
Begin VB.Form message 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "infomation"
   ClientHeight    =   1680
   ClientLeft      =   4830
   ClientTop       =   4665
   ClientWidth     =   5475
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox msg 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   975
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "message.frx":0000
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   0
      Picture         =   "message.frx":0015
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Option Explicit

Public Function msgresult() As Integer

    msgresult = i
End Function

Public Sub CancelButton_Click()
    
    i = 0
   
    Unload Me
   
End Sub



Public Sub OKButton_Click()
    
    i = 1
    Unload Me
End Sub

