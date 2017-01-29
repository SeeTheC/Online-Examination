VERSION 5.00
Begin VB.Form mousepos 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   1560
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "mousepos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Interval = 100

End Sub

Private Sub Timer1_Timer()
 Dim rect As POINTAPI
      ' Get the current mouse cursor coordinates:
      Call GetCursorPos(rect)
      mousepos.Cls
      ' Print out current position on the form:
      Print "Current X = " & rect.X
      Print "Current Y = " & rect.Y
End Sub
