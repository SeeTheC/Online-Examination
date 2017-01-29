VERSION 5.00
Begin VB.Form instruction 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Instruction"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14895
   FillColor       =   &H00FF0000&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   14895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox final_ins 
      Height          =   1095
      Left            =   4320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Top             =   9240
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9495
      Left            =   -120
      TabIndex        =   2
      Top             =   -480
      Width           =   17535
      Begin VB.CommandButton delete 
         BackColor       =   &H0000FFFF&
         Caption         =   "DELETE"
         Height          =   375
         Left            =   15600
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6120
         Width           =   1095
      End
      Begin VB.CommandButton ADD 
         BackColor       =   &H0000FFFF&
         Caption         =   "ADD"
         Height          =   375
         Left            =   14040
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6120
         Width           =   1215
      End
      Begin VB.ListBox inslist 
         BackColor       =   &H0080FFFF&
         Height          =   3120
         Left            =   12720
         TabIndex        =   4
         Top             =   2640
         Width           =   4095
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   5775
         LargeChange     =   50
         Left            =   9600
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Top             =   2760
         Width           =   375
      End
      Begin VB.PictureBox P1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5775
         Left            =   1440
         ScaleHeight     =   5775
         ScaleWidth      =   8295
         TabIndex        =   9
         Top             =   2760
         Width           =   8295
         Begin VB.CommandButton ADD1 
            BackColor       =   &H000080FF&
            Caption         =   "ADD"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   840
            MaskColor       =   &H0080FFFF&
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox serial 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   0
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   1
            Text            =   "1)"
            Top             =   0
            Width           =   495
         End
         Begin VB.TextBox ins1 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   525
            Index           =   0
            Left            =   720
            MultiLine       =   -1  'True
            TabIndex        =   0
            Top             =   0
            Width           =   7455
         End
         Begin VB.Line vline1 
            BorderColor     =   &H8000000D&
            BorderWidth     =   2
            X1              =   720
            X2              =   720
            Y1              =   0
            Y2              =   6480
         End
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Instructions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   10
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000D&
         BorderWidth     =   2
         X1              =   2160
         X2              =   2160
         Y1              =   1920
         Y2              =   3720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   2
         X1              =   1440
         X2              =   9960
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   1440
         TabIndex        =   6
         Top             =   1920
         Width           =   8535
      End
      Begin VB.Image Image2 
         Height          =   4425
         Left            =   12600
         Picture         =   "instruction.frx":0000
         Top             =   2520
         Width           =   4410
      End
      Begin VB.Image Image1 
         Height          =   9000
         Left            =   0
         Picture         =   "instruction.frx":1A9A
         Top             =   720
         Width           =   10800
      End
   End
   Begin VB.Image back 
      Height          =   1200
      Index           =   2
      Left            =   840
      Picture         =   "instruction.frx":11EDB
      Top             =   9240
      Width           =   2745
   End
   Begin VB.Image next 
      Height          =   1290
      Index           =   2
      Left            =   11040
      Picture         =   "instruction.frx":13222
      Top             =   9120
      Width           =   2835
   End
End
Attribute VB_Name = "instruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim size
Dim prevvalue, prev_add_position As Integer
Dim ins1count As Integer
Dim keypressed As Boolean
Private Const scalling As Integer = 100
Dim length

Private Sub insertintolist()
Dim i
i = 0
inslist.Clear
While i < ins1count
    inslist.AddItem ((i + 1) & ") " & ins1(i).Text)
    i = i + 1
Wend

End Sub


Private Sub add_Click()
       
      Dim str1 As String
     
      If (Len(ins1(ins1count - 1).Text) = 0 Or Len(Trim(ins1(ins1count - 1).Text)) = 0) Then
            errormsg.msg.Text = " Instruction " + str(ins1count) + " is EMPTY ."
            errormsg.Show 1
            Exit Sub
              
      End If
      Load ins1(ins1count)
      Load serial(ins1count)
      
      ins1(ins1count).Visible = True
      ins1(ins1count).TabIndex = ins1(ins1count - 1).TabIndex + 1
      ADD1.TabIndex = ins1count
      ADD1.TabStop = True
      
      serial(ins1count).Visible = True
      str1 = ins1count + 1
      serial(ins1count).Text = " " + str1 + ")"
      
      
      ins1(ins1count).top = ins1(ins1count - 1).top + ins1(ins1count - 1).Height + 20
       
      ins1(ins1count).Height = ins1(ins1count).FontSize * 25
      serial(ins1count).top = ins1(ins1count).top
      ins1(ins1count).Text = ""
      ins1count = ins1count + 1

      
      Call ins1_Change(ins1count)
      'Call insertintolist

End Sub



Private Sub ADD1_Click()
 Call add_Click
End Sub



Private Sub back_Click(index As Integer)
tc1.step(2).Visible = True
Me.Hide
tc1.Visible = True
End Sub

Private Sub delete_Click()

    
    If inslist.ListIndex >= 0 Then
        inslist.RemoveItem (inslist.ListIndex)
        Call remove
    Else
        If inslist.ListCount > 0 Then
            inslist.RemoveItem (inslist.ListCount - 1)
            Call remove
        End If
    End If
    
    Call ins1_Change(ins1count)

    
End Sub

Private Sub remove()
      If ins1count <> 1 Then
    
        Unload ins1(ins1count - 1)
        Unload serial(ins1count - 1)
    
        ins1count = ins1count - 1
    End If
   
    '===========copying from list to text file
    Dim str()  As String
    Dim liststr(500) As String
    Dim i, j, n
    i = 0
    n = inslist.ListCount
  
    While i < n
        
        
        str = Split(inslist.List(i), ") ", 4)
        liststr(i) = str(1)
        i = i + 1
    Wend
    i = 0
    While i < n
        ins1(i).Text = liststr(i)
        i = i + 1
    Wend
  
     
End Sub



Private Sub Form_Load()
    Call initvscroll
    prevvalue = 0
    ins1count = 1
    length = ins1(0).Height + ADD1.Height
  
End Sub

Private Sub findmax()
    Dim i
    Dim gap
    i = 0
    length = 0
    gap = 20
    While (i < ins1.count)
                
        length = length + ins1(i).Height + gap ' 20 is for gap between 2 ins.
        i = i + 1
    Wend
    length = length + ADD1.Height + gap
    
    ' extra length
End Sub
Private Sub ins1_Change(index As Integer)

    Dim i, no
    Dim new_add_pos As Integer
    i = index + 1
    no = ins1.count
    While i < no
            
            ins1(i).Move ins1(i).Left, (ins1(i - 1).top + ins1(i - 1).Height + 20)
            serial(i).Move serial(i).Left, (ins1(i - 1).top + ins1(i - 1).Height + 20)
            i = i + 1
    Wend
    '=======add button
    ADD1.Move ADD1.Left, (ins1(no - 1).top + ins1(no - 1).Height + 20)
            
    '===========SCROLL BAR CONDITION====================
     Call findmax
     Dim temp As Integer
     If (length > P1.Height) Then
        length = length - P1.Height
        temp = (length / scalling)
        
        VScroll1.max = temp * -1
        VScroll1.min = 0
     End If
     
Call insertintolist
End Sub

Private Sub ins1_KeyPress(index As Integer, KeyAscii As Integer)
    Dim intialheight
    Dim tempselstart
    Dim i, top
    If KeyAscii = 13 Then
        intialheight = (ins1(index).FontSize * 25)
        ins1(index).Height = ins1(index).Height + intialheight
        
    End If
    If KeyAscii = 8 Then
    
        If ins1(index).SelStart <> 0 Then
               
         tempselstart = ins1(index).SelStart
       
        ins1(index).SelStart = ins1(index).SelStart - 1
        ins1(index).SelLength = 1
            If (Asc(ins1(index).SelText) = 10) Then
                ins1(index).SelText = vbNullChar
              intialheight = (ins1(index).FontSize * 25)
              ins1(index).Height = ins1(index).Height - intialheight
                       
            End If
       
        End If
    End If
    
    keypressed = True
    
End Sub


Private Sub initvscroll()
    VScroll1.min = 0
    VScroll1.max = 0
    no = ins1.count
    

End Sub




Private Sub next_Click(index As Integer)
 Dim i
 i = 0
 final_ins = ""
 While i < ins1.count
        final_ins.Text = final_ins.Text + str(i + 1) + " . " + ins1(i).Text + vbCrLf
        i = i + 1
 Wend
 
 Me.Hide
 overview.Visible = True
End Sub

Private Sub VScroll1_Change()
    
    
    Dim i, no, size
    no = ins1.count
    i = 0
  
    While i < no
            
            ins1(i).Move ins1(i).Left, (ins1(i).top + (VScroll1.Value - prevvalue) * scalling)
            serial(i).Move serial(i).Left, (serial(i).top + (VScroll1.Value - prevvalue) * scalling)
            i = i + 1
        
    Wend
    ADD1.Move ADD1.Left, (ADD1.top + (VScroll1.Value - prevvalue) * scalling)
    'ADD1.Move ADD1.Left, (ins1(no - 1).Top + ins1(no - 1).Height + 20)
    prevvalue = VScroll1.Value
finish:
End Sub

Private Sub VScroll1_GotFocus()
ADD1.SetFocus
End Sub

Private Sub VScroll1_LostFocus()
Dim i
End Sub

