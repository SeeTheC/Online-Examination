VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form insert_que_by_file 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13950
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9990
   ScaleWidth      =   13950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2880
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8535
      Left            =   3600
      TabIndex        =   0
      Top             =   1560
      Width           =   9975
      Begin VB.CommandButton done 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Done"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   8040
         Width           =   1935
      End
      Begin VB.TextBox example 
         Appearance      =   0  'Flat
         Height          =   1935
         Left            =   1800
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "insert_que_by_file.frx":0000
         Top             =   6120
         Width           =   5775
      End
      Begin VB.ComboBox options 
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   4680
         Width           =   2535
      End
      Begin VB.ComboBox question 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   390
         ItemData        =   "insert_que_by_file.frx":0008
         Left            =   2400
         List            =   "insert_que_by_file.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3720
         Width           =   2535
      End
      Begin VB.CommandButton browse 
         Caption         =   "Browse"
         Height          =   390
         Left            =   7200
         TabIndex        =   3
         Top             =   1905
         Width           =   1215
      End
      Begin VB.TextBox path 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   2
         Text            =   "c:\"
         Top             =   1920
         Width           =   4695
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EXAMPLE :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   840
         TabIndex        =   10
         Top             =   5520
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Options :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   840
         TabIndex        =   8
         Top             =   4680
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Question :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   720
         TabIndex        =   6
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Signature :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Path :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select the file :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   2655
      End
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   480
      Picture         =   "insert_que_by_file.frx":000C
      Top             =   360
      Width           =   2250
   End
End
Attribute VB_Name = "insert_que_by_file"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim eg As String
Private Sub Dir1_Change()
File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
 Dir1.path = Drive1.Drive
End Sub

Private Sub init()
    Dim que(50) As String, op(50) As String
    Dim n As Integer, m As Integer
    Dim i
    'n:  for no. of que
    'm: for no. of op
    n = 5
    que(0) = "Q. 1 ."
    que(1) = "Q . 1 )"
    que(2) = "Q 1."
    que(3) = "Q 1)"
    que(4) = "1 ."
    que(5) = "1 )"
    i = 0
    While i < n
        
        question.AddItem (que(i))
        i = i + 1
    Wend
    
    m = 4
    op(0) = "a ."
    op(1) = "a )"
    op(2) = "A ."
    op(3) = "A )"
    op(4) = "1 ."
    op(5) = "1 )"
    i = 0
    While i < m
        
        options.AddItem (op(i))
        i = i + 1
    Wend


End Sub

Private Sub Command1_Click()

End Sub

Private Sub browse_Click()
 cd1.Filter = "text file|*.txt"
 cd1.ShowOpen
 
 If (Len(cd1.filename) <> 0) Then
    path.Text = cd1.filename
 End If
End Sub

Private Sub chkdone()

    If (Len(question.Text) <> 0 And Len(options.Text) <> 0 And Len(path.Text) <> 0) Then
    
            done.Enabled = True
    Else
            done.Enabled = False
    End If
End Sub
Private Sub done_Click()
    
    If (Len(question.Text) <> 0 And Len(options.Text) <> 0 And Len(path.Text) <> 0) Then
    
        Me.Hide
        que_file.Show
    End If
End Sub

Private Sub Form_Load()
    Call init

End Sub

Private Sub setexample()


    Dim str1
    
    If (Len(question.Text) <> 0) Then
 
            eg = question.Text + " What is the capital of INDIA "
    End If
    If (Len(options.Text) <> 0) Then
    str1 = Mid(options.Text, 3, 1)
    
    If (Asc(Mid(options.Text, 1, 1)) = 49) Then
        eg = eg + vbCrLf + "1" + str1 + " MUMBAI"
        eg = eg + vbCrLf + "2" + str1 + " NEW DELHI"
        eg = eg + vbCrLf + "3" + str1 + " CHENNAI"
        eg = eg + vbCrLf + "4" + str1 + " PUNE"
    End If
    If (Mid(options.Text, 1, 1) = "a") Then
        eg = eg + vbCrLf + "a" + str1 + " MUMBAI"
        eg = eg + vbCrLf + "b" + str1 + " NEW DELHI"
        eg = eg + vbCrLf + "c" + str1 + " CHENNAI"
        eg = eg + vbCrLf + "d" + str1 + " PUNE"
    End If
    If (Mid(options.Text, 1, 1) = "A") Then
        eg = eg + vbCrLf + "A" + str1 + " MUMBAI"
        eg = eg + vbCrLf + "B" + str1 + " NEW DELHI"
        eg = eg + vbCrLf + "C" + str1 + " CHENNAI"
        eg = eg + vbCrLf + "D" + str1 + " PUNE"
    End If
    
End If

End Sub
Private Sub options_Click()
 
    Call setexample
    Call chkdone
    If (Len(question.Text) <> 0 And Len(options.Text) <> 0) Then
        example.Text = eg
    End If
 
End Sub

Private Sub question_Click()
     
    Call setexample
    Call chkdone
    
    If (Len(question.Text) <> 0 And Len(options.Text) <> 0) Then
        example.Text = eg
    End If
 
End Sub
