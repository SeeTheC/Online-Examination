VERSION 5.00
Begin VB.Form overview 
   BackColor       =   &H8000000E&
   Caption         =   "Form2"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   870
   ClientWidth     =   9330
   FillColor       =   &H000000C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form2"
   ScaleHeight     =   11040
   ScaleWidth      =   20280
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton finish 
      BackColor       =   &H80000010&
      Height          =   1095
      Left            =   15120
      Picture         =   "overview.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   2520
      TabIndex        =   0
      Top             =   2880
      Width           =   14295
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   8
         Left            =   10440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   7
         Left            =   10440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   5280
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   6
         Left            =   10440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   3840
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   5
         Left            =   10440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   2400
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   4
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Index           =   3
         Left            =   3480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   3720
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   2
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   1
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   0
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "-ve Marking :"
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   8
         Left            =   7800
         TabIndex        =   18
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   7560
         X2              =   7560
         Y1              =   120
         Y2              =   6720
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Hard Level  :"
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   7
         Left            =   7800
         TabIndex        =   15
         Top             =   5280
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Medium Level:"
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   6
         Left            =   7800
         TabIndex        =   13
         Top             =   3840
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Low Level :"
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   5
         Left            =   7800
         TabIndex        =   11
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Marks :"
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   4
         Left            =   7800
         TabIndex        =   9
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Question  Types :"
         ForeColor       =   &H00000080&
         Height          =   855
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   3720
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Duration :"
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   2
         Left            =   1560
         TabIndex        =   5
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject :"
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   1
         Left            =   1680
         TabIndex        =   3
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TYPE :"
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   0
         Left            =   2040
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "DONE"
      Height          =   495
      Left            =   15240
      TabIndex        =   20
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Image back 
      Height          =   1200
      Index           =   2
      Left            =   2520
      Picture         =   "overview.frx":18E4
      Top             =   10080
      Width           =   2745
   End
   Begin VB.Image Image2 
      Height          =   3090
      Left            =   2400
      Picture         =   "overview.frx":2C2B
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   8160
      Left            =   2280
      Picture         =   "overview.frx":2FEC5
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   14955
   End
End
Attribute VB_Name = "overview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command
Dim temp
Dim qflag As Boolean
Dim counter As Integer
Private Sub back_Click(Index As Integer)
    Me.Hide
    instruction.Visible = True
End Sub

Private Sub excquery(ByVal query As String)
    
    With cmd
        On Error GoTo Exit1
        .CommandText = query
        .ActiveConnection = cn
        Set rs = .Execute
       counter = counter + 1
        Exit Sub
    End With
Exit1:
    'MsgBox "Error in Excecuting query"
    qflag = True
End Sub

Private Sub finish_Click()
    Dim str1 As String
    Dim i
    Dim que_type, name, diff_level
    Dim query As String
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    str1 = "Provider=MSDAORA.1" ';User ID=pro;Password=pro"
    cn.Open str1, "pro", "pro"
    
    query = "select count(*) from examsetting"
    Call excquery(query)
    
        
    If (rs(0) <> 0) Then
        query = "select max(sno) from examsetting"
        Call excquery(query)
    
        counter = rs(0)
    Else
        counter = 0
    End If
    'MsgBox counter
    If (counter <= 0) Then
        counter = 1
    Else
        counter = counter + 1
    End If
   

    With tc1                                                                                   'catagory               'subject                'duration               'totalmarks             'negative
    query = "insert into examsetting values('" + str(counter) + "',current_date,'" + Text1(0).Text + "','" + Text1(1).Text + "','" + Text1(2).Text + "','" + Text1(4).Text + "','" + Text1(8).Text + "','" + .nq(0).Text + "','" + .nm(0).Text + "','" + .nq(1).Text + "','" + .nm(1).Text + "','" + .nq(2).Text + "','" + .nm(2).Text + "','" + .quetype.Text + "','" + instruction.final_ins.Text + "' )"
    Call excquery(query)

    End With
    MsgBox "SETTING HAS BEEN SAVED"
    first.Show
    Unload Me
End Sub

Private Sub Form_Load()
Dim i
With tc1
        Text1(0).Text = .year.Text 'type
        Text1(1).Text = .subject.Text 'subject
        Text1(2).Text = .hr.Text + " : " + .min.Text  ' duration
        i = 0
        While i < 5
            If (.que(i).Value = 1) Then
                Text1(3).Text = Text1(3).Text + "" + str(i) + " ) " + .que(i).Caption + vbCrLf
            End If
             i = i + 1
        Wend
        Text1(4).Text = .Tmarks.Text
        Text1(5).Text = " 1) No. of questions :" + .nq(0).Text + vbCrLf + " 2) Mark for each question  :" + .nm(0).Text
        Text1(6).Text = " 1) No. of questions :" + .nq(1).Text + vbCrLf + " 2) Mark for each question  :" + .nm(1).Text
        Text1(7).Text = " 1) No. of questions :" + .nq(2).Text + vbCrLf + " 2) Mark for each question  :" + .nm(2).Text
        
        Text1(8).Text = .negativemarks.Text
        i = 0
        While i < 3
        
            If (.level(i).Value = 0) Then
                   Text1(5 + i).Enabled = False
            End If
            i = i + 1
        Wend
        
        
        If (.Option1.Value = False) Then
                Text1(8).Enabled = False
                Text1(8).Text = 0
        End If
End With

End Sub

