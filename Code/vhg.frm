VERSION 5.00
Begin VB.Form tlogin 
   BackColor       =   &H8000000E&
   Caption         =   "Teacher Login For Manual Entry"
   ClientHeight    =   8880
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   12240
      TabIndex        =   17
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox addcat 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7080
      TabIndex        =   16
      Top             =   4320
      Width           =   3975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11160
      TabIndex        =   14
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Add New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11160
      TabIndex        =   15
      Top             =   4320
      Width           =   1335
   End
   Begin VB.ComboBox ccombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7080
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   4320
      Width           =   3975
   End
   Begin VB.CommandButton sql 
      Caption         =   "refresh"
      Height          =   735
      Left            =   240
      TabIndex        =   11
      Top             =   7920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.OptionButton i_f 
      BackColor       =   &H8000000E&
      Caption         =   "Import From File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7920
      TabIndex        =   10
      Top             =   6120
      Width           =   2535
   End
   Begin VB.OptionButton m_e 
      BackColor       =   &H8000000E&
      Caption         =   "Manual Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7920
      TabIndex        =   9
      Top             =   5640
      Value           =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   12240
      TabIndex        =   7
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11160
      TabIndex        =   6
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox addsub 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7080
      TabIndex        =   5
      Top             =   5040
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Add New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11160
      TabIndex        =   4
      Top             =   5040
      Width           =   1335
   End
   Begin VB.ComboBox scombo 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7080
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   5040
      Width           =   3975
   End
   Begin VB.TextBox tname 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7080
      TabIndex        =   2
      Top             =   3360
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Select Catagory:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   12
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Image exitup 
      Height          =   630
      Left            =   12120
      Picture         =   "vhg.frx":0000
      Top             =   1600
      Width           =   630
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "Select Way of Question Entry :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   5640
      Width           =   3975
   End
   Begin VB.Image exitbtn 
      Height          =   960
      Left            =   12000
      Picture         =   "vhg.frx":050B
      ToolTipText     =   "Exit"
      Top             =   1440
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   4650
      Left            =   14640
      Picture         =   "vhg.frx":0BF9
      Top             =   5520
      Width           =   6150
   End
   Begin VB.Image login 
      Height          =   495
      Left            =   8160
      Picture         =   "vhg.frx":D313
      ToolTipText     =   "Login "
      Top             =   7200
      Width           =   1380
   End
   Begin VB.Image Image3 
      Height          =   3705
      Left            =   -240
      Picture         =   "vhg.frx":DA95
      Top             =   -600
      Width           =   4980
   End
   Begin VB.Image Image1 
      Height          =   4095
      Left            =   13080
      Picture         =   "vhg.frx":10C6D
      Top             =   1440
      Width           =   4290
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Select Subject :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   1
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Enter Teacher's Name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   3360
      Width           =   3375
   End
End
Attribute VB_Name = "tlogin"
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
'---------------------------
Private Sub Command1_Click()
meform.Show
tlogin.Hide
End Sub

Private Sub ccombo_Click()
    scombo.Enabled = True
    query = "select catagory,subject from tabledetail"
    Call excquery(query)
    scombo.Clear
    While (rs.EOF = False)
            If (ccombo.Text = rs(0)) Then
                  scombo.AddItem (rs(1))
            End If
                rs.MoveNext
        Wend
End Sub

Private Sub Command2_Click()
scombo.Visible = False
Command2.Visible = False
Command3.Visible = True
Command4.Visible = True
addsub.Visible = True
addsub.SetFocus
End Sub

Private Sub createtable()
Dim flag As Boolean
Dim i
i = 0
flag = False
While i < scombo.ListCount
        
        If (addsub.Text = scombo.List(i)) Then
                MsgBox " Can not add subject. As it is already Present"
                flag = True
        End If
        i = 1 + i
Wend
If (flag = False) Then

    query = "create table " + addsub.Text + " (qno integer,quetype varchar(200),name varchar(100),que varchar(1000),opa varchar(100),opb varchar(100),obc varchar(100),opd varchar(100),ans varchar(20),diff_level integer)"
    Call excquery(query)
    query = "insert into tabledetail values('" + ccombo.Text + "','" + addsub.Text + " ')"
    Call excquery(query)
    Call sql_Click
    'MsgBox "table created"
End If
End Sub
Private Sub Command3_Click()
scombo.Visible = True
Command2.Visible = True
Command3.Visible = False
Command4.Visible = False
addsub.Visible = False
Call createtable

End Sub

Private Sub Command4_Click()
scombo.Visible = True
Command2.Visible = True
Command3.Visible = False
addsub.Visible = False
Command4.Visible = False
End Sub





Private Sub exit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


End Sub



Private Sub Command5_Click()
ccombo.Visible = False
Command5.Visible = False
Command6.Visible = True
Command7.Visible = True
addcat.Visible = True
addcat.SetFocus
End Sub

Private Sub Command6_Click()
ccombo.Visible = True
Command5.Visible = True
Command6.Visible = False
Command7.Visible = False
addcat.Visible = False
ccombo.AddItem (addcat.Text)
ccombo.ListIndex = ccombo.ListCount - 1
'Call createtable
End Sub

Private Sub Command7_Click()
ccombo.Visible = True
Command5.Visible = True
Command6.Visible = False
addcat.Visible = False
Command7.Visible = False
End Sub

Private Sub exitbtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
exitbtn.Visible = False
exitup.Visible = True
End Sub

Private Sub exitbtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
exitbtn.Visible = True
exitup.Visible = False
teachers.Show
Unload Me
End Sub



Private Sub Form_Load()
'addsub.SetFocus
exitup.Visible = False
exitbtn.Visible = True
scombo.Visible = True
Command2.Visible = True
Command3.Visible = False
Command4.Visible = False
addsub.Visible = False

ccombo.Visible = True
Command5.Visible = True
Command6.Visible = False
Command7.Visible = False
addcat.Visible = False
Call sql_Click
End Sub

Private Sub login_Click()
Dim flag As Boolean
flag = False
If tname.Text = "" Then
MsgBox "Please enter Name"
flag = True
End If
If scombo.Text = "" Then
MsgBox "Please select Subject Name"
flag = True
End If
If m_e.Value = False And i_f.Value = False Then
MsgBox "Please select Way of entering Questions"
flag = True
End If
If flag = True Then
Exit Sub
End If


tlogin.Hide

If (m_e.Value = True) Then
    meform.Show
Else
    insert_que_by_file.Show
End If

End Sub
Private Sub excquery(ByVal query As String)
    
    With cmd
        'On Error GoTo Exit1
            .CommandText = query
            .ActiveConnection = cn
            Set rs = .Execute
       counter = counter + 1
        Exit Sub
    End With
Exit1:
   MsgBox "Error in Excecuting query"
    qflag = True
End Sub

Private Sub sql_Click()
    Dim str1 As String
    Dim i
    Dim query As String
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    str1 = "Provider=MSDAORA.1" ';User ID=pro;Password=pro"
    cn.Open str1, "pro", "pro"
    
    query = "select count(*) from tabledetail"
    Call excquery(query)
    'MsgBox rs(0)
        
    If (rs(0) <> 0) Then
        
        ' catagory
        query = "select distinct catagory from tabledetail"
        Call excquery(query)
        ccombo.Clear
        While (rs.EOF = False)
                ccombo.AddItem (rs(0))
                rs.MoveNext
        Wend
        
'        query = "select subject from tabledetail"
'        Call excquery(query)
'        scombo.Clear
'        While (rs.EOF = False)
'                scombo.AddItem (rs(0))
'                rs.MoveNext
'        Wend
'
        
        
    Else
        counter = 0
    End If
End Sub
