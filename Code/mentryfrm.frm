VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form meform 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Manual Entry"
   ClientHeight    =   11010
   ClientLeft      =   1065
   ClientTop       =   -570
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox opb 
      DataField       =   "OPB"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3120
      TabIndex        =   73
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox qtext 
      Height          =   375
      Left            =   480
      TabIndex        =   71
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox question 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   7680
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   70
      Top             =   2760
      Width           =   7095
   End
   Begin VB.CommandButton add 
      BackColor       =   &H00FFFF80&
      Caption         =   "Save This"
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   8760
      Width           =   2055
   End
   Begin VB.TextBox quetype 
      DataField       =   "QUETYPE"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   240
      TabIndex        =   68
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox level 
      DataField       =   "DIFF_LEVEL"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   67
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox ans 
      DataField       =   "ANS"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      TabIndex        =   66
      Top             =   6120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox opd 
      DataField       =   "OPD"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3120
      TabIndex        =   65
      Top             =   5640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox opc 
      DataField       =   "OBC"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3120
      TabIndex        =   64
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox opa 
      DataField       =   "OPA"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3120
      TabIndex        =   63
      Top             =   3840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox que 
      DataField       =   "QUE"
      DataSource      =   "Adodc1"
      Height          =   1335
      Left            =   2280
      TabIndex        =   62
      Top             =   2280
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton ADDNEW 
      BackColor       =   &H8000000D&
      Caption         =   "ADD NEW"
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   8160
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   240
      Top             =   1080
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;User ID=pro;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=pro;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "pro"
      Password        =   "pro"
      RecordSource    =   "select * from c"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame qframe 
      BackColor       =   &H8000000E&
      Caption         =   "Multiple Answers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8775
      Index           =   0
      Left            =   7200
      TabIndex        =   6
      Top             =   2040
      Width           =   8775
      Begin VB.TextBox mans 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   3
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   59
         Top             =   7320
         Width           =   6135
      End
      Begin VB.TextBox mans 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   2
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   58
         Top             =   6255
         Width           =   6135
      End
      Begin VB.TextBox mans 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   1
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   57
         Top             =   5175
         Width           =   6135
      End
      Begin VB.TextBox mans 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   0
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   4095
         Width           =   6135
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   10
         Top             =   4440
         Width           =   255
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   9
         Top             =   5400
         Width           =   255
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   960
         TabIndex        =   8
         Top             =   6480
         Width           =   255
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   960
         TabIndex        =   7
         Top             =   7560
         Width           =   255
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Enter Question Here"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   3615
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "Enter And Select Correct Answers"
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
         Left            =   240
         TabIndex        =   11
         Top             =   3480
         Width           =   4695
      End
   End
   Begin VB.Frame qframe 
      BackColor       =   &H8000000E&
      Caption         =   "Single Answer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Index           =   1
      Left            =   7200
      TabIndex        =   13
      Top             =   2040
      Width           =   8775
      Begin VB.OptionButton soptn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   960
         TabIndex        =   56
         Top             =   7680
         Width           =   255
      End
      Begin VB.OptionButton soptn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   960
         TabIndex        =   55
         Top             =   6600
         Width           =   255
      End
      Begin VB.OptionButton soptn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   54
         Top             =   5400
         Width           =   255
      End
      Begin VB.TextBox sans 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   3
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   7320
         Width           =   6135
      End
      Begin VB.TextBox sans 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   2
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   6240
         Width           =   6135
      End
      Begin VB.TextBox sans 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   1
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   5160
         Width           =   6135
      End
      Begin VB.TextBox sans 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   0
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   4080
         Width           =   6135
      End
      Begin VB.OptionButton soptn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   15
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Enter And Select Correct Answer"
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
         Left            =   240
         TabIndex        =   16
         Top             =   3480
         Width           =   5175
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Enter Question Here"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame qframe 
      BackColor       =   &H8000000E&
      Caption         =   "True Or False"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Index           =   2
      Left            =   7200
      TabIndex        =   17
      Top             =   2040
      Width           =   8775
      Begin VB.OptionButton tffalse 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "False"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3240
         TabIndex        =   2
         Top             =   6960
         Width           =   1575
      End
      Begin VB.OptionButton tftrue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "True"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   1
         Top             =   6120
         Width           =   1095
      End
      Begin VB.Image Image3 
         Height          =   5055
         Left            =   3720
         Picture         =   "mentryfrm.frx":0000
         Top             =   3360
         Width           =   4980
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Select Correct Answer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   19
         Top             =   4680
         Width           =   2895
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Enter Question Here"
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
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame qframe 
      BackColor       =   &H8000000E&
      Caption         =   "Fill in the Blanks"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Index           =   3
      Left            =   7200
      TabIndex        =   25
      Top             =   2040
      Width           =   8775
      Begin VB.TextBox fans 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   72
         Top             =   4320
         Width           =   6735
      End
      Begin VB.Image blank 
         Height          =   675
         Left            =   3360
         Picture         =   "mentryfrm.frx":385A
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Enter Correct Answer :"
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
         Left            =   360
         TabIndex        =   27
         Top             =   3480
         Width           =   5175
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000E&
         Caption         =   "Enter Question Here"
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
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame levelfrem 
      BackColor       =   &H8000000E&
      Height          =   855
      Left            =   7200
      TabIndex        =   20
      Top             =   960
      Width           =   8775
      Begin VB.OptionButton optlevel 
         BackColor       =   &H8000000E&
         Caption         =   "Low"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optlevel 
         BackColor       =   &H8000000E&
         Caption         =   "Midium"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   5280
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optlevel 
         BackColor       =   &H8000000E&
         Caption         =   "High"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   7440
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   2520
         X2              =   2520
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000E&
         Caption         =   "Select Level"
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
         Left            =   480
         TabIndex        =   21
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.ComboBox qtcombo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "mentryfrm.frx":509B
      Left            =   720
      List            =   "mentryfrm.frx":50AB
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   7440
      Width           =   5415
   End
   Begin VB.Frame qframe 
      BackColor       =   &H8000000E&
      Caption         =   "Match the Following"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Index           =   4
      Left            =   7200
      TabIndex        =   31
      Top             =   2040
      Width           =   8775
      Begin VB.TextBox match 
         Appearance      =   0  'Flat
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
         Index           =   3
         Left            =   6840
         TabIndex        =   53
         Top             =   7440
         Width           =   615
      End
      Begin VB.TextBox match 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   6840
         TabIndex        =   52
         Top             =   6615
         Width           =   615
      End
      Begin VB.TextBox match 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   6840
         TabIndex        =   51
         Top             =   5760
         Width           =   615
      End
      Begin VB.TextBox match 
         Appearance      =   0  'Flat
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
         Index           =   0
         Left            =   6840
         TabIndex        =   50
         Top             =   4920
         Width           =   615
      End
      Begin VB.TextBox mfans 
         Appearance      =   0  'Flat
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
         Index           =   7
         Left            =   3720
         TabIndex        =   49
         Top             =   7440
         Width           =   2295
      End
      Begin VB.TextBox mfans 
         Appearance      =   0  'Flat
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
         Index           =   6
         Left            =   3720
         TabIndex        =   47
         Top             =   6600
         Width           =   2295
      End
      Begin VB.TextBox mfans 
         Appearance      =   0  'Flat
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
         Index           =   5
         Left            =   3720
         TabIndex        =   45
         Top             =   5760
         Width           =   2295
      End
      Begin VB.TextBox mfans 
         Appearance      =   0  'Flat
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
         Index           =   4
         Left            =   3720
         TabIndex        =   43
         Top             =   4920
         Width           =   2295
      End
      Begin VB.TextBox mfans 
         Appearance      =   0  'Flat
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
         Index           =   3
         Left            =   720
         TabIndex        =   41
         Top             =   7440
         Width           =   2295
      End
      Begin VB.TextBox mfans 
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   720
         TabIndex        =   39
         Top             =   6600
         Width           =   2295
      End
      Begin VB.TextBox mfans 
         Appearance      =   0  'Flat
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
         Index           =   1
         Left            =   720
         TabIndex        =   37
         Top             =   5760
         Width           =   2295
      End
      Begin VB.TextBox mfans 
         Appearance      =   0  'Flat
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
         Index           =   0
         Left            =   720
         TabIndex        =   35
         Top             =   4920
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "D"
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
         Index           =   10
         Left            =   3240
         TabIndex        =   48
         Top             =   7440
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "C"
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
         Index           =   9
         Left            =   3240
         TabIndex        =   46
         Top             =   6600
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "B"
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
         Index           =   8
         Left            =   3240
         TabIndex        =   44
         Top             =   5760
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "A"
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
         Index           =   7
         Left            =   3240
         TabIndex        =   42
         Top             =   4920
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "4"
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
         Index           =   6
         Left            =   360
         TabIndex        =   40
         Top             =   7440
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "3"
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
         Index           =   5
         Left            =   360
         TabIndex        =   38
         Top             =   6600
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "2"
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
         Index           =   4
         Left            =   360
         TabIndex        =   36
         Top             =   5760
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "1"
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
         Index           =   3
         Left            =   360
         TabIndex        =   34
         Top             =   4920
         Width           =   255
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000E&
         Caption         =   "Enter Question Tag, If Any"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Enter Both Side And  Correct Answer "
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
         Index           =   2
         Left            =   240
         TabIndex        =   32
         Top             =   3960
         Width           =   5175
      End
   End
   Begin VB.Image deleteup 
      Height          =   720
      Left            =   3000
      Picture         =   "mentryfrm.frx":50F3
      Top             =   8400
      Width           =   720
   End
   Begin VB.Image deletebtn 
      Height          =   720
      Left            =   3000
      Picture         =   "mentryfrm.frx":62FC
      Top             =   8400
      Width           =   720
   End
   Begin VB.Image exitup 
      Height          =   480
      Left            =   1080
      Picture         =   "mentryfrm.frx":7505
      Top             =   8520
      Width           =   480
   End
   Begin VB.Line Line2 
      X1              =   3960
      X2              =   3960
      Y1              =   8280
      Y2              =   9240
   End
   Begin VB.Image prev 
      Height          =   720
      Left            =   4320
      Picture         =   "mentryfrm.frx":7B82
      ToolTipText     =   "Previous Question"
      Top             =   8400
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image nextbtn 
      Height          =   720
      Left            =   5280
      Picture         =   "mentryfrm.frx":81EA
      ToolTipText     =   "Next Question"
      Top             =   8400
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image exitbtn 
      Height          =   720
      Left            =   960
      Picture         =   "mentryfrm.frx":A22E
      ToolTipText     =   "Exit"
      Top             =   8400
      Width           =   720
   End
   Begin VB.Image back 
      Height          =   720
      Left            =   2040
      Picture         =   "mentryfrm.frx":AB1D
      ToolTipText     =   "Back to Login"
      Top             =   8400
      Width           =   720
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Teacher's Manual Question Entry"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   735
      Left            =   5040
      TabIndex        =   60
      Top             =   120
      Width           =   12375
   End
   Begin VB.Image Image2 
      Height          =   9540
      Left            =   16560
      Picture         =   "mentryfrm.frx":B17A
      Top             =   1440
      Width           =   10800
   End
   Begin VB.Image Image1 
      Height          =   4500
      Index           =   1
      Left            =   360
      Picture         =   "mentryfrm.frx":1560E
      Top             =   2040
      Width           =   6360
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Caption         =   "Select Question Type"
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
      Left            =   720
      TabIndex        =   24
      Top             =   6840
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   4545
      Index           =   0
      Left            =   240
      Picture         =   "mentryfrm.frx":1AD86
      Top             =   2040
      Width           =   6060
   End
   Begin VB.Image Image1 
      Height          =   4500
      Index           =   4
      Left            =   360
      Picture         =   "mentryfrm.frx":28866
      Top             =   2040
      Width           =   6150
   End
   Begin VB.Image Image1 
      Height          =   4980
      Index           =   3
      Left            =   240
      Picture         =   "mentryfrm.frx":31772
      Top             =   1680
      Width           =   6315
   End
   Begin VB.Image Image1 
      Height          =   4500
      Index           =   2
      Left            =   360
      Picture         =   "mentryfrm.frx":35A28
      Top             =   2040
      Width           =   6735
   End
End
Attribute VB_Name = "meform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim err As Boolean, optnb, ansb, bpress As Boolean



Private Sub add_Click()
       Adodc1.Recordset.ADDNEW
   ' validation
    err = False
        If qtext.Enabled = False Then
        MsgBox "Please Select Question Type", vbOKOnly, "ERROR"
        err = True
        Exit Sub
    End If
    If (meform.question.Text = "") Then
        MsgBox "Please Enter Question", vbOKOnly, "ERROR"

        err = True
        Exit Sub
    End If
    'Call Validate
    'List1.AddItem (qtext.Text)
    

    If meform.qtcombo.ListIndex = 0 Then
        Call Validfrm_mul
    End If
    If meform.qtcombo.ListIndex = 1 Then
        Call Validfrm_Sans
    End If
    If meform.qtcombo.ListIndex = 2 Then
        Call Validfrm_TorF
    End If
    If meform.qtcombo.ListIndex = 3 Then
        Call Validfrm_Fill
    End If
    If meform.qtcombo.ListIndex = 4 Then
        Call Validfrm_Match
    End If
    
    If (err = True) Then
        Exit Sub
    End If
   '----------------------------------------
    que.Text = question.Text
    opa.Text = ""
    opb.Text = ""
    opc.Text = ""
    opd.Text = ""
    level.Text = ""
    
    ans.Text = ""

'    Multiple Answers
'    Single Answer
'    True Or False
'    Fill in the Blanks
'    Match the Following
    i = 0
    While i < 3
        If (optlevel(i).Value = True) Then
                
                level.Text = str(i + 1)
        End If
        i = i + 1
    Wend

If (qtcombo.Text = "Multiple Answers") Then
        quetype = "2"
        opa.Text = mans(0).Text
        opb.Text = mans(1).Text
        opc.Text = mans(2).Text
        opd.Text = mans(3).Text
        
        i = 0
        While (i < 4)
            If (chk(i).Value = 1) Then
                    ans.Text = ans.Text + str(i + 1) + " "
            End If
            i = i + 1
            
        Wend
        'MsgBox ans.Text

End If
If (qtcombo.Text = "Single Answer") Then
        quetype = "1"
        opa.Text = sans(0).Text
        opb.Text = sans(1).Text
        opc.Text = sans(2).Text
        opd.Text = sans(3).Text
        
        i = 0
        While (i < 4)
            If (soptn(i).Value = True) Then
                    ans.Text = str(i + 1)
            End If
            i = i + 1
        Wend
        'MsgBox ans.Text
End If
If (qtcombo.Text = "True Or False") Then
        quetype = "3"
        
      
       
        If (tftrue.Value = True) Then
                  ans.Text = "1"
        
        End If
        If (tffalse.Value = True) Then
                  ans.Text = "2"
        End If
      '  MsgBox ans.Text

End If
If (qtcombo.Text = "Fill in the Blanks") Then
        quetype = "4"
        
        ans.Text = fans.Text
       ' MsgBox ans.Text

End If

    Adodc1.Recordset.MoveLast
    
    
    


End Sub

Private Sub ADDNEW_Click()
  
   question.Text = ""
   For i = 0 To 3
   mans(i).Text = ""
   sans(i).Text = ""
   match(i).Text = ""
   mfans(i).Text = ""
   Next
   For i = 4 To 7
   mfans(i).Text = ""
   Next

    
End Sub

Private Sub back_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
back.Left = back.Left - 300
End Sub

Private Sub back_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
back.Left = back.Left + 300
Unload Me
tlogin.Show
End Sub

Private Sub blank_Click()
question.Text = question.Text + "_____________"
bpress = True
question.SetFocus
question.SelStart = Len(question.Text)
question.SelLength = 0
End Sub

Private Sub Blankbtn_Click()
qtext.Text = question.Text + "_____________"
bpress = True
qtext.SetFocus
qtext.SelStart = Len(question.Text)
qtext.SelLength = 0
End Sub


Private Sub deletebtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
deletebtn.Visible = False
deleteup.Visible = True
End Sub

Private Sub deletebtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
deletebtn.Visible = True
deleteup.Visible = False
Adodc1.Recordset.delete
MsgBox " Question is deleted", vbOKOnly, "info"
End Sub


Private Sub exitbtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
exitup.Visible = True
exitbtn.Visible = False
End Sub

Private Sub exitbtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
exitup.Visible = False
exitbtn.Visible = True
End
End Sub

Private Sub Form_Click()
If qtext.Enabled = False Then
    MsgBox "Please Select Question Type", vbOKOnly, "ERROR"
End If

End Sub

Private Sub Form_Load()
deletebtn.Visible = True
deleteup.Visible = False
exitbtn.Visible = True
exitup.Visible = False
For i = 0 To qtcombo.ListCount - 1
    qframe(i).Enabled = False
Next
qtext.Enabled = False
meform.optlevel(0).Value = True
i = 0
For i = 0 To 2
    optlevel(i).Enabled = False
Next
Adodc1.RecordSource = "select * from " & tlogin.scombo.Text
Adodc1.Refresh
End Sub

Private Sub loaddata()

    question.Text = que.Text

        If (Val(level.Text) <> 0) Then
        
            optlevel(Val(level.Text) - 1).Value = True
        End If
        If (quetype.Text = "1") Then
        
         qtcombo.ListIndex = 1
    
         
         'MsgBox opa.Text
         
         sans(0).Text = opa.Text
        ' MsgBox sans(0).Text
         
         sans(1).Text = opb.Text
         sans(2).Text = opc.Text
         sans(3).Text = opd.Text
        
        i = Val(ans.Text)
        soptn(i - 1).Value = True
        'MsgBox ans.Text
ElseIf (quetype.Text = "2") Then
        qtcombo.ListIndex = 0
        
        'MsgBox opa.Text
        mans(0).Text = opa.Text
        mans(1).Text = opb.Text
        mans(2).Text = opc.Text
        mans(3).Text = opd.Text
        
        i = 1
       
        While (i < Len(ans.Text) - 1)
               If (Mid(ans.Text, i, 1) = " ") Then

                Else
                       chk(Val(Mid(ans.Text, i, 1))).Value = 1
                 
                End If
            i = i + 1

        Wend
        'MsgBox ans.Text

ElseIf (quetype.Text = "3") Then
         qtcombo.ListIndex = 2
       
      
       
        If (ans.Text = "1") Then
              tftrue.Value = True
        
        End If
        If (ans.Text = "2") Then
                tffalse.Value = True
        End If
      '  MsgBox ans.Text

ElseIf (quetype.Text = "4") Then
      qtcombo.ListIndex = 3
          
        fans.Text = ans.Text
       ' MsgBox ans.Text

End If


End Sub

Private Sub nextbtn_Click()

 If (Adodc1.Recordset.EOF = False) Then
    Adodc1.Recordset.MoveNext
    Call loaddata
End If

End Sub
Private Sub Validfrm_TorF()
If (tftrue.Value = False And tffalse.Value = False) Then
    MsgBox "Please Select Answer", vbOKOnly, "ERROR"
End If

End Sub
Private Sub Validfrm_mul()
Dim mcnt As Integer
ansb = False
optnb = True
mcnt = 0
For i = 0 To 3
    If mans(i) = "" Then
    ansb = True
    End If
    If chk(i).Value = 1 Then
    mcnt = mcnt + 1
    End If
Next
If ansb = True Then
    MsgBox "Please Enter All Options", vbOKOnly, "ERROR"

    err = True
ElseIf mcnt = 0 Then
    MsgBox "Please Select Answer", vbOKOnly, "ERROR"

    err = True
ElseIf mcnt = 1 Then
    MsgBox "Please Select Multiple Choices Or Switch to Single Answer type", vbOKOnly, "ERROR"

    err = True
End If

End Sub
Private Sub Validfrm_Fill()

If (fans.Text = "") Then
    MsgBox "Please Enter the correct answer", vbOKOnly, "ERROR"
End If

End Sub
Private Sub Validfrm_Match()
Dim mtch As Integer, mnum As Boolean
mnum = False
ansb = False
optnb = False
For i = 0 To 7
    If mfans(i) = "" Then
    ansb = True
    End If
Next
For i = 0 To 3
    If match(i) = "" Then
    optnb = True
    End If
Next
If ansb = True Then
MsgBox "Please Enter All Elements of Both Sides", vbOKOnly, "ERROR"

err = True
End If
If optnb = True Then
MsgBox "Please Enter All Matches", vbOKOnly, "ERROR"

err = True
Exit Sub
End If
For i = 0 To 3
mtch = Val(match(i))
If mtch <> 1 And mtch <> 2 And mtch <> 3 And mtch <> 4 Then
'mnum = True
'End If
MsgBox "Please Enter All Matches Properly", vbOKOnly, "ERROR"

err = True
Exit Sub
End If
Next
For i = 0 To 3
j = Len(match(i))
If j > 1 Then
MsgBox "Please Enter All Matches Properly", vbOKOnly, "ERROR"

err = True
Exit Sub
End If
Next
For i = 0 To 3
mtch = Val(match(i))
    For j = 0 To 3
        If mtch = Val(match(j)) And i <> j Then
        MsgBox "Two matches cannot be same", vbOKOnly, "ERROR"

        err = True
        Exit Sub
        End If
    Next
Next
'If mnum = False Then
'MsgBox "Please Enter All Matches Properly"
'End If

End Sub
Private Sub Validfrm_Sans()
ansb = False
optnb = True
For i = 0 To 3
    If sans(i) = "" Then
    ansb = True
    End If
    If soptn(i).Value = True Then
    optnb = False
    End If
Next
If ansb = True Then
MsgBox "Please Enter All Options", vbOKOnly, "ERROR"

err = True
End If
If optnb = True Then
MsgBox "Please Select Answer", vbOKOnly, "ERROR"

err = True
End If
End Sub

Private Sub nextbtn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'nextbtn.Left = nextbtn.Left + 300
End Sub

Private Sub nextbtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'nextbtn.Left = nextbtn.Left - 300
End Sub

Private Sub prev_Click()
If (Adodc1.Recordset.BOF = False) Then
    Adodc1.Recordset.MovePrevious
    Call loaddata

End If
End Sub

Private Sub prev_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'prev.Left = prev.Left - 300
End Sub

Private Sub prev_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'prev.Left = prev.Left + 300
End Sub

Private Sub qtcombo_Click()
For i = 0 To qtcombo.ListCount - 1
    qframe(i).Enabled = True
Next
For i = 0 To 2
    optlevel(i).Enabled = True
Next
qtext.Enabled = True
meform.qframe(qtcombo.ListIndex).ZOrder
meform.Image1(qtcombo.ListIndex).ZOrder
meform.qtext.ZOrder
bpress = False
question.ZOrder
End Sub

