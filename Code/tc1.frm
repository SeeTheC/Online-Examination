VERSION 5.00
Begin VB.Form tc1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   12435
   FillColor       =   &H8000000D&
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   12435
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton initsub 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1440
      TabIndex        =   44
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame step 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Step3"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   12615
      Index           =   2
      Left            =   6000
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   9615
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   3975
         Left            =   2280
         TabIndex        =   24
         Top             =   2640
         Width           =   6855
         Begin VB.TextBox nq 
            Alignment       =   1  'Right Justify
            Height          =   510
            Index           =   0
            Left            =   2760
            TabIndex        =   37
            Text            =   "0"
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox total 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   36
            Text            =   "0"
            Top             =   2640
            Width           =   1095
         End
         Begin VB.TextBox total 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   35
            Text            =   "0"
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox total 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   34
            Text            =   "0"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox nm 
            Alignment       =   1  'Right Justify
            Height          =   465
            Index           =   2
            Left            =   4200
            TabIndex        =   33
            Text            =   "0"
            Top             =   2520
            Width           =   855
         End
         Begin VB.TextBox nm 
            Alignment       =   1  'Right Justify
            Height          =   465
            Index           =   1
            Left            =   2760
            TabIndex        =   32
            Text            =   "0"
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox nq 
            Alignment       =   1  'Right Justify
            Height          =   510
            Index           =   2
            Left            =   2760
            TabIndex        =   31
            Text            =   "0"
            Top             =   2520
            Width           =   855
         End
         Begin VB.TextBox nq 
            Alignment       =   1  'Right Justify
            Height          =   510
            Index           =   1
            Left            =   4200
            TabIndex        =   30
            Text            =   "0"
            Top             =   1800
            Width           =   855
         End
         Begin VB.CheckBox level 
            Appearance      =   0  'Flat
            BackColor       =   &H80000009&
            Caption         =   "Hard"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   510
            Index           =   2
            Left            =   360
            TabIndex        =   29
            Top             =   2520
            Width           =   1935
         End
         Begin VB.CheckBox level 
            Appearance      =   0  'Flat
            BackColor       =   &H80000009&
            Caption         =   "Medium"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   570
            Index           =   1
            Left            =   360
            TabIndex        =   28
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox nm 
            Alignment       =   1  'Right Justify
            Height          =   465
            Index           =   0
            Left            =   4200
            TabIndex        =   27
            Text            =   "0"
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox level 
            Appearance      =   0  'Flat
            BackColor       =   &H80000009&
            Caption         =   "Low"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   510
            Index           =   0
            Left            =   360
            TabIndex        =   26
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox gtotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   25
            Text            =   "0"
            Top             =   3240
            Width           =   1095
         End
         Begin VB.Line Line12 
            X1              =   240
            X2              =   6480
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Line Line11 
            X1              =   240
            X2              =   6600
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Line Line5 
            BorderWidth     =   3
            X1              =   240
            X2              =   240
            Y1              =   360
            Y2              =   3120
         End
         Begin VB.Line Line6 
            BorderWidth     =   3
            X1              =   6600
            X2              =   6600
            Y1              =   360
            Y2              =   3720
         End
         Begin VB.Line Line3 
            BorderWidth     =   2
            X1              =   3720
            X2              =   3720
            Y1              =   360
            Y2              =   3120
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000008&
            BackStyle       =   0  'Transparent
            Caption         =   "NO.   of  Que."
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   735
            Index           =   0
            Left            =   2520
            TabIndex        =   42
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Each que marks"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   735
            Index           =   0
            Left            =   3600
            TabIndex        =   41
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   " Total"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   375
            Index           =   0
            Left            =   5400
            TabIndex        =   40
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000007&
            BackStyle       =   0  'Transparent
            Caption         =   "Grand Total  :"
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   0
            Left            =   2880
            TabIndex        =   39
            Top             =   3240
            Width           =   2175
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Levels"
            ForeColor       =   &H00008000&
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   38
            Top             =   480
            Width           =   1935
         End
         Begin VB.Line Line1 
            BorderWidth     =   3
            X1              =   240
            X2              =   6600
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line4 
            BorderWidth     =   3
            X1              =   240
            X2              =   6600
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Line Line7 
            BorderWidth     =   3
            X1              =   240
            X2              =   6600
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line Line8 
            BorderWidth     =   2
            X1              =   2520
            X2              =   2520
            Y1              =   360
            Y2              =   3720
         End
         Begin VB.Line Line9 
            BorderWidth     =   2
            X1              =   5280
            X2              =   5280
            Y1              =   360
            Y2              =   3120
         End
         Begin VB.Line Line10 
            BorderWidth     =   3
            X1              =   6600
            X2              =   2520
            Y1              =   3720
            Y2              =   3720
         End
      End
      Begin VB.ComboBox negativemarks 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         ItemData        =   "tc1.frx":0000
         Left            =   6000
         List            =   "tc1.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "No"
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   3840
         TabIndex        =   12
         Top             =   1440
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Yes"
         ForeColor       =   &H0000C000&
         Height          =   420
         Left            =   4920
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Tmarks 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   10
         Text            =   "100"
         Top             =   840
         Width           =   1095
      End
      Begin VB.Image back 
         Height          =   1200
         Index           =   2
         Left            =   120
         Picture         =   "tc1.frx":004F
         Top             =   7200
         Width           =   2745
      End
      Begin VB.Image next 
         Height          =   1290
         Index           =   2
         Left            =   6600
         Picture         =   "tc1.frx":1396
         Top             =   7080
         Width           =   2835
      End
      Begin VB.Image Image2 
         Height          =   4500
         Left            =   1800
         Picture         =   "tc1.frx":2800
         Top             =   2640
         Width           =   7275
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "3) Marking scheme:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   720
         TabIndex        =   16
         Top             =   2160
         Width           =   3255
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "2) -ve marking :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   720
         TabIndex        =   14
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "1) Total marks :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   720
         TabIndex        =   13
         Top             =   840
         Width           =   2535
      End
   End
   Begin VB.Frame step 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Step2"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6615
      Index           =   1
      Left            =   6000
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   9615
      Begin VB.TextBox quetype 
         Height          =   540
         Left            =   6480
         TabIndex        =   43
         Text            =   "1"
         Top             =   480
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   3120
         TabIndex        =   18
         Top             =   1080
         Width           =   5415
         Begin VB.CheckBox que 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Caption         =   "Single ans question"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Index           =   0
            Left            =   1080
            Picture         =   "tc1.frx":6D282
            TabIndex        =   23
            Top             =   480
            Width           =   3135
         End
         Begin VB.CheckBox que 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Match the following"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   660
            Index           =   4
            Left            =   1080
            TabIndex        =   22
            Top             =   2400
            Width           =   3135
         End
         Begin VB.CheckBox que 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "True and  False"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   2
            Left            =   1080
            TabIndex        =   21
            Top             =   1440
            Width           =   3135
         End
         Begin VB.CheckBox que 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Fill in the blanks"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   3
            Left            =   1080
            TabIndex        =   20
            Top             =   1920
            Width           =   3135
         End
         Begin VB.CheckBox que 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   " Mutiple ans question"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   1
            Left            =   1080
            TabIndex        =   19
            Top             =   960
            Width           =   3135
         End
         Begin VB.Image Image5 
            Height          =   3840
            Left            =   120
            Picture         =   "tc1.frx":6EA6A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   5130
         End
      End
      Begin VB.Image next 
         Height          =   1290
         Index           =   1
         Left            =   6720
         Picture         =   "tc1.frx":BDEAC
         Top             =   5280
         Width           =   2835
      End
      Begin VB.Image back 
         Height          =   1200
         Index           =   1
         Left            =   120
         Picture         =   "tc1.frx":BF316
         Top             =   5280
         Width           =   2745
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Select the type of  Question :"
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   6135
      End
   End
   Begin VB.Frame step 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Step1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   5535
      Index           =   0
      Left            =   6000
      TabIndex        =   0
      Top             =   2760
      Width           =   9615
      Begin VB.ComboBox min 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3360
         Width           =   1095
      End
      Begin VB.ComboBox hr 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "tc1.frx":C065D
         Left            =   3480
         List            =   "tc1.frx":C0670
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3360
         Width           =   855
      End
      Begin VB.ComboBox subject 
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2520
         Width           =   3255
      End
      Begin VB.ComboBox year 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   450
         ItemData        =   "tc1.frx":C0688
         Left            =   3480
         List            =   "tc1.frx":C068A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Image Image3 
         Height          =   2250
         Left            =   7800
         Picture         =   "tc1.frx":C068C
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "3) Duration:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   1200
         TabIndex        =   17
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Image next 
         Height          =   1290
         Index           =   0
         Left            =   6720
         Picture         =   "tc1.frx":C1B07
         Top             =   4200
         Width           =   2835
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "2) Subject :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   1200
         TabIndex        =   6
         Top             =   2520
         Width           =   3135
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "1) Catogory :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   1200
         TabIndex        =   5
         Top             =   1800
         Width           =   2775
      End
   End
   Begin VB.Image Image4 
      Height          =   2445
      Left            =   7200
      Picture         =   "tc1.frx":C2F71
      Top             =   120
      Width           =   7830
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   6480
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   -600
      Picture         =   "tc1.frx":101613
      Top             =   -960
      Width           =   6000
   End
   Begin VB.Image Image11 
      Height          =   1920
      Left            =   240
      Picture         =   "tc1.frx":10FEFC
      Top             =   8640
      Width           =   1920
   End
   Begin VB.Image img1 
      Height          =   4485
      Left            =   -3960
      Picture         =   "tc1.frx":112D16
      Top             =   240
      Width           =   4485
   End
End
Attribute VB_Name = "tc1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================
'      GLOBAL Variable
'========================
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command
Dim temp
Dim qflag As Boolean
Dim counter As Integer
'----------------------

Dim textval As TextBox
Dim fsys As New Scripting.FileSystemObject
Dim tstream As TextBox


'============END of Global varible===
'*********************
Private Sub insertmin()
    Dim i As Integer
        
    i = 0
    While i < 60
        min.AddItem i
        i = i + 5
    Wend
    

End Sub
'******************


Private Sub back_Click(Index As Integer)
Call hideall

step(Index - 1).ZOrder
step(Index - 1).Visible = True
End Sub





Private Sub Command1_Click()

End Sub

Private Sub Form_Load()

     Call initsub_Click
     Module2.initlevels
    
    
End Sub
Private Sub addsubject()
    subject.AddItem "DSP"
    subject.AddItem "DSA"
    subject.AddItem "SPL"
    subject.AddItem "DBMS"
    subject.AddItem "DEL"
    subject.AddItem "APPTITUDE"
    subject.AddItem "C"
    subject.AddItem "C++"
    subject.AddItem "JAVA"
End Sub
'********************************
Private Sub fe()

    subject.AddItem "FPL"
    subject.AddItem "Mechanics"
    subject.AddItem "Electronics"


End Sub
Private Sub Se()

    subject.AddItem "DEL"
    subject.AddItem "DSA"
    subject.AddItem "PPS"


End Sub

Private Sub Te()

    subject.AddItem "DBMS"
    subject.AddItem "TOC"
    subject.AddItem "SPL"
  
End Sub

Private Sub Be()

    subject.AddItem "COMPILER"
    
End Sub
Private Sub lang()

    subject.AddItem "C"
    subject.AddItem "C++"
    subject.AddItem "JAVA"
    subject.AddItem ".NET"
    subject.AddItem "PHP"
    
End Sub

'*******************************
Private Function subremove(comb As ComboBox)
    Dim count As Integer
    count = comb.ListCount
    
    While count > 0
        comb.RemoveItem (count - 1)
        count = count - 1
    Wend
    
End Function
'**********************************


Private Sub loadsub()

    Dim fsys As New Scripting.FileSystemObject
    Dim tstream As TextStream
    Dim filename As String, path As String
    
    filename = File.filepath()
    
    If year.Text <> "" Then
         Set tstream = fsys.OpenTextFile(filename, ForReading, False)
        subremove subject
        While tstream.AtEndOfStream = False
            subject.AddItem (tstream.ReadLine)
        Wend
    End If
    

End Sub




Private Sub gtotal_Change(Index As Integer)

    'if val(gtotal(0).Text)>
End Sub





Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'     Leftarr(0).Visible = False
'     Leftarr(1).Visible = True
    
End Sub

Private Sub Leftarr_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
        
   
'       Leftarr(0).Visible = True
'       Leftarr(1).Visible = False
       T1.Enabled = True
       
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
 '   MsgBox "Error in Excecuting query"
    qflag = True
End Sub

Private Sub initsub_Click()
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
        
        query = "select distinct catagory from tabledetail"
        Call excquery(query)
        While (rs.EOF = False)
                year.AddItem (rs(0))
                rs.MoveNext
        Wend
        
        
    Else
        counter = 0
    End If
    
End Sub

Private Sub level_Click(Index As Integer)
If (level(Index).Value = 1) Then
    nq(Index).Enabled = True
    nm(Index).Enabled = True
    total(Index).Enabled = True
Else
    nq(Index).Enabled = False
    nm(Index).Enabled = False
    nq(Index).Text = "0"
    nm(Index).Text = "0"
    
    total(Index).Enabled = False
    
End If
Call grandtotal(Index)
End Sub

Private Sub min_GotFocus()
    Call insertmin
End Sub
'===========================
' text validation
'============================
Private Sub textvalid(TextBox1 As TextBox)



    Dim i, j
    Dim ch As String
    
    
    i = TextBox1.SelStart
    If i <> 0 Then
    
          ch = Mid(TextBox1.Text, i, 1)
        If (Asc(ch) < 48 Or Asc(ch) >= 58) Then
            TextBox1.SelStart = i - 1
            TextBox1.SelLength = 1
            TextBox1.SelText = vbNullChar
    
        
        End If
    End If

End Sub

Private Sub setcursor(TextBox1 As TextBox)

    If (Len(TextBox1.Text) = 0) Then
        TextBox1.Text = "0"
        TextBox1.SelStart = 1
    Else
             If ((Asc(Left(TextBox1.Text, 1)) = 48) And Len(TextBox1.Text) > 1) Then
                With TextBox1
                    .SelStart = 0
                    .SelLength = 1
                    .SelText = vbNullChar
                    .SelStart = 1
                End With
                
            End If
    End If


End Sub
Private Sub grandtotal(n As Integer)
Dim ans As Double
Dim i
ans = 0
i = 0
    While i < 3
    
        If (total(i).Enabled = True) Then
             ans = Val(total(i).Text) + ans
        End If
        i = i + 1
    Wend
    If (ans > Tmarks.Text) Then
    
        errormsg.msg.Text = " ERROR :Grand Total is exceeding the Total marks"
        errormsg.Show 1
        nm(n).Text = "0"
        Exit Sub
        
    End If
gtotal(0).Text = ans


End Sub
'==========END of text validation =====



Private Sub hideall()
    step(0).Visible = False
    step(1).Visible = False
    step(2).Visible = False
    
End Sub
Private Sub next_Click(Index As Integer)
i = 0
Dim flag As Boolean
If (Index = 1) Then
    i = 0
    flag = False
    While (i < 5)
        If (que(i).Value = 1) Then
                flag = True
        End If
        i = i + 1
    Wend
    If (flag = False) Then
        Exit Sub
    End If
End If

Call hideall
If (Index = 1) Then
    i = 0
    quetype.Text = ""
    While (i < 5)
        If (que(i).Value = 1) Then
            quetype.Text = quetype.Text + str(i) + ","
        End If
        i = i + 1
    Wend
    quetype.Text = Left(quetype.Text, Len(quetype.Text) - 1)
     
End If

If (Index + 1 < 3) Then

    step(Index + 1).ZOrder
    step(Index + 1).Visible = True
Else
       tc1.Hide
       instruction.Visible = True
End If


End Sub

Private Sub nm_Change(Index As Integer)
    textvalid nm(Index)
    setcursor nm(Index)
    total(Index).Text = Val(nq(Index).Text) * Val(nm(Index).Text)
    Call grandtotal(Index)
   
End Sub

Private Sub nm_GotFocus(Index As Integer)
nm(Index).SelStart = 1
   
End Sub

Private Sub nq_Change(Index As Integer)
    textvalid nq(Index)
    setcursor nq(Index)
    total(Index).Text = Val(nq(Index).Text) * Val(nm(Index).Text)
    
    Call grandtotal(Index)


End Sub

Private Sub nq_GotFocus(Index As Integer)
   nq(Index).SelStart = 1
   
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
        negativemarks.Enabled = True
Else
        negativemarks.Enabled = False
End If
    
End Sub

Private Sub Option2_Click()
If Option1.Value = True Then
        negativemarks.Enabled = True
Else
        negativemarks.Enabled = False
End If
End Sub

Private Sub que_Click(Index As Integer)
    If que(Index).Value = 1 Then
        que(Index).ForeColor = &HFF00&
    Else
        que(Index).ForeColor = &H80C0FF
    End If
    
End Sub

Private Sub que_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    
    que(0).ForeColor = &H80000008
    que(1).ForeColor = &H80000008
    que(2).ForeColor = &H80000008
    que(3).ForeColor = &H80000008
    que(4).ForeColor = &H80000008
    Select Case Index
        
            Case 0:
                    que(0).ForeColor = &HFF00&
  
            Case 1:  que(1).ForeColor = &HFF00&
            
            Case 2:  que(2).ForeColor = &HFF00&
  
            Case 3:  que(3).ForeColor = &HFF00&
  
            Case 4:  que(4).ForeColor = &HFF00&
     
    
    
    End Select
  
    
    
    
End Sub



Private Sub T1_Timer()
'    Leftarr(0).Width = Leftarr(0).Width + 1
          
'      If Leftarr(0).Width <= 1800 Then
'            T1.Enabled = False
'      End If
'
    
       
    
End Sub

Private Sub Text2_Change(Index As Integer)

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Tmarks_Change()
    textvalid Tmarks
    setcursor Tmarks
 End Sub

Private Sub loadsubject1()
         subject.Clear
        query = "select catagory,subject from tabledetail"
        Call excquery(query)
        While (rs.EOF = False)
                If (year.Text = rs(0)) Then
                    subject.AddItem (rs(1))
                End If
                rs.MoveNext
                
        Wend
End Sub

Private Sub year_Click()
    subject.Enabled = True
    Call loadsubject1
End Sub
