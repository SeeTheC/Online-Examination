VERSION 5.00
Begin VB.Form que_file 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10710
   FillColor       =   &H000080FF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton finish 
      BackColor       =   &H008080FF&
      Caption         =   "FINISH"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   9360
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   2160
      TabIndex        =   13
      Top             =   1680
      Width           =   13935
      Begin VB.TextBox txtop 
         Appearance      =   0  'Flat
         Height          =   1440
         Index           =   3
         Left            =   10800
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   28
         Top             =   4800
         Width           =   2655
      End
      Begin VB.TextBox txtop 
         Appearance      =   0  'Flat
         Height          =   1440
         Index           =   2
         Left            =   7320
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   27
         Top             =   4800
         Width           =   2775
      End
      Begin VB.TextBox txtop 
         Appearance      =   0  'Flat
         Height          =   1440
         Index           =   1
         Left            =   3960
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   26
         Top             =   4800
         Width           =   2775
      End
      Begin VB.TextBox txtop 
         Appearance      =   0  'Flat
         Height          =   1440
         Index           =   0
         Left            =   600
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   25
         Top             =   4800
         Width           =   2775
      End
      Begin VB.ListBox quelist 
         BackColor       =   &H00C0FFFF&
         Height          =   3300
         Left            =   480
         TabIndex        =   18
         Top             =   960
         Width           =   13215
      End
      Begin VB.ListBox options 
         BackColor       =   &H00C0FFFF&
         Height          =   2940
         Index           =   0
         Left            =   9360
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ListBox options 
         BackColor       =   &H00C0FFFF&
         Height          =   2940
         Index           =   1
         Left            =   10440
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ListBox options 
         BackColor       =   &H00C0FFFF&
         Height          =   2940
         Index           =   2
         Left            =   11520
         TabIndex        =   15
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ListBox options 
         BackColor       =   &H00C0FFFF&
         Height          =   2940
         Index           =   3
         Left            =   12600
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   615
         Index           =   0
         Left            =   1560
         TabIndex        =   23
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Index           =   1
         Left            =   4920
         TabIndex        =   22
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   615
         Index           =   2
         Left            =   8400
         TabIndex        =   21
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   615
         Index           =   3
         Left            =   12240
         TabIndex        =   20
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Questions :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Index           =   4
         Left            =   600
         TabIndex        =   19
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "load signature"
      Height          =   495
      Left            =   9360
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame ERROR 
      Caption         =   "Error"
      Height          =   7815
      Left            =   12000
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox continue_text 
         Height          =   495
         Left            =   1800
         TabIndex        =   10
         Text            =   "0"
         Top             =   5640
         Width           =   615
      End
      Begin VB.TextBox error_del 
         Height          =   495
         Left            =   1800
         TabIndex        =   7
         Text            =   "0"
         Top             =   4920
         Width           =   615
      End
      Begin VB.ListBox op_error 
         Height          =   1500
         Left            =   600
         TabIndex        =   6
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox que_error 
         Height          =   975
         Left            =   600
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "continue :"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   5760
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "delete:"
         Height          =   615
         Left            =   720
         TabIndex        =   9
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "RETURN:"
         Height          =   615
         Left            =   480
         TabIndex        =   8
         Top             =   4440
         Width           =   1935
      End
   End
   Begin VB.TextBox display 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label status 
      BackStyle       =   0  'Transparent
      Caption         =   "LOADING...."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   975
      Left            =   8040
      TabIndex        =   24
      Top             =   840
      Visible         =   0   'False
      Width           =   6015
   End
End
Attribute VB_Name = "que_file"
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
'--------------------------
Dim sig_que, sig_op, sig_que_flag, sig_op_flag
Dim sig_removed_line As String

Private Const max As Integer = 2000
Dim que(max) As String, k As Integer
Dim op(max) As String
Dim opcount, shiftcount
Dim opa(max) As String, opb(max) As String, opc(max) As String, opd(max) As String
Dim stropa As String, stropb As String, stropc As String, stropd As String, strope As String, strop As String
'strop:-store option signature
Dim delerror As Boolean
Dim debugflag As Boolean

Dim fsys As New Scripting.FileSystemObject
Dim filein As TextStream
' for file handling
Dim path, cantopen As Boolean
Private Sub openfile()
path = insert_que_by_file.path
If (fsys.FileExists(path) = False) Then
        cantopen = True
        Exit Sub
End If
Set filein = fsys.OpenTextFile(path, ForReading)
   
End Sub

Private Function chkstart(ByVal line As String, ByVal signature As String, ByVal flag As Integer) As Boolean
    'flag =1 number in signature
    'flag =0 char in signatutre
    Dim noofchar, opcount
    Dim questart As Boolean
    Dim i, j As Integer
    questart = False
    
     If filein.AtEndOfStream = True Then
         
                 If (Len(line) = 0) Then
                        Exit Function
                        
                 End If
     End If
      
    'If (Mid(signature, 1, 1) = Mid(line, 1, 1)) Then
           
     Dim found As Boolean, chk As Boolean
       
       chk = True
       nofound = False
       
        i = 1
        j = 1
        If (InStr(1, signature, "x") = 0) Then ' if "x" is not there
              found = True
        Else
            found = False
        End If
                        
       While chk = True
                                
                ' i is denonating number
                If found = True Then
                    
                    If (Mid(line, i, 1) = Mid(signature, Len(signature), 1)) Then
                          sig_removed_line = Mid(line, i + 1, Len(line) - i)
                          strop = Mid(line, 1, i)
                          'MsgBox strop
                          
                          questart = True
                          
                          GoTo break
                    End If
                End If
                
                ' finding the in between no.
                If (Mid(signature, j, 1) <> Mid(line, i, 1)) Then
                                                
                        ' checking x . as x will be in between number or alphabet
                        ' InStr(1, sig_que, "x") means x is there in signature
                        If (Mid(signature, j, 1) <> "x" And InStr(1, signature, "x") <> 0) Then
                                
                                questart = False
                                GoTo break
                        
                        End If
                        
                        ' is a no.
                        If (InStr(1, signature, "x") = 0) Then ' if "x" is not there
                                questart = False
                                GoTo break
                        
                                    
                        
                        Else
                        ' x is a number
                        If (flag = 1) Then
                               
                                
                                If (Asc(Mid(line, i, 1)) >= 48 And Asc(Mid(line, i, 1)) < 58) Then
                                         While (Asc(Mid(line, i, 1)) >= 48 And Asc(Mid(line, i, 1)) <= 58)
                                    
                                              i = i + 1
                                         Wend
                                         found = True
                            
                                        i = i - 1
                                Else
                                        questart = False
                                        GoTo break
                                End If
                        End If
                        ' x is a char
                        If (flag = 0) Then
                            ' not alphbet
                            If (Asc(Mid(line, i, 1)) < 65 Or Asc(Mid(line, i, 1)) >= 91) And (Asc(Mid(line, i, 1)) < 97 Or Asc(Mid(line, i, 1)) >= 123) Then
                                        
                                        questart = False
                                        GoTo break
                            Else    ' is a aplhabet
                                    If noofchar = 0 Then
                                        noofchar = 1
                                        found = True
                                    Else
                                        questart = False
                                        GoTo break
                                    End If
                                 
                            End If
                            
                        
                        End If
                        'MsgBox "i=" & i & "==" & Mid(line, i, 1) & nofound
                        
                    
                    End If ' of instr
                
                End If
          i = i + 1
          j = j + 1
       Wend
break:
    
    'End If
    
    chkstart = questart
     
End Function
Private Function trimline(ByRef line) As String
         Dim i
         While Len(line) = 0 And filein.AtEndOfStream = False
                line = filein.ReadLine
         Wend
    
         If filein.AtEndOfStream = False Then
         
                 i = 1
                While (Mid(line, i, 1) = " ")
                    i = i + 1
                Wend
                line = Mid(line, i, Len(line))
         
                trimline = line
        End If

End Function
Private Sub shiftoption(ByVal n As Integer)
    Dim i
    i = 0
    While (i < n)
            opa(k) = opb(k)
            opb(k) = opc(k)
            opc(k) = opd(k)
            opd(k) = op(k)
            op(k) = ""
            i = i + 1
    Wend
End Sub
Private Sub correct_error()
                                
                               
                                'MsgBox "--->" & stropa
                                
                                shiftcount = 0
                                que_error.Text = ""
                                que_error.SelText = "Q . "
                                que_error.SelText = que(k)
                                que_error.SelText = vbCrLf
                                que_error.SelText = "a ) " + opa(k)
                                que_error.SelText = vbCrLf
                                que_error.SelText = "b ) " + opb(k)
                                que_error.SelText = vbCrLf
                                que_error.SelText = "c ) " + opc(k)
                                que_error.SelText = vbCrLf
                                que_error.SelText = "d ) " + opd(k)
                                que_error.SelText = vbCrLf
                                que_error.SelText = "e ) " + op(k)
                                
                                op_error.Clear
                                op_error.AddItem (opa(k))
                                op_error.AddItem (opb(k))
                                op_error.AddItem (opc(k))
                                op_error.AddItem (opd(k))
                                op_error.AddItem (op(k))
                                
                                
                                'fileerror.Show 1
                                error_transparent.Show
                                Unload error_transparent
                                
                                
                                
                                If error_del.Text = "1" Then
                                   delerror = True
                                End If
                                
                                If continue_text = "1" Then
                                        
                                       'from last
                                       If (op(k) = op_error.List(0)) Then
                                                
                                            shiftcount = 0
                                            op_error.RemoveItem (0)
                                       
                                       ' checking from the top of the error list
                                      
                                       ElseIf (opa(k) = op_error.List(0)) Then
                                             
                                             que(k) = que(k) + "`" + stropa + op_error.List(0)
                                             op_error.RemoveItem (0)
                                             opcount = opcount - 1
                                            
                                            shiftcount = shiftcount + 1
                                            
                                            If (op_error.ListCount <> 0) Then
                                                If (opb(k) = op_error.List(0)) Then
                                                    
                                                    que(k) = que(k) + "`" + stropb + op_error.List(0)
                                                    op_error.RemoveItem (0)
                                                 
                                                    opcount = opcount - 1
                                                    shiftcount = shiftcount + 1
                                                
                                                    If (op_error.ListCount <> 0) Then
                                                            If (opc(k) = op_error.List(0)) Then
                                                                que(k) = que(k) + "`" + stropc + op_error.List(0)
                                                                op_error.RemoveItem (0)
                                                    
                                                                opcount = opcount - 1
                                                                shiftcount = shiftcount + 1
                                            
                                                                If (op_error.ListCount <> 0) Then
                                                                        If (opd(k) = op_error.List(0)) Then
                                                                            que(k) = que(k) + "`" + stropd + op_error.List(0)
                                                                            op_error.RemoveItem (0)
                                                                        
                                                                            opcount = opcount - 1
                                                                            shiftcount = shiftcount + 1
                                            
                                                                            If (op_error.ListCount <> 0) Then
                                                                                    If (opb(k) = op_error.List(0)) Then
                                                                                    que(k) = que(k) + "`" + strope + op_error.List(0)
                                                                                    op_error.RemoveItem (0)
                                                                                    opcount = opcount - 1
                                                                                    shiftcount = shiftcount + 1
                                            
                                                                                    Else
                                                                                            delerror = True
                                                                                            MsgBox "Cannot correct the question. Press ok to delete"
                                                                                    End If
                                            
                                                                            End If
                                                                            
                                                                        Else
                                                                                delerror = True
                                                                                MsgBox "Cannot correct the question. Press ok to delete"
                                                                        End If
                                                                End If
                                                         Else
                                                            delerror = True
                                                            MsgBox "Cannot correct the question. Press ok to delete"
                                                        End If
                                                    End If
                                                                                                    
                                                Else
                                                     delerror = True
                                                    MsgBox "Cannot correct the question. Press ok to delete"
                                                End If
                                            End If
                                    Else
                                         delerror = True
                                         MsgBox "Cannot correct the question. Press ok to delete"
                                    End If
                                  
                                   debugflag = True
                                   Call shiftoption(shiftcount)
                                End If
                                   
                                
                                
End Sub
Private Sub generate_signature()

    
    Dim temp_sig_que As String, temp_sig_op As String
    'temp_sig_que = "Q. 1 ."
    'temp_sig_op = "a ."
    With insert_que_by_file
        temp_sig_que = .question.Text
        temp_sig_op = .options.Text
    End With
        
    
    Dim i As Integer
    i = 1
    While i <= Len(temp_sig_que)
            
            ' for number
            If (Asc(Mid(temp_sig_que, i, 1)) >= 48) And (Asc(Mid(temp_sig_que, i, 1)) < 58) Then
                
                temp_sig_que = Replace(temp_sig_que, Mid(temp_sig_que, i, 1), "x")
                sig_que_flag = 1
                GoTo break1
            End If
            ' que of type Q. a . i.e alphabet in between is noot allowed
            ' when question is of format Q. a .
            'If ((Asc(Mid(temp_sig_que, i, 1)) >= 65) And (Asc(Mid(temp_sig_que, i, 1)) < 91)) Or ((Asc(Mid(temp_sig_que, i, 1)) >= 97) And (Asc(Mid(temp_sig_que, i, 1)) < 123)) Then
                
             '   temp_sig_que = Replace(temp_sig_que, Mid(temp_sig_que, i, 1), "x")
              '  GoTo break1
            'End If
    
        i = i + 1
    Wend
    
break1:
    i = 1
    While i <= Len(temp_sig_op)
            
            ' for number
            If (Asc(Mid(temp_sig_op, i, 1)) >= 48) And (Asc(Mid(temp_sig_op, i, 1)) < 58) Then
                
                temp_sig_op = Replace(temp_sig_op, Mid(temp_sig_op, i, 1), "x")
                sig_op_flag = 1
                GoTo break2
            End If
            If ((Asc(Mid(temp_sig_op, i, 1)) >= 65) And (Asc(Mid(temp_sig_op, i, 1)) < 91)) Or ((Asc(Mid(temp_sig_op, i, 1)) >= 97) And (Asc(Mid(temp_sig_op, i, 1)) < 123)) Then
                
                temp_sig_op = Replace(temp_sig_op, Mid(temp_sig_op, i, 1), "x")
                sig_op_flag = 0
                GoTo break2
            End If
    
        i = i + 1
    Wend
break2:
    'MsgBox temp_sig_que
    
    'MsgBox temp_sig_op
    
    sig_que = temp_sig_que
    sig_op = temp_sig_op

    
End Sub

Private Sub Command1_Click()
       status.Visible = True
       Command1.Enabled = False

     Call generate_signature
     Call openfile
     If (cantopen = True) Then
            MsgBox " CAN'T OPEN THE FILE "
            insert_que_by_file.Visible = True
            Unload Me
            Exit Sub
     End If
    
    Dim line, chkque As Boolean, chkop As Boolean, chk As Boolean
    Dim splitstr() As String
   
    k = 1
    
    'sig_que = "Q ."
    'sig_op = "x ."
    
    'MsgBox InStr(1, sig_que, "x")
    Dim sperator As String
    sperator = "||_________________________________________________||"
    
        line = filein.ReadLine
        line = trimline(line)
        
        chkque = chkstart(line, sig_que, 1)
        
        ' finding the first question
        While chkque = False And filein.AtEndOfStream = False
            
            line = filein.ReadLine
            line = trimline(line)
            chkque = chkstart(line, sig_que, 1)
                
        Wend
        If filein.AtEndOfStream = True Then
                
                MsgBox "INVALID FILE. As NO questions found", vbOKOnly, ERROR
                Exit Sub
        End If
        'Text1.Text = line
    While filein.AtEndOfStream = False
    
         
         
         Text1.Text = line
         display.SelText = line
           
          delerror = False
        If chkque = True Then
            
            chkque = chkstart(line, sig_que, 1)
            que(k) = sig_removed_line
         
     
            line = filein.ReadLine
            line = trimline(line)
            
            chkque = chkstart(line, sig_que, 1)
            chkop = chkstart(line, sig_op, sig_op_flag)
             
            
            ' reading question
            While chkque = False And chkop = False And filein.AtEndOfStream = False

                 que(k) = que(k) + "`" + line
                 display.SelText = vbCrLf
                 display.SelText = line
            
                 line = filein.ReadLine
                 line = trimline(line)
                    
                 chkque = chkstart(line, sig_que, 1)
                 chkop = chkstart(line, sig_op, sig_op_flag)
                 optionlen = 50
                 
            Wend
            opcount = 0
            
            ' checking for the option
            While chkque = False And chkop = True And filein.AtEndOfStream = False
                 
                 opcount = opcount + 1
                
                 op(k) = sig_removed_line
                 
                 display.SelText = vbCrLf
                 display.SelText = "-->" + line
                 
                 'MsgBox "-->" & strop
                 
                 Select Case opcount
                 
                        Case 1: stropa = strop
                                
                        Case 2:
                                stropb = strop
                        Case 3:
                                stropc = strop
                        Case 4:
                                stropd = strop
                        Case Else:
                                strope = strop
                End Select
                 
                 line = filein.ReadLine
                 line = trimline(line)
                 
                 chkque = chkstart(line, sig_que, 1)
                 chkop = chkstart(line, sig_op, sig_op_flag)
                    
                   'MsgBox line
                   'MsgBox "que=" & chkque & " op=" & chkop
                  
                 While chkque = False And chkop = False And filein.AtEndOfStream = False
                        
                      op(k) = op(k) + "`" + line
                      display.SelText = vbCrLf
                      display.SelText = line
            
                      line = filein.ReadLine
                      line = trimline(line)
                           
                      chkop = chkstart(line, sig_op, sig_op_flag)
                      chkque = chkstart(line, sig_que, 1)
                  
                 Wend
                 ' putting the value in the option a,b,c,d, array
                 
                 Select Case opcount
                 
                        Case 1: opa(k) = op(k)
                                                       
                        Case 2: opb(k) = op(k)
                               
                        Case 3: opc(k) = op(k)
                               
                        Case 4: opd(k) = op(k)
                               
                        Case Else:
                                
                                'MsgBox " no. of option are not 4 "
                                If (delerror = False) Then
                                        Call correct_error
                                End If
                                ' reset error
                                     error_del.Text = 0
                                     continue_text.Text = 0
                                     que_error.Text = ""
                                     op_error.Clear
                                     
                                
                 End Select
                 
                 
                 
            Wend
            If (opd(k) = "") Or que(k) = "" Then ' que(k) :-> this is for,when only option is there
                    delerror = True
                    MsgBox "==>Cannot correct the question. Press ok to delete"
            End If
            If (debugflag = True) Then
                debugflag = False
                'MsgBox que(k)
                'MsgBox "a)" & opa(k)
                'MsgBox "b)" & opb(k)
                'MsgBox "c)" & opc(k)
                'MsgBox "d)" & opd(k)
            End If
            
            
            display.SelText = vbCrLf
            display.SelText = sperator
            display.SelText = vbCrLf
           ' que(k) = line
                        
        End If
     'If delerror = False Then
     '    k = k + 1
     'End If
     If delerror = True Then
         k = k - 1
     Else
        quelist.AddItem (que(k))
        options(0).AddItem (opa(k))
        options(1).AddItem (opb(k))
        options(2).AddItem (opc(k))
        options(3).AddItem (opd(k))
        quelist.ListIndex = quelist.ListCount - 1
     End If
     
     k = k + 1
    Wend
listprint:
    
    
    'Call put_value_in_list
    
    status.Caption = " FINISHED"
    finish.Enabled = True
  End Sub
Private Sub put_value_in_list()
    
    Dim i
    Dim str() As String
    i = 1
    While i < k
        quelist.AddItem (que(i))
        options(0).AddItem (opa(i))
        options(1).AddItem (opb(i))
        options(2).AddItem (opc(i))
        options(3).AddItem (opd(i))
        'MsgBox que(i)
        i = i + 1
    Wend

End Sub



Private Sub Command2_Click()
Call generate_signature
End Sub

Private Sub Command3_Click()
 
 Call finish_Click
 
End Sub
Private Sub excquery(ByVal query As String)
    With cmd
        On Error GoTo Exit1
        .CommandText = query
        .ActiveConnection = cn
        Set rs = .Execute
        'If (temp >= 340) Then
        '        MsgBox query
        'End If
       counter = counter + 1
        Exit Sub
    End With
Exit1:
    'MsgBox "error in loading question"
    qflag = True
End Sub




Private Sub finish_Click()
    Dim str1 As String
    Dim i
    Dim que_type, name, diff_level
    Dim query As String
    Dim subject
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    str1 = "Provider=MSDAORA.1" ';User ID=pro;Password=pro"
    cn.Open str1, "pro", "pro"
    
  
   

    'query = " insert into mytable values(3327,'r')"
    'query = "insert into clang values ('3','a','a','a','a','a','a','a','a','2')"
    subject = tlogin.scombo.Text
    query = "select count(qno) from " + subject
    Call excquery(query)
    
        
    If (rs(0) <> 0) Then
        query = "select max(qno) from " + subject
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
    i = 0
    que_type = "SINGLE"
    name = tlogin.tname
    diff_level = "2"
    While i < quelist.ListCount
        'query = "insert into clang values ( '" + str(counter) + "','" + que_type + "','" + name + "','" + que(i) + "','" + opa(i) + "','" + opb(i) + "','" + opc(i) + "','" + opd(i) + "','" + ans + "','" + diff_level + "')"
        'query = "insert into clang values (str(counter),'que_type','name','que(i)','opa(i)','opb(i)','opc(i)','opa(i)d','ans',diff_line)"
        
        ' removing the  " ' " with  (@)
        que(i) = Replace(quelist.List(i), "'", "(@)")
        opa(i) = Replace(options(0).List(i), "'", "(@)")
        opb(i) = Replace(options(1).List(i), "'", "(@)")
        opc(i) = Replace(options(2).List(i), "'", "(@)")
        opd(i) = Replace(options(3).List(i), "'", "(@)")
        
        query = "insert into  " + subject + "  values ('" + str(counter) + "','" + que_type + "','" + name + "','" + que(i) + "','" + opa(i) + "','" + opb(i) + "','" + opc(i) + "','" + opd(i) + "','" + ans + "','" + diff_line + "')"
        temp = i
        qflag = False
        Call excquery(query)
        If (qflag = True) Then
                
               ' MsgBox i
        End If
        
        i = i + 1

    Wend
    teachers.Show
    Unload Me
End Sub

Private Sub quelist_Click()
    Dim i, j
    i = quelist.ListIndex
    
    While (j < 4)
    
        txtop(j).Text = options(j).List(i)
        j = j + 1
    Wend
    'MsgBox " ==>index" & i
End Sub
