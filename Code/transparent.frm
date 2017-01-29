VERSION 5.00
Begin VB.Form error_transparent 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7965
   ClientLeft      =   105
   ClientTop       =   885
   ClientWidth     =   11145
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "error_transparent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prevline

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bDefaut As Byte, ByVal dwFlags As Long) As Long

Private Const GWL_EXSTYLE  As Long = (-20)
Private Const LWA_COLORKEY  As Long = &H1
Private Const LWA_Defaut  As Long = &H2
Private Const WS_EX_LAYERED  As Long = &H80000

 Public Function Transparency(ByVal hWnd As Long, Optional ByVal Col As Long = vbBlack, _
 Optional ByVal PcTransp As Byte = 255, Optional ByVal TrMode As Boolean = True) As Boolean
' Return : True if there is no error.
' hWnd   : hWnd of the window to make transparent
' Col : Color to make transparent if TrMode=False
' PcTransp  : 0 à 255 >> 0 = transparent  -:- 255 = Opaque
Dim DisplayStyle As Long
    On Error GoTo Exit1
    VoirStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    If DisplayStyle <> (DisplayStyle Or WS_EX_LAYERED) Then
        DisplayStyle = (DisplayStyle Or WS_EX_LAYERED)
        Call SetWindowLong(hWnd, GWL_EXSTYLE, DisplayStyle)
    End If
    Transparency = (SetLayeredWindowAttributes(hWnd, Col, PcTransp, IIf(TrMode, LWA_COLORKEY Or LWA_Defaut, LWA_COLORKEY)) <> 0)
    
Exit1:
    If Not err.Number = 0 Then err.Clear
End Function

Public Sub ActiveTransparency(m As Form, d As Boolean, F As Boolean, _
     T_Transparency As Integer, Optional Color As Long)
Dim B As Boolean
        If d And F Then
        'Makes color (here the background color of the shape) transparent
        'upon value of T_Transparency
            B = Transparency(m.hWnd, Color, T_Transparency, False)
        ElseIf d Then
            'Makes form, including all components, transparent
            'upon value of T_Transparency
            B = Transparency(m.hWnd, 0, T_Transparency, True)
        Else
            'Restores the form opaque.
            B = Transparency(m.hWnd, , 255, True)
        End If
End Sub

Private Sub Form_Load()
prevline = 0
Dim i As Integer
    
     ActiveTransparency Me, True, False, 0
     Me.Show
    
    'For i = 0 To 200 Step 5
        ActiveTransparency Me, True, False, 200
        Me.Refresh
    'Next i
    
    fileerror.Show 1
  
    
End Sub

Private Sub rtb_KeyPress(KeyAscii As Integer)
 Dim str
    Dim length As Integer, remain As Integer, noofchar As Integer, q As Integer
    Dim noofline, line
    noofchar = 16
    length = Len(rtb.Text) + 1
    q = (length) / noofchar
        
    line = 0
    If ((length > q * noofchar) And q <> 0) Then
           line = q + 1
    Else
    
    End If

    If prevline <> line And line <> 0 Then
            intial = rtb.Font.size * 25
            rtb.Height = rtb.Height + intial
            
    'MsgBox "q=" & q & "=" & prevline
    prevline = line
    End If
   
End Sub
