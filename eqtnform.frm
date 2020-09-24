VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form eqtnform 
   BackColor       =   &H8000000A&
   Caption         =   "Equation Formatter"
   ClientHeight    =   3480
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9465
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   390
      TabIndex        =   22
      Top             =   1605
      Width           =   1680
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Refresh View"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6360
      TabIndex        =   21
      Top             =   1620
      Width           =   1680
   End
   Begin RichTextLib.RichTextBox rtb2 
      Height          =   525
      Left            =   225
      TabIndex        =   18
      Top             =   2520
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   926
      _Version        =   393217
      TextRTF         =   $"eqtnform.frx":0000
   End
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   525
      Left            =   195
      TabIndex        =   17
      Top             =   1950
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   926
      _Version        =   393217
      TextRTF         =   $"eqtnform.frx":00D7
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1065
      Left            =   135
      ScaleHeight     =   1005
      ScaleWidth      =   9120
      TabIndex        =   14
      Top             =   465
      Width           =   9180
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   8955
      Top             =   1350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
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
      Height          =   345
      Left            =   30
      ScaleHeight     =   285
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   1305
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
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
      Height          =   345
      Left            =   420
      ScaleHeight     =   285
      ScaleWidth      =   375
      TabIndex        =   8
      Top             =   1335
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Courier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7800
      TabIndex        =   7
      Top             =   105
      Width           =   1230
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Greek"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6525
      TabIndex        =   6
      Top             =   105
      Width           =   825
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Super"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4890
      TabIndex        =   5
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Subscript"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3900
      TabIndex        =   4
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DIV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2715
      TabIndex        =   3
      Top             =   105
      Width           =   795
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SQRT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1890
      TabIndex        =   2
      Top             =   105
      Width           =   795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2 line"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   915
      TabIndex        =   1
      Top             =   105
      Width           =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1 line"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   195
      TabIndex        =   0
      Top             =   105
      Width           =   690
   End
   Begin VB.Label Label11 
      Caption         =   "Denom."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8235
      TabIndex        =   20
      Top             =   2625
      Width           =   675
   End
   Begin VB.Label Label10 
      Caption         =   "Num."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8250
      TabIndex        =   19
      Top             =   2100
      Width           =   525
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Editor Window"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3090
      TabIndex        =   16
      Top             =   3060
      Width           =   2550
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Image Preview Window"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3225
      TabIndex        =   15
      Top             =   1575
      Width           =   2550
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Height          =   105
      Left            =   9120
      TabIndex        =   13
      Top             =   195
      Width           =   105
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Height          =   105
      Left            =   6315
      TabIndex        =   12
      Top             =   195
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   105
      Left            =   5940
      TabIndex        =   11
      Top             =   195
      Width           =   105
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   100
      Left            =   3675
      TabIndex        =   10
      Top             =   195
      Width           =   100
   End
   Begin VB.Menu dd8 
      Caption         =   ""
   End
   Begin VB.Menu mnuLoadformula 
      Caption         =   "Load Formula"
   End
   Begin VB.Menu dd5 
      Caption         =   ""
   End
   Begin VB.Menu mnuSaveformula 
      Caption         =   "Save Formula"
   End
   Begin VB.Menu dd3 
      Caption         =   ""
   End
   Begin VB.Menu dd1 
      Caption         =   ""
   End
   Begin VB.Menu mnuSave 
      Caption         =   "Save Image"
   End
   Begin VB.Menu dd6 
      Caption         =   ""
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "Save/Insert Image"
   End
   Begin VB.Menu dd4 
      Caption         =   ""
   End
   Begin VB.Menu dd2 
      Caption         =   ""
   End
   Begin VB.Menu mnuClear 
      Caption         =   "Clear"
   End
   Begin VB.Menu dd7 
      Caption         =   ""
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "eqtnform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   Equation formatter
'

Option Explicit

Dim pixwidth As Long
Dim divwidth As Long
Dim toggleflag As Integer
Dim drawflag As Integer
Dim startpos As Long
Dim endpos As Long
Dim ycor As Long
Dim fnum1 As Integer
Dim x As Integer
Dim y As Integer
Dim i As Integer
Dim filename As String
Dim lineflag As Integer
Dim divflag As Integer
Dim subflag As Integer
Dim superflag As Integer
Dim greekflag As Integer
Dim TNRflag As Integer
Dim focusedon As Integer
Dim txt1 As String
Dim txt2 As String
Dim toplineX1 As Long
Dim toplineX2 As Long
Dim botlineX1 As Long
Dim botlineX2 As Long
Dim toplineY As Long
Dim botlineY As Long
Dim toplineflag As Integer
Dim botlineflag As Integer
Dim holdoff As Integer


Private Sub Command1_Click() '1 line

lineflag = 1
rtb2.Text = ""
rtb2.Visible = False
Picture1.Height = 450
divflag = 0
Picture1.Picture = Picture3.Image

End Sub

Private Sub Command10_Click() 'help

EQhelp.Show

End Sub

Private Sub Command2_Click() '2 line

lineflag = 2
rtb2.Visible = True
Picture1.Height = 1065


End Sub

Private Sub Command3_Click() 'sqrt

rtb1.SelFontName = "Symbol"
rtb2.SelFontName = "Symbol"
greekflag = 1
TNRflag = 0

If focusedon = 0 Then
rtb1.SetFocus
Else
rtb2.SetFocus
End If

SendKeys Chr(214)
DoEvents
rtb1.Refresh
rtb2.Refresh

rtb1.SelFontName = "Courier New"
rtb2.SelFontName = "Courier New"
greekflag = 0
TNRflag = 1
rtb1.SelFontSize = 12
rtb2.SelFontSize = 12
rtb1.Refresh
rtb2.Refresh


End Sub

Private Sub Command4_Click() 'divider

If divflag = 0 Then
divflag = 1
Picture1.Line (100, 435)-(divwidth, 435)
Picture1.Picture = Picture1.Image
Else
divflag = 0
Picture1.Picture = Picture3.Image
Command9_Click
End If

End Sub

Private Sub Command5_Click() 'sub

If subflag = 0 Then
subflag = 1
Label2.BackColor = vbRed
Label3.BackColor = vbWhite
rtb1.SelCharOffset = -60
rtb1.SelFontSize = 8
rtb2.SelCharOffset = -60
rtb2.SelFontSize = 8
If focusedon = 0 Then
rtb1.SetFocus
Else
rtb2.SetFocus
End If
superflag = 0
Else
Label2.BackColor = vbWhite
rtb1.SelCharOffset = 0
rtb1.SelFontSize = 12
rtb2.SelCharOffset = 0
rtb2.SelFontSize = 12
If focusedon = 0 Then
rtb1.SetFocus
Else
rtb2.SetFocus
End If
subflag = 0
End If

End Sub

Private Sub Command6_Click() 'super

If superflag = 0 Then
superflag = 1
Label2.BackColor = vbWhite
Label3.BackColor = vbRed
rtb1.SelCharOffset = 60
rtb1.SelFontSize = 8
rtb2.SelCharOffset = 60
rtb2.SelFontSize = 8
If focusedon = 0 Then
rtb1.SetFocus
Else
rtb2.SetFocus
End If
subflag = 0
Else
Label3.BackColor = vbWhite
rtb1.SelCharOffset = 0
rtb1.SelFontSize = 12
rtb2.SelCharOffset = 0
rtb2.SelFontSize = 12
If focusedon = 0 Then
rtb1.SetFocus
Else
rtb2.SetFocus
End If
superflag = 0
End If

End Sub

Private Sub Command7_Click() 'Greek

greekflag = 1
TNRflag = 0
Label5.BackColor = vbWhite
Label4.BackColor = vbRed
rtb1.SelFontName = "Symbol"
rtb2.SelFontName = "Symbol"
If focusedon = 0 Then
rtb1.SetFocus
Else
rtb2.SetFocus
End If

End Sub

Private Sub Command8_Click() 'Courier

TNRflag = 1
greekflag = 0
rtb1.SelFontName = "Courier New"
rtb2.SelFontName = "Courier New"
Label5.BackColor = vbRed
Label4.BackColor = vbWhite
If focusedon = 0 Then
rtb1.SetFocus
Else
rtb2.SetFocus
End If

End Sub


Private Sub Form_Load()

divwidth = 10000
rtb1.SelFontName = "Courier New"
rtb2.SelFontName = "Courier New"
Label5.BackColor = vbRed
lineflag = 2
TNRflag = 1
focusedon = 0
rtb1.TextRTF = ""
rtb2.TextRTF = ""
txt1 = ""
txt2 = ""
Picture3.Width = 9180
Picture3.Height = 1065
Picture1.Picture = Picture3.Image
pixwidth = 9180

End Sub

Private Sub mnuClear_Click() 'clear

rtb1.TextRTF = ""
rtb2.TextRTF = ""
txt1 = ""
txt2 = ""
Picture1.Picture = Picture3.Image
Picture1.Width = 9180
toplineflag = 0
botlineflag = 0
pixwidth = 9180

End Sub

Private Sub mnuExit_Click()  'exit

Unload eqtnform

End Sub


Private Sub mnuInsert_Click()  'save and insert

dlgFile.Filter = "All Files (*.bmp)|*.bmp|"
dlgFile.filename = ""
dlgFile.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames
    
    On Error Resume Next
    dlgFile.ShowSave
    If Err.Number = cdlCancel Then Exit Sub
    If Err.Number <> 0 Then
        MsgBox "Error" & Str$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
    End If
    On Error GoTo 0

filename = dlgFile.filename

If filename <> "" Then
makepicture
On Error Resume Next
SavePicture Picture2.Image, filename
End If

FormNB.imgfilename = filename
FormNB.PlaceImage

End Sub


Private Sub mnuLoadformula_Click()

dlgFile.Filter = "All Files (*.eqn)|*.eqn|"
dlgFile.filename = ""
dlgFile.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames
    
    On Error Resume Next
    dlgFile.ShowOpen
    If Err.Number = cdlCancel Then
    Exit Sub
    End If
    If Err.Number <> 0 Then
        MsgBox "Error" & Str$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
    End If
    On Error GoTo 0

filename = dlgFile.filename

If filename <> "" Then

mnuClear_Click
fnum1 = FreeFile
On Error Resume Next
Open filename For Input As #fnum1
On Error Resume Next
Input #fnum1, divflag
Input #fnum1, lineflag
Input #fnum1, txt1
Input #fnum1, txt2
Input #fnum1, toplineflag
Input #fnum1, botlineflag
Input #fnum1, toplineX1
Input #fnum1, toplineX2
Input #fnum1, botlineX1
Input #fnum1, botlineX2
Input #fnum1, toplineY
Input #fnum1, botlineY
Input #fnum1, divwidth
Input #fnum1, pixwidth
Close fnum1

rtb1.TextRTF = txt1
rtb2.TextRTF = txt2
Picture1.Width = pixwidth
Picture2.Width = pixwidth
If divflag = 1 Then
Picture1.Line (100, 435)-(10000, 435)
Picture1.Picture = Picture1.Image
Else
Picture1.Picture = Picture3.Image
End If

If lineflag = 1 Then
Command1_Click
Else
Command2_Click
End If

makepicture
Command9_Click


End If


End Sub

Private Sub mnuSave_Click()  'save as image file .bmp

dlgFile.Filter = "All Files (*.bmp)|*.bmp|"
dlgFile.filename = ""
dlgFile.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames
    
    On Error Resume Next
    dlgFile.ShowSave
    If Err.Number = cdlCancel Then Exit Sub
    If Err.Number <> 0 Then
        MsgBox "Error" & Str$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
    End If
    On Error GoTo 0

filename = dlgFile.filename

If filename <> "" Then
makepicture
On Error Resume Next
SavePicture Picture2.Image, filename
End If

End Sub


Private Function makepicture()

Picture2.Picture = Picture3.Image
Picture2.Height = Picture1.Height
Picture2.Width = Picture1.Width
pixwidth = Picture1.Width
Picture2.Picture = Picture1.Image
If Len(rtb1.Text) > Len(rtb2.Text) Then
'Picture2.Width = 500 + 150 * (Len(rtb1.Text)) '(Me.TextWidth(rtb1.Text))
Else
'Picture2.Width = 500 + 150 * (Len(rtb2.Text)) '(Me.TextWidth(rtb2.Text))
End If
Picture2.FontSize = 12

Picture2.CurrentX = 190
Picture2.CurrentY = 90
For x = 1 To Len(rtb1.Text)
rtb1.SetFocus
rtb1.SelStart = x
Picture2.FontName = rtb1.SelFontName
Picture2.FontSize = rtb1.SelFontSize
Picture2.CurrentX = 155 + (x * 150)
If rtb1.SelCharOffset > 0 Then
Picture2.CurrentY = 90 - 0.4 * rtb1.SelCharOffset
Else
Picture2.CurrentY = 90 - 2.5 * rtb1.SelCharOffset
End If
Picture2.Print Mid$(rtb1.Text, x, 1)
Next x


Picture2.CurrentX = 210
Picture2.CurrentY = 550
  If lineflag = 2 Then
For x = 1 To Len(rtb2.Text)
rtb2.SetFocus
rtb2.SelStart = x
Picture2.FontName = rtb2.SelFontName
Picture2.FontSize = rtb2.SelFontSize
Picture2.CurrentX = 185 + (x * 150)
If rtb2.SelCharOffset > 0 Then
Picture2.CurrentY = 550 - 0.25 * rtb2.SelCharOffset
Else
Picture2.CurrentY = 550 - 2.5 * rtb2.SelCharOffset
End If
Picture2.Print Mid$(rtb2.Text, x, 1)
Next x
  End If

End Function

Private Sub mnuSaveformula_Click()


Command9_Click

dlgFile.Filter = "All Files (*.eqn)|*.eqn|"
dlgFile.filename = ""
dlgFile.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames
    
    On Error Resume Next
    dlgFile.ShowSave
    If Err.Number = cdlCancel Then Exit Sub
    If Err.Number <> 0 Then
        MsgBox "Error" & Str$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
    End If
    On Error GoTo 0

filename = dlgFile.filename
txt1 = rtb1.TextRTF
txt2 = rtb2.TextRTF

If filename <> "" Then
fnum1 = FreeFile
On Error Resume Next
Open filename For Output As #fnum1
On Error Resume Next
Write #fnum1, divflag
Write #fnum1, lineflag
Write #fnum1, txt1
Write #fnum1, txt2
Write #fnum1, toplineflag
Write #fnum1, botlineflag
Write #fnum1, toplineX1
Write #fnum1, toplineX2
Write #fnum1, botlineX1
Write #fnum1, botlineX2
Write #fnum1, toplineY
Write #fnum1, botlineY
Write #fnum1, divwidth
Write #fnum1, pixwidth
Close fnum1
End If

End Sub

Private Sub rtb1_Click()

focusedon = 0

If greekflag = 1 Then
rtb1.SelFontName = "Symbol"
End If

If TNRflag = 1 Then
rtb1.SelFontName = "Courier New"
End If

End Sub

Private Sub rtb2_Click()

focusedon = 1

If greekflag = 1 Then
rtb2.SelFontName = "Symbol"
End If

If TNRflag = 1 Then
rtb2.SelFontName = "Courier New"
End If

End Sub


Private Sub picture1_MouseDown(Button As Integer, Shift As Integer, X2 As Single, Y2 As Single)

If Button = 1 Then
drawflag = 1
startpos = X2
ycor = Y2
End If


If Button = 2 Then
 If drawflag = 1 Then
 endpos = X2
 
  If toplineflag = 0 Then
   toplineflag = 1
   Picture1.Line (startpos, ycor)-(endpos, ycor)
   toplineX1 = startpos
   toplineX2 = endpos
   toplineY = ycor
  Else
   If botlineflag = 0 Then
   botlineflag = 1
   Picture1.Line (startpos, ycor)-(endpos, ycor)
   botlineX1 = startpos
   botlineX2 = endpos
   botlineY = ycor
   End If
  End If
  
  If toplineflag = 1 And botlineflag = 1 Then
    If toggleflag = 0 Then
    toggleflag = 1
     Picture1.Line (startpos, ycor)-(endpos, ycor)
   toplineX1 = startpos
   toplineX2 = endpos
   toplineY = ycor
   Command9_Click
    Else
    toggleflag = 0
    Picture1.Line (startpos, ycor)-(endpos, ycor)
   botlineX1 = startpos
   botlineX2 = endpos
   botlineY = ycor
   Command9_Click
    End If
  
  End If
  
 End If
drawflag = 0
End If



End Sub


Private Sub picture1_MouseMove(Button As Integer, Shift As Integer, X2 As Single, Y2 As Single)

If Button = 1 And Shift = 0 Then
Picture1.Width = X2
Picture2.Width = X2
pixwidth = X2
Command9_Click
End If

If Button = 1 And Shift = 1 Then
divwidth = X2
Command9_Click
End If


End Sub


Private Sub Command9_Click() 'refresh picture view

Picture1.Picture = Picture3.Image
Picture1.Refresh
makepicture
Picture1.Width = pixwidth
Picture1.Picture = Picture2.Image
Picture1.Picture = Picture1.Image

If divflag = 1 Then
Picture1.Line (100, 435)-(divwidth, 435)
End If
If toplineflag = 1 Then
Picture1.Line (toplineX1, toplineY)-(toplineX2, toplineY)
End If
If botlineflag = 1 Then
Picture1.Line (botlineX1, botlineY)-(botlineX2, botlineY)
End If

End Sub



