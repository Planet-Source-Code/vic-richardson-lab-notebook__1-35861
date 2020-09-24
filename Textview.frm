VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Textview 
   AutoRedraw      =   -1  'True
   Caption         =   "Text Viewer"
   ClientHeight    =   8400
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   8250
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   14552
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Textview.frx":0000
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   9630
      Top             =   75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuRich 
         Caption         =   "Load File"
      End
      Begin VB.Menu dd3 
         Caption         =   "-"
      End
      Begin VB.Menu flastfile 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu flastfile 
         Caption         =   ""
         Index           =   2
      End
      Begin VB.Menu flastfile 
         Caption         =   ""
         Index           =   3
      End
   End
   Begin VB.Menu dummm 
      Caption         =   ""
   End
   Begin VB.Menu mnuClear 
      Caption         =   "Clear"
   End
   Begin VB.Menu dda 
      Caption         =   ""
   End
   Begin VB.Menu mnuSelectall 
      Caption         =   "Select All"
   End
   Begin VB.Menu dd5 
      Caption         =   ""
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
   Begin VB.Menu mnupopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnucopy 
         Caption         =   "Copy"
      End
   End
End
Attribute VB_Name = "Textview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim textfilename As String
Dim fnum1 As Integer
Dim xx As Long
Dim x As Long
Dim txt As Byte
Dim lastused(4) As String
Dim lf As Integer


Private Sub Form_Load()

rtb1.Visible = True
rtb1.SelColor = vbBlack
rtb1.SelFontSize = 12
rtb1.SelFontName = "Times New Roman"

  fnum1 = FreeFile
On Error Resume Next
Open "C:/DWLNBfiles/lastused.txt" For Input As #fnum1
On Error Resume Next

Input #fnum1, lastused(1)
Input #fnum1, lastused(2)
Input #fnum1, lastused(3)
Close fnum1

flastfile(1).Caption = lastused(1)
flastfile(2).Caption = lastused(2)
flastfile(3).Caption = lastused(3)


End Sub



Private Sub mnuExit_Click()

Unload Textview

End Sub


Private Sub mnuClear_Click()

rtb1.Text = ""

End Sub

Private Sub mnuRich_Click() 'input rt file



dlgfile.Filter = "Text Files  *.rtf  *.txt|*.rtf;*.txt|*;"
dlgfile.filename = ""
dlgfile.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames
    
    On Error Resume Next
    dlgfile.ShowOpen
    If Err.Number = cdlCancel Then Exit Sub
    If Err.Number <> 0 Then
        MsgBox "Error" & Str$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
    End If
    On Error GoTo 0

textfilename = dlgfile.filename

If textfilename <> "" Then
On Error Resume Next

lf = lf + 1
If lf > 3 Then
lf = 1
End If

flastfile(lf).Caption = textfilename

  fnum1 = FreeFile
On Error Resume Next
Open "C:/DWLNBfiles/lastused.txt" For Output As #fnum1
On Error Resume Next

Write #fnum1, flastfile(1).Caption
Write #fnum1, flastfile(2).Caption
Write #fnum1, flastfile(3).Caption
Close fnum1


rtb1.LoadFile (dlgfile.filename)
rtb1.SetFocus
rtb1.Refresh



End If


End Sub


Private Sub mnuSelectall_click()

rtb1.SelStart = 0
rtb1.SelLength = Len(rtb1.Text)

End Sub


Private Sub rtb1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 2 Then
PopupMenu Textview.mnupopup
End If

End Sub


Private Sub mnucopy_click() 'copy

On Error Resume Next

Clipboard.SetText Textview.rtb1.SelRTF

End Sub


Private Sub flastfile_Click(index As Integer)


textfilename = flastfile(index).Caption

If textfilename <> "" Then
On Error Resume Next
rtb1.LoadFile (textfilename)
rtb1.SetFocus
rtb1.Refresh
End If

End Sub
