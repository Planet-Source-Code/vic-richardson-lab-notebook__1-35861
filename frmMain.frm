VERSION 5.00
Begin VB.Form frmimgload 
   AutoRedraw      =   -1  'True
   Caption         =   "Image Loader"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11895
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll 
      Height          =   6060
      Left            =   11535
      TabIndex        =   10
      Top             =   1080
      Width           =   285
   End
   Begin VB.PictureBox Picture2 
      ClipControls    =   0   'False
      Height          =   6300
      Left            =   5595
      ScaleHeight     =   416
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   388
      TabIndex        =   8
      Top             =   810
      Width           =   5880
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5865
         Left            =   75
         ScaleHeight     =   391
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   379
         TabIndex        =   9
         Top             =   60
         Width           =   5685
         Begin VB.Shape Pic2 
            BackColor       =   &H80000001&
            BorderWidth     =   5
            FillColor       =   &H00FFFFFF&
            Height          =   840
            Index           =   0
            Left            =   315
            Top             =   4425
            Width           =   1695
         End
         Begin VB.Image Img 
            Height          =   510
            Index           =   0
            Left            =   3105
            Top             =   4575
            Width           =   1275
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   4020
            Left            =   15
            Top             =   15
            Width           =   5100
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "File Selection"
      Height          =   7080
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3075
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   135
         TabIndex        =   2
         Top             =   225
         Width           =   2880
      End
      Begin VB.DirListBox Dir1 
         Height          =   6165
         Left            =   180
         TabIndex        =   1
         Top             =   660
         Width           =   2835
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Path"
      Height          =   675
      Left            =   5490
      TabIndex        =   5
      Top             =   60
      Width           =   6045
      Begin VB.TextBox txtFileSelected 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   105
         TabIndex        =   7
         Top             =   210
         Width           =   5820
      End
   End
   Begin VB.Frame Filtro 
      Caption         =   "Filter Extension"
      Height          =   7050
      Left            =   3270
      TabIndex        =   3
      Top             =   60
      Width           =   2220
      Begin VB.FileListBox File1 
         Height          =   6135
         Left            =   120
         TabIndex        =   6
         Top             =   615
         Width           =   2040
      End
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2040
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Left click on image filename to preview, Right click to load into notebook."
      Height          =   315
      Left            =   270
      TabIndex        =   11
      Top             =   7215
      Width           =   5415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoadit 
         Caption         =   "Load"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewActualSize 
         Caption         =   "Actual Size"
      End
      Begin VB.Menu mnuViewAdjusted 
         Caption         =   "Adjusted to Window"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmimgload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'      Author: Ramon Antonio Gimenez (ramtonio@yahoo.com)
'                  Formosa  Argentina
'                Created  14-Dec-01
'     Modified by: Vic Richardson for DWL use  6-June-02
'

Option Explicit

Private AppPath                As String
Private FileSelected           As String
Private filename               As String
Private ViewSize               As Integer
Private mPath                  As String
Private Enum ViewSizes
    RealSize
    AdjustedSize
    ZoomedSize
    Thumbnailed
End Enum
Dim Cargado                    As Boolean
Dim ImagenesCargadas           As Boolean
Private mAnchoPic              As Integer
Private mAnchoCelda            As Integer
Private mmargenizq             As Integer
Private mmargensup             As Integer
Private mSep                   As Integer
Private mNroPics               As Integer
Private mNroCols               As Integer
Private mNroFilas              As Integer
Private mAnchoLibre            As Integer
Private mThumbnailView         As Boolean
 


Public Function CalcularNroFilas() As Integer
    On Error GoTo errorHandler
    'Calculate the number of rows
    'in thumbnail view.
    CalcularNroFilas = Int(mNroPics / mNroCols) + 1
    
    Exit Function
errorHandler:
    MsgBox "Error en frmMain.CalcularNroFilas ; " & Err.Number & vbCrLf & _
           Err.Description
End Function


Public Function CalcularNroCols() As Integer
    On Error GoTo errorHandler
    'Calculate the number of columns
    'in thumbnail view.
    CalcularNroCols = Int(Picture1.Width / mAnchoCelda)

    Exit Function
errorHandler:
    MsgBox "Error en frmMain.CalcularNroCols ; " & Err.Number & vbCrLf & _
           Err.Description
End Function


Public Sub CargarPics()
    On Error GoTo errorHandler
    'Load the images
    Dim i As Integer
    
    For i = 1 To mNroPics - 1
        Load Pic2(i)
        Load Img(i)
    Next i

    Exit Sub
errorHandler:
    MsgBox "Error en frmMain.CargarPics ; " & Err.Number & vbCrLf & _
           Err.Description
End Sub



Public Sub UbicarPics()
    On Error GoTo errorHandler
    'place the pics
    Dim i As Integer, j As Integer, n As Integer
    
    mAnchoLibre = Picture1.Width - mAnchoCelda * mNroCols 'Free width
    mmargenizq = mAnchoLibre / 2 'Left margin
    mmargensup = 10     'Top margen
    
    Picture1.Height = mNroFilas * mAnchoCelda
    VScroll.Max = Picture1.Height - Picture2.Height + 10
    n = 0
    For i = 0 To mNroFilas - 1
        For j = 0 To mNroCols - 1
            Pic2(n).Left = j * mAnchoCelda + mmargenizq
            Pic2(n).Top = i * mAnchoCelda + mmargensup
            
            If n = mNroPics - 1 Then
                Exit Sub
            End If
            
            n = n + 1
        Next j
    Next i
    

    Exit Sub
errorHandler:
    MsgBox "Error en frmMain.UbicarPics ; " & Err.Number & vbCrLf & _
           Err.Description
End Sub




Private Sub CargarImagenes()
    On Error GoTo errorHandler
    Dim ratio As Double
    Dim i As Integer
        
    'Load the image files
    For i = 0 To mNroPics - 1
        Img(i).Stretch = True
        Img(i).Picture = LoadPicture(mPath & File1.List(i))
        ratio = Img(i).Picture.Width / Img(i).Picture.Height

        If Img(i).Picture.Height >= Img(i).Picture.Width Then
            Img(i).Height = Pic2(i).Height
            Img(i).Width = Pic2(i).Height * ratio
        Else
            Img(i).Width = Pic2(i).Width
            Img(i).Height = Pic2(i).Width / ratio
        End If


        Img(i).Left = Pic2(i).Left + (Pic2(i).Width - Img(i).Width) / 2
        Img(i).Top = Pic2(i).Top + (Pic2(i).Height - Img(i).Height) / 2
        
    Next i

    Exit Sub
errorHandler:
    MsgBox "Error en frmMain.CargarImagenes ; " & Err.Number & vbCrLf & _
           Err.Description
End Sub


Public Sub DisplayPics()
    On Error GoTo errorHandler

    Dim i As Integer
    Dim n As Integer
    n = mNroPics
    For i = 0 To n - 1
        Pic2(i).Visible = True
        Img(i).Visible = True
    Next i

    Exit Sub
errorHandler:
    MsgBox "Error en frmMain.DisplayPics ; " & Err.Number & vbCrLf & _
           Err.Description
End Sub



Private Sub DescargarPics()
    'Unload image controls
    On Error GoTo errorHandler

    Dim i As Integer
    
    For i = 1 To mNroPics - 1
        Unload Pic2(i)
        Unload Img(i)
'        Img(i).Picture = LoadPicture()
    Next i
    
    Pic2(0).Visible = False
    Img(0).Visible = False
    
    Exit Sub
errorHandler:
    MsgBox "Error en frmMain.DescargarPics ; " & Err.Number & vbCrLf & _
           Err.Description
End Sub


Private Sub Img_Click(index As Integer)
    On Error GoTo errorHandler

    Dim i As Integer
    
    For i = 0 To mNroPics - 1
        Pic2(i).BorderWidth = 3
        Pic2(i).BorderColor = vbBlack
    Next i
    
    Pic2(index).BorderWidth = 5
    Pic2(index).BorderColor = vbRed

    Exit Sub
errorHandler:
    MsgBox "Error en frmMain.Img_Click ; " & Err.Number & vbCrLf & _
           Err.Description
End Sub


Private Sub VScroll_Change()

    Picture1.Top = -VScroll.Value
    
End Sub


Private Sub cboFiltro_Click()
    On Error GoTo errorHandler
    
    Dim pattern As String
    
    If cboFiltro.Text = "All image files" Then
        pattern = "*.jpg;*.jpeg;*.bmp;*.gif;*.ico;*.wmf"
    Else
        pattern = cboFiltro.Text
    End If
    File1.pattern = pattern

    Exit Sub
errorHandler:
    MsgBox "Error en frmMain.cboFiltro_Click ; " & Err.Number & vbCrLf & _
           Err.Description
End Sub


Private Sub cmdAdjustInWindow_Click()

    mnuViewAdjusted_Click
    
End Sub


Public Sub cmdAvanzar_Click()
    On Error GoTo errorHandler

    If File1.ListIndex < File1.ListCount - 1 Then
        Image1.Visible = False
        File1.ListIndex = File1.ListIndex + 1
        Image1.Visible = True
                
    End If
    
    Exit Sub
errorHandler:
    MsgBox "Error en frmMain.cmdAvanzar_Click ; " & Err.Number & vbCrLf & _
           Err.Description
End Sub


Private Sub cmdExit_Click()
    mnuFileExit_Click
End Sub


Public Sub cmdPrimero_Click()

    File1.ListIndex = 0
    cmdRetroceder_Click
    
End Sub


Private Sub cmdRealSize_Click()

    mnuViewActualSize_Click
    
End Sub


Public Sub cmdRetroceder_Click()

    If File1.ListIndex > 0 Then
        Image1.Visible = False
        File1.ListIndex = File1.ListIndex - 1
        Image1.Visible = True
       
        
    End If
    
End Sub


Public Sub cmdUltimo_Click()

    File1.ListIndex = File1.ListCount - 1
    cmdAvanzar_Click
    
   

End Sub


Private Sub Dir1_Change()
    On Error GoTo errorHandler
    
    File1.Path = Dir1.Path
    Picture1.Visible = False
    Image1.Visible = False
    
    If Right(File1.Path, 1) <> "\" Then
        mPath = File1.Path & "\"
    Else
        mPath = File1.Path
    End If
  
    If File1.ListCount <> 0 Then
        mNroPics = File1.ListCount
    Else
        mNroPics = 0
        Image1.Visible = False
        Picture1.Visible = False
        FileSelected = ""
        Exit Sub
    End If

    
    If mNroPics > 1 Then
        ReDim Img(mNroPics - 1)
    End If
    
    If mNroPics = 0 Then
        Picture1.Visible = False
       
       
    Else
       
        
        File1.ListIndex = 0
        txtFileSelected.Text = mPath & File1.List(0)
        FileSelected = File1.List(0)
        Image1.Picture = LoadPicture(mPath & File1.List(0))

      
    End If
    
    
    Exit Sub
errorHandler:
    'MsgBox "Error en frmMain.Dir1_Change ; " & Err.Number & vbCrLf & Err.Description

End Sub


Private Sub Drive1_Change()
    
    Dir1.Path = Drive1.Drive

End Sub


Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X2 As Single, Y2 As Single)

    On Error GoTo errorHandler
    
    If mNroPics = 0 Then
       
        Picture1.Visible = False
        Exit Sub
    End If
        
        filename = mPath & File1.List(File1.ListIndex)
        txtFileSelected.Text = filename
        Image1.Picture = LoadPicture(mPath & File1.List(File1.ListIndex))

    If ViewSize = ViewSizes.AdjustedSize Then
        Dim ratio As Double
        Picture1.Visible = False
        Image1.Visible = False
        With Image1
            If .Picture.Height >= .Picture.Width Then
                .Height = Picture1.Height
                ratio = .Picture.Width / .Picture.Height
                .Width = .Height * ratio
            Else
                .Width = Picture1.Width
                ratio = .Picture.Height / .Picture.Width
                .Height = .Width * ratio
            End If
            .Stretch = True
            .Visible = True
            Picture1.Visible = True
        End With
    ElseIf ViewSize = ViewSizes.RealSize Then
        Picture1.Visible = True
        Image1.Stretch = False
        Image1.Visible = True
    End If
    
    
    If Button = 2 Then
    mnuLoadit_Click
    End If
    
    FileSelected = ""
    cboFiltro_Click
    
    Exit Sub
errorHandler:
    'MsgBox "Error en frmMain.File1_Click ; " & Err.Number & vbCrLf & Err.Description

End Sub


Private Sub File1a_Click()


    On Error GoTo errorHandler
    
    If mNroPics = 0 Then
       
        Picture1.Visible = False
        Exit Sub
    End If
        
        filename = mPath & File1.List(File1.ListIndex)
        txtFileSelected.Text = filename
        Image1.Picture = LoadPicture(mPath & File1.List(File1.ListIndex))

    If ViewSize = ViewSizes.AdjustedSize Then
        Dim ratio As Double
        Picture1.Visible = False
        Image1.Visible = False
        With Image1
            If .Picture.Height >= .Picture.Width Then
                .Height = Picture1.Height
                ratio = .Picture.Width / .Picture.Height
                .Width = .Height * ratio
            Else
                .Width = Picture1.Width
                ratio = .Picture.Height / .Picture.Width
                .Height = .Width * ratio
            End If
            .Stretch = True
            .Visible = True
            Picture1.Visible = True
        End With
    ElseIf ViewSize = ViewSizes.RealSize Then
        Picture1.Visible = True
        Image1.Stretch = False
        Image1.Visible = True
    End If
    
    FileSelected = ""
    cboFiltro_Click
    
    Exit Sub
errorHandler:
    'MsgBox "Error en frmMain.File1_Click ; " & Err.Number & vbCrLf & Err.Description

End Sub


Private Sub Form_Load()
    On Error GoTo errorHandler
    
    ScaleMode = vbPixels
    Picture2.ScaleMode = vbPixels
    Picture1.ScaleMode = vbPixels
    Picture1.Top = 0
    Picture1.Left = 0
    Picture1.Width = Picture2.Width
    Picture1.Height = Picture2.Height
    
    Dir1.Path = App.Path
    Drive1.Drive = Left(App.Path, 3)
    File1.Path = Dir1.Path
    
    Pic2(0).Visible = False
                               
    
    mAnchoPic = 75
    mSep = 10
    
    Pic2(0).Width = mAnchoPic
    Pic2(0).Height = mAnchoPic
    Pic2(0).Visible = False
    Img(0).Visible = False
    
    With VScroll
        .Min = Picture1.Top
        .Max = Picture1.Height
        .SmallChange = 0.1 * (.Max - .Min)
        .LargeChange = 0.3 * (.Max - .Min)
        .Visible = False
    End With

    With cboFiltro
        
        .AddItem "All image files", 0
        .AddItem "*.jpg;*.jpeg", 1
        .AddItem "*.bmp", 2
        .AddItem "*.gif", 3
        .AddItem "*.ico", 4
        .AddItem "*.wmf", 5
        
        .Text = .List(0)
    End With
    
  
    Dim pattern As String
    If cboFiltro.Text = "All image files" Then
        pattern = "*.jpg;*.jpeg;*.bmp;*.gif;*.ico;*.wmf"
    Else
        pattern = cboFiltro.Text
    End If
    File1.pattern = pattern
    ViewSize = ViewSizes.AdjustedSize

    Image1.Stretch = True
    Image1.Visible = True
    Image1.Left = 0
    Image1.Top = 0
    
    Dir1_Change
    
    Exit Sub
errorHandler:
    MsgBox "Error en frmMain.Form_Load ; " & Err.Number & vbCrLf & Err.Description

End Sub


Private Sub mnuFileExit_Click()

   
    Unload Me
    
End Sub


Private Sub mnuViewActualSize_Click()

    mnuViewActualSize.Checked = True
    mnuViewAdjusted.Checked = False
    
    ViewSize = ViewSizes.RealSize
    Image1.Visible = False
    File1a_Click
    
End Sub


Private Sub mnuViewAdjusted_Click()

    mnuViewActualSize.Checked = False
    mnuViewAdjusted.Checked = True

    ViewSize = ViewSizes.AdjustedSize
    Image1.Visible = False
    File1a_Click

    
End Sub


Private Sub mnuLoadit_Click() 'send "filename" to notebook

FormNB.imgfilename = filename
FormNB.PlaceImage


End Sub
