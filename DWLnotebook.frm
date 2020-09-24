VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FormNB 
   AutoRedraw      =   -1  'True
   Caption         =   "EL-7000  Notebook"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   13110
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DWLnotebook.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10530
   ScaleWidth      =   13110
   Begin VB.Frame Frame1 
      Caption         =   "Copy Page"
      Height          =   825
      Left            =   7230
      TabIndex        =   66
      Top             =   285
      Visible         =   0   'False
      Width           =   1665
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1035
         TabIndex        =   68
         Text            =   "1"
         Top             =   360
         Width           =   510
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   67
         Text            =   "1"
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "to"
         Height          =   270
         Left            =   660
         TabIndex        =   69
         Top             =   405
         Width           =   345
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   4605
      TabIndex        =   50
      Top             =   120
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   35
      Left            =   7290
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   49
      Top             =   7035
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   34
      Left            =   6075
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   48
      Top             =   8010
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   33
      Left            =   6030
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   47
      Top             =   7500
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   32
      Left            =   6015
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   46
      Top             =   7035
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   31
      Left            =   4770
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   45
      Top             =   8010
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   30
      Left            =   4770
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   44
      Top             =   7545
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   29
      Left            =   4740
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   43
      Top             =   7080
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   28
      Left            =   3360
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   42
      Top             =   7905
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   27
      Left            =   3330
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   41
      Top             =   7410
      Width           =   810
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   10
      Left            =   6840
      ScaleHeight     =   615
      ScaleWidth      =   765
      TabIndex        =   40
      Top             =   5220
      Width           =   765
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   9
      Left            =   6825
      ScaleHeight     =   615
      ScaleWidth      =   765
      TabIndex        =   39
      Top             =   4320
      Width           =   765
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   8
      Left            =   6780
      ScaleHeight     =   615
      ScaleWidth      =   765
      TabIndex        =   38
      Top             =   3390
      Width           =   765
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   7
      Left            =   6795
      ScaleHeight     =   615
      ScaleWidth      =   765
      TabIndex        =   37
      Top             =   2490
      Width           =   765
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   6
      Left            =   6825
      ScaleHeight     =   615
      ScaleWidth      =   765
      TabIndex        =   36
      Top             =   1545
      Width           =   765
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   5
      Left            =   5445
      ScaleHeight     =   615
      ScaleWidth      =   765
      TabIndex        =   35
      Top             =   6060
      Width           =   765
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   4
      Left            =   5460
      ScaleHeight     =   615
      ScaleWidth      =   765
      TabIndex        =   34
      Top             =   5145
      Width           =   765
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   3
      Left            =   5430
      ScaleHeight     =   615
      ScaleWidth      =   765
      TabIndex        =   33
      Top             =   4260
      Width           =   765
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   2
      Left            =   5355
      ScaleHeight     =   615
      ScaleWidth      =   765
      TabIndex        =   32
      Top             =   3345
      Width           =   765
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   1
      Left            =   5370
      ScaleHeight     =   615
      ScaleWidth      =   765
      TabIndex        =   31
      Top             =   2490
      Width           =   765
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   0
      Left            =   5370
      ScaleHeight     =   615
      ScaleWidth      =   765
      TabIndex        =   30
      Top             =   1575
      Width           =   765
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   26
      Left            =   3315
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   29
      Top             =   6930
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   25
      Left            =   3345
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   28
      Top             =   6420
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   24
      Left            =   3360
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   27
      Top             =   6000
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   23
      Left            =   3300
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   26
      Top             =   5565
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   22
      Left            =   3405
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   25
      Top             =   5070
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   21
      Left            =   3360
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   24
      Top             =   4590
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   20
      Left            =   3270
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   23
      Top             =   4065
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   19
      Left            =   3240
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   22
      Top             =   3555
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   18
      Left            =   3300
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   21
      Top             =   3075
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   17
      Left            =   3285
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   20
      Top             =   2610
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   16
      Left            =   3120
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   19
      Top             =   2145
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   15
      Left            =   3180
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   18
      Top             =   1635
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   14
      Left            =   3060
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   17
      Top             =   1095
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   13
      Left            =   1875
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   16
      Top             =   8220
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   12
      Left            =   1920
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   15
      Top             =   7605
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   11
      Left            =   1905
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   14
      Top             =   7020
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   10
      Left            =   1890
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   13
      Top             =   6525
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   9
      Left            =   1860
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   12
      Top             =   5955
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   8
      Left            =   1890
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   11
      Top             =   5370
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   7
      Left            =   1950
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   10
      Top             =   4740
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   6
      Left            =   1965
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   9
      Top             =   4185
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   5
      Left            =   1890
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   8
      Top             =   3690
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   4
      Left            =   1860
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   7
      Top             =   3165
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   3
      Left            =   1875
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   6
      Top             =   2580
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   2
      Left            =   1875
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   5
      Top             =   2025
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   1
      Left            =   1830
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   4
      Top             =   1530
      Width           =   810
   End
   Begin VB.PictureBox Picture3 
      Height          =   285
      Index           =   0
      Left            =   1890
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   3
      Top             =   1035
      Width           =   810
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   1455
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   135
      Width           =   9750
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   915
      Top             =   180
   End
   Begin MSComDlg.CommonDialog dlgfile 
      Left            =   720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   10470
      Left            =   45
      Picture         =   "DWLnotebook.frx":030A
      ScaleHeight     =   10410
      ScaleWidth      =   12120
      TabIndex        =   1
      Top             =   30
      Width           =   12180
      Begin RichTextLib.RichTextBox rtb1 
         Height          =   9690
         Left            =   780
         TabIndex        =   51
         Top             =   690
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   17092
         _Version        =   393217
         BorderStyle     =   0
         Appearance      =   0
         TextRTF         =   $"DWLnotebook.frx":CB36
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11055
         TabIndex        =   2
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "    F1 =        Undo                      F9 = Indent"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   12330
      TabIndex        =   65
      Top             =   5325
      Width           =   735
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CENTER"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12330
      TabIndex        =   64
      Top             =   2895
      Width           =   735
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BLUE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12315
      TabIndex        =   63
      Top             =   4350
      Width           =   735
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RED"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12315
      TabIndex        =   62
      Top             =   4710
      Width           =   735
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ITALIC"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12330
      TabIndex        =   61
      Top             =   3930
      Width           =   735
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ULINE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12330
      TabIndex        =   60
      Top             =   3585
      Width           =   735
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BOLD"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12330
      TabIndex        =   59
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Char Pos"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   12285
      TabIndex        =   58
      Top             =   2235
      Width           =   765
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12285
      TabIndex        =   57
      Top             =   2460
      Width           =   780
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "LINES"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12360
      TabIndex        =   56
      Top             =   225
      Width           =   705
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12330
      TabIndex        =   55
      Top             =   1335
      Width           =   750
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Current"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   12330
      TabIndex        =   54
      Top             =   615
      Width           =   750
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   12315
      TabIndex        =   53
      Top             =   1605
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   12300
      TabIndex        =   52
      Top             =   855
      Width           =   795
   End
   Begin VB.Menu mnuSpacer12 
      Caption         =   ""
   End
   Begin VB.Menu mnuSpacer11 
      Caption         =   ""
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLoadbook 
         Caption         =   "Load Notebook"
      End
      Begin VB.Menu mnuSavebook 
         Caption         =   "Save Notebook"
      End
      Begin VB.Menu mnuNewbook 
         Caption         =   "New Notebook"
      End
      Begin VB.Menu mnuTextview 
         Caption         =   "Textfile Viewer"
      End
      Begin VB.Menu mnuSavetext 
         Caption         =   "Save Text on Page"
      End
      Begin VB.Menu mnuPagesetup 
         Caption         =   "Print Page Setup"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Page Text only"
      End
      Begin VB.Menu mnuPrintform 
         Caption         =   "Print Form"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo text change"
      End
      Begin VB.Menu mniImage 
         Caption         =   "Insert Image"
      End
      Begin VB.Menu mnuDelimage 
         Caption         =   "Delete Image (Rt Click)"
      End
      Begin VB.Menu mnuEquation 
         Caption         =   "Equation formatter"
      End
      Begin VB.Menu mniInspage 
         Caption         =   "Insert Page"
      End
      Begin VB.Menu mnuCopypage 
         Caption         =   "Copy Page"
      End
      Begin VB.Menu mnuDelpage 
         Caption         =   "Delete Page"
      End
      Begin VB.Menu mnuClearpage 
         Caption         =   "Clear Page"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuCalcassoc 
         Caption         =   "Calculator association"
      End
      Begin VB.Menu mnuApp1assoc 
         Caption         =   "User App1 association"
      End
      Begin VB.Menu mnuApp2assoc 
         Caption         =   "User App2 association"
      End
      Begin VB.Menu mnuMargins 
         Caption         =   "Margins"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
   Begin VB.Menu mnuSpacer9 
      Caption         =   ""
   End
   Begin VB.Menu mnuSpacer3 
      Caption         =   ""
   End
   Begin VB.Menu mnuSpacer1 
      Caption         =   " "
   End
   Begin VB.Menu mnuPageDown 
      Caption         =   "< Page"
   End
   Begin VB.Menu mnuGoTo 
      Caption         =   "GOTO"
   End
   Begin VB.Menu mnuPageUp 
      Caption         =   "Page >"
   End
   Begin VB.Menu mnuspacer2 
      Caption         =   ""
   End
   Begin VB.Menu mnuTOC 
      Caption         =   "TOC"
   End
   Begin VB.Menu dummy1 
      Caption         =   ""
   End
   Begin VB.Menu dummy2 
      Caption         =   ""
   End
   Begin VB.Menu mnuSpacer4 
      Caption         =   ""
   End
   Begin VB.Menu mnuCalculator 
      Caption         =   "Calculator"
   End
   Begin VB.Menu dd6 
      Caption         =   ""
   End
   Begin VB.Menu mnuApp1 
      Caption         =   "User App1"
   End
   Begin VB.Menu dd7 
      Caption         =   ""
   End
   Begin VB.Menu mnuApp2 
      Caption         =   "User App2"
   End
   Begin VB.Menu mnupopup 
      Caption         =   "popuprtclick"
      Visible         =   0   'False
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuSelectall 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuIndent 
         Caption         =   "Indent (F9)"
      End
      Begin VB.Menu mnuBold 
         Caption         =   "Bold (F10)"
      End
      Begin VB.Menu mnuUnderline 
         Caption         =   "Underline (F11)"
      End
      Begin VB.Menu mnuItalics 
         Caption         =   "Italics (F12)"
      End
      Begin VB.Menu mnuHighred 
         Caption         =   "Highlight Red"
      End
      Begin VB.Menu mnuHighblue 
         Caption         =   "Highlight Blue"
      End
      Begin VB.Menu mnuCenter 
         Caption         =   "Center Align"
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "Left Align"
      End
   End
End
Attribute VB_Name = "FormNB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'  DazyWeb Laboratories Notebook   17-June-02  rev 1.06
'
'

Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long

Private Const EM_GETLINECOUNT = &HBA
Private Const EM_GETLINE = &HC4
Private Const EM_LINEFROMCHAR = &HC9

Dim copyfrompage As Integer
Dim copytopage As Integer
Dim charpos As Long
Dim CurrentLine2 As Long
Dim Totallines As Long
Dim Topline As Long
Dim app1assoc As String
Dim app2assoc As String
Dim ret As Variant
Dim calcassoc As String
Dim undoflag As Integer
Dim cc As Integer
Dim textundo(10) As String
Dim nextpage As Integer
Dim chunk As Integer
Dim temphold As Integer
Dim refstart As Integer
Dim lastone As Integer
Dim notemptyflag As Integer
Dim temptext$
Dim textarray$(100)
Dim textarray2$(30)
Dim pagetitle(100) As String
Dim imagename(100, 10) As String
Dim imagecoordX(100, 10) As Long
Dim imagecoordY(100, 10) As Long
Dim pageimages(100) As Integer
Dim fnum1 As Integer
Dim currentpage As Integer
Dim numpages As Integer
Dim fonttype As String
Dim fontcolor As String
Dim xx As Long
Dim yy As Long
Dim i As Integer
Dim j As Long
Dim tt As Long
Dim cntr As Integer
Dim indexflag As Integer
Dim TOCpage As Integer
Dim bookfilename As String
Dim bookfilepath As String
Dim CurrentLine As Integer
Dim stillloading As Integer
Dim holdoff As Integer
Public imgfilename$
Dim cursorX As Long
Dim cursorY As Long
Dim delimgflag As Integer
Dim pwid As Single
Dim phgt As Single
Dim xmid As Single
Dim ymid As Single
Dim screensize As Integer
Dim filename As String
Dim newfilename As String
Dim txt As String
Dim endflag As Integer
Dim boldflag As Integer
Dim underlineflag As Integer
Dim italicsflag As Integer
Dim highredflag As Integer
Dim highblueflag As Integer
Dim centerflag As Integer
Dim marginflag As Integer


Private Sub Form_Load()

Text1.FontSize = 17
Text1.FontName = "Times New Roman"
Text1.Height = 400
Text1.Top = 150


rtb1.SelColor = vbBlack
rtb1.SelFontSize = 12
rtb1.SelFontName = "Times New Roman"
rtb1.SelBold = False
rtb1.Top = 350 + 300

cc = 0

fonttype = "Times New Roman"
fontcolor = vbBlack

stillloading = 1

  On Error Resume Next
MkDir "c:\DWLNBfiles"
On Error Resume Next

  fnum1 = FreeFile
On Error Resume Next
Open "c:\DWLNBfiles\notebookini" For Input As #fnum1
On Error Resume Next
Input #fnum1, currentpage
Input #fnum1, numpages
Input #fnum1, fonttype
Input #fnum1, fontcolor
For xx = 1 To 100
For yy = 1 To 10
Input #fnum1, imagename(xx, yy)
Next yy
Next xx
For xx = 1 To 100
For yy = 1 To 10
Input #fnum1, imagecoordX(xx, yy)
Next yy
Next xx
For xx = 1 To 100
For yy = 1 To 10
Input #fnum1, imagecoordY(xx, yy)
Next yy
Next xx
For yy = 1 To 100
Input #fnum1, pageimages(yy)
Next yy
Input #fnum1, calcassoc
For xx = 1 To 100
Input #fnum1, pagetitle(xx)
Next xx
Input #fnum1, app1assoc
Input #fnum1, app2assoc
Input #fnum1, marginflag
Close fnum1

If marginflag = 1 Then
mnuMargins.Checked = True
rtb1.Left = 1280
rtb1.Width = 10295
Else
mnuMargins.Checked = False
rtb1.Left = 780
rtb1.Width = 11295
End If

If calcassoc = "" Then
calcassoc = "C:\WINDOWS\CALC.EXE"
End If

If currentpage = 0 Then
currentpage = 1
End If

If numpages = 0 Then
numpages = 1
End If

If fontcolor = "" Then
fontcolor = vbBlack
End If


On Error Resume Next
filename = "c:\DWLNBfiles\page" + CStr(currentpage) + ".nbt"
On Error Resume Next
rtb1.LoadFile (filename)
rtb1.SetFocus
rtb1.Refresh

Text1.Text = pagetitle(currentpage)
Label1.Caption = CStr(currentpage) + "/" + CStr(numpages)

For xx = 0 To 10
Picture2(xx).Visible = False
Next xx

For xx = 1 To pageimages(currentpage)
If imagename(currentpage, xx) <> "" Then
Picture2(xx).Visible = True
Picture2(xx).ZOrder (0)
Picture2(xx).Picture = LoadPicture(imagename(currentpage, xx))
Picture2(xx).Left = imagecoordX(currentpage, xx)
Picture2(xx).Top = imagecoordY(currentpage, xx)
End If
Next xx


For xx = 0 To 35

Picture3(xx).ForeColor = "&HFFcccc"
Picture3(xx).BackColor = "&HFFcccc"
Picture3(xx).Height = 15
Picture3(xx).Width = 11400
Picture3(xx).Top = 700 + (xx * 285)
Picture3(xx).Left = 800
Picture3(xx).Refresh

Next xx


stillloading = 0

rtb1.SetFocus

End Sub


Private Sub mniInspage_Click() 'insert page


If numpages < 100 Then


i = currentpage

If numpages = 100 Then
numpages = 99
End If

For xx = i To numpages
On Error Resume Next
pagetitle(numpages - xx + i + 1) = pagetitle(numpages - xx + i)
'rename files here
filename = "c:\DWLNBfiles\page" + CStr(numpages - xx + i) + ".nbt"
newfilename = "c:\DWLNBfiles\page" + CStr(numpages - xx + i + 1) + ".nbt"
Name filename As newfilename
Next xx

On Error Resume Next
Kill filename

rtb1.Text = ""
resetflags

For xx = i To 99
For yy = 1 To 10
imagename(99 - xx + i + 1, yy) = imagename(99 - xx + i, yy)
Next yy
Next xx

For xx = i To 99
For yy = 1 To 10
imagecoordX(99 - xx + i + 1, yy) = imagecoordX(99 - xx + i, yy)
Next yy
Next xx

For xx = i To 99
For yy = 1 To 10
imagecoordY(99 - xx + i + 1, yy) = imagecoordY(99 - xx + i, yy)
Next yy
Next xx

For xx = i To 99
pageimages(99 - xx + i + 1) = pageimages(99 - xx + i)
Next xx


currentpage = i
mnuClearpage_Click


numpages = numpages + 1

Label1.Caption = CStr(currentpage) + "/" + CStr(numpages)
End If

End Sub


Private Sub mnuCopypage_Click() 'copy page

Frame1.Visible = True

End Sub


Private Sub Text4_KeyDown(keycode As Integer, Shift As Integer) 'act on copy page return

If keycode = vbKeyReturn Then
copyfrompage = Abs(Val(Text2.Text))
copytopage = Abs(Val(Text4.Text))
Frame1.Visible = False

If copyfrompage = copytopage Then
Exit Sub
End If

If copyfrompage > numpages Or copytopage > numpages Then
Exit Sub
End If

On Error Resume Next
filename = "c:\DWLNBfiles\page" + CStr(copyfrompage) + ".nbt"
newfilename = "c:\DWLNBfiles\page" + CStr(copytopage) + ".nbt"
Kill newfilename
On Error Resume Next
ret = FileCopy(filename, newfilename)

pagetitle(copytopage) = pagetitle(copyfrompage)
For yy = 1 To 10
imagename(copytopage, yy) = imagename(copyfrompage, yy)
Next yy
For yy = 1 To 10
imagecoordX(copytopage, yy) = imagecoordX(copyfrompage, yy)
Next yy
For yy = 1 To 10
imagecoordY(copytopage, yy) = imagecoordY(copyfrompage, yy)
Next yy
pageimages(copytopage) = pageimages(copyfrompage)

Update

End If

End Sub


Private Sub mnuDelpage_Click() 'del page

If numpages > 0 Then

mnuClearpage_Click

i = currentpage

For xx = i To numpages
On Error Resume Next
pagetitle(xx) = pagetitle(xx + 1)
Next xx

Text1.Text = pagetitle(i)
filename = "c:\DWLNBfiles\page" + CStr(i) + ".nbt"
Kill filename

For xx = i To numpages
filename = "c:\DWLNBfiles\page" + CStr(xx + 1) + ".nbt"
newfilename = "c:\DWLNBfiles\page" + CStr(xx) + ".nbt"
'On Error Resume Next
Name filename As newfilename
Next xx

On Error Resume Next


If i < 100 Then 'redisplay current page
On Error Resume Next
currentpage = i
filename = "c:\DWLNBfiles\page" + CStr(currentpage) + ".nbt"
On Error Resume Next
rtb1.LoadFile (filename)
rtb1.SetFocus
rtb1.Refresh

End If

For xx = i To 99
For yy = 1 To 10
imagename(xx, yy) = imagename(xx + 1, yy)
Next yy
Next xx

For xx = i To 99
For yy = 1 To 10
imagecoordX(xx, yy) = imagecoordX(xx + 1, yy)
Next yy
Next xx

For xx = i To 99
For yy = 1 To 10
imagecoordY(xx, yy) = imagecoordY(xx + 1, yy)
Next yy
Next xx

For yy = i To 99
pageimages(yy) = pageimages(yy + 1)
Next yy

For xx = 0 To 10
Picture2(xx).Visible = False
Next xx



For xx = 1 To pageimages(currentpage) 'redisplay images for revised currentpage
If imagename(currentpage, xx) <> "" Then
Picture2(xx).Visible = True
Picture2(xx).Picture = LoadPicture(imagename(currentpage, xx))
Picture2(xx).Left = imagecoordX(currentpage, xx)
Picture2(xx).Top = imagecoordY(currentpage, xx)
End If
Next xx

numpages = numpages - 1

Label1.Caption = CStr(currentpage) + "/" + CStr(numpages)
End If

End Sub



Private Sub mniImage_Click() 'insert image at cursor position

frmimgload.Show

End Sub

Public Function PlaceImage()

 If imgfilename$ <> "" And indexflag = 0 Then

If pageimages(currentpage) < 10 Then
pageimages(currentpage) = pageimages(currentpage) + 1
End If

imagename(currentpage, pageimages(currentpage)) = imgfilename$
imagecoordX(currentpage, pageimages(currentpage)) = cursorX
imagecoordY(currentpage, pageimages(currentpage)) = cursorY

Picture2(pageimages(currentpage)).Visible = True
Picture2(pageimages(currentpage)).ZOrder (0)
Picture2(pageimages(currentpage)).Picture = LoadPicture(imgfilename$)  'dlgfile.filename)
Picture2(pageimages(currentpage)).Left = cursorX
Picture2(pageimages(currentpage)).Top = cursorY

End If


End Function


Private Sub mnuClearpage_Click() 'clear page


rtb1.Text = ""
Text1.Text = ""
resetflags

For xx = 1 To pageimages(currentpage)
Picture2(xx).Visible = False
imagename(currentpage, xx) = ""
imagecoordX(currentpage, xx) = 0
imagecoordY(currentpage, xx) = 0
Next xx

End Sub

Private Sub mnuDelimage_Click() 'del image

If delimgflag = 0 Then
delimgflag = 1
Else
delimgflag = 0
End If

End Sub




Private Sub mnuPrintform_Click()  'screen print form

On Error Resume Next
PrintForm
On Error Resume Next

End Sub


Private Sub mnuSavetext_Click() 'save text on page

dlgfile.Filter = "Richtext Files|*.rtf;"
dlgfile.filename = ""
dlgfile.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames
    
    On Error Resume Next
    dlgfile.ShowSave
    If Err.Number = cdlCancel Then Exit Sub
    If Err.Number <> 0 Then
        MsgBox "Error" & Str$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
    End If
    On Error GoTo 0

filename = dlgfile.filename
If filename <> "" Then
On Error Resume Next
rtb1.SaveFile (filename)
End If

End Sub


Private Sub Picture2_Click(index As Integer)


If delimgflag = 1 Then

Picture2(index).Visible = False
imagename(currentpage, index) = ""
imagecoordX(currentpage, index) = 0
imagecoordY(currentpage, index) = 0
delimgflag = 0

'compact array

If pageimages(currentpage) > 0 Then
For xx = index To pageimages(currentpage) - 1
imagename(currentpage, xx) = imagename(currentpage, (xx + 1))
imagecoordX(currentpage, xx) = imagecoordX(currentpage, (xx + 1))
imagecoordY(currentpage, xx) = imagecoordY(currentpage, (xx + 1))
Next xx
imagename(currentpage, pageimages(currentpage)) = ""
imagecoordX(currentpage, pageimages(currentpage)) = 0
imagecoordY(currentpage, pageimages(currentpage)) = 0
pageimages(currentpage) = pageimages(currentpage) - 1
End If

End If


End Sub


Private Sub Form_Unload(Cancel As Integer)

mnuExit_Click
    
End Sub


Private Sub mnuExit_Click()

If indexflag = 1 Then
mnuTOC_Click
End If



pagetitle(currentpage) = Text1.Text


On Error Resume Next
filename = "c:\DWLNBfiles\page" + CStr(currentpage) + ".nbt"
On Error Resume Next
rtb1.SaveFile (filename)

  fnum1 = FreeFile
On Error Resume Next
Open "c:\DWLNBfiles\notebookini" For Output As #fnum1
On Error Resume Next
Write #fnum1, currentpage
Write #fnum1, numpages
Write #fnum1, fonttype
Write #fnum1, fontcolor
For xx = 1 To 100
For yy = 1 To 10
Write #fnum1, imagename(xx, yy)
Next yy
Next xx
For xx = 1 To 100
For yy = 1 To 10
Write #fnum1, imagecoordX(xx, yy)
Next yy
Next xx
For xx = 1 To 100
For yy = 1 To 10
Write #fnum1, imagecoordY(xx, yy)
Next yy
Next xx
For yy = 1 To 100
Write #fnum1, pageimages(yy)
Next yy
Write #fnum1, calcassoc
For xx = 1 To 100
Write #fnum1, pagetitle(xx)
Next xx
Write #fnum1, app1assoc
Write #fnum1, app2assoc
Write #fnum1, marginflag
Close fnum1

Unload frmimgload
Unload Textview
Unload Help
Unload eqtnform
Unload FormNB

End Sub


Private Sub mnuHelp_Click()

Help.Show

End Sub

Private Sub mnuLoadbook_Click()

ChDir ("C:/DWLNBfiles")

dlgfile.Filter = "Notebook Files|*.lnb;"
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

bookfilename = dlgfile.filename

'parse for filepath
For xx = 0 To Len(bookfilename) - 1
 If Mid$(bookfilename, Len(bookfilename) - xx, 1) = "\" Then
 endflag = 1
 bookfilepath = txt
 Else
   If endflag = 0 Then
   txt = Left$(bookfilename, Len(bookfilename) - xx - 1)
   End If
 End If
Next xx
endflag = 0


If bookfilename <> "" Then
On Error Resume Next
stillloading = 1

mnuNewbook_Click

  fnum1 = FreeFile
On Error Resume Next
Open bookfilename For Input As #fnum1
On Error Resume Next
Input #fnum1, currentpage
Input #fnum1, numpages
Input #fnum1, fonttype
Input #fnum1, fontcolor
For xx = 1 To 100
For yy = 1 To 10
Input #fnum1, imagename(xx, yy)
Next yy
Next xx
For xx = 1 To 100
For yy = 1 To 10
Input #fnum1, imagecoordX(xx, yy)
Next yy
Next xx
For xx = 1 To 100
For yy = 1 To 10
Input #fnum1, imagecoordY(xx, yy)
Next yy
Next xx
For yy = 1 To 100
Input #fnum1, pageimages(yy)
Next yy
Input #fnum1, calcassoc
For xx = 1 To 100
Input #fnum1, pagetitle(xx)
Next xx
Input #fnum1, app1assoc
Input #fnum1, app2assoc
Input #fnum1, marginflag
Close fnum1


If marginflag = 1 Then
mnuMargins.Checked = True
rtb1.Left = 1280
rtb1.Width = 10295
Else
mnuMargins.Checked = False
rtb1.Left = 780
rtb1.Width = 11295
End If

If currentpage = 0 Then
currentpage = 1
End If

If numpages = 0 Then
numpages = 1
End If


'use bookfilepath to copy A to B times numpages
For xx = 1 To numpages
On Error Resume Next
newfilename = "c:\DWLNBfiles\page" + CStr(xx) + ".nbt"
filename = bookfilepath + "page" + CStr(xx) + ".nbt"
On Error Resume Next
ret = FileCopy(filename, newfilename)
Next xx

For xx = numpages + 1 To 100  'kill leftovers from last notebook
filename = "c:\DWLNBfiles\page" + CStr(xx) + ".nbt"
On Error Resume Next
Kill filename
Next xx

On Error Resume Next 'load currentpage
filename = "c:\DWLNBfiles\page" + CStr(currentpage) + ".nbt"
On Error Resume Next
rtb1.LoadFile (filename)
rtb1.SetFocus
rtb1.Refresh

Text1.Text = pagetitle(currentpage)
Label1.Caption = CStr(currentpage) + "/" + CStr(numpages)

For xx = 0 To 10
Picture2(xx).Visible = False
Next xx

For xx = 1 To pageimages(currentpage)

If imagename(currentpage, xx) <> "" Then
Picture2(xx).Visible = True
Picture2(xx).ZOrder (0)
Picture2(xx).Picture = LoadPicture(imagename(currentpage, xx))
Picture2(xx).Left = imagecoordX(currentpage, xx)
Picture2(xx).Top = imagecoordY(currentpage, xx)
End If

Next xx


stillloading = 0

End If


End Sub


Private Sub mnuSavebook_Click()

ChDir ("C:/DWLNBfiles")

dlgfile.Filter = "Notebook Files|*.lnb;"
dlgfile.filename = ""
dlgfile.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames
    
    On Error Resume Next
    dlgfile.ShowSave
    If Err.Number = cdlCancel Then Exit Sub
    If Err.Number <> 0 Then
        MsgBox "Error" & Str$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
    End If
    On Error GoTo 0


bookfilename = dlgfile.filename  'parse for filepath
For xx = 0 To Len(bookfilename) - 1
 If Mid$(bookfilename, Len(bookfilename) - xx, 1) = "\" Then
 endflag = 1
 bookfilepath = txt
 Else
   If endflag = 0 Then
   txt = Left$(bookfilename, Len(bookfilename) - xx - 1)
   End If
 End If
Next xx
endflag = 0


If bookfilename <> "" Then
On Error Resume Next



pagetitle(currentpage) = Text1.Text

On Error Resume Next
filename = "c:\DWLNBfiles\page" + CStr(currentpage) + ".nbt"
On Error Resume Next
rtb1.SaveFile (filename)

  fnum1 = FreeFile
On Error Resume Next
Open bookfilename For Output As #fnum1
On Error Resume Next
Write #fnum1, currentpage
Write #fnum1, numpages
Write #fnum1, fonttype
Write #fnum1, fontcolor
For xx = 1 To 100
For yy = 1 To 10
Write #fnum1, imagename(xx, yy)
Next yy
Next xx
For xx = 1 To 100
For yy = 1 To 10
Write #fnum1, imagecoordX(xx, yy)
Next yy
Next xx
For xx = 1 To 100
For yy = 1 To 10
Write #fnum1, imagecoordY(xx, yy)
Next yy
Next xx
For yy = 1 To 100
Write #fnum1, pageimages(yy)
Next yy
Write #fnum1, calcassoc
For xx = 1 To 100
Write #fnum1, pagetitle(xx)
Next xx
Write #fnum1, app1assoc
Write #fnum1, app2assoc
Write #fnum1, marginflag
Close fnum1

'use bookfilepath to copy A to B times numpages
For xx = 1 To numpages
On Error Resume Next
filename = "c:\DWLNBfiles\page" + CStr(xx) + ".nbt"
newfilename = bookfilepath + "page" + CStr(xx) + ".nbt"
On Error Resume Next
ret = FileCopy(filename, newfilename)
Next xx

End If


End Sub


Public Function FileCopy(SourceFile$, TargetFile$)

    Dim FSO As Variant
    Dim Src As Variant
    Dim TRG As Variant

    On Error Resume Next

        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set Src = FSO.GetFile(SourceFile)
        Src.Copy TargetFile
    

End Function


Private Sub mnuNewbook_Click()

Kill "c:\DWLNBfiles\*.nbt"

For xx = 1 To 100
pagetitle(xx) = ""
Next xx

Text1.Text = ""
rtb1.Text = ""
resetflags

For xx = 1 To 100
For yy = 1 To 10
imagename(xx, yy) = ""
Next yy
Next xx
For xx = 1 To 100
For yy = 1 To 10
imagecoordX(xx, yy) = 0
Next yy
Next xx
For xx = 1 To 100
For yy = 1 To 10
imagecoordY(xx, yy) = 0
Next yy
Next xx

For xx = 0 To 10
Picture2(xx).Visible = False
Next xx

For yy = 1 To 100
pageimages(yy) = 0
Next yy

marginflag = 0

mnuMargins.Checked = False
rtb1.Left = 780
rtb1.Width = 11295


currentpage = 1
numpages = 1
Label1.Caption = CStr(currentpage) + "/" + CStr(numpages)


End Sub


Private Sub mnuPageDown_Click()

If indexflag = 0 Then


Text1_Change
pagetitle(currentpage) = Text1.Text

On Error Resume Next
filename = "c:\DWLNBfiles\page" + CStr(currentpage) + ".nbt"
On Error Resume Next
rtb1.SaveFile (filename)

SaveSettings

If currentpage > 1 Then
currentpage = currentpage - 1
End If

If currentpage > numpages Then
numpages = currentpage
End If

rtb1.Text = ""
resetflags

On Error Resume Next
filename = "c:\DWLNBfiles\page" + CStr(currentpage) + ".nbt"
On Error Resume Next
rtb1.LoadFile (filename)
rtb1.SetFocus
rtb1.Refresh

rtb1.Refresh
Text1.Text = pagetitle(currentpage)
Label1.Caption = CStr(currentpage) + "/" + CStr(numpages)


For xx = 0 To 10
Picture2(xx).Visible = False
Next xx



For xx = 1 To pageimages(currentpage)

If imagename(currentpage, xx) <> "" Then
Picture2(xx).Visible = True
Picture2(xx).ZOrder (0)
Picture2(xx).Picture = LoadPicture(imagename(currentpage, xx))
Picture2(xx).Left = imagecoordX(currentpage, xx)
Picture2(xx).Top = imagecoordY(currentpage, xx)
End If

Next xx


End If

If indexflag = 1 Then
rtb1.Text = ""
  If TOCpage > 1 Then
  TOCpage = TOCpage - 1
  End If
    If TOCpage = 2 Then
    For yy = 1 To 25
    rtb1.Text = rtb1.Text + CStr(yy + 25) + ".    " + pagetitle(yy + 25) + vbCrLf
    Next yy
    End If
    
    If TOCpage = 3 Then
    For yy = 1 To 25
    rtb1.Text = rtb1.Text + CStr(yy + 50) + ".    " + pagetitle(yy + 50) + vbCrLf
    Next yy
    End If
    
    If TOCpage = 1 Then
    For yy = 1 To 25
    If yy < 10 Then
    rtb1.Text = rtb1.Text + CStr(yy) + ".    " + pagetitle(yy) + vbCrLf
    Else
    rtb1.Text = rtb1.Text + CStr(yy) + ".   " + pagetitle(yy) + vbCrLf
    End If
    Next yy
    End If
  
End If

rtb1.SetFocus
refresh_linecount

End Sub


Private Sub mnuPageUp_Click()

If indexflag = 0 Then


Text1_Change
pagetitle(currentpage) = Text1.Text

On Error Resume Next
filename = "c:\DWLNBfiles\page" + CStr(currentpage) + ".nbt"
On Error Resume Next
rtb1.SaveFile (filename)

SaveSettings

If currentpage < 100 Then
currentpage = currentpage + 1
End If

If currentpage > numpages Then
numpages = currentpage
End If

rtb1.Text = ""
resetflags

On Error Resume Next
filename = "c:\DWLNBfiles\page" + CStr(currentpage) + ".nbt"
On Error Resume Next
rtb1.LoadFile (filename)
rtb1.SetFocus
rtb1.Refresh

rtb1.Refresh
Text1.Text = pagetitle(currentpage)
Label1.Caption = CStr(currentpage) + "/" + CStr(numpages)


For xx = 0 To 10
Picture2(xx).Visible = False
Next xx



For xx = 1 To pageimages(currentpage)

If imagename(currentpage, xx) <> "" Then
Picture2(xx).Visible = True
Picture2(xx).ZOrder (0)
Picture2(xx).Picture = LoadPicture(imagename(currentpage, xx))
Picture2(xx).Left = imagecoordX(currentpage, xx)
Picture2(xx).Top = imagecoordY(currentpage, xx)
End If

Next xx


End If



If indexflag = 1 Then
  rtb1.Text = ""
  If TOCpage < 4 Then
  TOCpage = TOCpage + 1
  End If
    If TOCpage = 2 Then
    For yy = 1 To 25
    rtb1.Text = rtb1.Text + CStr(yy + 25) + ".    " + pagetitle(yy + 25) + vbCrLf
    Next yy
    End If
    
    If TOCpage = 3 Then
    For yy = 1 To 25
    rtb1.Text = rtb1.Text + CStr(yy + 50) + ".    " + pagetitle(yy + 50) + vbCrLf
    Next yy
    End If
    
    If TOCpage = 4 Then
    For yy = 1 To 25
    rtb1.Text = rtb1.Text + CStr(yy + 75) + ".    " + pagetitle(yy + 75) + vbCrLf
    Next yy
    End If
  
End If

rtb1.SetFocus
refresh_linecount

End Sub


Private Sub mnuTOC_Click()

TOCpage = 1

If indexflag = 0 Then
 indexflag = 1
 
 For xx = 0 To 10
Picture2(xx).Visible = False
Next xx
 
   
     pagetitle(currentpage) = Text1.Text
     
     
   On Error Resume Next
   filename = "c:\DWLNBfiles\page" + CStr(currentpage) + ".nbt"
   On Error Resume Next
   rtb1.SaveFile (filename)
     
     Text1.Text = "TABLE  OF  CONTENTS"
   rtb1.Text = ""
   rtb1.SelFontSize = 12
   rtb1.SelFontName = "Times new Roman"
   
   For yy = 1 To 25
    If yy < 10 Then
    rtb1.Text = rtb1.Text + CStr(yy) + ".    " + pagetitle(yy) + vbCrLf
    Else
    rtb1.Text = rtb1.Text + CStr(yy) + ".   " + pagetitle(yy) + vbCrLf
    End If
   Next yy
     Label1.Caption = "TOC"
Else
  indexflag = 0
  On Error Resume Next
  filename = "c:\DWLNBfiles\page" + CStr(currentpage) + ".nbt"
  On Error Resume Next
  rtb1.LoadFile (filename)
  rtb1.SetFocus
  rtb1.Refresh
  resetflags
    
For xx = 1 To pageimages(currentpage)
If imagename(currentpage, xx) <> "" Then
Picture2(xx).Visible = True
Picture2(xx).ZOrder (0)
Picture2(xx).Picture = LoadPicture(imagename(currentpage, xx))
Picture2(xx).Left = imagecoordX(currentpage, xx)
Picture2(xx).Top = imagecoordY(currentpage, xx)
End If
Next xx

  Text1.Text = pagetitle(currentpage)
  Label1.Caption = CStr(currentpage) + "/" + CStr(numpages)
 
End If

 rtb1.SetFocus
End Sub


Private Function Update()

indexflag = 0
  On Error Resume Next
  filename = "c:\DWLNBfiles\page" + CStr(currentpage) + ".nbt"
  On Error Resume Next
  rtb1.LoadFile (filename)
  rtb1.SetFocus
  rtb1.Refresh
  resetflags
    
For xx = 1 To pageimages(currentpage)
If imagename(currentpage, xx) <> "" Then
Picture2(xx).Visible = True
Picture2(xx).ZOrder (0)
Picture2(xx).Picture = LoadPicture(imagename(currentpage, xx))
Picture2(xx).Left = imagecoordX(currentpage, xx)
Picture2(xx).Top = imagecoordY(currentpage, xx)
End If
Next xx

  Text1.Text = pagetitle(currentpage)
  Label1.Caption = CStr(currentpage) + "/" + CStr(numpages)
 
End Function



Private Sub Picture2_MouseDown(index As Integer, Button As Integer, Shift As Integer, X2 As Single, Y2 As Single)

If Button = 2 Then 'delete image

Picture2(index).Visible = False
imagename(currentpage, index) = ""
imagecoordX(currentpage, index) = 0
imagecoordY(currentpage, index) = 0
delimgflag = 0

'compact array

If pageimages(currentpage) > 0 Then
For xx = index To pageimages(currentpage) - 1
imagename(currentpage, xx) = imagename(currentpage, (xx + 1))
imagecoordX(currentpage, xx) = imagecoordX(currentpage, (xx + 1))
imagecoordY(currentpage, xx) = imagecoordY(currentpage, (xx + 1))
Next xx
imagename(currentpage, pageimages(currentpage)) = ""
imagecoordX(currentpage, pageimages(currentpage)) = 0
imagecoordY(currentpage, pageimages(currentpage)) = 0
pageimages(currentpage) = pageimages(currentpage) - 1
End If

End If

End Sub


Private Sub Picture2_MouseMove(index As Integer, Button As Integer, Shift As Integer, X2 As Single, Y2 As Single)
On Error Resume Next
If Button = 1 Then
    Picture2(index).Left = Picture2(index).Left + X2
    Picture2(index).Top = Picture2(index).Top + Y2
    imagecoordX(currentpage, index) = Picture2(index).Left
    imagecoordY(currentpage, index) = Picture2(index).Top
End If
End Sub



Private Sub mnuPrint_Click() 'print page

 ' Get the printer's dimensions in twips.
    pwid = Printer.ScaleX(Printer.ScaleWidth, Printer.ScaleMode, vbTwips)
    phgt = Printer.ScaleY(Printer.ScaleHeight, Printer.ScaleMode, vbTwips)
    
    ' Convert the printer's dimensions into the
    ' object's coordinates.
    pwid = FormNB.ScaleX(pwid, vbTwips, FormNB.ScaleMode)
    phgt = FormNB.ScaleY(phgt, vbTwips, FormNB.ScaleMode)
    
    ' Compute the center of the object.
    xmid = FormNB.ScaleLeft + FormNB.ScaleWidth / 2
    ymid = FormNB.ScaleTop + FormNB.ScaleHeight / 2
    
    ' Pass the coordinates of the upper left and
    ' lower right corners into the Scale method.
    Printer.Scale _
        (xmid - pwid / 2, ymid - phgt / 2)- _
        (xmid + pwid / 2, ymid + phgt / 2)


   'Printer.PaintPicture Picture1.Image, 30, 100  'background
  
  ' If indexflag = 0 Then
  ' For xx = 1 To pageimages(currentpage)
  ' Printer.PaintPicture Picture2(xx).Image, imagecoordX(currentpage, xx), imagecoordY(currentpage, xx)
  ' Next xx
  ' End If
   
  ' Printer.CurrentX = 1545 + 3000
  ' Printer.CurrentY = 60
  ' Printer.FontName = Text1.FontName
  ' Printer.FontSize = Text1.FontSize
  ' Printer.Print Text1.Text
   
  ' Printer.CurrentX = 11100
  ' Printer.CurrentY = 120
   Printer.FontName = rtb1.SelFontName
   Printer.FontSize = rtb1.SelFontSize
  ' Printer.Print Label1.Caption
   Printer.FontSize = 12
   Printer.CurrentX = 0 '500
   Printer.CurrentY = 470 + (371)
   
   'Printer.EndDoc
   rtb1.SelStart = 0
   rtb1.SelPrint (Printer.hDC)
       
   On Error Resume Next
   Printer.EndDoc

 

End Sub


Private Sub mnuPagesetup_Click() 'printer setup

  On Error Resume Next
        '
        ' Show printer dialog
        With dlgfile
                .DialogTitle = "Page Setup"
                .CancelError = True
                .ShowPrinter
        End With

End Sub


Private Sub mnuCalcassoc_Click() 'associate with calc.exe

dlgfile.Filter = "*.*;"
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

If dlgfile.Name <> "" Then
calcassoc = dlgfile.filename
End If

End Sub


Private Sub mnuApp1assoc_Click()

dlgfile.Filter = "*.*;"
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

If dlgfile.Name <> "" Then
app1assoc = dlgfile.filename
End If

End Sub


Private Sub mnuApp2assoc_Click()

dlgfile.Filter = "*.*;"
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

If dlgfile.Name <> "" Then
app2assoc = dlgfile.filename
End If

End Sub


Private Sub mnuApp1_Click()

On Error Resume Next

Dim retval As Variant
retval = Shell(app1assoc, 1)

End Sub


Private Sub mnuApp2_Click()

On Error Resume Next

Dim retval As Variant
retval = Shell(app2assoc, 1)

End Sub


Private Sub mnuCalculator_Click() 'launch outboard calc

Dim retval As Variant
retval = Shell(calcassoc, 1)

End Sub

Private Sub mnuTextview_Click() 'view external text files

Textview.Show

End Sub


Private Sub mnuGoTo_Click()  'goto page

Text3.Visible = True
Text3.SetFocus

End Sub


Private Sub Text1_Change() 'check for quotes in title and convert to spaces

For xx = 1 To Len(Text1.Text)
If Mid(Text1.Text, xx, 1) = Chr(34) Then
Text1.Text = Left(Text1.Text, xx - 1) + Chr$(32) + Right$(Text1.Text, Len(Text1.Text) - xx)
End If
Next xx

End Sub

Private Sub Text3_KeyDown(keycode As Integer, Shift As Integer)

If indexflag = 0 Then

If keycode = vbKeyReturn Then
nextpage = Val(Text3.Text)

If nextpage = 0 Or nextpage > numpages Then
Exit Sub
End If


Text1_Change
pagetitle(currentpage) = Text1.Text

On Error Resume Next
filename = "c:\DWLNBfiles\page" + CStr(currentpage) + ".nbt"
On Error Resume Next
rtb1.SaveFile (filename)

SaveSettings

currentpage = nextpage
If numpages < currentpage Then
numpages = currentpage
End If

rtb1.Text = ""
resetflags

On Error Resume Next
filename = "c:\DWLNBfiles\page" + CStr(currentpage) + ".nbt"
On Error Resume Next
rtb1.LoadFile (filename)
rtb1.SetFocus
rtb1.Refresh

Text1.Text = pagetitle(currentpage)
Label1.Caption = CStr(currentpage) + "/" + CStr(numpages)

For xx = 0 To 10
Picture2(xx).Visible = False
Next xx

For xx = 1 To pageimages(currentpage)
If imagename(currentpage, xx) <> "" Then
Picture2(xx).Visible = True
Picture2(xx).Picture = LoadPicture(imagename(currentpage, xx))
Picture2(xx).Left = imagecoordX(currentpage, xx)
Picture2(xx).Top = imagecoordY(currentpage, xx)
End If
Next xx
Text3.Visible = False
rtb1.SetFocus
End If
End If

If indexflag = 1 Then
Text3.Visible = False
rtb1.SetFocus
End If

refresh_linecount

End Sub


Private Sub rtb1_change()

If undoflag = 0 Then
textundo(cc) = rtb1.Text
cc = cc + 1
If cc = 11 Then
cc = 1
End If
Else
undoflag = 0
End If

rtb1.SelFontSize = 12
rtb1.SelFontName = "Times New Roman"
If highredflag = 0 And highblueflag = 0 Then
rtb1.SelColor = vbBlack
End If

refresh_linecount

End Sub


Private Sub mnuUndo_Click()

undoflag = 1

If cc - 1 > 0 Then
textundo(cc - 1) = ""
Else
textundo(10) = ""
End If

If cc = 1 And textundo(9) <> "" Then
rtb1.Text = textundo(9)
textundo(9) = ""
End If
If cc = 2 And textundo(10) <> "" Then
rtb1.Text = textundo(10)
textundo(10) = ""
End If
If cc = 3 And textundo(1) <> "" Then
rtb1.Text = textundo(1)
textundo(1) = ""
End If
If cc = 4 And textundo(2) <> "" Then
rtb1.Text = textundo(2)
textundo(2) = ""
End If
If cc = 5 And textundo(3) <> "" Then
rtb1.Text = textundo(3)
textundo(3) = ""
End If
If cc = 6 And textundo(4) <> "" Then
rtb1.Text = textundo(4)
textundo(4) = ""
End If
If cc = 7 And textundo(5) <> "" Then
rtb1.Text = textundo(5)
textundo(5) = ""
End If
If cc = 8 And textundo(6) <> "" Then
rtb1.Text = textundo(6)
textundo(6) = ""
End If
If cc = 9 And textundo(7) <> "" Then
rtb1.Text = textundo(7)
textundo(7) = ""
End If
If cc = 10 And textundo(8) <> "" Then
rtb1.Text = textundo(8)
textundo(8) = ""
End If

cc = cc - 1
If cc = 0 Then
cc = 10
End If

End Sub

Private Sub rtb1_KeyDown(keycode As Integer, Shift As Integer)

If keycode = vbKeyF1 Then  'undo
mnuUndo_Click
End If

If keycode = vbKeyF2 Then  'equation
mnuEquation_Click
End If

If keycode = vbKeyF3 Then  'insert img
mniImage_Click
End If

If keycode = vbKeyF4 Then  'calc launch
mnuCalculator_Click
End If

If keycode = vbKeyF5 Then  'back a page
mnuPageDown_Click
End If

If keycode = vbKeyF6 Then  'goto page
mnuGoTo_Click
End If

If keycode = vbKeyF7 Then  'forward page
mnuPageUp_Click
End If

If keycode = vbKeyF8 Then  'TOC
mnuTOC_Click
End If

If keycode = vbKeyF9 Then  'Indent
mnuIndent_click
End If

If keycode = vbKeyF10 Then  'Bold
mnuBold_click
End If

If keycode = vbKeyF11 Then  'Underline
mnuUnderline_click
End If

If keycode = vbKeyF12 Then  'Italics
mnuItalics_click
End If


refresh_linecount
 
End Sub


Private Sub rtb1_KeyUp(keycode As Integer, Shift As Integer)

refresh_linecount

End Sub


Private Sub mnuEquation_Click()  'launch equation formatter form

eqtnform.Show

End Sub

Private Sub rtb1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 2 Then
PopupMenu FormNB.mnupopup
End If
rtb1.SelFontSize = 12


End Sub


Private Sub rtb1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

 refresh_linecount
 
End Sub


Private Sub mnuPaste_click() 'paste

On Error Resume Next

FormNB.rtb1.SelRTF = Clipboard.GetText
rtb1.SelStart = 0
rtb1.SelLength = Len(rtb1.Text)
rtb1.SelFontSize = 12
rtb1.SelFontName = "Times New Roman"
If highredflag = 0 And highblueflag = 0 Then
rtb1.SelColor = vbBlack
End If
rtb1.SelLength = 0

End Sub

Private Sub mnucopy_click() 'copy

On Error Resume Next

Clipboard.SetText FormNB.rtb1.SelRTF

End Sub

Private Sub mnuCut_click()  'cut

On Error Resume Next

Clipboard.SetText FormNB.rtb1.SelRTF
FormNB.rtb1.SelText = vbNullString

End Sub

Private Sub mnuDelete_click()  'delete

On Error Resume Next

FormNB.rtb1.SelText = vbNullString

End Sub

Private Sub mnuIndent_click()  'send 10 spaces

On Error Resume Next

For xx = 1 To 10
SendKeys " "
Next xx

End Sub


Private Sub mnuBold_click()  'Bold

If boldflag = 0 Then
boldflag = 1
Label9.Enabled = True
rtb1.SelBold = True
Else
boldflag = 0
Label9.Enabled = False
rtb1.SelBold = False
End If

End Sub


Private Sub mnuUnderline_click()  'Underline

If underlineflag = 0 Then
underlineflag = 1
Label10.Enabled = True
rtb1.SelUnderline = True
Else
underlineflag = 0
Label10.Enabled = False
rtb1.SelUnderline = False
End If

End Sub


Private Sub mnuItalics_click()  'Italics

If italicsflag = 0 Then
italicsflag = 1
Label11.Enabled = True
rtb1.SelItalic = True
Else
italicsflag = 0
Label11.Enabled = False
rtb1.SelItalic = False
End If

End Sub



Private Sub mnuHighred_click()  'Highlight red

If highredflag = 0 Then
highredflag = 1
Label12.Enabled = True
rtb1.SelColor = vbRed
highblueflag = 0
Label13.Enabled = False
Else
highredflag = 0
Label12.Enabled = False
rtb1.SelColor = vbBlack
End If

End Sub



Private Sub mnuHighblue_click()  'Highlight blue

If highblueflag = 0 Then
highblueflag = 1
Label13.Enabled = True
rtb1.SelColor = vbBlue
Label12.Enabled = False
highredflag = 0
Else
highblueflag = 0
Label13.Enabled = False
rtb1.SelColor = vbBlack
End If

End Sub

Private Sub mnuCenter_click()  'Center align

If centerflag = 0 Then
centerflag = 1
Label14.Enabled = True
rtb1.SelAlignment = 2
Else
centerflag = 0
Label14.Enabled = False
rtb1.SelAlignment = 0
End If

End Sub


Private Sub mnuLeft_click()  'Left align

centerflag = 0
Label14.Enabled = False
rtb1.SelAlignment = 0

End Sub


Private Sub mnuSelectall_click()  'Select all

rtb1.SelStart = 0
rtb1.SelLength = Len(rtb1.Text)

End Sub


Private Function SaveSettings()



pagetitle(currentpage) = Text1.Text


On Error Resume Next
filename = "c:\DWLNBfiles\page" + CStr(currentpage) + ".nbt"
On Error Resume Next
rtb1.SaveFile (filename)

  fnum1 = FreeFile
On Error Resume Next
Open "c:\DWLNBfiles\notebookini" For Output As #fnum1
On Error Resume Next
Write #fnum1, currentpage
Write #fnum1, numpages
Write #fnum1, fonttype
Write #fnum1, fontcolor
For xx = 1 To 100
For yy = 1 To 10
Write #fnum1, imagename(xx, yy)
Next yy
Next xx
For xx = 1 To 100
For yy = 1 To 10
Write #fnum1, imagecoordX(xx, yy)
Next yy
Next xx
For xx = 1 To 100
For yy = 1 To 10
Write #fnum1, imagecoordY(xx, yy)
Next yy
Next xx
For yy = 1 To 100
Write #fnum1, pageimages(yy)
Next yy
Write #fnum1, calcassoc
For xx = 1 To 100
Write #fnum1, pagetitle(xx)
Next xx
Write #fnum1, app1assoc
Write #fnum1, app2assoc
Write #fnum1, marginflag
Close fnum1

End Function


Private Function refresh_linecount()

 CurrentLine2 = SendMessage(rtb1.hwnd, EM_LINEFROMCHAR, -1, 0&) + 1
 Label2.Caption = CStr(CurrentLine2)
 Totallines = SendMessage(rtb1.hwnd, EM_GETLINECOUNT, 0, 0&)
 Label3.Caption = CStr(Totallines)
 charpos = rtb1.SelStart
 Label7.Caption = CStr(charpos)
 
If CurrentLine2 > 33 And imagename(currentpage, 1) <> "" Then
SendKeys "{PGUP}"
End If

End Function


Private Function resetflags() 'on page change

Label9.Enabled = False
Label10.Enabled = False
Label11.Enabled = False
Label12.Enabled = False
Label13.Enabled = False
Label14.Enabled = False
highredflag = 0
highblueflag = 0
boldflag = 0
italicsflag = 0
underlineflag = 0
centerflag = 0

End Function

Private Sub mnuMargins_Click()

If marginflag = 0 Then
marginflag = 1
mnuMargins.Checked = True
rtb1.Left = 1280
rtb1.Width = 10295
Else
marginflag = 0
mnuMargins.Checked = False
rtb1.Left = 780
rtb1.Width = 11295
End If


End Sub
