VERSION 5.00
Begin VB.Form Help 
   Caption         =   "Help"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10665
   LinkTopic       =   "Form2"
   ScaleHeight     =   7680
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   7500
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Help.frx":0000
      Top             =   120
      Width           =   10470
   End
End
Attribute VB_Name = "Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Text1.Text = "                      Lab Notebook Help Page" + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + "    DazyWeb Laboratories  EL-7000  Rev 1.06  build 17-June-02 " + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + "      Visit:    Dazyweblabs.com for updates or email:     vrbalthezr@earthlink.net" + vbCrLf
Text1.Text = Text1.Text + "              to comment or report bugs, suggest improvements." + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf + vbCrLf
Text1.Text = Text1.Text + " SHORTCUT KEYS:" + vbCrLf
Text1.Text = Text1.Text + " F1  - Undo text change (up to 9 deep)" + vbCrLf
Text1.Text = Text1.Text + " F2  - Equation Formatter" + vbCrLf
Text1.Text = Text1.Text + " F3  - Insert Image" + vbCrLf
Text1.Text = Text1.Text + " F4  - Launch Calculator" + vbCrLf
Text1.Text = Text1.Text + " F5  - Back a Page" + vbCrLf
Text1.Text = Text1.Text + " F6  - GoTo a Page" + vbCrLf
Text1.Text = Text1.Text + " F7  - Forward a Page" + vbCrLf
Text1.Text = Text1.Text + " F8  - Table Of Contents" + vbCrLf
Text1.Text = Text1.Text + " F9  - Indent 10 spaces" + vbCrLf
Text1.Text = Text1.Text + " F10 - Bold" + vbCrLf
Text1.Text = Text1.Text + " F11 - Underline" + vbCrLf
Text1.Text = Text1.Text + " F12 - Italics" + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + "     HOW TO RUN THE LAB NOTEBOOK:" + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " There are 100 pages possible in the notebook with up to 10 images per page." + vbCrLf
Text1.Text = Text1.Text + " 33 lines of text are allowed with images or unlimited no. of lines if no images are placed." + vbCrLf
Text1.Text = Text1.Text + " The images can be loaded in the Edit menu. BMP, JPG or GIF style images may be used" + vbCrLf
Text1.Text = Text1.Text + " and once loaded can be moved around by dragging and dropping with the mouse. The image " + vbCrLf
Text1.Text = Text1.Text + " location on the hard drive is logged in the notebook settings so either don't move them" + vbCrLf
Text1.Text = Text1.Text + " or pre-save a copy for the notebook in a file folder of your choice. Right-clicking on" + vbCrLf
Text1.Text = Text1.Text + " an image deletes it. Assorted maintenance features are in Edit and File such as New" + vbCrLf
Text1.Text = Text1.Text + " Notebook, Save Notebook, Load Notebook. The current one is saved at Exit if you use the" + vbCrLf
Text1.Text = Text1.Text + " Exit under the File pulldown menu. New Notebook will delete data from your current" + vbCrLf
Text1.Text = Text1.Text + " notebook so Save it first before hitting New Notebook. " + vbCrLf
Text1.Text = Text1.Text + " SAVING A NOTEBOOK - Because a separate file is created for each page (page1.nbt , page2.nbt ...)" + vbCrLf
Text1.Text = Text1.Text + " you will want to create a file folder before saving your notebook. Do this with Windows Explorer" + vbCrLf
Text1.Text = Text1.Text + " before starting to save your notebook and use it as a repository for your saved notebook. File" + vbCrLf
Text1.Text = Text1.Text + " copying of all the page files is automatic." + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " EDITTING - The main text section uses the standard Windows textbox control. Right clicking" + vbCrLf
Text1.Text = Text1.Text + " gives you the normal editting features and the Insert command supports cutting and pasting" + vbCrLf
Text1.Text = Text1.Text + " from the Windows clipboard. Richtext options: Bold, Italic, Underline and colored text." + vbCrLf
Text1.Text = Text1.Text + " NOTE: Formatting options are cleared when changing pages. To undo a formatted section of a newly" + vbCrLf
Text1.Text = Text1.Text + " loaded page, highlight the area, then rightclick on the function you want undone twice to change it back." + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " TEXTFILE PREVIEW under the File Menu will open the first 20000 characters of any text file" + vbCrLf
Text1.Text = Text1.Text + " and you can cut and paste from there to the main Notebook page." + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " SAVE TEXT ON PAGE - Save the richtext on the current page as a .rtf file (readable by most doc programs)." + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " PAGES - The top line is a larger font and is the Heading for the page as it will appear" + vbCrLf
Text1.Text = Text1.Text + " in the Table of Contents (TOC). The page up/down buttons at the top also navigate inside" + vbCrLf
Text1.Text = Text1.Text + " the TOC area. Pressing the TOC button once gets you into the TOC, Press again to get out." + vbCrLf
Text1.Text = Text1.Text + " The GOTO button drops down a textbox that you can put a page number into and the hit ENTER." + vbCrLf
Text1.Text = Text1.Text + " In the Edit menu, you can Clear a Page, Delete a Page (auto compresses the notebook" + vbCrLf
Text1.Text = Text1.Text + " around the deleted page), Insert a Page (expands the notebook) or Copy a Page (press Return in" + vbCrLf
Text1.Text = Text1.Text + " Copy To textbox after typing page number to activate function)." + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " PRINT FORM - There is a Print Page Setup item in the File menu. Print in Landscape mode" + vbCrLf
Text1.Text = Text1.Text + " and you will get a slow printout of a low res screenshot of your page in color (if your" + vbCrLf
Text1.Text = Text1.Text + " printer supports color)." + vbCrLf
Text1.Text = Text1.Text + " PRINT PAGE TEXT ONLY - This will print the body of the text saved on the page using as many "
Text1.Text = Text1.Text + " printer sheets as necessary and using quick text printing by the printer - good for big inserted documents."
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " CALCULATOR - takes you to the Windows default calculator. You can associate the Calculator" + vbCrLf
Text1.Text = Text1.Text + " button with a different calculator (such as the MA-2002) by going to Options and changing" + vbCrLf
Text1.Text = Text1.Text + " the file path to a different .EXE application." + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " USERAPP1 , USERAPP2 - Works like the calculator button above but assign executable program" + vbCrLf
Text1.Text = Text1.Text + " files of your choice (image editor, spreadsheet, schematic capture, etc. Look in Options to" + vbCrLf
Text1.Text = Text1.Text + " assign the association." + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " EQUATION FORMATTER - Launches a window where you can build a .BMP image file for use in the" + vbCrLf
Text1.Text = Text1.Text + " notebook. Supports Greek and Times New Roman font (mixed), superscripts, subscripts, divide" + vbCrLf
Text1.Text = Text1.Text + " bar, square root overbar and it autosizes the formula to fit in the smallest BMP size." + vbCrLf
Text1.Text = Text1.Text + " After choosing a 1 or 2 Line formula, just type in the two text boxes in the picture window" + vbCrLf
Text1.Text = Text1.Text + " and choose whether you want a divide by separator bar and any squareroot symbols. Afterwards" + vbCrLf
Text1.Text = Text1.Text + " save your work and then return to the notebook and insert those images." + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " NOTE: This App is optimized for 768 x 1024 screensize but may be used in 600 x 800 if you" + vbCrLf
Text1.Text = Text1.Text + " don't use all 26 lines. Part of the TOC will be cut off also but the Notebook would still" + vbCrLf
Text1.Text = Text1.Text + " be 100% functional." + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf
Text1.Text = Text1.Text + " " + vbCrLf













End Sub
