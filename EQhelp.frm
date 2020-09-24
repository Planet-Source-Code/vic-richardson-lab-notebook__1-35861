VERSION 5.00
Begin VB.Form EQhelp 
   AutoRedraw      =   -1  'True
   Caption         =   "Equation Help"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4650
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   10350
   End
End
Attribute VB_Name = "EQhelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Text1.Text = ""

Text1.Text = Text1.Text + vbCrLf + "     EQUATION HELP: "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "The upper box is a picture preview area of how the formula will look and also is an"
Text1.Text = Text1.Text + vbCrLf + "edit area for drawing lines (overbars) from the top left edge of square root symbols"
Text1.Text = Text1.Text + vbCrLf + "to where you would like the length to end. To see the latest view of the preview"
Text1.Text = Text1.Text + vbCrLf + "window use the REFRESH button. The imported image will look just like the preview box "
Text1.Text = Text1.Text + vbCrLf + "including it's exact width."
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "You can set the width of the image by left clicking in the preview window and dragging"
Text1.Text = Text1.Text + vbCrLf + "left or right and then letting go."
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "1 LINE , 2 LINE - choice of just a one line formula or use of a numerator and a "
Text1.Text = Text1.Text + vbCrLf + "denominator."
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "DIV - Toggles whether a dividing line needs to be between the numerator and denominator. "
Text1.Text = Text1.Text + vbCrLf + "The DIV length may be changed by holding down SHIFT and the left mouse button and "
Text1.Text = Text1.Text + vbCrLf + "dragging left or right, then letting go. "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "SQRT - special symbol ascii chr(214) in the Symbol char set, pressing this button "
Text1.Text = Text1.Text + vbCrLf + "places it in the text string. Use line drawing to complete the symbol. Do this by "
Text1.Text = Text1.Text + vbCrLf + "positioning the mouse cursor over the left upper endpoit of the SQRT sign, click once "
Text1.Text = Text1.Text + vbCrLf + "then position where you want it to end to the right and right click again. "
Text1.Text = Text1.Text + vbCrLf + "One overbar per line is allowed and the editor alternates between them to erase the "
Text1.Text = Text1.Text + vbCrLf + "old overbar and replace it with a new one. To eliminate overbars, left click then right "
Text1.Text = Text1.Text + vbCrLf + "click in the same place. "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "SUBSCRIPT - toggle button to make lowered variable identifiers. "
Text1.Text = Text1.Text + vbCrLf + "SUPERSCRIPT - toggle button to do squared or cubed or superscripts as needed. "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "GREEK - sets the font to Symbol, basically the Greek alphabet for hardcore variable naming. "
Text1.Text = Text1.Text + vbCrLf + "COURIER - safe monotype font that everyone has, good for the rest of the text and simple "
Text1.Text = Text1.Text + vbCrLf + "variable naming. Greek and Courier toggle the other one off. "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "NUM and DENOM textboxes are where you type in your rich text formulas. You can mix and match "
Text1.Text = Text1.Text + vbCrLf + "everything here. Hit refresh to see how it will look. "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "CLEAR erases everything to start over. "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "LOAD FORMULA , SAVE FORMULA - this saves the editor area for future modifications. "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "SAVE IMAGE - saves the preview window as it appears as a .bmp ready for use by the "
Text1.Text = Text1.Text + vbCrLf + "notebook. Use SAVE/INSERT IMAGE to save it and automatically insert it into the "
Text1.Text = Text1.Text + vbCrLf + "notebook. Remember you can right click on the image once it is in the notebook to "
Text1.Text = Text1.Text + vbCrLf + "quickly erase it. "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "Future revisions may include multiple overbars, (num1/den1) = (num2/den2), etc. "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + " "










End Sub
