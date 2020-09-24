VERSION 5.00
Begin VB.Form EQhelp 
   AutoRedraw      =   -1  'True
   Caption         =   "Equation Help"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10770
   Icon            =   "EQhelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10770
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
      Height          =   7275
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   10710
   End
End
Attribute VB_Name = "EQhelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Text1.Text = ""

Text1.Text = Text1.Text + vbCrLf + "             EQUATION PAINTER 1.01 HELP:             build 10 August 2002"
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "    Enter formulas using numbers, +-/*, greek symbol names and keywords."
Text1.Text = Text1.Text + vbCrLf + "    The keyboard driven formula entry may seem harder at first but after memorizing"
Text1.Text = Text1.Text + vbCrLf + "    some of the keywords, will seem much more efficient than clicking assorted buttons."
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "    To optimize the look pad where needed with vertical and horizontal spaces in assorted"
Text1.Text = Text1.Text + vbCrLf + "    increments. The Symbol vs Courier font change in normal and superscript causes much of"
Text1.Text = Text1.Text + vbCrLf + "    the need. Enter the formula first then pad with spacing commands as required. Right click"
Text1.Text = Text1.Text + vbCrLf + "    in the textbox to cut, paste from the clipboard."
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "    Example formula:  beta = sqr alpha sup 2 + delta sup 2 esqr over  3"
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "    Start and end square roots, inversions (numerator over denominator); super and subscripts"
Text1.Text = Text1.Text + vbCrLf + "    end themselves as does over. Use over only in very simple cases, i.e. 1 over x."
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "    KEYWORDS:"
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "    @ - use at start and end of formula remarks (never image maps them)."
Text1.Text = Text1.Text + vbCrLf + "    times - puts an x (or you can use nothing, * or lowercase x)"
Text1.Text = Text1.Text + vbCrLf + "    times2 - puts a dot in for a multiplication sign."
Text1.Text = Text1.Text + vbCrLf + "    divide -  inserts a real divide sign as opposed to a slash."
Text1.Text = Text1.Text + vbCrLf + "    over -  sets the breakpoint for the numerator to denominator. Only one"
Text1.Text = Text1.Text + vbCrLf + "            divider bar per equation is allowed."
Text1.Text = Text1.Text + vbCrLf + "    sqr - starts a real square root sign (not a font character)."
Text1.Text = Text1.Text + vbCrLf + "    esqr - use at the end of the items you want the square root bar over"
Text1.Text = Text1.Text + vbCrLf + "    tsqr - starts a tall square root sign for 2 line square roots (use esqr after)."
Text1.Text = Text1.Text + vbCrLf + "    sqr2 - starts a 2nd real square root sign (not a font character)."
Text1.Text = Text1.Text + vbCrLf + "    esqr2 - use at the end of the items you want the 2nd square root bar over"
Text1.Text = Text1.Text + vbCrLf + "    tsqr2 - starts a 2nd tall square root sign for 2 line square roots (use esqr2 after)."
Text1.Text = Text1.Text + vbCrLf + "    deg - inserts a degree circle."
Text1.Text = Text1.Text + vbCrLf + "    nospc - no space, use to defeat the autospacing between characters."
Text1.Text = Text1.Text + vbCrLf + "    sub - subscript, applies to next whole word separated by spaces."
Text1.Text = Text1.Text + vbCrLf + "    sup - superscript, applies to next whole word."
Text1.Text = Text1.Text + vbCrLf + "    ssub - double subscript, applies to next whole word separated by spaces."
Text1.Text = Text1.Text + vbCrLf + "    ssup - double superscript, offsets twice as far vertically, applies to next whole word."
Text1.Text = Text1.Text + vbCrLf + "    +/- - inserts a + over a - as a single character."
Text1.Text = Text1.Text + vbCrLf + "    sinv - start local invert divider bar"
Text1.Text = Text1.Text + vbCrLf + "    inv -  set divider point for local invert divider bar"
Text1.Text = Text1.Text + vbCrLf + "    einv - end denominator entry of local invert bar"
Text1.Text = Text1.Text + vbCrLf + "    sinv2 - start 2nd local invert divider bar"
Text1.Text = Text1.Text + vbCrLf + "    inv2 -  set 2nd divider point for local invert divider bar"
Text1.Text = Text1.Text + vbCrLf + "    einv2 - end 2nd denominator entry of local invert bar"
Text1.Text = Text1.Text + vbCrLf + "    dnspc - downspace one row"
Text1.Text = Text1.Text + vbCrLf + "    upspc - upspace one row"
Text1.Text = Text1.Text + vbCrLf + "    dnhspc - downspace one half row"
Text1.Text = Text1.Text + vbCrLf + "    uphspc - upspace one half row"
Text1.Text = Text1.Text + vbCrLf + "    dnqspc - downspace a quarter row"
Text1.Text = Text1.Text + vbCrLf + "    upqspc - upspace a quarter row"
Text1.Text = Text1.Text + vbCrLf + "    dnespc - downspace an eighth row"
Text1.Text = Text1.Text + vbCrLf + "    upespc - upspace an eighth row"
Text1.Text = Text1.Text + vbCrLf + "    dnstspc - downspace a sixteenth row"
Text1.Text = Text1.Text + vbCrLf + "    upstspc - upspace a sixteenth row"
Text1.Text = Text1.Text + vbCrLf + "    dntsspc - downspace a thirtysecond row"
Text1.Text = Text1.Text + vbCrLf + "    uptsspc - upspace a thirtysecond row"
Text1.Text = Text1.Text + vbCrLf + "    crlf - downspace one row in comments and start a beginning of line"
Text1.Text = Text1.Text + vbCrLf + "    upaline - move up a line in comments"
Text1.Text = Text1.Text + vbCrLf + "    tab5 - go to the right 5 characters"
Text1.Text = Text1.Text + vbCrLf + "    tab10 - go to the right 10 characters"
Text1.Text = Text1.Text + vbCrLf + "    tab20 - go to the right 20 characters"
Text1.Text = Text1.Text + vbCrLf + "    back10 - go to the left 10 characters"
Text1.Text = Text1.Text + vbCrLf + "    back20 - go to the left 20 characters"
Text1.Text = Text1.Text + vbCrLf + "    cmmt - use at start and end of text that you don't want tokenized"
Text1.Text = Text1.Text + vbCrLf + "    inf -  draws a normal size infinity sign"
Text1.Text = Text1.Text + vbCrLf + "    inf2 -  draws a smaller sized infinity sign"
Text1.Text = Text1.Text + vbCrLf + "    subinf -  draws a subscript infinity sign"
Text1.Text = Text1.Text + vbCrLf + "    supinf -  draws a superscript infinity sign"
Text1.Text = Text1.Text + vbCrLf + "    ssubinf -  draws a double spaced down subscript infinity sign"
Text1.Text = Text1.Text + vbCrLf + "    ssupinf -  draws a double spaced up superscript infinity sign"
Text1.Text = Text1.Text + vbCrLf + "    bkspc -  backspace 1 character (use with subinf supinf)."
Text1.Text = Text1.Text + vbCrLf + "    hbkspc -  half backspace (use with 1 to supinf)."
Text1.Text = Text1.Text + vbCrLf + "    qspc -  quarter space 1 character (use with subinf supinf)."
Text1.Text = Text1.Text + vbCrLf + "    qbkspc -  quarter backspace (use with 1 to supinf)."
Text1.Text = Text1.Text + vbCrLf + "    ebkspc -  eighth backspace (use with 1 to supinf)."
Text1.Text = Text1.Text + vbCrLf + "    stbkspc -  sixteenth backspace (use with - j omega)."
Text1.Text = Text1.Text + vbCrLf + "    vbar -  do crlf first then make matrices end bars with this."
Text1.Text = Text1.Text + vbCrLf + "    sum - draws summation sign."
Text1.Text = Text1.Text + vbCrLf + "    int - draws integral sign."
Text1.Text = Text1.Text + vbCrLf + "    therefore - draws the 3 dot math symbol for therefore"
Text1.Text = Text1.Text + vbCrLf + "    noteql - not equal to"
Text1.Text = Text1.Text + vbCrLf + "    lessoreql - less than or equal to"
Text1.Text = Text1.Text + vbCrLf + "    greatoreql - greater than or equal to"
Text1.Text = Text1.Text + vbCrLf + "    semieql - wiggly over equal sign"
Text1.Text = Text1.Text + vbCrLf + "    vyeql - 3 bar high equal sign"
Text1.Text = Text1.Text + vbCrLf + "    approx - two wiggly lines as equal"
Text1.Text = Text1.Text + vbCrLf + "    tparl - left side of 2-line high parenthesis"
Text1.Text = Text1.Text + vbCrLf + "    tparr - right side of 2-line high parenthesis"
Text1.Text = Text1.Text + vbCrLf + "    tbktl - left side of 2-line high square bracket"
Text1.Text = Text1.Text + vbCrLf + "    tbktr - right side of 2-line high square bracket"
Text1.Text = Text1.Text + vbCrLf + "    lbkt - left side of 1-line high square bracket"
Text1.Text = Text1.Text + vbCrLf + "    rbkt - right side of 1-line high square bracket"
Text1.Text = Text1.Text + vbCrLf + "    lbkt2 - left side of 1-line high wiggly bracket"
Text1.Text = Text1.Text + vbCrLf + "    rbkt2 - right side of 1-line high wiggly bracket"
Text1.Text = Text1.Text + vbCrLf + "   "
Text1.Text = Text1.Text + vbCrLf + "   "
Text1.Text = Text1.Text + vbCrLf + "    (Use all uppercase for Greek name for uppercase Greek letters,"
Text1.Text = Text1.Text + vbCrLf + "     all lowercase name entry for lowercase Greek letters)."
Text1.Text = Text1.Text + vbCrLf + "    alpha - "
Text1.Text = Text1.Text + vbCrLf + "    beta - "
Text1.Text = Text1.Text + vbCrLf + "    gamma - "
Text1.Text = Text1.Text + vbCrLf + "    delta - "
Text1.Text = Text1.Text + vbCrLf + "    epsilon - "
Text1.Text = Text1.Text + vbCrLf + "    zeta - "
Text1.Text = Text1.Text + vbCrLf + "    eta - "
Text1.Text = Text1.Text + vbCrLf + "    theta - "
Text1.Text = Text1.Text + vbCrLf + "    iota - "
Text1.Text = Text1.Text + vbCrLf + "    kappa - "
Text1.Text = Text1.Text + vbCrLf + "    lambda - "
Text1.Text = Text1.Text + vbCrLf + "    mu - "
Text1.Text = Text1.Text + vbCrLf + "    nu - "
Text1.Text = Text1.Text + vbCrLf + "    xi - "
Text1.Text = Text1.Text + vbCrLf + "    omicron - "
Text1.Text = Text1.Text + vbCrLf + "    pi - "
Text1.Text = Text1.Text + vbCrLf + "    rho - "
Text1.Text = Text1.Text + vbCrLf + "    sigma - "
Text1.Text = Text1.Text + vbCrLf + "    tau - "
Text1.Text = Text1.Text + vbCrLf + "    upsilon - "
Text1.Text = Text1.Text + vbCrLf + "    phi - "
Text1.Text = Text1.Text + vbCrLf + "    chi - "
Text1.Text = Text1.Text + vbCrLf + "    psi - "
Text1.Text = Text1.Text + vbCrLf + "    omega - "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "   Remember that spaces are the separator item to use between keywords and variables."
Text1.Text = Text1.Text + vbCrLf + "   DO NOT use quotes in your text anywhere or the save formula function will work incorrectly."
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "   CROP IMAGE - starts a rectangular box that you click, hold and drag to autocrop"
Text1.Text = Text1.Text + vbCrLf + "   before saving to disk."
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "   CLEAR - erases everything to start over. "
Text1.Text = Text1.Text + vbCrLf + "   RESET PICTURE - erases the image view and sets picture to default size but retains the formula. "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "   LOAD FORMULA , SAVE FORMULA - this saves the Editor Window formula for future modifications. "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "   SAVE IMAGE - saves the preview window as it appears as a black and white .bmp"
Text1.Text = Text1.Text + vbCrLf + "   for use in any word processor or drawing program. Using B&W makes the file size nearly"
Text1.Text = Text1.Text + vbCrLf + "   as small as a .jpg. Always crop in the edit menu for absolute smallest file size."
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "   AUTO REDRAW OFF - If the automatic attempt to refresh the picture becomes annoying"
Text1.Text = Text1.Text + vbCrLf + "   it can be turned off in the options menu."
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + " "
Text1.Text = Text1.Text + vbCrLf + "   Visit: Dazyweblabs.com for the latest update on this program or email:"
Text1.Text = Text1.Text + vbCrLf + "                                                vrbalthezr@earthlink.net"










End Sub
