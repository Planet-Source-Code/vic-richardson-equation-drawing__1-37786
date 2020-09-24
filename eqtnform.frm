VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form eqtnform 
   BackColor       =   &H8000000A&
   Caption         =   "Equation Drawing Program"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "eqtnform.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   11880
   Begin VB.CommandButton Command3 
      Caption         =   "Crop"
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
      Left            =   8160
      TabIndex        =   15
      Top             =   2760
      Width           =   1680
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset Picture"
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
      Left            =   2040
      TabIndex        =   14
      Top             =   2760
      Width           =   1680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
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
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   1680
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   360
      TabIndex        =   10
      Top             =   6840
      Width           =   11295
   End
   Begin VB.Frame Frame1 
      Caption         =   "KEYWORDS"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   8
      Top             =   5040
      Width           =   11415
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   11055
      End
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   7
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "eqtnform.frx":030A
      Top             =   3120
      Width           =   11415
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   2400
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   9960
      TabIndex        =   5
      Top             =   2760
      Width           =   1680
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2625
      Left            =   240
      ScaleHeight     =   2565
      ScaleWidth      =   11340
      TabIndex        =   2
      Top             =   0
      Width           =   11400
      Begin VB.Line Line4 
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   6000
         X2              =   7560
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   6000
         X2              =   7560
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   5880
         X2              =   7440
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   5880
         X2              =   7440
         Y1              =   360
         Y2              =   360
      End
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
      Left            =   1800
      ScaleHeight     =   285
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   4680
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
      Left            =   1320
      ScaleHeight     =   285
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   4680
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      TabIndex        =   13
      Top             =   7920
      Width           =   11295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Sample Formulas   -   Highlight desired formula and double click to move to Editor Window. "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   6600
      Width           =   10335
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
      Left            =   4560
      TabIndex        =   4
      Top             =   4680
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
      Left            =   4560
      TabIndex        =   3
      Top             =   2640
      Width           =   2550
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLoadformula 
         Caption         =   "Load Formula"
      End
      Begin VB.Menu mnuSaveformula 
         Caption         =   "Save Formula"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Image"
      End
      Begin VB.Menu qw 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu we 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCrop 
         Caption         =   "Crop"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuClip 
         Caption         =   "Copy to Clipboard"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "Reset Picture"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuAutoredraw 
         Caption         =   "Auto redraw on"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "eqtnform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'              DazyWeb Laboratories
'   Equation Painter  rev 1.01  build 10-August-02
'

Option Explicit

Dim justplusminflag As Integer
Dim justsymbolflag As Integer
Dim justsubsupflag As Integer
Dim nOldX
Dim nOldY
Dim mbDown
Dim cropflag As Integer
Dim cntr As Integer
Dim pixwidth As Long
Dim startpos As Long
Dim endpos As Long
Dim temppos As Long
Dim ycor As Long
Dim fnum1 As Integer
Dim X As Long
Dim Y As Long
Dim i As Long
Dim FileName As String
Dim txt As String
Dim txt1 As String
Dim txt2 As String
Dim txt3 As String
Dim pword(1000) As String
Dim pwordcnt As Integer
Dim mPath As String
Dim LastPath As String
Dim xx As Long
Dim endflag As Integer
Dim equalpos As Long
Dim caseflag As Integer
Dim overflag As Integer
Dim overpos As Long
Dim subflag As Integer
Dim supflag As Integer
Dim ssubflag As Integer
Dim ssupflag As Integer
Dim sqrflag As Integer
Dim sqr2flag As Integer
Dim maxX As Long
Dim tempX As Long
Dim tempX2 As Long
Dim tempY As Long
Dim temp2X As Long
Dim temp2X2 As Long
Dim temp2Y As Long
Dim insetflag As Integer
Dim autoredrawflag As Integer
Dim formula(100) As String
Dim invflag As Integer
Dim inv2flag As Integer
Dim invstpos As Long
Dim invstpos2 As Long
Dim invendpos As Long
Dim invendpos2 As Long
Dim invosflag As Integer
Dim invos2flag As Integer
Dim invos3flag As Integer
Dim overflag2 As Integer
Dim den_os_used As Integer
Dim uinvflag As Integer
Dim linvflag As Integer
Dim commentflag As Integer
Dim remflag As Integer


Private Sub Command1_Click() 'clear the formula and picture window

mnuClear_Click

End Sub


Private Sub Command2_Click() 'reset picture

mnuReset_Click

End Sub


Private Sub Command3_Click() 'crop picture

mnuCrop_Click

End Sub

Private Sub List1_dblClick() 'move sample to editor window

Text2.Text = formula(List1.ListIndex)
Command9_Click

End Sub


Private Sub mnuAutoredraw_Click() 'toggle autoredraw

If autoredrawflag = 0 Then
mnuAutoredraw.Checked = True
autoredrawflag = 1
Else
mnuAutoredraw.Checked = False
autoredrawflag = 0
End If

End Sub


Private Sub mnuClip_Click() 'copy to clipboard

        Clipboard.Clear
        DoEvents
        Clipboard.SetData Picture1.Picture
        DoEvents
        
End Sub


Private Sub mnuCrop_Click() 'crop pix before saving

cropflag = 1

End Sub


Private Sub mnuHelp_Click() 'help

EQhelp.Show

End Sub


Private Sub Form_Load()

autoredrawflag = 1
Text2.FontName = "Courier"
Text2.FontSize = 10
txt1 = ""
Picture3.Width = 11400
Picture3.Height = 2625
Picture2.Width = 11400
Picture2.Height = 2625
Picture4.Width = 11400
Picture4.Height = 2625
Picture1.Picture = Picture3.Image
pixwidth = 11400

Label1.Caption = "@   over  sinv   inv   einv   sinv2   inv2   einv2   sub   ssub   sup   ssup   sqr   esqr   tsqr   sqr2   tsqr2   esqr2   divide   times   times2   therefore   deg   +/-   inf   inf2   subinf   supinf   ssubinf   ssupinf   "
Label1.Caption = Label1.Caption + "bkspc   hbkspc   qbkspace   ebkspc   stbkspc   dnspc   upspc   uphspc   dnhspc   upqspc   dnqspc   upespc   dnespc   upstspc   dnstspc   uptsspc   dntsspc   qspc   nospc   crlf   upaline   tab5   tab10   tab20   back10   back20   cmmt   sum   int   vbar   noteql   lessoreql   greatoreql   semieql   vyeql   approx   tparl   tparr   tbktl   tbktr   lbkt   rbkt   lbkt2   rbkt2   alpha   beta   gamma   delta   epsilon   zeta   eta   theta  "
Label1.Caption = Label1.Caption + "iota   kappa   lambda   mu   nu   xi   omicron   pi   rho   "
Label1.Caption = Label1.Caption + "sigma   tau   upsilon   phi   chi   psi   omega   "

formula(0) = "@    Quadratic solver     @ Answer = -b +/-  sqr b sup 2 - 4ac esqr over      2a"
formula(1) = "@    Chained inversions   @ sinv x sup 2 inv a sup 2 einv + upespc  sinv y sup 2 inv b sup 2 einv - upespc  sinv z sup 2 inv c sup 2 einv = 0 cmmt   Elliptic cone  cmmt "
formula(2) = "@    Integral example     @ GAMMA (n) = int ssub 0  ssupinf nospc ebkspc x sup n-1 e sup -x dx  (n > 0)   GAMMA (n) = sinv  GAMMA (n+1) inv   n    einv (0 > n  noteql  -1,-2,-3,...) cmmt tab20 tab5 Gamma Function cmmt "
formula(3) = "@     Limit example       @ lim qbkspc ssub x->0 1/x = inf"
formula(4) = "@     Matrix example      @ crlf vbar a b vbar crlf vbar c d vbar upqspc upespc  = crlf tab5    vbar e f vbar crlf tab5     vbar g h vbar uphspc"
formula(5) = "@  Embedded square roots  @ dnspc Q =  sinv2 bkspc bkspc 1 inv2 dnhspc tsqr2 sinv C sub 2 inv C sub 1 einv  upqspc tparl nospc tsqr sinv R sub 3 inv R sub 2  einv esqr  dnhspc qbkspc +  uphspc tsqr sinv R sub 2 inv R sub 3  einv esqr dnhspc +  uphspc tsqr  sinv R sub 3 / R sub 2 inv   R sub 1    einv esqr  tparr uphspc esqr2 einv2 "
formula(6) = "@  Superscript example    @ y = sinv x inv 2 einv nospc (e sup x/a  +  e sup -x/a ) = a cosh sinv x inv a einv cmmt upaline  Caternary, Hyperbolic cosine cmmt"
formula(7) = "@  Differential example   @ sinv dz inv dx einv =  upstspc sinv dz inv du einv hbkspc  sinv du inv dx einv hbkspc  + sinv dz inv du einv hbkspc  sinv du inv dx einv  "
formula(8) = "@   Integral example      @ int u  delta v = uv - int v  delta u"
formula(9) = "@  Multiple inversions    @ (sinh sup -1 x)' =      sinv 1 inv y einv dnspc bkspc bkspc over   sinv    1 inv sqr 1 + x sup 2  esqr einv   "
formula(10) = "@  Superscript example    @ (x sup 2 + y sup 2 ) sup 2 = ax sup 2 y      r = a sin ebkspc  theta  cos sup 2 theta cmmt upaline upaline tab5 tab5 tab5  Bifolium cmmt "

For i = 0 To 10
List1.AddItem formula(i)
Next i

End Sub


Private Sub mnuClear_Click() 'clear

Text2.Text = ""
txt1 = ""
Picture1.Picture = Picture3.Image
Picture1.Width = 11400
Picture1.Height = 2625
pixwidth = 11400
cntr = 0

End Sub


Private Sub mnuExit_Click()  'exit
  
Unload eqtnform

End Sub


Private Sub Form_Unload(Cancel As Integer)

mnuExit_Click
    
End Sub


Private Sub mnuLoadformula_Click()

dlgFile.Filter = "All Files (*.eqn)|*.eqn|"
dlgFile.FileName = ""
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

FileName = dlgFile.FileName

If FileName <> "" Then

'parse for filepath
For xx = 0 To Len(FileName) - 1
 If Mid$(FileName, Len(FileName) - xx, 1) = "\" Then
 endflag = 1
 LastPath = txt3
 Else
   If endflag = 0 Then
   txt3 = Left$(FileName, Len(FileName) - xx - 1)
   End If
 End If
Next xx
endflag = 0


mnuClear_Click
fnum1 = FreeFile
On Error Resume Next
Open FileName For Input As #fnum1
On Error Resume Next
Input #fnum1, txt1
Close fnum1
Text2.Text = txt1
Label3.Caption = FileName
makepicture
Command9_Click
End If


End Sub


Private Sub mnuPrint_Click() 'print picture

Printer.PaintPicture Picture1.Image, 0, 0
Printer.EndDoc

End Sub


Private Sub mnuReset_Click() 'reset picturebox

Picture1.Picture = Picture3.Image
Picture1.Width = 11400
Picture1.Height = 2625
pixwidth = 11400

End Sub


Private Sub mnuSave_Click()  'save as image file .bmp

dlgFile.Filter = "All Files (*.bmp)|*.bmp|"
dlgFile.FileName = ""
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

FileName = dlgFile.FileName

If FileName <> "" Then

'parse for filepath
For xx = 0 To Len(FileName) - 1
 If Mid$(FileName, Len(FileName) - xx, 1) = "\" Then
 endflag = 1
 LastPath = txt3
 Else
   If endflag = 0 Then
   txt3 = Left$(FileName, Len(FileName) - xx - 1)
   End If
 End If
Next xx
endflag = 0
Label3.Caption = FileName

makepicture
On Error Resume Next
SaveNewBMP(Picture1, FileName, 1) = True
End If

Picture1.ScaleMode = 1

End Sub


Private Function makepicture()

Picture2.Picture = Picture3.Image
Picture2.Height = Picture1.Height
Picture2.Width = Picture1.Width
pixwidth = Picture1.Width
Picture2.Picture = Picture1.Image
Picture2.CurrentX = 190
Picture2.CurrentY = 390
Text2.SetFocus
Text2.SelStart = X
Picture2.FontName = "Courier New"
Picture2.FontSize = 12
cntr = 0
parseit
drawit

End Function


Private Sub mnuSaveformula_Click()


Command9_Click

dlgFile.Filter = "All Files (*.eqn)|*.eqn|"
dlgFile.FileName = ""
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

FileName = dlgFile.FileName
txt1 = Text2.Text


If FileName <> "" Then

'parse for filepath
For xx = 0 To Len(FileName) - 1
 If Mid$(FileName, Len(FileName) - xx, 1) = "\" Then
 endflag = 1
 LastPath = txt3
 Else
   If endflag = 0 Then
   txt3 = Left$(FileName, Len(FileName) - xx - 1)
   End If
 End If
Next xx
endflag = 0

On Error Resume Next
  fnum1 = FreeFile
  Open FileName For Output As #fnum1
  On Error Resume Next
  Write #fnum1, txt1
  Close fnum1
  Label3.Caption = FileName
End If

End Sub


Private Sub Command9_Click() 'refresh picture view

Picture1.Picture = Picture3.Image
Picture1.Refresh
makepicture
Picture1.Picture = Picture2.Image
Picture1.Picture = Picture1.Image

End Sub


Private Function parseit()

maxX = 0
subflag = 0
supflag = 0
sqrflag = 0
sqr2flag = 0
overflag = 0
overflag2 = 0
invflag = 0
inv2flag = 0
invosflag = 0
invos2flag = 0
invos3flag = 0
caseflag = 0
den_os_used = 0
commentflag = 0
uinvflag = 0  '
linvflag = 0  '
ssubflag = 0  '
ssupflag = 0  '
remflag = 0
pwordcnt = 1

For i = 1 To 1000  'clear old parse array (max 1000 chars)
pword(i) = ""
Next i

For i = 1 To Len(Text2.Text) 'parse text1.text
txt2 = Mid$(Text2.Text, i, 1)
If txt2 <> " " Then
pword(pwordcnt) = pword(pwordcnt) + txt2
End If

If txt2 = " " Then
pwordcnt = pwordcnt + 1
End If

Next i

End Function


Private Function detect_case()

caseflag = 0

If pword(i) = "@" Then
caseflag = 1
If remflag = 0 Then
remflag = 1
Else
remflag = 0
End If
End If

If remflag = 0 Then

If pword(i) = "cmmt" Then
caseflag = 1
If commentflag = 0 Then
commentflag = 1
tempX = Picture2.CurrentX
tempY = Picture2.CurrentY
Picture2.CurrentX = 150
Picture2.CurrentY = 1500
Else
commentflag = 0
Picture2.CurrentX = tempX
Picture2.CurrentY = tempY
End If
End If

If pword(i) = "crlf" Then  'go down a line
caseflag = 1
Picture2.CurrentY = Picture2.CurrentY + 280
Picture2.CurrentX = 150
End If

If pword(i) = "upaline" Then  'go down a line
caseflag = 1
Picture2.CurrentY = Picture2.CurrentY - 300
Picture2.CurrentX = 150
End If

If pword(i) = "tab5" Then  'go over 5 characters
caseflag = 1
Picture2.CurrentX = Picture2.CurrentX + 750
End If

If pword(i) = "tab10" Then  'go over 10 characters
caseflag = 1
Picture2.CurrentX = Picture2.CurrentX + 1500
End If

If pword(i) = "tab20" Then  'go over 20 characters
caseflag = 1
Picture2.CurrentX = Picture2.CurrentX + 3000
End If


If commentflag = 0 Then

If subflag = 2 Then 'past subscript value so reset offset, font
subflag = 0
Picture2.CurrentY = Picture2.CurrentY - 100
Picture2.FontSize = 12
End If

If supflag = 2 Then 'past superscript value so reset offset, font
supflag = 0
Picture2.CurrentY = Picture2.CurrentY + 50
Picture2.FontSize = 12
End If

If ssubflag = 2 Then 'past subsubscript value so reset offset, font
ssubflag = 0
Picture2.CurrentY = Picture2.CurrentY - 400
Picture2.FontSize = 12
End If

If ssupflag = 2 Then 'past supersuperscript value so reset offset, font
ssupflag = 0
Picture2.CurrentY = Picture2.CurrentY + 230
Picture2.FontSize = 12
End If


If subflag = 1 Then 'use offset for subscript value
subflag = 2
End If

If supflag = 1 Then 'use offset for superscript value
supflag = 2
End If

If ssubflag = 1 Then 'use offset for double subscript value
ssubflag = 2
End If

If ssupflag = 1 Then 'use offset for double superscript value
ssupflag = 2
End If





If pword(i) = "nospc" Then  'leave out a space between variables
caseflag = 1
Picture2.CurrentX = Picture2.CurrentX - 150
End If

If pword(i) = "dnspc" Then  'go down a line
caseflag = 1
Picture2.CurrentY = Picture2.CurrentY + 450
End If


If pword(i) = "upspc" Then  'go up a line
caseflag = 1
Picture2.CurrentY = Picture2.CurrentY - 450
End If


If pword(i) = "uphspc" Then  'go up a half line
caseflag = 1
Picture2.CurrentY = Picture2.CurrentY - 225
End If


If pword(i) = "dnhspc" Then  'go down a half line
caseflag = 1
Picture2.CurrentY = Picture2.CurrentY + 225
End If


If pword(i) = "upqspc" Then  'go up a quarter line
caseflag = 1
Picture2.CurrentY = Picture2.CurrentY - 112
End If


If pword(i) = "dnqspc" Then  'go down a quarter line
caseflag = 1
Picture2.CurrentY = Picture2.CurrentY + 112
End If


If pword(i) = "upespc" Then  'go up an eighth line
caseflag = 1
Picture2.CurrentY = Picture2.CurrentY - 56
End If


If pword(i) = "dnespc" Then  'go down an eighth line
caseflag = 1
Picture2.CurrentY = Picture2.CurrentY + 56
End If

If pword(i) = "upstspc" Then  'go up a sixteenth line
caseflag = 1
Picture2.CurrentY = Picture2.CurrentY - 28
End If


If pword(i) = "dnstspc" Then  'go down a sixteenth line
caseflag = 1
Picture2.CurrentY = Picture2.CurrentY + 28
End If

If pword(i) = "uptsspc" Then  'go up a thirtysecond line
caseflag = 1
Picture2.CurrentY = Picture2.CurrentY - 14
End If


If pword(i) = "dntsspc" Then  'go down a thirtysecond line
caseflag = 1
Picture2.CurrentY = Picture2.CurrentY + 14
End If


If pword(i) = "inv" And overflag2 = 1 Then 'set offset flag for denominator if numerator has inversion
invos2flag = 1
End If

If pword(i) = "inv2" And overflag2 = 1 Then 'set offset flag for denominator if numerator has inversion
invos3flag = 1
End If

If pword(i) = "over" And invosflag = 0 Then 'flag that this formula has a divider bar
caseflag = 1
overflag = 1
overflag2 = 1
overpos = Picture2.CurrentX
Picture2.CurrentX = equalpos + 300
Picture2.CurrentY = 710
End If

If pword(i) = "over" And invosflag = 1 Then 'flag that this formula has a divider bar
caseflag = 1
overflag = 1
overflag2 = 1
overpos = Picture2.CurrentX
Picture2.CurrentX = equalpos + 300
Picture2.CurrentY = 710 + 200
End If
  
If pword(i) = "sinv" And overflag2 = 1 Then
invos2flag = 1
caseflag = 1
End If
  
If pword(i) = "sinv2" And overflag2 = 1 Then
invos3flag = 1
caseflag = 1
End If
  
If pword(i) = "sinv" Then 'flag that this formula has a local divider bar
 caseflag = 1
 invflag = 1
 invstpos = Picture2.CurrentX + 150
 Picture2.CurrentX = invstpos - 150
   If invos2flag = 0 Then
   Picture2.CurrentY = Picture2.CurrentY - 150
   Else
      Picture2.CurrentY = Picture2.CurrentY + 150
        If den_os_used = 1 Then
        Picture2.CurrentY = Picture2.CurrentY - 300
        End If
   End If
End If

If pword(i) = "sinv2" Then 'flag that this formula has a 2nd local divider bar
 caseflag = 1
 inv2flag = 1
 invstpos2 = Picture2.CurrentX + 150
 Picture2.CurrentX = invstpos - 150
   If invosflag = 0 Then
   Picture2.CurrentY = Picture2.CurrentY - 150
   Else
      Picture2.CurrentY = Picture2.CurrentY + 150
        If den_os_used = 1 Then
        Picture2.CurrentY = Picture2.CurrentY - 300
        End If
   End If
End If



If pword(i) = "inv" Then 'set local divider bar
caseflag = 1
invflag = 1
Picture2.CurrentX = invstpos - 150
  If justsubsupflag = 0 Then
  Picture2.CurrentY = Picture2.CurrentY + 300
  Else
  Picture2.CurrentY = Picture2.CurrentY + 356
  End If
End If


If pword(i) = "inv2" Then 'set local divider bar2
caseflag = 1
inv2flag = 1
Picture2.CurrentX = invstpos2 - 150
 If justsubsupflag = 0 Then
  Picture2.CurrentY = Picture2.CurrentY + 300
  Else
  Picture2.CurrentY = Picture2.CurrentY + 356
  End If
End If


If pword(i) = "einv" Then 'flag that this formula has ended local divider bar
caseflag = 1
invflag = 0
invendpos = Picture2.CurrentX
Picture2.Line (invstpos, Picture2.CurrentY - 30)-(invendpos, Picture2.CurrentY - 30) 'draw divider bar
Picture2.CurrentX = Picture2.CurrentX + 150
Picture2.CurrentY = Picture2.CurrentY - 120
   If overflag2 = 1 Then
   den_os_used = 1
   End If
End If


If pword(i) = "einv2" Then 'flag that this formula has ended local divider bar2
caseflag = 1
inv2flag = 0
invendpos2 = Picture2.CurrentX
Picture2.Line (invstpos2, Picture2.CurrentY - 30)-(invendpos2, Picture2.CurrentY - 30) 'draw divider bar2
Picture2.CurrentX = Picture2.CurrentX + 150
Picture2.CurrentY = Picture2.CurrentY - 120
   If overflag2 = 1 Then
   den_os_used = 1
   End If
End If


If pword(i) = "=" Then 'save x position of equal sign to set divider bar start position
equalpos = Picture2.CurrentX
End If


If pword(i) = "sqr" Then 'begin square root symbol
caseflag = 1
sqrflag = 1
tempX = Picture2.CurrentX
tempY = Picture2.CurrentY
Picture2.Line (tempX + 300, tempY)-(tempX + 450, tempY) 'draw square root symbol
Picture2.Line (tempX + 300, tempY)-(tempX + 200, tempY + 200)
Picture2.Line (tempX + 200, tempY + 200)-(tempX + 150, tempY + 100)
Picture2.CurrentX = tempX + 150
Picture2.CurrentY = tempY
End If


If pword(i) = "sqr2" Then 'begin square root symbol
caseflag = 1
sqr2flag = 1
temp2X = Picture2.CurrentX
temp2Y = Picture2.CurrentY
Picture2.Line (temp2X + 300, temp2Y)-(temp2X + 450, temp2Y) 'draw square root symbol
Picture2.Line (temp2X + 300, temp2Y)-(temp2X + 200, temp2Y + 200)
Picture2.Line (temp2X + 200, temp2Y + 200)-(temp2X + 150, temp2Y + 100)
Picture2.CurrentX = temp2X + 150
Picture2.CurrentY = temp2Y
End If


If pword(i) = "esqr" Then 'finish square root overbar
caseflag = 1
sqrflag = 0
tempX2 = Picture2.CurrentX
Picture2.Line (tempX + 300, tempY)-(tempX2, tempY)
Picture2.CurrentX = tempX2
End If


If pword(i) = "esqr2" Then 'finish square root overbar
caseflag = 1
sqr2flag = 0
temp2X2 = Picture2.CurrentX
Picture2.Line (temp2X + 300, temp2Y)-(temp2X2, temp2Y)
Picture2.CurrentX = temp2X2
End If


If pword(i) = "tsqr" Then 'begin tall square root symbol
caseflag = 1
sqrflag = 1
tempX = Picture2.CurrentX
tempY = Picture2.CurrentY
Picture2.Line (tempX + 300, tempY)-(tempX + 450, tempY) 'draw square root symbol
Picture2.Line (tempX + 300, tempY)-(tempX + 200, tempY + 600)
Picture2.Line (tempX + 200, tempY + 600)-(tempX + 150, tempY + 400)
Picture2.CurrentX = tempX + 150 + 150
Picture2.CurrentY = tempY + 168
End If

If pword(i) = "tsqr2" Then 'begin tall square root symbol
caseflag = 1
sqr2flag = 1
temp2X = Picture2.CurrentX
temp2Y = Picture2.CurrentY
Picture2.Line (temp2X + 300, temp2Y)-(temp2X + 450, temp2Y) 'draw square root symbol
Picture2.Line (temp2X + 300, temp2Y)-(temp2X + 200, temp2Y + 600)
Picture2.Line (temp2X + 200, temp2Y + 600)-(temp2X + 150, temp2Y + 400)
Picture2.CurrentX = temp2X + 150 + 150
Picture2.CurrentY = temp2Y + 168
End If


If pword(i) = "tparl" Then 'big left paranth
caseflag = 1
tempX = Picture2.CurrentX
tempY = Picture2.CurrentY
Picture2.Circle (tempX + 1200, tempY + 300), 1200, , 3.14 * 0.92, 1.08 * 3.14
Picture2.CurrentX = tempX + 150
Picture2.CurrentY = tempY
End If


If pword(i) = "tparr" Then 'big right paranth
caseflag = 1
tempX = Picture2.CurrentX
tempY = Picture2.CurrentY
Picture2.Circle (tempX - 1200, tempY + 300), 1200, , 3.14 * 1.92, 0.08 * 3.14
Picture2.CurrentX = tempX + 150
Picture2.CurrentY = tempY
End If

If pword(i) = "tbktl" Then 'big left bracket
caseflag = 1
tempX = Picture2.CurrentX
tempY = Picture2.CurrentY
Picture2.Line (tempX, tempY)-(tempX, tempY + 600)
Picture2.Line (tempX, tempY)-(tempX + 80, tempY)
Picture2.Line (tempX, tempY + 600)-(tempX + 80, tempY + 600)
Picture2.CurrentX = tempX + 150
Picture2.CurrentY = tempY
End If


If pword(i) = "tbktr" Then 'big right bracket
caseflag = 1
tempX = Picture2.CurrentX
tempY = Picture2.CurrentY
Picture2.Line (tempX, tempY)-(tempX, tempY + 600)
Picture2.Line (tempX, tempY)-(tempX - 80, tempY)
Picture2.Line (tempX, tempY + 600)-(tempX - 80, tempY + 600)
Picture2.CurrentX = tempX + 150
Picture2.CurrentY = tempY
End If


If pword(i) = "times2" Then 'inline dot for multiplication sign
caseflag = 1
tempX = Picture2.CurrentX
tempY = Picture2.CurrentY
Picture2.Circle (tempX + 100, tempY + 150), 8
Picture2.CurrentX = tempX + 50
Picture2.CurrentY = tempY
End If


If pword(i) = "int" Then 'inline integral sign
caseflag = 1
tempX = Picture2.CurrentX
tempY = Picture2.CurrentY
Picture2.Circle Step(300, 0), 50, , 0, 1.1 * 3.14
Picture2.Circle Step(-100, 300), 50, , 3.14, 1.99 * 3.14
Picture2.Line (tempX + 250, tempY)-(tempX + 250, tempY + 310)
Picture2.CurrentX = tempX + 300
Picture2.CurrentY = tempY
End If


If pword(i) = "sum" Then 'summation symbol
caseflag = 1
tempX = Picture2.CurrentX
tempY = Picture2.CurrentY
Picture2.Line (tempX + 300, tempY + 50)-(tempX + 500, tempY + 50)
Picture2.Line (tempX + 300, tempY + 250)-(tempX + 500, tempY + 250)
Picture2.Line (tempX + 500, tempY + 50)-(tempX + 508, tempY + 85)
Picture2.Line (tempX + 500, tempY + 250)-(tempX + 508, tempY + 215)
Picture2.Line (tempX + 300, tempY + 50)-(tempX + 420, tempY + 158)
Picture2.Line (tempX + 300, tempY + 250)-(tempX + 420, tempY + 142)
Picture2.CurrentX = tempX + 150
Picture2.CurrentY = tempY
End If


If pword(i) = "inf" Then 'inline infinity sign
caseflag = 1
tempX = Picture2.CurrentX
tempY = Picture2.CurrentY
Picture2.Circle Step(225, 150), 5
Picture2.Circle Step(-25, 0), 50
Picture2.Circle Step(115, 0), 50
Picture2.CurrentX = tempX + 450
Picture2.CurrentY = tempY
End If


If pword(i) = "supinf" Then 'superscript infinity sign
caseflag = 1
tempX = Picture2.CurrentX
tempY = Picture2.CurrentY
Picture2.Circle Step(225, 50), 5
Picture2.Circle Step(-25, 0), 27
Picture2.Circle Step(60, 0), 27
Picture2.CurrentX = tempX + 450
Picture2.CurrentY = tempY
End If

If pword(i) = "subinf" Then 'subscript infinity sign
caseflag = 1
tempX = Picture2.CurrentX
tempY = Picture2.CurrentY
Picture2.Circle Step(225, 250), 5
Picture2.Circle Step(-25, 0), 27
Picture2.Circle Step(60, 0), 27
Picture2.CurrentX = tempX + 450
Picture2.CurrentY = tempY
End If


If pword(i) = "ssupinf" Then 'double superscript infinity sign
caseflag = 1
tempX = Picture2.CurrentX
tempY = Picture2.CurrentY
Picture2.Circle Step(0, -150), 5
Picture2.Circle Step(-25, 0), 27
Picture2.Circle Step(60, 0), 27
Picture2.CurrentX = tempX + 250
Picture2.CurrentY = tempY
End If


If pword(i) = "ssubinf" Then 'double subscript infinity sign
caseflag = 1
tempX = Picture2.CurrentX - 300
tempY = Picture2.CurrentY
Picture2.Circle Step(0, 500), 5
Picture2.Circle Step(-25, 0), 27
Picture2.Circle Step(60, 0), 27
Picture2.CurrentX = tempX + 250
Picture2.CurrentY = tempY
End If


If pword(i) = "back10" Then 'offset one character backwards
caseflag = 1
Picture2.CurrentX = Picture2.CurrentX - 2250
End If


If pword(i) = "back20" Then 'offset one character backwards
caseflag = 1
Picture2.CurrentX = Picture2.CurrentX - 4500
End If


If pword(i) = "bkspc" Then 'offset one character backwards
caseflag = 1
Picture2.CurrentX = Picture2.CurrentX - 450
End If


If pword(i) = "hbkspc" Then 'offset one half character backwards
caseflag = 1
Picture2.CurrentX = Picture2.CurrentX - 300
End If


If pword(i) = "qbkspc" Then 'offset one quarter character backwards
caseflag = 1
Picture2.CurrentX = Picture2.CurrentX - 150
End If

If pword(i) = "ebkspc" Then 'offset one eighth character backwards
caseflag = 1
Picture2.CurrentX = Picture2.CurrentX - 75
End If

If pword(i) = "stbkspc" Then 'offset one eighth character backwards
caseflag = 1
Picture2.CurrentX = Picture2.CurrentX - 37
End If


If pword(i) = "qspc" Then 'offset one half character backwards
caseflag = 1
Picture2.CurrentX = Picture2.CurrentX + 75
End If


If pword(i) = "sub" Then 'offset for subscript
caseflag = 1
subflag = 1
Picture2.CurrentY = Picture2.CurrentY + 100
Picture2.CurrentX = Picture2.CurrentX - 70
Picture2.FontSize = 8
End If


If pword(i) = "ssub" Then 'double offset for subscript
caseflag = 1
ssubflag = 1
Picture2.CurrentY = Picture2.CurrentY + 400
Picture2.CurrentX = Picture2.CurrentX - 300
Picture2.FontSize = 8
End If


If pword(i) = "sup" Then 'offset for superscript
caseflag = 1
supflag = 1
Picture2.CurrentY = Picture2.CurrentY - 50
Picture2.CurrentX = Picture2.CurrentX - 70
Picture2.FontSize = 8
End If


If pword(i) = "ssup" Then 'double offset for superscript
caseflag = 1
ssupflag = 1
Picture2.CurrentY = Picture2.CurrentY - 230
Picture2.CurrentX = Picture2.CurrentX - 250
Picture2.FontSize = 8
End If


If pword(i) = "+/-" Then
insetflag = 1 'set flag to trim trailing space to value
Picture2.Font = "Symbol"
pword(i) = "±"
End If

check_for_greek

End If  'end test for comments flag (commentflag)

End If  ' test for remarks flag (remflag)


End Function


Private Function check_for_greek()

Picture2.Font = "Symbol"

Select Case pword(i)
   Case "deg"
      pword(i) = "°"
      Picture2.CurrentX = Picture2.CurrentX - 50
   Case "times"
      pword(i) = " ´"
   Case "divide"
      pword(i) = " ¸"
   Case "therefore"
      pword(i) = "\"
   Case "lbkt"
      pword(i) = Chr(91)
   Case "rbkt"
      pword(i) = Chr(93)
   Case "lbkt2"
      pword(i) = Chr(123)
   Case "rbkt2"
      pword(i) = Chr(125)
   Case "alpha"
      pword(i) = "a"
   Case "beta"
      pword(i) = "b"
   Case "gamma"
      pword(i) = "g"
   Case "delta"
      pword(i) = "d"
   Case "epsilon"
      pword(i) = "e"
   Case "zeta"
      pword(i) = "z"
   Case "eta"
      pword(i) = "h"
   Case "theta"
      pword(i) = "q"
      
   Case "noteql"
      pword(i) = Chr(185)
   Case "lessoreql"
      pword(i) = Chr(163)
   Case "greatoreql"
      pword(i) = Chr(179)
   Case "semieql"
      pword(i) = Chr(64)
   Case "vyeql"
      pword(i) = Chr(186)
   Case "inf2"
      pword(i) = Chr(165)
   Case "approx"
      pword(i) = Chr(187)
   Case "vbar"
      pword(i) = Chr(189)
      
   Case "iota"
      pword(i) = "i"
   Case "kappa"
      pword(i) = "k"
   Case "lambda"
      pword(i) = "l"
   Case "mu"
      pword(i) = "m"
   Case "nu"
      pword(i) = "n"
   Case "xi"
      pword(i) = "x"
   Case "omicron"
      pword(i) = "o"
   Case "pi"
      pword(i) = "p"
   Case "rho"
      pword(i) = "r"
   Case "sigma"
      pword(i) = "s"
   Case "tau"
      pword(i) = "t"
   Case "upsilon"
      pword(i) = "u"
   Case "phi"
      pword(i) = "f"
   Case "chi"
      pword(i) = "c"
   Case "psi"
      pword(i) = "y"
   Case "omega"
      pword(i) = "w"
   Case "ALPHA"
      pword(i) = "A"
   Case "BETA"
      pword(i) = "B"
   Case "GAMMA"
      pword(i) = "G"
   Case "DELTA"
      pword(i) = "D"
   Case "EPSILON"
      pword(i) = "E"
   Case "ZETA"
      pword(i) = "Z"
   Case "ETA"
      pword(i) = "H"
   Case "THETA"
      pword(i) = "Q"
   Case "IOTA"
      pword(i) = "I"
   Case "KAPPA"
      pword(i) = "K"
   Case "LAMBDA"
      pword(i) = "L"
   Case "MU"
      pword(i) = "M"
   Case "NU"
      pword(i) = "N"
   Case "XI"
      pword(i) = "X"
   Case "OMICRON"
      pword(i) = "O"
   Case "PI"
      pword(i) = "P"
   Case "RHO"
      pword(i) = "R"
   Case "SIGMA"
      pword(i) = "S"
   Case "TAU"
      pword(i) = "T"
   Case "UPSILON"
      pword(i) = "U"
   Case "PHI"
      pword(i) = "F"
   Case "CHI"
      pword(i) = "C"
   Case "PSI"
      pword(i) = "Y"
   Case "OMEGA"
      pword(i) = "W"
   Case Else
      Picture2.Font = "Courier New"
End Select

End Function



Private Function drawit()

uinvflag = 0
linvflag = 0

For i = 1 To pwordcnt 'precalc numerator overs, denominator overs, major over
    
    If pword(i) = "cmmt" Then
      If commentflag = 0 Then
      commentflag = 1
      Else
      commentflag = 0
      End If
   End If

   If pword(i) = "@" Then
      If remflag = 0 Then
      remflag = 1
      Else
      remflag = 0
      End If
   End If
   

   If commentflag = 0 And remflag = 0 Then
   
   If pword(i) = "over" Then 'detect whether to start left of equation lower if over used
   overflag = 1
   End If
   
   If overflag = 1 And (pword(i) = "inv" Or pword(i) = "inv2") Then 'flag that invert is used in denominator
   linvflag = 1
   End If
      
   If overflag = 0 And (pword(i) = "inv" Or pword(i) = "inv2") Then 'flag that invert is used in numerator
   invosflag = 1
   uinvflag = 1
   End If
   
   End If
   
   
Next i
            'set left side of equation Y offset
            
   commentflag = 0
   remflag = 0
   
   If overflag = 1 And linvflag = 0 And uinvflag = 0 Then
   Picture2.CurrentY = 550
   End If
   
   If overflag = 1 And linvflag = 1 And uinvflag = 0 Then
   Picture2.CurrentY = 800
   End If
   
   If overflag = 1 And linvflag = 0 And uinvflag = 1 Then
   Picture2.CurrentY = 500
   End If
   
   If overflag = 1 And linvflag = 1 And uinvflag = 1 Then
   Picture2.CurrentY = 780
   End If

   If overflag = 0 And linvflag = 0 And uinvflag = 1 Then
   Picture2.CurrentY = 350
   End If
   
   
   
For i = 1 To pwordcnt  'draw the formula now

detect_case
     
  If caseflag = 0 And remflag = 0 Then 'not a format word
    
      If Picture2.Font = "Symbol" And justplusminflag = 1 Then
      Picture2.CurrentX = Picture2.CurrentX + 150
      End If
      
     If Picture2.Font = "Symbol" And justplusminflag = 0 Then
     Picture2.CurrentX = Picture2.CurrentX - 75 'do eighth backspace with Greek font
     justsymbolflag = 1
     End If
     
     If Picture2.Font = "Symbol" And justsubsupflag = 1 Then 'counter for subscripts after Greek
     Picture2.CurrentX = Picture2.CurrentX + 150
     justsymbolflag = 1
     End If
     
     If justsymbolflag = 1 And justsubsupflag = 0 And Picture2.FontName = "Courier New" Then 'couter for Courier after Greek
     Picture2.CurrentX = Picture2.CurrentX - 150
     justsymbolflag = 0
     End If
     
     If pword(i) = "+" Or pword(i) = "-" And justsubsupflag = 1 Then
     Picture2.CurrentX = Picture2.CurrentX + 75
     End If
       
     If pword(i) = "+" Or pword(i) = "-" And justsymbolflag = 0 And justsubsupflag = 0 And Picture2.FontName = "Courier New" Then
     Picture2.CurrentX = Picture2.CurrentX - 75
     End If
       
     If subflag = 0 And supflag = 0 Then
     Picture2.Print " " + pword(i);
     justsubsupflag = 0
     Else
     Picture2.Print " " + pword(i);
     Picture2.CurrentX = Picture2.CurrentX - 125
     justsubsupflag = 1
     End If
      
     If pword(i) = "+" Or pword(i) = "-" Then
     justplusminflag = 1
     Picture2.CurrentX = Picture2.CurrentX - 75
     Else
     justplusminflag = 0
     End If
     
    Picture2.FontName = "Courier New"
     If Picture2.CurrentX > maxX Then
     maxX = Picture2.CurrentX
     End If
  End If
  
  If insetflag = 1 Then
  insetflag = 0
  Picture2.CurrentX = Picture2.CurrentX - 150
  End If
  
  
  If pword(i) = "=" And commentflag = 0 Then  'force numerator higher in position
  Picture2.CurrentY = 370
  End If
Next i

'logic on the X,Y metrics to draw the OVER bar

If (overflag = 1) And (invosflag = 0) And (invos2flag = 0) Then
Picture2.Line (equalpos + 450, Picture2.CurrentY + 0)-(maxX, Picture2.CurrentY + 0) 'redraw divider bar to compensate for denominator length
End If

If (overflag = 1) And (invosflag = 0) And (invos2flag = 1) Then
Picture2.Line (equalpos + 450, Picture2.CurrentY - 250)-(maxX, Picture2.CurrentY - 250) 'redraw divider bar to compensate for denominator length
End If

If (overflag = 1) And (invosflag = 1) And (invos2flag = 0) Then
Picture2.Line (equalpos + 450, Picture2.CurrentY - 100)-(maxX, Picture2.CurrentY - 100) 'redraw divider bar to compensate for denominator length
End If

If (overflag = 1) And (invosflag = 1) And (invos2flag = 1) Then
Picture2.Line (equalpos + 450, Picture2.CurrentY - 300)-(maxX, Picture2.CurrentY - 300) 'redraw divider bar to compensate for denominator length
End If

End Function


Private Sub Text2_Change() 'update picture constantly

If autoredrawflag = 1 Then
temppos = Text2.SelStart
Command9_Click
Text2.SelStart = temppos
End If

End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--- This where we set the Beginning of the Box
'--- that will be Drawn around the Capture Area
    

    If cropflag = 1 Then
        mbDown = (Button = 1)
        Picture1.MousePointer = vbCrosshair
        
        With Line1
            .X1 = X
            .X2 = X
            .y1 = Y
            .Y2 = Y
        End With
            
        With Line2
            .X1 = X
            .X2 = X
            .y1 = Y
            .Y2 = Y
        End With
            
        With Line3
            .X1 = X
            .X2 = X
            .y1 = Y
            .Y2 = Y
        End With
            
        With Line4
            .X1 = X
            .X2 = X
            .y1 = Y
            .Y2 = Y
        End With
            
        Line1.Visible = True
        Line2.Visible = True
        Line3.Visible = True
        Line4.Visible = True
        
        nOldX = X
        nOldY = Y
    End If

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--- Where we Draw the Box around the Choosen Area as you hold down the Left Mouse
'--- button and Drag in any direction to create a rectangle
    If mbDown Then
        With Line1
            .X1 = nOldX
            .X2 = X
            .y1 = nOldY
            .Y2 = nOldY
        End With
        
        With Line2
            .X1 = nOldX
            .X2 = nOldX
            .y1 = nOldY
            .Y2 = Y
        End With
        
        With Line3
            .X1 = X
            .X2 = X
            .y1 = nOldY
            .Y2 = Y
        End With
        
        With Line4
            .X1 = nOldX
            .X2 = X
            .y1 = Y
            .Y2 = Y
        End With
    End If

End Sub


Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
   If cropflag = 1 Then
   
    On Error Resume Next
    
    Dim XUpperLeft As Long
    Dim YUpperLeft As Long
    Dim XLowerRight As Long
    Dim YLowerRight As Long
    
    Line1.Visible = False
    Line2.Visible = False
    Line3.Visible = False
    Line4.Visible = False
    Picture2.MousePointer = vbDefault
   
    cropflag = 0
    
    '--- Determine the upper left hand corner & lower right hand corner
    '--- XY coordinates.  By doing this, it doesn't matter which
    '--- direction the user "dragged" the rectangle:
    XUpperLeft = Line1.X1
    If Line1.X2 < XUpperLeft Then
        XUpperLeft = Line1.X2
    End If
    With Line2
        If .X1 < XUpperLeft Then
            XUpperLeft = .X1
        End If
        If .X2 < XUpperLeft Then
            XUpperLeft = .X2
        End If
    End With
    
    YUpperLeft = Line1.y1
    If Line1.Y2 < YUpperLeft Then
        YUpperLeft = Line1.Y2
    End If
    With Line2
        If .y1 < YUpperLeft Then
            YUpperLeft = .y1
        End If
        If .Y2 < YUpperLeft Then
            YUpperLeft = .Y2
        End If
    End With
    
    XLowerRight = Line1.X1
    If Line1.X2 > XLowerRight Then
        XLowerRight = Line1.X2
    End If
    With Line2
        If .X1 > XLowerRight Then
            XLowerRight = .X1
        End If
        If .X2 > XLowerRight Then
            XLowerRight = .X2
        End If
    End With
    
    YLowerRight = Line1.y1
    If Line1.Y2 > YLowerRight Then
        YLowerRight = Line1.Y2
    End If
    With Line2
        If .y1 > YLowerRight Then
            YLowerRight = .y1
        End If
        If .Y2 > YLowerRight Then
            YLowerRight = .Y2
        End If
    End With
    
    '--- Selected a single pixel (clicked, no drag)
    If XUpperLeft = XLowerRight Then XLowerRight = XLowerRight + 1
    If YUpperLeft = YLowerRight Then YLowerRight = YLowerRight + 1

    '--- Set Picture1 to the size
    '--- we will paint the Image to
    With Picture4
        .Picture = Picture3.Image
        '.Cls
        DoEvents
        .Width = Abs(Line1.X2 - Line1.X1)  '* Screen.TwipsPerPixelX
        .Height = Abs(Line2.Y2 - Line2.y1) '* Screen.TwipsPerPixelY
    
        '--- Paint the Captured part of the screen to
        '--- Picture4
        .PaintPicture Picture1.Picture, 0, 0, _
            (XLowerRight - XUpperLeft), _
            (YLowerRight - YUpperLeft), _
            XUpperLeft, YUpperLeft, _
            (XLowerRight - XUpperLeft), _
            (YLowerRight - YUpperLeft)  ', opcode
        
        '--- IMPORTANT: DO NOT REMOVE THIS DoEvents! Windows needs to "catchup"
        '--- before can use the "painted" picture.
        DoEvents
        mbDown = False
    End With
    
    '--- Load selected rectangle image into picture box:
    With Picture1
        '--- Incase picture was scrolled over (via scrollbars), reset it's position
        .Left = 240
        .Top = 0
        .Width = Picture4.Width
        .Height = Picture4.Height
        '--- Just to be safe, clear picture before
        '--- loading new image:
        .Picture = Picture3.Image
        '.Cls
        DoEvents
        .Picture = Picture4.Image
        Picture4.Picture = Picture3.Image
       
    End With


End If



End Sub

