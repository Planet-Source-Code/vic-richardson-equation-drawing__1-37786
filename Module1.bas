Attribute VB_Name = "BMPconvert"
 Type BITMAPFILEHEADER   '14 Bytes
         bfType As Integer
         bfSize As Long
         bfReserved1 As Integer
         bfReserved2 As Integer
         bfOffBits As Long
  End Type

 Type BITMAPINFOHEADER '40 bytes
          biSize As Long
          biWidth As Long
          biHeight As Long
          biPlanes As Integer
          biBitCount As Integer
          biCompression As Long
          biSizeImage As Long
          biXPelsPerMeter As Long
          biYPelsPerMeter As Long
          biClrUsed As Long
          biClrImportant As Long
  End Type

Private Type RGBQUAD
   rgbBlue           As Byte
   rgbGreen          As Byte
   rgbRed            As Byte
   rgbReserved       As Byte
End Type

Type BITMAPINFO_1   ' For monochrome
        bmiHeader As BITMAPINFOHEADER
      '  bmiColors As String * 8 ' Byte '* 8
       bmiColors(1) As RGBQUAD

End Type

 Type BITMAPINFO_4   ' For 4 bits per pixel (16 colors)
        bmiHeader As BITMAPINFOHEADER
       ' bmiColors As String * 64
        bmiColors(15) As RGBQUAD
 End Type

 Type BITMAPINFO_8   ' For 8 bits per pixel (256 colors)
        bmiHeader As BITMAPINFOHEADER
       ' bmiColors As String * 1024
        bmiColors(255) As RGBQUAD
 End Type

 Declare Function GlobalAlloc& Lib "Kernel32" (ByVal wFlags&, ByVal dwBytes&)
 Declare Function GlobalLock& Lib "Kernel32" (ByVal hMem&)
 Declare Function GlobalFree& Lib "Kernel32" (ByVal hMem&)
 Declare Function GlobalUnlock& Lib "Kernel32" (ByVal hMem&)
 Declare Function DeleteDC& Lib "GDI32" (ByVal hDC&)
 Declare Function hwrite& Lib "Kernel32" Alias "_hwrite" (ByVal hf&, ByVal hpvBuffer&, ByVal cbBuffer&)
 Declare Function GetDIBits Lib "GDI32" (ByVal aHDC&, ByVal hBitmap, ByVal nStartScan&, ByVal nNumScans&, ByVal LpBits As Any, lpBI As BITMAPINFO_1, ByVal wUsage&) As Long
 Declare Function GetDIBits1 Lib "GDI32" Alias "GetDIBits" (ByVal aHDC&, ByVal hBitmap&, ByVal nStartScan&, ByVal nNumScans&, LpBits As Any, lpBI As BITMAPINFO_1, ByVal wUsage&) As Long
 Declare Function GetDIBits4 Lib "GDI32" Alias "GetDIBits" (ByVal aHDC&, ByVal hBitmap&, ByVal nStartScan&, ByVal nNumScans&, LpBits As Any, lpBI As BITMAPINFO_4, ByVal wUsage&) As Long
 Declare Function GetDIBits8 Lib "GDI32" Alias "GetDIBits" (ByVal aHDC&, ByVal hBitmap&, ByVal nStartScan&, ByVal nNumScans&, LpBits As Any, lpBI As BITMAPINFO_8, ByVal wUsage&) As Long
 
 Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal LpszDir As String, ByVal FsShowCmd As Long) As Long




' SubRoutine follows
Public Function SaveNewBMP(pic As PictureBox, FileName As String, ByVal NumColors As Integer)
 ' For VB5 /6 (32 bit)
 ' Save picture box at reduced color depth bitmap to reduce file size when full color is not needed.
 ' NumColors can be either:
 '   2, 16, 256 for the number of colors, or
 '   1, 4, 8    for the number color bits (bits per pixel.

 ' Has provision to customize foreground and background colors in 2-color version

 '  ** Note: this SUB sets the source picturebox to scalemode 3 (pixels).
 '     Restore scalemode and any custom scale properties if you may
 '     add more graphics or print to the picture box.
 '
 ' Type definitions and declarations needed for this sub are shown at the end
 '
 ' Programmer = Dick Petschauer;  RJPetsch@Aol.com   Apr, 1999
 ' Tested a little.  Not gauranteed to be bug free.

 Dim SaveFileHeader       As BITMAPFILEHEADER
 Dim SaveBITMAPINFO_1     As BITMAPINFO_1
 Dim SaveBITMAPINFO_4     As BITMAPINFO_4
 Dim SaveBITMAPINFO_8     As BITMAPINFO_8
 Dim SaveBits()           As Byte
 Dim BitsPerPixel         As Integer

 Dim Num32bitWords        As Integer
 Dim Buffersize           As Long
 Dim FileNum              As Integer
 Dim Retval&              ' Temporary returns and handles follow

 ' Set the Scalemode to pixels (*** Note: this also sets the source PictureBox to scalemode 3)
  pic.ScaleMode = 3  ' Pixels

 ' Allow for use of color bits to be used instead of the number of colors:
  
  If NumColors = 1 Then
  NumColors = 2
  End If
  
  If NumColors = 4 Then
  NumColors = 16
  End If
  
  If NumColors = 8 Then
  NumColors = 256
  End If
  
 ' Check for illegal NumColors. Set to default as monochrome.
  If NumColors <> 2 And NumColors <> 16 And NumColors <> 256 Then
  NumColors = 2
  End If
  
  BitsPerPixel = Log(NumColors) / Log(2)

 ' *** Calculate the buffer for the pixel data
  Num32bitWords = (pic.ScaleWidth * BitsPerPixel) \ 32   ' Integer divide
  
  If pic.ScaleWidth Mod 32 > 0 Then
  Num32bitWords = Num32bitWords + 1 'End each scan line on 32-bit boundary
  End If
  
  Buffersize = Num32bitWords * 4 * pic.ScaleHeight  ' 8-bit Bytes; 8 pixels per byte for mono; 2 for 16 color; 4 for 256 color
 ' Buffersize can be larger than this; results in larger bitmap file.

  ReDim SaveBits(0 To Buffersize - 1)

  Debug.Print Buffersize; UBound(SaveBits)
  ' *** Fill the Bitmap info
  If BitsPerPixel = 1 Then
   SaveBITMAPINFO_1.bmiHeader.biSize = 40
   SaveBITMAPINFO_1.bmiHeader.biWidth = pic.ScaleWidth
   SaveBITMAPINFO_1.bmiHeader.biHeight = pic.ScaleHeight
   SaveBITMAPINFO_1.bmiHeader.biPlanes = 1
   SaveBITMAPINFO_1.bmiHeader.biBitCount = BitsPerPixel
   SaveBITMAPINFO_1.bmiHeader.biCompression = 0
   SaveBITMAPINFO_1.bmiHeader.biClrUsed = 0
   SaveBITMAPINFO_1.bmiHeader.biClrImportant = 0
   SaveBITMAPINFO_1.bmiHeader.biSizeImage = Buffersize
  End If

 If BitsPerPixel = 4 Then
   SaveBITMAPINFO_4.bmiHeader.biSize = 40
   SaveBITMAPINFO_4.bmiHeader.biWidth = pic.ScaleWidth
   SaveBITMAPINFO_4.bmiHeader.biHeight = pic.ScaleHeight
   SaveBITMAPINFO_4.bmiHeader.biPlanes = 1
   SaveBITMAPINFO_4.bmiHeader.biBitCount = BitsPerPixel
   SaveBITMAPINFO_4.bmiHeader.biCompression = 0
   SaveBITMAPINFO_4.bmiHeader.biClrUsed = 0
   SaveBITMAPINFO_4.bmiHeader.biClrImportant = 0
   SaveBITMAPINFO_4.bmiHeader.biSizeImage = Buffersize
  End If

 If BitsPerPixel = 8 Then
   SaveBITMAPINFO_8.bmiHeader.biSize = 40
   SaveBITMAPINFO_8.bmiHeader.biWidth = pic.ScaleWidth
   SaveBITMAPINFO_8.bmiHeader.biHeight = pic.ScaleHeight
   SaveBITMAPINFO_8.bmiHeader.biPlanes = 1
   SaveBITMAPINFO_8.bmiHeader.biBitCount = BitsPerPixel
   SaveBITMAPINFO_8.bmiHeader.biCompression = 0
   SaveBITMAPINFO_8.bmiHeader.biClrUsed = 0
   SaveBITMAPINFO_8.bmiHeader.biClrImportant = 0
   SaveBITMAPINFO_8.bmiHeader.biSizeImage = Buffersize
 End If

If BitsPerPixel = 1 Then
Retval& = GetDIBits1(pic.hDC, pic.Image, 0, pic.ScaleHeight, SaveBits(0), SaveBITMAPINFO_1, DIB_RGB_COLORS)
End If

If BitsPerPixel = 4 Then
Retval& = GetDIBits4(pic.hDC, pic.Image, 0, pic.ScaleHeight, SaveBits(0), SaveBITMAPINFO_4, DIB_RGB_COLORS)
End If

If BitsPerPixel = 8 Then
Retval& = GetDIBits8(pic.hDC, pic.Image, 0, pic.ScaleHeight, SaveBits(0), SaveBITMAPINFO_8, DIB_RGB_COLORS)
End If

  If BitsPerPixel = 1 Then
  BiLen = Len(SaveBITMAPINFO_1)
  End If
  
  If BitsPerPixel = 4 Then
  BiLen = Len(SaveBITMAPINFO_4)
  End If
  
  If BitsPerPixel = 8 Then
  BiLen = Len(SaveBITMAPINFO_8)
  End If
  
  ' *** Make and fill a Header for the new bitmap
    SaveFileHeader.bfType = &H4D42      ' "BM" for Bitmap; first two characters in file
    SaveFileHeader.bfSize = Len(SaveFileHeader) + BiLen + Buffersize
    SaveFileHeader.bfOffBits = Len(SaveFileHeader) + BiLen

  ' Here is where you can customize the foreground and background colors for a 2-color bitmap
  ' For the Foreground color add this line:
     ' Mid$(SaveBITMAPINFO_1.bmiColors, 1, 3) = Chr$(ForeBlue) + Chr$ (ForeGreen) + Chr$(ForeRed)
  ' For the Foreground color add this line:
     ' Mid$(SaveBITMAPINFO_1.bmiColors, 5, 3) = Chr$(BackBlue) + Chr$ (BackGreen) + Chr$(BackRed)
  ' Where ForeBlue, etc are integers from 0 to 255 that represent the respective color strength.
  ' For white, set all to 255; Black all to 0.
  ' Defaults: Black foreground, White background.

  ' *** Save the Bitmap Header and BitmapInfo to disk
  ' First remove old bitmap file if there
    If Dir$(FileName) <> "" Then
    Kill FileName
    End If
    
    FileNum = FreeFile
    Open FileName For Binary As FileNum

    Put FileNum, , SaveFileHeader

    If BitsPerPixel = 1 Then
    Put FileNum, , SaveBITMAPINFO_1
    End If
    If BitsPerPixel = 4 Then
    Put FileNum, , SaveBITMAPINFO_4
    End If
    If BitsPerPixel = 8 Then
    Put FileNum, , SaveBITMAPINFO_8
    End If
    
   Put FileNum, , SaveBits()


    Close FileNum


End Function
