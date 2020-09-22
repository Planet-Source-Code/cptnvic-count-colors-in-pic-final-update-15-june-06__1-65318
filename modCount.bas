Attribute VB_Name = "Module1"
Option Explicit
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ Sorry I had to take a vacation from this project for awhile, this code has been laying
'++ around for over a month now... and I'm just now polishing off the edges.  Sometimes,
'++ you've just got to go to the lake and forget about counting colors!
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ This version moves the main API to the GetDIBits api.  While still quirky in early OS
'++ versions, the changes below make it substantially more OS friendly.  I have tested it in
'++ ME and XP... but NOT 98/95 or NT yet... and would be interested to hear of results from
'++ users of those systems.  As usual, I have a high (though often misplaced) level of
'++ confidence that these earlier versions should work fine.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ Since I started this project, results have dropped from several seconds per count (for
'++ an average sized picture, to: .125 - .203 seconds on my XP tests today (1024x768 pic).
'++ Special thanks to Cobein and Robert Rayment for improving on my original code... which
'++ sort of forced me to improve(?) on this one!  I will certify to all interested that
'++ the result of having 2 coders that are smarter than you are taking an interest in your
'++ project is that the results are improved by gazillions... or, atleast alot!  It is
'++ exactly for that reason that I try to visit PSC atleast once a day!
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ Since I've been trying to evaluate several methods of counting these colors, I've tried
'++ to keep the surface level for all methods.  If you spot something un-fair... change it!
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'create a structure to hold bitmap information
Private Type BITMAPINFOHEADER
  biSize          As Long
  biWidth         As Long
  biHeight        As Long
  biPlanes        As Integer
  biBitCount      As Integer
  biCompression   As Long
  biSizeImage     As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed       As Long
  biClrImportant  As Long
End Type

'declare the GetDIBits api
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long

'declare a constant for GetDIBits
Private Const DIB_RGB_COLORS As Long = 0

'just used to score time... not too accurate, and not necessary... but good enough for this purpose
Public Declare Function GetTickCount& Lib "kernel32" ()

Public Function UniqColors(tstPic As Object) As Long
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++ Oddly, I expected this function to be the fastest... which it is not.  So much for
    '++ what I know!  Not only that, but it's a temporary memory pig.
    '++ It's built so that all array memory is dumped on function exit.
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'As in all the functions below, you do not need to use GetObject an object to use the GetDIBits
    'function.  But to really save time, if you can either load the DIB array when you load
    'the picture, or do the equivalent with DMA... the time savings would be significant...
    'though lugging that memory requirement around on the chance that someone would want to
    'count colors in an image should be a large consideration.
    
    'Set-up the DIB information
    Dim picInf As BITMAPINFOHEADER
    'fill in part of the header information that will receive the picture data
    With picInf
        .biSize = 40 'could use: Len(picInf) but I'm trying to save every nano-second
        .biPlanes = 1
        .biHeight = -tstPic.Height ' a top down scan...
            'not needed... just helps keep me from going nuts during de-bug...
            'set to: TestPic.Height for bottom up scan.  Makes no difference speed-wise.
        .biWidth = tstPic.Width
        .biBitCount = 32
    End With
    
    'ColorNdx will tell me if color has been found before... Colour is the long color value to check on ColorNdx.
    Dim ColorNdx(16777216) As Byte, colour As Long ' ColorNdx is a memory pig!  Releases memory on exit
    
    Dim picBits() As Long 'the array that will hold the images color info... built here so it will be destroyed on exiting the function and release memory
    ReDim picBits((tstPic.Height * tstPic.Width) - 1) 'size array to hold the existing long color values
    
    'load the picture's color info into the picBits() array... in this case, a zero based array.
    GetDIBits tstPic.hDC, tstPic.Image, 0&, tstPic.Height, picBits(0), picInf, DIB_RGB_COLORS
    'for those of you not following along, picBits() now holds the long color value (+ alpha value) of all colors
    'in the picture from top-left (because of top down scan) to bottom right.  This is where I believe the problems
    'between OS versions originates.  Of course, I'm a man, I'm probably wrong.  See more below.
    
    ' Count the colors
    UniqColors = 0 'not possible to have a picture with zero colors loaded... but initialize the count to zero anyway
    '--> which reminds me that I should place an error trap for no picture loaded...
    Dim i As Long
    For i = 0 To UBound(picBits)
        colour = picBits(i) And &HFFFFFF 'get long color value and mask white for early OS
        'this should just leave the long "color" value...
        'note: XP will handle picBits(x) without issue... but ME, and probably all less versions
        'will throw a negative value into Colour... so ColorNdx(Colour) generates a subscript error
        'the white mask should remove this issue.
        If ColorNdx(colour) < 1 Then 'Found this color before?
            ColorNdx(colour) = 1 'If not, mark it as found...
            UniqColors = UniqColors + 1 '... and increment the counter
        End If
    Next i
    'and you're done.
End Function
Public Function UniqBitColors(tstPic As Object) As Long
    '++ This function turns out to be pretty fast and more memory efficient than the
    '++ function above.  I got the idea from Robert Rayment (but couldn't get his version
    '++ to work across old OS)... so built this version... speed wise, it seems to be slighly
    '++ faster in the ide but more or less the same as the ColorCountBitsNew function at run time.
    
    'Set-up the DIB information
    Dim picInf As BITMAPINFOHEADER
    'fill in part of the header information that will receive the picture data
    With picInf
        .biSize = 40
        .biPlanes = 1
        .biHeight = -tstPic.Height ' a top down scan...
        .biWidth = tstPic.Width
        .biBitCount = 32
    End With
    
    Dim i As Long, colour As Long
    Dim BArray() As Long, BMask(31) As Long 'dim an array for the bit masks and the bit array
    
    'create bit masks for keeping track of found colors
    BMask(0) = 1
    BMask(31) = &H80000000
    For i = 1 To 30
        BMask(i) = BMask(i - 1) * 2
    Next i
      
    'create a bit array for 2^19 longs = 524,288... zero based = 524287 and don't do the math here
    ReDim BArray(524287)
    
    Dim picBits() As Long 'the array that will hold the images color info... built here so it will be destroyed on exiting the function and release memory
    ReDim picBits((tstPic.Height * tstPic.Width) - 1) 'size array to hold the existing long color values
    UniqBitColors = 0
    
    'load the picture's color info into the picBits() array
    GetDIBits tstPic.hDC, tstPic.Image, 0&, tstPic.Height, picBits(0), picInf, DIB_RGB_COLORS
    
    'count the colors...
    For i = 0 To UBound(picBits)
        colour = picBits(i) And &HFFFFFF
        'Found this color before?
        If (BArray(colour \ 32) And BMask(colour And 31)) = 0 Then
            BArray(colour \ 32) = BArray(colour \ 32) Or BMask(colour And 31) 'If not, mark it as found...
            UniqBitColors = UniqBitColors + 1 '... and increment the counter
        End If
    Next i
End Function
Public Function ColorCountBits(tstPic As Object) As Long
'Public Function ColorCountBits(lhBitmap As Long) As Long

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ Modified from Robert Rayment's submission to PSC:
'++ http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=65420&lngWid=1
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ Modified to use GetDIBitsn by CptnVic
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ WARNING>>>>> This version is XP Friendly only. ME (for sure) and probably other older
'++ OS versions (95/98/NT3.?) will crash with subscript out of range errors.
'++ See the function following this one for a version of this code that is OS friendly.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'Set-up the DIB information
    Dim picInf As BITMAPINFOHEADER
    'fill in part of the header information that will receive the picture data
    With picInf
        .biSize = 40
        .biPlanes = 1
        .biHeight = -tstPic.Height ' a top down scan...
        .biWidth = tstPic.Width
        .biBitCount = 32
    End With
    '------------------->
    'the following used to be in a sub... moved here for fairness in testing since has to be done
    Dim k As Long, Pow2(7) As Byte
    For k = 0 To 7
        Pow2(k) = 2 ^ k
    Next k
    '-------------------<
    
    Dim bPicBytes() As Long
    Dim lByteLen As Long
    Dim i As Long
    Dim lCounter As Long
    Dim bColorTable(0 To 16777216 \ 8) As Byte ' 16/8 = 2 MB
    Dim ind As Long
    Dim bitpos As Long
    'Dim addr As Single' not used here
    'addr = VarPtr(bColorTable(0)) ' Check if Aligned on 8 byte boundary' not used here
    
    'GetObject lhBitmap, Len(tBitmap), tBitmap ' not needed here
    'lByteLen = (tBitmap.bmWidth * 4) * tBitmap.bmHeight '< ditto
    lByteLen = (tstPic.Height * tstPic.Width) - 1
    ReDim bPicBytes(0 To lByteLen)
    'GetBitmapBits lhBitmap, UBound(bPicBytes), bPicBytes(1)
    GetDIBits tstPic.hDC, tstPic.Image, 0&, tstPic.Height, bPicBytes(0), picInf, DIB_RGB_COLORS
    
    'For i = 1 To lByteLen \ 4
    For i = 0 To lByteLen
        ind = bPicBytes(i) \ 8 ' <-- this will blow up in ME/others... see the function that follows for fix
         bitpos = bPicBytes(i) And 7
         If (bColorTable(ind) And Pow2(bitpos)) = 0 Then
               bColorTable(ind) = bColorTable(ind) Or Pow2(bitpos)
               lCounter = lCounter + 1
         End If
    Next
    ColorCountBits = lCounter
End Function

Public Function ColorCountBitsNew(tstPic As Object) As Long
'Public Function ColorCountBits(lhBitmap As Long) As Long

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ Modified from Robert Rayment's submission to PSC:
'++ http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=65420&lngWid=1
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ Modified to use GetDIBits by CptnVic
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ This version should be more OS friendly... I've only tested it on ME and XP so far.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'Set-up the DIB information
    Dim picInf As BITMAPINFOHEADER
    'fill in part of the header information that will receive the picture data
    With picInf
        .biSize = 40
        .biPlanes = 1
        .biHeight = -tstPic.Height ' a top down scan...
        .biWidth = tstPic.Width
        .biBitCount = 32
    End With
    '------------------->
    'the following used to be in a sub... moved here for fairness in testing since has to be done
    Dim k As Long, Pow2(7) As Byte
    For k = 0 To 7
        Pow2(k) = 2 ^ k
    Next k
    '-------------------<
        
    'Dim tBitmap As BITMAP
    Dim bPicBytes() As Long
    Dim lByteLen As Long
    Dim i As Long
    Dim lCounter As Long
    Dim bColorTable(0 To 16777216 \ 8) As Byte ' 16/8 = 2 MB
    Dim ind As Long
    Dim bitpos As Long
    'Dim addr As Single' not used here
    'addr = VarPtr(bColorTable(0)) ' Check if Aligned on 8 byte boundary' not used here
    'GetObject lhBitmap, Len(tBitmap), tBitmap  ' not needed here
    'lByteLen = (tBitmap.bmWidth * 4) * tBitmap.bmHeight  ' not needed here
    lByteLen = (tstPic.Height * tstPic.Width) - 1
    ReDim bPicBytes(0 To lByteLen)
    'GetBitmapBits lhBitmap, UBound(bPicBytes), bPicBytes(1)
    GetDIBits tstPic.hDC, tstPic.Image, 0&, tstPic.Height, bPicBytes(0), picInf, DIB_RGB_COLORS
    
    'For i = 1 To lByteLen \ 4
    For i = 0 To lByteLen
        ind = (bPicBytes(i) And &HFFFFFF) \ 8 '<--- apply white mask
         bitpos = bPicBytes(i) And 7
         If (bColorTable(ind) And Pow2(bitpos)) = 0 Then
               bColorTable(ind) = bColorTable(ind) Or Pow2(bitpos)
               lCounter = lCounter + 1
         End If
    Next
    ColorCountBitsNew = lCounter
End Function

