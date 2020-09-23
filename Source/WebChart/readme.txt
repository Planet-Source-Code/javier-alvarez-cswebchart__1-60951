---------------------------------------------------
Overview

CSWebChart is a dll Activex write in vb 6. 
create pie and bar charts.

Use GDI+ to draw and can save the chart in
bmp, gif, jpg, png y tiff. 	 	

---------------------------------------------------
Dependencies:

gdiplus.dll        - http://www.microsoft.com/downloads/details.aspx?FamilyID=6a63ab9c-df12-4d41-933c-be590feaa05a&DisplayLang=en
GpGDIPlus10.dll (*)- http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=45451&lngWId=1

(*) This dll have some bugs. I fixed this and send to author.
You can check if the code is fixed if compare with de code 
is bellow.

---------------------------------------------------
Version .Net

This code is based in Blong'source that you can 
find in:

http://www.codeproject.com/aspnet/webchart.asp

It's vb.net code.


---------------------------------------------------
GDI+

To replace .Net framework I use GDI+. 
Thanks to Genghis Khan's GDIPlus dll. 
That dll implement a wrapper COM over GDI+.

http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=45451&lngWId=1

This code have 3 bugs.

After you download the code have to replace this functions:

------------------------------------------------------------------------------------------------------------------------------------------
1) I detect a bug in:

Public Function MeasureStringSizeFormat(ByVal sText As String, _
                                        ByVal lLength As Long, _
                                        ByVal font As cFont, _
                                        ByVal layoutRectSize As cSizeF, _
                                        ByVal StringFormat As cStringFormat, _
                                        ByRef size As cSizeF, _
                                        Optional ByRef codepointsFitted As Long = 0, _
                                        Optional ByRef linesFilled As Long = 0) As GpStatus
    Dim uR                              As RECTF
    Dim lFont                           As Long
    Dim lFormat                         As Long
    Dim clsRect                         As cRectF
    
    If layoutRectSize Is Nothing Then m_uLastStatus = InvalidParameter: MeasureStringSizeFormat = m_uLastStatus: Exit Function
    If font Is Nothing Then lFont = 0& Else lFont = font.NativeFont
    If StringFormat Is Nothing Then lFormat = 0& Else lFormat = StringFormat.NativeStringFormat

'bug missing new
    Set clsRect = New cRectF
'end bug

    clsRect.Create 0#, 0#, layoutRectSize.Width, layoutRectSize.Height
    m_uLastStatus = GdipMeasureString(m_lngGraphics, StrConv(sText, vbUnicode), _
                                      lLength, lFont, clsRect.ToType, lFormat, _
                                      uR, codepointsFitted, linesFilled)
    If m_uLastStatus = Ok Then
       If size Is Nothing Then Set size = New cSizeF
       size.Create uR.Right - uR.Left, uR.Bottom - uR.Top
    End If
    Set clsRect = Nothing
    MeasureStringSizeFormat = m_uLastStatus
End Function

------------------------------------------------------------------------------------------------------------------------------------------
2) I detect a bug in:


' draws a string based on a font and an origin for the string.
Public Function DrawStringPointF(ByVal sText As String, _
                                 ByVal lLength As Long, _
                                 ByVal font As cFont, _
                                 ByVal origin As cPointF, _
                                 ByVal Brush As cBrush) As GpStatus
    Dim lFont                    As Long
    Dim lBrush                   As Long
    Dim clsRect                  As New cRectF

'bug call with this params fails because the rect dimension produce vertical output when you call with x=0 and y=0
'    clsRect.Create origin.x, origin.y, 0#, 0#
    clsRect.Create origin.x, origin.y, &HFFFFF, &HFFFFF
'end bug

    If font Is Nothing Then lFont = 0& Else lFont = font.NativeFont
    If Brush Is Nothing Then lBrush = 0& Else lBrush = Brush.NativeBrush
    m_uLastStatus = GdipDrawString(m_lngGraphics, _
                                   StrConv(sText, vbUnicode), _
                                   lLength, _
                                   lFont, _
                                   clsRect.ToType, _
                                   ByVal 0&, _
                                   lBrush)
    Set clsRect = Nothing
    DrawStringPointF = m_uLastStatus
End Function

------------------------------------------------------------------------------------------------------------------------------------------
3) other bug

Public Function SaveImageToFile(ByVal filename As String, Optional ByVal SaveParameters As cImageSaveParameters = Nothing) As Boolean
   Dim uclsidEncoder As Clsid
   Dim uencoderParams As EncoderParameters

'---------------------------------------------------------------     
'  this parasit call fixe de bug. Horrible but effective.
'
   SaveParameters.GetEncoderCLSID uclsidEncoder
'---------------------------------------------------------------
 
   If SaveParameters.GetEncoderCLSID(uclsidEncoder) = False Then Exit Function
   If SaveParameters.GetEncoderParameters(uencoderParams) Then
      m_uLastStatus = GdipSaveImageToFile(m_lngImage, StrConv(filename, vbUnicode), uclsidEncoder, uencoderParams)
   Else
      m_uLastStatus = GdipSaveImageToFile(m_lngImage, StrConv(filename, vbUnicode), uclsidEncoder, ByVal 0&)
   End If
   If m_uLastStatus <> Ok Then Exit Function
   SaveImageToFile = True
End Function

