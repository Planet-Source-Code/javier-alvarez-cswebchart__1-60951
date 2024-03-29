VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'/////////////////////////////////////////////////////////////////////////////////////////
' CopyRight © 2003-2005 Javier Alvarez (javier@crowsoft.com.ar)

' This library is free software; you can redistribute it and/or modify
' it under the terms of the GNU Lesser Gereral Public Licence as published
' by the Free Software Foundation; either version 2 of the Licence,
' or (at your opinion) any later version.

' This library is distributed in the hope that it will be usefull,
' but WITHOUT ANY WARRANTY; without even the implied warranty of merchantability
' or fitness for a particular purpose. See the GNU Lesser General Public Licence
' for more details.

' You should have received a copy of the GNU Lesser General Public Licence
' along with this library; if not, write to the Free Software Foundation, Inc.,
' 59 Temple Place, Suite 330, Boston, Ma 02111-1307 USA.

' Visit CrowSoft.
'    http://www.crowsoft.com.ar
'/////////////////////////////////////////////////////////////////////////////////////////

Public Enum ePixelFormat
  Alpha
  Canonical
  DontCare
  Extended
  Format1bppIndexed = &H30101
  Format4bppIndexed = &H30402
  Format8bppIndexed = &H30803
  Format16bppGreyScale = &H101004
  Format16bppRGB555 = &H21005
  Format16bppRGB565 = &H21006
  Format16bppArgb1555 = &H61007
  Format24bppRGB = &H21808
  Format32bppRGB = &H22009
  Format32bppARGB = &H26200A
  Format32bppPARGB = &HE200B
  Format48bppRGB = &H10300C
  Format64bppARGB = &H34400D
  Format64bppPARGB = &H1C400E
  Gdi
  Indexed
  Max
  PAlpha
  Undefined
End Enum
