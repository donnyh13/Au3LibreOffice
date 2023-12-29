#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once
; Common includes for Writer
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Helper.au3"
#include "LibreOfficeWriter_Internal.au3"

; Other includes for Writer
#include "LibreOfficeWriter_Cell.au3"
#include "LibreOfficeWriter_Char.au3"
#include "LibreOfficeWriter_Cursor.au3"
#include "LibreOfficeWriter_DirectFormating.au3"
#include "LibreOfficeWriter_Doc.au3"
#include "LibreOfficeWriter_Field.au3"
#include "LibreOfficeWriter_Font.au3"
#include "LibreOfficeWriter_FootEndNotes.au3"
#include "LibreOfficeWriter_Frame.au3"
#include "LibreOfficeWriter_Images.au3"
#include "LibreOfficeWriter_Num.au3"
#include "LibreOfficeWriter_Page.au3"
#include "LibreOfficeWriter_Par.au3"
#include "LibreOfficeWriter_Shapes.au3"
#include "LibreOfficeWriter_Table.au3"


; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for interacting with Libre Office Writer.
; Author(s) .....: donnyh13, mLipok
; Sources .......: Andrew Pitonyak & Laurent Godard. Useful Macro Information, section 5.7.1. OOo version. Used for VersionGet;
;				   jguinch -- Printmgr.au3. Function used: _PrintMgr_EnumPrinter.
;				   Leagnus & GMK -- OOoCalc.au3. Function used: SetPropertyValue.
;				   mLipok  -- OOoCalc.au3. Function used: __OOoCalc_ComErrorHandler_UserFunction,_InternalComErrorHandler.
;						   -- WriterDemo.au3. Function used: _CreateStruct;
; Dll ...........:
; Note...........: Tips/templates taken from OOoCalc UDF written by user GMK; also from Word UDF by user water.
;				   I found the book by Andrew Pitonyak very helpful also, titled, "OpenOffice.org Macros Explained; OOME Third Edition".
;				   Of course, this UDF is written using the English version of LibreOffice, and may only work for the English version of LibreOffice installations.
;				   Many functions in this UDF may or may not work with OpenOffice Writer, however some settings are definitely for LibreOffice only.
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; ===============================================================================================================================
