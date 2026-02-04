#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once

; Common includes for Base
#include "LibreOfficeBase_Constants.au3"
#include "LibreOfficeBase_Helper.au3"
#include "LibreOfficeBase_Internal.au3"

; Other includes for Base
#include "LibreOfficeBase_Database.au3"
#include "LibreOfficeBase_Doc.au3"
#include "LibreOfficeBase_Form.au3"
#include "LibreOfficeBase_Query.au3"
#include "LibreOfficeBase_Report.au3"
#include "LibreOfficeBase_SQLStatement.au3"
#include "LibreOfficeBase_Table.au3"

;~ _LOBase_ComError_UserFunction(ConsoleWrite)

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for interacting with Libre Office Base.
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
;				   Many functions in this UDF may or may not work with OpenOffice Base, however some settings are definitely for LibreOffice only.
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; ===============================================================================================================================
