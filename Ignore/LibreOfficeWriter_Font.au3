#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"

; Common includes for Writer
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Helper.au3"
#include "LibreOfficeWriter_Internal.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Listing and querying available L.O. Writer Fonts.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOWriter_FontExists
; _LOWriter_FontsList
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FontExists
; Description ...: Tests whether a Document has a specific font available by name.
; Syntax ........: _LOWriter_FontExists(ByRef $oDoc, $sFontName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sFontName           - a string value. The Font name to search for.
; Return values .: Success: Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sFontName not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Font list.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean  = Success. Returns True if the Font is available, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function may cause a processor usage spike for a moment or two. If you wish to eliminate this, comment out the current sleep function and place a sleep(10) in its place.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FontExists(ByRef $oDoc, $sFontName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atFonts

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	$atFonts = $oDoc.getCurrentController().getFrame().getContainerWindow().getFontDescriptors()
	If Not IsArray($atFonts) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	For $i = 0 To UBound($atFonts) - 1
		If $atFonts[$i].Name = $sFontName Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next
	Return SetError($__LO_STATUS_SUCCESS, 0, False)

EndFunc   ;==>_LOWriter_FontExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FontsList
; Description ...: Retrieve a list of currently available fonts.
; Syntax ........: _LOWriter_FontsList(ByRef $oDoc)
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Font list.
;				   --Success--
;				   @Error 0 @Extended ? Return Array  = Success. Returns a 4 Column Array, @extended is set to the number of results. See remarks
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Many fonts will be listed multiple times, this is because of the varying settings for them, such as bold,
;					Italic, etc. Style Name is really a repeat of weight(Bold) and Slant (Italic) settings, but is included for
;					easier processing if required. From personal tests, Slant only returns 0 or 2. This function may cause a
;					 processor usage spike for a moment or two.
;				   The returned array will be as follows:
;				   The first column (Array[1][0]) contains the Font Name.
;				   The Second column (Array [1][1] contains the style name (Such as Bold Italic etc.)
;				   The third column (Array[1][2]) contains the Font weight (Bold) See Constants, $LOW_WEIGHT_* as defined in LibreOfficeWriter_Constants.au3;
;				   The fourth column (Array[1][3]) contains the font slant (Italic) See constants, $LOW_POSTURE_* as defined in LibreOfficeWriter_Constants.au3.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FontsList(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atFonts
	Local $asFonts[0][4]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	$atFonts = $oDoc.getCurrentController().getFrame().getContainerWindow().getFontDescriptors()
	If Not IsArray($atFonts) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	ReDim $asFonts[UBound($atFonts)][4]

	For $i = 0 To UBound($atFonts) - 1
		$asFonts[$i][0] = $atFonts[$i].Name()
		$asFonts[$i][1] = $atFonts[$i].StyleName()
		$asFonts[$i][2] = $atFonts[$i].Weight
		$asFonts[$i][3] = $atFonts[$i].Slant() ; only 0 or 2?
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, UBound($atFonts), $asFonts)
EndFunc   ;==>_LOWriter_FontsList
