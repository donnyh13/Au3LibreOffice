#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Internal.au3"

; #INDEX# =======================================================================================================================
; Title .........: Libre Office Writer (LOWriter)
; AutoIt Version : v3.3.16.1
; UDF Version    : 0.0.0.3
; Description ...: Provides basic functionality through Autoit for interacting with Libre Office Writer.
; Author(s) .....: donnyh13, mLipok
; Sources .......: jguinch -- Printmgr.au3, used (_PrintMgr_EnumPrinter);
;					mLipok -- OOoCalc.au3, used (__OOoCalc_ComErrorHandler_UserFunction,_InternalComErrorHandler,
;						-- WriterDemo.au3, used _CreateStruct;
;					Andrew Pitonyak & Laurent Godard (VersionGet);
;					Leagnus & GMK -- OOoCalc.au3, used (SetPropertyValue)
; Dll ...........:
; Note...........: Tips/templates taken from OOoCalc UDF written by user GMK; also from Word UDF by user water.
;					I found the book by Andrew Pitonyak very helpful also, titled, "OpenOffice.org Macros Explained;
;						OOME Third Edition".
;					Of course, this UDF is written using the English version of LibreOffice, and may only work for the English
;						version of LibreOffice installations. Many functions in this UDF may or may not work with OpenOffice
;						Writer, however some settings are definitely for LibreOffice only.
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOWriter_ShapesGetNames
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ShapesGetNames
; Description ...: Return a list of Shape names contained in a document.
; Syntax ........: _LOWriter_ShapesGetNames(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object.  A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 2D Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Shapes Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning 2D Array containing a list of Shape names contained in a document, the first column ($aArray[0][0] contains the shape name, the second column ($aArray[0][1] contains the shape's Implementation name. See Remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Implementation name identifies what type of shape object it is, as there can be multiple things counted as "Shapes", such as Text Frames etc.
;				   I have found the three Implementation names being returned, SwXTextFrame, indicating the shape is actually a Text Frame, SwXShape, is a regular shape such as a line, circle etc. And "SwXTextGraphicObject", which is an image / picture. There may be other return types I haven't found yet.
;				   Images inserted into the document are also listed as TextFrames in the shapes category. There isn't an easy way to differentiate between them yet, see _LOWriter_FramesGetNames, to search for Frames in the shapes category.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ShapesGetNames(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asShapeNames[0][2]
	Local $oShapes

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	$oShapes = $oDoc.DrawPage()
	If Not IsObj($oShapes) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If $oShapes.hasElements() Then
		ReDim $asShapeNames[$oShapes.getCount()][2]
		For $i = 0 To $oShapes.getCount() - 1
			$asShapeNames[$i][0] = $oShapes.getByIndex($i).Name()
			If $oShapes.getByIndex($i).supportsService("com.sun.star.drawing.Text") Then
				; If Supports Text Method, then get that impl. name, else just the regular impl. name.
				$asShapeNames[$i][1] = $oShapes.getByIndex($i).Text.ImplementationName()
			Else
				$asShapeNames[$i][1] = $oShapes.getByIndex($i).ImplementationName()
			EndIf

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, UBound($asShapeNames), $asShapeNames)
EndFunc   ;==>_LOWriter_ShapesGetNames
