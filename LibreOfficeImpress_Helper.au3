#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Impress
#include "LibreOfficeImpress_Constants.au3"
#include "LibreOfficeImpress_Internal.au3"

; Other includes for Impress

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Functions used for creating, modifying and retrieving data for use in various functions in LibreOffice UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOImpress_ComError_UserFunction
; _LOImpress_CursorInsertString
; _LOImpress_FontExists
; _LOImpress_FontsGetNames
; _LOImpress_GradientMulticolorAdd
; _LOImpress_GradientMulticolorDelete
; _LOImpress_GradientMulticolorModify
; _LOImpress_TransparencyGradientMultiAdd
; _LOImpress_TransparencyGradientMultiDelete
; _LOImpress_TransparencyGradientMultiModify
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ComError_UserFunction
; Description ...: Set a UserFunction to receive the Fired COM Error Error outside of the UDF.
; Syntax ........: _LOImpress_ComError_UserFunction([$vUserFunction = Default[, $vParam1 = Null[, $vParam2 = Null[, $vParam3 = Null[, $vParam4 = Null[, $vParam5 = Null]]]]]])
; Parameters ....: $vUserFunction       - [optional] a Function or Keyword. Default value is Default. Accepts a Function, or the Keyword Default and Null. If called with a User function, the function may have up to 5 required parameters.
;                  $vParam1             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
;                  $vParam2             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
;                  $vParam3             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
;                  $vParam4             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
;                  $vParam5             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
; Return values .: Success: 1 or UserFunction.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $vUserFunction Not a Function, or Default keyword, or Null Keyword.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Successfully set the UserFunction.
;                  @Error 0 @Extended 0 Return 2 = Successfully cleared the set UserFunction.
;                  @Error 0 @Extended 0 Return Function = Returning the set UserFunction.
; Author ........: mLipok
; Modified ......: donnyh13 - Added a clear UserFunction without error option. Also added parameters option.
; Remarks .......: The first parameter passed to the User function will always be the COM Error object. See below.
;                  Every COM Error will be passed to that function. The user can then read the following properties. (As Found in the COM Reference section in Autoit HelpFile.) Using the first parameter in the UserFunction.
;                  For Example MyFunc($oMyError)
;                  - $oMyError.number The Windows HRESULT value from a COM call
;                  - $oMyError.windescription The FormatWinError() text derived from .number
;                  - $oMyError.source Name of the Object generating the error (contents from ExcepInfo.source)
;                  - $oMyError.description Source Object's description of the error (contents from ExcepInfo.description)
;                  - $oMyError.helpfile Source Object's help file for the error (contents from ExcepInfo.helpfile)
;                  - $oMyError.helpcontext Source Object's help file context id number (contents from ExcepInfo.helpcontext)
;                  - $oMyError.lastdllerror The number returned from GetLastError()
;                  - $oMyError.scriptline The script line on which the error was generated
;                  - NOTE: Not all properties will necessarily contain data, some will be blank.
;                  If MsgBox or ConsoleWrite functions are passed to this function, the error details will be displayed using that function automatically.
;                  If called with Default keyword, the current UserFunction, if set, will be returned.
;                  If called with Null keyword, the currently set UserFunction is cleared and only the internal ComErrorHandler will be called for COM Errors.
;                  The stored UserFunction (besides MsgBox and ConsoleWrite) will be called as follows: UserFunc($oComError,$vParam1,$vParam2,$vParam3,$vParam4,$vParam5)
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_ComError_UserFunction($vUserFunction = Default, $vParam1 = Null, $vParam2 = Null, $vParam3 = Null, $vParam4 = Null, $vParam5 = Null)
	#forceref $vParam1, $vParam2, $vParam3, $vParam4, $vParam5

	; If user does not set a function, UDF must use internal function to avoid AutoItError.
	Local Static $vUserFunction_Static = Default
	Local $avUserFuncWParams[@NumParams]

	If $vUserFunction = Default Then
		; just return stored static User Function variable

		Return SetError($__LO_STATUS_SUCCESS, 0, $vUserFunction_Static)

	ElseIf IsFunc($vUserFunction) Then
		; If User called Parameters, then add to array.
		If @NumParams > 1 Then
			$avUserFuncWParams[0] = $vUserFunction
			For $i = 1 To @NumParams - 1
				$avUserFuncWParams[$i] = Eval("vParam" & $i)
				; set static variable
			Next
			$vUserFunction_Static = $avUserFuncWParams

		Else
			$vUserFunction_Static = $vUserFunction
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 0, 1)

	ElseIf $vUserFunction = Null Then
		; Clear User Function.
		$vUserFunction_Static = Default

		Return SetError($__LO_STATUS_SUCCESS, 0, 2)

	Else
		; return error as an incorrect parameter was passed to this function

		Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	EndIf
EndFunc   ;==>_LOImpress_ComError_UserFunction

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_CursorInsertString
; Description ...: Insert a string at a cursor position.
; Syntax ........: _LOImpress_CursorInsertString(ByRef $oCursor, $sString[, $bOverwrite = False])
; Parameters ....: $oCursor             - [in/out] an object. A Text Cursor Object returned from any Cursor Object creation or retrieval functions.
;                  $sString             - a string value. A String to insert.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, and the cursor object has text selected, the selection is overwritten, else if False, the string is inserted to the left of the selection. If there are multiple selections, the string is inserted to the left of the last selection, and none are overwritten.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sString not a string..
;                  @Error 1 @Extended 3 Return 0 = $bOverwrite not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. String was successfully inserted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Warning! For some reason this function doesn't seem to set the modified status to True. Changes could be inadvertently lost due to this, if the user closes without saving.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_CursorInsertString(ByRef $oCursor, $sString, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oCursor.Text.insertString($oCursor, $sString, $bOverwrite)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_CursorInsertString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_FontExists
; Description ...: Tests whether a specific font exists by name.
; Syntax ........: _LOImpress_FontExists($sFontName[, $oDoc = Null])
; Parameters ....: $sFontName           - a string value. The Font name to search for.
;                  $oDoc                - [optional] an object. Default is Null. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sFontName not a String.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a "com.sun.star.frame.Desktop" Object.
;                  @Error 2 @Extended 3 Return 0 = Failed to create a Property Struct.
;                  @Error 2 @Extended 4 Return 0 = Failed to create a new Document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Font list.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returning True if the Font is available, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $oDoc is optional, if not called, an Impress Document is created invisibly to perform the check.
; Related .......: _LOImpress_FontsGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_FontExists($sFontName, $oDoc = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atFonts, $atProperties[1]
	Local Const $iURLFrameCreate = 8 ; Frame will be created if not found
	Local $oServiceManager, $oDesktop
	Local $bClose = False

	If Not IsString($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If Not IsObj($oDoc) Then
		$oServiceManager = __LO_ServiceManager()
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
		If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$atProperties[0] = __LO_SetPropertyValue("Hidden", True)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

		$oDoc = $oDesktop.loadComponentFromURL("private:factory/simpress", "_blank", $iURLFrameCreate, $atProperties)
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

		$bClose = True
	EndIf

	$atFonts = $oDoc.getCurrentController().getFrame().getContainerWindow().getFontDescriptors()
	If Not IsArray($atFonts) Then
		If $bClose Then $oDoc.Close(True)

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	For $i = 0 To UBound($atFonts) - 1
		If $atFonts[$i].Name = $sFontName Then
			If $bClose Then $oDoc.Close(True)

			Return SetError($__LO_STATUS_SUCCESS, 0, True)
		EndIf
		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	If $bClose Then $oDoc.Close(True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOImpress_FontExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_FontsGetNames
; Description ...: Retrieve an array of currently available font names.
; Syntax ........: _LOImpress_FontsGetNames([$oDoc = Null])
; Parameters ....: $oDoc                - [optional] an object. Default is Null. A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a "com.sun.star.frame.Desktop" Object.
;                  @Error 2 @Extended 3 Return 0 = Failed to create a Property Struct.
;                  @Error 2 @Extended 4 Return 0 = Failed to create a new Document.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Font list.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning a 4 Column Array, @Extended is set to the number of results. See remarks
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $oDoc is optional, if not called, an Impress Document is created invisibly to perform the check.
;                  Many fonts will be listed multiple times, this is because of the varying settings for them, such as bold, Italic, etc. Style Name is really a repeat of weight(Bold) and Slant (Italic) settings, but is included for easier processing if required.
;                  From personal tests, Slant only returns 0 or 2.
;                  The returned array will be as follows:
;                  The first column (Array[1][0]) contains the Font Name.
;                  The Second column (Array [1][1] contains the style name (Such as Bold Italic etc.)
;                  The third column (Array[1][2]) contains the Font weight (Bold) See Constants, $LOI_WEIGHT_* as defined in LibreOfficeImpress_Constants.au3;
;                  The fourth column (Array[1][3]) contains the font slant (Italic) See constants, $LOI_POSTURE_* as defined in LibreOfficeImpress_Constants.au3.
; Related .......: _LOImpress_FontExists
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_FontsGetNames($oDoc = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asFonts[0][4]
	Local $atFonts, $atProperties[1]
	Local Const $iURLFrameCreate = 8 ; Frame will be created if not found
	Local $oServiceManager, $oDesktop
	Local $bClose = False

	If Not IsObj($oDoc) Then
		$oServiceManager = __LO_ServiceManager()
		If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
		If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$atProperties[0] = __LO_SetPropertyValue("Hidden", True)
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

		$oDoc = $oDesktop.loadComponentFromURL("private:factory/simpress", "_blank", $iURLFrameCreate, $atProperties)
		If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)

		$bClose = True
	EndIf

	$atFonts = $oDoc.getCurrentController().getFrame().getContainerWindow().getFontDescriptors()
	If Not IsArray($atFonts) Then
		If $bClose Then $oDoc.Close(True)

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	ReDim $asFonts[UBound($atFonts)][4]

	For $i = 0 To UBound($atFonts) - 1
		$asFonts[$i][0] = $atFonts[$i].Name()
		$asFonts[$i][1] = $atFonts[$i].StyleName()
		$asFonts[$i][2] = $atFonts[$i].Weight
		$asFonts[$i][3] = $atFonts[$i].Slant() ; only 0 or 2?
		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	If $bClose Then $oDoc.Close(True)

	Return SetError($__LO_STATUS_SUCCESS, UBound($atFonts), $asFonts)
EndFunc   ;==>_LOImpress_FontsGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_GradientMulticolorAdd
; Description ...: Add a ColorStop to a Gradient ColorStop Array.
; Syntax ........: _LOImpress_GradientMulticolorAdd(ByRef $avColorStops, $iIndex, $nStopOffset, $iColor)
; Parameters ....: $avColorStops        - [in/out] an array of variants. A two column array of ColorStops. Array will be directly modified.
;                  $iIndex              - an integer value. The array index to insert the color stop. 0 Based. Call the last element index plus 1 to insert at the end.
;                  $nStopOffset         - a general number value (0-1.0). The ColorStop offset value.
;                  $iColor              - an integer value (0-16777215). The ColorStop color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $avColorStops not an Array.
;                  @Error 1 @Extended 2 Return 0 = $avColorStops does not contain two columns.
;                  @Error 1 @Extended 3 Return 0 = $iIndex not an Integer, less than 0 or greater than last element plus 1.
;                  @Error 1 @Extended 4 Return 0 = $nStopOffset not a number, less than 0 or greater than 1.0.
;                  @Error 1 @Extended 5 Return 0 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. ColorStop successfully added to array.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_GradientMulticolorAdd(ByRef $avColorStops, $iIndex, $nStopOffset, $iColor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__UBOUND_COLUMNS = 2

	If Not IsArray($avColorStops) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iIndex, 0, UBound($avColorStops)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_NumIsBetween($nStopOffset, 0, 1.0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	ReDim $avColorStops[UBound($avColorStops) + 1][2]

	For $iToWrite = (UBound($avColorStops) - 1) To 0 Step -1
		If $iToWrite = $iIndex Then
			$avColorStops[$iToWrite][0] = $nStopOffset
			$avColorStops[$iToWrite][1] = $iColor
			ExitLoop

		Else
			$avColorStops[$iToWrite][0] = $avColorStops[$iToWrite - 1][0]
			$avColorStops[$iToWrite][1] = $avColorStops[$iToWrite - 1][1]
		EndIf

		Sleep((IsInt($iToWrite / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_GradientMulticolorAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_GradientMulticolorDelete
; Description ...: Delete a ColorStop from a Gradient ColorStop Array.
; Syntax ........: _LOImpress_GradientMulticolorDelete(ByRef $avColorStops, $iIndex)
; Parameters ....: $avColorStops        - [in/out] an array of variants. A two column array of ColorStops. Array will be directly modified.
;                  $iIndex              - an integer value. The array index to delete. 0 Based.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $avColorStops not an Array.
;                  @Error 1 @Extended 2 Return 0 = $avColorStops does not contain two columns.
;                  @Error 1 @Extended 3 Return 0 = $iIndex not an Integer, less than 0 or greater than last element plus 1.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. ColorStop successfully removed from array.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_GradientMulticolorDelete(ByRef $avColorStops, $iIndex)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__UBOUND_COLUMNS = 2
	Local $iToRead = 0

	If Not IsArray($avColorStops) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iIndex, 0, UBound($avColorStops) - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	For $iToWrite = 0 To UBound($avColorStops) - 2
		If $iToWrite = $iIndex Then $iToRead += 1

		$avColorStops[$iToWrite][0] = $avColorStops[$iToWrite + $iToRead][0]
		$avColorStops[$iToWrite][1] = $avColorStops[$iToWrite + $iToRead][1]

		Sleep((IsInt($iToWrite / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	ReDim $avColorStops[UBound($avColorStops) - 1][2]

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_GradientMulticolorDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_GradientMulticolorModify
; Description ...: Modify a ColorStop in a Gradient ColorStop Array.
; Syntax ........: _LOImpress_GradientMulticolorModify(ByRef $avColorStops, $iIndex, $nStopOffset, $iColor)
; Parameters ....: $avColorStops        - [in/out] an array of variants. A two column array of ColorStops. Array will be directly modified.
;                  $iIndex              - an integer value. The array index to modify. 0 Based.
;                  $nStopOffset         - a general number value (0-1.0). The ColorStop offset value.
;                  $iColor              - an integer value (0-16777215). The ColorStop color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $avColorStops not an Array.
;                  @Error 1 @Extended 2 Return 0 = $avColorStops does not contain two columns.
;                  @Error 1 @Extended 3 Return 0 = $iIndex not an Integer, less than 0 or greater than last element.
;                  @Error 1 @Extended 4 Return 0 = $nStopOffset not a number, less than 0 or greater than 1.0.
;                  @Error 1 @Extended 5 Return 0 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. ColorStop successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_GradientMulticolorModify(ByRef $avColorStops, $iIndex, $nStopOffset, $iColor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__UBOUND_COLUMNS = 2

	If Not IsArray($avColorStops) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iIndex, 0, UBound($avColorStops) - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_NumIsBetween($nStopOffset, 0, 1.0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	For $iToWrite = 0 To UBound($avColorStops) - 1
		If $iToWrite = $iIndex Then
			$avColorStops[$iToWrite][0] = $nStopOffset
			$avColorStops[$iToWrite][1] = $iColor
			ExitLoop
		EndIf

		Sleep((IsInt($iToWrite / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_GradientMulticolorModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TransparencyGradientMultiAdd
; Description ...: Add a ColorStop to a Gradient ColorStop Array.
; Syntax ........: _LOImpress_TransparencyGradientMultiAdd(ByRef $avColorStops, $iIndex, $nStopOffset, $iTransparency)
; Parameters ....: $avColorStops        - [in/out] an array of variants. A two column array of ColorStops. Array will be directly modified.
;                  $iIndex              - an integer value. The array index to insert the color stop. 0 Based. Call the last element index plus 1 to insert at the end.
;                  $nStopOffset         - a general number value (0-1.0). The ColorStop offset value.
;                  $iTransparency       - an integer value (0-100). The ColorStop Transparency value percentage. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $avColorStops not an Array.
;                  @Error 1 @Extended 2 Return 0 = $avColorStops does not contain two columns.
;                  @Error 1 @Extended 3 Return 0 = $iIndex not an Integer, less than 0 or greater than last element plus 1.
;                  @Error 1 @Extended 4 Return 0 = $nStopOffset not a number, less than 0 or greater than 1.0.
;                  @Error 1 @Extended 5 Return 0 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. ColorStop successfully added to array.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TransparencyGradientMultiAdd(ByRef $avColorStops, $iIndex, $nStopOffset, $iTransparency)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__UBOUND_COLUMNS = 2

	If Not IsArray($avColorStops) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iIndex, 0, UBound($avColorStops)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_NumIsBetween($nStopOffset, 0, 1.0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	ReDim $avColorStops[UBound($avColorStops) + 1][2]

	For $iToWrite = (UBound($avColorStops) - 1) To 0 Step -1
		If $iToWrite = $iIndex Then
			$avColorStops[$iToWrite][0] = $nStopOffset
			$avColorStops[$iToWrite][1] = $iTransparency
			ExitLoop

		Else
			$avColorStops[$iToWrite][0] = $avColorStops[$iToWrite - 1][0]
			$avColorStops[$iToWrite][1] = $avColorStops[$iToWrite - 1][1]
		EndIf

		Sleep((IsInt($iToWrite / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_TransparencyGradientMultiAdd

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TransparencyGradientMultiDelete
; Description ...: Delete a ColorStop from a Gradient ColorStop Array.
; Syntax ........: _LOImpress_TransparencyGradientMultiDelete(ByRef $avColorStops, $iIndex)
; Parameters ....: $avColorStops        - [in/out] an array of variants. A two column array of ColorStops. Array will be directly modified.
;                  $iIndex              - an integer value. The array index to delete. 0 Based.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $avColorStops not an Array.
;                  @Error 1 @Extended 2 Return 0 = $avColorStops does not contain two columns.
;                  @Error 1 @Extended 3 Return 0 = $iIndex not an Integer, less than 0 or greater than last element plus 1.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. ColorStop successfully removed from array.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TransparencyGradientMultiDelete(ByRef $avColorStops, $iIndex)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__UBOUND_COLUMNS = 2
	Local $iToRead = 0

	If Not IsArray($avColorStops) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iIndex, 0, UBound($avColorStops) - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	For $iToWrite = 0 To UBound($avColorStops) - 2
		If $iToWrite = $iIndex Then $iToRead += 1

		$avColorStops[$iToWrite][0] = $avColorStops[$iToWrite + $iToRead][0]
		$avColorStops[$iToWrite][1] = $avColorStops[$iToWrite + $iToRead][1]

		Sleep((IsInt($iToWrite / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	ReDim $avColorStops[UBound($avColorStops) - 1][2]

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_TransparencyGradientMultiDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TransparencyGradientMultiModify
; Description ...: Modify a ColorStop in a Gradient ColorStop Array.
; Syntax ........: _LOImpress_TransparencyGradientMultiModify(ByRef $avColorStops, $iIndex, $nStopOffset, $iTransparency)
; Parameters ....: $avColorStops        - [in/out] an array of variants. A two column array of ColorStops. Array will be directly modified.
;                  $iIndex              - an integer value. The array index to modify. 0 Based.
;                  $nStopOffset         - a general number value (0-1.0). The ColorStop offset value.
;                  $iTransparency       - an integer value (0-100). The ColorStop Transparency value percentage. 0% is fully opaque and 100% is fully transparent.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $avColorStops not an Array.
;                  @Error 1 @Extended 2 Return 0 = $avColorStops does not contain two columns.
;                  @Error 1 @Extended 3 Return 0 = $iIndex not an Integer, less than 0 or greater than last element.
;                  @Error 1 @Extended 4 Return 0 = $nStopOffset not a number, less than 0 or greater than 1.0.
;                  @Error 1 @Extended 5 Return 0 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. ColorStop successfully modified.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TransparencyGradientMultiModify(ByRef $avColorStops, $iIndex, $nStopOffset, $iTransparency)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__UBOUND_COLUMNS = 2

	If Not IsArray($avColorStops) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iIndex, 0, UBound($avColorStops) - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_NumIsBetween($nStopOffset, 0, 1.0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	For $iToWrite = 0 To UBound($avColorStops) - 1
		If $iToWrite = $iIndex Then
			$avColorStops[$iToWrite][0] = $nStopOffset
			$avColorStops[$iToWrite][1] = $iTransparency
			ExitLoop
		EndIf

		Sleep((IsInt($iToWrite / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_TransparencyGradientMultiModify
