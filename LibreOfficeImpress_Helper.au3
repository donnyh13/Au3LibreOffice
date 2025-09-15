#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel
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
; _LOImpress_DrawShapeDelete
; _LOImpress_DrawShapeGetType
; _LOImpress_DrawShapeTextboxCreateTextCursor
; _LOImpress_FontExists
; _LOImpress_FontsGetNames
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_ComError_UserFunction
; Description ...: Set a UserFunction to receive the Fired COM Error Error outside of the UDF.
; Syntax ........: _LOImpress_ComError_UserFunction([$vUserFunction = Default[, $vParam1 = Null[, $vParam2 = Null[, $vParam3 = Null[, $vParam4 = Null[, $vParam5 = Null]]]]]])
; Parameters ....: $vUserFunction       - [optional] a Function or Keyword. Default value is Default. Accepts a Function, or the Keyword Default and Null. If set to a User function, the function may have up to 5 required parameters.
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
; Name ..........: _LOImpress_DrawShapeDelete
; Description ...: Delete a Shape.
; Syntax ........: _LOImpress_DrawShapeDelete(ByRef $oShape)
; Parameters ....: $oShape                - [in/out] an object. A Shape object returned by a previous _LOImpress_ShapeInsert, or _LOImpress_SlideShapesGetList function.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Shape's containing Slide.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve count of shapes.
;                  @Error 3 @Extended 3 Return 0 = Same number of shapes still present. Failed to delete the Shape.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Shape was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeDelete(ByRef $oShape)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iShapes
	Local $oDrawPage

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDrawPage = $oShape.Parent()
	If Not IsObj($oDrawPage) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iShapes = $oDrawPage.getCount()
	If Not IsInt($iShapes) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oDrawPage.remove($oShape)

	Return ($oDrawPage.getCount() = $iShapes) ? (SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_DrawShapeDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeGetType
; Description ...: Return the Drawing Shape's Type corresponding to the constants $LOI_DRAWSHAPE_TYPE_*
; Syntax ........: _LOImpress_DrawShapeGetType(ByRef $oShape)
; Parameters ....: $oShape              - [in/out] an object. A Shape object returned by a previous _LOImpress_ShapeInsert, or _LOImpress_SlideShapesGetList function.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oShape not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve CustomShapeGeometry Array.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve CustomShapeGeometry "Type" value.
;                  @Error 3 @Extended 3 Return 0 = Failed to determine CustomShape's type.
;                  @Error 3 @Extended 4 Return 0 = Failed to identify what type of "com.sun.star.drawing.EllipseShape" called shape is.
;                  @Error 3 @Extended 5 Return 0 = Called Shape is a unknown shape type.
;                  --Success--
;                  @Error 0 @Extended 1 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_CONNECTOR_* Type Shape. Returning $LOI_DRAWSHAPE_TYPE_CONNECTOR Constant Value. See Remarks #4.
;                  @Error 0 @Extended 2 Return Integer = Success. Shape is a Custom Shape Type. Returning appropriate Constant for shape type if successfully identified, else -1 if identification failed. See Remarks #1. See Constants, $LOI_DRAWSHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error 0 @Extended 3 Return Integer = Success. Shape is a*_BASIC_CIRCLE_SEGMENT or *_BASIC_ARC Type Shape. Returning appropriate Constant, See Constants, $LOI_DRAWSHAPE_TYPE_* as defined in LibreOfficeImpress_Constants.au3
;                  @Error 0 @Extended 4 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_LINE_CURVE Shape.
;                  @Error 0 @Extended 5 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_LINE_CURVE_FILLED Shape.
;                  @Error 0 @Extended 6 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE Shape.
;                  @Error 0 @Extended 7 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE_FILLED Shape.
;                  @Error 0 @Extended 8 Return Integer = Success. Shape is a *_LINE_LINE, *_LINE_LINE_45 or $LOI_DRAWSHAPE_TYPE_LINE_ARROW_* Type Shape. Returning $LOI_DRAWSHAPE_TYPE_LINE_LINE Constant Value. See Remarks #4.
;                  @Error 0 @Extended 9 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_LINE_DIMENSION Shape.
;                  @Error 0 @Extended 10 Return Integer = Success. Shape is a *_LINE_POLYGON, or *_LINE_POLYGON_45 Type Shape. Returning $LOI_DRAWSHAPE_TYPE_LINE_POLYGON Constant Value. See Remarks #2.
;                  @Error 0 @Extended 11 Return Integer = Success. Shape is a *_LINE_POLYGON_FILLED, or *_LINE_POLYGON_45_FILLED Type Shape. Returning $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED Constant Value. See Remarks #2.
;                  @Error 0 @Extended 11 Return Integer = Success. Shape is a $LOI_DRAWSHAPE_TYPE_3D_* Type Shape. Returning $LOI_DRAWSHAPE_TYPE_3D_CONE Constant Value. See Remarks #5.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: #1 Some shapes are not implemented, or not fully implemented into LibreOffice for automation, consequently they do not have appropriate type names as of yet. Many have simply ambiguous names, such as "non-primitive".
;                  Because of this the following Custom shape types cannot be identified, and this function will return -1:
;                  - $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT, known as "mso-spt100".
;                  - $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT, known as "non-primitive", should be "corner-right-arrow".
;                  - $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT, known as "non-primitive", should be "right-left-arrow".
;                  - $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_S_SHAPED, known as "non-primitive", should be "s-sharped-arrow".
;                  - $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_SPLIT, known as "non-primitive", should be "split-arrow".
;                  - $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_STRIPED_RIGHT, known as "mso-spt100", should be "striped-right-arrow".
;                  - $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT, known as "mso-spt89", should be "up-right-arrow-callout".
;                  - $LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN, known as "mso-spt100", should be "up-right-down-arrow".
;                  - $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_PIE, known as "mso-spt100", should be "circle-pie".
;                  - $LOI_DRAWSHAPE_TYPE_STARS_6_POINT, known as "non-primitive", should be "star6".
;                  - $LOI_DRAWSHAPE_TYPE_STARS_6_POINT_CONCAVE, known as "non-primitive", should be "concave-star6".
;                  - $LOI_DRAWSHAPE_TYPE_STARS_12_POINT, known as "non-primitive", should be "star12".
;                  - $LOI_DRAWSHAPE_TYPE_STARS_SIGNET, known as "non-primitive", should be "signet".
;                  - $LOI_DRAWSHAPE_TYPE_SYMBOL_CLOUD, known as "non-primitive", should be "cloud"?
;                  - $LOI_DRAWSHAPE_TYPE_SYMBOL_FLOWER, known as "non-primitive", should be "flower"?
;                  - $LOI_DRAWSHAPE_TYPE_SYMBOL_LIGHTNING, known as "non-primitive", should be "lightning".
;                  #2 The following Shapes implement the same type names, and are consequently indistinguishable:
;                  - $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE, $LOI_DRAWSHAPE_TYPE_BASIC_ELLIPSE (The Value of $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE is returned for either one.)
;                  - $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE, $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE (The Value of $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE is returned for either one.)
;                  - $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE_ROUNDED, $LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE_ROUNDED (The Value of $LOI_DRAWSHAPE_TYPE_BASIC_SQUARE_ROUNDED is returned for either one.)
;                  - $LOI_DRAWSHAPE_TYPE_LINE_POLYGON, $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_45 (The Value of $LOI_DRAWSHAPE_TYPE_LINE_POLYGON is returned for either of these.)
;                  - $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED, $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_45_FILLED (The Value of $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED is returned for either of these.)
;                  #3 The following Shapes have strange names that may change in the future, but currently are able to be identified:
;                  - $LOI_DRAWSHAPE_TYPE_STARS_DOORPLATE, known as, "mso-spt21", should be "doorplate"
;                  - $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_DIAMOND, known as, "col-502ad400", should be ??
;                  - $LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_OCTAGON, known as, "col-60da8460", should be ??
;                  #4 The following Shapes are customizable one to another, and are consequently indistinguishable:
;                  - $LOI_DRAWSHAPE_TYPE_CONNECTOR_* (The Value of $LOI_DRAWSHAPE_TYPE_CONNECTOR is returned for any of these.)
;                  - $LOI_DRAWSHAPE_TYPE_FONTWORK_* (The Value of $LOI_DRAWSHAPE_TYPE_FONTWORK_AIR_MAIL is returned for any of these.)
;                  - $LOI_DRAWSHAPE_TYPE_LINE_ARROW_* or $LOI_DRAWSHAPE_TYPE_LINE_LINE_45 (The Value of $LOI_DRAWSHAPE_TYPE_LINE_LINE is returned for any of these.)
;                  #5 The following Shapes are have nothing unique that I have found yet to identify each, and are consequently indistinguishable:
;                  - $LOI_DRAWSHAPE_TYPE_3D_* (The Value of $LOI_DRAWSHAPE_TYPE_3D_CONE is returned for any of these.)
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeGetType(ByRef $oShape)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atCusShapeGeo[0]
	Local Const $iCircleKind_CUT = 2 ; a circle with a cut connected by a line.
	Local Const $iCircleKind_ARC = 3 ; a circle with an open cut.
	Local $sType
	Local $iReturn

	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Switch $oShape.ShapeType()
		Case "com.sun.star.drawing.ConnectorShape" ; No way to differentiate between these??

			Return SetError($__LO_STATUS_SUCCESS, 1, $LOI_DRAWSHAPE_TYPE_CONNECTOR)

		Case "com.sun.star.drawing.CustomShape"
			$atCusShapeGeo = $oShape.CustomShapeGeometry()
			If Not IsArray($atCusShapeGeo) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			For $i = 0 To UBound($atCusShapeGeo) - 1
				If ($atCusShapeGeo[$i].Name() = "Type") Then
					$sType = $atCusShapeGeo[$i].Value()
					If Not IsString($sType) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

					ExitLoop
				EndIf

				Sleep((IsInt($i / $__LOICONST_SLEEP_DIV)) ? (10) : (0))
			Next

			$iReturn = __LOImpress_DrawShape_GetCustomType($sType)
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			Return SetError($__LO_STATUS_SUCCESS, 2, $iReturn)

		Case "com.sun.star.drawing.EllipseShape"
			If ($oShape.CircleKind() = $iCircleKind_CUT) Then ; Circle Segment = CircleKind_CUT(2), Arc = CircleKind_ARC(3)

				Return SetError($__LO_STATUS_SUCCESS, 3, $LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_SEGMENT)

			ElseIf ($oShape.CircleKind() = $iCircleKind_ARC) Then

				Return SetError($__LO_STATUS_SUCCESS, 3, $LOI_DRAWSHAPE_TYPE_BASIC_ARC)

			Else

				Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
			EndIf

		Case "com.sun.star.drawing.OpenBezierShape"

			Return SetError($__LO_STATUS_SUCCESS, 4, $LOI_DRAWSHAPE_TYPE_LINE_CURVE)

		Case "com.sun.star.drawing.ClosedBezierShape"

			Return SetError($__LO_STATUS_SUCCESS, 5, $LOI_DRAWSHAPE_TYPE_LINE_CURVE_FILLED)

		Case "com.sun.star.drawing.OpenFreeHandShape"

			Return SetError($__LO_STATUS_SUCCESS, 6, $LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE)

		Case "com.sun.star.drawing.ClosedFreeHandShape"

			Return SetError($__LO_STATUS_SUCCESS, 7, $LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE_FILLED)

		Case "com.sun.star.drawing.LineShape" ; No way to differentiate between these?? (Lines + Arrows)

			Return SetError($__LO_STATUS_SUCCESS, 8, $LOI_DRAWSHAPE_TYPE_LINE_LINE)

		Case "com.sun.star.drawing.MeasureShape"

			Return SetError($__LO_STATUS_SUCCESS, 9, $LOI_DRAWSHAPE_TYPE_LINE_DIMENSION)

		Case "com.sun.star.drawing.PolyLineShape"
			;~ $LOI_DRAWSHAPE_TYPE_LINE_POLYGON ; No way to differentiate between these??
			;~ $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_45

			Return SetError($__LO_STATUS_SUCCESS, 10, $LOI_DRAWSHAPE_TYPE_LINE_POLYGON)

		Case "com.sun.star.drawing.PolyPolygonShape"

			Return SetError($__LO_STATUS_SUCCESS, 11, $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED)
			;~ $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED ; No way to differentiate between these??
			;~ $LOI_DRAWSHAPE_TYPE_LINE_POLYGON_45_FILLED

		Case "com.sun.star.drawing.Shape3DSceneObject" ; No way to differentiate between these??

			Return SetError($__LO_STATUS_SUCCESS, 12, $LOI_DRAWSHAPE_TYPE_3D_CONE)

		Case Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0) ; Unknown shape type.
	EndSwitch
EndFunc   ;==>_LOImpress_DrawShapeGetType

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_DrawShapeTextboxCreateTextCursor
; Description ...: Create a Text Cursor in a Textbox for inserting text etc.
; Syntax ........: _LOImpress_DrawShapeTextboxCreateTextCursor(ByRef $oTextbox)
; Parameters ....: $oTextbox              - [in/out] an object. A Textbox Shape object returned by a previous _LOImpress_ShapeInsert, or _LOImpress_SlideShapesGetList function.
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTextbox not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. A Text Cursor Object located in the Frame.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_DrawShapeTextboxCreateTextCursor(ByRef $oTextbox)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextCursor
	Local $iShapeType

	If Not IsObj($oTextbox) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iShapeType = __LOImpress_DrawShapeGetType($oTextbox)
	If ($iShapeType <> $LOI_SHAPE_TYPE_TEXTBOX) And ($iShapeType <> $LOI_SHAPE_TYPE_TEXTBOX_TITLE) And ($iShapeType <> $LOI_SHAPE_TYPE_TEXTBOX_SUBTITLE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oTextCursor = $oTextbox.Text.createTextCursor()
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextCursor)
EndFunc   ;==>_LOImpress_DrawShapeTextboxCreateTextCursor

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
;                  @Error 0 @Extended 0 Return Boolean = Success. Returns True if the Font is available, else False.
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
;                  @Error 0 @Extended ? Return Array = Success. Returns a 4 Column Array, @extended is set to the number of results. See remarks
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
