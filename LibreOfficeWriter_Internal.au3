#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Writer
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Helper.au3"

; Required AutoIt Include
#include <WinAPIGdiDC.au3>

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Various functions for internal data processing, data retrieval, retrieving and applying settings for LibreOffice UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __LOWriter_AnyAreDefault
; __LOWriter_Border
; __LOWriter_CharBorder
; __LOWriter_CharBorderPadding
; __LOWriter_CharEffect
; __LOWriter_CharFont
; __LOWriter_CharFontColor
; __LOWriter_CharOverLine
; __LOWriter_CharPosition
; __LOWriter_CharRotateScale
; __LOWriter_CharShadow
; __LOWriter_CharSpacing
; __LOWriter_CharStrikeOut
; __LOWriter_CharStyleNameToggle
; __LOWriter_CharUnderLine
; __LOWriter_ColorRemoveAlpha
; __LOWriter_CreatePoint
; __LOWriter_CursorGetText
; __LOWriter_DateStructCompare
; __LOWriter_DirFrmtCheck
; __LOWriter_FieldCountType
; __LOWriter_FieldsGetList
; __LOWriter_FieldTypeServices
; __LOWriter_FilterNameGet
; __LOWriter_FindFormatAddSetting
; __LOWriter_FindFormatDeleteSetting
; __LOWriter_FindFormatRetrieveSetting
; __LOWriter_FooterBorder
; __LOWriter_FormConGetObj
; __LOWriter_FormConIdentify
; __LOWriter_FormConSetGetFontDesc
; __LOWriter_GetPrinterSetting
; __LOWriter_GetShapeName
; __LOWriter_GradientNameInsert
; __LOWriter_GradientPresets
; __LOWriter_HeaderBorder
; __LOWriter_ImageGetSuggestedSize
; __LOWriter_Internal_CursorGetDataType
; __LOWriter_Internal_CursorGetType
; __LOWriter_InternalComErrorHandler
; __LOWriter_IsCellRange
; __LOWriter_IsTableInDoc
; __LOWriter_NumRuleCreateMap
; __LOWriter_NumStyleCreateScript
; __LOWriter_NumStyleDeleteScript
; __LOWriter_NumStyleInitiateDocument
; __LOWriter_NumStyleListFormat
; __LOWriter_NumStyleModify
; __LOWriter_ObjRelativeSize
; __LOWriter_PageStyleNameToggle
; __LOWriter_ParAlignment
; __LOWriter_ParBackColor
; __LOWriter_ParBorderPadding
; __LOWriter_ParDropCaps
; __LOWriter_ParHasTabStop
; __LOWriter_ParHyphenation
; __LOWriter_ParIndent
; __LOWriter_ParOutLineAndList
; __LOWriter_ParPageBreak
; __LOWriter_ParShadow
; __LOWriter_ParSpace
; __LOWriter_ParStyleNameToggle
; __LOWriter_ParTabStopCreate
; __LOWriter_ParTabStopDelete
; __LOWriter_ParTabStopMod
; __LOWriter_ParTabStopsGetList
; __LOWriter_ParTxtFlowOpt
; __LOWriter_Shape_CreateArrow
; __LOWriter_Shape_CreateBasic
; __LOWriter_Shape_CreateCallout
; __LOWriter_Shape_CreateFlowchart
; __LOWriter_Shape_CreateLine
; __LOWriter_Shape_CreateStars
; __LOWriter_Shape_CreateSymbol
; __LOWriter_Shape_GetCustomType
; __LOWriter_ShapeArrowStyleName
; __LOWriter_ShapeLineStyleName
; __LOWriter_ShapePointGetSettings
; __LOWriter_ShapePointModify
; __LOWriter_TableBorder
; __LOWriter_TableCursorMove
; __LOWriter_TableHasCellName
; __LOWriter_TableHasColumnRange
; __LOWriter_TableHasRowRange
; __LOWriter_TableRowSplitToggle
; __LOWriter_TextCursorMove
; __LOWriter_TransparencyGradientConvert
; __LOWriter_TransparencyGradientNameInsert
; __LOWriter_ViewCursorMove
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_AnyAreDefault
; Description ...: Tests whether any input parameters are equal to Default keyword.
; Syntax ........: __LOWriter_AnyAreDefault($vVar1[, $vVar2 = Null[, $vVar3 = Null[, $vVar4 = Null[, $vVar5 = Null[, $vVar6 = Null[, $vVar7 = Null[, $vVar8 = Null]]]]]]])
; Parameters ....: $vVar1               - a variant value.
;                  $vVar2               - [optional] a variant value. Default is Null.
;                  $vVar3               - [optional] a variant value. Default is Null.
;                  $vVar4               - [optional] a variant value. Default is Null.
;                  $vVar5               - [optional] a variant value. Default is Null.
;                  $vVar6               - [optional] a variant value. Default is Null.
;                  $vVar7               - [optional] a variant value. Default is Null.
;                  $vVar8               - [optional] a variant value. Default is Null.
; Return values .: Success: Boolean
;                  Failure: False
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = If Any parameters are equal to Default, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_AnyAreDefault($vVar1, $vVar2 = Null, $vVar3 = Null, $vVar4 = Null, $vVar5 = Null, $vVar6 = Null, $vVar7 = Null, $vVar8 = Null)
	Local $bAnyDefault1, $bAnyDefault2
	$bAnyDefault1 = (($vVar1 = Default) Or ($vVar2 = Default) Or ($vVar3 = Default) Or ($vVar4 = Default)) ? (True) : (False)
	$bAnyDefault2 = (($vVar5 = Default) Or ($vVar6 = Default) Or ($vVar7 = Default) Or ($vVar8 = Default)) ? (True) : (False)

	Return ($bAnyDefault1 Or $bAnyDefault2) ? (SetError($__LO_STATUS_SUCCESS, 0, True)) : (SetError($__LO_STATUS_SUCCESS, 0, False))
EndFunc   ;==>__LOWriter_AnyAreDefault

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Border
; Description ...: Border Setting Internal function. Libre Office Version 3.4 and Up.
; Syntax ........: __LOWriter_Border(ByRef $oObj, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
; Parameters ....: $oObj                - [in/out] an object. An Object that implements BorderLine2 service for border properties.
;                  $bWid                - a boolean value. If True, the calling function is for setting Border Line Width.
;                  $bSty                - a boolean value. If True, the calling function is for setting Border Line Style.
;                  $bCol                - a boolean value. If True, the calling function is for setting Border Line Color.
;                  $iTop                - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iBottom             - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iLeft               - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iRight              - an integer value. See Border Style, Width, and Color functions for possible values.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style/Color when Top Border width not set.
;                  @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border style/Color when Bottom Border width not set.
;                  @Error 4 @Extended 3 Return 0 = Cannot set Left Border style/Color when Left Border width not set.
;                  @Error 4 @Extended 4 Return 0 = Cannot set Right Border style/Color when Right Border width not set.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with all other parameters set to Null keyword, and $bWid, or $bSty, or $bCol set to true to get the corresponding current settings.
;                  All distance values are set in Micrometers. Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_Border(ByRef $oObj, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avBorder[4]
	Local $tBL2

	If Not __LO_VersionCheck(3.4) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (($bWid + $bSty + $bCol) <> 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; If more than one Boolean is true = error

	If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then
		If $bWid Then
			__LO_ArrayFill($avBorder, $oObj.TopBorder.LineWidth(), $oObj.BottomBorder.LineWidth(), $oObj.LeftBorder.LineWidth(), _
					$oObj.RightBorder.LineWidth())

		ElseIf $bSty Then
			__LO_ArrayFill($avBorder, $oObj.TopBorder.LineStyle(), $oObj.BottomBorder.LineStyle(), $oObj.LeftBorder.LineStyle(), _
					$oObj.RightBorder.LineStyle())

		ElseIf $bCol Then
			__LO_ArrayFill($avBorder, $oObj.TopBorder.Color(), $oObj.BottomBorder.Color(), $oObj.LeftBorder.Color(), $oObj.RightBorder.Color())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avBorder)
	EndIf

	$tBL2 = __LO_CreateStruct("com.sun.star.table.BorderLine2")
	If Not IsObj($tBL2) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If $iTop <> Null Then
		If Not $bWid And ($oObj.TopBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0) ; If Width not set, cant set color or style.

		; Top Line
		$tBL2.LineWidth = ($bWid) ? ($iTop) : ($oObj.TopBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iTop) : ($oObj.TopBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iTop) : ($oObj.TopBorder.Color()) ; copy Color over to new size structure
		$oObj.TopBorder = $tBL2
	EndIf

	If $iBottom <> Null Then
		If Not $bWid And ($oObj.BottomBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 2, 0) ; If Width not set, cant set color or style.

		; Bottom Line
		$tBL2.LineWidth = ($bWid) ? ($iBottom) : ($oObj.BottomBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iBottom) : ($oObj.BottomBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iBottom) : ($oObj.BottomBorder.Color()) ; copy Color over to new size structure
		$oObj.BottomBorder = $tBL2
	EndIf

	If $iLeft <> Null Then
		If Not $bWid And ($oObj.LeftBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 3, 0) ; If Width not set, cant set color or style.

		; Left Line
		$tBL2.LineWidth = ($bWid) ? ($iLeft) : ($oObj.LeftBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iLeft) : ($oObj.LeftBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iLeft) : ($oObj.LeftBorder.Color()) ; copy Color over to new size structure
		$oObj.LeftBorder = $tBL2
	EndIf

	If $iRight <> Null Then
		If Not $bWid And ($oObj.RightBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 4, 0) ; If Width not set, cant set color or style.

		; Right Line
		$tBL2.LineWidth = ($bWid) ? ($iRight) : ($oObj.RightBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iRight) : ($oObj.RightBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iRight) : ($oObj.RightBorder.Color()) ; copy Color over to new size structure
		$oObj.RightBorder = $tBL2
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_Border

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharBorder
; Description ...: Character Border Setting and retrieving Internal function.
; Syntax ........: __LOWriter_CharBorder(ByRef $oObj, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $bWid                - a boolean value. If True, the calling function is for setting Border Line Width.
;                  $bSty                - a boolean value. If True, the calling function is for setting Border Line Style.
;                  $bCol                - a boolean value. If True, the calling function is for setting Border Line Color.
;                  $iTop                - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iBottom             - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iLeft               - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iRight              - an integer value. See Border Style, Width, and Color functions for possible values.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj Variable not Object type variable.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style/Color when Top Border width not set.
;                  @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border style/Color when Bottom Border width not set.
;                  @Error 4 @Extended 3 Return 0 = Cannot set Left Border style/Color when Left Border width not set.
;                  @Error 4 @Extended 4 Return 0 = Cannot set Right Border style/Color when Right Border width not set.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, and $bWid, or $bSty, or $bCol set to true to get the corresponding current settings.
;                  All distance values are set in Micrometers.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharBorder(ByRef $oObj, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avBorder[4]
	Local $tBL2

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (($bWid + $bSty + $bCol) <> 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; If more than one Boolean is true = error

	If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then
		If $bWid Then
			__LO_ArrayFill($avBorder, $oObj.CharTopBorder.LineWidth(), $oObj.CharBottomBorder.LineWidth(), $oObj.CharLeftBorder.LineWidth(), _
					$oObj.CharRightBorder.LineWidth())

		ElseIf $bSty Then
			__LO_ArrayFill($avBorder, $oObj.CharTopBorder.LineStyle(), $oObj.CharBottomBorder.LineStyle(), $oObj.CharLeftBorder.LineStyle(), _
					$oObj.CharRightBorder.LineStyle())

		ElseIf $bCol Then
			__LO_ArrayFill($avBorder, $oObj.CharTopBorder.Color(), $oObj.CharBottomBorder.Color(), $oObj.CharLeftBorder.Color(), _
					$oObj.CharRightBorder.Color())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avBorder)
	EndIf

	$tBL2 = __LO_CreateStruct("com.sun.star.table.BorderLine2")
	If Not IsObj($tBL2) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If $iTop <> Null Then
		If Not $bWid And ($oObj.CharTopBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0) ; If Width not set, cant set color or style.

		; Top Line
		$tBL2.LineWidth = ($bWid) ? ($iTop) : ($oObj.CharTopBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iTop) : ($oObj.CharTopBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iTop) : ($oObj.CharTopBorder.Color()) ; copy Color over to new size structure
		$oObj.CharTopBorder = $tBL2
	EndIf

	If $iBottom <> Null Then
		If Not $bWid And ($oObj.CharBottomBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 2, 0) ; If Width not set, cant set color or style.

		; Bottom Line
		$tBL2.LineWidth = ($bWid) ? ($iBottom) : ($oObj.CharBottomBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iBottom) : ($oObj.CharBottomBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iBottom) : ($oObj.CharBottomBorder.Color()) ; copy Color over to new size structure
		$oObj.CharBottomBorder = $tBL2
	EndIf

	If $iLeft <> Null Then
		If Not $bWid And ($oObj.CharLeftBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 3, 0) ; If Width not set, cant set color or style.

		; Left Line
		$tBL2.LineWidth = ($bWid) ? ($iLeft) : ($oObj.CharLeftBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iLeft) : ($oObj.CharLeftBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iLeft) : ($oObj.CharLeftBorder.Color()) ; copy Color over to new size structure
		$oObj.CharLeftBorder = $tBL2
	EndIf

	If $iRight <> Null Then
		If Not $bWid And ($oObj.CharRightBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 4, 0) ; If Width not set, cant set color or style.

		; Right Line
		$tBL2.LineWidth = ($bWid) ? ($iRight) : ($oObj.CharRightBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iRight) : ($oObj.CharRightBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iRight) : ($oObj.CharRightBorder.Color()) ; copy Color over to new size structure
		$oObj.CharRightBorder = $tBL2
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_CharBorder

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharBorderPadding
; Description ...: Set and retrieve the distance between the border and the characters.
; Syntax ........: __LOWriter_CharBorderPadding(ByRef $oObj, $iAll, $iTop, $iBottom, $iLeft, $iRight)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $iAll                - an integer value. Set all four padding values to the same value. When used, all other parameters are ignored. In Micrometers.
;                  $iTop                - an integer value. Set the Top border distance in Micrometers.
;                  $iBottom             - an integer value. Set the Bottom border distance in Micrometers.
;                  $iLeft               - an integer value. Set the left border distance in Micrometers.
;                  $iRight              - an integer value. Set the Right border distance in Micrometers.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iAll not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iTop not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iBottom not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $Left not an Integer.
;                  @Error 1 @Extended 8 Return 0 = $iRight not an Integer.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iAll border distance
;                  |                               2 = Error setting $iTop border distance
;                  |                               4 = Error setting $iBottom border distance
;                  |                               8 = Error setting $iLeft border distance
;                  |                               16 = Error setting $iRight border distance
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  All distance values are set in Micrometers. Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharBorderPadding(ByRef $oObj, $iAll, $iTop, $iBottom, $iLeft, $iRight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiBPadding[5]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($iAll, $iTop, $iBottom, $iLeft, $iRight) Then
		__LO_ArrayFill($aiBPadding, $oObj.CharBorderDistance(), $oObj.CharTopBorderDistance(), $oObj.CharBottomBorderDistance(), _
				$oObj.CharLeftBorderDistance(), $oObj.CharRightBorderDistance())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiBPadding)
	EndIf

	If ($iAll <> Null) Then
		If Not __LO_IntIsBetween($iAll, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.CharBorderDistance = $iAll
		$iError = (__LO_IntIsBetween($oObj.CharBorderDistance(), $iAll - 1, $iAll + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iTop <> Null) Then
		If Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.CharTopBorderDistance = $iTop
		$iError = (__LO_IntIsBetween($oObj.CharTopBorderDistance(), $iTop - 1, $iTop + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iBottom <> Null) Then
		If Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.CharBottomBorderDistance = $iBottom
		$iError = (__LO_IntIsBetween($oObj.CharBottomBorderDistance(), $iBottom - 1, $iBottom + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iLeft <> Null) Then
		If Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oObj.CharLeftBorderDistance = $iLeft
		$iError = (__LO_IntIsBetween($oObj.CharLeftBorderDistance(), $iLeft - 1, $iLeft + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iRight <> Null) Then
		If Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oObj.CharRightBorderDistance = $iRight
		$iError = (__LO_IntIsBetween($oObj.CharRightBorderDistance(), $iRight - 1, $iRight + 1)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_CharBorderPadding

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharEffect
; Description ...: Set or Retrieve the Font Effect settings.
; Syntax ........: __LOWriter_CharEffect(ByRef $oObj, $iRelief, $iCase, $bHidden, $bOutline, $bShadow)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $iRelief             - an integer value (0-2). The Character Relief style. See Constants, $LOW_RELIEF_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iCase               - an integer value (0-4). The Character Case Style. See Constants, $LOW_CASEMAP_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bHidden             - a boolean value. If True, the Characters are hidden.
;                  $bOutline            - a boolean value. If True, the characters have an outline around the outside.
;                  $bShadow             - a boolean value. If True, the characters have a shadow.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iRelief not an integer or less than 0, or greater than 2. See Constants, $LOW_RELIEF_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iCase not an integer or less than 0, or greater than 4. See Constants, $LOW_CASEMAP_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bHidden not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bOutline not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bShadow not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iRelief
;                  |                               2 = Error setting $iCase
;                  |                               4 = Error setting $bHidden
;                  |                               8 = Error setting $bOutline
;                  |                               16 = Error setting $bShadow
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharEffect(ByRef $oObj, $iRelief, $iCase, $bHidden, $bOutline, $bShadow)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avEffect[5]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($iRelief, $iCase, $bHidden, $bOutline, $bShadow) Then
		__LO_ArrayFill($avEffect, $oObj.CharRelief(), $oObj.CharCaseMap(), $oObj.CharHidden(), $oObj.CharContoured(), $oObj.CharShadowed())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avEffect)
	EndIf

	If ($iRelief <> Null) Then
		If Not __LO_IntIsBetween($iRelief, $LOW_RELIEF_NONE, $LOW_RELIEF_ENGRAVED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.CharRelief = $iRelief
		$iError = ($oObj.CharRelief() = $iRelief) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iCase <> Null) Then
		If Not __LO_IntIsBetween($iCase, $LOW_CASEMAP_NONE, $LOW_CASEMAP_SM_CAPS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.CharCaseMap = $iCase
		$iError = ($oObj.CharCaseMap() = $iCase) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bHidden <> Null) Then
		If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.CharHidden = $bHidden
		$iError = ($oObj.CharHidden() = $bHidden) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bOutline <> Null) Then
		If Not IsBool($bOutline) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oObj.CharContoured = $bOutline
		$iError = ($oObj.CharContoured() = $bOutline) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bShadow <> Null) Then
		If Not IsBool($bShadow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oObj.CharShadowed = $bShadow
		$iError = ($oObj.CharShadowed() = $bShadow) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_CharEffect

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharFont
; Description ...: Set and Retrieve the Font Settings
; Syntax ........: __LOWriter_CharFont(ByRef $oObj, $sFontName, $nFontSize, $iPosture, $iWeight)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $sFontName           - a string value. The Font Name to change to.
;                  $nFontSize           - a general number value. The new Font size.
;                  $iPosture            - an integer value (0-5). Italic setting. See Constants, $LOW_POSTURE_* as defined in LibreOfficeWriter_Constants.au3. Also see remarks.
;                  $iWeight             - an integer value (0,50-200). Bold settings see Constants, $LOW_WEIGHT_* as defined in LibreOfficeWriter_Constants.au3. Also see remarks.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 4 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 5 Return 0 = $sFontName not a String.
;                  @Error 1 @Extended 6 Return 0 = $nFontSize not a Number.
;                  @Error 1 @Extended 7 Return 0 = $iPosture not an Integer, less than 0, or greater than 5. See Constants, $LOW_POSTURE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $iWeight less than 50 and not 0, or more than 200. See Constants, $LOW_WEIGHT_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sFontName
;                  |                               2 = Error setting $nFontSize
;                  |                               4 = Error setting $iPosture
;                  |                               8 = Error setting $iWeight
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Not every font accepts Bold and Italic settings, and not all settings for bold and Italic are accepted, such as oblique, ultra Bold etc.
;                  Libre Writer accepts only the predefined weight values, any other values are changed automatically to an acceptable value, which could trigger a settings error.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharFont(ByRef $oObj, $sFontName, $nFontSize, $iPosture, $iWeight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFont[4]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If __LO_VarsAreNull($sFontName, $nFontSize, $iPosture, $iWeight) Then
		__LO_ArrayFill($avFont, $oObj.CharFontName(), $oObj.CharHeight(), $oObj.CharPosture(), $oObj.CharWeight())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avFont)
	EndIf

	If ($sFontName <> Null) Then
		If Not IsString($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.CharFontName = $sFontName
		$iError = ($oObj.CharFontName() = $sFontName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($nFontSize <> Null) Then
		If Not IsNumber($nFontSize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.CharHeight = $nFontSize
		$iError = ($oObj.CharHeight() = $nFontSize) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iPosture <> Null) Then
		If Not __LO_IntIsBetween($iPosture, $LOW_POSTURE_NONE, $LOW_POSTURE_ITALIC) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oObj.CharPosture = $iPosture
		$iError = ($oObj.CharPosture() = $iPosture) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iWeight <> Null) Then
		If Not __LO_IntIsBetween($iWeight, $LOW_WEIGHT_THIN, $LOW_WEIGHT_BLACK, "", $LOW_WEIGHT_DONT_KNOW) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oObj.CharWeight = $iWeight
		$iError = ($oObj.CharWeight() = $iWeight) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_CharFont

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharFontColor
; Description ...: Set or retrieve the font color, transparency and highlighting values.
; Syntax ........: __LOWriter_CharFontColor(ByRef $oObj, $iFontColor, $iTransparency, $iHighlight)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $iFontColor          - an integer value (-1-16777215). The desired font Color value in Long Integer format, can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Set to $LO_COLOR_OFF(-1) for Auto color.
;                  $iTransparency       - an integer value (0-100). Transparency percentage. 0 is visible, 100 is invisible. Available for Libre Office 7.0 and up.
;                  $iHighlight          - an integer value (-1-16777215). The highlight Color value in Long Integer format, can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Set to $LO_COLOR_OFF(-1) for No color.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iFontColor not an integer, less than -1, or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $iTransparency not an Integer, or less than 0, or greater than 100%.
;                  @Error 1 @Extended 6 Return 0 = $iHighlight not an integer, less than -1, or greater than 16777215.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve old Transparency value.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $FontColor
;                  |                               2 = Error setting $iTransparency.
;                  |                               4 = Error setting $iHighlight
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 7.0.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 or 3 Element Array with values in order of function parameters. If The current Libre Office version is below 7.0 the returned array will contain 2 elements, because $iTransparency is not available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharFontColor(ByRef $oObj, $iFontColor, $iTransparency, $iHighlight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iOldTransparency
	Local $avColor[2]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($iFontColor, $iTransparency, $iHighlight) Then
		If __LO_VersionCheck(7.0) Then
			__LO_ArrayFill($avColor, __LOWriter_ColorRemoveAlpha($oObj.CharColor()), $oObj.CharTransparence(), $oObj.CharBackColor())

		Else
			__LO_ArrayFill($avColor, __LOWriter_ColorRemoveAlpha($oObj.CharColor()), $oObj.CharBackColor())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avColor)
	EndIf

	If ($iFontColor <> Null) Then
		If Not __LO_IntIsBetween($iFontColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		If __LO_VersionCheck(7.0) Then
			$iOldTransparency = $oObj.CharTransparence()
			If Not IsInt($iOldTransparency) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
		EndIf

		$oObj.CharColor = $iFontColor
		$iError = ($oObj.CharColor() = $iFontColor) ? ($iError) : (BitOR($iError, 1))

		If IsInt($iOldTransparency) Then $oObj.CharTransparence = $iOldTransparency
	EndIf

	If ($iTransparency <> Null) Then
		If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If Not __LO_VersionCheck(7.0) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oObj.CharTransparence = $iTransparency
		$iError = ($oObj.CharTransparence() = $iTransparency) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iHighlight <> Null) Then
		If Not __LO_IntIsBetween($iHighlight, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		; CharHighlight; same as CharBackColor---Libre seems to use back color for highlighting however, so using that for setting.
		;~ 		If Not __LO_VersionCheck(4.2) Then Return SetError($__LO_STATUS_VER_ERROR, 2, 0)
		;~ 		$oObj.CharHighlight = $iHighlight ;-- keeping old method in case.
		;~ 		$iError = ($oObj.CharHighlight() = $iHighlight) ? ($iError) : (BitOR($iError, 4)
		$oObj.CharBackColor = $iHighlight
		$iError = ($oObj.CharBackColor() = $iHighlight) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_CharFontColor

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharOverLine
; Description ...: Set and retrieve the OverLine settings.
; Syntax ........: __LOWriter_CharOverLine(ByRef $oObj, $bWordOnly, $iOverLineStyle, $bOLHasColor, $iOLColor)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $bWordOnly           - a boolean value. If true, white spaces are not Overlined.
;                  $iOverLineStyle      - an integer value (0-18). The line style of the Overline, see constants, $LOW_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $bOLHasColor         - a boolean value. If True, the Overline is colored, must be set to true in order to set the Overline color.
;                  $iOLColor            - an integer value (-1-16777215). The color of the Overline, set in Long integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Set to $LO_COLOR_OFF(-1) for automatic color mode.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iOverLineStyle not an Integer, or less than 0, or greater than 18. See constants, $LOW_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bOLHasColor not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $iOLColor not an Integer, or less than -1, or greater than 16777215.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bWordOnly
;                  |                               2 = Error setting $iOverLineStyle
;                  |                               4 = Error setting $OLHasColor
;                  |                               8 = Error setting $iOLColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: OverLine line style uses the same constants as underline style.
;                  Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $bOLHasColor must be set to true in order to set the Overline color.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharOverLine(ByRef $oObj, $bWordOnly, $iOverLineStyle, $bOLHasColor, $iOLColor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avOverLine[4]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($bWordOnly, $iOverLineStyle, $bOLHasColor, $iOLColor) Then
		__LO_ArrayFill($avOverLine, $oObj.CharWordMode(), $oObj.CharOverline(), $oObj.CharOverlineHasColor(), $oObj.CharOverlineColor())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avOverLine)
	EndIf

	If ($bWordOnly <> Null) Then
		If Not IsBool($bWordOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.CharWordMode = $bWordOnly
		$iError = ($oObj.CharWordMode() = $bWordOnly) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iOverLineStyle <> Null) Then
		If Not __LO_IntIsBetween($iOverLineStyle, $LOW_UNDERLINE_NONE, $LOW_UNDERLINE_BOLD_WAVE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.CharOverline = $iOverLineStyle
		$iError = ($oObj.CharOverline() = $iOverLineStyle) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bOLHasColor <> Null) Then
		If Not IsBool($bOLHasColor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.CharOverlineHasColor = $bOLHasColor
		$iError = ($oObj.CharOverlineHasColor() = $bOLHasColor) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iOLColor <> Null) Then
		If Not __LO_IntIsBetween($iOLColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oObj.CharOverlineColor = $iOLColor
		$iError = ($oObj.CharOverlineColor() = $iOLColor) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_CharOverLine

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharPosition
; Description ...: Set and retrieve settings related to Sub/Super Script and relative size.
; Syntax ........: __LOWriter_CharPosition(ByRef $oObj, $bAutoSuper, $iSuperScript, $bAutoSub, $iSubScript, $iRelativeSize)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $bAutoSuper          - a boolean value. If True, automatic sizing for Superscript is active.
;                  $iSuperScript        - an integer value. The Superscript percentage value. See Remarks.
;                  $bAutoSub            - a boolean value. If True, automatic sizing for Subscript is active.
;                  $iSubScript          - an integer value. The Subscript percentage value. See Remarks.
;                  $iRelativeSize       - an integer value (1-100). The size percentage relative to current font size.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bAutoSuper not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bAutoSub not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iSuperScript not an integer, less than 0, higher than 100 and Not 14000.
;                  @Error 1 @Extended 7 Return 0 = $iSubScript not an integer, less than -100, higher than 100 and Not 14000.
;                  @Error 1 @Extended 8 Return 0 = $iRelativeSize not an integer, or less than 1, higher than 100.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iSuperScript
;                  |                               2 = Error setting $iSubScript
;                  |                               4 = Error setting $iRelativeSize.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Set either $iSubScript or $iSuperScript to 0 to return it to Normal setting.
;                  The way LibreOffice is set up Super/Subscript are set in the same setting, Super is a positive number from 1 to 100 (percentage), Subscript is a negative number set to 1 to 100 percentage.
;                  For the user's convenience this function accepts both positive and negative numbers for Subscript, if a positive number is called for Subscript, it is automatically set to a negative.
;                  Automatic Superscript has a integer value of 14000, Auto Subscript has a integer value of -14000. There is no settable setting of Automatic Super/Sub Script, though one exists, it is read-only in LibreOffice, consequently I have made two separate parameters to be able to determine if the user wants to automatically set Superscript or Subscript.
;                  If you set both Auto Superscript to True and Auto Subscript to True, or $iSuperScript to an integer and $iSubScript to an integer, Subscript will be set as it is the last in the line to be set in this function, and thus will over-write any Superscript settings.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharPosition(ByRef $oObj, $bAutoSuper, $iSuperScript, $bAutoSub, $iSubScript, $iRelativeSize)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avPosition[5]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($bAutoSuper, $iSuperScript, $bAutoSub, $iSubScript, $iRelativeSize) Then
		__LO_ArrayFill($avPosition, ($oObj.CharEscapement() = 14000) ? (True) : (False), ($oObj.CharEscapement() > 0) ? ($oObj.CharEscapement()) : (0), _
				($oObj.CharEscapement() = -14000) ? (True) : (False), ($oObj.CharEscapement() < 0) ? ($oObj.CharEscapement()) : (0), $oObj.CharEscapementHeight())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPosition)
	EndIf

	If ($bAutoSuper <> Null) Then
		If Not IsBool($bAutoSuper) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		; If $bAutoSuper = True set it to 14000 (automatic Superscript) else if $iSuperScript is set, let that overwrite
		;	the current setting, else if subscript is true or set to an integer, it will overwrite the setting. If nothing
		; else set Subscript to 1
		$iSuperScript = ($bAutoSuper) ? (14000) : ((IsInt($iSuperScript)) ? ($iSuperScript) : ((IsInt($iSubScript) Or ($bAutoSub = True)) ? ($iSuperScript) : (1)))
	EndIf

	If ($bAutoSub <> Null) Then
		If Not IsBool($bAutoSub) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		; If $bAutoSub = True set it to -14000 (automatic Subscript) else if $iSubScript is set, let that overwrite
		;	the current setting, else if superscript is true or set to an integer, it will overwrite the setting.
		$iSubScript = ($bAutoSub) ? (-14000) : ((IsInt($iSubScript)) ? ($iSubScript) : ((IsInt($iSuperScript)) ? ($iSubScript) : (1)))
	EndIf

	If ($iSuperScript <> Null) Then
		If Not __LO_IntIsBetween($iSuperScript, 0, 100, "", 14000) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.CharEscapement = $iSuperScript
		$iError = ($oObj.CharEscapement() = $iSuperScript) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iSubScript <> Null) Then
		If Not __LO_IntIsBetween($iSubScript, -100, 100, "", "-14000:14000") Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$iSubScript = ($iSubScript > 0) ? Int("-" & $iSubScript) : $iSubScript
		$oObj.CharEscapement = $iSubScript
		$iError = ($oObj.CharEscapement() = $iSubScript) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iRelativeSize <> Null) Then
		If Not __LO_IntIsBetween($iRelativeSize, 1, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oObj.CharEscapementHeight = $iRelativeSize
		$iError = ($oObj.CharEscapementHeight() = $iRelativeSize) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_CharPosition

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharRotateScale
; Description ...: Set or retrieve the character rotational and Scale settings.
; Syntax ........: __LOWriter_CharRotateScale(ByRef $oObj, $iRotation, $iScaleWidth[, $bRotateFitLine = Null])
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $iRotation           - an integer value (0,90,270). Degrees to rotate the text.
;                  $iScaleWidth         - an integer value (1-100). The percentage to horizontally stretch or compress the text. 100 is normal sizing.
;                  $bRotateFitLine      - [optional] a boolean value. Default is Null. If True, Stretches or compresses the selected text so that it fits between the line that is above the text and the line that is below the text. Only works with Direct Formatting.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iRotation not an Integer or not equal to 0, 90 or 270 degrees.
;                  @Error 1 @Extended 5 Return 0 = $iScaleWidth not an Integer or less than 1% or greater than 100%.
;                  @Error 1 @Extended 6 Return 0 = $bRotateFitLine not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iRotation
;                  |                               2 = Error setting $iScaleWidth
;                  |                               4 = Error setting $bRotateFitLine
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters. Note: Excludes $bRotateFitLine, which is added onto the Direct Formatting function return.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharRotateScale(ByRef $oObj, $iRotation, $iScaleWidth, $bRotateFitLine = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avRotation[2]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($iRotation, $iScaleWidth, $bRotateFitLine) Then
		; rotation set in hundredths (90 deg = 900 etc)so divide by 10.
		__LO_ArrayFill($avRotation, Int($oObj.CharRotation() / 10), $oObj.CharScaleWidth())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avRotation)
	EndIf

	If ($iRotation <> Null) Then
		If Not __LO_IntIsBetween($iRotation, 0, 0, "", "90:270") Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$iRotation = Int($iRotation * 10) ; Rotation set in hundredths (90 deg = 900 etc)so times by 10.
		$oObj.CharRotation = $iRotation
		$iError = ($oObj.CharRotation() = $iRotation) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iScaleWidth <> Null) Then ; can't be less than 1%
		If Not __LO_IntIsBetween($iScaleWidth, 1, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.CharScaleWidth = $iScaleWidth
		$iError = ($oObj.CharScaleWidth() = $iScaleWidth) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bRotateFitLine <> Null) Then
		; works only on Direct Formatting:
		If Not IsBool($bRotateFitLine) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.CharRotationIsFitToLine = $bRotateFitLine
		$iError = ($oObj.CharRotationIsFitToLine() = $bRotateFitLine) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_CharRotateScale

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharShadow
; Description ...: Set and retrieve the Shadow for a Character Style.
; Syntax ........: __LOWriter_CharShadow(ByRef $oObj, $iWidth, $iColor, $bTransparent, $iLocation)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $iWidth              - an integer value. The Shadow width, set in Micrometers.
;                  $iColor              - an integer value (0-16777215). The Shadow color. See Remarks. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bTransparent        - a boolean value. If True, the shadow is transparent.
;                  $iLocation           - an integer value (0-4). Location of the shadow compared to the characters. See Constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iColor not an Integer, or less than 0, or greater than 16777215.
;                  @Error 1 @Extended 6 Return 0 = $bTransparent not a boolean.
;                  @Error 1 @Extended 7 Return 0 = $iLocation not an Integer, or less than 0, or greater than 4. See Constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Shadow format Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Shadow format Object for Error Checking.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iWidth
;                  |                               2 = Error setting $iColor
;                  |                               4 = Error setting $bTransparent
;                  |                               8 = Error setting $iLocation
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  LibreOffice may adjust the set width +/- 1 Micrometer after setting.
;                  Color is set in Long Integer format.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharShadow(ByRef $oObj, $iWidth, $iColor, $bTransparent, $iLocation)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $tShdwFrmt
	Local $avShadow[4]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tShdwFrmt = $oObj.CharShadowFormat()
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iWidth, $iColor, $bTransparent, $iLocation) Then
		__LO_ArrayFill($avShadow, $tShdwFrmt.ShadowWidth(), $tShdwFrmt.Color(), $tShdwFrmt.IsTransparent(), $tShdwFrmt.Location())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avShadow)
	EndIf

	If ($iWidth <> Null) Then
		If Not IsInt($iWidth) Or ($iWidth < 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tShdwFrmt.ShadowWidth = $iWidth
	EndIf

	If ($iColor <> Null) Then
		If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tShdwFrmt.Color = $iColor
	EndIf

	If ($bTransparent <> Null) Then
		If Not IsBool($bTransparent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tShdwFrmt.IsTransparent = $bTransparent
	EndIf

	If ($iLocation <> Null) Then
		If Not __LO_IntIsBetween($iLocation, $LOW_SHADOW_NONE, $LOW_SHADOW_BOTTOM_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tShdwFrmt.Location = $iLocation
	EndIf

	$oObj.CharShadowFormat = $tShdwFrmt
	$tShdwFrmt = $oObj.CharShadowFormat
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$iError = ($iWidth = Null) ? ($iError) : (($tShdwFrmt.ShadowWidth() = $iWidth) ? ($iError) : (BitOR($iError, 1)))
	$iError = ($iColor = Null) ? ($iError) : (($tShdwFrmt.Color() = $iColor) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($bTransparent = Null) ? ($iError) : (($tShdwFrmt.IsTransparent() = $bTransparent) ? ($iError) : (BitOR($iError, 4)))
	$iError = ($iLocation = Null) ? ($iError) : (($tShdwFrmt.Location() = $iLocation) ? ($iError) : (BitOR($iError, 8)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_CharShadow

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharSpacing
; Description ...: Set and retrieve the spacing between characters (Kerning).
; Syntax ........: __LOWriter_CharSpacing(ByRef $oObj, $bAutoKerning, $nKerning)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $bAutoKerning        - a boolean value. If True, applies a spacing in between certain pairs of characters.
;                  $nKerning            - a general number value (-2-928.8). The kerning value of the characters. See Remarks. Values are in Printer's Points as set in the Libre Office UI.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bAutoKerning not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $nKerning not a number, or less than -2 or greater than 928.8 Points.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bAutoKerning
;                  |                               2 = Error setting $nKerning.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  When setting Kerning values in LibreOffice, the measurement is listed in Pt (Printer's Points) in the User Display, however the internal setting is measured in Micrometers. They will be automatically converted from Points to MicroMeters and back for retrieval of settings.
;                  The acceptable values for $nKerning are from -2 Pt to 928.8 Pt. the figures can be directly converted easily, however, for an unknown reason to myself, LibreOffice begins counting backwards and in negative Micrometers internally from 928.9 up to 1000 Pt (Max setting).
;                  For example, 928.8Pt is the last correct value, which equals 32766 uM (Micrometers), after this LibreOffice reports the following: ;928.9 Pt = -32766 uM; 929 Pt = -32763 uM; 929.1 = -32759; 1000 pt = -30258.
;                  Attempting to set Libre's kerning value to anything over 32768 uM causes a COM exception, and attempting to set the kerning to any of these negative numbers sets the User viewable kerning value to -2.0 Pt. For these reasons the max settable kerning is -2.0 Pt to 928.8 Pt.
; Related .......: _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharSpacing(ByRef $oObj, $bAutoKerning, $nKerning)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avKerning[2]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($bAutoKerning, $nKerning) Then
		$nKerning = __LO_UnitConvert($oObj.CharKerning(), $__LOCONST_CONVERT_UM_PT)
		__LO_ArrayFill($avKerning, $oObj.CharAutoKerning(), (($nKerning > 928.8) ? (1000) : ($nKerning)))

		Return SetError($__LO_STATUS_SUCCESS, 1, $avKerning)
	EndIf

	If ($bAutoKerning <> Null) Then
		If Not IsBool($bAutoKerning) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.CharAutoKerning = $bAutoKerning
		$iError = ($oObj.CharAutoKerning() = $bAutoKerning) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($nKerning <> Null) Then
		If Not __LO_NumIsBetween($nKerning, -2, 928.8) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$nKerning = __LO_UnitConvert($nKerning, $__LOCONST_CONVERT_PT_UM)

		$oObj.CharKerning = $nKerning
		$iError = ($oObj.CharKerning() = $nKerning) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_CharSpacing

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharStrikeOut
; Description ...: Set or Retrieve the StrikeOut settings,
; Syntax ........: __LOWriter_CharStrikeOut(ByRef $oObj, $bWordOnly, $bStrikeOut, $iStrikeLineStyle)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $bWordOnly           - a boolean value. If True, strikeout is applied to words only skipping whitespaces.
;                  $bStrikeOut          - a boolean value. If True, strikeout is applied to characters.
;                  $iStrikeLineStyle    - an integer value (0-6). The Strikeout Line Style, see constants, $LOW_STRIKEOUT_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bStrikeOut not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iStrikeLineStyle not an Integer, or less than 0, or greater than 6. See constants, $LOW_STRIKEOUT_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bWordOnly
;                  |                               2 = Error setting $bStrikeOut
;                  |                               4 = Error setting $iStrikeLineStyle
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Strikeout is converted to single line in Ms word document format.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharStrikeOut(ByRef $oObj, $bWordOnly, $bStrikeOut, $iStrikeLineStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avStrikeOut[3]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($bWordOnly, $bStrikeOut, $iStrikeLineStyle) Then
		__LO_ArrayFill($avStrikeOut, $oObj.CharWordMode(), $oObj.CharCrossedOut(), $oObj.CharStrikeout())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avStrikeOut)
	EndIf

	If ($bWordOnly <> Null) Then
		If Not IsBool($bWordOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.CharWordMode = $bWordOnly
		$iError = ($oObj.CharWordMode() = $bWordOnly) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bStrikeOut <> Null) Then
		If Not IsBool($bStrikeOut) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.CharCrossedOut = $bStrikeOut
		$iError = ($oObj.CharCrossedOut() = $bStrikeOut) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iStrikeLineStyle <> Null) Then
		If Not __LO_IntIsBetween($iStrikeLineStyle, $LOW_STRIKEOUT_NONE, $LOW_STRIKEOUT_X) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.CharStrikeout = $iStrikeLineStyle
		$iError = ($oObj.CharStrikeout() = $iStrikeLineStyle) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_CharStrikeOut

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharStyleNameToggle
; Description ...: Toggle from Character Style Display Name to Internal Name for error checking and setting retrieval.
; Syntax ........: __LOWriter_CharStyleNameToggle(ByRef $sCharStyle[, $bReverse = False])
; Parameters ....: $sCharStyle          - a string value. The Character Style Name to Toggle.
;                  $bReverse            - [optional] a boolean value. Default is False. If True, the Character Style name is reverse toggled.
; Return values .: Success: String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sCharStyle not a String.
;                  @Error 1 @Extended 2 Return 0 = $bReverse not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Character Style Name was successfully toggled. Returning toggled name as a string.
;                  @Error 0 @Extended 1 Return String = Success. Character Style Name was successfully reverse toggled. Returning toggled name as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharStyleNameToggle($sCharStyle, $bReverse = False)
	If Not IsString($sCharStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReverse) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If ($bReverse = False) Then
		$sCharStyle = ($sCharStyle = "Footnote Characters") ? ("Footnote Symbol") : ($sCharStyle)
		$sCharStyle = ($sCharStyle = "Bullets") ? ("Bullet Symbols") : ($sCharStyle)
		$sCharStyle = ($sCharStyle = "Endnote Characters") ? ("Endnote Symbol") : ($sCharStyle)
		$sCharStyle = ($sCharStyle = "Quotation") ? ("Citation") : ($sCharStyle)
		$sCharStyle = ($sCharStyle = "No Character Style") ? ("Standard") : ($sCharStyle)

		Return SetError($__LO_STATUS_SUCCESS, 0, $sCharStyle)

	Else
		$sCharStyle = ($sCharStyle = "Footnote Symbol") ? ("Footnote Characters") : ($sCharStyle)
		$sCharStyle = ($sCharStyle = "Bullet Symbols") ? ("Bullets") : ($sCharStyle)
		$sCharStyle = ($sCharStyle = "Endnote Symbol") ? ("Endnote Characters") : ($sCharStyle)
		$sCharStyle = ($sCharStyle = "Citation") ? ("Quotation") : ($sCharStyle)
		$sCharStyle = ($sCharStyle = "Standard") ? ("No Character Style") : ($sCharStyle)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sCharStyle)
	EndIf
EndFunc   ;==>__LOWriter_CharStyleNameToggle

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharUnderLine
; Description ...: Set and retrieve the Underline settings.
; Syntax ........: __LOWriter_CharUnderLine(ByRef $oObj, $bWordOnly, $iUnderLineStyle, $bULHasColor, $iULColor)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If true, white spaces are not underlined.
;                  $iUnderLineStyle     - [optional] an integer value (0-18). Default is Null. The line style of the Underline, see constants, $LOW_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bULHasColor         - [optional] a boolean value. Default is Null. If True, the underline is colored, must be set to true in order to set the underline color.
;                  $iULColor            - [optional] an integer value (-1-16777215). Default is Null. The color of the underline, set in Long integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Set to $LO_COLOR_OFF(-1) for automatic color mode.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iUnderLineStyle not an Integer, or less than 0, or greater than 18. See constants, $LOW_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bULHasColor not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $iULColor not an Integer, or less than -1, or greater than 16777215.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bWordOnly
;                  |                               2 = Error setting $iUnderLineStyle
;                  |                               4 = Error setting $ULHasColor
;                  |                               8 = Error setting $iULColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  $bULHasColor must be set to true in order to set the underline color.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharUnderLine(ByRef $oObj, $bWordOnly, $iUnderLineStyle, $bULHasColor, $iULColor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avUnderLine[4]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($bWordOnly, $iUnderLineStyle, $bULHasColor, $iULColor) Then
		__LO_ArrayFill($avUnderLine, $oObj.CharWordMode(), $oObj.CharUnderline(), $oObj.CharUnderlineHasColor(), $oObj.CharUnderlineColor())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avUnderLine)
	EndIf

	If ($bWordOnly <> Null) Then
		If Not IsBool($bWordOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.CharWordMode = $bWordOnly
		$iError = ($oObj.CharWordMode() = $bWordOnly) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iUnderLineStyle <> Null) Then
		If Not __LO_IntIsBetween($iUnderLineStyle, $LOW_UNDERLINE_NONE, $LOW_UNDERLINE_BOLD_WAVE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.CharUnderline = $iUnderLineStyle
		$iError = ($oObj.CharUnderline() = $iUnderLineStyle) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bULHasColor <> Null) Then
		If Not IsBool($bULHasColor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.CharUnderlineHasColor = $bULHasColor
		$iError = ($oObj.CharUnderlineHasColor() = $bULHasColor) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iULColor <> Null) Then
		If Not __LO_IntIsBetween($iULColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oObj.CharUnderlineColor = $iULColor
		$iError = ($oObj.CharUnderlineColor() = $iULColor) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_CharUnderLine

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ColorRemoveAlpha
; Description ...: Remove the Alpha value from a Long color value.
; Syntax ........: __LOWriter_ColorRemoveAlpha($iColor)
; Parameters ....: $iColor              - an integer value. A Long Color value to remove Alpha from.
; Return values .: Success: Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return $iColor = $iColor not an Integer. Returning $iColor to be sure not to lose the value.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Color already has no Alpha value, returning same color.
;                  @Error 0 @Extended 1 Return Integer = Success. Removed Alpha value from Long Color value, returning new Color value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: In functions which return the current color value, generally background colors, if Transparency (alpha) is set, the background color value is not the literal color set, but also includes the transparency value added to it. This functions removes that value for simpler color values.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ColorRemoveAlpha($iColor)
	Local $iRed, $iGreen, $iBlue, $iLong

	If Not IsInt($iColor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, $iColor)

	If __LO_IntIsBetween($iColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_SUCCESS, 0, $iColor) ; If Color value is not greater than White(16777215) or less than -1, then there is no alpha to remove.

	; Obtain individual color values.
	$iRed = BitAND(BitShift($iColor, 16), 0xff)
	$iGreen = BitAND(BitShift($iColor, 8), 0xff)
	$iBlue = BitAND($iColor, 0xff)
	$iLong = BitShift($iRed, -16) + BitShift($iGreen, -8) + $iBlue

	Return SetError($__LO_STATUS_SUCCESS, 1, $iLong)
EndFunc   ;==>__LOWriter_ColorRemoveAlpha

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CreatePoint
; Description ...: Creates a Position structure.
; Syntax ........: __LOWriter_CreatePoint($iX, $iY)
; Parameters ....: $iX                  - an integer value. The X position, in Micrometers.
;                  $iY                  - an integer value. The Y position, in Micrometers.
; Return values .: Success: Structure
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iX not an Integer.
;                  @Error 1 @Extended 2 Return 0 = $iY not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create a Position Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Structure = Success. Returning created Position Structure set to $iX and $iY values.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Modified from A. Pitonyak, Listing 493. in OOME 3.0
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CreatePoint($iX, $iY)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tPoint

	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tPoint = __LO_CreateStruct("com.sun.star.awt.Point")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tPoint.X = $iX
	$tPoint.Y = $iY

	Return SetError($__LO_STATUS_SUCCESS, 0, $tPoint)
EndFunc   ;==>__LOWriter_CreatePoint

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CursorGetText
; Description ...: Retrieves a Text object appropriate for the type of cursor.
; Syntax ........: __LOWriter_CursorGetText(ByRef $oDoc, $oCursor)
; Parameters ....: $oDoc                - [in/out] A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Text or View Cursor Object returned from any Cursor Object creation or retrieval functions.
; Return values .: Success: Object.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to get Cursor data type.
;                  @Error 3 @Extended 2 Return 0 = Failed to create Object for creating TextObject.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Text Object.
;                  @Error 3 @Extended 4 Return 0 = Cursor is in an unknown data field.
;                  --Success--
;                  @Error 0 @Extended ? Return Object = Success, Text object was returned. @Extended will be one of the constants, $LOW_CURDATA_* as defined in LibreOfficeWriter_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Also returns what type of cursor, such as a text Table, footnote etc.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CursorGetText(ByRef $oDoc, ByRef $oCursor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oText, $oReturnedObj
	Local $iCursorDataType

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oReturnedObj = __LOWriter_Internal_CursorGetDataType($oDoc, $oCursor, True)
	$iCursorDataType = @extended
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not IsObj($oReturnedObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Switch $iCursorDataType
		Case $LOW_CURDATA_BODY_TEXT, $LOW_CURDATA_FRAME, $LOW_CURDATA_FOOTNOTE, $LOW_CURDATA_ENDNOTE, $LOW_CURDATA_HEADER_FOOTER
			$oText = $oReturnedObj.getText()
			If Not IsObj($oText) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			Return SetError($__LO_STATUS_SUCCESS, $iCursorDataType, $oText)

		Case $LOW_CURDATA_CELL
			$oText = $oReturnedObj.getCellByName($oCursor.Cell.CellName)
			If Not IsObj($oText) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			Return SetError($__LO_STATUS_SUCCESS, $iCursorDataType, $oText)

		Case Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndSwitch
EndFunc   ;==>__LOWriter_CursorGetText

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_DateStructCompare
; Description ...: Compare two date Structures to see if they are the same Date, Time, etc.
; Syntax ........: __LOWriter_DateStructCompare($tDateStruct1, $tDateStruct2[, $bIsDate = False[, $bIsTime = False]])
; Parameters ....: $tDateStruct1        - a dll struct value. The First Date Structure.
;                  $tDateStruct2        - a dll struct value. The Second Date Structure.
;                  $bIsDate             - [optional] a boolean value. Default is False. If True, the comparison is two Date Structures.
;                  $bIsTime             - [optional] a boolean value. Default is False. If True, the comparison is two Time Structures.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return False = $tDateStruct1 not an Object.
;                  @Error 1 @Extended 2 Return False = $tDateStruct2 not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If the Dates/Times in $tDateStruct1 and $tDateStruct2 are the same, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If both $bIsDate and $bIsTime are False, the comparison is two Date and Time Structures
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_DateStructCompare($tDateStruct1, $tDateStruct2, $bIsDate = False, $bIsTime = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($tDateStruct1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, False)
	If Not IsObj($tDateStruct2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, False)
	If Not IsBool($bIsDate) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, False)
	If Not IsBool($bIsTime) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, False)

	If $bIsDate Then
		If $tDateStruct1.Year() <> $tDateStruct2.Year() Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
		If $tDateStruct1.Month() <> $tDateStruct2.Month() Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
		If $tDateStruct1.Day() <> $tDateStruct2.Day() Then Return SetError($__LO_STATUS_SUCCESS, 0, False)

	ElseIf $bIsTime Then
		If $tDateStruct1.Hours() <> $tDateStruct2.Hours() Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
		If $tDateStruct1.Minutes() <> $tDateStruct2.Minutes() Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
		If $tDateStruct1.Seconds() <> $tDateStruct2.Seconds() Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
		If $tDateStruct1.NanoSeconds() <> $tDateStruct2.NanoSeconds() Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
		If __LO_VersionCheck(4.1) Then
			If $tDateStruct1.IsUTC() <> $tDateStruct2.IsUTC() Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
		EndIf

	Else
		If $tDateStruct1.Year() <> $tDateStruct2.Year() Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
		If $tDateStruct1.Month() <> $tDateStruct2.Month() Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
		If $tDateStruct1.Day() <> $tDateStruct2.Day() Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
		If $tDateStruct1.Hours() <> $tDateStruct2.Hours() Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
		If $tDateStruct1.Minutes() <> $tDateStruct2.Minutes() Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
		If $tDateStruct1.Seconds() <> $tDateStruct2.Seconds() Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
		If $tDateStruct1.NanoSeconds() <> $tDateStruct2.NanoSeconds() Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
		If __LO_VersionCheck(4.1) Then
			If $tDateStruct1.IsUTC() <> $tDateStruct2.IsUTC() Then Return SetError($__LO_STATUS_SUCCESS, 0, False)
		EndIf
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, True)
EndFunc   ;==>__LOWriter_DateStructCompare

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_DirFrmtCheck
; Description ...: Do checks on Dirformat input object.
; Syntax ........: __LOWriter_DirFrmtCheck(ByRef $oSelection[, $bCheckSelection = False])
; Parameters ....: $oSelection          - [in/out] an object. The Object to check, which should be either a cursor with data selected or a paragraph object.
;                  $bCheckSelection     - [optional] a boolean value. Default is False. If True, check for whether the cursor object is collapsed (no data selected).
; Return values .: Success: Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Returns True, if called Object is fit for Direct Formatting use, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_DirFrmtCheck(ByRef $oSelection, $bCheckSelection = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	; If Object supports Paragraph or TextPortion services, then return True.
	If $oSelection.supportsService("com.sun.star.text.Paragraph") Or _
			$oSelection.supportsService("com.sun.star.text.TextPortion") Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

	; If Object is a cursor then return true if $bcheckSelection is false. Else test if cursor selection is collapsed, return
	; false if it is.
	If $oSelection.supportsService("com.sun.star.text.TextCursor") Or _
			$oSelection.supportsService("com.sun.star.text.TextViewCursor") Then
		If $bCheckSelection Then Return SetError($__LO_STATUS_SUCCESS, 0, ($oSelection.IsCollapsed()) ? (False) : (True)) ; If collapsed return false meaning fail.

		Return SetError($__LO_STATUS_SUCCESS, 0, True)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>__LOWriter_DirFrmtCheck

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_FieldCountType
; Description ...: Determine a Count Field's type.
; Syntax ........: __LOWriter_FieldCountType($vInput)
; Parameters ....: $vInput              - a variant value. Either a Field Object to determine the appropriate integer Constant to return, or a Integer Constant to return the appropriate Field type String. See constants, $LOW_FIELD_COUNT_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: String or Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $vInput is neither a String nor an Integer.
;                  @Error 1 @Extended 2 Return 0 = $vInput was an Object, but did not match any known counting fields.
;                  @Error 1 @Extended 3 Return 0 = $vInput was an Integer but is higher than the size of the array of Field types. See Constants, $LOW_FIELD_COUNT_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Counting Field type identified, returning FieldCountType constant Integer. See Constants, $LOW_FIELD_COUNT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 0 @Extended 1 Return String = Success. Counting Field type identified, returning Field Count Type String for CreateInstance function.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_FieldCountType($vInput)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asFieldTypes[7]
	$asFieldTypes[$LOW_FIELD_COUNT_TYPE_CHARACTERS] = "com.sun.star.text.TextField.CharacterCount"
	$asFieldTypes[$LOW_FIELD_COUNT_TYPE_IMAGES] = "com.sun.star.text.TextField.GraphicObjectCount"
	$asFieldTypes[$LOW_FIELD_COUNT_TYPE_OBJECTS] = "com.sun.star.text.TextField.EmbeddedObjectCount"
	$asFieldTypes[$LOW_FIELD_COUNT_TYPE_PAGES] = "com.sun.star.text.TextField.PageCount"
	$asFieldTypes[$LOW_FIELD_COUNT_TYPE_PARAGRAPHS] = "com.sun.star.text.TextField.ParagraphCount"
	$asFieldTypes[$LOW_FIELD_COUNT_TYPE_TABLES] = "com.sun.star.text.TextField.TableCount"
	$asFieldTypes[$LOW_FIELD_COUNT_TYPE_WORDS] = "com.sun.star.text.TextField.WordCount"

	If IsObj($vInput) Then
		For $i = 0 To UBound($asFieldTypes) - 1
			If $vInput.supportsService($asFieldTypes[$i]) Then Return SetError($__LO_STATUS_SUCCESS, 0, $i)
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
		Next

		Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; No Hits

	ElseIf IsInt($vInput) Then
		If Not __LO_IntIsBetween($vInput, 0, UBound($asFieldTypes) - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $asFieldTypes[$vInput])

	Else

		Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0) ; Wrong VarType
	EndIf
EndFunc   ;==>__LOWriter_FieldCountType

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_FieldsGetList
; Description ...: Internal Function to retrieve a Field Object list.
; Syntax ........: __LOWriter_FieldsGetList(ByRef $oDoc, $bSupportedServices, $bFieldType, $bFieldTypeNum, ByRef $avFieldTypes)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bSupportedServices  - a boolean value. If True, adds a column to the array that has the supported service String for that particular Field, To assist in identifying the Field type.
;                  $bFieldType          - [optional] a boolean value. Default is True. If True, adds a column to the array that has the Field Type String for that particular Field as described by Libre Office. To assist in identifying the Field type.
;                  $bFieldTypeNum       - [optional] a boolean value. Default is True. If True, adds a column to the array that has the Field Type Constant for that particular Field, to assist in identifying the Field type. See remarks.
;                  $avFieldTypes        - [in/out] an array of variants. An Array of Field types to search for to return. Array will not be modified.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 3 Return 0 = $bSupportedServices not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bFieldType not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bFieldTypeNum not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $avFieldTypes not an Array.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create enumeration of fields in document.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning Array of Text Field Objects with @Extended set to number of results. See Remarks for Array sizing.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Array can vary in the number of columns, if $bSupportedServices, $bFieldType, and $bFieldTypeNum are set to False, the Array will be a single column. With each of the above listed options being set to True, a column will be added in the order they are listed in the UDF parameters.
;                  The First column will always be the Field Object.
;                  Setting $bSupportedServices to True will add a Supported Service String column for the found Field.
;                  Setting $bFieldType to True will add a Field type column for the found Field.
;                  Setting $bFieldTypeNum to True will add a Field type Number column, matching one of the following constants for the found Field. $LOW_FIELD_TYPE_*,$LOW_FIELD_ADV_TYPE_*, and $LOW_FIELD_DOCINFO_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_FieldsGetList(ByRef $oDoc, $bSupportedServices, $bFieldType, $bFieldTypeNum, ByRef $avFieldTypes)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextFields, $oTextField
	Local $iCount = 0, $iColumns = 4, $iFieldTypeCol = 2, $iFieldTypeNumCol = 3
	Local $avTextFields[50][4]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	; Skip 2 to match other Funcs.
	If Not IsBool($bSupportedServices) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bFieldType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bFieldTypeNum) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsArray($avFieldTypes) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$iColumns = ($bSupportedServices = False) ? ($iColumns - 1) : ($iColumns)
	$iColumns = ($bFieldType = False) ? ($iColumns - 1) : ($iColumns)
	$iColumns = ($bFieldTypeNum = False) ? ($iColumns - 1) : ($iColumns)

	If ($iColumns = 1) Then ReDim $avTextFields[UBound($avTextFields)]

	; If Supported Services Option is False, change the column position of FieldType
	$iFieldTypeCol = ($bSupportedServices = False) ? ($iFieldTypeCol - 1) : ($iFieldTypeCol)

	$iFieldTypeNumCol = ($bSupportedServices = False) ? ($iFieldTypeNumCol - 1) : ($iFieldTypeNumCol)
	$iFieldTypeNumCol = ($bFieldType = False) ? ($iFieldTypeNumCol - 1) : ($iFieldTypeNumCol)

	$oTextFields = $oDoc.getTextFields.createEnumeration()
	If Not IsObj($oTextFields) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	While $oTextFields.hasMoreElements()
		$oTextField = $oTextFields.nextElement()

		For $i = 0 To UBound($avFieldTypes) - 1
			If $oTextField.supportsService($avFieldTypes[$i][1]) Then
				If ($iColumns = 1) Then
					$avTextFields[$iCount] = $oTextField

					$iCount += 1
					If ($iCount = UBound($avTextFields)) Then ReDim $avTextFields[$iCount * 2]
					ExitLoop

				Else
					$avTextFields[$iCount][0] = $oTextField

					If ($bSupportedServices = True) Then $avTextFields[$iCount][1] = $avFieldTypes[$i][1]
					If ($bFieldType = True) Then $avTextFields[$iCount][$iFieldTypeCol] = $oTextField.getPresentation(True)
					If ($bFieldTypeNum = True) Then $avTextFields[$iCount][$iFieldTypeNumCol] = $avFieldTypes[$i][0]

					$iCount += 1
					If ($iCount = UBound($avTextFields)) Then ReDim $avTextFields[$iCount * 2][$iColumns]
					ExitLoop
				EndIf
			EndIf
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
		Next
	WEnd

	If ($iColumns = 1) Then
		ReDim $avTextFields[$iCount]

	Else
		ReDim $avTextFields[$iCount][$iColumns]
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $avTextFields)
EndFunc   ;==>__LOWriter_FieldsGetList

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_FieldTypeServices
; Description ...: Retrieve an Array of Supported Service Names and Integer Constants to search for Fields.
; Syntax ........: __LOWriter_FieldTypeServices($iFieldType[, $bAdvancedServices = False[, $bDocInfoServices = False]])
; Parameters ....: $iFieldType          - an integer value. The Constant Field type.
;                  $bAdvancedServices   - [optional] a boolean value. Default is False. If True, search in Advanced Field Type Array.
;                  $bDocInfoServices    - [optional] a boolean value. Default is False. If True, search in Document Information Field Type Array.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iFieldType not an Integer.
;                  @Error 1 @Extended 2 Return 0 = $bAdvancedServices not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bDocInfoServices not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Something went wrong determining what Array to search/Return.
;                  --Success--
;                  @Error 0 @Extended 0 Return Array = Success. $iFieldType set to All, $bAdvancedServices and $bDocInfoServices both set to false, returning full regular Field Service list String Array.
;                  @Error 0 @Extended 1 Return Array = Success. $iFieldType set to All, $bAdvancedServices set to True and $bDocInfoServices set to false, returning full Advanced Field Service String list Array.
;                  @Error 0 @Extended 2 Return Array = Success. $iFieldType set to All, $bAdvancedServices set to False and $bDocInfoServices set to True, returning full DocInfo Field Service String list Array.
;                  @Error 0 @Extended 3 Return Array = Success. $iFieldType BitOr'd together, determining which flags are called from the specified Array. Returning Field Service String list Array.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_FieldTypeServices($iFieldType, $bAdvancedServices = False, $bDocInfoServices = False)
	Local $avFieldTypes[29][2] = [[$LOW_FIELD_TYPE_COMMENT, "com.sun.star.text.TextField.Annotation"], _
			[$LOW_FIELD_TYPE_AUTHOR, "com.sun.star.text.TextField.Author"], [$LOW_FIELD_TYPE_CHAPTER, "com.sun.star.text.TextField.Chapter"], _
			[$LOW_FIELD_TYPE_CHAR_COUNT, "com.sun.star.text.TextField.CharacterCount"], [$LOW_FIELD_TYPE_COMBINED_CHAR, "com.sun.star.text.TextField.CombinedCharacters"], _
			[$LOW_FIELD_TYPE_COND_TEXT, "com.sun.star.text.TextField.ConditionalText"], [$LOW_FIELD_TYPE_DATE_TIME, "com.sun.star.text.TextField.DateTime"], _
			[$LOW_FIELD_TYPE_INPUT_LIST, "com.sun.star.text.TextField.DropDown"], [$LOW_FIELD_TYPE_EMB_OBJ_COUNT, "com.sun.star.text.TextField.EmbeddedObjectCount"], _
			[$LOW_FIELD_TYPE_SENDER, "com.sun.star.text.TextField.ExtendedUser"], [$LOW_FIELD_TYPE_FILENAME, "com.sun.star.text.TextField.FileName"], _
			[$LOW_FIELD_TYPE_SHOW_VAR, "com.sun.star.text.TextField.GetExpression"], [$LOW_FIELD_TYPE_INSERT_REF, "com.sun.star.text.TextField.GetReference"], _
			[$LOW_FIELD_TYPE_IMAGE_COUNT, "com.sun.star.text.TextField.GraphicObjectCount"], [$LOW_FIELD_TYPE_HIDDEN_PAR, "com.sun.star.text.TextField.HiddenParagraph"], _
			[$LOW_FIELD_TYPE_HIDDEN_TEXT, "com.sun.star.text.TextField.HiddenText"], [$LOW_FIELD_TYPE_INPUT, "com.sun.star.text.TextField.Input"], _
			[$LOW_FIELD_TYPE_PLACEHOLDER, "com.sun.star.text.TextField.JumpEdit"], [$LOW_FIELD_TYPE_MACRO, "com.sun.star.text.TextField.Macro"], _
			[$LOW_FIELD_TYPE_PAGE_COUNT, "com.sun.star.text.TextField.PageCount"], [$LOW_FIELD_TYPE_PAGE_NUM, "com.sun.star.text.TextField.PageNumber"], _
			[$LOW_FIELD_TYPE_PAR_COUNT, "com.sun.star.text.TextField.ParagraphCount"], [$LOW_FIELD_TYPE_SHOW_PAGE_VAR, "com.sun.star.text.TextField.ReferencePageGet"], _
			[$LOW_FIELD_TYPE_SET_PAGE_VAR, "com.sun.star.text.TextField.ReferencePageSet"], [$LOW_FIELD_TYPE_SCRIPT, "com.sun.star.text.TextField.Script"], _
			[$LOW_FIELD_TYPE_SET_VAR, "com.sun.star.text.TextField.SetExpression"], [$LOW_FIELD_TYPE_TABLE_COUNT, "com.sun.star.text.TextField.TableCount"], _
			[$LOW_FIELD_TYPE_TEMPLATE_NAME, "com.sun.star.text.TextField.TemplateName"], [$LOW_FIELD_TYPE_WORD_COUNT, "com.sun.star.text.TextField.WordCount"]]

	Local $avFieldAdvTypes[9][2] = [[$LOW_FIELD_ADV_TYPE_BIBLIOGRAPHY, "com.sun.star.text.TextField.Bibliography"], _
			[$LOW_FIELD_ADV_TYPE_DATABASE, "com.sun.star.text.TextField.Database"], [$LOW_FIELD_ADV_TYPE_DATABASE_NAME, "com.sun.star.text.TextField.DatabaseName"], _
			[$LOW_FIELD_ADV_TYPE_DATABASE_NEXT_SET, "com.sun.star.text.TextField.DatabaseNextSet"], [$LOW_FIELD_ADV_TYPE_DATABASE_NAME_OF_SET, "com.sun.star.text.TextField.DatabaseNumberOfSet"], _
			[$LOW_FIELD_ADV_TYPE_DATABASE_SET_NUM, "com.sun.star.text.TextField.DatabaseSetNumber"], [$LOW_FIELD_ADV_TYPE_DDE, "com.sun.star.text.TextField.DDE"], _
			[$LOW_FIELD_ADV_TYPE_INPUT_USER, "com.sun.star.text.TextField.InputUser"], [$LOW_FIELD_ADV_TYPE_USER, "com.sun.star.text.TextField.User"]]

	Local $avFieldDocInfoTypes[13][2] = [[$LOW_FIELD_DOCINFO_TYPE_MOD_AUTH, "com.sun.star.text.TextField.DocInfo.ChangeAuthor"], _
			[$LOW_FIELD_DOCINFO_TYPE_MOD_DATE_TIME, "com.sun.star.text.TextField.DocInfo.ChangeDateTime"], _
			[$LOW_FIELD_DOCINFO_TYPE_CREATE_AUTH, "com.sun.star.text.TextField.DocInfo.CreateAuthor"], [$LOW_FIELD_DOCINFO_TYPE_CREATE_DATE_TIME, "com.sun.star.text.TextField.DocInfo.CreateDateTime"], _
			[$LOW_FIELD_DOCINFO_TYPE_CUSTOM, "com.sun.star.text.TextField.DocInfo.Custom"], [$LOW_FIELD_DOCINFO_TYPE_COMMENTS, "com.sun.star.text.TextField.DocInfo.Description"], _
			[$LOW_FIELD_DOCINFO_TYPE_EDIT_TIME, "com.sun.star.text.TextField.DocInfo.EditTime"], [$LOW_FIELD_DOCINFO_TYPE_KEYWORDS, "com.sun.star.text.TextField.DocInfo.KeyWords"], _
			[$LOW_FIELD_DOCINFO_TYPE_PRINT_AUTH, "com.sun.star.text.TextField.DocInfo.PrintAuthor"], [$LOW_FIELD_DOCINFO_TYPE_PRINT_DATE_TIME, "com.sun.star.text.TextField.DocInfo.PrintDateTime"], _
			[$LOW_FIELD_DOCINFO_TYPE_REVISION, "com.sun.star.text.TextField.DocInfo.Revision"], [$LOW_FIELD_DOCINFO_TYPE_SUBJECT, "com.sun.star.text.TextField.DocInfo.Subject"], _
			[$LOW_FIELD_DOCINFO_TYPE_TITLE, "com.sun.star.text.TextField.DocInfo.Title"]]

	Local $avSearch[0][0], $avFieldResults[UBound($avFieldTypes)][2]
	Local $iCount = 0

	If Not IsInt($iFieldType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bAdvancedServices) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bDocInfoServices) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If ($bAdvancedServices = False) And ($bDocInfoServices = False) Then
		If (BitAND($iFieldType, $LOW_FIELD_TYPE_ALL)) Then Return SetError($__LO_STATUS_SUCCESS, 0, $avFieldTypes)
		$avSearch = $avFieldTypes

	ElseIf ($bAdvancedServices = True) And ($bDocInfoServices = False) Then
		If (BitAND($iFieldType, $LOW_FIELD_ADV_TYPE_ALL)) Then Return SetError($__LO_STATUS_SUCCESS, 1, $avFieldAdvTypes)
		$avSearch = $avFieldAdvTypes

	ElseIf ($bDocInfoServices = True) And ($bAdvancedServices = False) Then
		If (BitAND($iFieldType, $LOW_FIELD_DOCINFO_TYPE_ALL)) Then Return SetError($__LO_STATUS_SUCCESS, 2, $avFieldDocInfoTypes)
		$avSearch = $avFieldDocInfoTypes

	Else

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	For $i = 0 To UBound($avSearch) - 1
		If BitAND($avSearch[$i][0], $iFieldType) Then
			$avFieldResults[$iCount][0] = $avSearch[$i][0]
			$avFieldResults[$iCount][1] = $avSearch[$i][1]
			$iCount += 1
		EndIf
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	ReDim $avFieldResults[$iCount][2]

	Return SetError($__LO_STATUS_SUCCESS, 3, $avFieldResults)
EndFunc   ;==>__LOWriter_FieldTypeServices

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_FilterNameGet
; Description ...: Retrieves the correct L.O. Filtername for use in SaveAs and Export.
; Syntax ........: __LOWriter_FilterNameGet(ByRef $sDocSavePath[, $bExportFilters = False])
; Parameters ....: $sDocSavePath        - [in/out] a string value. Full path with extension.
;                  $bExportFilters      - [optional] a boolean value. Default is False. If True, includes the FilterNames that can be used to Export only, in the search.
; Return values .: Success: String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sDocSavePath is not a string.
;                  @Error 1 @Extended 2 Return 0 = $bExportFilters not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $sDocSavePath is not a correct path or URL.
;                  --Success--
;                  @Error 0 @Extended 1 Return String = Success. Returns required filtername from "SaveAs" FilterNames.
;                  @Error 0 @Extended 2 Return String = Success. Returns required filtername from "Export" FilterNames.
;                  @Error 0 @Extended 3 Return String = FilterName not found for given file extension, defaulting to .odt file format and updating save path accordingly.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Searches a predefined list of extensions stored in an array.
;                  Not all FilterNames are listed, where multiple options were available for a given extension, the most recent Filter was used; Such as for .doc, there is the a 97 MsWord Filter available, and also a 95 MsWord, in this case 97 MsWord was used, as is listed in the "SaveAs" and "Export" dialogs.
;                  For finding your own FilterNames, see convertfilters.html found in L.O. Install Folder: LibreOffice\help\en-US\text\shared\guide -- Or See: "OOME_3_0", "OpenOffice.org Macros Explained OOME Third Edition" by Andrew D. Pitonyak, which has a handy Macro for listing all FilterNames, found on page 284 of the above book in the ODT format.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_FilterNameGet(ByRef $sDocSavePath, $bExportFilters = False)
	Local $iLength, $iSlashLocation, $iDotLocation
	Local Const $STR_NOCASESENSE = 0, $STR_STRIPALL = 8
	Local $sFileExtension, $sFilterName
	Local $msSaveAsFilters[], $msExportFilters[]

	If Not IsString($sDocSavePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bExportFilters) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$iLength = StringLen($sDocSavePath)

	$msSaveAsFilters[".doc"] = "MS Word 97"
	$msSaveAsFilters[".docm"] = "MS Word 2007 XML VBA"
	$msSaveAsFilters[".docx"] = "MS Word 2007 XML"
	$msSaveAsFilters[".dot"] = "MS Word 97"
	$msSaveAsFilters[".dotx"] = "MS Word 2007 XML Template"
	$msSaveAsFilters[".fodt"] = "OpenDocument Text Flat XML"
	$msSaveAsFilters[".html"] = "HTML (StarWriter)"
	$msSaveAsFilters[".odt"] = "writer8"
	$msSaveAsFilters[".ott"] = "writer8_template"
	$msSaveAsFilters[".rtf"] = "Rich Text Format"
	$msSaveAsFilters[".txt"] = "Text"
	$msSaveAsFilters[".uot"] = "UOF text"
	$msSaveAsFilters[".xml"] = "MS Word 2003 XML"

	If $bExportFilters Then
		$msExportFilters[".epub"] = "EPUB"
		$msExportFilters[".jfif"] = "writer_jpg_Export"
		$msExportFilters[".jif"] = "writer_jpg_Export"
		$msExportFilters[".jpe"] = "writer_jpg_Export"
		$msExportFilters[".jpg"] = "writer_jpg_Export"
		$msExportFilters[".jpeg"] = "writer_jpg_Export"
		$msExportFilters[".pdf"] = "writer_pdf_Export"
		$msExportFilters[".png"] = "writer_png_Export"
		$msExportFilters[".xhtml"] = "XHTML Writer File"
	EndIf

	If StringInStr($sDocSavePath, "file:///") Then ;  If L.O. URL Then
		$iSlashLocation = StringInStr($sDocSavePath, "/", $STR_NOCASESENSE, -1)
		$iDotLocation = StringInStr($sDocSavePath, ".", $STR_NOCASESENSE, -1, $iLength, $iLength - $iSlashLocation)
		$sFileExtension = StringRight($sDocSavePath, ($iLength - $iDotLocation) + 1)

	ElseIf StringInStr($sDocSavePath, "\") Then ;  Else if PC Path Then
		$iSlashLocation = StringInStr($sDocSavePath, "\", $STR_NOCASESENSE, -1)
		$iDotLocation = StringInStr($sDocSavePath, ".", $STR_NOCASESENSE, -1, $iLength, $iLength - $iSlashLocation)
		$sFileExtension = StringRight($sDocSavePath, $iLength - $iDotLocation + 1)

	Else

		Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	EndIf

	If $sFileExtension = $sDocSavePath Then ;  If no file extension identified, append .odt extension and return.
		$sDocSavePath = $sDocSavePath & ".odt"

		Return SetError($__LO_STATUS_SUCCESS, 3, "writer8")

	Else
		$sFileExtension = StringLower(StringStripWS($sFileExtension, $STR_STRIPALL))
	EndIf

	$sFilterName = $msSaveAsFilters[$sFileExtension]

	If IsString($sFilterName) Then Return SetError($__LO_STATUS_SUCCESS, 1, $sFilterName)

	If $bExportFilters Then $sFilterName = $msExportFilters[$sFileExtension]

	If IsString($sFilterName) Then Return SetError($__LO_STATUS_SUCCESS, 2, $sFilterName)

	$sDocSavePath = StringReplace($sDocSavePath, $sFileExtension, ".odt") ; If No results, replace with ODT extension.

	Return SetError($__LO_STATUS_SUCCESS, 3, "writer8")
EndFunc   ;==>__LOWriter_FilterNameGet

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_FindFormatAddSetting
; Description ...: Add or Update a setting in a Find Format Array.
; Syntax ........: __LOWriter_FindFormatAddSetting(ByRef $atArray, $tSetting)
; Parameters ....: $atArray             - [in/out] an array of structs. A Find Format Array of Settings to Search. Array will be directly modified.
;                  $tSetting            - a struct value. A Libre Office Structure setting.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $aArray not an Array.
;                  @Error 1 @Extended 2 Return 0 = $tSetting not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sSettingName not a String.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Setting was successfully updated or added.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_FindFormatAddSetting(ByRef $atArray, $tSetting)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bFound = False
	Local $sSettingName

	If Not IsArray($atArray) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($tSetting) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$sSettingName = $tSetting.Name()
	If Not IsString($sSettingName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	For $i = 0 To UBound($atArray) - 1
		If $atArray[$i].Name() = $sSettingName Then
			$atArray[$i].Value = $tSetting.Value()
			$bFound = True
			ExitLoop
		EndIf

		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	If ($bFound = False) Then
		ReDim $atArray[UBound($atArray) + 1]
		$atArray[UBound($atArray) - 1] = $tSetting
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_FindFormatAddSetting

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_FindFormatDeleteSetting
; Description ...: Delete a setting from a Find Format Array.
; Syntax ........: __LOWriter_FindFormatDeleteSetting(ByRef $atArray, $sSettingName)
; Parameters ....: $atArray             - [in/out] an array of structs. A Find Format Array of Settings to Search. Array will be directly modified.
;                  $sSettingName        - a string value. The setting name to search and delete.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $aArray not an Array
;                  @Error 1 @Extended 2 Return 0 = $sSettingName not a String.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Setting was either not found or was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_FindFormatDeleteSetting(ByRef $atArray, $sSettingName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCount = 0

	If Not IsArray($atArray) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sSettingName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	For $i = 0 To UBound($atArray) - 1
		If $atArray[$i].Name() <> $sSettingName Then
			$atArray[$iCount] = $atArray[$i]
			$iCount += 1
		EndIf
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next
	ReDim $atArray[$iCount]

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_FindFormatDeleteSetting

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_FindFormatRetrieveSetting
; Description ...: Retrieve a specific setting from a Find Format Array of Settings.
; Syntax ........: __LOWriter_FindFormatRetrieveSetting(ByRef $atArray, $sSettingName)
; Parameters ....: $atArray             - [in/out] an array of structs. A Find Format Array of Settings to Search. Array will not be modified.
;                  $sSettingName        - a string value. The Setting name to search for.
; Return values .: Success: Object or 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $aArray not an Array.
;                  @Error 1 @Extended 2 Return 0 = $sSettingName not a String.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Search was successful, but setting was not found.
;                  @Error 0 @Extended 1 Return Object = Success. Setting found, returning requested setting Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_FindFormatRetrieveSetting(ByRef $atArray, $sSettingName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsArray($atArray) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sSettingName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	For $i = 0 To UBound($atArray) - 1
		If $atArray[$i].Name() = $sSettingName Then Return SetError($__LO_STATUS_SUCCESS, 1, $atArray[$i])
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_FindFormatRetrieveSetting

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_FooterBorder
; Description ...: Header Border Setting Internal function.
; Syntax ........: __LOWriter_FooterBorder(ByRef $oObj, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
; Parameters ....: $oObj                - [in/out] an object. Footer Object.
;                  $bWid                - a boolean value. If True the calling function is for setting Border Line Width.
;                  $bSty                - a boolean value. If True the calling function is for setting Border Line Style.
;                  $bCol                - a boolean value. If True the calling function is for setting Border Line Color.
;                  $iTop                - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iBottom             - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iLeft               - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iRight              - an integer value. See Border Style, Width, and Color functions for possible values.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj Variable not Object type variable.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style/Color when Top Border width not set.
;                  @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border style/Color when Bottom Border width not set.
;                  @Error 4 @Extended 3 Return 0 = Cannot set Left Border style/Color when Left Border width not set.
;                  @Error 4 @Extended 4 Return 0 = Cannot set Right Border style/Color when Right Border width not set.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with all other parameters set to Null keyword, and $bWid, or $bSty, or $bCol set to true to get the corresponding current settings.
;                  All distance values are set in Micrometers. Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_FooterBorder(ByRef $oObj, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aiBorder[4]
	Local $tBL2

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (($bWid + $bSty + $bCol) <> 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; If more than one Boolean is true = error

	If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then
		If $bWid Then
			__LO_ArrayFill($aiBorder, $oObj.FooterTopBorder.LineWidth(), $oObj.FooterBottomBorder.LineWidth(), _
					$oObj.FooterLeftBorder.LineWidth(), $oObj.FooterRightBorder.LineWidth())

		ElseIf $bSty Then
			__LO_ArrayFill($aiBorder, $oObj.FooterTopBorder.LineStyle(), $oObj.FooterBottomBorder.LineStyle(), _
					$oObj.FooterLeftBorder.LineStyle(), $oObj.FooterRightBorder.LineStyle())

		ElseIf $bCol Then
			__LO_ArrayFill($aiBorder, $oObj.FooterTopBorder.Color(), $oObj.FooterBottomBorder.Color(), $oObj.FooterLeftBorder.Color(), _
					$oObj.FooterRightBorder.Color())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiBorder)
	EndIf

	$tBL2 = __LO_CreateStruct("com.sun.star.table.BorderLine2")
	If Not IsObj($tBL2) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If $iTop <> Null Then
		If Not $bWid And ($oObj.FooterTopBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0) ; If Width not set, cant set color or style.

		; Top Line
		$tBL2.LineWidth = ($bWid) ? ($iTop) : ($oObj.FooterTopBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iTop) : ($oObj.FooterTopBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iTop) : ($oObj.FooterTopBorder.Color()) ; copy Color over to new size structure
		$oObj.FooterTopBorder = $tBL2
	EndIf

	If $iBottom <> Null Then
		If Not $bWid And ($oObj.FooterBottomBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 2, 0) ; If Width not set, cant set color or style.

		; Bottom Line
		$tBL2.LineWidth = ($bWid) ? ($iBottom) : ($oObj.FooterBottomBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iBottom) : ($oObj.FooterBottomBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iBottom) : ($oObj.FooterBottomBorder.Color()) ; copy Color over to new size structure
		$oObj.FooterBottomBorder = $tBL2
	EndIf

	If $iLeft <> Null Then
		If Not $bWid And ($oObj.FooterLeftBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 3, 0) ; If Width not set, cant set color or style.

		; Left Line
		$tBL2.LineWidth = ($bWid) ? ($iLeft) : ($oObj.FooterLeftBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iLeft) : ($oObj.FooterLeftBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iLeft) : ($oObj.FooterLeftBorder.Color()) ; copy Color over to new size structure
		$oObj.FooterLeftBorder = $tBL2
	EndIf

	If $iRight <> Null Then
		If Not $bWid And ($oObj.FooterRightBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 4, 0) ; If Width not set, cant set color or style.

		; Right Line
		$tBL2.LineWidth = ($bWid) ? ($iRight) : ($oObj.FooterRightBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iRight) : ($oObj.FooterRightBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iRight) : ($oObj.FooterRightBorder.Color()) ; copy Color over to new size structure
		$oObj.FooterRightBorder = $tBL2
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_FooterBorder

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_FormConGetObj
; Description ...: Returns the Shape Object for a Control.
; Syntax ........: __LOWriter_FormConGetObj($oControl, $iControlType)
; Parameters ....: $oControl            - an object. A Control Object rather than the Shape Object.
;                  $iControlType        - an integer value. The Shape type being called and looked for. See Constants $LOW_FORM_CON_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oControl not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iControlType not an Integer, less than 0 or greater than 18. See Constants $LOW_FORM_CON_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Parent of Control.
;                  @Error 3 @Extended 2 Return 0 = Failed to identify parent Document.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Document Draw Page.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve shape object.
;                  @Error 3 @Extended 5 Return 0 = Failed to identify control type.
;                  @Error 3 @Extended 6 Return 0 = Failed to identify control's parent Form.
;                  @Error 3 @Extended 7 Return 0 = Failed to identify control's shape container.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning Shape container for called Control.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: In This UDF I have provided Shape Objects for the Controls as some properties are found only in the Shape container, rather than the Control Object. In Some cases I need to Identify the Shape container from a Control Object.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_FormConGetObj($oControl, $iControlType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oParent, $oShapes, $oShape, $oDoc
	Local $aoControls[0][2]
	Local $iCount = 0

	If Not IsObj($oControl) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iControlType, $LOW_FORM_CON_TYPE_CHECK_BOX, $LOW_FORM_CON_TYPE_TIME_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oParent = $oControl.Parent()
	If Not IsObj($oParent) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $oParent.supportsService("com.sun.star.form.component.Form") Then ; Get only controls in the form.

		$oDoc = $oParent ; Identify the parent document.

		Do
			$oDoc = $oDoc.getParent()
			If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
		Until $oDoc.supportsService("com.sun.star.text.TextDocument")

		$oShapes = $oDoc.DrawPage()
		If Not IsObj($oShapes) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		ReDim $aoControls[$oShapes.Count()][2]

		For $i = 0 To $oShapes.Count() - 1
			$oShape = $oShapes.getByIndex($i)
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

			If $oShape.supportsService("com.sun.star.drawing.ControlShape") And ($oShape.Control.Parent() = $oParent) Then ; If shape is a single control, and is contained in the form.

				$aoControls[$iCount][0] = $oShape
				$aoControls[$iCount][1] = __LOWriter_FormConIdentify($oShape)
				If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

				$iCount += 1

			ElseIf $oShape.supportsService("com.sun.star.drawing.GroupShape") And ($oShape.getByIndex(0).Control.Parent() = $oParent) Then ; If shape is a group control, and the first control contained in it is contained in the form.
				$aoControls[$iCount][0] = $oShape
				$aoControls[$iCount][1] = $LOW_FORM_CON_TYPE_GROUP_BOX

				$iCount += 1
			EndIf
			Sleep((IsInt(($i / $__LOWCONST_SLEEP_DIV))) ? (10) : (0))
		Next

	Else

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)
	EndIf

	ReDim $aoControls[$iCount][2]

	For $i = 0 To UBound($aoControls) - 1
		If ($aoControls[$i][1] = $iControlType) And ($aoControls[$i][0].Control() = $oControl) Then Return SetError($__LO_STATUS_SUCCESS, 0, $aoControls[$i][0])
		Sleep((IsInt(($i / $__LOWCONST_SLEEP_DIV))) ? (10) : (0))
	Next

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)
EndFunc   ;==>__LOWriter_FormConGetObj

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_FormConIdentify
; Description ...: Identify the type of Control being called, or return the Service name of a control type.
; Syntax ........: __LOWriter_FormConIdentify($oObj[, $iControlType = Null])
; Parameters ....: $oObj                - an object. A Control object returned by a previous _LOWriter_FormConInsert, _LOWriter_FormConTableConColumnAdd or _LOWriter_FormConsGetList function.
;                  $iControlType        - [optional] an integer value (1-524288). Default is Null. The Control Type Constant. See Constants $LOW_FORM_CON_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: Integer or String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object, and $iControlType not an Integer, less than 1 or greater than 524288. See Constants $LOW_FORM_CON_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Control, or return requested Service name.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Returning Constant value for Control type.
;                  @Error 0 @Extended 1 Return String = Success. Returning requested Control type's service name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_FormConIdentify($oObj, $iControlType = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avControls[20][2] = [["com.sun.star.form.component.CheckBox", $LOW_FORM_CON_TYPE_CHECK_BOX], ["com.sun.star.form.component.ComboBox", $LOW_FORM_CON_TYPE_COMBO_BOX], _
			["com.sun.star.form.component.CurrencyField", $LOW_FORM_CON_TYPE_CURRENCY_FIELD], ["com.sun.star.form.component.DateField", $LOW_FORM_CON_TYPE_DATE_FIELD], _
			["com.sun.star.form.component.FileControl", $LOW_FORM_CON_TYPE_FILE_SELECTION], ["com.sun.star.form.component.FormattedField", $LOW_FORM_CON_TYPE_FORMATTED_FIELD], _
			["com.sun.star.form.component.GroupBox", $LOW_FORM_CON_TYPE_GROUP_BOX], ["com.sun.star.form.component.GroupBox", $LOW_FORM_CON_TYPE_GROUPED_CONTROL], _ ; This will never match as the previous group box will be tested first. This is added here for completeness of the constants and for creating a grouped control. com.sun.star.drawing.GroupShape
			["com.sun.star.form.component.ImageButton", $LOW_FORM_CON_TYPE_IMAGE_BUTTON], ["com.sun.star.form.component.DatabaseImageControl", $LOW_FORM_CON_TYPE_IMAGE_CONTROL], _ ; This seems to be used instead of this--> "com.sun.star.form.control.ImageControl"
			["com.sun.star.form.component.FixedText", $LOW_FORM_CON_TYPE_LABEL], ["com.sun.star.form.component.ListBox", $LOW_FORM_CON_TYPE_LIST_BOX], _
			["com.sun.star.form.component.NavigationToolBar", $LOW_FORM_CON_TYPE_NAV_BAR], ["com.sun.star.form.component.NumericField", $LOW_FORM_CON_TYPE_NUMERIC_FIELD], _
			["com.sun.star.form.component.RadioButton", $LOW_FORM_CON_TYPE_OPTION_BUTTON], ["com.sun.star.form.component.PatternField", $LOW_FORM_CON_TYPE_PATTERN_FIELD], _
			["com.sun.star.form.component.CommandButton", $LOW_FORM_CON_TYPE_PUSH_BUTTON], ["com.sun.star.form.component.GridControl", $LOW_FORM_CON_TYPE_TABLE_CONTROL], _
			["com.sun.star.form.component.TextField", $LOW_FORM_CON_TYPE_TEXT_BOX], ["com.sun.star.form.component.TimeField", $LOW_FORM_CON_TYPE_TIME_FIELD]]

	If Not IsObj($oObj) And Not __LO_IntIsBetween($iControlType, $LOW_FORM_CON_TYPE_CHECK_BOX, $LOW_FORM_CON_TYPE_TIME_FIELD) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If IsObj($oObj) Then
		If $oObj.Parent.supportsService("com.sun.star.form.component.GridControl") Then ; TableControl column, these controls aren't housed in a shape.
			For $i = 0 To UBound($avControls) - 1
				If ($oObj.ColumnServiceName() = $avControls[$i][0]) Then Return SetError($__LO_STATUS_SUCCESS, 0, $avControls[$i][1])
			Next

		ElseIf $oObj.supportsService("com.sun.star.drawing.GroupShape") Then ; If a Group shape, it is a Grouped control.

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_FORM_CON_TYPE_GROUPED_CONTROL)

		Else     ; Normal control housed in a shape.
			For $i = 0 To UBound($avControls) - 1
				If $oObj.Control.supportsService($avControls[$i][0]) Then Return SetError($__LO_STATUS_SUCCESS, 0, $avControls[$i][1])
			Next
		EndIf

	ElseIf IsInt($iControlType) Then
		For $i = 0 To UBound($avControls) - 1
			If ($avControls[$i][1] = $iControlType) Then Return SetError($__LO_STATUS_SUCCESS, 1, $avControls[$i][0])
		Next
	EndIf

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
EndFunc   ;==>__LOWriter_FormConIdentify

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_FormConSetGetFontDesc
; Description ...: Set or Retrieve a Control's Font values.
; Syntax ........: __LOWriter_FormConSetGetFontDesc(ByRef $oControl[, $mFontDesc = Null])
; Parameters ....: $oControl            - [in/out] an object. A Control object returned by a previous _LOWriter_FormConInsert or _LOWriter_FormConsGetList function.
;                  $mFontDesc           - [optional] a map. Default is Null. A Font descriptor Map returned by a previous _LOWriter_FontDescCreate or _LOWriter_FontDescEdit function.
; Return values .: Success: 1 or Map
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oControl not an Object.
;                  @Error 1 @Extended 2 Return 0 = $mFontDesc not a Map.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting Font Name.
;                  |                               2 = Error setting Font Weight.
;                  |                               4 = Error setting Font Posture.
;                  |                               8 = Error setting Font Size.
;                  |                               16 = Error setting Font Color.
;                  |                               32 = Error setting Font Underline Style.
;                  |                               64 = Error setting Font Underline Color.
;                  |                               128 = Error setting Font Strikeout Style.
;                  |                               256 = Error setting Individual Word mode.
;                  |                               512 = Error setting Font Relief.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Map = Success. All optional parameters were set to Null, returning current settings as a Map.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_FormConSetGetFontDesc(ByRef $oControl, $mFontDesc = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $mControlFontDesc[]

	If Not IsObj($oControl) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($mFontDesc) Then
		$mControlFontDesc.CharFontName = $oControl.CharFontName()
		$mControlFontDesc.CharWeight = $oControl.CharWeight()
		$mControlFontDesc.CharPosture = $oControl.CharPosture()
		$mControlFontDesc.CharHeight = $oControl.CharHeight()
		$mControlFontDesc.CharColor = $oControl.CharColor()
		$mControlFontDesc.CharUnderline = $oControl.CharUnderline()
		$mControlFontDesc.CharUnderlineColor = $oControl.CharUnderlineColor()
		$mControlFontDesc.CharStrikeout = $oControl.CharStrikeout()
		$mControlFontDesc.CharWordMode = $oControl.CharWordMode()
		$mControlFontDesc.CharRelief = $oControl.CharRelief()

		Return SetError($__LO_STATUS_SUCCESS, 1, $mControlFontDesc)
	EndIf

	If Not IsMap($mFontDesc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oControl.CharFontName = $mFontDesc.CharFontName
	$iError = ($oControl.CharFontName() = $mFontDesc.CharFontName) ? ($iError) : (BitOR($iError, 1))

	$oControl.CharWeight = $mFontDesc.CharWeight
	$iError = (__LO_IntIsBetween($oControl.CharWeight(), $mFontDesc.CharWeight - 50, $mFontDesc.CharWeight + 50)) ? ($iError) : (BitOR($iError, 2))

	$oControl.CharPosture = $mFontDesc.CharPosture
	$iError = ($oControl.CharPosture() = $mFontDesc.CharPosture) ? ($iError) : (BitOR($iError, 4))

	$oControl.CharHeight = $mFontDesc.CharHeight
	$iError = ($oControl.CharHeight() = $mFontDesc.CharHeight) ? ($iError) : (BitOR($iError, 8))

	$oControl.CharColor = $mFontDesc.CharColor
	$iError = ($oControl.CharColor() = $mFontDesc.CharColor) ? ($iError) : (BitOR($iError, 16))

	$oControl.CharUnderline = $mFontDesc.CharUnderline
	$iError = ($oControl.CharUnderline() = $mFontDesc.CharUnderline) ? ($iError) : (BitOR($iError, 32))

	$oControl.CharUnderlineColor = $mFontDesc.CharUnderlineColor
	$iError = ($oControl.CharUnderlineColor() = $mFontDesc.CharUnderlineColor) ? ($iError) : (BitOR($iError, 64))

	$oControl.CharStrikeout = $mFontDesc.CharStrikeout
	$iError = ($oControl.CharStrikeout() = $mFontDesc.CharStrikeout) ? ($iError) : (BitOR($iError, 128))

	$oControl.CharWordMode = $mFontDesc.CharWordMode
	$iError = ($oControl.CharWordMode() = $mFontDesc.CharWordMode) ? ($iError) : (BitOR($iError, 256))

	$oControl.CharRelief = $mFontDesc.CharRelief
	$iError = ($oControl.CharRelief() = $mFontDesc.CharRelief) ? ($iError) : (BitOR($iError, 512))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_FormConSetGetFontDesc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_GetPrinterSetting
; Description ...: Internal function for retrieving Printer settings.
; Syntax ........: __LOWriter_GetPrinterSetting(ByRef $oDoc, $sSetting)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sSetting            - a string value. The setting Name.
; Return values .: Success: Variable Value.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Array of Printer setting objects.
;                  @Error 3 @Extended 2 Return 0 = Requested setting not found.
;                  --Success--
;                  @Error 0 @Extended 0 Return Variable = Success. The requested setting's value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_GetPrinterSetting(ByRef $oDoc, $sSetting)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aoPrinterProperties

	$aoPrinterProperties = $oDoc.getPrinter()
	If Not IsArray($aoPrinterProperties) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	For $i = 0 To UBound($aoPrinterProperties) - 1
		If (($aoPrinterProperties[$i].Name()) = $sSetting) Then Return SetError($__LO_STATUS_SUCCESS, 0, $aoPrinterProperties[$i].Value())
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0) ; No Matches
EndFunc   ;==>__LOWriter_GetPrinterSetting

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_GetShapeName
; Description ...: Create a Shape Name that hasn't been used yet in the document.
; Syntax ........: __LOWriter_GetShapeName(ByRef $oDoc, $sShapeName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sShapeName          - a string value. The Shape name to begin with.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sShapeName not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve DrawPage object.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Document contained no shapes, returns the Shape name with a "1" appended.
;                  @Error 0 @Extended 1 Return String = Success. Returns the unique Shape name to use.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function adds a digit after the shape name, incrementing it until that name hasn't been used yet in L.O.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_GetShapeName(ByRef $oDoc, $sShapeName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShapes

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sShapeName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oShapes = $oDoc.DrawPage()
	If Not IsObj($oShapes) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $oShapes.hasElements() Then
		For $i = 1 To $oShapes.getCount() - 1
			For $j = 0 To $oShapes.getCount() - 1
				If ($oShapes.getByIndex($j).Name() = $sShapeName & $i) Then ExitLoop

				Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
			Next

			If ($oShapes.getByIndex($j).Name() <> $sShapeName & $i) Then ExitLoop ; If no matches, exit loop with current name.
		Next

	Else

		Return SetError($__LO_STATUS_SUCCESS, 0, $sShapeName & "1") ; If Doc has no shapes, just return the name with a "1" appended.
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 1, $sShapeName & $i)
EndFunc   ;==>__LOWriter_GetShapeName

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_GradientNameInsert
; Description ...: Create and insert a new Gradient name.
; Syntax ........: __LOWriter_GradientNameInsert(ByRef $oDoc, $tGradient[, $sGradientName = "Gradient "])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $tGradient           - a dll struct value. A Gradient Structure to copy settings from.
;                  $sGradientName       - [optional] a string value. Default is "Gradient ". The Gradient name to create.
; Return values .: Success: String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $tGradient not an Object.
;                  @Error 1 @Extended 3 Return 0 = $sGradientName not a string.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.drawing.GradientTable" Object.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.awt.Gradient" or "com.sun.star.awt.Gradient2" structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error creating Gradient Name.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. A new Gradient name was created. Returning the new name as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If The Gradient name is blank, I need to create a new name and apply it. I think I could re-use an old one without problems, but I'm not sure, so to be safe, I will create a new one.
;                  If there are no names that have been already created, then I need to create and apply one before the gradient will be displayed.
;                  Else if a preset Gradient is called, I need to create its name before it can be used.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_GradientNameInsert(ByRef $oDoc, $tGradient, $sGradientName = "Gradient ")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tNewGradient
	Local $oGradTable
	Local $iCount = 1
	Local $sGradient = "com.sun.star.awt.Gradient2"

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($tGradient) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sGradientName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If Not __LO_VersionCheck(7.6) Then $sGradient = "com.sun.star.awt.Gradient"

	$oGradTable = $oDoc.createInstance("com.sun.star.drawing.GradientTable")
	If Not IsObj($oGradTable) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($sGradientName = "Gradient ") Then
		While $oGradTable.hasByName($sGradientName & $iCount)
			$iCount += 1
			Sleep((IsInt($iCount / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
		WEnd
		$sGradientName = $sGradientName & $iCount
	EndIf

	$tNewGradient = __LO_CreateStruct($sGradient)
	If Not IsObj($tNewGradient) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	; Copy the settings over from the input Style Gradient to my new one. This may not be necessary? But just in case.
	With $tNewGradient
		.Style = $tGradient.Style()
		.XOffset = $tGradient.XOffset()
		.YOffset = $tGradient.YOffset()
		.Angle = $tGradient.Angle()
		.Border = $tGradient.Border()
		.StartColor = $tGradient.StartColor()
		.EndColor = $tGradient.EndColor()
		.StartIntensity = $tGradient.StartIntensity()
		.EndIntensity = $tGradient.EndIntensity()

		If __LO_VersionCheck(7.6) Then .ColorStops = $tGradient.ColorStops()

	EndWith

	If Not $oGradTable.hasByName($sGradientName) Then
		$oGradTable.insertByName($sGradientName, $tNewGradient)
		If Not ($oGradTable.hasByName($sGradientName)) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $sGradientName)
EndFunc   ;==>__LOWriter_GradientNameInsert

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_GradientPresets
; Description ...: Set Page background Gradient to preset settings.
; Syntax ........: __LOWriter_GradientPresets(ByRef $oDoc, ByRef $oObject, ByRef $tGradient, $sGradientName[, $bFooter = False[, $bHeader = False]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oObject             - [in/out] an object. The Object to modify the Gradient settings for.
;                  $tGradient           - [in/out] an object. The Fill Gradient Object to modify the Gradient settings for.
;                  $sGradientName       - a string value. The Gradient Preset name to apply.
;                  $bFooter             - [optional] a boolean value. Default is False. If True, settings are being set for footer Fill Gradient. If both are false, settings are for The Page itself.
;                  $bHeader             - [optional] a boolean value. Default is False. If True, settings are being set for Header Fill Gradient. If both are false, settings are for The Page itself.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create a "com.sun.star.awt.ColorStop" Struct.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to create Gradient name.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. The Style Gradient settings were successfully set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_GradientPresets(ByRef $oDoc, ByRef $oObject, ByRef $tGradient, $sGradientName, $bFooter = False, $bHeader = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tColorStop, $tStopColor
	Local $atColorStop[2]

	If __LO_VersionCheck(7.6) Then
		$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
		If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$atColorStop[0] = $tColorStop

		$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
		If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$atColorStop[1] = $tColorStop
	EndIf

	Switch $sGradientName
		Case $LOW_GRAD_NAME_PASTEL_BOUQUET
			With $tGradient
				.Style = $LOW_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 300
				.Border = 0
				.StartColor = 14543051
				.EndColor = 16766935
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOW_GRAD_NAME_PASTEL_DREAM
			With $tGradient
				.Style = $LOW_GRAD_TYPE_RECT
				.StepCount = 0
				.XOffset = 50
				.YOffset = 50
				.Angle = 450
				.Border = 0
				.StartColor = 16766935
				.EndColor = 11847644
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOW_GRAD_NAME_BLUE_TOUCH
			With $tGradient
				.Style = $LOW_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 10
				.Border = 0
				.StartColor = 11847644
				.EndColor = 14608111
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOW_GRAD_NAME_BLANK_W_GRAY
			With $tGradient
				.Style = $LOW_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 900
				.Border = 75
				.StartColor = $LO_COLOR_WHITE
				.EndColor = 14540253
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOW_GRAD_NAME_SPOTTED_GRAY
			With $tGradient
				.Style = $LOW_GRAD_TYPE_RADIAL
				.StepCount = 0
				.XOffset = 50
				.YOffset = 50
				.Angle = 0
				.Border = 0
				.StartColor = 11711154
				.EndColor = 15658734
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOW_GRAD_NAME_LONDON_MIST
			With $tGradient
				.Style = $LOW_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 300
				.Border = 0
				.StartColor = 13421772
				.EndColor = 6710886
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOW_GRAD_NAME_TEAL_TO_BLUE
			With $tGradient
				.Style = $LOW_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 300
				.Border = 0
				.StartColor = 5280650
				.EndColor = 5866416
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOW_GRAD_NAME_MIDNIGHT
			With $tGradient
				.Style = $LOW_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 0
				.Border = 0
				.StartColor = $LO_COLOR_BLACK
				.EndColor = 2777241
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOW_GRAD_NAME_DEEP_OCEAN
			With $tGradient
				.Style = $LOW_GRAD_TYPE_RADIAL
				.StepCount = 0
				.XOffset = 50
				.YOffset = 50
				.Angle = 0
				.Border = 0
				.StartColor = $LO_COLOR_BLACK
				.EndColor = 7512015
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOW_GRAD_NAME_SUBMARINE
			With $tGradient
				.Style = $LOW_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 0
				.Border = 0
				.StartColor = 14543051
				.EndColor = 11847644
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOW_GRAD_NAME_GREEN_GRASS
			With $tGradient
				.Style = $LOW_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 300
				.Border = 0
				.StartColor = 16776960
				.EndColor = 8508442
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOW_GRAD_NAME_NEON_LIGHT
			With $tGradient
				.Style = $LOW_GRAD_TYPE_ELLIPTICAL
				.StepCount = 0
				.XOffset = 50
				.YOffset = 50
				.Angle = 0
				.Border = 15
				.StartColor = 1209890
				.EndColor = $LO_COLOR_WHITE
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOW_GRAD_NAME_SUNSHINE
			With $tGradient
				.Style = $LOW_GRAD_TYPE_RADIAL
				.StepCount = 0
				.XOffset = 66
				.YOffset = 33
				.Angle = 0
				.Border = 33
				.StartColor = 16760576
				.EndColor = 16776960
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOW_GRAD_NAME_PRESENT
			With $tGradient
				.Style = $LOW_GRAD_TYPE_SQUARE
				.StepCount = 0
				.XOffset = 70
				.YOffset = 60
				.Angle = 450
				.Border = 72
				.StartColor = 8468233
				.EndColor = 16728064
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOW_GRAD_NAME_MAHOGANY
			With $tGradient
				.Style = $LOW_GRAD_TYPE_SQUARE
				.StepCount = 0
				.XOffset = 50
				.YOffset = 50
				.Angle = 450
				.Border = 0
				.StartColor = $LO_COLOR_BLACK
				.EndColor = 9250846
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = $atColorStop[1]
					$tColorStop.StopOffset = 1
					$atColorStop[1] = $tColorStop
				EndIf
			EndWith

		Case $LOW_GRAD_NAME_RAINBOW
			With $tGradient
				.Style = $LOW_GRAD_TYPE_RADIAL
				.StepCount = 0
				.XOffset = 50
				.YOffset = 100
				.Angle = 0
				.Border = 0
				.StartColor = $LO_COLOR_WHITE
				.EndColor = $LO_COLOR_WHITE
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					ReDim $atColorStop[7]

					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0.2
					$atColorStop[0] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.2

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 1
					$tStopColor.Green = 0
					$tStopColor.Blue = 0
					$tColorStop.StopColor = $tStopColor

					$atColorStop[1] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.4

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 1
					$tStopColor.Green = 1
					$tStopColor.Blue = 0
					$tColorStop.StopColor = $tStopColor

					$atColorStop[2] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.5

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0
					$tStopColor.Green = 1
					$tStopColor.Blue = 0
					$tColorStop.StopColor = $tStopColor

					$atColorStop[3] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.65

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0
					$tStopColor.Green = 1
					$tStopColor.Blue = 1
					$tColorStop.StopColor = $tStopColor

					$atColorStop[4] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.8

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 1
					$tStopColor.Green = 0
					$tStopColor.Blue = 1
					$tColorStop.StopColor = $tStopColor

					$atColorStop[5] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.8
					$atColorStop[6] = $tColorStop
				EndIf
			EndWith

		Case $LOW_GRAD_NAME_SUNRISE
			With $tGradient
				.Style = $LOW_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 0
				.Border = 0
				.StartColor = 3713206
				.EndColor = 14065797
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					ReDim $atColorStop[4]

					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.5

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0.505882352941176
					$tStopColor.Green = 0.784313725490196
					$tStopColor.Blue = 0.768627450980392
					$tColorStop.StopColor = $tStopColor

					$atColorStop[1] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.75

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0.717647058823529
					$tStopColor.Green = 0.807843137254902
					$tStopColor.Blue = 0.698039215686275
					$tColorStop.StopColor = $tStopColor

					$atColorStop[2] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 1
					$atColorStop[3] = $tColorStop
				EndIf
			EndWith

		Case $LOW_GRAD_NAME_SUNDOWN
			With $tGradient
				.Style = $LOW_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 0
				.Border = 0
				.StartColor = 985943
				.EndColor = 16759664
				.StartIntensity = 100
				.EndIntensity = 100

				If __LO_VersionCheck(7.6) Then
					ReDim $atColorStop[5]

					$tColorStop = $atColorStop[0]
					$tColorStop.StopOffset = 0
					$atColorStop[0] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.3

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0.392156862745098
					$tStopColor.Green = 0.305882352941177
					$tStopColor.Blue = 0.690196078431373
					$tColorStop.StopColor = $tStopColor

					$atColorStop[1] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.5

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0.827450980392157
					$tStopColor.Green = 0.572549019607843
					$tStopColor.Blue = 0.83921568627451
					$tColorStop.StopColor = $tStopColor

					$atColorStop[2] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 0.75

					$tStopColor = $tColorStop.StopColor()
					$tStopColor.Red = 0.996078431372549
					$tStopColor.Green = 0.733333333333333
					$tStopColor.Blue = 0.76078431372549
					$tColorStop.StopColor = $tStopColor

					$atColorStop[3] = $tColorStop

					$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
					If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tColorStop.StopOffset = 1
					$atColorStop[4] = $tColorStop
				EndIf
			EndWith

		Case Else ; Custom Gradient Name
			__LOWriter_GradientNameInsert($oDoc, $tGradient, $sGradientName)
			If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			If $bFooter Then
				$oObject.FooterFillGradientName = $sGradientName

			ElseIf $bHeader Then
				$oObject.HeaderFillGradientName = $sGradientName

			Else
				$oObject.FillGradientName = $sGradientName
			EndIf

			Return SetError($__LO_STATUS_SUCCESS, 0, 1)
	EndSwitch

	If __LO_VersionCheck(7.6) Then
		$tColorStop = $atColorStop[0] ; "Start Color" Value.

		$tStopColor = $tColorStop.StopColor()

		$tStopColor.Red = (BitAND(BitShift($tGradient.StartColor(), 16), 0xff) / 255)
		$tStopColor.Green = (BitAND(BitShift($tGradient.StartColor(), 8), 0xff) / 255)
		$tStopColor.Blue = (BitAND($tGradient.StartColor(), 0xff) / 255)

		$tColorStop.StopColor = $tStopColor

		$atColorStop[0] = $tColorStop

		$tColorStop = $atColorStop[UBound($atColorStop) - 1] ; Last element is "End Color" Value.

		$tStopColor = $tColorStop.StopColor()

		$tStopColor.Red = (BitAND(BitShift($tGradient.EndColor(), 16), 0xff) / 255)
		$tStopColor.Green = (BitAND(BitShift($tGradient.EndColor(), 8), 0xff) / 255)
		$tStopColor.Blue = (BitAND($tGradient.EndColor(), 0xff) / 255)

		$tColorStop.StopColor = $tStopColor

		$atColorStop[UBound($atColorStop) - 1] = $tColorStop

		$tGradient.ColorStops = $atColorStop
	EndIf

	__LOWriter_GradientNameInsert($oDoc, $tGradient, $sGradientName)
	If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $bFooter Then
		$oObject.FooterFillGradient = $tGradient
		$oObject.FooterFillGradientName = $sGradientName
		$oObject.FooterFillGradientStepCount = $tGradient.StepCount()

	ElseIf $bHeader Then
		$oObject.HeaderFillGradient = $tGradient
		$oObject.HeaderFillGradientName = $sGradientName
		$oObject.HeaderFillGradientStepCount = $tGradient.StepCount()

	Else
		$oObject.FillGradient = $tGradient
		$oObject.FillGradientName = $sGradientName
		$oObject.FillGradientStepCount = $tGradient.StepCount()
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_GradientPresets

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_HeaderBorder
; Description ...: Header Border Setting Internal function.
; Syntax ........: __LOWriter_HeaderBorder(ByRef $oObj, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
; Parameters ....: $oObj                - [in/out] an object. A Header object.
;                  $bWid                - a boolean value. If True the calling function is for setting Border Line Width.
;                  $bSty                - a boolean value. If True the calling function is for setting Border Line Style.
;                  $bCol                - a boolean value. If True the calling function is for setting Border Line Color.
;                  $iTop                - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iBottom             - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iLeft               - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iRight              - an integer value. See Border Style, Width, and Color functions for possible values.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style/Color when Top Border width not set.
;                  @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border style/Color when Bottom Border width not set.
;                  @Error 4 @Extended 3 Return 0 = Cannot set Left Border style/Color when Left Border width not set.
;                  @Error 4 @Extended 4 Return 0 = Cannot set Right Border style/Color when Right Border width not set.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with all other parameters set to Null keyword, and $bWid, or $bSty, or $bCol set to true to get the corresponding current settings.
;                  All distance values are set in Micrometers.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_HeaderBorder(ByRef $oObj, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tBL2
	Local $aiBorder[4]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (($bWid + $bSty + $bCol) <> 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; If more than one Boolean is true = error

	If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then
		If $bWid Then
			__LO_ArrayFill($aiBorder, $oObj.HeaderTopBorder.LineWidth(), $oObj.HeaderBottomBorder.LineWidth(), _
					$oObj.HeaderLeftBorder.LineWidth(), $oObj.HeaderRightBorder.LineWidth())

		ElseIf $bSty Then
			__LO_ArrayFill($aiBorder, $oObj.HeaderTopBorder.LineStyle(), $oObj.HeaderBottomBorder.LineStyle(), _
					$oObj.HeaderLeftBorder.LineStyle(), $oObj.HeaderRightBorder.LineStyle())

		ElseIf $bCol Then
			__LO_ArrayFill($aiBorder, $oObj.HeaderTopBorder.Color(), $oObj.HeaderBottomBorder.Color(), $oObj.HeaderLeftBorder.Color(), _
					$oObj.HeaderRightBorder.Color())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiBorder)
	EndIf

	$tBL2 = __LO_CreateStruct("com.sun.star.table.BorderLine2")
	If Not IsObj($tBL2) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If $iTop <> Null Then
		If Not $bWid And ($oObj.HeaderTopBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0) ; If Width not set, cant set color or style.

		; Top Line
		$tBL2.LineWidth = ($bWid) ? ($iTop) : ($oObj.HeaderTopBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iTop) : ($oObj.HeaderTopBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iTop) : ($oObj.HeaderTopBorder.Color()) ; copy Color over to new size structure
		$oObj.HeaderTopBorder = $tBL2
	EndIf

	If $iBottom <> Null Then
		If Not $bWid And ($oObj.HeaderBottomBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 2, 0) ; If Width not set, cant set color or style.

		; Bottom Line
		$tBL2.LineWidth = ($bWid) ? ($iBottom) : ($oObj.HeaderBottomBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iBottom) : ($oObj.HeaderBottomBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iBottom) : ($oObj.HeaderBottomBorder.Color()) ; copy Color over to new size structure
		$oObj.HeaderBottomBorder = $tBL2
	EndIf

	If $iLeft <> Null Then
		If Not $bWid And ($oObj.HeaderLeftBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 3, 0) ; If Width not set, cant set color or style.

		; Left Line
		$tBL2.LineWidth = ($bWid) ? ($iLeft) : ($oObj.HeaderLeftBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iLeft) : ($oObj.HeaderLeftBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iLeft) : ($oObj.HeaderLeftBorder.Color()) ; copy Color over to new size structure
		$oObj.HeaderLeftBorder = $tBL2
	EndIf

	If $iRight <> Null Then
		If Not $bWid And ($oObj.HeaderRightBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 4, 0) ; If Width not set, cant set color or style.

		; Right Line
		$tBL2.LineWidth = ($bWid) ? ($iRight) : ($oObj.HeaderRightBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iRight) : ($oObj.HeaderRightBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iRight) : ($oObj.HeaderRightBorder.Color()) ; copy Color over to new size structure
		$oObj.HeaderRightBorder = $tBL2
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_HeaderBorder

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ImageGetSuggestedSize
; Description ...: Return a suggested image width/height based on an image's original size.
; Syntax ........: __LOWriter_ImageGetSuggestedSize(ByRef $oGraphic, $oPageStyle)
; Parameters ....: $oGraphic            - [in/out] an object. A graphic Object returned from a queryGraphicDescriptor call.
;                  $oPageStyle          - an object. A Page Style object returned by a previous _LOWriter_PageStyleGetObj function.
; Return values .: Success: Structure.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oGraphic not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oPageStyle not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error calculating Width and Height.
;                  --Success--
;                  @Error 0 @Extended 0 Return Structure = Successfully calculated suggested Width and Height, returning size Structure.
; Author ........: Andrew Pitonyak ("Useful Macro Information For OpenOffice.org", Page 62, listing 5.28)
; Modified ......: donnyh13, converted code from L.O. Basic to AutoIt. Added a max W/H based on current page size.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ImageGetSuggestedSize($oGraphic, $oPageStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSize
	Local $iMaxH, $iMaxW

	If Not IsObj($oGraphic) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	; Retrieve the Current Page Style's height minus top/bottom margins
	$iMaxH = Int($oPageStyle.Height() - $oPageStyle.LeftMargin() - $oPageStyle.RightMargin())
	If ($iMaxH = 0) Then $iMaxH = 24130 ; If error or is equal to 0, then set to 9.5 Inches in Micrometers

	; Retrieve the Current Page Style's width minus left/right margins
	$iMaxW = Int($oPageStyle.Width() - $oPageStyle.TopMargin() - $oPageStyle.BottomMargin())
	If ($iMaxW = 0) Then $iMaxW = 17145 ; If error or is equal to 0, then set to 6.75 Inches in Micrometers.

	$oSize = $oGraphic.Size100thMM()

	If ($oSize.Height = 0) Or ($oSize.Width = 0) Then
		; 2540 Micrometers per Inch, 1440 TWIPS per inch
		$oSize.Height = Int($oGraphic.SizePixel.Height * 2540 * _WinAPI_TwipsPerPixelY() / 1440)
		$oSize.Width = Int($oGraphic.SizePixel.Width * 2540 * _WinAPI_TwipsPerPixelX() / 1440)
	EndIf

	If ($oSize.Height = 0) Or ($oSize.Width = 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($oSize.Width() > $iMaxW) Then
		$oSize.Height = Int($oSize.Height * $iMaxW / $oSize.Width())
		$oSize.Width = $iMaxW
	EndIf

	If ($oSize.Height() > $iMaxH) Then
		$oSize.Width = Int($oSize.Width() * $iMaxH / $oSize.Height)
		$oSize.Height = $iMaxH
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSize)
EndFunc   ;==>__LOWriter_ImageGetSuggestedSize

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Internal_CursorGetDataType
; Description ...: Get what type of Text data the cursor object is currently in. Internal version of CursorGetDataType.
; Syntax ........: __LOWriter_Internal_CursorGetDataType(ByRef $oDoc, ByRef $oCursor[, $bReturnObject = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions.
;                  $bReturnObject       - [optional] a boolean value. Default is False. If True, return the object used for creating a Text Object etc.
; Return values .: Success: Object or Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 3 Return 0 = $bReturnObject not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $oCursor is a Table Cursor, or a View Cursor with table cells selected. Can't get data type from these types of Cursors.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving TextFrame Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving TextCell Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve Footnotes Object for document.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Endnotes Object for document.
;                  @Error 3 @Extended 5 Return 0 = Unable to identify Foot/EndNote.
;                  @Error 3 @Extended 6 Return 0 = Cursor in unknown DataType
;                  --Success--
;                  @Error 0 @Extended ? Return Object = Success, If $bReturnObject is True, returns an object used for creating a Text Object, @Extended is set to one of the constants, $LOW_CURDATA_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 0 @Extended 0 Return Integer = Success, If $bReturnObject is False, Return value will be one of constants, $LOW_CURDATA_* as defined in LibreOfficeWriter_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Returns what type of cursor, such as a TextTable, Footnote etc.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_Internal_CursorGetDataType(ByRef $oDoc, ByRef $oCursor, $bReturnObject = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oEndNotes, $oFootNotes, $oReturnObject

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bReturnObject) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If (($oCursor.ImplementationName()) = "SwXTextTableCursor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0) ; Can't get data type from Table Cursor.

	Switch $oCursor.Text.getImplementationName()
		Case "SwXBodyText"
			$oReturnObject = $oDoc

			Return ($bReturnObject) ? (SetError($__LO_STATUS_SUCCESS, $LOW_CURDATA_BODY_TEXT, $oReturnObject)) : (SetError($__LO_STATUS_SUCCESS, 0, $LOW_CURDATA_BODY_TEXT))

		Case "SwXTextFrame"
			$oReturnObject = $oDoc.TextFrames.getByName($oCursor.TextFrame.Name)
			If Not IsObj($oReturnObject) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			Return ($bReturnObject) ? (SetError($__LO_STATUS_SUCCESS, $LOW_CURDATA_FRAME, $oReturnObject)) : (SetError($__LO_STATUS_SUCCESS, 0, $LOW_CURDATA_FRAME))

		Case "SwXCell"
			$oReturnObject = $oDoc.TextTables.getByName($oCursor.TextTable.Name)
			If Not IsObj($oReturnObject) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			Return ($bReturnObject) ? (SetError($__LO_STATUS_SUCCESS, $LOW_CURDATA_CELL, $oReturnObject)) : (SetError($__LO_STATUS_SUCCESS, 0, $LOW_CURDATA_CELL))

		Case "SwXHeadFootText"
			$oReturnObject = $oCursor

			Return ($bReturnObject) ? (SetError($__LO_STATUS_SUCCESS, $LOW_CURDATA_HEADER_FOOTER, $oReturnObject)) : (SetError($__LO_STATUS_SUCCESS, 0, $LOW_CURDATA_HEADER_FOOTER))

		Case "SwXFootnote"
			$oFootNotes = $oDoc.getFootnotes()
			If Not IsObj($oFootNotes) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			For $i = 0 To $oFootNotes.getCount() - 1
				If ($oFootNotes.getByIndex($i).ReferenceId() = $oCursor.Text.ReferenceId()) And _
						($oFootNotes.getByIndex($i).Text() = $oCursor.Text()) Then Return ($bReturnObject) ? (SetError($__LO_STATUS_SUCCESS, $LOW_CURDATA_FOOTNOTE, $oFootNotes.getByIndex($i))) : (SetError($__LO_STATUS_SUCCESS, 0, $LOW_CURDATA_FOOTNOTE))
			Next

			$oEndNotes = $oDoc.getEndnotes()     ; Not found in Footnotes, check Endnotes.
			If Not IsObj($oEndNotes) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

			For $i = 0 To $oEndNotes.getCount() - 1
				If ($oEndNotes.getByIndex($i).ReferenceId() = $oCursor.Text.ReferenceId()) And _
						($oEndNotes.getByIndex($i).Text() = $oCursor.Text()) Then Return ($bReturnObject) ? (SetError($__LO_STATUS_SUCCESS, $LOW_CURDATA_ENDNOTE, $oEndNotes.getByIndex($i))) : (SetError($__LO_STATUS_SUCCESS, 0, $LOW_CURDATA_ENDNOTE))
			Next

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0) ; no matches
	EndSwitch

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)         ; unknown data type.
EndFunc   ;==>__LOWriter_Internal_CursorGetDataType

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Internal_CursorGetType
; Description ...: Get what type of cursor the object is. Internal version of CursorGetType.
; Syntax ........: __LOWriter_Internal_CursorGetType(ByRef $oCursor)
; Parameters ....: $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Unknown Cursor type.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success, Return value will be one of the constants, $LOW_CURTYPE_* as defined in LibreOfficeWriter_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Returns what type of cursor the input Object is, such as a Table Cursor, Text Cursor or a View Cursor. Can also be a Paragraph or Text Portion.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_Internal_CursorGetType(ByRef $oCursor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Switch $oCursor.getImplementationName()
		Case "SwXTextViewCursor"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_CURTYPE_VIEW_CURSOR)

		Case "SwXTextTableCursor"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_CURTYPE_TABLE_CURSOR)

		Case "SwXTextCursor", "SvxUnoTextCursor" ; SvxUnoTextCursor is a Text Cursor created in a TextBox Form Control.

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_CURTYPE_TEXT_CURSOR)

		Case "SwXParagraph"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_CURTYPE_PARAGRAPH)

		Case "SwXTextPortion"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_CURTYPE_TEXT_PORTION)

		Case Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; unknown Cursor type.
	EndSwitch
EndFunc   ;==>__LOWriter_Internal_CursorGetType

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_InternalComErrorHandler
; Description ...: ComError Handler
; Syntax ........: __LOWriter_InternalComErrorHandler(ByRef $oComError)
; Parameters ....: $oComError           - [in/out] an object. The Com Error Object passed by Autoit.Error.
; Return values .: None
; Author ........: mLipok
; Modified ......: donnyh13 - Added parameters option. Also added MsgBox & ConsoleWrite options.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_InternalComErrorHandler(ByRef $oComError)
	; If not defined ComError_UserFunction then this function does nothing, in which case you can only check @error / @extended after suspect functions.
	Local $avUserFunction = _LOWriter_ComError_UserFunction(Default)
	Local $vUserFunction, $avUserParams[2] = ["CallArgArray", $oComError]

	If IsArray($avUserFunction) Then
		$vUserFunction = $avUserFunction[0]
		ReDim $avUserParams[UBound($avUserFunction) + 1]
		For $i = 1 To UBound($avUserFunction) - 1
			$avUserParams[$i + 1] = $avUserFunction[$i]
		Next

	Else
		$vUserFunction = $avUserFunction
	EndIf
	If IsFunc($vUserFunction) Then
		Switch $vUserFunction
			Case ConsoleWrite
				ConsoleWrite("!--COM Error-Begin--" & @CRLF & _
						"Number: 0x" & Hex($oComError.number, 8) & @CRLF & _
						"WinDescription: " & $oComError.windescription & @CRLF & _
						"Source: " & $oComError.source & @CRLF & _
						"Error Description: " & $oComError.description & @CRLF & _
						"HelpFile: " & $oComError.helpfile & @CRLF & _
						"HelpContext: " & $oComError.helpcontext & @CRLF & _
						"LastDLLError: " & $oComError.lastdllerror & @CRLF & _
						"At line: " & $oComError.scriptline & @CRLF & _
						"!--COM-Error-End--" & @CRLF)

			Case MsgBox
				MsgBox(0, "COM Error", "Number: 0x" & Hex($oComError.number, 8) & @CRLF & _
						"WinDescription: " & $oComError.windescription & @CRLF & _
						"Source: " & $oComError.source & @CRLF & _
						"Error Description: " & $oComError.description & @CRLF & _
						"HelpFile: " & $oComError.helpfile & @CRLF & _
						"HelpContext: " & $oComError.helpcontext & @CRLF & _
						"LastDLLError: " & $oComError.lastdllerror & @CRLF & _
						"At line: " & $oComError.scriptline)

			Case Else
				Call($vUserFunction, $avUserParams)
		EndSwitch
	EndIf
EndFunc   ;==>__LOWriter_InternalComErrorHandler

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_IsCellRange
; Description ...: Check whether a Cell Object is a single cell or a Cell Range.
; Syntax ........: __LOWriter_IsCellRange(ByRef $oCell)
; Parameters ....: $oCell               - [in/out] an object. A Table Cell or Cell Range Object returned by a previous _LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, or _LOWriter_TableGetCellObjByPosition function.
; Return values .: Success: Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = If the cell object is a Cell Range, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_IsCellRange(ByRef $oCell)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return ($oCell.supportsService("com.sun.star.text.CellRange")) ? (SetError($__LO_STATUS_SUCCESS, 0, True)) : (SetError($__LO_STATUS_SUCCESS, 0, False))
EndFunc   ;==>__LOWriter_IsCellRange

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_IsTableInDoc
; Description ...: Check if Table is inserted in a Document or has only been created and not inserted.
; Syntax ........: __LOWriter_IsTableInDoc(ByRef $oTable)
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableInsert, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Table cell names.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = If True, Table is inserted into the document, If false Table has been created with _LOWriter_TableCreate but not inserted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_IsTableInDoc(ByRef $oTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aTableNames

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$aTableNames = $oTable.getCellNames()
	If Not IsArray($aTableNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, (UBound($aTableNames)) ? (True) : (False)) ; If 0 elements = False = not in doc.
EndFunc   ;==>__LOWriter_IsTableInDoc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_NumRuleCreateMap
; Description ...: Creates a map with values for each setting location in the Array.
; Syntax ........: __LOWriter_NumRuleCreateMap(ByRef $atNumLevel)
; Parameters ....: $atNumLevel          - [in/out] an array of dll structs. An Array of Property Structures for a Numbering Rule.
; Return values .: Success: Map
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $atNumLevel not an Array.
;                  --Success--
;                  @Error 0 @Extended 0 Return Map = Success. Returning a Map containing the location in the array for each setting.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_NumRuleCreateMap(ByRef $atNumLevel)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $mNumLevel[]

	If Not IsArray($atNumLevel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	For $i = 0 To UBound($atNumLevel) - 1
		$mNumLevel[$atNumLevel[$i].Name()] = $i
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, $mNumLevel)
EndFunc   ;==>__LOWriter_NumRuleCreateMap

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_NumStyleCreateScript
; Description ...: Part of the Numbering Style Modification workaround, creates a Macro in a document.
; Syntax ........: __LOWriter_NumStyleCreateScript(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Standard Macro Library.
;                  @Error 3 @Extended 2 Return 0 = Error creating Macro in Document.
;                  @Error 3 @Extended 3 Return 0 = Error retrieving Script Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Function successfully created the Macro in Document. Returning Script Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_NumStyleCreateScript(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sNumStyleScript = "Function ReplaceByIndex(oNumRules As Object, iIndex%, vSettings As Variant)" & @CRLF & _
			"oNumRules.replaceByIndex(iIndex,vSettings)" & @CRLF & _
			"ReplaceByIndex = oNumRules" & @CRLF & _
			"End Function"
	Local $oStandardLibrary, $oScript

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	; Retrieving the BasicLibrary.Standard Object fails when using a newly opened document, I found a workaround by updating the
	; following setting.
	$oDoc.BasicLibraries.VBACompatibilityMode = $oDoc.BasicLibraries.VBACompatibilityMode()

	$oStandardLibrary = $oDoc.BasicLibraries.Standard()
	If Not IsObj($oStandardLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $oStandardLibrary.hasByName("AU3LibreOffice_UDF_Macros") Then $oStandardLibrary.removeByName("AU3LibreOffice_UDF_Macros")

	$oStandardLibrary.insertByName("AU3LibreOffice_UDF_Macros", $sNumStyleScript)
	If Not $oStandardLibrary.hasByName("AU3LibreOffice_UDF_Macros") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oScript = $oDoc.getScriptProvider().getScript("vnd.sun.star.script:Standard.AU3LibreOffice_UDF_Macros.ReplaceByIndex?language=Basic&location=document")
	If Not IsObj($oScript) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oScript)
EndFunc   ;==>__LOWriter_NumStyleCreateScript

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_NumStyleDeleteScript
; Description ...: Part of the Numbering Style Modification workaround, deletes a Macro in a document.
; Syntax ........: __LOWriter_NumStyleDeleteScript(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Standard Macro Library.
;                  @Error 3 @Extended 2 Return 0 = Error deleting Macro.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Function successfully deleted the Macro in Document.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_NumStyleDeleteScript(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oStandardLibrary

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	; Retrieving the BasicLibrary.Standard Object fails when using a newly opened document, I found a workaround by updating the
	; following setting.
	$oDoc.BasicLibraries.VBACompatibilityMode = $oDoc.BasicLibraries.VBACompatibilityMode()

	$oStandardLibrary = $oDoc.BasicLibraries.Standard()
	If Not IsObj($oStandardLibrary) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If $oStandardLibrary.hasByName("AU3LibreOffice_UDF_Macros") Then $oStandardLibrary.removeByName("AU3LibreOffice_UDF_Macros")

	If $oStandardLibrary.hasByName("AU3LibreOffice_UDF_Macros") Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_NumStyleDeleteScript

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_NumStyleInitiateDocument
; Description ...: Part of the work around method for modifying Numbering Style settings.
; Syntax ........: __LOWriter_NumStyleInitiateDocument()
; Parameters ....: None
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.ServiceManager" Object.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.frame.Desktop" Object.
;                  @Error 2 @Extended 3 Return 0 = Error Creating document.
;                  @Error 2 @Extended 4 Return 0 = Error retrieving standard Macro Library Object from Document.
;                  @Error 2 @Extended 5 Return 0 = Error creating AU3LibreOffice_UDF_Macros Module in document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting Hidden
;                  |                               2 = Error setting MacroExecutionMode
;                  |                               4 = Error setting ReadOnly
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. The Numbering Style Modification Document was successfully created.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_NumStyleInitiateDocument()
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $iMacroExecMode_ALWAYS_EXECUTE_NO_WARN = 4, $iURLFrameCreate = 8 ; Frame will be created if not found
	Local $iError = 0
	Local $oNumStyleDoc, $oServiceManager, $oDesktop
	Local $atProperties[3]
	Local $vProperty

	$oServiceManager = __LO_ServiceManager()
	If Not IsObj($oServiceManager) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$vProperty = __LO_SetPropertyValue("Hidden", True)
	If @error Then $iError = BitOR($iError, 1)
	If Not BitAND($iError, 1) Then $atProperties[0] = $vProperty

	$vProperty = __LO_SetPropertyValue("MacroExecutionMode", $iMacroExecMode_ALWAYS_EXECUTE_NO_WARN)
	If @error Then $iError = BitOR($iError, 2)
	If Not BitAND($iError, 2) Then $atProperties[1] = $vProperty

	$vProperty = __LO_SetPropertyValue("ReadOnly", True)
	If @error Then $iError = BitOR($iError, 4)
	If Not BitAND($iError, 4) Then $atProperties[2] = $vProperty

	$oNumStyleDoc = $oDesktop.loadComponentFromURL("private:factory/swriter", "_blank", $iURLFrameCreate, $atProperties)
	If Not IsObj($oNumStyleDoc) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	__LOWriter_NumStyleCreateScript($oNumStyleDoc)
	If (@error > 0) Then Return SetError($__LO_STATUS_INIT_ERROR, 5, 0)

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $oNumStyleDoc)) : (SetError($__LO_STATUS_SUCCESS, 1, $oNumStyleDoc))
EndFunc   ;==>__LOWriter_NumStyleInitiateDocument

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_NumStyleListFormat
; Description ...: Creates a string for modifying List Format Number Style setting.
; Syntax ........: __LOWriter_NumStyleListFormat(ByRef $atNumLevel, $iLevel, $iSubLevels[, $sPrefix = Null[, $sSuffix = Null]])
; Parameters ....: $atNumLevel          - [in/out] an array of dll structs. An array of Numbering Rule settings retrieved from a Numbering Style.
;                  $iLevel              - an integer value. The Level to create the ListFormat string for
;                  $iSubLevels          - an integer value. The number of levels to go up from $iLevel.
;                  $sPrefix             - [optional] a string value. Default is Null. If Null, retrieves the current Prefix, else use the input Prefix.
;                  $sSuffix             - [optional] a string value. Default is Null. If Null, retrieves the current Suffix, else use the input Suffix.
; Return values .: Success: String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oNumRules not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iLevel not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iSubLevels not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 =  Error mapping setting values.
;                  --Success--
;                  @Error 0 @Extended 1 Return String = Success. A String used for modifying ListFormat Numbering Style setting.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_NumStyleListFormat(ByRef $atNumLevel, $iLevel, $iSubLevels, $sPrefix = Null, $sSuffix = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sListFormat = "", $sSeperator = "."
	Local $iEndLevel
	Local $mNumLevel[]

	If Not IsArray($atNumLevel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iLevel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iSubLevels) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$mNumLevel = __LOWriter_NumRuleCreateMap($atNumLevel)
	If Not IsMap($mNumLevel) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iEndLevel = ($iLevel - $iSubLevels + 1) ; Start at the called level minus any Sub-levels.

	$sPrefix = ($sPrefix = Null) ? ($atNumLevel[$mNumLevel["Prefix"]].Value()) : ($sPrefix)    ; If I'm modifying a specific level, retrieve its prefix/Suffix
	$sSuffix = ($sSuffix = Null) ? ($atNumLevel[$mNumLevel["Suffix"]].Value()) : ($sSuffix)

	For $i = $iLevel To $iEndLevel Step -1     ; Cycle Through the levels if any Sub levels are set.
		If ($i = $iEndLevel) Then $sSeperator = ""
		$sListFormat = $sSeperator & "%" & ($i + 1) & "%" & $sListFormat
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	$sListFormat = $sPrefix & $sListFormat & $sSuffix

	Return SetError($__LO_STATUS_SUCCESS, 1, $sListFormat)
EndFunc   ;==>__LOWriter_NumStyleListFormat

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_NumStyleModify
; Description ...: Internal function for modifying Numbering Style settings.
; Syntax ........: __LOWriter_NumStyleModify(ByRef $oDoc, ByRef $oNumRules, $iLevel, $atNumLevel)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function, to modify NumberingRules for.
;                  $oNumRules           - [in/out] an object. The Numbering Rules object retrieved from a Numbering Style.
;                  $iLevel              - an integer value (-1-9). The Numbering Style level to modify. -1 = all levels.
;                  $atNumLevel          - an array of dll structs. An array of Numbering Rule settings retrieved from a Numbering Style.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oNumRules not an Object.
;                  @Error 1 @Extended 3 Return 0 = $iLevel not between -1 and 9 to indicate correct level.
;                  @Error 1 @Extended 4 Return 0 = $atNumLevel not an array.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error opening new document, and inserting ReplaceByIndex Script.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving "Standard.AU3LibreOffice_UDF_Macros.ReplaceByIndex" Macro in new document.
;                  @Error 3 @Extended 3 Return 0 = Error deleting ReplaceByIndex Macro from Document.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully set the requested settings.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This works, but only with a work-around method, see inside this function for a description of why a work-around method is necessary.
;                  When a lot of settings are set, especially for all levels, this function can be a bit slow.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_NumStyleModify(ByRef $oDoc, ByRef $oNumRules, $iLevel, $atNumLevel)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oNumStyleDoc, $oScript
	Local $aDummyArray[0], $avParamArray[3]
	Local $bNumDocOpen = False

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oNumRules) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iLevel, -1, 9) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsArray($atNumLevel) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oScript = __LOWriter_NumStyleCreateScript($oDoc) ; Create my modification Script.

	If Not IsObj($oScript) Then ; If creating my Mod. Script fails, open a new document and create a script in there.
		$oNumStyleDoc = __LOWriter_NumStyleInitiateDocument()
		If Not IsObj($oNumStyleDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$oScript = $oNumStyleDoc.getScriptProvider().getScript("vnd.sun.star.script:Standard.AU3LibreOffice_UDF_Macros.ReplaceByIndex?language=Basic&location=document")
		If Not IsObj($oScript) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$bNumDocOpen = True
	EndIf

	; $oNumRules.replaceByIndex($iGetLevel, $atNumLevel); This should work but doesn't -- It would seem that the Array passed by
	; AutoIt is not recognized as an appropriate array(or Sequence) by LibreOffice, or perhaps as variable type "Any", which is
	; what LibreOffice replace by index is expecting, and consequently causes a com.sun.star.lang.IllegalArgumentException COM error.

	$avParamArray[0] = $oNumRules
	$avParamArray[1] = $iLevel
	$avParamArray[2] = $atNumLevel

	$oNumRules = $oScript.Invoke($avParamArray, $aDummyArray, $aDummyArray)

	If ($bNumDocOpen = True) Then
		$oNumStyleDoc.Close(True)

	Else
		__LOWriter_NumStyleDeleteScript($oDoc)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_NumStyleModify

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ObjRelativeSize
; Description ...: Calculate appropriate values to set Frame, Frame Style or Image Width or Height, when using relative values.
; Syntax ........: __LOWriter_ObjRelativeSize(ByRef $oDoc, ByRef $oObj[, $bRelativeWidth = False[, $bRelativeHeight = False]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oObj                - [in/out] an object. A Frame or Frame Style object returned by a previous _LOWriter_FrameStyleCreate, _LOWriter_FrameCreate, _LOWriter_FrameStyleGetObj, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function. Can also be an Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $bRelativeWidth      - [optional] a boolean value. Default is False. If True, modify Width based on relative Width percentage.
;                  $bRelativeHeight     - [optional] a boolean value. Default is False. If True, modify Height based on relative Height percentage.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 3 Return 0 = $bRelativeWidth not a boolean.
;                  @Error 1 @Extended 4 Return 0 = $bRelativeHeight not a boolean.
;                  @Error 1 @Extended 5 Return 0 = $bRelativeHeight and $bRelativeWidth both set to False.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error Retrieving Page Style Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function isn't totally necessary, because when setting Relative Width/Height, for a Frame/Frame style, the frame is still appropriately set to the correct percentage. However, the L.O. U.I. does not show the percentage unless you set a width value for the frame or frame style based on the Page width.
;                  For Frame Styles, If you notice in L.O. when you set the relative value, while the Viewcursor is in one Page Style, and then move the cursor to another type of page style, the percentage changes. So when I am modifying a Frame Style obtain the ViewCursor, retrieve what Page Style it is currently in, and calculate the Width/Height values based on that sizing.
;                  Or when modifying a Frame, I obtain its anchor, and retrieve the page style name, and get the page size settings for setting Frame Width/Height.
;                  However, it makes no material difference, as the frame still is set to the correct width/height regardless.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ObjRelativeSize(ByRef $oDoc, ByRef $oObj, $bRelativeWidth = False, $bRelativeHeight = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iPageWidth, $iPageHeight, $iObjWidth, $iObjHeight
	Local $oPageStyle

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bRelativeWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bRelativeHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If (($bRelativeHeight = False) And ($bRelativeWidth = False)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If ($oObj.supportsService("com.sun.star.text.TextFrame")) Or ($oObj.supportsService("com.sun.star.text.TextGraphicObject")) Then
		$oPageStyle = $oDoc.StyleFamilies().getByName("PageStyles").getByName($oObj.Anchor.PageStyleName())

	Else
		$oPageStyle = $oDoc.StyleFamilies().getByName("PageStyles").getByName($oDoc.CurrentController.getViewCursor().PageStyleName())
	EndIf

	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($bRelativeWidth = True) Then
		$iPageWidth = $oPageStyle.Width() ; Retrieve total Page Style width
		$iPageWidth = $iPageWidth - $oPageStyle.RightMargin()
		$iPageWidth = $iPageWidth - $oPageStyle.LeftMargin() ; Minus off both margins.

		$iObjWidth = Int($iPageWidth * ($oObj.RelativeWidth() / 100)) ; Times Page width minus margins by relative width percentage.

		$oObj.Width = $iObjWidth
	EndIf

	If ($bRelativeHeight = True) Then
		$iPageHeight = $oPageStyle.Height() ; Retrieve total Page Style Height
		$iPageHeight = $iPageHeight - $oPageStyle.TopMargin()
		$iPageHeight = $iPageHeight - $oPageStyle.BottomMargin() ; Minus off both margins.

		$iObjHeight = Int($iPageHeight * ($oObj.RelativeHeight() / 100)) ; Times Page Height minus margins by relative Height percentage.

		$oObj.Height = $iObjHeight
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_ObjRelativeSize

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_PageStyleNameToggle
; Description ...: Toggle from Page Style Display Name to Internal Name for error checking and setting retrieval.
; Syntax ........: __LOWriter_PageStyleNameToggle(ByRef $sPageStyle[, $bReverse = False])
; Parameters ....: $sPageStyle          - a string value. The Page Style Name to Toggle.
;                  $bReverse            - [optional] a boolean value. Default is False. If True Reverse toggles the Page Style Name.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sPageStyle not a String.
;                  @Error 1 @Extended 2 Return 0 = $bReverse not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Page Style Name successfully toggled. Returning changed name as a string.
;                  @Error 0 @Extended 1 Return String = Success. Page Style Name successfully reverse toggled. Returning changed name as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_PageStyleNameToggle($sPageStyle, $bReverse = False)
	If Not IsString($sPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReverse) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If ($bReverse = False) Then
		$sPageStyle = ($sPageStyle = "Default Page Style") ? ("Standard") : ($sPageStyle)

		Return SetError($__LO_STATUS_SUCCESS, 0, $sPageStyle)

	Else
		$sPageStyle = ($sPageStyle = "Standard") ? ("Default Page Style") : ($sPageStyle)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sPageStyle)
	EndIf
EndFunc   ;==>__LOWriter_PageStyleNameToggle

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParAlignment
; Description ...: Set and Retrieve Alignment settings.
; Syntax ........: __LOWriter_ParAlignment(ByRef $oObj, $iHorAlign, $iVertAlign, $iLastLineAlign, $bExpandSingleWord, $bSnapToGrid, $iTxtDirection)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iHorAlign           - an integer value (0-3). The Horizontal alignment of the paragraph. See Constants, $LOW_PAR_ALIGN_HOR_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $iVertAlign          - an integer value (0-4). The Vertical alignment of the paragraph. See Constants, $LOW_PAR_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLastLineAlign      - an integer value (0-3). Specify the alignment for the last line in the paragraph. See Constants, $LOW_PAR_LAST_LINE_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $bExpandSingleWord   - a boolean value. If True, and the last line of a justified paragraph consists of one word, the word is stretched to the width of the paragraph.
;                  $bSnapToGrid         - a boolean value. If True, Aligns the paragraph to a text grid (if one is active).
;                  $iTxtDirection       - an integer value (0-5). The Text Writing Direction. See Constants, $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3. [Libre Office Default is 4]
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iHorAlign not an integer, less than 0, or greater than 3. See Constants, $LOW_PAR_ALIGN_HOR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iVertAlign not an integer, less than 0 or more than 4. See Constants, $LOW_PAR_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iLastLineAlign not an integer, less than 0 or more than 3. See Constants, $LOW_PAR_LAST_LINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $bExpandSingleWord not a Boolean.
;                  @Error 1 @Extended 8 Return 0 = $bSnapToGrid not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $iTxtDirection not an Integer, less than 0, or greater than 5, See Constants, $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iHorAlign
;                  |                               2 = Error setting $iVertAlign
;                  |                               4 = Error setting $iLastLineALign
;                  |                               8 = Error setting $bExpandSIngleWord
;                  |                               16 = Error setting $bSnapToGrid
;                  |                               32 = Error setting $iTxtDirection
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iHorAlign must be set to $LOW_PAR_ALIGN_HOR_JUSTIFIED(2) before you can set $iLastLineAlign, and $iLastLineAlign must be set to $LOW_PAR_LAST_LINE_JUSTIFIED(2) before $bExpandSingleWord can be set.
;                  $iTxtDirection constants 2,3, and 5 may not be available depending on your language settings.
;                  Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParAlignment(ByRef $oObj, $iHorAlign, $iVertAlign, $iLastLineAlign, $bExpandSingleWord, $bSnapToGrid, $iTxtDirection)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avAlignment[6]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($iHorAlign, $iVertAlign, $iLastLineAlign, $bExpandSingleWord, $bSnapToGrid, $iTxtDirection) Then
		__LO_ArrayFill($avAlignment, $oObj.ParaAdjust(), $oObj.ParaVertAlignment(), $oObj.ParaLastLineAdjust(), $oObj.ParaExpandSingleWord(), _
				$oObj.SnapToGrid(), $oObj.WritingMode())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avAlignment)
	EndIf

	If ($iHorAlign <> Null) Then
		If Not __LO_IntIsBetween($iHorAlign, $LOW_PAR_ALIGN_HOR_LEFT, $LOW_PAR_ALIGN_HOR_CENTER) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.ParaAdjust = $iHorAlign
		$iError = ($oObj.ParaAdjust() = $iHorAlign) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOW_PAR_ALIGN_VERT_AUTO, $LOW_PAR_ALIGN_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.ParaVertAlignment = $iVertAlign
		$iError = ($oObj.ParaVertAlignment() = $iVertAlign) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iLastLineAlign <> Null) Then
		If Not __LO_IntIsBetween($iLastLineAlign, $LOW_PAR_LAST_LINE_JUSTIFIED, $LOW_PAR_LAST_LINE_CENTER, "", $LOW_PAR_LAST_LINE_START) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.ParaLastLineAdjust = $iLastLineAlign
		$iError = ($oObj.ParaLastLineAdjust() = $iLastLineAlign) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bExpandSingleWord <> Null) Then
		If Not IsBool($bExpandSingleWord) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oObj.ParaExpandSingleWord = $bExpandSingleWord
		$iError = ($oObj.ParaExpandSingleWord() = $bExpandSingleWord) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bSnapToGrid <> Null) Then
		If Not IsBool($bSnapToGrid) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oObj.SnapToGrid = $bSnapToGrid
		$iError = ($oObj.SnapToGrid() = $bSnapToGrid) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($iTxtDirection <> Null) Then
		If Not __LO_IntIsBetween($iTxtDirection, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oObj.WritingMode = $iTxtDirection
		$iError = ($oObj.WritingMode() = $iTxtDirection) ? ($iError) : (BitOR($iError, 32))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_ParAlignment

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParBackColor
; Description ...: Set or Retrieve background color settings.
; Syntax ........: __LOWriter_ParBackColor(ByRef $oObj, $iBackColor, $bBackTransparent)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iBackColor          - an integer value (-1-16777215). The background color. Set in Long integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Set to $LO_COLOR_OFF(-1) for "None".
;                  $bBackTransparent    - a boolean value. If True, the background color is transparent
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iBackColor not an integer, less than -1, or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $bBackTransparent not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBackColor
;                  |                               2 = Error setting $bBackTransparent
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParBackColor(ByRef $oObj, $iBackColor, $bBackTransparent)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avColor[2]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($iBackColor, $bBackTransparent) Then
		__LO_ArrayFill($avColor, $oObj.ParaBackColor(), $oObj.ParaBackTransparent())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avColor)
	EndIf

	If ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.ParaBackColor = $iBackColor
		$iError = ($oObj.ParaBackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bBackTransparent <> Null) Then
		If Not IsBool($bBackTransparent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.ParaBackTransparent = $bBackTransparent
		$iError = ($oObj.ParaBackTransparent() = $bBackTransparent) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_ParBackColor

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParBorderPadding
; Description ...: Set or retrieve the Border Padding (spacing between the Paragraph and border) settings.
; Syntax ........: __LOWriter_ParBorderPadding(ByRef $oObj, $iAll, $iTop, $iBottom, $iLeft, $iRight)
; Parameters ....: $oObj                - [in/out] an object. A Paragraph Style object returned by a previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $iAll                - an integer value. Set all four padding distances to one distance in Micrometers (uM).
;                  $iTop                - an integer value. Set the Top Distance between the Border and Paragraph in Micrometers(uM).
;                  $iBottom             - an integer value. Set the Bottom Distance between the Border and Paragraph in Micrometers(uM).
;                  $iLeft               - an integer value. Set the Left Distance between the Border and Paragraph in Micrometers(uM).
;                  $iRight              - an integer value. Set the Right Distance between the Border and Paragraph in Micrometers(uM).
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Passed Object to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iAll not an Integer.
;                  @Error 1 @Extended 5 Return 0 = $iTop not an Integer.
;                  @Error 1 @Extended 6 Return 0 = $iBottom not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $Left not an Integer.
;                  @Error 1 @Extended 8 Return 0 = $iRight not an Integer.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iAll border distance
;                  |                               2 = Error setting $iTop border distance
;                  |                               4 = Error setting $iBottom border distance
;                  |                               8 = Error setting $iLeft border distance
;                  |                               16 = Error setting $iRight border distance
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParBorderPadding(ByRef $oObj, $iAll, $iTop, $iBottom, $iLeft, $iRight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiBPadding[5]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($iAll, $iTop, $iBottom, $iLeft, $iRight) Then
		__LO_ArrayFill($aiBPadding, $oObj.BorderDistance(), $oObj.TopBorderDistance(), $oObj.BottomBorderDistance(), _
				$oObj.LeftBorderDistance(), $oObj.RightBorderDistance())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiBPadding)
	EndIf

	If ($iAll <> Null) Then
		If Not __LO_IntIsBetween($iAll, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.BorderDistance = $iAll
		$iError = (__LO_IntIsBetween($oObj.BorderDistance(), $iAll - 1, $iAll + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iTop <> Null) Then
		If Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.TopBorderDistance = $iTop
		$iError = (__LO_IntIsBetween($oObj.TopBorderDistance(), $iTop - 1, $iTop + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iBottom <> Null) Then
		If Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.BottomBorderDistance = $iBottom
		$iError = (__LO_IntIsBetween($oObj.BottomBorderDistance(), $iBottom - 1, $iBottom + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iLeft <> Null) Then
		If Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oObj.LeftBorderDistance = $iLeft
		$iError = (__LO_IntIsBetween($oObj.LeftBorderDistance(), $iLeft - 1, $iLeft + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iRight <> Null) Then
		If Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oObj.RightBorderDistance = $iRight
		$iError = (__LO_IntIsBetween($oObj.RightBorderDistance(), $iRight - 1, $iRight + 1)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_ParBorderPadding

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParDropCaps
; Description ...: Set or Retrieve DropCaps settings
; Syntax ........: __LOWriter_ParDropCaps(ByRef $oObj, $iNumChar, $iLines, $iSpcTxt, $bWholeWord, $sCharStyle)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iNumChar            - an integer value (0-9). The number of characters to make into DropCaps.
;                  $iLines              - an integer value (0, 2-9). The number of lines to drop down.
;                  $iSpcTxt             - an integer value. The distance between the drop cap and the following text. In Micrometers.
;                  $bWholeWord          - a boolean value. If True, DropCap the whole first word. (Nullifys $iNumChars.)
;                  $sCharStyle          - a string value. The character style to use for the DropCaps. See Remarks.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 5 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 6 Return 0 = $iNumChar not an integer, less than 0, or greater than 9.
;                  @Error 1 @Extended 7 Return 0 = $iLines not an Integer, less than 0, equal to 1 or greater than 9
;                  @Error 1 @Extended 8 Return 0 = $iSpaceTxt not an Integer, or less than 0.
;                  @Error 1 @Extended 9 Return 0 = $bWholeWord not a Boolean.
;                  @Error 1 @Extended 10 Return 0 = $sCharStyle not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving DropCap Format Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iNumChar
;                  |                               2 = Error setting $iLines
;                  |                               4 = Error setting $iSpcTxt
;                  |                               8 = Error setting $bWholeWord
;                  |                               16 = Error setting $sCharStyle
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Set $iNumChars, $iLines, $iSpcTxt to 0 to disable DropCaps.
;                  I am unable to find a way to set Drop Caps character style to "None" as is available in the User Interface. When it is set to "None" Libre returns a blank string ("") but setting it to a blank string throws a COM error/Exception, even when attempting to set it to Libre's own return value without any in-between variables, in case I was mistaken as to it being a blank string, but this still caused a COM error. So consequently, you cannot set Character Style to "None", but you can still disable Drop Caps as noted above.
;                  Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParDropCaps(ByRef $oObj, $iNumChar, $iLines, $iSpcTxt, $bWholeWord, $sCharStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $tDCFrmt
	Local $avDropCaps[5]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$tDCFrmt = $oObj.DropCapFormat()
	If Not IsObj($tDCFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iNumChar, $iLines, $iSpcTxt, $bWholeWord, $sCharStyle) Then
		__LO_ArrayFill($avDropCaps, $tDCFrmt.Count(), $tDCFrmt.Lines(), $tDCFrmt.Distance(), $oObj.DropCapWholeWord(), _
				__LOWriter_CharStyleNameToggle($oObj.DropCapCharStyleName(), True))

		Return SetError($__LO_STATUS_SUCCESS, 1, $avDropCaps)
	EndIf

	If Not __LO_VarsAreNull($iNumChar, $iLines, $iSpcTxt) Then
		If ($iNumChar <> Null) Then
			If Not __LO_IntIsBetween($iNumChar, 0, 9) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

			$tDCFrmt.Count = $iNumChar
		EndIf

		If ($iLines <> Null) Then
			If Not __LO_IntIsBetween($iLines, 0, 9, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

			$tDCFrmt.Lines = $iLines
		EndIf

		If ($iSpcTxt <> Null) Then
			If Not __LO_IntIsBetween($iSpcTxt, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

			$tDCFrmt.Distance = $iSpcTxt
		EndIf

		$oObj.DropCapFormat = $tDCFrmt
		$iError = ($iNumChar = Null) ? ($iError) : (($tDCFrmt.Count() = $iNumChar) ? ($iError) : (BitOR($iError, 1)))
		$iError = ($iLines = Null) ? ($iError) : (($tDCFrmt.Lines() = $iLines) ? ($iError) : (BitOR($iError, 2)))
		$iError = ($iSpcTxt = Null) ? ($iError) : ((__LO_IntIsBetween($tDCFrmt.Distance(), $iSpcTxt - 1, $iSpcTxt + 1)) ? ($iError) : (BitOR($iError, 4)))
	EndIf

	If ($bWholeWord <> Null) Then
		If Not IsBool($bWholeWord) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oObj.DropCapWholeWord = $bWholeWord
		$iError = ($oObj.DropCapWholeWord() = $bWholeWord) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($sCharStyle <> Null) Then
		If Not IsString($sCharStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$sCharStyle = __LOWriter_CharStyleNameToggle($sCharStyle)
		$oObj.DropCapCharStyleName = $sCharStyle
		$iError = ($oObj.DropCapCharStyleName() = $sCharStyle) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_ParDropCaps

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParHasTabStop
; Description ...: Check whether a Paragraph has a requested TabStop created already.
; Syntax ........: __LOWriter_ParHasTabStop(ByRef $oObj, $iTabStop)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iTabStop            - an integer value. The Tab Stop to look for.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iTabStop not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve ParaTabStops Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = True if Paragraph has the requested TabStop. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParHasTabStop(ByRef $oObj, $iTabStop)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atTabStops

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$atTabStops = $oObj.ParaTabStops()
	If Not IsArray($atTabStops) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	For $i = 0 To UBound($atTabStops) - 1
		If ($atTabStops[$i].Position() = $iTabStop) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 0, False)
EndFunc   ;==>__LOWriter_ParHasTabStop

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParHyphenation
; Description ...: Set or Retrieve Hyphenation settings.
; Syntax ........: __LOWriter_ParHyphenation(ByRef $oObj, $bAutoHyphen, $bHyphenNoCaps, $iMaxHyphens, $iMinLeadingChar, $iMinTrailingChar)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $bAutoHyphen         - a boolean value. If True, automatic hyphenation is applied.
;                  $bHyphenNoCaps       - a boolean value. If True, hyphenation will be disabled for words written in CAPS for this paragraph. Libre 6.4 and up.
;                  $iMaxHyphens         - an integer value (0-99). The maximum number of consecutive hyphens.
;                  $iMinLeadingChar     - an integer value (2-9). Specifies the minimum number of characters to remain before the hyphen character (when hyphenation is applied).
;                  $iMinTrailingChar    - an integer value (2-9). Specifies the minimum number of characters to remain after the hyphen character (when hyphenation is applied).
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bAutoHyphen not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bHyphenNoCaps not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iMaxHyphens not an Integer, less than 0, or greater than 99.
;                  @Error 1 @Extended 7 Return 0 = $iMinLeadingChar not an Integer, less than 2, or greater than 9.
;                  @Error 1 @Extended 8 Return 0 = $iMinTrailingChar not an Integer, less than 2, or greater than 9.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bAutoHyphen
;                  |                               2 = Error setting $bHyphenNoCaps
;                  |                               4 = Error setting $iMaxHyphens
;                  |                               8 = Error setting $iMinLeadingChar
;                  |                               16 = Error setting $iMinTrailingChar
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 6.4.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 or 5 Element Array with values in order of function parameters. If the current Libre Office Version is below 6.4, then the Array returned will contain 4 elements because $bHyphenNoCaps is not available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $bAutoHyphen needs to be set to True for the rest of the settings to be activated, but they will be still successfully be set regardless.
;                  Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParHyphenation(ByRef $oObj, $bAutoHyphen, $bHyphenNoCaps, $iMaxHyphens, $iMinLeadingChar, $iMinTrailingChar)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avHyphenation[4]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($bAutoHyphen, $bHyphenNoCaps, $iMaxHyphens, $iMinLeadingChar, $iMinTrailingChar) Then
		If __LO_VersionCheck(6.4) Then
			__LO_ArrayFill($avHyphenation, $oObj.ParaIsHyphenation(), $oObj.ParaHyphenationNoCaps(), $oObj.ParaHyphenationMaxHyphens(), _
					$oObj.ParaHyphenationMaxLeadingChars(), $oObj.ParaHyphenationMaxTrailingChars())

		Else
			__LO_ArrayFill($avHyphenation, $oObj.ParaIsHyphenation(), $oObj.ParaHyphenationMaxHyphens(), _
					$oObj.ParaHyphenationMaxLeadingChars(), $oObj.ParaHyphenationMaxTrailingChars())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avHyphenation)
	EndIf

	If ($bAutoHyphen <> Null) Then
		If Not IsBool($bAutoHyphen) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.ParaIsHyphenation = $bAutoHyphen
		$iError = ($oObj.ParaIsHyphenation = $bAutoHyphen) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bHyphenNoCaps <> Null) Then
		If Not IsBool($bHyphenNoCaps) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If Not __LO_VersionCheck(6.4) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oObj.ParaHyphenationNoCaps = $bHyphenNoCaps
		$iError = ($oObj.ParaHyphenationNoCaps = $bHyphenNoCaps) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iMaxHyphens <> Null) Then
		If Not __LO_IntIsBetween($iMaxHyphens, 0, 99) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.ParaHyphenationMaxHyphens = $iMaxHyphens
		$iError = ($oObj.ParaHyphenationMaxHyphens = $iMaxHyphens) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iMinLeadingChar <> Null) Then
		If Not __LO_IntIsBetween($iMinLeadingChar, 2, 9) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oObj.ParaHyphenationMaxLeadingChars = $iMinLeadingChar
		$iError = ($oObj.ParaHyphenationMaxLeadingChars = $iMinLeadingChar) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iMinTrailingChar <> Null) Then
		If Not __LO_IntIsBetween($iMinTrailingChar, 2, 9) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oObj.ParaHyphenationMaxTrailingChars = $iMinTrailingChar
		$iError = ($oObj.ParaHyphenationMaxTrailingChars = $iMinTrailingChar) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_ParHyphenation

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParIndent
; Description ...: Set or Retrieve Indent settings.
; Syntax ........: __LOWriter_ParIndent(ByRef $oObj, $iBeforeTxt, $iAfterTxt, $iFirstLine, $bAutoFirstLine)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iBeforeTxt          - an integer value (-9998989-17094). The amount of space that you want to indent the paragraph from the page margin. If you want the paragraph to extend into the page margin, enter a negative number. Set in Micrometers(uM).
;                  $iAfterTxt           - an integer value (-9998989-17094). The amount of space that you want to indent the paragraph from the page margin. If you want the paragraph to extend into the page margin, enter a negative number. Set in Micrometers(uM)
;                  $iFirstLine          - an integer value (-57785-17094). Indentation distance of the first line of a paragraph. Set in Micrometers(uM).
;                  $bAutoFirstLine      - a boolean value. If True, the first line will be indented automatically.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iBeforeText not an integer, less than -9998989 or more than 17094 uM.
;                  @Error 1 @Extended 5 Return 0 = $iAfterText not an integer, less than -9998989 or more than 17094 uM.
;                  @Error 1 @Extended 6 Return 0 = $iFirstLine not an integer, less than -57785 or more than 17094 uM.
;                  @Error 1 @Extended 7 Return 0 = $bAutoFirstLine not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBeforeTxt
;                  |                               2 = Error setting $iAfterTxt
;                  |                               4 = Error setting $iFirstLine
;                  |                               8 = Error setting $bAutoFirstLine
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iFirstLine Indent cannot be set if $bAutoFirstLine is set to True.
;                  Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParIndent(ByRef $oObj, $iBeforeTxt, $iAfterTxt, $iFirstLine, $bAutoFirstLine)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avIndent[4]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($iBeforeTxt, $iAfterTxt, $iFirstLine, $bAutoFirstLine) Then
		__LO_ArrayFill($avIndent, $oObj.ParaLeftMargin(), $oObj.ParaRightMargin(), $oObj.ParaFirstLineIndent(), $oObj.ParaIsAutoFirstLineIndent())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avIndent)
	EndIf

	; Min: -9998989; Max: 17094
	If ($iBeforeTxt <> Null) Then
		If Not __LO_IntIsBetween($iBeforeTxt, -9998989, 17094) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.ParaLeftMargin = $iBeforeTxt
		$iError = (__LO_NumIsBetween(($oObj.ParaLeftMargin()), ($iBeforeTxt - 1), ($iBeforeTxt + 1))) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iAfterTxt <> Null) Then
		If Not __LO_IntIsBetween($iAfterTxt, -9998989, 17094) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.ParaRightMargin = $iAfterTxt
		$iError = (__LO_NumIsBetween(($oObj.ParaRightMargin()), ($iAfterTxt - 1), ($iAfterTxt + 1))) ? ($iError) : (BitOR($iError, 2))
	EndIf

	; max 17094; min;-57785
	If ($iFirstLine <> Null) Then
		If Not __LO_IntIsBetween($iFirstLine, -57785, 17094) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.ParaFirstLineIndent = $iFirstLine
		$iError = (__LO_NumIsBetween(($oObj.ParaFirstLineIndent()), ($iFirstLine - 1), ($iFirstLine + 1))) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bAutoFirstLine <> Null) Then
		If Not IsBool($bAutoFirstLine) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oObj.ParaIsAutoFirstLineIndent = $bAutoFirstLine
		$iError = ($oObj.ParaIsAutoFirstLineIndent() = $bAutoFirstLine) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_ParIndent

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParOutLineAndList
; Description ...: Set and Retrieve the Outline and List settings.
; Syntax ........: __LOWriter_ParOutLineAndList(ByRef $oObj, $iOutline, $sNumStyle, $bParLineCount, $iLineCountVal)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iOutline            - an integer value (0-10). The Outline Level, see Constants, $LOW_OUTLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sNumStyle           - a string value. Specifies the name of the style for the Paragraph numbering. Set to "" for None.
;                  $bParLineCount       - a boolean value. If True, the paragraph is included in the line numbering.
;                  $iLineCountVal       - an integer value. The start value for numbering if a new numbering starts at this paragraph. Set to 0 for no line numbering restart.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 5 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 6 Return 0 = $iOutline not an integer, less than 0, or greater than 10. See constants, $LOW_OUTLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $sNumStyle not a String.
;                  @Error 1 @Extended 8 Return 0 = $bParLineCount not a Boolean.
;                  @Error 1 @Extended 9 Return 0 = $iLineCountVal not an Integer or less than 0.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iOutline
;                  |                               2 = Error setting $sNumStyle
;                  |                               4 = Error setting $bParLineCount
;                  |                               8 = Error setting $iLineCountVal
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParOutLineAndList(ByRef $oObj, $iOutline, $sNumStyle, $bParLineCount, $iLineCountVal)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avOutlineNList[4]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If __LO_VarsAreNull($iOutline, $sNumStyle, $bParLineCount, $iLineCountVal) Then
		__LO_ArrayFill($avOutlineNList, $oObj.OutlineLevel(), $oObj.NumberingStyleName(), $oObj.ParaLineNumberCount(), _
				$oObj.ParaLineNumberStartValue())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avOutlineNList)
	EndIf

	If ($iOutline <> Null) Then
		If Not __LO_IntIsBetween($iOutline, $LOW_OUTLINE_BODY, $LOW_OUTLINE_LEVEL_10) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.OutlineLevel = $iOutline
		$iError = ($oObj.OutlineLevel = $iOutline) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sNumStyle <> Null) Then
		If Not IsString($sNumStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oObj.NumberingStyleName = $sNumStyle
		$iError = ($oObj.NumberingStyleName = $sNumStyle) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bParLineCount <> Null) Then
		If Not IsBool($bParLineCount) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oObj.ParaLineNumberCount = $bParLineCount
		$iError = ($oObj.ParaLineNumberCount = $bParLineCount) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iLineCountVal <> Null) Then
		If Not __LO_IntIsBetween($iLineCountVal, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$oObj.ParaLineNumberStartValue = $iLineCountVal
		$iError = ($oObj.ParaLineNumberStartValue = $iLineCountVal) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_ParOutLineAndList

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParPageBreak
; Description ...: Set or Retrieve Page Break Settings.
; Syntax ........: __LOWriter_ParPageBreak(ByRef $oObj, $iBreakType, $sPageStyle, $iPgNumOffSet)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iBreakType          - an integer value (0-6). The Page Break Type. See Constants, $LOW_BREAK_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sPageStyle          - a string value. Creates a page break before the paragraph it belongs to and assigns the new page style to use. Note: If you set this parameter, to remove the page break setting you must set this to "".
;                  $iPgNumOffSet        - an integer value. If a page break property is set at a paragraph, this property contains the new value for the page number.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 5 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 6 Return 0 = $iBreakType not an integer, less than 0, or greater than 6. See Constants, $LOW_BREAK_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 7 Return 0 = $sPageStyle not a String.
;                  @Error 1 @Extended 8 Return 0 = $iPgNumOffSet not an Integer or less than 0.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBreakType
;                  |                               2 = Error setting $sPageStyle
;                  |                               4 = Error setting $iPgNumOffSet
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Break Type must be set before Page Style will be able to be set, and page style needs set before $iPgNumOffSet can be set.
;                  Libre doesn't directly show in its User interface options for Break type constants #3 and #6 (Column both) and (Page both), but doesn't throw an error when being set to either one, so they are included here, though I'm not sure if they will work correctly.
;                  Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParPageBreak(ByRef $oObj, $iBreakType, $sPageStyle, $iPgNumOffSet)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avPageBreak[3]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If __LO_VarsAreNull($iBreakType, $sPageStyle, $iPgNumOffSet) Then
		__LO_ArrayFill($avPageBreak, $oObj.BreakType(), $oObj.PageDescName(), $oObj.PageNumberOffset())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avPageBreak)
	EndIf

	If ($iBreakType <> Null) Then
		If Not __LO_IntIsBetween($iBreakType, $LOW_BREAK_NONE, $LOW_BREAK_PAGE_BOTH) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.BreakType = $iBreakType
		$iError = ($oObj.BreakType = $iBreakType) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sPageStyle <> Null) Then
		If Not IsString($sPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oObj.PageDescName = $sPageStyle
		$iError = ($oObj.PageDescName = $sPageStyle) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iPgNumOffSet <> Null) Then
		If Not __LO_IntIsBetween($iPgNumOffSet, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oObj.PageNumberOffset = $iPgNumOffSet
		$iError = ($oObj.PageNumberOffset = $iPgNumOffSet) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_ParPageBreak

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParShadow
; Description ...: Set or Retrieve the Shadow settings for a Paragraph.
; Syntax ........: __LOWriter_ParShadow(ByRef $oObj, $iWidth, $iColor, $bTransparent, $iLocation)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iWidth              - an integer value. The shadow width in Micrometers.
;                  $iColor              - an integer value (0-16777215). The shadow color, set in Long Integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bTransparent        - a boolean value. If True, the shadow is transparent.
;                  $iLocation           - an integer value (0-4). The location of the shadow compared to the paragraph. See Constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iWidth not an integer or less than 0.
;                  @Error 1 @Extended 5 Return 0 = $iColor not an integer, less than 0, or greater than 16777215.
;                  @Error 1 @Extended 6 Return 0 = $bTransparent not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $iLocation not an Integer, less than 0, or greater than 4. See Constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Shadow Format Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Shadow Format Object for Error Checking.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iWidth
;                  |                               2 = Error setting $iColor
;                  |                               4 = Error setting $bTransparent
;                  |                               8 = Error setting $iLocation
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  LibreOffice may change the shadow width +/- a Micrometer.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParShadow(ByRef $oObj, $iWidth, $iColor, $bTransparent, $iLocation)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $tShdwFrmt
	Local $avShadow[4]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tShdwFrmt = $oObj.ParaShadowFormat()
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iWidth, $iColor, $bTransparent, $iLocation) Then
		__LO_ArrayFill($avShadow, $tShdwFrmt.ShadowWidth(), $tShdwFrmt.Color(), $tShdwFrmt.IsTransparent(), $tShdwFrmt.Location())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avShadow)
	EndIf

	If ($iWidth <> Null) Then
		If Not __LO_IntIsBetween($iWidth, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tShdwFrmt.ShadowWidth = $iWidth
	EndIf

	If ($iColor <> Null) Then
		If Not __LO_IntIsBetween($iColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tShdwFrmt.Color = $iColor
	EndIf

	If ($bTransparent <> Null) Then
		If Not IsBool($bTransparent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tShdwFrmt.IsTransparent = $bTransparent
	EndIf

	If ($iLocation <> Null) Then
		If Not __LO_IntIsBetween($iLocation, $LOW_SHADOW_NONE, $LOW_SHADOW_BOTTOM_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tShdwFrmt.Location = $iLocation
	EndIf

	$oObj.ParaShadowFormat = $tShdwFrmt
	; Error Checking
	$tShdwFrmt = $oObj.ParaShadowFormat()
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$iError = ($iWidth = Null) ? ($iError) : (($tShdwFrmt.ShadowWidth() = $iWidth) ? ($iError) : (BitOR($iError, 1)))
	$iError = ($iColor = Null) ? ($iError) : (($tShdwFrmt.Color() = $iColor) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($bTransparent = Null) ? ($iError) : (($tShdwFrmt.IsTransparent() = $bTransparent) ? ($iError) : (BitOR($iError, 4)))
	$iError = ($iLocation = Null) ? ($iError) : (($tShdwFrmt.Location() = $iLocation) ? ($iError) : (BitOR($iError, 8)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_ParShadow

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParSpace
; Description ...: Set and Retrieve Line Spacing settings.
; Syntax ........: __LOWriter_ParSpace(ByRef $oObj, $iAbovePar, $iBelowPar, $bAddSpace, $iLineSpcMode, $iLineSpcHeight, $bPageLineSpc)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iAbovePar           - an integer value (0-10008). The Space above a paragraph, in Micrometers.
;                  $iBelowPar           - an integer value (0-10008). The Space Below a paragraph, in Micrometers.
;                  $bAddSpace           - a boolean value. If true, the top and bottom margins of the paragraph should not be applied when the previous and next paragraphs have the same style. Libre Office Version 3.6 and Up.
;                  $iLineSpcMode        - an integer value (0-3). The line spacing type of the paragraph. See Constants, $LOW_LINE_SPC_MODE_* as defined in LibreOfficeWriter_Constants.au3, also notice min and max values for each.
;                  $iLineSpcHeight      - an integer value. This value specifies the height in regard to Mode. See Remarks.
;                  $bPageLineSpc        - a boolean value. If True, register mode is applied to a paragraph. See Remarks.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iAbovePar not an integer, less than 0 or more than 10008 uM.
;                  @Error 1 @Extended 5 Return 0 = $iBelowPar not an integer, less than 0 or more than 10008 uM.
;                  @Error 1 @Extended 6 Return 0 = $bAddSpc not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $iLineSpcMode not an integer, less than 0, or greater than 3. See Constants, $LOW_LINE_SPC_MODE_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $iLineSpcHeight not an Integer.
;                  @Error 1 @Extended 9 Return 0 = $iLineSpcMode set to 0(Proportional) and $iLineSpcHeight less than 6(%) or greater than 65535(%).
;                  @Error 1 @Extended 10 Return 0 = $iLineSpcMode set to 1 or 2(Minimum, or Leading) and $iLineSpcHeight less than 0 uM or greater than 10008 uM
;                  @Error 1 @Extended 11 Return 0 = $iLineSpcMode set to 3(Fixed) and $iLineSpcHeight less than 51 uM or greater than 10008 uM.
;                  @Error 1 @Extended 12 Return 0 = $bPageLineSpc not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving ParaLineSpacing Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iAbovePar
;                  |                               2 = Error setting $iBelowPar
;                  |                               4 = Error setting $bAddSpace
;                  |                               8 = Error setting $iLineSpcMode
;                  |                               16 = Error setting $iLineSpcHeight
;                  |                               32 = Error setting $bPageLineSpc
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $bPageLineSpc(Register mode) is only used if the register mode property of the page style is switched on. $bPageLineSpc(Register Mode) Aligns the baseline of each line of text to a vertical document grid, so that each line is the same height.
;                  The settings in Libre Office, (Single, 1.15, 1.5, Double), Use the Proportional mode, and are just varying percentages. e.g Single = 100, 1.15 = 115%, 1.5 = 150%, Double = 200%.
;                  $iLineSpcHeight depends on the $iLineSpcMode used, see constants for accepted Input values.
;                  $iAbovePar, $iBelowPar, $iLineSpcHeight may change +/- 1 Micrometer once set.
;                  Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParSpace(ByRef $oObj, $iAbovePar, $iBelowPar, $bAddSpace, $iLineSpcMode, $iLineSpcHeight, $bPageLineSpc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tLine
	Local $iError = 0
	Local $avSpacing[5]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($iAbovePar, $iBelowPar, $bAddSpace, $iLineSpcMode, $iLineSpcHeight, $bPageLineSpc) Then
		If __LO_VersionCheck(3.6) Then
			__LO_ArrayFill($avSpacing, $oObj.ParaTopMargin(), $oObj.ParaBottomMargin(), $oObj.ParaContextMargin(), _
					$oObj.ParaLineSpacing.Mode(), $oObj.ParaLineSpacing.Height(), $oObj.ParaRegisterModeActive())

		Else
			__LO_ArrayFill($avSpacing, $oObj.ParaTopMargin(), $oObj.ParaBottomMargin(), $oObj.ParaLineSpacing.Mode(), $oObj.ParaLineSpacing.Height(), _
					$oObj.ParaRegisterModeActive())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avSpacing)
	EndIf

	If ($iAbovePar <> Null) Then
		If Not __LO_IntIsBetween($iAbovePar, 0, 10008) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.ParaTopMargin = $iAbovePar
		$iError = (__LO_NumIsBetween(($oObj.ParaTopMargin()), ($iAbovePar - 1), ($iAbovePar + 1))) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iBelowPar <> Null) Then
		If Not __LO_IntIsBetween($iBelowPar, 0, 10008) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.ParaBottomMargin = $iBelowPar
		$iError = (__LO_NumIsBetween(($oObj.ParaBottomMargin()), ($iBelowPar - 1), ($iBelowPar + 1))) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bAddSpace <> Null) Then
		If Not IsBool($bAddSpace) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If Not __LO_VersionCheck(3.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

		$oObj.ParaContextMargin = $bAddSpace
		$iError = ($oObj.ParaContextMargin = $bAddSpace) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iLineSpcMode <> Null) Then
		If Not __LO_IntIsBetween($iLineSpcMode, $LOW_LINE_SPC_MODE_PROP, $LOW_LINE_SPC_MODE_FIX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tLine = $oObj.ParaLineSpacing()
		If Not IsObj($tLine) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$tLine.Mode = $iLineSpcMode
		$oObj.ParaLineSpacing = $tLine
		$iError = ($oObj.ParaLineSpacing.Mode() = $iLineSpcMode) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iLineSpcHeight <> Null) Then
		If Not IsInt($iLineSpcHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$tLine = $oObj.ParaLineSpacing()
		If Not IsObj($tLine) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Switch $tLine.Mode()
			Case $LOW_LINE_SPC_MODE_PROP ; Proportional
				If Not __LO_IntIsBetween($iLineSpcHeight, 6, 65535) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0) ; Min setting on Proportional is 6%

			Case $LOW_LINE_SPC_MODE_MIN, $LOW_LINE_SPC_MODE_LEADING ; Minimum and Leading Modes
				If Not __LO_IntIsBetween($iLineSpcHeight, 0, 10008) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

			Case $LOW_LINE_SPC_MODE_FIX ; Fixed Line Spacing Mode
				If Not __LO_IntIsBetween($iLineSpcHeight, 51, 10008) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0) ; Min spacing is 51 when Fixed Mode
		EndSwitch
		$tLine.Height = $iLineSpcHeight
		$oObj.ParaLineSpacing = $tLine
		$iError = (__LO_NumIsBetween(($oObj.ParaLineSpacing.Height()), ($iLineSpcHeight - 1), ($iLineSpcHeight + 1))) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bPageLineSpc <> Null) Then
		If Not IsBool($bPageLineSpc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$oObj.ParaRegisterModeActive = $bPageLineSpc
		$iError = ($oObj.ParaRegisterModeActive() = $bPageLineSpc) ? ($iError) : (BitOR($iError, 32))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_ParSpace

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParStyleNameToggle
; Description ...: Toggle from Par Style Display Name to Internal Name for error checking, or setting retrieval.
; Syntax ........: __LOWriter_ParStyleNameToggle(ByRef $sParStyle[, $bReverse = False])
; Parameters ....: $sParStyle           - a string value. The Paragraph Style Name to Toggle.
;                  $bReverse            - [optional] a boolean value. Default is False. If True, Reverse toggles the name.
; Return values .: Success: String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sParStyle not a String.
;                  @Error 1 @Extended 2 Return 0 = $bReverse not a Boolean.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Paragraph Style Name was Successfully toggled. Returning toggled name as a string.
;                  @Error 0 @Extended 1 Return String = Success. Paragraph Style Name was Successfully reverse toggled. Returning toggled name as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParStyleNameToggle($sParStyle, $bReverse = False)
	If Not IsString($sParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReverse) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If ($bReverse = False) Then
		$sParStyle = ($sParStyle = "Default Paragraph Style") ? ("Standard") : ($sParStyle)
		$sParStyle = ($sParStyle = "Complimentary Close") ? ("Salutation") : ($sParStyle)

		Return SetError($__LO_STATUS_SUCCESS, 0, $sParStyle)

	Else
		$sParStyle = ($sParStyle = "Standard") ? ("Default Paragraph Style") : ($sParStyle)
		$sParStyle = ($sParStyle = "Salutation") ? ("Complimentary Close") : ($sParStyle)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sParStyle)
	EndIf
EndFunc   ;==>__LOWriter_ParStyleNameToggle

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParTabStopCreate
; Description ...: Create a new TabStop for a Paragraph.
; Syntax ........: __LOWriter_ParTabStopCreate(ByRef $oObj, $iPosition, $iAlignment, $iFillChar, $iDecChar)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iPosition           - an integer value. The TabStop position to set the new TabStop to. Set in Micrometers (uM). See Remarks.
;                  $iAlignment          - an integer value (0-4). The position of where the end of a Tab is aligned to compared to the text. See Constants, $LOW_TAB_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iFillChar           - an integer value. The Asc (see AutoIt function) value of any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
;                  $iDecChar            - an integer value. Enter a character(in Asc Value(See AutoIt Function)) that you want the decimal tab to use as a decimal separator. Can only be set if $iAlignment is set to $LOW_TAB_ALIGN_DECIMAL.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 5 Return 0 = Passed Object to internal function not an Object.
;                  @Error 1 @Extended 6 Return 0 = $iFillChar not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $iAlignment not an Integer, less than 0, or greater than 4. See Constants, $LOW_TAB_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $iDecChar not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.style.TabStop" Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving ParaTabStops Array Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving list of TabStop Positions.
;                  @Error 3 @Extended 3 Return 0 = Failed to identify the new Tabstop once inserted.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return Integer = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                                     1 = Error setting $iPosition
;                  |                                     2 = Error setting $iFillChar
;                  |                                     4 = Error setting $iAlignment
;                  |                                     8 = Error setting $iDecChar
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Settings were successfully set. New TabStop position is returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iPosition once set can vary +/- 1 uM. To ensure you can identify the tabstop to modify it again, This function returns the new TabStop position.
;                  Since $iPosition can fluctuate +/- 1 uM when it is inserted into LibreOffice, it is possible to accidentally overwrite an already existing TabStop.
;                  $iFillChar, Libre's Default value, "None" is in reality a space character which is Asc value 32. The other values offered by Libre are: Period (ASC 46), Dash (ASC 45) and Underscore (ASC 95). You can also enter a custom ASC value. See ASC AutoIt Func. and "ASCII Character Codes" in the AutoIt help file.
;                  Call any optional parameter with Null keyword to skip it.
;                  $iNewTabStop position is still returned as even though some settings weren't successfully set, the new TabStop was still created.
; Related .......: _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParTabStopCreate(ByRef $oObj, $iPosition, $iAlignment, $iFillChar, $iDecChar)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aiTabList
	Local $bFound = False
	Local $iNewPosition = -1
	Local $atTabStops, $atNewTabStops
	Local $tFoundTabStop, $tTabStruct
	Local $iError = 0

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$atTabStops = $oObj.ParaTabStops()
	If Not IsArray($atTabStops) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tTabStruct = __LO_CreateStruct("com.sun.star.style.TabStop")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tTabStruct.Position = $iPosition
	$tTabStruct.FillChar = 32
	; If set to 0 Libre sets fill character to Null instead of setting to None. 32 = None.(Space character)
	$tTabStruct.Alignment = 0
	$tTabStruct.DecimalChar = 0

	If ($iFillChar <> Null) Then
		If Not IsInt($iFillChar) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tTabStruct.FillChar = ($iFillChar = 0) ? (32) : ($iFillChar)
	EndIf

	If ($iAlignment <> Null) Then
		If Not __LO_IntIsBetween($iAlignment, $LOW_TAB_ALIGN_LEFT, $LOW_TAB_ALIGN_DEFAULT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tTabStruct.Alignment = $iAlignment
	EndIf

	If ($iDecChar <> Null) Then
		If Not IsInt($iDecChar) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$tTabStruct.DecimalChar = $iDecChar
	EndIf

	If ($atTabStops[0].Alignment() = $LOW_TAB_ALIGN_DEFAULT) And (UBound($atTabStops) = 1) Then ; if inserting a  Tabstop for the first time, overwrite the "Default blank TabStop.
		$atTabStops[0] = $tTabStruct
		$oObj.ParaTabStops = $atTabStops ; Insert the new TabStop
		$atNewTabStops = $oObj.ParaTabStops()
		$tFoundTabStop = $atNewTabStops[0]
		$iNewPosition = $tFoundTabStop.Position()

	Else
		__LO_AddTo1DArray($atTabStops, $tTabStruct)

		$aiTabList = __LOWriter_ParTabStopsGetList($oObj) ; Get an array of existing tabstops to compare with
		If Not IsArray($aiTabList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		__LO_AddTo1DArray($aiTabList, 0) ; Add a dummy to make Array sizes equal.

		$oObj.ParaTabStops = $atTabStops ; Insert the new TabStop

		$atNewTabStops = $oObj.ParaTabStops() ; Now retrieve a new list to find the final Tab Stop position.
		For $i = 0 To UBound($atNewTabStops) - 1
			If ($atNewTabStops[$i].Position()) <> $aiTabList[$i] Then
				$iNewPosition = $atNewTabStops[$i].Position()
				$tFoundTabStop = $atNewTabStops[$i]
				$bFound = True
				ExitLoop
			EndIf
		Next

		If Not $bFound Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; Didn't find the new TabStop
	EndIf

	$iError = (__LO_NumIsBetween(($tFoundTabStop.Position()), ($iPosition - 1), ($iPosition + 1))) ? ($iError) : (BitOR($iError, 1))
	$iError = ($iFillChar = Null) ? ($iError) : (($tFoundTabStop.FillChar = $iFillChar) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($iAlignment = Null) ? ($iError) : (($tFoundTabStop.Alignment = $iAlignment) ? ($iError) : (BitOR($iError, 4)))
	$iError = ($iDecChar = Null) ? ($iError) : (($tFoundTabStop.DecimalChar = $iDecChar) ? ($iError) : (BitOR($iError, 8)))

	Return ($iError > 0) ? SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, $iNewPosition) : SetError($__LO_STATUS_SUCCESS, 0, $iNewPosition)
EndFunc   ;==>__LOWriter_ParTabStopCreate

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParTabStopDelete
; Description ...: Delete a TabStop from a Paragraph
; Syntax ........: __LOWriter_ParTabStopDelete(ByRef $oObj, ByRef $oDoc, $iTabStop)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iTabStop            - an integer value. The Tab position of the TabStop to modify. See Remarks.
; Return values .: Success: Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 6 Return 0 = Passed Object to internal function not an Object.
;                  @Error 1 @Extended 7 Return 0 = Passed Document Object to internal function not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving ParaTabStops Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to identify and delete TabStop in Paragraph.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Returns true if TabStop was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iTabStop refers to the position, or essential the "length" of a TabStop from the edge of a page margin. This is the only reliable way to identify a Tabstop to be able to interact with it, as there can only be one of a certain length per paragraph style.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParTabStopDelete(ByRef $oObj, ByRef $oDoc, $iTabStop)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDefaults
	Local $tTabStruct
	Local $atOldTabStops[0], $atNewTabStops[0]
	Local $bDeleted = False
	Local $iCount = 0

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$atOldTabStops = $oObj.ParaTabStops()
	ReDim $atNewTabStops[UBound($atOldTabStops)]
	If Not IsArray($atOldTabStops) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If (UBound($atOldTabStops) = 1) Then
		$oDefaults = $oDoc.createInstance("com.sun.star.text.Defaults")
		$tTabStruct = $atOldTabStops[0]
		$tTabStruct.Alignment = $LOW_TAB_ALIGN_DEFAULT
		$tTabStruct.Position = $oDefaults.TabStopDistance()
		$tTabStruct.FillChar = 32 ; Space
		$tTabStruct.DecimalChar = 46 ; Period
		$atNewTabStops[0] = $tTabStruct
		$bDeleted = True

	Else
		For $i = 0 To UBound($atOldTabStops) - 1
			If ($atOldTabStops[$i].Position() = $iTabStop) Then
				$bDeleted = True

			Else
				$atNewTabStops[$iCount] = $atOldTabStops[$i]
				$iCount += 1
			EndIf
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
		Next
	EndIf
	ReDim $atNewTabStops[(($bDeleted) ? (UBound($atNewTabStops) - 1) : (UBound($atNewTabStops)))]

	$oObj.ParaTabStops = $atNewTabStops

	Return ($bDeleted) ? (SetError($__LO_STATUS_SUCCESS, 0, $bDeleted)) : (SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0))
EndFunc   ;==>__LOWriter_ParTabStopDelete

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParTabStopMod
; Description ...: Modify or retrieve the properties of an existing TabStop.
; Syntax ........: __LOWriter_ParTabStopMod(ByRef $oObj, $iTabStop, $iPosition, $iFillChar, $iAlignment, $iDecChar)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iTabStop            - an integer value. The Tab position of the TabStop to modify. See Remarks.
;                  $iPosition           - an integer value. The New position to set the input position to. Set in Micrometers (uM). See Remarks.
;                  $iFillChar           - an integer value. The Asc (see AutoIt function) value of any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
;                  $iAlignment          - an integer value. The position of where the end of a Tab is aligned to compared to the text. See Constants, $LOW_TAB_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iDecChar            - an integer value. Enter a character(in Asc Value(See AutoIt Function)) that you want the decimal tab to use as a decimal separator. Can only be set if $iAlignment is set to $LOW_TAB_ALIGN_DECIMAL.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 5 Return 0 = Passed Object to internal function not an Object.
;                  @Error 1 @Extended 6 Return 0 = $iPosition not an Integer.
;                  @Error 1 @Extended 7 Return 0 = $iFillChar not an Integer.
;                  @Error 1 @Extended 8 Return 0 = $iAlignment not an Integer, less than 0, or greater than 4. See Constants, $LOW_TAB_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 9 Return 0 = $iDecChar not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving ParaTabStops Object.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Requested TabStop Object.
;                  @Error 3 @Extended 3 Return 0 = Paragraph style already contains a TabStop at the length/Position specified in $iPosition.
;                  @Error 3 @Extended 4 Return 0 = Error retrieving list of TabStop Positions.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iPosition
;                  |                               2 = Error setting $iFillChar
;                  |                               4 = Error setting $iAlignment
;                  |                               8 = Error setting $iDecChar
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  @Error 0 @Extended ? Return 2 = Success. Settings were successfully set. New TabStop position is returned in @Extended.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iTabStop refers to the position, or essential the "length" of a TabStop from the edge of a page margin. This is the only reliable way to identify a Tabstop to be able to interact with it, as there can only be one of a certain length per Paragraph Style.
;                  $iPosition once set can vary +/- 1 uM. To ensure you can identify the tabstop to modify it again, This function returns the new TabStop position in @Extended when $iPosition is set, return value will be set to 2. See Return Values.
;                  Since $iPosition can fluctuate +/- 1 uM when it is inserted into LibreOffice, it is possible to accidentally overwrite an already existing TabStop.
;                  $iFillChar, Libre's Default value, "None" is in reality a space character which is Asc value 32. The other values offered by Libre are: Period (ASC 46), Dash (ASC 45) and Underscore (ASC 95). You can also enter a custom ASC value. See ASC AutoIt Func. and "ASCII Character Codes" in the AutoIt help file.
;                  Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParTabStopMod(ByRef $oObj, $iTabStop, $iPosition, $iFillChar, $iAlignment, $iDecChar)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atTabStops, $atNewTabStops
	Local $iError = 0, $iNewPosition = 0
	Local $tTabStruct
	Local $bNewPosition = False
	Local $aiTabList
	Local $aiTabSettings[4]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$atTabStops = $oObj.ParaTabStops()
	If Not IsArray($atTabStops) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	For $i = 0 To UBound($atTabStops) - 1
		$tTabStruct = ($atTabStops[$i].Position() = $iTabStop) ? ($atTabStops[$i]) : (Null)
		If IsObj($tTabStruct) Then ExitLoop
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next
	If Not IsObj($tTabStruct) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If __LO_VarsAreNull($iPosition, $iFillChar, $iAlignment, $iDecChar) Then
		__LO_ArrayFill($aiTabSettings, $tTabStruct.Position(), $tTabStruct.FillChar(), $tTabStruct.Alignment(), $tTabStruct.DecimalChar())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiTabSettings)
	EndIf

	If ($iPosition <> Null) Then
		If Not IsInt($iPosition) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If __LOWriter_ParHasTabStop($oObj, $iPosition) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$tTabStruct.Position = $iPosition
		$iError = ($tTabStruct.Position() = $iPosition) ? ($iError) : (BitOR($iError, 1))
		$bNewPosition = True
	EndIf

	If ($iFillChar <> Null) Then
		If Not IsInt($iFillChar) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tTabStruct.FillChar = $iFillChar
		$tTabStruct.FillChar = ($tTabStruct.FillChar() = 0) ? (32) : ($tTabStruct.FillChar())
		$iError = ($tTabStruct.FillChar = $iFillChar) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iAlignment <> Null) Then
		If Not __LO_IntIsBetween($iAlignment, $LOW_TAB_ALIGN_LEFT, $LOW_TAB_ALIGN_DEFAULT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$tTabStruct.Alignment = $iAlignment
		$iError = ($tTabStruct.Alignment = $iAlignment) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iDecChar <> Null) Then
		If Not IsInt($iDecChar) And ($iDecChar <> Null) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$tTabStruct.DecimalChar = $iDecChar
		$iError = ($tTabStruct.DecimalChar = $iDecChar) ? ($iError) : (BitOR($iError, 8))
	EndIf

	$atTabStops[$i] = $tTabStruct

	If $bNewPosition Then
		$aiTabList = __LOWriter_ParTabStopsGetList($oObj)
		If Not IsArray($aiTabList) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	$oObj.ParaTabStops = $atTabStops

	If $bNewPosition Then
		$atNewTabStops = $oObj.ParaTabStops()
		For $j = 0 To UBound($atNewTabStops) - 1
			If ($atNewTabStops[$j].Position()) <> $aiTabList[$j] Then
				$iNewPosition = $atNewTabStops[$j].Position()
				ExitLoop
			EndIf
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
		Next

		Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, $iNewPosition, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_ParTabStopMod

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParTabStopsGetList
; Description ...: Retrieve an array of TabStops available in a Paragraph.
; Syntax ........: __LOWriter_ParTabStopsGetList(ByRef $oObj)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
; Return values .: Success: Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving ParaTabStops Object.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. An Array of TabStops. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParTabStopsGetList(ByRef $oObj)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atTabStops[0]
	Local $aiTabList[0]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$atTabStops = $oObj.ParaTabStops()
	If Not IsArray($atTabStops) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	ReDim $aiTabList[UBound($atTabStops)]

	For $i = 0 To UBound($atTabStops) - 1
		$aiTabList[$i] = $atTabStops[$i].Position()
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, $i, $aiTabList)
EndFunc   ;==>__LOWriter_ParTabStopsGetList

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParTxtFlowOpt
; Description ...: Set and Retrieve Text Flow settings.
; Syntax ........: __LOWriter_ParTxtFlowOpt(ByRef $oObj, $bParSplit, $bKeepTogether, $iParOrphans, $iParWidows)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $bParSplit           - a boolean value. If False, prevents the paragraph from getting split between two pages or columns
;                  $bKeepTogether       - a boolean value. If True, prevents page or column breaks between this and the following paragraph
;                  $iParOrphans         - an integer value (0, 2-9). Specifies the minimum number of lines of the paragraph that have to be at bottom of a page if the paragraph is spread over more than one page. 0 = disabled.
;                  $iParWidows          - an integer value (0, 2-9). Specifies the minimum number of lines of the paragraph that have to be at top of a page if the paragraph is spread over more than one page. 0 = disabled.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bParSplit not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bKeepTogether not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iParOrphans not an Integer, less than 0, equal to 1, or greater than 9.
;                  @Error 1 @Extended 7 Return 0 = $iParWidows not an Integer, less than 0, equal to 1, or greater than 9.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bParSplit
;                  |                               2 = Error setting $bKeepTogether
;                  |                               4 = Error setting $iParOrphans
;                  |                               8 = Error setting $iParWidows
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If you do not set ParSplit to True, the rest of the settings will still show to have been set but will not become active until $bParSplit is set to true.
;                  Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParTxtFlowOpt(ByRef $oObj, $bParSplit, $bKeepTogether, $iParOrphans, $iParWidows)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avTxtFlowOpt[4]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($bParSplit, $bKeepTogether, $iParOrphans, $iParWidows) Then
		__LO_ArrayFill($avTxtFlowOpt, $oObj.ParaSplit(), $oObj.ParaKeepTogether(), $oObj.ParaOrphans(), $oObj.ParaWidows())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avTxtFlowOpt)
	EndIf

	If ($bParSplit <> Null) Then
		If Not IsBool($bParSplit) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.ParaSplit = $bParSplit
		$iError = ($oObj.ParaSplit = $bParSplit) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bKeepTogether <> Null) Then
		If Not IsBool($bKeepTogether) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.ParaKeepTogether = $bKeepTogether
		$iError = ($oObj.ParaKeepTogether = $bKeepTogether) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iParOrphans <> Null) Then
		If Not __LO_IntIsBetween($iParOrphans, 0, 9, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.ParaOrphans = $iParOrphans
		$iError = ($oObj.ParaOrphans = $iParOrphans) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iParWidows <> Null) Then
		If Not __LO_IntIsBetween($iParWidows, 0, 9, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oObj.ParaWidows = $iParWidows
		$iError = ($oObj.ParaWidows = $iParWidows) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOWriter_ParTxtFlowOpt

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Shape_CreateArrow
; Description ...: Create a Arrow type Shape.
; Syntax ........: __LOWriter_Shape_CreateArrow($oDoc, $iWidth, $iHeight, $iShapeType)
; Parameters ....: $oDoc                - an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iWidth              - an integer value. The Shape's Width in Micrometers.
;                  $iHeight             - an integer value. The Shape's Height in Micrometers.
;                  $iShapeType          - an integer value (0-25). The Type of shape to create. See $LOW_SHAPE_TYPE_ARROWS_* as defined in LibreOfficeWriter_Constants.au3
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iShapeType not an Integer
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" or "com.sun.star.drawing.EllipseShape" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;                  @Error 2 @Extended 3 Return 0 = Failed to create "MirroredX" property structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve the Size Structure.
;                  @Error 3 @Extended 3 Return 0 = Failed to create a unique Shape name.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The following shapes are not implemented into LibreOffice as of L.O. Version 7.3.4.2 for automation, and thus will not work:
;                  $LOW_SHAPE_TYPE_ARROWS_ARROW_S_SHAPED, $LOW_SHAPE_TYPE_ARROWS_ARROW_SPLIT, $LOW_SHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT,
;                  $LOW_SHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT, $LOW_SHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN, $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_Shape_CreateArrow($oDoc, $iWidth, $iHeight, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape
	Local $tProp, $tProp2, $tSize, $tPos
	Local $atCusShapeGeo[1]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oShape = $oDoc.createInstance("com.sun.star.drawing.CustomShape")
	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tProp = __LO_SetPropertyValue("Type", "")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Switch $iShapeType
		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_4_WAY
			$tProp.Value = "quad-arrow"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_4_WAY
			$tProp.Value = "quad-arrow-callout"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_DOWN
			$tProp.Value = "down-arrow-callout"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT
			$tProp.Value = "left-arrow-callout"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT_RIGHT
			$tProp.Value = "left-right-arrow-callout"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_RIGHT
			$tProp.Value = "right-arrow-callout"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP
			$tProp.Value = "up-arrow-callout"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_DOWN
			$tProp.Value = "up-down-arrow-callout"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT
			$tProp.Value = "mso-spt100"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_CIRCULAR
			$tProp.Value = "circular-arrow"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT
			$tProp.Value = "corner-right-arrow" ; "non-primitive"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_DOWN
			$tProp.Value = "down-arrow"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_LEFT
			$tProp.Value = "left-arrow"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_LEFT_RIGHT
			$tProp.Value = "left-right-arrow"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_NOTCHED_RIGHT
			$tProp.Value = "notched-right-arrow"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_RIGHT
			$tProp.Value = "right-arrow"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT
			$tProp.Value = "split-arrow" ; "non-primitive"??

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_S_SHAPED
			$tProp.Value = "s-sharped-arrow" ; "non-primitive"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_SPLIT
			$tProp.Value = "split-arrow" ; "non-primitive"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_STRIPED_RIGHT
			$tProp.Value = "striped-right-arrow" ; "mso-spt100"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_UP
			$tProp.Value = "up-arrow"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_UP_DOWN
			$tProp.Value = "up-down-arrow"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_UP_RIGHT
			$tProp.Value = "up-right-arrow-callout" ; "mso-spt89"

		Case $LOW_SHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN
			$tProp.Value = "up-right-down-arrow" ; "mso-spt100"

			$tProp2 = __LO_SetPropertyValue("MirroredX", True) ; Shape is an up and left arrow without this Property.
			If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

			ReDim $atCusShapeGeo[2]
			$atCusShapeGeo[1] = $tProp2

		Case $LOW_SHAPE_TYPE_ARROWS_CHEVRON
			$tProp.Value = "chevron"

		Case $LOW_SHAPE_TYPE_ARROWS_PENTAGON
			$tProp.Value = "pentagon-right"
	EndSwitch

	$atCusShapeGeo[0] = $tProp
	$oShape.CustomShapeGeometry = $atCusShapeGeo

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tPos.X = 0
	$tPos.Y = 0

	$oShape.Position = $tPos

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	$oShape.Name = __LOWriter_GetShapeName($oDoc, "Shape ")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOWriter_Shape_CreateArrow

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Shape_CreateBasic
; Description ...: Create a Basic type Shape.
; Syntax ........: __LOWriter_Shape_CreateBasic($oDoc, $iWidth, $iHeight, $iShapeType)
; Parameters ....: $oDoc                - an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iWidth              - an integer value. The Shape's Width in Micrometers.
;                  $iHeight             - an integer value. The Shape's Height in Micrometers.
;                  $iShapeType          - an integer value (26-49). The Type of shape to create. See $LOW_SHAPE_TYPE_BASIC_* as defined in LibreOfficeWriter_Constants.au3
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iShapeType not an Integer
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" or "com.sun.star.drawing.EllipseShape" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to create a unique Shape name.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the Size Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The following shapes are not implemented into LibreOffice as of L.O. Version 7.3.4.2 for automation, and thus will not work:
;                  $LOW_SHAPE_TYPE_BASIC_CIRCLE_PIE, $LOW_SHAPE_TYPE_BASIC_FRAME
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_Shape_CreateBasic($oDoc, $iWidth, $iHeight, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape
	Local $tProp, $tSize, $tPos
	Local $atCusShapeGeo[1]
	Local $iCircleKind_CUT = 2 ; a circle with a cut connected by a line.
	Local $iCircleKind_ARC = 3 ; a circle with an open cut.

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If ($iShapeType = $LOW_SHAPE_TYPE_BASIC_CIRCLE_SEGMENT) Or ($iShapeType = $LOW_SHAPE_TYPE_BASIC_ARC) Then ; These two shapes need special procedures.
		$oShape = $oDoc.createInstance("com.sun.star.drawing.EllipseShape")
		If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Else
		$oShape = $oDoc.createInstance("com.sun.star.drawing.CustomShape")
		If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tProp = __LO_SetPropertyValue("Type", "")
		If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		$oShape.Name = __LOWriter_GetShapeName($oDoc, "Shape ")
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	Switch $iShapeType
		Case $LOW_SHAPE_TYPE_BASIC_ARC
			$oShape.FillColor = $LO_COLOR_OFF

			$oShape.Name = __LOWriter_GetShapeName($oDoc, "Elliptical arc ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			$oShape.CircleKind = $iCircleKind_ARC
			$oShape.CircleStartAngle = 0
			$oShape.CircleEndAngle = 25000

		Case $LOW_SHAPE_TYPE_BASIC_ARC_BLOCK
			$tProp.Value = "block-arc"

		Case $LOW_SHAPE_TYPE_BASIC_CIRCLE_PIE
			$tProp.Value = "circle-pie" ; "mso-spt100"

		Case $LOW_SHAPE_TYPE_BASIC_CIRCLE_SEGMENT
			$oShape.Name = __LOWriter_GetShapeName($oDoc, "Ellipse Segment ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			$oShape.CircleKind = $iCircleKind_CUT
			$oShape.CircleStartAngle = 0
			$oShape.CircleEndAngle = 25000

		Case $LOW_SHAPE_TYPE_BASIC_CROSS
			$tProp.Value = "cross"

		Case $LOW_SHAPE_TYPE_BASIC_CUBE
			$tProp.Value = "cube"

		Case $LOW_SHAPE_TYPE_BASIC_CYLINDER
			$tProp.Value = "can"

		Case $LOW_SHAPE_TYPE_BASIC_DIAMOND
			$tProp.Value = "diamond"

		Case $LOW_SHAPE_TYPE_BASIC_ELLIPSE, $LOW_SHAPE_TYPE_BASIC_CIRCLE
			$tProp.Value = "ellipse"

		Case $LOW_SHAPE_TYPE_BASIC_FOLDED_CORNER
			$tProp.Value = "paper"

		Case $LOW_SHAPE_TYPE_BASIC_FRAME
			$tProp.Value = "frame" ; Not working

		Case $LOW_SHAPE_TYPE_BASIC_HEXAGON
			$tProp.Value = "hexagon"

		Case $LOW_SHAPE_TYPE_BASIC_OCTAGON
			$tProp.Value = "octagon"

		Case $LOW_SHAPE_TYPE_BASIC_PARALLELOGRAM
			$tProp.Value = "parallelogram"

		Case $LOW_SHAPE_TYPE_BASIC_RECTANGLE, $LOW_SHAPE_TYPE_BASIC_SQUARE
			$tProp.Value = "rectangle"

		Case $LOW_SHAPE_TYPE_BASIC_RECTANGLE_ROUNDED, $LOW_SHAPE_TYPE_BASIC_SQUARE_ROUNDED
			$tProp.Value = "round-rectangle"

		Case $LOW_SHAPE_TYPE_BASIC_REGULAR_PENTAGON
			$tProp.Value = "pentagon"

		Case $LOW_SHAPE_TYPE_BASIC_RING
			$tProp.Value = "ring"

		Case $LOW_SHAPE_TYPE_BASIC_TRAPEZOID
			$tProp.Value = "trapezoid"

		Case $LOW_SHAPE_TYPE_BASIC_TRIANGLE_ISOSCELES
			$tProp.Value = "isosceles-triangle"

		Case $LOW_SHAPE_TYPE_BASIC_TRIANGLE_RIGHT
			$tProp.Value = "right-triangle"
	EndSwitch

	If ($iShapeType <> $LOW_SHAPE_TYPE_BASIC_CIRCLE_SEGMENT) And ($iShapeType <> $LOW_SHAPE_TYPE_BASIC_ARC) Then
		$atCusShapeGeo[0] = $tProp
		$oShape.CustomShapeGeometry = $atCusShapeGeo
	EndIf

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$tPos.X = 0
	$tPos.Y = 0

	$oShape.Position = $tPos

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOWriter_Shape_CreateBasic

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Shape_CreateCallout
; Description ...: Create a Callout type Shape.
; Syntax ........: __LOWriter_Shape_CreateCallout($oDoc, $iWidth, $iHeight, $iShapeType)
; Parameters ....: $oDoc                - an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iWidth              - an integer value. The Shape's Width in Micrometers.
;                  $iHeight             - an integer value. The Shape's Height in Micrometers.
;                  $iShapeType          - an integer value (50-56). The Type of shape to create. See $LOW_SHAPE_TYPE_CALLOUT_* as defined in LibreOfficeWriter_Constants.au3
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iShapeType not an Integer
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve the Size Structure.
;                  @Error 3 @Extended 3 Return 0 = Failed to create a unique Shape name.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_Shape_CreateCallout($oDoc, $iWidth, $iHeight, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape
	Local $tProp, $tSize, $tPos
	Local $atCusShapeGeo[1]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oShape = $oDoc.createInstance("com.sun.star.drawing.CustomShape")
	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tProp = __LO_SetPropertyValue("Type", "")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Switch $iShapeType
		Case $LOW_SHAPE_TYPE_CALLOUT_CLOUD
			$tProp.Value = "cloud-callout"

		Case $LOW_SHAPE_TYPE_CALLOUT_LINE_1
			$tProp.Value = "line-callout-1"

		Case $LOW_SHAPE_TYPE_CALLOUT_LINE_2
			$tProp.Value = "line-callout-2"

		Case $LOW_SHAPE_TYPE_CALLOUT_LINE_3
			$tProp.Value = "line-callout-3"

		Case $LOW_SHAPE_TYPE_CALLOUT_RECTANGULAR
			$tProp.Value = "rectangular-callout"

		Case $LOW_SHAPE_TYPE_CALLOUT_RECTANGULAR_ROUNDED
			$tProp.Value = "round-rectangular-callout"

		Case $LOW_SHAPE_TYPE_CALLOUT_ROUND
			$tProp.Value = "round-callout"
	EndSwitch

	$atCusShapeGeo[0] = $tProp
	$oShape.CustomShapeGeometry = $atCusShapeGeo

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tPos.X = 0
	$tPos.Y = 0

	$oShape.Position = $tPos

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	$oShape.Name = __LOWriter_GetShapeName($oDoc, "Shape ")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOWriter_Shape_CreateCallout

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Shape_CreateFlowchart
; Description ...: Create a FlowChart type Shape.
; Syntax ........: __LOWriter_Shape_CreateFlowchart($oDoc, $iWidth, $iHeight, $iShapeType)
; Parameters ....: $oDoc                - an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iWidth              - an integer value. The Shape's Width in Micrometers.
;                  $iHeight             - an integer value. The Shape's Height in Micrometers.
;                  $iShapeType          - an integer value (57-84). The Type of shape to create. See $LOW_SHAPE_TYPE_FLOWCHART_* as defined in LibreOfficeWriter_Constants.au3
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iShapeType not an Integer
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve the Size Structure.
;                  @Error 3 @Extended 3 Return 0 = Failed to create a unique Shape name.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_Shape_CreateFlowchart($oDoc, $iWidth, $iHeight, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape
	Local $tProp, $tSize, $tPos
	Local $atCusShapeGeo[1]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oShape = $oDoc.createInstance("com.sun.star.drawing.CustomShape")
	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tProp = __LO_SetPropertyValue("Type", "")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Switch $iShapeType
		Case $LOW_SHAPE_TYPE_FLOWCHART_CARD
			$tProp.Value = "flowchart-card"

		Case $LOW_SHAPE_TYPE_FLOWCHART_COLLATE
			$tProp.Value = "flowchart-collate"

		Case $LOW_SHAPE_TYPE_FLOWCHART_CONNECTOR
			$tProp.Value = "flowchart-connector"

		Case $LOW_SHAPE_TYPE_FLOWCHART_CONNECTOR_OFF_PAGE
			$tProp.Value = "flowchart-off-page-connector"

		Case $LOW_SHAPE_TYPE_FLOWCHART_DATA
			$tProp.Value = "flowchart-data"

		Case $LOW_SHAPE_TYPE_FLOWCHART_DECISION
			$tProp.Value = "flowchart-decision"

		Case $LOW_SHAPE_TYPE_FLOWCHART_DELAY
			$tProp.Value = "flowchart-delay"

		Case $LOW_SHAPE_TYPE_FLOWCHART_DIRECT_ACCESS_STORAGE
			$tProp.Value = "flowchart-direct-access-storage"

		Case $LOW_SHAPE_TYPE_FLOWCHART_DISPLAY
			$tProp.Value = "flowchart-display"

		Case $LOW_SHAPE_TYPE_FLOWCHART_DOCUMENT
			$tProp.Value = "flowchart-document"

		Case $LOW_SHAPE_TYPE_FLOWCHART_EXTRACT
			$tProp.Value = "flowchart-extract"

		Case $LOW_SHAPE_TYPE_FLOWCHART_INTERNAL_STORAGE
			$tProp.Value = "flowchart-internal-storage"

		Case $LOW_SHAPE_TYPE_FLOWCHART_MAGNETIC_DISC
			$tProp.Value = "flowchart-magnetic-disk"

		Case $LOW_SHAPE_TYPE_FLOWCHART_MANUAL_INPUT
			$tProp.Value = "flowchart-manual-input"

		Case $LOW_SHAPE_TYPE_FLOWCHART_MANUAL_OPERATION
			$tProp.Value = "flowchart-manual-operation"

		Case $LOW_SHAPE_TYPE_FLOWCHART_MERGE
			$tProp.Value = "flowchart-merge"

		Case $LOW_SHAPE_TYPE_FLOWCHART_MULTIDOCUMENT
			$tProp.Value = "flowchart-multidocument"

		Case $LOW_SHAPE_TYPE_FLOWCHART_OR
			$tProp.Value = "flowchart-or"

		Case $LOW_SHAPE_TYPE_FLOWCHART_PREPARATION
			$tProp.Value = "flowchart-preparation"

		Case $LOW_SHAPE_TYPE_FLOWCHART_PROCESS
			$tProp.Value = "flowchart-process"

		Case $LOW_SHAPE_TYPE_FLOWCHART_PROCESS_ALTERNATE
			$tProp.Value = "flowchart-alternate-process"

		Case $LOW_SHAPE_TYPE_FLOWCHART_PROCESS_PREDEFINED
			$tProp.Value = "flowchart-predefined-process"

		Case $LOW_SHAPE_TYPE_FLOWCHART_PUNCHED_TAPE
			$tProp.Value = "flowchart-punched-tape"

		Case $LOW_SHAPE_TYPE_FLOWCHART_SEQUENTIAL_ACCESS
			$tProp.Value = "flowchart-sequential-access"

		Case $LOW_SHAPE_TYPE_FLOWCHART_SORT
			$tProp.Value = "flowchart-sort"

		Case $LOW_SHAPE_TYPE_FLOWCHART_STORED_DATA
			$tProp.Value = "flowchart-stored-data"

		Case $LOW_SHAPE_TYPE_FLOWCHART_SUMMING_JUNCTION
			$tProp.Value = "flowchart-summing-junction"

		Case $LOW_SHAPE_TYPE_FLOWCHART_TERMINATOR
			$tProp.Value = "flowchart-terminator"
	EndSwitch

	$atCusShapeGeo[0] = $tProp
	$oShape.CustomShapeGeometry = $atCusShapeGeo

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tPos.X = 0
	$tPos.Y = 0

	$oShape.Position = $tPos

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	$oShape.Name = __LOWriter_GetShapeName($oDoc, "Shape ")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOWriter_Shape_CreateFlowchart

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Shape_CreateLine
; Description ...: Create a Line type Shape.
; Syntax ........: __LOWriter_Shape_CreateLine($oDoc, $iWidth, $iHeight, $iShapeType)
; Parameters ....: $oDoc                - an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iWidth              - an integer value. The Shape's Width in Micrometers.
;                  $iHeight             - an integer value. The Shape's Height in Micrometers.
;                  $iShapeType          - an integer value (85-92). The Type of shape to create. See $LOW_SHAPE_TYPE_LINE_* as defined in LibreOfficeWriter_Constants.au3
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iShapeType not an Integer
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create the requested Line type Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a Position structure.
;                  @Error 2 @Extended 3 Return 0 = Failed to create "com.sun.star.drawing.PolyPolygonBezierCoords" Structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to create a unique Shape name.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the Size Structure.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_Shape_CreateLine($oDoc, $iWidth, $iHeight, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape
	Local $tSize, $tPos, $tPolyCoords
	Local $atPoint[0], $aiFlags[0]
	Local $avArray[1]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$tPolyCoords = __LO_CreateStruct("com.sun.star.drawing.PolyPolygonBezierCoords")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	Switch $iShapeType
		Case $LOW_SHAPE_TYPE_LINE_CURVE
			$oShape = $oDoc.createInstance("com.sun.star.drawing.OpenBezierShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			ReDim $atPoint[4]
			ReDim $aiFlags[4]

			$atPoint[0] = __LOWriter_CreatePoint(0, 0)
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOWriter_CreatePoint(Int($iWidth / 2), $iHeight)
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[2] = __LOWriter_CreatePoint(Int($iWidth / 2), Int($iHeight / 2))
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[3] = __LOWriter_CreatePoint($iWidth, 0)
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOW_SHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOW_SHAPE_POINT_TYPE_CONTROL
			$aiFlags[2] = $LOW_SHAPE_POINT_TYPE_CONTROL
			$aiFlags[3] = $LOW_SHAPE_POINT_TYPE_NORMAL

			$oShape.Name = __LOWriter_GetShapeName($oDoc, "Bézier curve ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOW_SHAPE_TYPE_LINE_CURVE_FILLED
			$oShape = $oDoc.createInstance("com.sun.star.drawing.ClosedBezierShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			ReDim $atPoint[4]
			ReDim $aiFlags[4]

			$atPoint[0] = __LOWriter_CreatePoint(0, 0)
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOWriter_CreatePoint(Int($iWidth / 2), $iHeight)
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[2] = __LOWriter_CreatePoint(Int($iWidth / 2), Int($iHeight / 2))
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[3] = __LOWriter_CreatePoint($iWidth, 0)
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOW_SHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOW_SHAPE_POINT_TYPE_CONTROL
			$aiFlags[2] = $LOW_SHAPE_POINT_TYPE_CONTROL
			$aiFlags[3] = $LOW_SHAPE_POINT_TYPE_NORMAL

			$oShape.Name = __LOWriter_GetShapeName($oDoc, "Bézier curve ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			$oShape.FillColor = 7512015 ; Light blue

		Case $LOW_SHAPE_TYPE_LINE_FREEFORM_LINE
			$oShape = $oDoc.createInstance("com.sun.star.drawing.OpenBezierShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			ReDim $atPoint[3]
			ReDim $aiFlags[3]

			$atPoint[0] = __LOWriter_CreatePoint(0, 0)
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOWriter_CreatePoint(Int($iWidth / 2), Int($iHeight / 2))
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[2] = __LOWriter_CreatePoint(Int($iWidth), Int($iHeight))
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOW_SHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOW_SHAPE_POINT_TYPE_CONTROL
			$aiFlags[2] = $LOW_SHAPE_POINT_TYPE_NORMAL

			$oShape.Name = __LOWriter_GetShapeName($oDoc, "Bézier curve ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Case $LOW_SHAPE_TYPE_LINE_FREEFORM_LINE_FILLED
			$oShape = $oDoc.createInstance("com.sun.star.drawing.ClosedBezierShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			ReDim $atPoint[4]
			ReDim $aiFlags[4]

			$atPoint[0] = __LOWriter_CreatePoint(0, 0)
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOWriter_CreatePoint($iWidth + Int(($iWidth / 8)), Int(($iHeight / 2)))
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[2] = __LOWriter_CreatePoint(Int($iWidth), Int($iHeight))
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[3] = __LOWriter_CreatePoint(0, 0)
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOW_SHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOW_SHAPE_POINT_TYPE_CONTROL
			$aiFlags[2] = $LOW_SHAPE_POINT_TYPE_NORMAL
			$aiFlags[3] = $LOW_SHAPE_POINT_TYPE_NORMAL

			$oShape.Name = __LOWriter_GetShapeName($oDoc, "Bézier curve ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			$oShape.FillColor = 7512015 ; Light blue

		Case $LOW_SHAPE_TYPE_LINE_LINE
			$oShape = $oDoc.createInstance("com.sun.star.drawing.LineShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			ReDim $atPoint[2]
			ReDim $aiFlags[2]

			$atPoint[0] = __LOWriter_CreatePoint(0, 0)
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOWriter_CreatePoint(Int($iWidth), Int($iHeight))
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOW_SHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOW_SHAPE_POINT_TYPE_NORMAL

			$oShape.Name = __LOWriter_GetShapeName($oDoc, "Line ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Case $LOW_SHAPE_TYPE_LINE_POLYGON, $LOW_SHAPE_TYPE_LINE_POLYGON_45
			$oShape = $oDoc.createInstance("com.sun.star.drawing.PolyPolygonShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			ReDim $atPoint[5]
			ReDim $aiFlags[5]

			$atPoint[0] = __LOWriter_CreatePoint(0, 0)
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOWriter_CreatePoint(Int($iWidth), 0)
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[2] = __LOWriter_CreatePoint(Int($iWidth), Int($iHeight))
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[3] = __LOWriter_CreatePoint(0, Int($iHeight))
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[4] = __LOWriter_CreatePoint(0, 0)
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOW_SHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOW_SHAPE_POINT_TYPE_NORMAL
			$aiFlags[2] = $LOW_SHAPE_POINT_TYPE_NORMAL
			$aiFlags[3] = $LOW_SHAPE_POINT_TYPE_NORMAL
			$aiFlags[4] = $LOW_SHAPE_POINT_TYPE_NORMAL

			$oShape.Name = __LOWriter_GetShapeName($oDoc, "Polygon 4 corners ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOW_SHAPE_TYPE_LINE_POLYGON_45_FILLED
			$oShape = $oDoc.createInstance("com.sun.star.drawing.PolyPolygonShape")
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			ReDim $atPoint[5]
			ReDim $aiFlags[5]

			$atPoint[0] = __LOWriter_CreatePoint(0, 0)
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[1] = __LOWriter_CreatePoint(Int($iWidth), 0)
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[2] = __LOWriter_CreatePoint(Int($iWidth), Int($iHeight))
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[3] = __LOWriter_CreatePoint(0, Int($iHeight))
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$atPoint[4] = __LOWriter_CreatePoint(0, 0)
			If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

			$aiFlags[0] = $LOW_SHAPE_POINT_TYPE_NORMAL
			$aiFlags[1] = $LOW_SHAPE_POINT_TYPE_NORMAL
			$aiFlags[2] = $LOW_SHAPE_POINT_TYPE_NORMAL
			$aiFlags[3] = $LOW_SHAPE_POINT_TYPE_NORMAL
			$aiFlags[4] = $LOW_SHAPE_POINT_TYPE_NORMAL

			$oShape.Name = __LOWriter_GetShapeName($oDoc, "Polygon 4 corners ")
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			$oShape.FillColor = 7512015 ; Light blue
	EndSwitch

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$tPos.X = 0
	$tPos.Y = 0

	$oShape.Position = $tPos

	$avArray[0] = $atPoint
	$tPolyCoords.Coordinates = $avArray

	$avArray[0] = $aiFlags
	$tPolyCoords.Flags = $avArray

	$oShape.PolyPolygonBezier = $tPolyCoords

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOWriter_Shape_CreateLine

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Shape_CreateStars
; Description ...: Create a Star or Banner type Shape.
; Syntax ........: __LOWriter_Shape_CreateStars($oDoc, $iWidth, $iHeight, $iShapeType)
; Parameters ....: $oDoc                - an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iWidth              - an integer value. The Shape's Width in Micrometers.
;                  $iHeight             - an integer value. The Shape's Height in Micrometers.
;                  $iShapeType          - an integer value (93-104). The Type of shape to create. See $LOW_SHAPE_TYPE_STARS_* as defined in LibreOfficeWriter_Constants.au3
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iShapeType not an Integer
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve the Size Structure.
;                  @Error 3 @Extended 3 Return 0 = Failed to create a unique Shape name.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The following shapes are not implemented into LibreOffice as of L.O. Version 7.3.4.2 for automation, and thus will not work:
;                  $LOW_SHAPE_TYPE_STARS_6_POINT, $LOW_SHAPE_TYPE_STARS_12_POINT, $LOW_SHAPE_TYPE_STARS_SIGNET, $LOW_SHAPE_TYPE_STARS_6_POINT_CONCAVE.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_Shape_CreateStars($oDoc, $iWidth, $iHeight, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape
	Local $tProp, $tSize, $tPos
	Local $atCusShapeGeo[1]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oShape = $oDoc.createInstance("com.sun.star.drawing.CustomShape")
	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tProp = __LO_SetPropertyValue("Type", "")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Switch $iShapeType
		Case $LOW_SHAPE_TYPE_STARS_4_POINT
			$tProp.Value = "star4"

		Case $LOW_SHAPE_TYPE_STARS_5_POINT
			$tProp.Value = "star5"

		Case $LOW_SHAPE_TYPE_STARS_6_POINT
			$tProp.Value = "star6" ; "non-primitive"

		Case $LOW_SHAPE_TYPE_STARS_6_POINT_CONCAVE
			$tProp.Value = "concave-star6" ; "non-primitive"

		Case $LOW_SHAPE_TYPE_STARS_8_POINT
			$tProp.Value = "star8"

		Case $LOW_SHAPE_TYPE_STARS_12_POINT
			$tProp.Value = "star12" ; "non-primitive"

		Case $LOW_SHAPE_TYPE_STARS_24_POINT
			$tProp.Value = "star24"

		Case $LOW_SHAPE_TYPE_STARS_DOORPLATE
			$tProp.Value = "mso-spt21" ; "doorplate"

		Case $LOW_SHAPE_TYPE_STARS_EXPLOSION
			$tProp.Value = "bang"

		Case $LOW_SHAPE_TYPE_STARS_SCROLL_HORIZONTAL
			$tProp.Value = "horizontal-scroll"

		Case $LOW_SHAPE_TYPE_STARS_SCROLL_VERTICAL
			$tProp.Value = "vertical-scroll"

		Case $LOW_SHAPE_TYPE_STARS_SIGNET
			$tProp.Value = "signet" ; "non-primitive"
	EndSwitch

	$atCusShapeGeo[0] = $tProp
	$oShape.CustomShapeGeometry = $atCusShapeGeo

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tPos.X = 0
	$tPos.Y = 0

	$oShape.Position = $tPos

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	$oShape.Name = __LOWriter_GetShapeName($oDoc, "Shape ")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOWriter_Shape_CreateStars

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Shape_CreateSymbol
; Description ...: Create a Symbol type Shape.
; Syntax ........: __LOWriter_Shape_CreateSymbol($oDoc, $iWidth, $iHeight, $iShapeType)
; Parameters ....: $oDoc                - an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iWidth              - an integer value. The Shape's Width in Micrometers.
;                  $iHeight             - an integer value. The Shape's Height in Micrometers.
;                  $iShapeType          - an integer value (105-122). The Type of shape to create. See $LOW_SHAPE_TYPE_SYMBOL_* as defined in LibreOfficeWriter_Constants.au3
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iWidth not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iHeight not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $iShapeType not an Integer
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.drawing.CustomShape" Object.
;                  @Error 2 @Extended 2 Return 0 = Failed to create a property structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve the Position Structure.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve the Size Structure.
;                  @Error 3 @Extended 3 Return 0 = Failed to create a unique Shape name.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning the newly created shape.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The following shapes are not implemented into LibreOffice as of L.O. Version 7.3.4.2 for automation, and thus will not work:
;                  $LOW_SHAPE_TYPE_SYMBOL_CLOUD, $LOW_SHAPE_TYPE_SYMBOL_FLOWER, $LOW_SHAPE_TYPE_SYMBOL_PUZZLE, $LOW_SHAPE_TYPE_SYMBOL_BEVEL_OCTAGON, $LOW_SHAPE_TYPE_SYMBOL_BEVEL_DIAMOND
;                  The following shape is visually different from the manually inserted one in L.O. 7.3.4.2:
;                  $LOW_SHAPE_TYPE_SYMBOL_LIGHTNING
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_Shape_CreateSymbol($oDoc, $iWidth, $iHeight, $iShapeType)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape
	Local $tProp, $tSize, $tPos
	Local $atCusShapeGeo[1]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oShape = $oDoc.createInstance("com.sun.star.drawing.CustomShape")
	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tProp = __LO_SetPropertyValue("Type", "")
	If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Switch $iShapeType
		Case $LOW_SHAPE_TYPE_SYMBOL_BEVEL_DIAMOND
			$tProp.Value = "col-502ad400"

		Case $LOW_SHAPE_TYPE_SYMBOL_BEVEL_OCTAGON
			$tProp.Value = "col-60da8460"

		Case $LOW_SHAPE_TYPE_SYMBOL_BEVEL_SQUARE
			$tProp.Value = "quad-bevel"

		Case $LOW_SHAPE_TYPE_SYMBOL_BRACE_DOUBLE
			$tProp.Value = "brace-pair"
			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOW_SHAPE_TYPE_SYMBOL_BRACE_LEFT
			$tProp.Value = "left-brace"
			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOW_SHAPE_TYPE_SYMBOL_BRACE_RIGHT
			$tProp.Value = "right-brace"
			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOW_SHAPE_TYPE_SYMBOL_BRACKET_DOUBLE
			$tProp.Value = "bracket-pair"
			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOW_SHAPE_TYPE_SYMBOL_BRACKET_LEFT
			$tProp.Value = "left-bracket"
			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOW_SHAPE_TYPE_SYMBOL_BRACKET_RIGHT
			$tProp.Value = "right-bracket"
			$oShape.FillColor = $LO_COLOR_OFF

		Case $LOW_SHAPE_TYPE_SYMBOL_CLOUD
			;~ Custom Shape Geometry Type = "non-primitive" ???? Try "cloud"
			$tProp.Value = "cloud"

		Case $LOW_SHAPE_TYPE_SYMBOL_FLOWER
			;~ Custom Shape Geometry Type = "non-primitive" ???? Try "flower"
			$tProp.Value = "flower"

		Case $LOW_SHAPE_TYPE_SYMBOL_HEART
			$tProp.Value = "heart"

		Case $LOW_SHAPE_TYPE_SYMBOL_LIGHTNING
			;~ Custom Shape Geometry Type = "non-primitive" ???? Try "lightning"
			$tProp.Value = "lightning"

		Case $LOW_SHAPE_TYPE_SYMBOL_MOON
			$tProp.Value = "moon"

		Case $LOW_SHAPE_TYPE_SYMBOL_SMILEY
			$tProp.Value = "smiley"

		Case $LOW_SHAPE_TYPE_SYMBOL_SUN
			$tProp.Value = "sun"

		Case $LOW_SHAPE_TYPE_SYMBOL_PROHIBITED
			$tProp.Value = "forbidden"

		Case $LOW_SHAPE_TYPE_SYMBOL_PUZZLE
			$tProp.Value = "puzzle"
	EndSwitch

	$atCusShapeGeo[0] = $tProp
	$oShape.CustomShapeGeometry = $atCusShapeGeo

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tPos.X = 0
	$tPos.Y = 0

	$oShape.Position = $tPos

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	$oShape.Name = __LOWriter_GetShapeName($oDoc, "Shape ")
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>__LOWriter_Shape_CreateSymbol

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Shape_GetCustomType
; Description ...: Return the Shape Type Constant corresponding to the Custom Shape Type string.
; Syntax ........: __LOWriter_Shape_GetCustomType($sCusShapeType)
; Parameters ....: $sCusShapeType       - a string value. The Returned Custom Shape Type Value from CustomShapeGeometry Array of properties.
; Return values .: Success: Integer or -1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sCusShapeType not a String.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. Custom Shape Type was successfully identified. Returning the Constant value of the Shape, see Constants $LOW_SHAPE_TYPE_* as defined in LibreOfficeWriter_Constants.au3
;                  @Error 0 @Extended 0 Return -1 = Success. Custom Shape is of an unimplemented type that has an ambiguous name, and cannot be identified. See Remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Some shapes are not implemented, or not fully implemented into LibreOffice for automation, consequently they do not have appropriate type names as of yet. Many have simply ambiguous names, such as "non-primitive".
;                  Because of this the following shape types cannot be identified, and this function will return -1:
;                   $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT, known as "mso-spt100".
;                   $LOW_SHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT, known as "non-primitive", should be "corner-right-arrow".
;                   $LOW_SHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT, known as "non-primitive", should be "right-left-arrow".
;                   $LOW_SHAPE_TYPE_ARROWS_ARROW_S_SHAPED, known as "non-primitive", should be "s-sharped-arrow".
;                   $LOW_SHAPE_TYPE_ARROWS_ARROW_SPLIT, known as "non-primitive", should be "split-arrow".
;                   $LOW_SHAPE_TYPE_ARROWS_ARROW_STRIPED_RIGHT, known as "mso-spt100", should be "striped-right-arrow".
;                   $LOW_SHAPE_TYPE_ARROWS_ARROW_UP_RIGHT, known as "mso-spt89", should be "up-right-arrow-callout".
;                   $LOW_SHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN, known as "mso-spt100", should be "up-right-down-arrow".
;                   $LOW_SHAPE_TYPE_BASIC_CIRCLE_PIE, known as "mso-spt100", should be "circle-pie".
;                   $LOW_SHAPE_TYPE_STARS_6_POINT, known as "non-primitive", should be "star6".
;                   $LOW_SHAPE_TYPE_STARS_6_POINT_CONCAVE, known as "non-primitive", should be "concave-star6".
;                   $LOW_SHAPE_TYPE_STARS_12_POINT, known as "non-primitive", should be "star12".
;                   $LOW_SHAPE_TYPE_STARS_SIGNET, known as "non-primitive", should be "signet".
;                   $LOW_SHAPE_TYPE_SYMBOL_CLOUD, known as "non-primitive", should be "cloud"?
;                   $LOW_SHAPE_TYPE_SYMBOL_FLOWER, known as "non-primitive", should be "flower"?
;                   $LOW_SHAPE_TYPE_SYMBOL_LIGHTNING, known as "non-primitive", should be "lightning".
;                  The following Shapes implement the same type names, and are consequently indistinguishable:
;                   $LOW_SHAPE_TYPE_BASIC_CIRCLE, $LOW_SHAPE_TYPE_BASIC_ELLIPSE (The Value of $LOW_SHAPE_TYPE_BASIC_CIRCLE is returned for either one.)
;                   $LOW_SHAPE_TYPE_BASIC_SQUARE, $LOW_SHAPE_TYPE_BASIC_RECTANGLE (The Value of $LOW_SHAPE_TYPE_BASIC_SQUARE is returned for either one.)
;                   $LOW_SHAPE_TYPE_BASIC_SQUARE_ROUNDED, $LOW_SHAPE_TYPE_BASIC_RECTANGLE_ROUNDED (The Value of $LOW_SHAPE_TYPE_BASIC_SQUARE_ROUNDED is returned for either one.)
;                  The following Shapes have strange names that may change in the future, but currently are able to be identified:
;                   $LOW_SHAPE_TYPE_STARS_DOORPLATE, known as, "mso-spt21", should be "doorplate"
;                   $LOW_SHAPE_TYPE_SYMBOL_BEVEL_DIAMOND, known as, "col-502ad400", should be ??
;                   $LOW_SHAPE_TYPE_SYMBOL_BEVEL_OCTAGON, known as, "col-60da8460", should be ??
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_Shape_GetCustomType($sCusShapeType)
	If Not IsString($sCusShapeType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Switch $sCusShapeType
		Case "quad-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_4_WAY)

		Case "quad-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_4_WAY)

		Case "down-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_DOWN)

		Case "left-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT)

		Case "left-right-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT_RIGHT)

		Case "right-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_RIGHT)

		Case "up-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP)

		Case "up-down-arrow-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_DOWN)

			;~ 	Case "mso-spt100" ; Can't include this one as other shapes return mso-spt100 also
			;~ Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT)

		Case "circular-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_CIRCULAR)

		Case "corner-right-arrow" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT)

		Case "down-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_DOWN)

		Case "left-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_LEFT)

		Case "left-right-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_LEFT_RIGHT)

		Case "notched-right-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_NOTCHED_RIGHT)

		Case "right-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_RIGHT)

		Case "right-left-arrow" ; "non-primitive"??

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT)

		Case "s-sharped-arrow" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_S_SHAPED)

		Case "split-arrow" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_SPLIT)

		Case "striped-right-arrow" ; "mso-spt100"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_STRIPED_RIGHT)

		Case "up-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_UP)

		Case "up-down-arrow"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_UP_DOWN)

		Case "up-right-arrow-callout", "mso-spt89" ; "mso-spt89"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_UP_RIGHT)

		Case "up-right-down-arrow" ; "mso-spt100"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN)

		Case "chevron"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_CHEVRON)

		Case "pentagon-right"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_ARROWS_PENTAGON)

		Case "block-arc"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_BASIC_ARC_BLOCK)

		Case "circle-pie" ; "mso-spt100"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_BASIC_CIRCLE_PIE)

		Case "cross"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_BASIC_CROSS)

		Case "cube"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_BASIC_CUBE)

		Case "can"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_BASIC_CYLINDER)

		Case "diamond"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_BASIC_DIAMOND)

		Case "ellipse"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_BASIC_CIRCLE)
			;~ $LOW_SHAPE_TYPE_BASIC_ELLIPSE

		Case "paper"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_BASIC_FOLDED_CORNER)

		Case "frame" ; Not working

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_BASIC_FRAME)

		Case "hexagon"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_BASIC_HEXAGON)

		Case "octagon"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_BASIC_OCTAGON)

		Case "parallelogram"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_BASIC_PARALLELOGRAM)

		Case "rectangle"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_BASIC_SQUARE)
			;~ $LOW_SHAPE_TYPE_BASIC_RECTANGLE

		Case "round-rectangle"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_BASIC_SQUARE_ROUNDED)
			;~ $LOW_SHAPE_TYPE_BASIC_RECTANGLE_ROUNDED

		Case "pentagon"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_BASIC_REGULAR_PENTAGON)

		Case "ring"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_BASIC_RING)

		Case "trapezoid"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_BASIC_TRAPEZOID)

		Case "isosceles-triangle"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_BASIC_TRIANGLE_ISOSCELES)

		Case "right-triangle"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_BASIC_TRIANGLE_RIGHT)

		Case "cloud-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_CALLOUT_CLOUD)

		Case "line-callout-1"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_CALLOUT_LINE_1)

		Case "line-callout-2"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_CALLOUT_LINE_2)

		Case "line-callout-3"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_CALLOUT_LINE_3)

		Case "rectangular-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_CALLOUT_RECTANGULAR)

		Case "round-rectangular-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_CALLOUT_RECTANGULAR_ROUNDED)

		Case "round-callout"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_CALLOUT_ROUND)

		Case "flowchart-card"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_CARD)

		Case "flowchart-collate"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_COLLATE)

		Case "flowchart-connector"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_CONNECTOR)

		Case "flowchart-off-page-connector"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_CONNECTOR_OFF_PAGE)

		Case "flowchart-data"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_DATA)

		Case "flowchart-decision"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_DECISION)

		Case "flowchart-delay"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_DELAY)

		Case "flowchart-direct-access-storage"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_DIRECT_ACCESS_STORAGE)

		Case "flowchart-display"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_DISPLAY)

		Case "flowchart-document"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_DOCUMENT)

		Case "flowchart-extract"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_EXTRACT)

		Case "flowchart-internal-storage"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_INTERNAL_STORAGE)

		Case "flowchart-magnetic-disk"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_MAGNETIC_DISC)

		Case "flowchart-manual-input"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_MANUAL_INPUT)

		Case "flowchart-manual-operation"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_MANUAL_OPERATION)

		Case "flowchart-merge"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_MERGE)

		Case "flowchart-multidocument"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_MULTIDOCUMENT)

		Case "flowchart-or"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_OR)

		Case "flowchart-preparation"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_PREPARATION)

		Case "flowchart-process"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_PROCESS)

		Case "flowchart-alternate-process"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_PROCESS_ALTERNATE)

		Case "flowchart-predefined-process"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_PROCESS_PREDEFINED)

		Case "flowchart-punched-tape"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_PUNCHED_TAPE)

		Case "flowchart-sequential-access"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_SEQUENTIAL_ACCESS)

		Case "flowchart-sort"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_SORT)

		Case "flowchart-stored-data"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_STORED_DATA)

		Case "flowchart-summing-junction"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_SUMMING_JUNCTION)

		Case "flowchart-terminator"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_FLOWCHART_TERMINATOR)

		Case "star4"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_STARS_4_POINT)

		Case "star5"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_STARS_5_POINT)

		Case "star6" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_STARS_6_POINT)

		Case "concave-star6" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_STARS_6_POINT_CONCAVE)

		Case "star8"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_STARS_8_POINT)

		Case "star12" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_STARS_12_POINT)

		Case "star24"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_STARS_24_POINT)

		Case "mso-spt21", "doorplate" ; "doorplate"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_STARS_DOORPLATE)

		Case "bang"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_STARS_EXPLOSION)

		Case "horizontal-scroll"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_STARS_SCROLL_HORIZONTAL)

		Case "vertical-scroll"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_STARS_SCROLL_VERTICAL)

		Case "signet" ; "non-primitive"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_STARS_SIGNET)

		Case "col-502ad400"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_SYMBOL_BEVEL_DIAMOND)

		Case "col-60da8460"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_SYMBOL_BEVEL_OCTAGON)

		Case "quad-bevel"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_SYMBOL_BEVEL_SQUARE)

		Case "brace-pair"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_SYMBOL_BRACE_DOUBLE)

		Case "left-brace"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_SYMBOL_BRACE_LEFT)

		Case "right-brace"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_SYMBOL_BRACE_RIGHT)

		Case "bracket-pair"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_SYMBOL_BRACKET_DOUBLE)

		Case "left-bracket"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_SYMBOL_BRACKET_LEFT)

		Case "right-bracket"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_SYMBOL_BRACKET_RIGHT)

		Case "cloud"
			;~ Custom Shape Geometry Type = "non-primitive" ???? Try "cloud"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_SYMBOL_CLOUD)

		Case "flower"
			;~ Custom Shape Geometry Type = "non-primitive" ???? Try "flower"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_SYMBOL_FLOWER)

		Case "heart"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_SYMBOL_HEART)

		Case "lightning"
			;~ Custom Shape Geometry Type = "non-primitive" ???? Try "lightning"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_SYMBOL_LIGHTNING)

		Case "moon"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_SYMBOL_MOON)

		Case "smiley"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_SYMBOL_SMILEY)

		Case "sun"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_SYMBOL_SUN)

		Case "forbidden"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_SYMBOL_PROHIBITED)

		Case "puzzle"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOW_SHAPE_TYPE_SYMBOL_PUZZLE)

		Case Else

			Return SetError($__LO_STATUS_SUCCESS, 0, -1)
	EndSwitch
EndFunc   ;==>__LOWriter_Shape_GetCustomType

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ShapeArrowStyleName
; Description ...: Convert a Arrow head Constant to the corresponding name or reverse.
; Syntax ........: __LOWriter_ShapeArrowStyleName([$iArrowStyle = Null[, $sArrowStyle = Null]])
; Parameters ....: $iArrowStyle         - [optional] an integer value (0-32). Default is Null. The Arrow Style Constant to convert to its corresponding name. See $LOW_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeWriter_Constants.au3
;                  $sArrowStyle         - [optional] a string value. Default is Null. The Arrow Style Name to convert to the corresponding constant if found.
; Return values .: Success: String or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iArrowStyle not set to Null, not an Integer, less than 0, or greater than Arrow type constants. See $LOW_SHAPE_LINE_ARROW_TYPE_* as defined in LibreOfficeWriter_Constants.au3
;                  @Error 1 @Extended 2 Return 0 = $sArrowStyle not a String and not set to Null.
;                  @Error 1 @Extended 3 Return 0 = Both $iArrowStyle and $sArrowStyle set to Null.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Constant called in $iArrowStyle was successfully converted to its corresponding Arrow Type Name.
;                  @Error 0 @Extended 1 Return Integer = Success. Arrow Type Name called in $sArrowStyle was successfully converted to its corresponding Constant value.
;                  @Error 0 @Extended 2 Return String = Success. Arrow Type Name called in $sArrowStyle was not matched to an existing Constant value, returning called name. Possibly a custom value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ShapeArrowStyleName($iArrowStyle = Null, $sArrowStyle = Null)
	Local $asArrowStyles[33]

	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_NONE] = ""
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_ARROW_SHORT] = "Arrow short"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_CONCAVE_SHORT] = "Concave short"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_ARROW] = "Arrow"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_TRIANGLE] = "Triangle"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_CONCAVE] = "Concave"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_ARROW_LARGE] = "Arrow large"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_CIRCLE] = "Circle"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_SQUARE] = "Square"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_SQUARE_45] = "Square 45"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_DIAMOND] = "Diamond"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_HALF_CIRCLE] = "Half Circle"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_DIMENSIONAL_LINES] = "Dimension Lines"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_DIMENSIONAL_LINE_ARROW] = "Dimension Line Arrow"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_DIMENSION_LINE] = "Dimension Line"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_LINE_SHORT] = "Line short"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_LINE] = "Line"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_TRIANGLE_UNFILLED] = "Triangle unfilled"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_DIAMOND_UNFILLED] = "Diamond unfilled"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_CIRCLE_UNFILLED] = "Circle unfilled"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_SQUARE_45_UNFILLED] = "Square 45 unfilled"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_SQUARE_UNFILLED] = "Square unfilled"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_HALF_CIRCLE_UNFILLED] = "Half Circle unfilled"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_HALF_ARROW_LEFT] = "Half Arrow left"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_HALF_ARROW_RIGHT] = "Half Arrow right"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_REVERSED_ARROW] = "Reversed Arrow"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_DOUBLE_ARROW] = "Double Arrow"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_CF_ONE] = "CF One"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_CF_ONLY_ONE] = "CF Only One"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_CF_MANY] = "CF Many"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_CF_MANY_ONE] = "CF Many One"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_CF_ZERO_ONE] = "CF Zero One"
	$asArrowStyles[$LOW_SHAPE_LINE_ARROW_TYPE_CF_ZERO_MANY] = "CF Zero Many"

	If ($iArrowStyle <> Null) Then
		If Not __LO_IntIsBetween($iArrowStyle, 0, UBound($asArrowStyles) - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 0, $asArrowStyles[$iArrowStyle]) ; Return the requested Arrow Style name.

	ElseIf ($sArrowStyle <> Null) Then
		If Not IsString($sArrowStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		For $i = 0 To UBound($asArrowStyles) - 1
			If ($asArrowStyles[$i] = $sArrowStyle) Then Return SetError($__LO_STATUS_SUCCESS, 1, $i) ; Return the array element where the matching Arrow Style was found.

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
		Next

		Return SetError($__LO_STATUS_SUCCESS, 2, $sArrowStyle) ; If no matches, just return the name, as it could be a custom value.

	Else

		Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; No values called.
	EndIf
EndFunc   ;==>__LOWriter_ShapeArrowStyleName

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ShapeLineStyleName
; Description ...: Convert a Line Style Constant to the corresponding name or reverse.
; Syntax ........: __LOWriter_ShapeLineStyleName([$iLineStyle = Null[, $sLineStyle = Null]])
; Parameters ....: $iLineStyle          - [optional] an integer value. Default is Null. The Line Style Constant to convert to its corresponding name. See $LOW_SHAPE_LINE_STYLE_* as defined in LibreOfficeWriter_Constants.au3
;                  $sLineStyle          - [optional] a string value. Default is Null. The Line Style Name to convert to the corresponding constant if found.
; Return values .: Success: String or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iLineStyle not set to Null, not an Integer, less than 0, or greater than Line Style constants. See $LOW_SHAPE_LINE_STYLE_* as defined in LibreOfficeWriter_Constants.au3
;                  @Error 1 @Extended 2 Return 0 = $sLineStyle not a String and not set to Null.
;                  @Error 1 @Extended 3 Return 0 = Both $iLineStyle and $sLineStyle set to Null.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. Constant called in $iLineStyle was successfully converted to its corresponding Line Style Name.
;                  @Error 0 @Extended 1 Return Integer = Success. Line Style Name called in $sLineStyle was successfully converted to its corresponding Constant value.
;                  @Error 0 @Extended 2 Return String = Success. Line Style Name called in $sLineStyle was not matched to an existing Constant value, returning called name. Possibly a custom value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ShapeLineStyleName($iLineStyle = Null, $sLineStyle = Null)
	Local $asLineStyles[32]

	; $LOW_SHAPE_LINE_STYLE_NONE, $LOW_SHAPE_LINE_STYLE_CONTINUOUS, don't have a name, so to keep things symmetrical I created my own, but those two won't be used.
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_NONE] = "NONE"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_CONTINUOUS] = "CONTINUOUS"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_DOT] = "Dot"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_DOT_ROUNDED] = "Dot (Rounded)"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_LONG_DOT] = "Long Dot"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_LONG_DOT_ROUNDED] = "Long Dot (Rounded)"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_DASH] = "Dash"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_DASH_ROUNDED] = "Dash (Rounded)"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_LONG_DASH] = "Long Dash"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_LONG_DASH_ROUNDED] = "Long Dash (Rounded)"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_DOUBLE_DASH] = "Double Dash"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_DOUBLE_DASH_ROUNDED] = "Double Dash (Rounded)"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_DASH_DOT] = "Dash Dot"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_DASH_DOT_ROUNDED] = "Dash Dot (Rounded)"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_LONG_DASH_DOT] = "Long Dash Dot"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_LONG_DASH_DOT_ROUNDED] = "Long Dash Dot (Rounded)"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT] = "Double Dash Dot"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT_ROUNDED] = "Double Dash Dot (Rounded)"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_DASH_DOT_DOT] = "Dash Dot Dot"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_DASH_DOT_DOT_ROUNDED] = "Dash Dot Dot (Rounded)"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT_DOT] = "Double Dash Dot Dot"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT_DOT_ROUNDED] = "Double Dash Dot Dot (Rounded)"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_ULTRAFINE_DOTTED] = "Ultrafine Dotted (var)"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_FINE_DOTTED] = "Fine Dotted"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_ULTRAFINE_DASHED] = "Ultrafine Dashed"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_FINE_DASHED] = "Fine Dashed"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_DASHED] = "Dashed (var)"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_LINE_STYLE_9] = "Line Style 9"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_3_DASHES_3_DOTS] = "3 Dashes 3 Dots (var)"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_ULTRAFINE_2_DOTS_3_DASHES] = "Ultrafine 2 Dots 3 Dashes"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_2_DOTS_1_DASH] = "2 Dots 1 Dash"
	$asLineStyles[$LOW_SHAPE_LINE_STYLE_LINE_WITH_FINE_DOTS] = "Line with Fine Dots"

	If ($iLineStyle <> Null) Then
		If Not __LO_IntIsBetween($iLineStyle, 0, UBound($asLineStyles) - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 0, $asLineStyles[$iLineStyle]) ; Return the requested Line Style name.

	ElseIf ($sLineStyle <> Null) Then
		If Not IsString($sLineStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		For $i = 0 To UBound($asLineStyles) - 1
			If ($asLineStyles[$i] = $sLineStyle) Then Return SetError($__LO_STATUS_SUCCESS, 1, $i) ; Return the array element where the matching Line Style was found.

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
		Next

		Return SetError($__LO_STATUS_SUCCESS, 2, $sLineStyle) ; If no matches, just return the name, as it could be a custom value.

	Else

		Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; No values called.
	EndIf
EndFunc   ;==>__LOWriter_ShapeLineStyleName

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ShapePointGetSettings
; Description ...: Retrieve the current settings for a particular point in a shape.
; Syntax ........: __LOWriter_ShapePointGetSettings(ByRef $avArray, ByRef $aiFlags, ByRef $atPoints, $iArrayElement)
; Parameters ....: $avArray             - [in/out] an array of variants. An array to fill with settings. Array will be directly modified.
;                  $aiFlags             - [in/out] an array of integers. An Array of Point Type Flags returned from the Shape.
;                  $atPoints            - [in/out] an array of dll structs. An Array of Points returned from the Shape.
;                  $iArrayElement       - an integer value. The Array element that contains the point to retrieve the settings for.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $avArray is not an Array.
;                  @Error 1 @Extended 2 Return 0 = $aiFlags is not an Array.
;                  @Error 1 @Extended 3 Return 0 = $atPoints is not an Array.
;                  @Error 1 @Extended 4 Return 0 = $iArrayElement not an Integer.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve the current X coordinate.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve the current Y coordinate.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve the current Point Type Flag.
;                  @Error 3 @Extended 4 Return 0 = Failed to determine if the Point is a Curve or not.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Current Settings were successfully retrieved, $avArray has been filled with the current settings.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ShapePointGetSettings(ByRef $avArray, ByRef $aiFlags, ByRef $atPoints, $iArrayElement)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iX, $iY, $iPointType
	Local $bIsCurve

	If Not IsArray($avArray) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsArray($aiFlags) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsArray($atPoints) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iArrayElement) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$iX = $atPoints[$iArrayElement].X()
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$avArray[0] = $iX

	$iY = $atPoints[$iArrayElement].Y()
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$avArray[1] = $iY

	$iPointType = $aiFlags[$iArrayElement]
	If Not IsInt($iPointType) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$avArray[2] = $iPointType

	If ($iPointType = $LOW_SHAPE_POINT_TYPE_NORMAL) Then
		If ($iArrayElement <> (UBound($atPoints) - 1)) Then ; Requested point is not at the end of the array of points.

			If ($aiFlags[$iArrayElement + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then ; Point after requested point is a Control Point.
				; If a Point and the following Control point have the same coordinates, the point is not a curve.
				$bIsCurve = (($atPoints[$iArrayElement].X() = $atPoints[$iArrayElement + 1].X()) And ($atPoints[$iArrayElement].Y() = $atPoints[$iArrayElement + 1].Y())) ? (False) : (True)

			Else ; Next point after requested point is not a control type point.
				$bIsCurve = False
			EndIf

		Else ; Point is the last point, cant be a curve.
			$bIsCurve = False
		EndIf

	Else ; Point is a Smooth, or Symmetrical Point type, point is a curve regardless.
		$bIsCurve = True
	EndIf

	If Not IsBool($bIsCurve) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$avArray[3] = $bIsCurve

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_ShapePointGetSettings

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ShapePointModify
; Description ...: Internal function for modifying A Shape's Points.
; Syntax ........: __LOWriter_ShapePointModify(ByRef $aiFlags, ByRef $atPoints, ByRef $iArrayElement[, $iX = Null[, $iY = Null[, $iPointType = Null[, $bIsCurve = Null]]]])
; Parameters ....: $aiFlags             - [in/out] an array of integers. An Array of Point Type Flags returned from the Shape. Array will be directly modified.
;                  $atPoints            - [in/out] an array of dll structs. An Array of Points returned from the Shape. Array will be directly modified.
;                  $iArrayElement       - [in/out] an integer value. The Array element that contains the point to modify. This may be directly modified, depending on the settings.
;                  $iX                  - [optional] an integer value. Default is Null. The X coordinate value, set in Micrometers.
;                  $iY                  - [optional] an integer value. Default is Null. The Y coordinate value, set in Micrometers.
;                  $iPointType          - [optional] an integer value (0,1,3). Default is Null. The Type of Point to change the called point to. See Remarks. See constants $LOW_SHAPE_POINT_TYPE_* as defined in LibreOfficeWriter_Constants.au3
;                  $bIsCurve            - [optional] a boolean value. Default is Null. If True, the Normal Point is a Curve. See remarks.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $aiFlags not an Array.
;                  @Error 1 @Extended 2 Return 0 = $atPoints not an Array.
;                  @Error 1 @Extended 3 Return 0 = $iArrayElement not an Integer.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to Create a new Position Point Structure for the First Control Point.
;                  @Error 2 @Extended 2 Return 0 = Failed to Create a new Position Point Structure for the Second Control Point.
;                  @Error 2 @Extended 3 Return 0 = Failed to Create a new Position Point Structure for the Third Control Point.
;                  @Error 2 @Extended 4 Return 0 = Failed to Create a new Position Point Structure for the Fourth Control Point.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify the next position point in the shape.
;                  @Error 3 @Extended 2 Return 0 = Failed to identify the previous position point in the shape.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call any optional parameter with Null keyword to skip it.
;                  Only $LOW_SHAPE_TYPE_LINE_* type shapes have Points that can be added to, removed, or modified.
;                  This is a homemade function as LibreOffice doesn't offer an easy way for modifying points in a shape. Consequently this will not produce similar results as when working with Libre office manually, and may wreck your shape's shape. Use with caution.
;                  For an unknown reason, I am unable to insert "SMOOTH" Points, and consequently, any smooth Points are reverted back to "Normal" points, but still having their Smooth control points upon insertion that were already present in the shape. If you modify a point to "SMOOTH" type, it will be, for now, replaced with "Symmetrical".
;                  The first and last points in a shape can only be a "Normal" Point Type. The last point cannot be Curved, but the first can be.
;                  Calling and Smooth or Symmetrical point types with $bIsCurve = True, will be ignored, as they are already a curve.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ShapePointModify(ByRef $aiFlags, ByRef $atPoints, ByRef $iArrayElement, $iX = Null, $iY = Null, $iPointType = Null, $bIsCurve = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iNextArrayElement, $iPreviousArrayElement, $iSymmetricalPointXValue, $iSymmetricalPointYValue, $iOffset, $iForOffset, $iReDimCount
	Local $tControlPoint1, $tControlPoint2, $tControlPoint3, $tControlPoint4
	Local $avArray[0], $avArray2[0]

	If Not IsArray($aiFlags) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsArray($atPoints) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iArrayElement) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; Error if point called is not between 0 or number of points.

	If ($iArrayElement <> UBound($atPoints) - 1) Then ; If The requested point to be modified is not at the end of the Array of points, find the next regular point.

		For $i = ($iArrayElement + 1) To UBound($aiFlags) - 1 ; Locate the next non-Control Point in the Array for later use.
			If ($aiFlags[$i] <> $LOW_SHAPE_POINT_TYPE_CONTROL) Then
				$iNextArrayElement = $i
				ExitLoop
			EndIf

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
		Next

		If Not IsInt($iNextArrayElement) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Else
		$iNextArrayElement = -1
	EndIf

	If ($iArrayElement > 0) Then ; If Point requested is not the first point, find the previous Point's position.

		For $i = ($iArrayElement - 1) To 0 Step -1 ; Locate the previous non-Control Point in the Array for later use.
			If ($aiFlags[$i] <> $LOW_SHAPE_POINT_TYPE_CONTROL) Then
				$iPreviousArrayElement = $i
				ExitLoop
			EndIf

			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
		Next

		If Not IsInt($iPreviousArrayElement) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Else
		$iPreviousArrayElement = -1
	EndIf

	If ($iX <> Null) Then
		If ($iArrayElement < UBound($atPoints) - 1) And ($aiFlags[$iArrayElement + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then ; Next point is a control point, check if this point is a curve.

			If ($atPoints[$iArrayElement].X() = $atPoints[$iArrayElement + 1].X()) And ($atPoints[$iArrayElement].Y() = $atPoints[$iArrayElement + 1].Y()) Then ; Update the coordinates, because the point is not a curve.
				$atPoints[$iArrayElement + 1].X = $iX
			EndIf
		EndIf

		$atPoints[$iArrayElement].X = $iX
	EndIf

	If ($iY <> Null) Then
		If ($iArrayElement < UBound($atPoints) - 1) And ($aiFlags[$iArrayElement + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then ; Next point is a control point, check if this point is a curve.

			If ($atPoints[$iArrayElement].X() = $atPoints[$iArrayElement + 1].X()) And ($atPoints[$iArrayElement].Y() = $atPoints[$iArrayElement + 1].Y()) Then ; Update the coordinates, because the point is not a curve.
				$atPoints[$iArrayElement + 1].Y = $iY
			EndIf
		EndIf

		$atPoints[$iArrayElement].Y = $iY
	EndIf

	If ($iPointType <> Null) Then
		If ($iPointType <> $LOW_SHAPE_POINT_TYPE_NORMAL) Then ; New point type is a curve.

			If ($aiFlags[$iArrayElement] = $LOW_SHAPE_POINT_TYPE_NORMAL) Then ; Converting point from Normal to a curve.

				; Pick the lowest X and Y value difference between previous point and current point and Next point and current Point.
				$iSymmetricalPointXValue = ((($atPoints[$iArrayElement].X() - $atPoints[$iPreviousArrayElement].X()) * .5) < (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5)) ? Int((($atPoints[$iArrayElement].X() - $atPoints[$iPreviousArrayElement].X()) * .5)) : Int((($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5))
				$iSymmetricalPointYValue = ((($atPoints[$iArrayElement].Y() - $atPoints[$iPreviousArrayElement].Y()) * .5) < (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5)) ? Int((($atPoints[$iArrayElement].Y() - $atPoints[$iPreviousArrayElement].Y()) * .5)) : Int((($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5))

				If ($aiFlags[$iArrayElement - 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then ; previous point is a control Point, might just need to modify it.

					If (($iArrayElement - 2 > $iPreviousArrayElement) And $aiFlags[$iArrayElement - 2] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then ; there are two control points before this point, I can just modify the first point before.
						$tControlPoint1 = $atPoints[$iArrayElement - 2]

						$tControlPoint2 = $atPoints[$iArrayElement - 1]
						$tControlPoint2.X = ($atPoints[$iArrayElement].X() - $iSymmetricalPointXValue)
						$tControlPoint2.Y = ($atPoints[$iArrayElement].Y() - $iSymmetricalPointYValue)

					Else ; There is only one control point, I need to create a new one.
						$tControlPoint1 = $atPoints[$iArrayElement - 1]

						$tControlPoint2 = __LOWriter_CreatePoint($atPoints[$iArrayElement].X() - $iSymmetricalPointXValue, $atPoints[$iArrayElement].Y() - $iSymmetricalPointYValue)
						If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
					EndIf

				Else ; Previous point is a normal point, need to create new control points.
					$tControlPoint1 = __LOWriter_CreatePoint($atPoints[$iPreviousArrayElement].X(), $atPoints[$iPreviousArrayElement].Y())
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

					$tControlPoint2 = __LOWriter_CreatePoint($atPoints[$iArrayElement].X() - $iSymmetricalPointXValue, $atPoints[$iArrayElement].Y() - $iSymmetricalPointYValue)
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)
				EndIf

				If ($aiFlags[$iArrayElement + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then ; Next point is a control Point, might just need to modify it.

					If (($iArrayElement + 2 < $iNextArrayElement) And $aiFlags[$iArrayElement + 2] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then ; there are two control points after this point, I can just modify the first point after.
						$tControlPoint4 = $atPoints[$iArrayElement + 2]

						$tControlPoint3 = $atPoints[$iArrayElement + 1]
						$tControlPoint3.X = ($atPoints[$iArrayElement].X() + $iSymmetricalPointXValue)
						$tControlPoint3.Y = ($atPoints[$iArrayElement].Y() + $iSymmetricalPointYValue)

					Else ; There is only one control point, I need to create a new one and modify the other.
						$tControlPoint3 = __LOWriter_CreatePoint($atPoints[$iArrayElement].X() + $iSymmetricalPointXValue, $atPoints[$iArrayElement].Y() + $iSymmetricalPointYValue)
						If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

						$tControlPoint4 = $atPoints[$iArrayElement + 1] ; Modify the Control Point.
						$tControlPoint4.X = ($atPoints[$iNextArrayElement].X() - (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5))
						$tControlPoint4.Y = ($atPoints[$iNextArrayElement].Y() - (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5))
					EndIf

				Else ; Next point is a normal point, need to create new control points.
					$tControlPoint3 = __LOWriter_CreatePoint(($atPoints[$iArrayElement].X() + $iSymmetricalPointXValue), ($atPoints[$iArrayElement].Y() + $iSymmetricalPointYValue))
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

					$tControlPoint4 = __LOWriter_CreatePoint(Int($atPoints[$iNextArrayElement].X() - (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5)), Int($atPoints[$iNextArrayElement].Y() - (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5)))
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)
				EndIf

				$iOffset = 0
				$iForOffset = 0
				$iReDimCount = 4
				; Check if there already was 4 control point present around this point I am modifying.
				$iReDimCount -= ($aiFlags[$iArrayElement - 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) ? (1) : (0)
				$iReDimCount -= (($iArrayElement - 2 > $iPreviousArrayElement) And ($aiFlags[$iArrayElement - 2] = $LOW_SHAPE_POINT_TYPE_CONTROL)) ? (1) : (0)
				$iReDimCount -= ($aiFlags[$iArrayElement + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) ? (1) : (0)
				$iReDimCount -= (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOW_SHAPE_POINT_TYPE_CONTROL)) ? (1) : (0)

				ReDim $avArray[UBound($atPoints) + $iReDimCount]
				ReDim $avArray2[UBound($aiFlags) + $iReDimCount]
				$iReDimCount = 0

				For $i = 0 To UBound($atPoints) - 1
					If ($iOffset = 0) Then
						$avArray[$i + $iForOffset] = $atPoints[$i] ; Add the rest of the points to the array.
						$avArray2[$i + $iForOffset] = $aiFlags[$i] ; Add the rest of the points to the array.

					Else
						$iOffset -= 1 ; minus 1 from offset per round so I don't go over array limits
						$iForOffset -= 1 ; Minus 1 from ForOffset as I am skipping one For cycle.
					EndIf

					If ($i = $iPreviousArrayElement) Then ; Insert the new or modified control points.

						$avArray[$i + 1] = $tControlPoint1
						$avArray2[$i + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL
						$avArray[$i + 2] = $tControlPoint2
						$avArray2[$i + 2] = $LOW_SHAPE_POINT_TYPE_CONTROL
						$avArray[$i + 3] = $atPoints[$iArrayElement]
						$avArray2[$i + 3] = $iPointType
						$avArray[$i + 4] = $tControlPoint3
						$avArray2[$i + 4] = $LOW_SHAPE_POINT_TYPE_CONTROL
						$avArray[$i + 5] = $tControlPoint4
						$avArray2[$i + 5] = $LOW_SHAPE_POINT_TYPE_CONTROL

						$iOffset = 1 ; Add one to offset to skip the point I am modifying.
						$iOffset += ($aiFlags[$iArrayElement - 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) ? (1) : (0) ; If the point I am modifying has a control point before it, I need to skip them in the PointsArray.
						$iOffset += (($iArrayElement - 2 > $iPreviousArrayElement) And ($aiFlags[$iArrayElement - 2] = $LOW_SHAPE_POINT_TYPE_CONTROL)) ? (1) : (0) ; If the point I am modifying has two control points before it, I need to skip them in the PointsArray.
						$iOffset += ($aiFlags[$iArrayElement + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) ? (1) : (0) ; If the point I am modifying has a control point after it, I need to skip them in the PointsArray.
						$iOffset += (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOW_SHAPE_POINT_TYPE_CONTROL)) ? (1) : (0) ; If the point I am modifying has two control points after it, I need to skip them in the PointsArray.

						$iForOffset += 5 ; Add to $i to skip the elements I manually added.
					EndIf

					Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
				Next

				; Update the ArrayElement value to its new position.
				$iArrayElement += ($aiFlags[$iArrayElement - 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) ? (0) : (1) ; If the point I am modifying has a control point before it, don't add one to array element, because I didn't have to create and insert a new control point.
				$iArrayElement += (($iArrayElement - 2 > $iPreviousArrayElement) And ($aiFlags[$iArrayElement - 2] = $LOW_SHAPE_POINT_TYPE_CONTROL)) ? (0) : (1) ; If the point I am modifying has two control points before it, don't add one to array element, because I didn't have to create and insert a new control point.

				$atPoints = $avArray
				$aiFlags = $avArray2

			Else ; Point is already a curve.
				; Do nothing?
			EndIf

		Else ; New Point is a Normal Point.
			If ($aiFlags[$iArrayElement] <> $LOW_SHAPE_POINT_TYPE_NORMAL) Then ; Point being modified is not a normal type of point.

				If ($aiFlags[$iPreviousArrayElement] = $LOW_SHAPE_POINT_TYPE_NORMAL) Then ; If previous point is a normal point, see if I need to delete control points or not.

					If ($aiFlags[$iPreviousArrayElement + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then ; Point after previous point is a control point, see if previous point is a curved point.

						If ($atPoints[$iPreviousArrayElement].X() <> $atPoints[$iPreviousArrayElement + 1].X()) And ($atPoints[$iPreviousArrayElement].Y() <> $atPoints[$iPreviousArrayElement + 1].Y()) Then
							; Previous Point is a Curved normal point, copy the control points present.

							$tControlPoint1 = $atPoints[$iPreviousArrayElement + 1]

							If ($iPreviousArrayElement + 2 < $iArrayElement) And ($atPoints[$iPreviousArrayElement + 2] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then $tControlPoint2 = $atPoints[$iPreviousArrayElement + 2] ; If two control points are present, copy them.
						EndIf
					EndIf

				Else ; Previous point is not a normal point.
					; Copy Control Points present.

					If ($aiFlags[$iPreviousArrayElement + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then $tControlPoint1 = $atPoints[$iPreviousArrayElement + 1]

					If ($iPreviousArrayElement + 2 < $iArrayElement) And ($aiFlags[$iPreviousArrayElement + 2] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then $tControlPoint2 = $atPoints[$iPreviousArrayElement + 2] ; If two control points are present, copy them.
				EndIf

				If ($aiFlags[$iNextArrayElement] <> $LOW_SHAPE_POINT_TYPE_NORMAL) Then
					; Next point is a curve of some form, copy the control points.

					If ($aiFlags[$iNextArrayElement - 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then $tControlPoint4 = $atPoints[$iNextArrayElement - 1]

					If ($iNextArrayElement - 2 > $iArrayElement) And ($aiFlags[$iNextArrayElement - 2] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then $tControlPoint3 = $atPoints[$iNextArrayElement - 2] ; If two control points are present, copy them.
				EndIf

				$iOffset = 0
				$iForOffset = 0
				$iReDimCount = 4
				; Check how many control points I am keeping.
				$iReDimCount -= (IsObj($tControlPoint1)) ? (1) : (0)
				$iReDimCount -= (IsObj($tControlPoint2)) ? (1) : (0)
				$iReDimCount -= (IsObj($tControlPoint3)) ? (1) : (0)
				$iReDimCount -= (IsObj($tControlPoint4)) ? (1) : (0)

				ReDim $avArray[UBound($atPoints) - $iReDimCount]
				ReDim $avArray2[UBound($aiFlags) - $iReDimCount]
				$iReDimCount = 0

				For $i = 0 To UBound($atPoints) - 1
					If ($iOffset = 0) Then
						$avArray[$i + $iForOffset] = $atPoints[$i + $iOffset] ; Add the rest of the points to the array.
						$avArray2[$i + $iForOffset] = $aiFlags[$i + $iOffset] ; Add the rest of the points to the array.

					Else
						$iOffset -= 1 ; minus 1 from offset per round so I don't go over array limits
						$iForOffset -= 1 ; Minus 1 from ForOffset as I am skipping one For cycle.
					EndIf

					If ($i = $iPreviousArrayElement) Then ; Insert the old control points or remove them.

						If IsObj($tControlPoint1) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + $iForOffset] = $tControlPoint1
							$avArray2[$i + $iForOffset] = $LOW_SHAPE_POINT_TYPE_CONTROL

						Else
							If ($aiFlags[$iPreviousArrayElement + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
						EndIf

						If IsObj($tControlPoint2) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + $iForOffset] = $tControlPoint2
							$avArray2[$i + $iForOffset] = $LOW_SHAPE_POINT_TYPE_CONTROL

						Else
							If (($iPreviousArrayElement + 2 < $iArrayElement) And ($aiFlags[$iPreviousArrayElement + 2] = $LOW_SHAPE_POINT_TYPE_CONTROL)) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
						EndIf

						$iForOffset += 1
						$iOffset += 1
						$avArray[$i + $iForOffset] = $atPoints[$iArrayElement]
						$avArray2[$i + $iForOffset] = $iPointType

						If IsObj($tControlPoint3) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + $iForOffset] = $tControlPoint3
							$avArray2[$i + $iForOffset] = $LOW_SHAPE_POINT_TYPE_CONTROL

						Else
							If ($aiFlags[$iNextArrayElement - 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then $iOffset += 1
						EndIf

						If IsObj($tControlPoint3) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + $iForOffset] = $tControlPoint4
							$avArray2[$i + $iForOffset] = $LOW_SHAPE_POINT_TYPE_CONTROL

						Else
							If (($iNextArrayElement - 2 > $iArrayElement) And ($aiFlags[$iNextArrayElement - 2] = $LOW_SHAPE_POINT_TYPE_CONTROL)) Then $iOffset += 1
						EndIf
					EndIf

					Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
				Next

				; Update the ArrayElement value to its new position.
				$iArrayElement -= (IsObj($tControlPoint1)) ? (0) : (1) ; If ControlPoint 1 is a object, it means I copied it, meaning I didn't remove that point, so Array element will be in the same position. Else I need to remove from from ArrayElement.
				$iArrayElement -= (IsObj($tControlPoint2)) ? (0) : (1) ; If ControlPoint 2 is a object, it means I copied it, meaning I didn't remove that point, so Array element will be in the same position. Else I need to remove from from ArrayElement.

				$atPoints = $avArray
				$aiFlags = $avArray2

			Else ; Point being modified is a normal point already.
				; Do nothing?
			EndIf
		EndIf
	EndIf

	If ($bIsCurve <> Null) Then
		If ($aiFlags[$iArrayElement] = $LOW_SHAPE_POINT_TYPE_NORMAL) Then ; If Point to modify is a normal point, then proceed, else point is a curve already.

			If ($aiFlags[$iArrayElement + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then ; Point after point to modify is a control point, just modify it.
				$tControlPoint3 = $atPoints[$iArrayElement + 1]

				If ($bIsCurve = True) Then
					$tControlPoint3.X = ($atPoints[$iArrayElement].X() + (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5))
					$tControlPoint3.Y = ($atPoints[$iArrayElement].Y() + (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5))

					If (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOW_SHAPE_POINT_TYPE_CONTROL)) Then
						$tControlPoint4 = $atPoints[$iArrayElement + 2] ; Copy the second control point.

					Else ; Create a new control point.
						$tControlPoint4 = __LOWriter_CreatePoint(Int($atPoints[$iNextArrayElement].X() - (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5)), Int($atPoints[$iArrayElement].Y() - (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5)))
						If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)
					EndIf

				ElseIf ($bIsCurve = False) And ($aiFlags[$iNextArrayElement] <> $LOW_SHAPE_POINT_TYPE_NORMAL) Then ; Next point is a curve, so just modify the control point.
					$tControlPoint3.X = $atPoints[$iArrayElement].X() ; When the control point after a point has the same coordinates, it means it is not a curve.
					$tControlPoint3.Y = $atPoints[$iArrayElement].Y()
					; Copy the second control point if it exists.
					If (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOW_SHAPE_POINT_TYPE_CONTROL)) Then $tControlPoint4 = $atPoints[$iArrayElement + 2]

				Else ; IsCurve = False, and next point is normal. delete control points.
					$tControlPoint3 = Null
				EndIf

			Else ; Need to create new control points if IsCurve = True.
				If ($bIsCurve = True) Then
					$tControlPoint3 = __LOWriter_CreatePoint(Int($atPoints[$iArrayElement].X() + (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5)), Int($atPoints[$iArrayElement].Y() + (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5)))
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

					$tControlPoint4 = __LOWriter_CreatePoint(Int($atPoints[$iNextArrayElement].X() - (($atPoints[$iNextArrayElement].X() - $atPoints[$iArrayElement].X()) * .5)), Int($atPoints[$iNextArrayElement].Y() - (($atPoints[$iNextArrayElement].Y() - $atPoints[$iArrayElement].Y()) * .5)))
					If @error Then Return SetError($__LO_STATUS_INIT_ERROR, 4, 0)
				EndIf
			EndIf

			$iOffset = 0
			$iForOffset = 0
			$iReDimCount = 0
			; Check how many control points I am keeping vs creating.
			$iReDimCount += (IsObj($tControlPoint3)) ? (1) : (0)
			$iReDimCount += (IsObj($tControlPoint4)) ? (1) : (0)
			$iReDimCount -= ($aiFlags[$iArrayElement + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) ? (1) : (0) ; If a control point already existed, minus one from ReDim as it is either not new, or I am deleting it.
			$iReDimCount -= (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOW_SHAPE_POINT_TYPE_CONTROL)) ? (1) : (0)

			ReDim $avArray[UBound($atPoints) + $iReDimCount]
			ReDim $avArray2[UBound($aiFlags) + $iReDimCount]

			$iReDimCount = 0

			For $i = 0 To UBound($atPoints) - 1
				If ($iOffset = 0) Then
					$avArray[$i + $iForOffset] = $atPoints[$i] ; Add the rest of the points to the array.
					$avArray2[$i + $iForOffset] = $aiFlags[$i] ; Add the rest of the points to the array.

				Else
					$iOffset -= 1 ; minus 1 from offset per round so I don't go over array limits
					$iForOffset -= 1 ; Minus 1 from ForOffset as I am skipping one For cycle.
				EndIf

				If ($i = $iArrayElement) Then ; Insert the new or modified control points.

					If IsObj($tControlPoint3) Then
						$iForOffset += 1
						If ($aiFlags[$iArrayElement + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then $iOffset += 1 ; If there is a control point present, I need to skip it.

						$avArray[$i + $iForOffset] = $tControlPoint3
						$avArray2[$i + $iForOffset] = $LOW_SHAPE_POINT_TYPE_CONTROL

					Else
						If ($aiFlags[$iArrayElement + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
					EndIf

					If IsObj($tControlPoint4) Then
						$iForOffset += 1
						If (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOW_SHAPE_POINT_TYPE_CONTROL)) Then $iOffset += 1

						$avArray[$i + $iForOffset] = $tControlPoint4
						$avArray2[$i + $iForOffset] = $LOW_SHAPE_POINT_TYPE_CONTROL

					Else
						If (($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOW_SHAPE_POINT_TYPE_CONTROL)) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
					EndIf
				EndIf

				Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
			Next

			$atPoints = $avArray
			$aiFlags = $avArray2

		Else ; Point is a Curve, see if bIsCurve = False.
			If ($bIsCurve = False) Then ; If bIsCurve = True, I can just skip it, as there is nothing to do when the point is a curve already.

				If ($aiFlags[$iNextArrayElement] <> $LOW_SHAPE_POINT_TYPE_NORMAL) Then ; Next point is a curve, need to keep the control points.
					If ($aiFlags[$iArrayElement + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then $tControlPoint3 = $atPoints[$iArrayElement + 1]
					If ($iArrayElement + 2 < $iNextArrayElement) And ($aiFlags[$iArrayElement + 2] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then $tControlPoint4 = $atPoints[$iArrayElement + 2]
				EndIf

				If ($iPreviousArrayElement <> -1) And ($aiFlags[$iPreviousArrayElement] <> $LOW_SHAPE_POINT_TYPE_NORMAL) Then ; There is a previous point, and it is a curve, I need to keep the control points.
					If ($aiFlags[$iPreviousArrayElement + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then $tControlPoint1 = $atPoints[$iPreviousArrayElement + 1]
					If ($iPreviousArrayElement + 2 < $iArrayElement) And ($aiFlags[$iPreviousArrayElement + 2] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then $tControlPoint2 = $atPoints[$iPreviousArrayElement + 2]

				ElseIf ($iPreviousArrayElement <> -1) And ($aiFlags[$iPreviousArrayElement] = $LOW_SHAPE_POINT_TYPE_NORMAL) Then ; There is a previous point, and it is a normal point.
					; See if it is curved.

					If ($aiFlags[$iPreviousArrayElement + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) And _
							(($atPoints[$iPreviousArrayElement].X() <> $atPoints[$iPreviousArrayElement + 1].X()) And _
							($atPoints[$iPreviousArrayElement].Y() <> $atPoints[$iPreviousArrayElement + 1].Y())) Then ; Previous Point is a curve, need to keep the control points.
						$tControlPoint1 = $atPoints[$iPreviousArrayElement + 1]

						If ($aiFlags[$iPreviousArrayElement + 2] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then $tControlPoint2 = $atPoints[$iPreviousArrayElement + 2]
					EndIf
				EndIf

				$iOffset = 0
				$iForOffset = 0
				$iReDimCount = 4
				; Check how many control points I am keeping vs deleting.
				$iReDimCount -= (IsObj($tControlPoint1)) ? (1) : (0)
				$iReDimCount -= (IsObj($tControlPoint2)) ? (1) : (0)
				$iReDimCount -= (IsObj($tControlPoint3)) ? (1) : (0)
				$iReDimCount -= (IsObj($tControlPoint4)) ? (1) : (0)

				ReDim $avArray[UBound($atPoints) - $iReDimCount]
				ReDim $avArray2[UBound($aiFlags) - $iReDimCount]
				$iReDimCount = 0

				For $i = 0 To UBound($atPoints) - 1
					If ($iOffset = 0) Then
						$avArray[$i + $iForOffset] = $atPoints[$i] ; Add the rest of the points to the array.
						$avArray2[$i + $iForOffset] = $aiFlags[$i] ; Add the rest of the points to the array.

					Else
						$iOffset -= 1 ; minus 1 from offset per round so I don't go over array limits
						$iForOffset -= 1 ; Minus 1 from ForOffset as I am skipping one For cycle.
					EndIf

					If ($i = $iPreviousArrayElement) Then
						If IsObj($tControlPoint1) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + $iForOffset] = $tControlPoint1
							$avArray2[$i + $iForOffset] = $LOW_SHAPE_POINT_TYPE_CONTROL

						Else
							If ($aiFlags[$iPreviousArrayElement + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
						EndIf

						If IsObj($tControlPoint2) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + $iForOffset] = $tControlPoint2
							$avArray2[$i + $iForOffset] = $LOW_SHAPE_POINT_TYPE_CONTROL

						Else
							If (($iPreviousArrayElement + 2 < $iArrayElement) And ($aiFlags[$iPreviousArrayElement + 2] = $LOW_SHAPE_POINT_TYPE_CONTROL)) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
						EndIf

					ElseIf ($i = $iArrayElement) Then ; Insert or skip Control Points as necessary.
						$avArray[$i] = $atPoints[$iArrayElement]
						$avArray2[$i] = $LOW_SHAPE_POINT_TYPE_NORMAL

						If IsObj($tControlPoint3) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + 1] = $tControlPoint3
							$avArray2[$i + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL

						Else
							If ($aiFlags[$iPreviousArrayElement + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
						EndIf

						If IsObj($tControlPoint4) Then
							$iForOffset += 1
							$iOffset += 1

							$avArray[$i + 2] = $tControlPoint4
							$avArray2[$i + 2] = $LOW_SHAPE_POINT_TYPE_CONTROL

						Else
							If (($iPreviousArrayElement + 2 < $iArrayElement) And ($aiFlags[$iPreviousArrayElement + 2] = $LOW_SHAPE_POINT_TYPE_CONTROL)) Then $iOffset += 1 ; If there is a control point present, I need to skip it.
						EndIf
					EndIf

					Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
				Next

				; Update the ArrayElement value to its new position.
				If ($iPreviousArrayElement <> -1) Then $iArrayElement -= ((IsObj($tControlPoint2) And ($iPreviousArrayElement + 2 < $iArrayElement) And ($aiFlags[$iPreviousArrayElement + 2] = $LOW_SHAPE_POINT_TYPE_CONTROL))) ? (0) : (1) ; If ControlPoint 2 is a object, it means I copied it, meaining I didn't remove that point, so Array element will be in the same position. Else I need to remove from from ArrayElement.
				If ($iPreviousArrayElement <> -1) Then $iArrayElement -= ((IsObj($tControlPoint1) And ($aiFlags[$iPreviousArrayElement + 1] = $LOW_SHAPE_POINT_TYPE_CONTROL))) ? (0) : (1) ; If ControlPoint 1 is a object, it means I copied it, meaning I didn't remove that point, so Array element will be in the same position. Else I need to remove from from ArrayElement.

				$atPoints = $avArray
				$aiFlags = $avArray2
			EndIf
		EndIf
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_ShapePointModify

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_TableBorder
; Description ...: Set or Retrieve Table Border settings -- internal function. Libre Office 3.6 and Up.
; Syntax ........: __LOWriter_TableBorder(ByRef $oTable, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight, $iVert, $iHori)
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableInsert, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $bWid                - a boolean value. If True the calling function is for setting Border Line Width.
;                  $bSty                - a boolean value. If True the calling function is for setting Border Line Style.
;                  $bCol                - a boolean value. If True the calling function is for setting Border Line Color.
;                  $iTop                - an integer value. See TableBorder Style, Width, and Color functions for possible values.
;                  $iBottom             - an integer value. See TableBorder Style, Width, and Color functions for possible values.
;                  $iLeft               - an integer value. See TableBorder Style, Width, and Color functions for possible values.
;                  $iRight              - an integer value. See TableBorder Style, Width, and Color functions for possible values.
;                  $iVert               - an integer value. See TableBorder Style, Width, and Color functions for possible values.
;                  $iHori               - an integer value. See TableBorder Style, Width, and Color functions for possible values.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Error retrieving Object "TableBorder2".
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style/Color when Top Border width not set.
;                  @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Style/Color when Bottom Border width not set.
;                  @Error 4 @Extended 3 Return 0 = Cannot set Left Border Style/Color when Left Border width not set.
;                  @Error 4 @Extended 4 Return 0 = Cannot set Right Border Style/Color when Right Border width not set.
;                  @Error 4 @Extended 5 Return 0 = Cannot set Vertical Border Style/Color when Vertical Border width not set.
;                  @Error 4 @Extended 6 Return 0 = Cannot set Horizontal Border Style/Color when Horizontal Border width not set.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Table Object, and either $bWid, $bSty, or $bCol set to true, with all other parameters set to Null keyword, to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_TableBorder(ByRef $oTable, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight, $iVert, $iHori)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avBorder[6]
	Local $tBL2, $tTB2

	If Not __LO_VersionCheck(3.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If (($bWid + $bSty + $bCol) <> 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight, $iVert, $iHori) Then
		If $bWid Then
			__LO_ArrayFill($avBorder, $oTable.TableBorder2.TopLine.LineWidth(), $oTable.TableBorder2.BottomLine.LineWidth(), _
					$oTable.TableBorder2.LeftLine.LineWidth(), $oTable.TableBorder2.RightLine.LineWidth(), $oTable.TableBorder2.VerticalLine.LineWidth(), _
					$oTable.TableBorder2.HorizontalLine.LineWidth())

		ElseIf $bSty Then
			__LO_ArrayFill($avBorder, $oTable.TableBorder2.TopLine.LineStyle(), $oTable.TableBorder2.BottomLine.LineStyle(), _
					$oTable.TableBorder2.LeftLine.LineStyle(), $oTable.TableBorder2.RightLine.LineStyle(), $oTable.TableBorder2.VerticalLine.LineStyle(), _
					$oTable.TableBorder2.HorizontalLine.LineStyle())

		ElseIf $bCol Then
			__LO_ArrayFill($avBorder, $oTable.TableBorder2.TopLine.Color(), $oTable.TableBorder2.BottomLine.Color(), _
					$oTable.TableBorder2.LeftLine.Color(), $oTable.TableBorder2.RightLine.Color(), $oTable.TableBorder2.VerticalLine.Color(), _
					$oTable.TableBorder2.HorizontalLine.Color())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avBorder)
	EndIf

	$tBL2 = __LO_CreateStruct("com.sun.star.table.BorderLine2")
	If Not IsObj($tBL2) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tTB2 = $oTable.TableBorder2
	If Not IsObj($tTB2) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	If $iTop <> Null Then
		If Not $bWid And ($tTB2.TopLine.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0) ; If Width not set, cant set color or style.

		; Top Line
		$tBL2.LineWidth = ($bWid) ? ($iTop) : ($tTB2.TopLine.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iTop) : ($tTB2.TopLine.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iTop) : ($tTB2.TopLine.Color()) ; copy Color over to new size structure
		$tTB2.TopLine = $tBL2
	EndIf

	If $iBottom <> Null Then
		If Not $bWid And ($tTB2.BottomLine.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 2, 0) ; If Width not set, cant set color or style.

		; Bottom Line
		$tBL2.LineWidth = ($bWid) ? ($iBottom) : ($tTB2.BottomLine.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iBottom) : ($tTB2.BottomLine.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iBottom) : ($tTB2.BottomLine.Color()) ; copy Color over to new size structure
		$tTB2.BottomLine = $tBL2
	EndIf

	If $iLeft <> Null Then
		If Not $bWid And ($tTB2.LeftLine.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 3, 0) ; If Width not set, cant set color or style.

		; Left Line
		$tBL2.LineWidth = ($bWid) ? ($iLeft) : ($tTB2.LeftLine.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iLeft) : ($tTB2.LeftLine.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iLeft) : ($tTB2.LeftLine.Color()) ; copy Color over to new size structure
		$tTB2.LeftLine = $tBL2
	EndIf

	If $iRight <> Null Then
		If Not $bWid And ($tTB2.RightLine.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 4, 0) ; If Width not set, cant set color or style.

		; Right Line
		$tBL2.LineWidth = ($bWid) ? ($iRight) : ($tTB2.RightLine.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iRight) : ($tTB2.RightLine.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iRight) : ($tTB2.RightLine.Color()) ; copy Color over to new size structure
		$tTB2.RightLine = $tBL2
	EndIf

	If $iVert <> Null Then
		If Not $bWid And ($tTB2.VerticalLine.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 5, 0) ; If Width not set, cant set color or style.

		; Vertical Line
		$tBL2.LineWidth = ($bWid) ? ($iVert) : ($tTB2.VerticalLine.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iVert) : ($tTB2.VerticalLine.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iVert) : ($tTB2.VerticalLine.Color()) ; copy Color over to new size structure
		$tTB2.VerticalLine = $tBL2
	EndIf

	If $iHori <> Null Then
		If Not $bWid And ($tTB2.HorizontalLine.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 6, 0) ; If Width not set, cant set color or style.

		; Horizontal Line
		$tBL2.LineWidth = ($bWid) ? ($iHori) : ($tTB2.HorizontalLine.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iHori) : ($tTB2.HorizontalLine.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iHori) : ($tTB2.HorizontalLine.Color()) ; copy Color over to new size structure
		$tTB2.HorizontalLine = $tBL2
	EndIf

	$oTable.TableBorder2 = $tTB2

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_TableBorder

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_TableCursorMove
; Description ...: Text-TableCursor related movements.
; Syntax ........: __LOWriter_TableCursorMove(ByRef $oCursor, $iMove, $iCount[, $bSelect = False])
; Parameters ....: $oCursor             - [in/out] an object. A TableCursor Object returned from _LOWriter_TableCreateCursor function.
;                  $iMove               - an Integer value. The movement command constant. See remarks and Constants, $LOW_TABLECUR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iCount              - an integer value. Number of movements to make.
;                  $bSelect             - [optional] a boolean value. Default is False. If True, select data during this cursor movement.
; Return values .: Success: Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iMove not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iMove mismatch with Cursor type. See Constants, $LOW_TABLECUR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iCount not an integer or is a negative.
;                  @Error 1 @Extended 5 Return 0 = $bSelect not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 2 Return 0 = Error processing cursor move.
;                  --Success--
;                  @Error 0 @Extended ? Return Boolean = Success, Cursor object movement was processed successfully. Returns True if the full count of movements were successful, else false if none or only partially successful. @Extended set to number of successful movements. Or Page Number for "gotoPage" command. See Remarks
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iMove may be set to any of the following constants. Only some movements accept movement amounts (such as "goRight" 2) etc. Also only some accept creating/ extending a selection of text/ data. They will be specified below.
;                  To Clear /Unselect a current selection, you can input a move such as "goRight", 0, False.
;                  #Cursor Movement Constants which accept number of Moves and Selecting:
;                   $LOW_TABLECUR_GO_LEFT, Move the cursor left/right n cells.
;                   $LOW_TABLECUR_GO_RIGHT, Move the cursor left/right n cells.
;                   $LOW_TABLECUR_GO_UP, Move the cursor up/down n cells.
;                   $LOW_TABLECUR_GO_DOWN, Move the cursor up/down n cells.
;                  #Cursor Movements which accept Selecting Only:
;                   $LOW_TABLECUR_GOTO_START, Move the cursor to the top left cell.
;                   $LOW_TABLECUR_GOTO_END, Move the cursor to the bottom right cell.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_TableCursorMove(ByRef $oCursor, $iMove, $iCount, $bSelect = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCounted = 0
	Local $bMoved = False
	Local $asMoves[6]

	$asMoves[$LOW_TABLECUR_GO_LEFT] = "goLeft"
	$asMoves[$LOW_TABLECUR_GO_RIGHT] = "goRight"
	$asMoves[$LOW_TABLECUR_GO_UP] = "goUp"
	$asMoves[$LOW_TABLECUR_GO_DOWN] = "goDown"
	$asMoves[$LOW_TABLECUR_GOTO_START] = "gotoStart"
	$asMoves[$LOW_TABLECUR_GOTO_END] = "gotoEnd"

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iMove) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iMove >= UBound($asMoves)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bSelect) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	Switch $iMove
		Case $LOW_TABLECUR_GO_LEFT, $LOW_TABLECUR_GO_RIGHT, $LOW_TABLECUR_GO_UP, $LOW_TABLECUR_GO_DOWN
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "(" & $iCount & "," & $bSelect & ")")
			$iCounted = ($bMoved) ? ($iCount) : (0)

			Return SetError($__LO_STATUS_SUCCESS, $iCounted, $bMoved)

		Case $LOW_TABLECUR_GOTO_START, $LOW_TABLECUR_GOTO_END
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "(" & $bSelect & ")")
			$iCounted = ($bMoved) ? (1) : (0)

			Return SetError($__LO_STATUS_SUCCESS, $iCounted, $bMoved)

		Case Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	EndSwitch
EndFunc   ;==>__LOWriter_TableCursorMove

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_TableHasCellName
; Description ...: Check whether the Table contains a Cell by the requested name.
; Syntax ........: __LOWriter_TableHasCellName(ByRef $oTable, ByRef $sCellName)
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableInsert, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $sCellName           - [in/out] a string value. The requested cell name.
; Return values .: Success: Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sCellName not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Cell Names.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = If the table contains the requested Cell Name, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_TableHasCellName(ByRef $oTable, ByRef $sCellName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aCellNames

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sCellName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$aCellNames = $oTable.getCellNames()
	If Not IsArray($aCellNames) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	For $i = 0 To UBound($aCellNames) - 1
		If StringInStr($aCellNames[$i], $sCellName) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, False) ; Cell not found
EndFunc   ;==>__LOWriter_TableHasCellName

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_TableHasColumnRange
; Description ...: Check if Table contains the requested Column.
; Syntax ........: __LOWriter_TableHasColumnRange(ByRef $oTable, ByRef $iColumn)
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableInsert, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $iColumn             - [in/out] an integer value. The requested Column.
; Return values .: Success: Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iColumn not an Integer.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = If True, the table contains the requested Column. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_TableHasColumnRange(ByRef $oTable, ByRef $iColumn)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iColumn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, ($iColumn <= ($oTable.getColumns.getCount() - 1)) ? (True) : (False))
EndFunc   ;==>__LOWriter_TableHasColumnRange

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_TableHasRowRange
; Description ...: Check if a Table contains the requested row.
; Syntax ........: __LOWriter_TableHasRowRange(ByRef $oTable, ByRef $iRow)
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableInsert, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $iRow                - [in/out] an integer value. The requested row.
; Return values .: Success: Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iRow not an Integer.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = If True, the table contains the requested row. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_TableHasRowRange(ByRef $oTable, ByRef $iRow)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iRow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, ($iRow <= ($oTable.getRows.getCount() - 1)) ? (True) : (False))
EndFunc   ;==>__LOWriter_TableHasRowRange

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_TableRowSplitToggle
; Description ...: Set or Retrieve Table Row split setting for an entire Table.
; Syntax ........: __LOWriter_TableRowSplitToggle(ByRef $oTable[, $bSplitRows = Null])
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned by a previous _LOWriter_TableInsert, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName function.
;                  $bSplitRows          - [optional] a boolean value. Default is Null. If True, the content in a Table row is allowed to split at page splits, else if False, Content is not allowed to split across pages.
; Return values .: Success: Integer or Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;                  @Error 1 @Extended 2 Return 0 = $bSplitRows not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Table's Row count.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve first Row's current split row setting.
;                  --Success--
;                  @Error 0 @Extended 0 Return 0 = Success. All optional parameters were set to Null, Table Rows have multiple SplitRow settings, returning 0 to indicate this.
;                  @Error 0 @Extended 1 Return Boolean = Success. All optional parameters were set to Null, returning current split row setting as a Boolean.
;                  @Error 0 @Extended 2 Return 1 = Success. Setting was successfully set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_TableRowSplitToggle(ByRef $oTable, $bSplitRows = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iRows
	Local $bSplitRowTest

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iRows = $oTable.getRows.getCount()
	If Not IsInt($iRows) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If ($bSplitRows = Null) Then ; Retrieve Split Rows Setting

		; Retrieve the First Row's Split Row setting.
		$bSplitRowTest = $oTable.getRows.getByIndex(0).IsSplitAllowed()
		If Not IsBool($bSplitRowTest) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		For $i = 1 To $iRows - 1
			If $bSplitRowTest <> ($oTable.getRows.getByIndex($i).IsSplitAllowed()) Then Return SetError($__LO_STATUS_SUCCESS, 0, 0) ; Table Rows have mixed settings, return 0.
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
		Next

		Return SetError($__LO_STATUS_SUCCESS, 1, $bSplitRowTest) ; All Table Rows are set the same as the first Row, return that setting.

	Else ; Set Split Rows Setting
		If Not IsBool($bSplitRows) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		For $i = 0 To $iRows - 1
			$oTable.getRows.getByIndex($i).IsSplitAllowed = $bSplitRows
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
		Next

		Return SetError($__LO_STATUS_SUCCESS, 2, 1)
	EndIf
EndFunc   ;==>__LOWriter_TableRowSplitToggle

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_TextCursorMove
; Description ...: For TextCursor related movements.
; Syntax ........: __LOWriter_TextCursorMove(ByRef $oCursor, $iMove, $iCount[, $bSelect = False])
; Parameters ....: $oCursor             - [in/out] an object. A TextCursor Object returned from any TextCursor creation functions.
;                  $iMove               - an Integer value. The movement command constant. See remarks and Constants, $LOW_TEXTCUR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iCount              - an integer value. Number of movements to make.
;                  $bSelect             - [optional] a boolean value. Default is False. If True, select data during this cursor movement.
; Return values .: Success: Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iMove not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iMove mismatch with Cursor type. See Constants, $LOW_TEXTCUR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iCount not an integer or is a negative.
;                  @Error 1 @Extended 5 Return 0 = $bSelect not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 2 Return 0 = Error processing cursor move.
;                  --Success--
;                  @Error 0 @Extended ? Return Boolean = Success, Cursor object movement was processed successfully. Returns True if the full count of movements were successful, else false if none or only partially successful. @Extended set to number of successful movements. Or Page Number for "gotoPage" command. See Remarks
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iMove may be set to any of the following constants.
;                  Only some movements accept movement amounts (such as "goRight" 2) etc.
;                  Only some accept creating/ extending a selection of text/ data. They will be specified below.
;                  To Clear /Unselect a current selection, you can input a move such as "goRight", 0, False.
;                  #Cursor Movement Constants which accept number of Moves and Selecting:
;                   $LOW_TEXTCUR_GO_LEFT, Move the cursor left by n characters.
;                   $LOW_TEXTCUR_GO_RIGHT, Move the cursor right by n characters.
;                   $LOW_TEXTCUR_GOTO_NEXT_WORD, Move to the start of the next word.
;                   $LOW_TEXTCUR_GOTO_PREV_WORD, Move to the end of the previous word.
;                   $LOW_TEXTCUR_GOTO_NEXT_SENTENCE, Move to the start of the next sentence.
;                   $LOW_TEXTCUR_GOTO_PREV_SENTENCE, Move to the end of the previous sentence.
;                   $LOW_TEXTCUR_GOTO_NEXT_PARAGRAPH, Move to the start of the next paragraph.
;                   $LOW_TEXTCUR_GOTO_PREV_PARAGRAPH, Move to the End of the previous paragraph.
;                  #Cursor Movements which accept Selecting Only:
;                   $LOW_TEXTCUR_GOTO_START, Move the cursor to the start of the text.
;                   $LOW_TEXTCUR_GOTO_END, Move the cursor to the end of the text.
;                   $LOW_TEXTCUR_GOTO_END_OF_WORD, Move to the end of the current word.
;                   $LOW_TEXTCUR_GOTO_START_OF_WORD, Move to the start of the current word.
;                   $LOW_TEXTCUR_GOTO_END_OF_SENTENCE, Move to the end of the current sentence.
;                   $LOW_TEXTCUR_GOTO_START_OF_SENTENCE, Move to the start of the current sentence.
;                   $LOW_TEXTCUR_GOTO_END_OF_PARAGRAPH, Move to the end of the current paragraph.
;                   $LOW_TEXTCUR_GOTO_START_OF_PARAGRAPH, Move to the start of the current paragraph.
;                  #Cursor Movements which accept nothing and are done once per call:
;                   $LOW_TEXTCUR_COLLAPSE_TO_START,
;                   $LOW_TEXTCUR_COLLAPSE_TO_END (Collapses the current selection and moves the cursor to start or End of selection.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_TextCursorMove(ByRef $oCursor, $iMove, $iCount, $bSelect = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCounted = 0
	Local $bMoved = False
	Local $asMoves[18]

	$asMoves[$LOW_TEXTCUR_COLLAPSE_TO_START] = "collapseToStart"
	$asMoves[$LOW_TEXTCUR_COLLAPSE_TO_END] = "collapseToEnd"
	$asMoves[$LOW_TEXTCUR_GO_LEFT] = "goLeft"
	$asMoves[$LOW_TEXTCUR_GO_RIGHT] = "goRight"
	$asMoves[$LOW_TEXTCUR_GOTO_START] = "gotoStart"
	$asMoves[$LOW_TEXTCUR_GOTO_END] = "gotoEnd"
	$asMoves[$LOW_TEXTCUR_GOTO_NEXT_WORD] = "gotoNextWord"
	$asMoves[$LOW_TEXTCUR_GOTO_PREV_WORD] = "gotoPreviousWord"
	$asMoves[$LOW_TEXTCUR_GOTO_END_OF_WORD] = "gotoEndOfWord"
	$asMoves[$LOW_TEXTCUR_GOTO_START_OF_WORD] = "gotoStartOfWord"
	$asMoves[$LOW_TEXTCUR_GOTO_NEXT_SENTENCE] = "gotoNextSentence"
	$asMoves[$LOW_TEXTCUR_GOTO_PREV_SENTENCE] = "gotoPreviousSentence"
	$asMoves[$LOW_TEXTCUR_GOTO_END_OF_SENTENCE] = "gotoEndOfSentence"
	$asMoves[$LOW_TEXTCUR_GOTO_START_OF_SENTENCE] = "gotoStartOfSentence"
	$asMoves[$LOW_TEXTCUR_GOTO_NEXT_PARAGRAPH] = "gotoNextParagraph"
	$asMoves[$LOW_TEXTCUR_GOTO_PREV_PARAGRAPH] = "gotoPreviousParagraph"
	$asMoves[$LOW_TEXTCUR_GOTO_END_OF_PARAGRAPH] = "gotoEndOfParagraph"
	$asMoves[$LOW_TEXTCUR_GOTO_START_OF_PARAGRAPH] = "gotoStartOfParagraph"

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iMove) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iMove >= UBound($asMoves)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bSelect) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	Switch $iMove
		Case $LOW_TEXTCUR_GO_LEFT, $LOW_TEXTCUR_GO_RIGHT
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "(" & $iCount & "," & $bSelect & ")")
			$iCounted = ($bMoved) ? ($iCount) : (0)

			Return SetError($__LO_STATUS_SUCCESS, $iCounted, $bMoved)

		Case $LOW_TEXTCUR_GOTO_NEXT_WORD, $LOW_TEXTCUR_GOTO_PREV_WORD, $LOW_TEXTCUR_GOTO_NEXT_SENTENCE, $LOW_TEXTCUR_GOTO_PREV_SENTENCE, _
				$LOW_TEXTCUR_GOTO_NEXT_PARAGRAPH, $LOW_TEXTCUR_GOTO_PREV_PARAGRAPH

			Do
				$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "(" & $bSelect & ")")
				$iCounted += ($bMoved) ? (1) : (0)
				Sleep((IsInt($iCounted / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
			Until ($iCounted >= $iCount) Or ($bMoved = False)

			Return SetError($__LO_STATUS_SUCCESS, $iCounted, $bMoved)

		Case $LOW_TEXTCUR_GOTO_START, $LOW_TEXTCUR_GOTO_END, $LOW_TEXTCUR_GOTO_END_OF_WORD, $LOW_TEXTCUR_GOTO_START_OF_WORD, _
				$LOW_TEXTCUR_GOTO_END_OF_SENTENCE, $LOW_TEXTCUR_GOTO_START_OF_SENTENCE, $LOW_TEXTCUR_GOTO_END_OF_PARAGRAPH, _
				$LOW_TEXTCUR_GOTO_START_OF_PARAGRAPH
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "(" & $bSelect & ")")
			$iCounted = ($bMoved) ? (1) : (0)

			Return SetError($__LO_STATUS_SUCCESS, $iCounted, $bMoved)

		Case $LOW_TEXTCUR_COLLAPSE_TO_START, $LOW_TEXTCUR_COLLAPSE_TO_END
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "()")
			$iCounted = ($bMoved) ? (1) : (0)

			Return SetError($__LO_STATUS_SUCCESS, $iCounted, $bMoved)

		Case Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	EndSwitch
EndFunc   ;==>__LOWriter_TextCursorMove

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_TransparencyGradientConvert
; Description ...: Convert a Transparency Gradient percentage value to a color value or from a color value to a percentage.
; Syntax ........: __LOWriter_TransparencyGradientConvert([$iPercentToLong = Null[, $iLongToPercent = Null]])
; Parameters ....: $iPercentToLong      - [optional] an integer value. Default is Null. The percentage to convert to Long color integer value.
;                  $iLongToPercent      - [optional] an integer value. Default is Null. The Long color integer value to convert to percentage.
; Return values .: Success: Integer.
;                  Failure: Null and sets the @Error and @Extended flags to non-zero.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return Null = No values called in parameters.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. The requested Integer value converted from percentage to Long color format.
;                  @Error 0 @Extended 1 Return Integer = Success. The requested Integer value from Long color format to percentage.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_TransparencyGradientConvert($iPercentToLong = Null, $iLongToPercent = Null)
	Local $iReturn

	If ($iPercentToLong <> Null) Then
		$iReturn = ((255 * ($iPercentToLong / 100)) + .50) ; Change percentage to decimal and times by White color (255 RGB) Add . 50 to round up if applicable.
		$iReturn = _LO_ConvertColorToLong(Int($iReturn), Int($iReturn), Int($iReturn))

		Return SetError($__LO_STATUS_SUCCESS, 0, $iReturn)

	ElseIf ($iLongToPercent <> Null) Then
		$iReturn = _LO_ConvertColorFromLong(Null, $iLongToPercent)
		$iReturn = Int((($iReturn[0] / 255) * 100) + .50) ; All return color values will be the same, so use only one. Add . 50 to round up if applicable.

		Return SetError($__LO_STATUS_SUCCESS, 1, $iReturn)

	Else

		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, Null)
	EndIf
EndFunc   ;==>__LOWriter_TransparencyGradientConvert

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_TransparencyGradientNameInsert
; Description ...: Create and insert a new Transparency Gradient name.
; Syntax ........: __LOWriter_TransparencyGradientNameInsert(ByRef $oDoc, $tTGradient)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $tTGradient          - a dll struct value. A Gradient Structure to copy settings from.
; Return values .: Success: String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $tTGradient not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.drawing.TransparencyGradientTable" Object.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.awt.Gradient" or "com.sun.star.awt.Gradient2" structure.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error creating Transparency Gradient Name.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. A new transparency Gradient name was created. Returning the new name as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If The Transparency Gradient name is blank, I need to create a new name and apply it. I think I could re-use an old one without problems, but I'm not sure, so to be safe, I will create a new one.
;                  If there are no names that have been already created, then I need to create and apply one before the transparency gradient will be displayed.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_TransparencyGradientNameInsert(ByRef $oDoc, $tTGradient)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tNewTGradient
	Local $oTGradTable
	Local $iCount = 1
	Local $sGradient = "com.sun.star.awt.Gradient2"

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($tTGradient) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If Not __LO_VersionCheck(7.6) Then $sGradient = "com.sun.star.awt.Gradient"

	$oTGradTable = $oDoc.createInstance("com.sun.star.drawing.TransparencyGradientTable")
	If Not IsObj($oTGradTable) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	While $oTGradTable.hasByName("Transparency " & $iCount)
		$iCount += 1
		Sleep((IsInt($iCount / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
	WEnd

	$tNewTGradient = __LO_CreateStruct($sGradient)
	If Not IsObj($tNewTGradient) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	; Copy the settings over from the input Style Gradient to my new one. This may not be necessary? But just in case.
	With $tNewTGradient
		.Style = $tTGradient.Style()
		.XOffset = $tTGradient.XOffset()
		.YOffset = $tTGradient.YOffset()
		.Angle = $tTGradient.Angle()
		.Border = $tTGradient.Border()
		.StartColor = $tTGradient.StartColor()
		.EndColor = $tTGradient.EndColor()

		If __LO_VersionCheck(7.6) Then .ColorStops = $tTGradient.ColorStops()
	EndWith

	$oTGradTable.insertByName("Transparency " & $iCount, $tNewTGradient)
	If Not ($oTGradTable.hasByName("Transparency " & $iCount)) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, "Transparency " & $iCount)
EndFunc   ;==>__LOWriter_TransparencyGradientNameInsert

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ViewCursorMove
; Description ...: For ViewCursor related movements.
; Syntax ........: __LOWriter_ViewCursorMove(ByRef $oCursor, $iMove, $iCount[, $bSelect = False])
; Parameters ....: $oCursor             - [in/out] an object. A ViewCursor Object returned from _LOWriter_DocGetViewCursor function.
;                  $iMove               - an integer value. The movement command. See remarks and Constants, $LOW_VIEWCUR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iCount              - an integer value. Number of movements to make.
;                  $bSelect             - [optional] a boolean value. Default is False. Whether to select data during this cursor movement.
; Return values .: Success: Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iMove not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iMove mismatch with Cursor type. See Constants, $LOW_VIEWCUR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iCount not an integer or is a negative.
;                  @Error 1 @Extended 5 Return 0 = $bSelect not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 2 Return 0 = Error processing cursor move.
;                  --Success--
;                  @Error 0 @Extended ? Return Boolean = Success, Cursor object movement was processed successfully. Returns True if the full count of movements were successful, else false if none or only partially successful. @Extended set to number of successful movements. Or Page Number for "gotoPage" command. See Remarks
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iMove may be set to any of the following constants.
;                  Only some movements accept movement amounts (such as "goRight" 2) etc.
;                  Only some accept creating/ extending a selection of text/ data. They will be specified below.
;                  To Clear /Unselect a current selection, you can input a move such as "goRight", 0, False.
;                  #Cursor Movement Constants which accept number of Moves and Selecting:
;                   $LOW_VIEWCUR_GO_DOWN, Move the cursor Down by n lines.
;                   $LOW_VIEWCUR_GO_UP, Move the cursor Up by n lines.
;                   $LOW_VIEWCUR_GO_LEFT, Move the cursor left by n characters.
;                   $LOW_VIEWCUR_GO_RIGHT, Move the cursor right by n characters.
;                  #Cursor Movements which accept number of Moves Only:
;                   $LOW_VIEWCUR_JUMP_TO_NEXT_PAGE, Move the cursor to the Next page.
;                   $LOW_VIEWCUR_JUMP_TO_PREV_PAGE, Move the cursor to the previous page.
;                   $LOW_VIEWCUR_SCREEN_DOWN, Scroll the view forward by one visible page.
;                   $LOW_VIEWCUR_SCREEN_UP, Scroll the view back by one visible page.
;                  #Cursor Movements which accept Selecting Only:
;                   $LOW_VIEWCUR_GOTO_END_OF_LINE, Move the cursor to the end of the current line.
;                   $LOW_VIEWCUR_GOTO_START_OF_LINE, Move the cursor to the start of the current line.
;                   $LOW_VIEWCUR_GOTO_START, Move the cursor to the start of the document or Table.
;                   $LOW_VIEWCUR_GOTO_END, Move the cursor to the end of the document or Table.
;                  #Cursor Movements which accept nothing and are done once per call:
;                   $LOW_VIEWCUR_JUMP_TO_FIRST_PAGE, Move the cursor to the first page.
;                   $LOW_VIEWCUR_JUMP_TO_LAST_PAGE, Move the cursor to the Last page.
;                   $LOW_VIEWCUR_JUMP_TO_END_OF_PAGE, Move the cursor to the end of the current page.
;                   $LOW_VIEWCUR_JUMP_TO_START_OF_PAGE, Move the cursor to the start of the current page.
;                  #Misc. Cursor Movements:
;                   $LOW_VIEWCUR_JUMP_TO_PAGE (accepts page number to jump to in $iCount, Returns what page was successfully jumped to.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ViewCursorMove(ByRef $oCursor, $iMove, $iCount, $bSelect = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCounted = 0
	Local $bMoved = False
	Local $asMoves[17]

	$asMoves[$LOW_VIEWCUR_GO_DOWN] = "goDown"
	$asMoves[$LOW_VIEWCUR_GO_UP] = "goUp"
	$asMoves[$LOW_VIEWCUR_GO_LEFT] = "goLeft"
	$asMoves[$LOW_VIEWCUR_GO_RIGHT] = "goRight"
	$asMoves[$LOW_VIEWCUR_GOTO_END_OF_LINE] = "gotoEndOfLine"
	$asMoves[$LOW_VIEWCUR_GOTO_START_OF_LINE] = "gotoStartOfLine"
	$asMoves[$LOW_VIEWCUR_JUMP_TO_FIRST_PAGE] = "jumpToFirstPage"
	$asMoves[$LOW_VIEWCUR_JUMP_TO_LAST_PAGE] = "jumpToLastPage"
	$asMoves[$LOW_VIEWCUR_JUMP_TO_PAGE] = "jumpToPage"
	$asMoves[$LOW_VIEWCUR_JUMP_TO_NEXT_PAGE] = "jumpToNextPage"
	$asMoves[$LOW_VIEWCUR_JUMP_TO_PREV_PAGE] = "jumpToPreviousPage"
	$asMoves[$LOW_VIEWCUR_JUMP_TO_END_OF_PAGE] = "jumpToEndOfPage"
	$asMoves[$LOW_VIEWCUR_JUMP_TO_START_OF_PAGE] = "jumpToStartOfPage"
	$asMoves[$LOW_VIEWCUR_SCREEN_DOWN] = "screenDown"
	$asMoves[$LOW_VIEWCUR_SCREEN_UP] = "screenUp"
	$asMoves[$LOW_VIEWCUR_GOTO_START] = "gotoStart"
	$asMoves[$LOW_VIEWCUR_GOTO_END] = "gotoEnd"

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iMove) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iMove >= UBound($asMoves)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iCount) Or ($iCount < 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bSelect) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	Switch $iMove
		Case $LOW_VIEWCUR_GO_DOWN, $LOW_VIEWCUR_GO_UP, $LOW_VIEWCUR_GO_LEFT, $LOW_VIEWCUR_GO_RIGHT
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "(" & $iCount & "," & $bSelect & ")")
			$iCounted = ($bMoved) ? ($iCount) : (0)

			Return SetError($__LO_STATUS_SUCCESS, $iCounted, $bMoved)

		Case $LOW_VIEWCUR_GOTO_END_OF_LINE, $LOW_VIEWCUR_GOTO_START_OF_LINE, $LOW_VIEWCUR_GOTO_START, $LOW_VIEWCUR_GOTO_END
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "(" & $bSelect & ")")
			$iCounted = ($bMoved) ? (1) : (0)

			Return SetError($__LO_STATUS_SUCCESS, $iCounted, $bMoved)

		Case $LOW_VIEWCUR_JUMP_TO_PAGE
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "(" & $iCount & ")")

			Return SetError($__LO_STATUS_SUCCESS, $oCursor.getPage(), $bMoved)

		Case $LOW_VIEWCUR_JUMP_TO_NEXT_PAGE, $LOW_VIEWCUR_JUMP_TO_PREV_PAGE, $LOW_VIEWCUR_SCREEN_DOWN, $LOW_VIEWCUR_SCREEN_UP
			Do
				$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "()")
				$iCounted += ($bMoved) ? (1) : (0)
				Sleep((IsInt($iCounted / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
			Until ($iCounted >= $iCount) Or ($bMoved = False)

			Return SetError($__LO_STATUS_SUCCESS, $iCounted, $bMoved)

		Case $LOW_VIEWCUR_JUMP_TO_FIRST_PAGE, $LOW_VIEWCUR_JUMP_TO_LAST_PAGE, $LOW_VIEWCUR_JUMP_TO_END_OF_PAGE, _
				$LOW_VIEWCUR_JUMP_TO_START_OF_PAGE
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "()")
			$iCounted = ($bMoved) ? (1) : (0)

			Return SetError($__LO_STATUS_SUCCESS, $iCounted, $bMoved)

		Case Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	EndSwitch
EndFunc   ;==>__LOWriter_ViewCursorMove
