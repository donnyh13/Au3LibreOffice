#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once
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
; __LOWriter_AddTo1DArray
; __LOWriter_AddTo2DArray
; __LOWriter_AnyAreDefault
; __LOWriter_ArrayFill
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
; __LOWriter_CreateStruct
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
; __LOWriter_GetPrinterSetting
; __LOWriter_GradientNameInsert
; __LOWriter_GradientPresets
; __LOWriter_HeaderBorder
; __LOWriter_ImageGetSuggestedSize
; __LOWriter_Internal_CursorGetDataType
; __LOWriter_Internal_CursorGetType
; __LOWriter_InternalComErrorHandler
; __LOWriter_IntIsBetween
; __LOWriter_IsCellRange
; __LOWriter_IsTableInDoc
; __LOWriter_NumIsBetween
; __LOWriter_NumStyleCreateScript
; __LOWriter_NumStyleDeleteScript
; __LOWriter_NumStyleInitiateDocument
; __LOWriter_NumStyleListFormat
; __LOWriter_NumStyleModify
; __LOWriter_NumStyleRetrieve
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
; __LOWriter_ParTabStopList
; __LOWriter_ParTabStopMod
; __LOWriter_ParTxtFlowOpt
; __LOWriter_RegExpConvert
; __LOWriter_SetPropertyValue
; __LOWriter_TableBorder
; __LOWriter_TableCursorMove
; __LOWriter_TableHasCellName
; __LOWriter_TableHasColumnRange
; __LOWriter_TableHasRowRange
; __LOWriter_TableRowSplitToggle
; __LOWriter_TextCursorMove
; __LOWriter_TransparencyGradientConvert
; __LOWriter_TransparencyGradientNameInsert
; __LOWriter_UnitConvert
; __LOWriter_VarsAreDefault
; __LOWriter_VarsAreNull
; __LOWriter_VersionCheck
; __LOWriter_ViewCursorMove
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_AddTo1DArray
; Description ...: Add data to a 1 Dimensional array.
; Syntax ........: __LOWriter_AddTo1DArray(Byref $aArray, $vData[, $bCountInFirst = False])
; Parameters ....: $aArray              - [in/out] an array of unknowns. The Array to directly add data to.  Array will be directly modified.
;                  $vData               - a variant value. The Data to add to the Array.
;                  $bCountInFirst       - [optional] a boolean value. Default is False. If True the first element of the array is a count of contained elements.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $aArray not an Array
;				   @Error 1 @Extended 2 Return 0 = $bCountinFirst not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $aArray contains too many columns.
;				   @Error 1 @Extended 4 Return 0 = $aArray[0] contains non integer data or is not empty, and $bCountInFirst is set to True.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Array item was successfully added.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_AddTo1DArray(ByRef $aArray, $vData, $bCountInFirst = False)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($aArray) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bCountInFirst) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If UBound($aArray, $UBOUND_COLUMNS) > 1 Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0) ; Too many columns

	If $bCountInFirst And (UBound($aArray) = 0) Then
		ReDim $aArray[1]
		$aArray[0] = 0
	EndIf

	If $bCountInFirst And (($aArray[0] <> "") And Not IsInt($aArray[0])) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	ReDim $aArray[UBound($aArray) + 1]
	$aArray[UBound($aArray) - 1] = $vData
	If $bCountInFirst Then $aArray[0] += 1
	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_AddTo1DArray

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_AddTo1DArray
; Description ...: Add data to a 2 Dimensional array.
; Syntax ........: __LOWriter_AddTo2DArray(Byref $aArray, $vDataCol1, $vDataCol2[, $bCountInFirst = False])
; Parameters ....: $aArray              - [in/out] an array of unknowns. The Array to directly add data to. Array will be directly modified.
;                  $vDataCol1           - a variant value. The Data to add to the first column of the Array.
;                  $vDataCol2           - a variant value. The Data to add to the Second column of the Array.
;                  $bCountInFirst       - [optional] a boolean value. Default is False. If True the first element of the array is a count of contained elements.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $aArray not an Array
;				   @Error 1 @Extended 2 Return 0 = $bCountinFirst not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $aArray does not contain two columns.
;				   @Error 1 @Extended 4 Return 0 = $aArray[0][0] contains non integer data or is not empty, and $bCountInFirst is set to True.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Array item was successfully added.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_AddTo2DArray(ByRef $aArray, $vDataCol1, $vDataCol2, $bCountInFirst = False)
	Local Const $UBOUND_COLUMNS = 2

	If Not IsArray($aArray) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bCountInFirst) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If UBound($aArray, $UBOUND_COLUMNS) <> 2 Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0) ; Too many or too few  columns

	If $bCountInFirst And (UBound($aArray) = 0) Then
		ReDim $aArray[1][2]
		$aArray[0][0] = 0
	EndIf

	If $bCountInFirst And (($aArray[0][0] <> "") And Not IsInt($aArray[0][0])) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	ReDim $aArray[UBound($aArray) + 1][2]
	$aArray[UBound($aArray) - 1][0] = $vDataCol1
	$aArray[UBound($aArray) - 1][1] = $vDataCol2
	If $bCountInFirst Then $aArray[0][0] += 1
	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_AddTo2DArray

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
;				   Failure: False
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = If Any parameters are equal to Default, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_AnyAreDefault($vVar1, $vVar2 = Null, $vVar3 = Null, $vVar4 = Null, $vVar5 = Null, $vVar6 = Null, $vVar7 = Null, $vVar8 = Null)
	Local $bAnyDefault1, $bAnyDefault2
	$bAnyDefault1 = (($vVar1 = Default) Or ($vVar2 = Default) Or ($vVar3 = Default) Or ($vVar4 = Default)) ? True : False
	$bAnyDefault2 = (($vVar5 = Default) Or ($vVar6 = Default) Or ($vVar7 = Default) Or ($vVar8 = Default)) ? True : False
	Return ($bAnyDefault1 Or $bAnyDefault2) ? SetError($__LOW_STATUS_SUCCESS, 0, True) : SetError($__LOW_STATUS_SUCCESS, 0, False)
EndFunc   ;==>__LOWriter_AnyAreDefault

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ArrayFill
; Description ...: Fill an Array with data.
; Syntax ........: __LOWriter_ArrayFill(Byref $aArrayToFill[, $vVar1 = Null[, $vVar2 = Null[, $vVar3 = Null[, $vVar4 = Null[, $vVar5 = Null[, $vVar6 = Null[, $vVar7 = Null[, $vVar8 = Null[, $vVar9 = Null[, $vVar10 = Null[, $vVar11 = Null[, $vVar12 = Null[, $vVar13 = Null[, $vVar14 = Null[, $vVar15 = Null[, $vVar16 = Null[, $vVar17 = Null[, $vVar18 = Null]]]]]]]]]]]]]]]]]])
; Parameters ....: $aArrayToFill        - [in/out] an array of unknowns. The Array to Fill.  Array will be directly modified.
;                  $vVar1               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar2               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar3               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar4               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar5               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar6               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar7               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar8               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar9               - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar10              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar11              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar12              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar13              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar14              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar15              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar16              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar17              - [optional] a variant value. Default is Null. The Data to add to the Array.
;                  $vVar18              - [optional] a variant value. Default is Null. The Data to add to the Array.
; Return values .: None
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call only how many you parameters you need to add to the Array. Automatically resizes the Array if it is the incorrect size.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ArrayFill(ByRef $aArrayToFill, $vVar1 = Null, $vVar2 = Null, $vVar3 = Null, $vVar4 = Null, $vVar5 = Null, $vVar6 = Null, $vVar7 = Null, $vVar8 = Null, $vVar9 = Null, $vVar10 = Null, $vVar11 = Null, $vVar12 = Null, $vVar13 = Null, _
		$vVar14 = Null, $vVar15 = Null, $vVar16 = Null, $vVar17 = Null, $vVar18 = Null)
	#forceref $vVar1, $vVar2, $vVar3, $vVar4, $vVar5, $vVar6, $vVar7, $vVar8, $vVar9, $vVar10, $vVar11, $vVar12, $vVar13, $vVar14
	#forceref $vVar15, $vVar16, $vVar17, $vVar18

	If UBound($aArrayToFill) < (@NumParams - 1) Then ReDim $aArrayToFill[@NumParams - 1]
	For $i = 0 To @NumParams - 2
		$aArrayToFill[$i] = Eval("vVar" & $i + 1)
	Next
EndFunc   ;==>__LOWriter_ArrayFill

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Border
; Description ...: Border Setting Internal function. Libre Office Version 3.4 and Up.
; Syntax ........: __LOWriter_Border(Byref $oObj, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
; Parameters ....: $oObj                - [in/out] an object. An Object that implements BorderLine2 service for border properties.
;                  $bWid                - a boolean value. If True the calling function is for setting Border Line Width.
;                  $bSty                - a boolean value. If True the calling function is for setting Border Line Style.
;                  $bCol                - a boolean value. If True the calling function is for setting Border Line Color.
;                  $iTop                - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iBottom             - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iLeft               - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iRight              - an integer value. See Border Style, Width, and Color functions for possible values.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oObj Variable not Object type variable.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style/Color when Border width not set.
;				   @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border style/Color when Border width not set.
;				   @Error 4 @Extended 3 Return 0 = Cannot set Left Border style/Color when Border width not set.
;				   @Error 4 @Extended 4 Return 0 = Cannot set Right Border style/Color when Border width not set.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with all other parameters set to Null keyword, and $bWid, or $bSty, or $bCol set to true to get the corresponding current settings.
;					All distance values are set in MicroMeters. Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_Border(ByRef $oObj, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avBorder[4]
	Local $tBL2

	If Not __LOWriter_VersionCheck(3.4) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If (($bWid + $bSty + $bCol) <> 1) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0) ; If more than one Boolean is true = error

	If __LOWriter_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then
		If $bWid Then
			__LOWriter_ArrayFill($avBorder, $oObj.TopBorder.LineWidth(), $oObj.BottomBorder.LineWidth(), $oObj.LeftBorder.LineWidth(), _
					$oObj.RightBorder.LineWidth())
		ElseIf $bSty Then
			__LOWriter_ArrayFill($avBorder, $oObj.TopBorder.LineStyle(), $oObj.BottomBorder.LineStyle(), $oObj.LeftBorder.LineStyle(), _
					$oObj.RightBorder.LineStyle())
		ElseIf $bCol Then
			__LOWriter_ArrayFill($avBorder, $oObj.TopBorder.Color(), $oObj.BottomBorder.Color(), $oObj.LeftBorder.Color(), $oObj.RightBorder.Color())
		EndIf
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avBorder)
	EndIf

	$tBL2 = __LOWriter_CreateStruct("com.sun.star.table.BorderLine2")
	If Not IsObj($tBL2) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If $iTop <> Null Then
		If Not $bWid And ($oObj.TopBorder.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 1, 0) ; If Width not set, cant set color or style.
		; Top Line
		$tBL2.LineWidth = ($bWid) ? $iTop : $oObj.TopBorder.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iTop : $oObj.TopBorder.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iTop : $oObj.TopBorder.Color() ; copy Color over to new size structure
		$oObj.TopBorder = $tBL2
	EndIf

	If $iBottom <> Null Then
		If Not $bWid And ($oObj.BottomBorder.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 2, 0) ; If Width not set, cant set color or style.
		; Bottom Line
		$tBL2.LineWidth = ($bWid) ? $iBottom : $oObj.BottomBorder.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iBottom : $oObj.BottomBorder.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iBottom : $oObj.BottomBorder.Color() ; copy Color over to new size structure
		$oObj.BottomBorder = $tBL2
	EndIf

	If $iLeft <> Null Then
		If Not $bWid And ($oObj.LeftBorder.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 3, 0) ; If Width not set, cant set color or style.
		; Left Line
		$tBL2.LineWidth = ($bWid) ? $iLeft : $oObj.LeftBorder.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iLeft : $oObj.LeftBorder.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iLeft : $oObj.LeftBorder.Color() ; copy Color over to new size structure
		$oObj.LeftBorder = $tBL2
	EndIf

	If $iRight <> Null Then
		If Not $bWid And ($oObj.RightBorder.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 4, 0) ; If Width not set, cant set color or style.
		; Right Line
		$tBL2.LineWidth = ($bWid) ? $iRight : $oObj.RightBorder.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iRight : $oObj.RightBorder.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iRight : $oObj.RightBorder.Color() ; copy Color over to new size structure
		$oObj.RightBorder = $tBL2
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_Border

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharBorder
; Description ...: Character Border Setting and retrieving Internal function.
; Syntax ........: __LOWriter_CharBorder(Byref $oObj, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $bWid                - a boolean value. If True the calling function is for setting Border Line Width.
;                  $bSty                - a boolean value. If True the calling function is for setting Border Line Style.
;                  $bCol                - a boolean value. If True the calling function is for setting Border Line Color.
;                  $iTop                - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iBottom             - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iLeft               - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iRight              - an integer value. See Border Style, Width, and Color functions for possible values.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oObj Variable not Object type variable.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style/Color when Border width not set.
;				   @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border style/Color when Border width not set.
;				   @Error 4 @Extended 3 Return 0 = Cannot set Left Border style/Color when Border width not set.
;				   @Error 4 @Extended 4 Return 0 = Cannot set Right Border style/Color when Border width not set.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, and $bWid,
;					or $bSty, or $bCol set to true to get the corresponding current settings.
;					All distance values are set in MicroMeters. Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharBorder(ByRef $oObj, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avBorder[4]
	Local $tBL2

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If (($bWid + $bSty + $bCol) <> 1) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0) ; If more than one Boolean is true = error

	If __LOWriter_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then

		If $bWid Then
			__LOWriter_ArrayFill($avBorder, $oObj.CharTopBorder.LineWidth(), $oObj.CharBottomBorder.LineWidth(), $oObj.CharLeftBorder.LineWidth(), _
					$oObj.CharRightBorder.LineWidth())
		ElseIf $bSty Then
			__LOWriter_ArrayFill($avBorder, $oObj.CharTopBorder.LineStyle(), $oObj.CharBottomBorder.LineStyle(), $oObj.CharLeftBorder.LineStyle(), _
					$oObj.CharRightBorder.LineStyle())
		ElseIf $bCol Then
			__LOWriter_ArrayFill($avBorder, $oObj.CharTopBorder.Color(), $oObj.CharBottomBorder.Color(), $oObj.CharLeftBorder.Color(), _
					$oObj.CharRightBorder.Color())
		EndIf
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avBorder)
	EndIf

	$tBL2 = __LOWriter_CreateStruct("com.sun.star.table.BorderLine2")
	If Not IsObj($tBL2) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If $iTop <> Null Then
		If Not $bWid And ($oObj.CharTopBorder.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 1, 0) ; If Width not set, cant set color or style.
		; Top Line
		$tBL2.LineWidth = ($bWid) ? $iTop : $oObj.CharTopBorder.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iTop : $oObj.CharTopBorder.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iTop : $oObj.CharTopBorder.Color() ; copy Color over to new size structure
		$oObj.CharTopBorder = $tBL2
	EndIf

	If $iBottom <> Null Then
		If Not $bWid And ($oObj.CharBottomBorder.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 2, 0) ; If Width not set, cant set color or style.
		; Bottom Line
		$tBL2.LineWidth = ($bWid) ? $iBottom : $oObj.CharBottomBorder.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iBottom : $oObj.CharBottomBorder.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iBottom : $oObj.CharBottomBorder.Color() ; copy Color over to new size structure
		$oObj.CharBottomBorder = $tBL2
	EndIf

	If $iLeft <> Null Then
		If Not $bWid And ($oObj.CharLeftBorder.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 3, 0) ; If Width not set, cant set color or style.
		; Left Line
		$tBL2.LineWidth = ($bWid) ? $iLeft : $oObj.CharLeftBorder.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iLeft : $oObj.CharLeftBorder.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iLeft : $oObj.CharLeftBorder.Color() ; copy Color over to new size structure
		$oObj.CharLeftBorder = $tBL2
	EndIf

	If $iRight <> Null Then
		If Not $bWid And ($oObj.CharRightBorder.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 4, 0) ; If Width not set, cant set color or style.
		; Right Line
		$tBL2.LineWidth = ($bWid) ? $iRight : $oObj.CharRightBorder.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iRight : $oObj.CharRightBorder.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iRight : $oObj.CharRightBorder.Color() ; copy Color over to new size structure
		$oObj.CharRightBorder = $tBL2
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_CharBorder

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharBorderPadding
; Description ...: Set and retrieve the distance between the border and the characters.
; Syntax ........: __LOWriter_CharBorderPadding(Byref $oObj, $iAll, $iTop, $iBottom, $iLeft, $iRight)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $iAll                - an integer value. Set all four values to the same value. When used, all other parameters are ignored. In MicroMeters.
;                  $iTop                - an integer value. Set the Top border distance in MicroMeters.
;                  $iBottom             - an integer value. Set the Bottom border distance in MicroMeters.
;                  $iLeft               - an integer value. Set the left border distance in MicroMeters.
;                  $iRight              - an integer value. Set the Right border distance in MicroMeters.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iAll not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iTop not an Integer.
;				   @Error 1 @Extended 6 Return 0 = $iBottom not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $Left not an Integer.
;				   @Error 1 @Extended 8 Return 0 = $iRight not an Integer.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iAll border distance
;				   |								2 = Error setting $iTop border distance
;				   |								4 = Error setting $iBottom border distance
;				   |								8 = Error setting $iLeft border distance
;				   |								16 = Error setting $iRight border distance
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;					All distance values are set in MicroMeters. Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharBorderPadding(ByRef $oObj, $iAll, $iTop, $iBottom, $iLeft, $iRight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiBPadding[5]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If __LOWriter_VarsAreNull($iAll, $iTop, $iBottom, $iLeft, $iRight) Then
		__LOWriter_ArrayFill($aiBPadding, $oObj.CharBorderDistance(), $oObj.CharTopBorderDistance(), $oObj.CharBottomBorderDistance(), _
				$oObj.CharLeftBorderDistance(), $oObj.CharRightBorderDistance())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $aiBPadding)
	EndIf

	If ($iAll <> Null) Then
		If Not __LOWriter_IntIsBetween($iAll, 0, $iAll) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oObj.CharBorderDistance = $iAll
		$iError = (__LOWriter_IntIsBetween($oObj.CharBorderDistance(), $iAll - 1, $iAll + 1)) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iTop <> Null) Then
		If Not __LOWriter_IntIsBetween($iTop, 0, $iTop) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oObj.CharTopBorderDistance = $iTop
		$iError = (__LOWriter_IntIsBetween($oObj.CharTopBorderDistance(), $iTop - 1, $iTop + 1)) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iBottom <> Null) Then
		If Not __LOWriter_IntIsBetween($iBottom, 0, $iBottom) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oObj.CharBottomBorderDistance = $iBottom
		$iError = (__LOWriter_IntIsBetween($oObj.CharBottomBorderDistance(), $iBottom - 1, $iBottom + 1)) ? $iError : BitOR($iError, 4)
	EndIf

	If ($iLeft <> Null) Then
		If Not __LOWriter_IntIsBetween($iLeft, 0, $iLeft) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oObj.CharLeftBorderDistance = $iLeft
		$iError = (__LOWriter_IntIsBetween($oObj.CharLeftBorderDistance(), $iLeft - 1, $iLeft + 1)) ? $iError : BitOR($iError, 8)
	EndIf

	If ($iRight <> Null) Then
		If Not __LOWriter_IntIsBetween($iRight, 0, $iRight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$oObj.CharRightBorderDistance = $iRight
		$iError = (__LOWriter_IntIsBetween($oObj.CharRightBorderDistance(), $iRight - 1, $iRight + 1)) ? $iError : BitOR($iError, 16)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_CharBorderPadding

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharEffect
; Description ...: Set or Retrieve the Font Effect settings.
; Syntax ........: __LOWriter_CharEffect(Byref $oObj, $iRelief, $iCase, $bHidden, $bOutline, $bShadow)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $iRelief             - an integer value (0-2). The Character Relief style. See Constants, $LOW_RELIEF_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iCase               - an integer value (0-4). The Character Case Style. See Constants, $LOW_CASEMAP_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bHidden             - a boolean value. Whether the Characters are hidden or not.
;                  $bOutline            - a boolean value. Whether the characters have an outline around the outside.
;                  $bShadow             - a boolean value. Whether the characters have a shadow.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iRelief not an integer or less than 0 or greater than 2. See Constants, $LOW_RELIEF_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 5 Return 0 = $iCase not an integer or less than 0 or greater than 4. See Constants, $LOW_CASEMAP_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 6 Return 0 = $bHidden not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $bOutline not a Boolean.
;				   @Error 1 @Extended 8 Return 0 = $bShadow not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iRelief
;				   |								2 = Error setting $iCase
;				   |								4 = Error setting $bHidden
;				   |								8 = Error setting $bOutline
;				   |								16 = Error setting $bShadow
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharEffect(ByRef $oObj, $iRelief, $iCase, $bHidden, $bOutline, $bShadow)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avEffect[5]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If __LOWriter_VarsAreNull($iRelief, $iCase, $bHidden, $bOutline, $bShadow) Then
		__LOWriter_ArrayFill($avEffect, $oObj.CharRelief(), $oObj.CharCaseMap(), $oObj.CharHidden(), $oObj.CharContoured(), $oObj.CharShadowed())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avEffect)
	EndIf

	If ($iRelief <> Null) Then
		If Not __LOWriter_IntIsBetween($iRelief, $LOW_RELIEF_NONE, $LOW_RELIEF_ENGRAVED) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oObj.CharRelief = $iRelief
		$iError = ($oObj.CharRelief() = $iRelief) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iCase <> Null) Then
		If Not __LOWriter_IntIsBetween($iCase, $LOW_CASEMAP_NONE, $LOW_CASEMAP_SM_CAPS) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oObj.CharCaseMap = $iCase
		$iError = ($oObj.CharCaseMap() = $iCase) ? $iError : BitOR($iError, 2)
	EndIf

	If ($bHidden <> Null) Then
		If Not IsBool($bHidden) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oObj.CharHidden = $bHidden
		$iError = ($oObj.CharHidden() = $bHidden) ? $iError : BitOR($iError, 4)
	EndIf

	If ($bOutline <> Null) Then
		If Not IsBool($bOutline) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oObj.CharContoured = $bOutline
		$iError = ($oObj.CharContoured() = $bOutline) ? $iError : BitOR($iError, 8)
	EndIf

	If ($bShadow <> Null) Then
		If Not IsBool($bShadow) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$oObj.CharShadowed = $bShadow
		$iError = ($oObj.CharShadowed() = $bShadow) ? $iError : BitOR($iError, 16)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_CharEffect

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharFont
; Description ...: Set and Retrieve the Font Settings
; Syntax ........: __LOWriter_CharFont(Byref $oObj, $sFontName, $nFontSize, $iPosture, $iWeight)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $sFontName           - a string value. The Font Name to change to.
;                  $nFontSize           - a general number value. The new Font size.
;                  $iPosture            - an integer value (0-5). Italic setting. See Constants, $LOW_POSTURE_* as defined in LibreOfficeWriter_Constants.au3. Also see remarks.
;                  $iWeight             - an integer value (0,50-200). Bold settings see Constants, $LOW_WEIGHT_* as defined in LibreOfficeWriter_Constants.au3. Also see remarks.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 5 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 6 Return 0 = $sFontName not a String.
;				   @Error 1 @Extended 7 Return 0 = $nFontSize not a Number.
;				   @Error 1 @Extended 8 Return 0 = $iPosture not an Integer, less than 0 or greater than 5. See Constants, $LOW_POSTURE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 9 Return 0 = $iWeight less than 50 and not 0, or more than 200. See Constants, $LOW_WEIGHT_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sFontName
;				   |								2 = Error setting $nFontSize
;				   |								4 = Error setting $iPosture
;				   |								8 = Error setting $iWeight
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;					Not every font accepts Bold and Italic settings, and not all settings for bold and Italic are accepted,
;					such as oblique, ultra Bold etc. Libre Writer accepts only the predefined weight values, any other values
;					are changed automatically to an acceptable value, which could trigger a settings error.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharFont(ByRef $oObj, $sFontName, $nFontSize, $iPosture, $iWeight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFont[4]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	If __LOWriter_VarsAreNull($sFontName, $nFontSize, $iPosture, $iWeight) Then
		__LOWriter_ArrayFill($avFont, $oObj.CharFontName(), $oObj.CharHeight(), $oObj.CharPosture(), $oObj.CharWeight())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avFont)
	EndIf

	If ($sFontName <> Null) Then
		If Not IsString($sFontName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oObj.CharFontName = $sFontName
		$iError = ($oObj.CharFontName() = $sFontName) ? $iError : BitOR($iError, 1)
	EndIf

	If ($nFontSize <> Null) Then
		If Not IsNumber($nFontSize) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oObj.CharHeight = $nFontSize
		$iError = ($oObj.CharHeight() = $nFontSize) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iPosture <> Null) Then
		If Not __LOWriter_IntIsBetween($iPosture, $LOW_POSTURE_NONE, $LOW_POSTURE_ITALIC) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$oObj.CharPosture = $iPosture
		$iError = ($oObj.CharPosture() = $iPosture) ? $iError : BitOR($iError, 4)
	EndIf

	If ($iWeight <> Null) Then
		If Not __LOWriter_IntIsBetween($iWeight, $LOW_WEIGHT_THIN, $LOW_WEIGHT_BLACK, "", $LOW_WEIGHT_DONT_KNOW) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		$oObj.CharWeight = $iWeight
		$iError = ($oObj.CharWeight() = $iWeight) ? $iError : BitOR($iError, 8)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_CharFont

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharFontColor
; Description ...: Set or retrieve the font color, transparency and highlighting values.
; Syntax ........: __LOWriter_CharFontColor(Byref $oObj, $iFontColor, $iTransparency, $iHighlight)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $iFontColor          - an integer value (-1-16777215). The desired Color value in Long Integer format, to make the font, can a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3. Set to $LOW_COLOR_OFF(-1) for Auto color.
;                  $iTransparency       - an integer value. Transparency percentage. 0 is not visible, 100 is fully visible. Available for Libre Office 7.0 and up.
;                  $iHighlight          - an integer value (-1-16777215). A Color value in Long Integer format, to highlight the text in, can a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3. Set to $LOW_COLOR_OFF(-1) for No color.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iFontColor not an integer, less than -1 or greater than 16777215.
;				   @Error 1 @Extended 5 Return 0 = $iTransparency not an Integer, or less than 0 or greater than 100%.
;				   @Error 1 @Extended 6 Return 0 = $iHighlight not an integer, less than -1 or greater than 16777215.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $FontColor
;				   |								2 = Error setting $iTransparency.
;				   |								4 = Error setting $iHighlight
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 7.0.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 or 3 Element Array with values in order of function parameters. If The current Libre Office version is below 7.0 the returned array will contain 2 elements, because $iTransparency is not available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;					Note: When setting transparency, the value of font color value is changed, this may lead to property
;					setting error messages for setting the Font color.
; Related .......: _LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharFontColor(ByRef $oObj, $iFontColor, $iTransparency, $iHighlight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avColor[2]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If __LOWriter_VarsAreNull($iFontColor, $iTransparency, $iHighlight) Then
		If __LOWriter_VersionCheck(7.0) Then
			__LOWriter_ArrayFill($avColor, $oObj.CharColor(), $oObj.CharTransparence(), $oObj.CharBackColor())
		Else
			__LOWriter_ArrayFill($avColor, $oObj.CharColor(), $oObj.CharBackColor())
		EndIf
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avColor)
	EndIf

	If ($iFontColor <> Null) Then
		If Not __LOWriter_IntIsBetween($iFontColor, $LOW_COLOR_OFF, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oObj.CharColor = $iFontColor
		$iError = ($oObj.CharColor() = $iFontColor) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iTransparency <> Null) Then
		If Not __LOWriter_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		If Not __LOWriter_VersionCheck(7.0) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
		$oObj.CharTransparence = $iTransparency
		$iError = ($oObj.CharTransparence() = $iTransparency) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iHighlight <> Null) Then
		If Not __LOWriter_IntIsBetween($iHighlight, $LOW_COLOR_OFF, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		; CharHighlight; same as CharBackColor---Libre seems to use back color for highlighting however, so using that for setting.
;~ 		If Not __LOWriter_VersionCheck(4.2) Then Return SetError($__LOW_STATUS_VER_ERROR,2,0)
;~ 		$oObj.CharHighlight = $iHighlight ;-- keeping old method in case.
;~ 		$iError = ($oObj.CharHighlight() = $iHighlight) ? $iError : BitOR($iError,4)
		$oObj.CharBackColor = $iHighlight
		$iError = ($oObj.CharBackColor() = $iHighlight) ? $iError : BitOR($iError, 4)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_CharFontColor

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharOverLine
; Description ...: Set and retrieve the OverLine settings.
; Syntax ........: __LOWriter_CharOverLine(Byref $oObj, $bWordOnly, $iOverLineStyle, $bOLHasColor, $iOLColor)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $bWordOnly           - a boolean value. If true, white spaces are not Overlined.
;                  $iOverLineStyle      - an integer value (0-18). The style of the Overline line, see constants, $LOW_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $bOLHasColor         - a boolean value. Whether the Overline is colored, must be set to true in order to set the Overline color.
;                  $iOLColor            - an integer value (-1-16777215). The color of the Overline, set in Long integer format. Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3. Set to $LOW_COLOR_OFF(-1) for automatic color mode.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iOverLineStyle not an Integer, or less than 0 or greater than 18. See constants, $LOW_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 6 Return 0 = $bOLHasColor not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $iOLColor not an Integer, or less than -1 or greater than 16777215.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bWordOnly
;				   |								2 = Error setting $iOverLineStyle
;				   |								4 = Error setting $OLHasColor
;				   |								8 = Error setting $iOLColor
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: OverLine line style uses the same constants as underline style.
;					Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;					Note: $bOLHasColor must be set to true in order to set the Overline color.
; Related .......: _LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharOverLine(ByRef $oObj, $bWordOnly, $iOverLineStyle, $bOLHasColor, $iOLColor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avOverLine[4]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If __LOWriter_VarsAreNull($bWordOnly, $iOverLineStyle, $bOLHasColor, $iOLColor) Then
		__LOWriter_ArrayFill($avOverLine, $oObj.CharWordMode(), $oObj.CharOverline(), $oObj.CharOverlineHasColor(), $oObj.CharOverlineColor())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avOverLine)
	EndIf

	If ($bWordOnly <> Null) Then
		If Not IsBool($bWordOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oObj.CharWordMode = $bWordOnly
		$iError = ($oObj.CharWordMode() = $bWordOnly) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iOverLineStyle <> Null) Then
		If Not __LOWriter_IntIsBetween($iOverLineStyle, $LOW_UNDERLINE_NONE, $LOW_UNDERLINE_BOLD_WAVE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oObj.CharOverline = $iOverLineStyle
		$iError = ($oObj.CharOverline() = $iOverLineStyle) ? $iError : BitOR($iError, 2)
	EndIf

	If ($bOLHasColor <> Null) Then
		If Not IsBool($bOLHasColor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oObj.CharOverlineHasColor = $bOLHasColor
		$iError = ($oObj.CharOverlineHasColor() = $bOLHasColor) ? $iError : BitOR($iError, 4)
	EndIf

	If ($iOLColor <> Null) Then
		If Not __LOWriter_IntIsBetween($iOLColor, $LOW_COLOR_OFF, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oObj.CharOverlineColor = $iOLColor
		$iError = ($oObj.CharOverlineColor() = $iOLColor) ? $iError : BitOR($iError, 8)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_CharOverLine

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharPosition
; Description ...: Set and retrieve settings related to Sub/Super Script and relative size.
; Syntax ........: __LOWriter_CharPosition(Byref $oObj, $bAutoSuper, $iSuperScript, $bAutoSub, $iSubScript, $iRelativeSize)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $bAutoSuper          -  a boolean value. Whether to active automatic sizing for SuperScript.
;                  $iSuperScript        -  an integer value.  SuperScript percentage value. See Remarks.
;                  $bAutoSub            -  a boolean value. Whether to active automatic sizing for SubScript.
;                  $iSubScript          -  an integer value. SubScript percentage value. See Remarks.
;                  $iRelativeSize       -  an integer value. Percentage relative to current font size. Min. 1, Max 100.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bAutoSuper not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bAutoSub not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iSuperScript not an integer, less than 0, higher than 100 and Not 14000.
;				   @Error 1 @Extended 7 Return 0 = $iSubScript not an integer, less than -100, higher than 100 and Not 14000.
;				   @Error 1 @Extended 8 Return 0 = $iRelativeSize not an integer, or less than 1, higher than 100.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iSuperScript
;				   |								2 = Error setting $iSubScript
;				   |								4 = Error setting $iRelativeSize.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;					Set either $iSubScript or $iSuperScript to 0 to return it to Normal setting.
;					The way LibreOffice is set up Super/SubScript are set in the same setting, Super is a positive number from
;						1 to 100 (percentage), SubScript is a negative number set to 1 to 100 percentage. For the user's
;						convenience this function accepts both positive and negative numbers for SubScript, if a positive number
;						is called for SubScript, it is automatically set to a negative. Automatic Superscript has a integer
;						value of 14000, Auto SubScript has a integer value of -14000. There is no settable setting of Automatic
;						Super/Sub Script, though one exists, it is read-only in LibreOffice, consequently I have made two
;						separate parameters to be able to determine if the user wants to automatically set SuperScript or
;						SubScript. If you set both Auto SuperScript to True and Auto SubScript to True, or $iSuperScript to an
;						integer and $iSubScript to an integer, Subscript will be set as it is the last in the line to be set in
;						this function, and thus will over-write any SuperScript settings.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharPosition(ByRef $oObj, $bAutoSuper, $iSuperScript, $bAutoSub, $iSubScript, $iRelativeSize)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avPosition[5]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If __LOWriter_VarsAreNull($bAutoSuper, $iSuperScript, $bAutoSub, $iSubScript, $iRelativeSize) Then
		__LOWriter_ArrayFill($avPosition, ($oObj.CharEscapement() = 14000) ? True : False, ($oObj.CharEscapement() > 0) ? $oObj.CharEscapement() : 0, _
				($oObj.CharEscapement() = -14000) ? True : False, ($oObj.CharEscapement() < 0) ? $oObj.CharEscapement() : 0, $oObj.CharEscapementHeight())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avPosition)
	EndIf

	If ($bAutoSuper <> Null) Then
		If Not IsBool($bAutoSuper) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		; If $bAutoSuper = True set it to 14000 (automatic superScript) else if $iSuperScript is set, let that overwrite
		;	the current setting, else if subscript is true or set to an integer, it will overwrite the setting. If nothing
		; else set SubScript to 1
		$iSuperScript = ($bAutoSuper) ? 14000 : (IsInt($iSuperScript)) ? $iSuperScript : (IsInt($iSubScript) Or ($bAutoSub = True)) ? $iSuperScript : 1
	EndIf

	If ($bAutoSub <> Null) Then
		If Not IsBool($bAutoSub) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		; If $bAutoSub = True set it to -14000 (automatic SubScript) else if $iSubScript is set, let that overwrite
		;	the current setting, else if superscript is true or set to an integer, it will overwrite the setting.
		$iSubScript = ($bAutoSub) ? -14000 : (IsInt($iSubScript)) ? $iSubScript : (IsInt($iSuperScript)) ? $iSubScript : 1

	EndIf

	If ($iSuperScript <> Null) Then
		If Not __LOWriter_IntIsBetween($iSuperScript, 0, 100, "", 14000) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oObj.CharEscapement = $iSuperScript
		$iError = ($oObj.CharEscapement() = $iSuperScript) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iSubScript <> Null) Then
		If Not __LOWriter_IntIsBetween($iSubScript, -100, 100, "", "-14000:14000") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$iSubScript = ($iSubScript > 0) ? Int("-" & $iSubScript) : $iSubScript
		$oObj.CharEscapement = $iSubScript
		$iError = ($oObj.CharEscapement() = $iSubScript) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iRelativeSize <> Null) Then
		If Not __LOWriter_IntIsBetween($iRelativeSize, 1, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$oObj.CharEscapementHeight = $iRelativeSize
		$iError = ($oObj.CharEscapementHeight() = $iRelativeSize) ? $iError : BitOR($iError, 4)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_CharPosition

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharRotateScale
; Description ...: Set or retrieve the character rotational and Scale settings.
; Syntax ........: __LOWriter_CharRotateScale(Byref $oObj, $iRotation, $iScaleWidth[, $bRotateFitLine = Null])
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $iRotation           - an integer value. Degrees to rotate the text. Accepts only 0, 90, and 270 degrees.
;                  $iScaleWidth         - an integer value. The percentage to  horizontally stretch or compress the text. Min. 1. Max 100. 100 is normal sizing.
;                  $bRotateFitLine      - [optional] a boolean value. Default is Null. If True, Stretches or compresses the selected text so that it fits between the line that is above the text and the line that is below the text. Only works with Direct Formatting.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iRotation not an Integer or not equal to 0, 90 or 270 degrees.
;				   @Error 1 @Extended 5 Return 0 = $iScaleWidth not an Integer or less than 1% or greater than 100%.
;				   @Error 1 @Extended 6 Return 0 = $bRotateFitLine not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iRotation
;				   |								2 = Error setting $iScaleWidth
;				   |								4 = Error setting $bRotateFitLine
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters. Note: Excludes $bRotateFitLine, which is added onto the Direct Formatting function return.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharRotateScale(ByRef $oObj, $iRotation, $iScaleWidth, $bRotateFitLine = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avRotation[2]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If __LOWriter_VarsAreNull($iRotation, $iScaleWidth, $bRotateFitLine) Then
		; rotation set in hundredths (90 deg = 900 etc)so divide by 10.
		__LOWriter_ArrayFill($avRotation, ($oObj.CharRotation() / 10), $oObj.CharScaleWidth())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avRotation)
	EndIf

	If ($iRotation <> Null) Then
		If Not __LOWriter_IntIsBetween($iRotation, 0, 0, "", "90:270") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$iRotation = ($iRotation > 0) ? ($iRotation * 10) : $iRotation ;rotation set in hundredths (90 deg = 900 etc)so times by 10.
		$oObj.CharRotation = $iRotation
		$iError = ($oObj.CharRotation() = $iRotation) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iScaleWidth <> Null) Then ; can't be less than 1%
		If Not __LOWriter_IntIsBetween($iScaleWidth, 1, 100) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oObj.CharScaleWidth = $iScaleWidth
		$iError = ($oObj.CharScaleWidth() = $iScaleWidth) ? $iError : BitOR($iError, 2)
	EndIf

	If ($bRotateFitLine <> Null) Then
		; works only on Direct Formatting:
		If Not IsBool($bRotateFitLine) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oObj.CharRotationIsFitToLine = $bRotateFitLine
		$iError = ($oObj.CharRotationIsFitToLine() = $bRotateFitLine) ? $iError : BitOR($iError, 4)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_CharRotateScale

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharShadow
; Description ...: Set and retrieve the Shadow for a CharacterStyle
; Syntax ........: __LOWriter_CharShadow(Byref $oObj, $iWidth, $iColor, $bTransparent, $iLocation)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $iWidth              - an integer value. Width of the shadow, set in Micrometers.
;                  $iColor              - an integer value (0-16777215). Color of the shadow. See Remarks. Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bTransparent        - a boolean value. Whether the shadow is transparent or not.
;                  $iLocation           - an integer value (0-4). Location of the shadow compared to the characters. See Constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iWidth not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iColor not an Integer, or less than 0 or greater than 16777215.
;				   @Error 1 @Extended 6 Return 0 = $bTransparent not a boolean.
;				   @Error 1 @Extended 7 Return 0 = $iLocation not an Integer, or less than 0 or greater than 4. See Constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Shadow format Object.
;				   @Error 2 @Extended 2 Return 0 = Error retrieving Shadow format Object for Error Checking.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iWidth
;				   |								2 = Error setting $iColor
;				   |								4 = Error setting $bTransparent
;				   |								8 = Error setting $iLocation
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;					Note: LibreOffice may adjust the set width +/- 1 Micrometer after setting.
;					Color is set in Long Integer format. You can use one of the below listed constants or a custom one.
; Related .......: _LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong,  _LOWriter_ConvertFromMicrometer,
;					_LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharShadow(ByRef $oObj, $iWidth, $iColor, $bTransparent, $iLocation)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $tShdwFrmt
	Local $avShadow[4]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	$tShdwFrmt = $oObj.CharShadowFormat()
	If Not IsObj($tShdwFrmt) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iWidth, $iColor, $bTransparent, $iLocation) Then
		__LOWriter_ArrayFill($avShadow, $tShdwFrmt.ShadowWidth(), $tShdwFrmt.Color(), $tShdwFrmt.IsTransparent(), $tShdwFrmt.Location())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avShadow)
	EndIf

	If ($iWidth <> Null) Then
		If Not IsInt($iWidth) Or ($iWidth < 0) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$tShdwFrmt.ShadowWidth = $iWidth
	EndIf

	If ($iColor <> Null) Then
		If Not __LOWriter_IntIsBetween($iColor, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$tShdwFrmt.Color = $iColor
	EndIf

	If ($bTransparent <> Null) Then
		If Not IsBool($bTransparent) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$tShdwFrmt.IsTransparent = $bTransparent
	EndIf

	If ($iLocation <> Null) Then
		If Not __LOWriter_IntIsBetween($iLocation, $LOW_SHADOW_NONE, $LOW_SHADOW_BOTTOM_RIGHT) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$tShdwFrmt.Location = $iLocation
	EndIf

	$oObj.CharShadowFormat = $tShdwFrmt
	$tShdwFrmt = $oObj.CharShadowFormat
	If Not IsObj($tShdwFrmt) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	$iError = ($iWidth = Null) ? $iError : ($tShdwFrmt.ShadowWidth() = $iWidth) ? $iError : BitOR($iError, 1)
	$iError = ($iColor = Null) ? $iError : ($tShdwFrmt.Color() = $iColor) ? $iError : BitOR($iError, 2)
	$iError = ($bTransparent = Null) ? $iError : ($tShdwFrmt.IsTransparent() = $bTransparent) ? $iError : BitOR($iError, 4)
	$iError = ($iLocation = Null) ? $iError : ($tShdwFrmt.Location() = $iLocation) ? $iError : BitOR($iError, 8)

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_CharShadow

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharSpacing
; Description ...: Set and retrieve the spacing between characters (Kerning).
; Syntax ........: __LOWriter_CharSpacing(Byref $oObj, $bAutoKerning, $nKerning)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $bAutoKerning        - a boolean value. True applies a spacing in between certain pairs of characters. False = disabled.
;                  $nKerning            - a general number value. The kerning value of the characters. Min is -2 Pt. Max is 928.8 Pt. See Remarks. Values are in Printer's Points as set in the Libre Office UI.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bAutoKerning not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $nKerning not a number, or less than -2 or greater than 928.8 Points.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bAutoKerning
;				   |								2 = Error setting $nKerning.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;					When setting Kerning values in LibreOffice, the measurement is listed in Pt (Printer's Points) in the User
;						Display, however the internal setting is measured in MicroMeters. They will be automatically converted
;						from Points to MicroMeters and back for retrieval of settings.
;						The acceptable values are from -2 Pt to  928.8 Pt. the figures can be directly converted easily,
;						however, for an unknown reason to myself, LibreOffice begins counting backwards and in negative
;						MicroMeters internally from 928.9 up to 1000 Pt (Max setting). For example, 928.8Pt is the last correct
;						value, which equals 32766 uM (MicroMeters), after this LibreOffice reports the following:
;						;928.9 Pt = -32766 uM;  929 Pt = -32763 uM; 929.1 = -32759; 1000 pt = -30258. Attempting to set Libre's
;						kerning value to  anything over 32768 uM causes a COM exception, and attempting to set the kerning to
;						any of these negative  numbers sets the User viewable kerning value to -2.0 Pt. For these reasons the
;						 max settable kerning  is -2.0 Pt to 928.8 Pt.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharSpacing(ByRef $oObj, $bAutoKerning, $nKerning)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avKerning[2]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If __LOWriter_VarsAreNull($bAutoKerning, $nKerning) Then
		$nKerning = __LOWriter_UnitConvert($oObj.CharKerning(), $__LOWCONST_CONVERT_UM_PT)
		__LOWriter_ArrayFill($avKerning, $oObj.CharAutoKerning(), (($nKerning > 928.8) ? 1000 : $nKerning))
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avKerning)
	EndIf

	If ($bAutoKerning <> Null) Then
		If Not IsBool($bAutoKerning) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oObj.CharAutoKerning = $bAutoKerning
		$iError = ($oObj.CharAutoKerning() = $bAutoKerning) ? $iError : BitOR($iError, 1)
	EndIf

	If ($nKerning <> Null) Then
		If Not __LOWriter_NumIsBetween($nKerning, -2, 928.8) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$nKerning = __LOWriter_UnitConvert($nKerning, $__LOWCONST_CONVERT_PT_UM)

		$oObj.CharKerning = $nKerning
		$iError = ($oObj.CharKerning() = $nKerning) ? $iError : BitOR($iError, 2)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_CharSpacing

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharStrikeOut
; Description ...: Set or Retrieve the StrikeOut settings,
; Syntax ........: __LOWriter_CharStrikeOut(Byref $oObj, $bWordOnly, $bStrikeOut, $iStrikeLineStyle)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $bWordOnly           - a boolean value. Whether to strike out words only and skip whitespaces. True = skip whitespaces.
;                  $bStrikeOut          - a boolean value. True = strikeout, False = no strike out.
;                  $iStrikeLineStyle    - an integer value. The Strikeout Line Style, see constants, $LOW_STRIKEOUT_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bStrikeOut not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iStrikeLineStyle not an Integer, or less than 0 or greater than 8. See constants, $LOW_STRIKEOUT_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bWordOnly
;				   |								2 = Error setting $bStrikeOut
;				   |								4 = Error setting $iStrikeLineStyle
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;					Note Strikeout converted to single line in Ms word document format.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharStrikeOut(ByRef $oObj, $bWordOnly, $bStrikeOut, $iStrikeLineStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avStrikeOut[3]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If __LOWriter_VarsAreNull($bWordOnly, $bStrikeOut, $iStrikeLineStyle) Then
		__LOWriter_ArrayFill($avStrikeOut, $oObj.CharWordMode(), $oObj.CharCrossedOut(), $oObj.CharStrikeout())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avStrikeOut)
	EndIf

	If ($bWordOnly <> Null) Then
		If Not IsBool($bWordOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oObj.CharWordMode = $bWordOnly
		$iError = ($oObj.CharWordMode() = $bWordOnly) ? $iError : BitOR($iError, 1)
	EndIf

	If ($bStrikeOut <> Null) Then
		If Not IsBool($bStrikeOut) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oObj.CharCrossedOut = $bStrikeOut
		$iError = ($oObj.CharCrossedOut() = $bStrikeOut) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iStrikeLineStyle <> Null) Then
		If Not __LOWriter_IntIsBetween($iStrikeLineStyle, $LOW_STRIKEOUT_NONE, $LOW_STRIKEOUT_X) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oObj.CharStrikeout = $iStrikeLineStyle
		$iError = ($oObj.CharStrikeout() = $iStrikeLineStyle) ? $iError : BitOR($iError, 4)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_CharStrikeOut

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharStyleNameToggle
; Description ...: Toggle from Character Style Display Name to Internal Name for error checking and setting retrieval.
; Syntax ........: __LOWriter_CharStyleNameToggle(Byref $sCharStyle[, $bReverse = False])
; Parameters ....: $sCharStyle          - a string value. The CharStyle Name to Toggle.
;                  $bReverse            - [optional] a boolean value. Default is False. If True, the Char Style name is reverse toggled.
; Return values .: Success: String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $sCharStyle not a String.
;				   @Error 1 @Extended 2 Return 0 = $bReverse not a Boolean.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. Character Style Name was successfully toggled. Returning toggled name as a string.
;				   @Error 0 @Extended 1 Return String = Success. Character Style Name was successfully reverse toggled. Returning toggled name as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharStyleNameToggle($sCharStyle, $bReverse = False)
	If Not IsString($sCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReverse) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If ($bReverse = False) Then
		$sCharStyle = ($sCharStyle = "Footnote Characters") ? "Footnote Symbol" : $sCharStyle
		$sCharStyle = ($sCharStyle = "Bullets") ? "Bullet Symbols" : $sCharStyle
		$sCharStyle = ($sCharStyle = "Endnote Characters") ? "Endnote Symbol" : $sCharStyle
		$sCharStyle = ($sCharStyle = "Quotation") ? "Citation" : $sCharStyle
		$sCharStyle = ($sCharStyle = "No Character Style") ? "Standard" : $sCharStyle
		Return SetError($__LOW_STATUS_SUCCESS, 0, $sCharStyle)
	Else
		$sCharStyle = ($sCharStyle = "Footnote Symbol") ? "Footnote Characters" : $sCharStyle
		$sCharStyle = ($sCharStyle = "Bullet Symbols") ? "Bullets" : $sCharStyle
		$sCharStyle = ($sCharStyle = "Endnote Symbol") ? "Endnote Characters" : $sCharStyle
		$sCharStyle = ($sCharStyle = "Citation") ? "Quotation" : $sCharStyle
		$sCharStyle = ($sCharStyle = "Standard") ? "No Character Style" : $sCharStyle
		Return SetError($__LOW_STATUS_SUCCESS, 1, $sCharStyle)
	EndIf
EndFunc   ;==>__LOWriter_CharStyleNameToggle

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CharUnderLine
; Description ...: Set and retrieve the UnderLine settings.
; Syntax ........: __LOWriter_CharUnderLine(Byref $oObj, $bWordOnly, $iUnderLineStyle, $bULHasColor, $iULColor)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.text.Paragraph" Or "com.sun.star.text.TextPortion" services, such as a Cursor with data selected or paragraph section.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If true, white spaces are not underlined.
;                  $iUnderLineStyle     - [optional] an integer value (0-18). Default is Null. The style of the Underline line, see constants, $LOW_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bULHasColor         - [optional] a boolean value. Default is Null. Whether the underline is colored, must be set to true in order to set the underline color.
;                  $iULColor            - [optional] an integer value (-1-16777215). Default is Null. The color of the underline, set in Long integer format. Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3. Set to $LOW_COLOR_OFF(-1) for automatic color mode.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iUnderLineStyle not an Integer, or less than 0 or greater than 18. See constants, $LOW_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 6 Return 0 = $bULHasColor not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $iULColor not an Integer, or less than -1 or greater than 16777215.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bWordOnly
;				   |								2 = Error setting $iUnderLineStyle
;				   |								4 = Error setting $ULHasColor
;				   |								8 = Error setting $iULColor
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;					Note: $bULHasColor must be set to true in order to set the underline color.
; Related .......: _LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CharUnderLine(ByRef $oObj, $bWordOnly, $iUnderLineStyle, $bULHasColor, $iULColor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avUnderLine[4]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If __LOWriter_VarsAreNull($bWordOnly, $iUnderLineStyle, $bULHasColor, $iULColor) Then
		__LOWriter_ArrayFill($avUnderLine, $oObj.CharWordMode(), $oObj.CharUnderline(), $oObj.CharUnderlineHasColor(), $oObj.CharUnderlineColor())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avUnderLine)
	EndIf

	If ($bWordOnly <> Null) Then
		If Not IsBool($bWordOnly) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oObj.CharWordMode = $bWordOnly
		$iError = ($oObj.CharWordMode() = $bWordOnly) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iUnderLineStyle <> Null) Then
		If Not __LOWriter_IntIsBetween($iUnderLineStyle, $LOW_UNDERLINE_NONE, $LOW_UNDERLINE_BOLD_WAVE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oObj.CharUnderline = $iUnderLineStyle
		$iError = ($oObj.CharUnderline() = $iUnderLineStyle) ? $iError : BitOR($iError, 2)
	EndIf

	If ($bULHasColor <> Null) Then
		If Not IsBool($bULHasColor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oObj.CharUnderlineHasColor = $bULHasColor
		$iError = ($oObj.CharUnderlineHasColor() = $bULHasColor) ? $iError : BitOR($iError, 4)
	EndIf

	If ($iULColor <> Null) Then
		If Not __LOWriter_IntIsBetween($iULColor, $LOW_COLOR_OFF, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oObj.CharUnderlineColor = $iULColor
		$iError = ($oObj.CharUnderlineColor() = $iULColor) ? $iError : BitOR($iError, 8)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_CharUnderLine

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CreateStruct
; Description ...: Retrieves a Struct.
; Syntax ........: __LOWriter_CreateStruct($sStructName)
; Parameters ....: $sStructName	- a string value. Name of structure to create.
; Return values .:Success: Structure.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $sStructName Value not a string
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object
;				   @Error 2 @Extended 2 Return 0 = Error retrieving requested Structure.
;				   --Success--
;				   @Error 0 @Extended 0 Return Structure = Success. Property Structure Returned
; Author ........: mLipok;
; Modified ......: donnyh13 - Added error checking.
; Remarks .......: From WriterDemo.au3 as modified by mLipok from WriterDemo.vbs found in the LibreOffice SDK examples.
; Related .......:
; Link ..........: https://www.autoitscript.com/forum/topic/204665-libreopenoffice-writer/?do=findComment&comment=1471711
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_CreateStruct($sStructName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oServiceManager, $tStruct

	If Not IsString($sStructName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$tStruct = $oServiceManager.Bridge_GetStruct($sStructName)
	If Not IsObj($tStruct) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $tStruct)
EndFunc   ;==>__LOWriter_CreateStruct

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_CursorGetText
; Description ...: Retrieves a Text object appropriate for the type of cursor.
; Syntax ........: __LOWriter_CursorGetText(Byref $oDoc, $oCursor)
; Parameters ....: $oDoc	    - [in/out] a Document object created by a preceding call to Open, Create, or Connect.
;                  $oCursor 	- [in/out] an object. A Text or View Cursor Object returned from any Cursor Object creation or retrieval functions.
; Return values .:Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc variable not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor variable not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to get Cursor data type.
;				   @Error 2 @Extended 2 Return 0 = Failed to create Text Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to create Object for creating TextObject.
;				   @Error 3 @Extended 2 Return 0 = Cursor is in an unknown data field.
;				   --Success--
;				   @Error 0 @Extended ? Return Object = Success, Text object was returned. @Extended will be one of the constants, $LOW_CURDATA_* as defined in LibreOfficeWriter_Constants.au3.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$oReturnedObj = __LOWriter_Internal_CursorGetDataType($oDoc, $oCursor, True)
	$iCursorDataType = @extended
	If @error Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	If Not IsObj($oReturnedObj) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

	Switch $iCursorDataType
		Case $LOW_CURDATA_BODY_TEXT, $LOW_CURDATA_FRAME, $LOW_CURDATA_FOOTNOTE, $LOW_CURDATA_ENDNOTE, $LOW_CURDATA_HEADER_FOOTER
			$oText = $oReturnedObj.getText()
			If Not IsObj($oText) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
			Return SetError($__LOW_STATUS_SUCCESS, $iCursorDataType, $oText)
		Case $LOW_CURDATA_CELL
			$oText = $oReturnedObj.getCellByName($oCursor.Cell.CellName)
			If Not IsObj($oText) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
			Return SetError($__LOW_STATUS_SUCCESS, $iCursorDataType, $oText)
		Case Else
			Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0)
	EndSwitch
EndFunc   ;==>__LOWriter_CursorGetText

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_DateStructCompare
; Description ...: Compare two date Structures to see if they are the same Date, Time, etc.
; Syntax ........: __LOWriter_DateStructCompare($tDateStruct1, $tDateStruct2)
; Parameters ....: $tDateStruct1        - a dll struct value. The First Date Structure.
;                  $tDateStruct2        - a dll struct value. The Second Date Structure.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return False = $tDateStruct1 not an Object.
;				   @Error 1 @Extended 2 Return False = $tDateStruct2 not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return False = Success. Dates/Times in $tDateStruct1 and $tDateStruct2 are not the same.
;				   @Error 0 @Extended 1 Return True = Success. Dates/Times in $tDateStruct1 and $tDateStruct2 are the same.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_DateStructCompare($tDateStruct1, $tDateStruct2)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($tDateStruct1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, False)
	If Not IsObj($tDateStruct2) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, False)

	If $tDateStruct1.Year() <> $tDateStruct2.Year() Then Return SetError($__LOW_STATUS_SUCCESS, 0, False)
	If $tDateStruct1.Month() <> $tDateStruct2.Month() Then Return SetError($__LOW_STATUS_SUCCESS, 0, False)
	If $tDateStruct1.Day() <> $tDateStruct2.Day() Then Return SetError($__LOW_STATUS_SUCCESS, 0, False)
	If $tDateStruct1.Hours() <> $tDateStruct2.Hours() Then Return SetError($__LOW_STATUS_SUCCESS, 0, False)
	If $tDateStruct1.Minutes() <> $tDateStruct2.Minutes() Then Return SetError($__LOW_STATUS_SUCCESS, 0, False)
	If $tDateStruct1.Seconds() <> $tDateStruct2.Seconds() Then Return SetError($__LOW_STATUS_SUCCESS, 0, False)
	If $tDateStruct1.NanoSeconds() <> $tDateStruct2.NanoSeconds() Then Return SetError($__LOW_STATUS_SUCCESS, 0, False)
	If __LOWriter_VersionCheck(4.1) Then
		If $tDateStruct1.IsUTC() <> $tDateStruct2.IsUTC() Then Return SetError($__LOW_STATUS_SUCCESS, 0, False)
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 1, True)
EndFunc   ;==>__LOWriter_DateStructCompare

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_DirFrmtCheck
; Description ...: Do checks on Dirformat input object.
; Syntax ........: __LOWriter_DirFrmtCheck(Byref $oSelection[, $bCheckSelection = False])
; Parameters ....: $oSelection          - [in/out] an object. The Object to check, which should be either a cursor with data selected or a paragraph object.
;                  $bCheckSelection     - [optional] a boolean value. Default is False. If True, check for whether the cursor object is collapsed (no data selected).
; Return values .: Success: Boolean.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = Success. If called Object is fit for DirectFormatting Use then return True, else False.
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
			$oSelection.supportsService("com.sun.star.text.TextPortion") Then Return SetError($__LOW_STATUS_SUCCESS, 0, True)

	; If Object is a cursor then return true if $bcheckSelection is false. Else test if cursor selection is collapsed, return
	; false if it is.
	If $oSelection.supportsService("com.sun.star.text.TextCursor") Or _
			$oSelection.supportsService("com.sun.star.text.TextViewCursor") Then
		If $bCheckSelection Then Return SetError($__LOW_STATUS_SUCCESS, 0, ($oSelection.IsCollapsed()) ? False : True) ; If collapsed return false meaning fail.
		Return SetError($__LOW_STATUS_SUCCESS, 0, True)
	EndIf
	Return SetError($__LOW_STATUS_SUCCESS, 0, False)
EndFunc   ;==>__LOWriter_DirFrmtCheck

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_FieldCountType
; Description ...: Determine a Count Field's type.
; Syntax ........: __LOWriter_FieldCountType($vInput)
; Parameters ....: $vInput              - a variant value. Either a Field Object to determine the appropriate integer Constant to return, or a Integer Constant to return the appropriate Field type String. See constants, $LOW_FIELD_COUNT_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: String or Integer.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $vInput is neither a String nor an Integer.
;				   @Error 1 @Extended 2 Return 0 = $vInput was an Object, but did not match any known counting fields.
;				   @Error 1 @Extended 3 Return 0 = $vInput was an Integer but is higher than the size of the array of Field types. See Constants, $LOW_FIELD_COUNT_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success. Counting Field type identified, returning FieldCountType Variable Integer.
;				   @Error 0 @Extended 1 Return String = Success. Counting Field type identified, returning Field Count Type String for CreateInstance function.
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
			If $vInput.supportsService($asFieldTypes[$i]) Then Return SetError($__LOW_STATUS_SUCCESS, 0, $i)
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
		Next
		Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0) ; No Hits

	ElseIf IsInt($vInput) Then
		If Not __LOWriter_IntIsBetween($vInput, 0, UBound($asFieldTypes) - 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
		Return SetError($__LOW_STATUS_SUCCESS, 1, $asFieldTypes[$vInput])
	Else
		Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0) ; Wrong VarType
	EndIf
EndFunc   ;==>__LOWriter_FieldCountType

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_FieldsGetList
; Description ...: Internal Function to retrieve a Field Object list.
; Syntax ........: __LOWriter_FieldsGetList(Byref $oDoc, $bSupportedServices, $bFieldType, $bFieldTypeNum, Byref $avFieldTypes)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bSupportedServices  - a boolean value. If True, adds a column to the array that has the supported service String for that particular Field, To assist in identifying the Field type.
;                  $bFieldType          - [optional] a boolean value. Default is True. If True, adds a column to the array that has the Field Type String for that particular Field as described by Libre Office. To assist in identifying the Field type.
;                  $bFieldTypeNum       - [optional] a boolean value. Default is True. If True, adds a column to the array that has the Field Type Constant Integer for that particular Field, to assist in identifying the Field type.
;                  $avFieldTypes        - [in/out] an array of variants. An Array of Field types to search for to return. Array will not be modified.
; Return values .: Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 3 Return 0 = $bSupportedServices not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bFieldType not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bFieldTypeNum not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $avFieldTypes not an Array.
;				   --Initialization Errors--
;				   @Error 2 @Extended 2 Return 0 = Failed to create enumeration of fields in document.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning Array of Text Field Objects with @Extended set to number of results. See Remarks for Array sizing.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:The Array can vary in the number of columns, if $bSupportedServices, $bFieldType, and $bFieldTypeNum are set
;					to False, the Array will be a single column. With each of the above listed options being set to True, a
;					column will be added in the order they are listed in the UDF parameters. The First column will always be the
;					Field Object.
;					Setting $bSupportedServices to True will add a Supported Service String column for the found Field.
;					Setting $bFieldType to True will add a Field type column for the found Field.
;					Setting $bFieldTypeNum to True will add a Field type Number column, matching the below constants, for the found Field.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	; Skip 2 to match other Funcs.
	If Not IsBool($bSupportedServices) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bFieldType) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bFieldTypeNum) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	If Not IsArray($avFieldTypes) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
	$iColumns = ($bSupportedServices = False) ? ($iColumns - 1) : $iColumns
	$iColumns = ($bFieldType = False) ? ($iColumns - 1) : $iColumns
	$iColumns = ($bFieldTypeNum = False) ? ($iColumns - 1) : $iColumns

	; If Supported Services Option is False, change the column position of FieldType
	$iFieldTypeCol = ($bSupportedServices = False) ? ($iFieldTypeCol - 1) : $iFieldTypeCol

	$iFieldTypeNumCol = ($bSupportedServices = False) ? ($iFieldTypeNumCol - 1) : $iFieldTypeNumCol
	$iFieldTypeNumCol = ($bFieldType = False) ? ($iFieldTypeNumCol - 1) : $iFieldTypeNumCol

	$oTextFields = $oDoc.getTextFields.createEnumeration()
	If Not IsObj($oTextFields) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	While $oTextFields.hasMoreElements()
		$oTextField = $oTextFields.nextElement()

		For $i = 0 To UBound($avFieldTypes) - 1

			If $oTextField.supportsService($avFieldTypes[$i][1]) Then

				$avTextFields[$iCount][0] = $oTextField

				If ($bSupportedServices = True) Then $avTextFields[$iCount][1] = $avFieldTypes[$i][1]
				If ($bFieldType = True) Then $avTextFields[$iCount][$iFieldTypeCol] = $oTextField.getPresentation(True)
				If ($bFieldTypeNum = True) Then $avTextFields[$iCount][$iFieldTypeNumCol] = $avFieldTypes[$i][0]

				$iCount += 1
				If ($iCount = UBound($avTextFields)) Then ReDim $avTextFields[$iCount * 2][$iColumns]
			EndIf
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next

	WEnd

	ReDim $avTextFields[$iCount][$iColumns]
	Return SetError($__LOW_STATUS_SUCCESS, $iCount, $avTextFields)
EndFunc   ;==>__LOWriter_FieldsGetList

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_FieldTypeServices
; Description ...: Retrieve an Array of Supported Service Names and Integer Constants to search for Fields.
; Syntax ........: __LOWriter_FieldTypeServices($iFieldType[, $bAdvancedServices = False[, $bDocInfoServices = False]])
; Parameters ....: $iFieldType          - an integer value. The Integer Constant Field type.
;                  $bAdvancedServices   - [optional] a boolean value. Default is False. If True, search in Advanced Field Type Array.
;                  $bDocInfoServices    - [optional] a boolean value. Default is False. If True, search in Document Information Field Type Array.
; Return values .: Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $iFieldType not an Integer.
;				   @Error 1 @Extended 2 Return 0 = $bAdvancedServices not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bDocInfoServices not a Boolean.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Something went wrong determining what Array to search/Return.
;				   --Success--
;				   @Error 0 @Extended 0 Return Array = Success. $iFieldType set to All, $bAdvancedServices and $bDocInfoServices both set to false, returning full regular Field Service list String Array.
;				   @Error 0 @Extended 1 Return Array = Success. $iFieldType set to All, $bAdvancedServices set to True and $bDocInfoServices set to false, returning full Advanced Field Service String list Array.
;				   @Error 0 @Extended 2 Return Array = Success. $iFieldType set to All, $bAdvancedServices set to False and $bDocInfoServices set to True, returning full DocInfo Field Service String list Array.
;				   @Error 0 @Extended 3 Return Array = Success. $iFieldType BitOr'd together, determining which flags are called from the specified Array. Returning Field Service String list Array.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_FieldTypeServices($iFieldType, $bAdvancedServices = False, $bDocInfoServices = False)
	Local $avFieldTypes[30][2] = [[$LOW_FIELD_TYPE_COMMENT, "com.sun.star.text.TextField.Annotation"], _
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
			[$LOW_FIELD_TYPE_TEMPLATE_NAME, "com.sun.star.text.TextField.TemplateName"], [$LOW_FIELD_TYPE_URL, "com.sun.star.text.TextField.URL"], _
			[$LOW_FIELD_TYPE_WORD_COUNT, "com.sun.star.text.TextField.WordCount"]]

	Local $avFieldAdvTypes[9][2] = [[$LOW_FIELDADV_TYPE_BIBLIOGRAPHY, "com.sun.star.text.TextField.Bibliography"], _
			[$LOW_FIELDADV_TYPE_DATABASE, "com.sun.star.text.TextField.Database"], [$LOW_FIELDADV_TYPE_DATABASE_NAME, "com.sun.star.text.TextField.DatabaseName"], _
			[$LOW_FIELDADV_TYPE_DATABASE_NEXT_SET, "com.sun.star.text.TextField.DatabaseNextSet"], [$LOW_FIELDADV_TYPE_DATABASE_NAME_OF_SET, "com.sun.star.text.TextField.DatabaseNumberOfSet"], _
			[$LOW_FIELDADV_TYPE_DATABASE_SET_NUM, "com.sun.star.text.TextField.DatabaseSetNumber"], [$LOW_FIELDADV_TYPE_DDE, "com.sun.star.text.TextField.DDE"], _
			[$LOW_FIELDADV_TYPE_INPUT_USER, "com.sun.star.text.TextField.InputUser"], [$LOW_FIELDADV_TYPE_USER, "com.sun.star.text.TextField.User"]]

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

	If Not IsInt($iFieldType) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bAdvancedServices) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bDocInfoServices) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If ($bAdvancedServices = False) And ($bDocInfoServices = False) Then
		If (BitAND($iFieldType, $LOW_FIELD_TYPE_ALL)) Then Return SetError($__LOW_STATUS_SUCCESS, 0, $avFieldTypes)
		$avSearch = $avFieldTypes

	ElseIf ($bAdvancedServices = True) And ($bDocInfoServices = False) Then
		If (BitAND($iFieldType, $LOW_FIELDADV_TYPE_ALL)) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $avFieldAdvTypes)
		$avSearch = $avFieldAdvTypes

	ElseIf ($bDocInfoServices = True) And ($bAdvancedServices = False) Then
		If (BitAND($iFieldType, $LOW_FIELD_DOCINFO_TYPE_ALL)) Then Return SetError($__LOW_STATUS_SUCCESS, 2, $avFieldDocInfoTypes)
		$avSearch = $avFieldDocInfoTypes

	Else
		Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	For $i = 0 To UBound($avSearch) - 1
		If BitAND($avSearch[$i][0], $iFieldType) Then
			$avFieldResults[$iCount][0] = $avSearch[$i][0]
			$avFieldResults[$iCount][1] = $avSearch[$i][1]
			$iCount += 1
		EndIf
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
	Next

	ReDim $avFieldResults[$iCount][2]

	Return SetError($__LOW_STATUS_SUCCESS, 3, $avFieldResults)
EndFunc   ;==>__LOWriter_FieldTypeServices

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_FilterNameGet
; Description ...: Retrieves the correct L.O. Filtername for use in SaveAs and Export.
; Syntax ........: __LOWriter_FilterNameGet(Byref $sDocSavePath[, $bIncludeExportFilters = False])
; Parameters ....: $sDocSavePath           - [in/out] a string value. Full path with extension.
;                  $bIncludeExportFilters  - [optional] a boolean value. Default is False. If True, includes the FilterNames that can be used to Export only, in the search.
; Return values .:Success: String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $sDocSavePath is not a string.
;				   @Error 1 @Extended 2 Return 0 = $bIncludeExportFilters not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sDocSavePath is not a correct path or URL.
;				   --Success--
;				   @Error 0 @Extended 1 Return String = Success. Returns required filtername from "SaveAs" FilterNames.
;				   @Error 0 @Extended 2 Return String = Success. Returns required filtername from "Export" FilterNames.
;				   @Error 0 @Extended 3 Return String = FilterName not found for given file extension, defaulting to .odt file format and updating savepath accordingly.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Searches a predefined list of extensions stored in an array. Not all FilterNames are listed, where multiple
;						options were available for a given extension, the most recent Filter was used; Such as for .doc, there
;						is the a 97 MsWord Filter available, and also a 95 MsWord, in this case 97 MsWord was used, as is listed
;						in the "SaveAs" and "Export" dialogs For finding you own FilterNames, see convertfilters.html found in
;						L.O Install Folder: LibreOffice\help\en-US\text\shared\guide -- Or See: "OOME_3_0.odt / .pdf",
;						"OpenOffice.org Macros Explained OOME Third Edition" by Andrew D. Pitonyak, which has a handy Macro for
;						listing all FilterNames, found on page 284 of the above book in the ODT format.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_FilterNameGet(ByRef $sDocSavePath, $bIncludeExportFilters = False)
	Local $iLength, $iSlashLocation, $iDotLocation
	Local Const $STR_NOCASESENSE = 0, $STR_STRIPALL = 8
	Local $sFileExtension, $sFilterName
	Local $msSaveAsFilters[], $msExportFilters[]

	If Not IsString($sDocSavePath) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bIncludeExportFilters) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
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

	If $bIncludeExportFilters Then
		$msExportFilters[".epub"] = "EPUB"
		$msExportFilters[".jfif"] = "writer_jpg_Export"
		$msExportFilters[".jif"] = "writer_jpg_Export"
		$msExportFilters[".jjpe"] = "writer_jpg_Export"
		$msExportFilters[".jpg"] = "writer_jpg_Export"
		$msExportFilters[".jpeg"] = "writer_jpg_Export"
		$msExportFilters[".pdf"] = "writer_pdf_Export"
		$msExportFilters[".png"] = "writer_png_Export"
		$msExportFilters[".xhtml"] = "XHTML Writer File"
	EndIf

	If StringInStr($sDocSavePath, "file:///") Then ;  If L.O. URl Then
		$iSlashLocation = StringInStr($sDocSavePath, "/", $STR_NOCASESENSE, -1)
		$iDotLocation = StringInStr($sDocSavePath, ".", $STR_NOCASESENSE, -1, $iLength, $iLength - $iSlashLocation)
		$sFileExtension = StringRight($sDocSavePath, ($iLength - $iDotLocation) + 1)
	ElseIf StringInStr($sDocSavePath, "\") Then ;  Else if PC Path Then
		$iSlashLocation = StringInStr($sDocSavePath, "\", $STR_NOCASESENSE, -1)
		$iDotLocation = StringInStr($sDocSavePath, ".", $STR_NOCASESENSE, -1, $iLength, $iLength - $iSlashLocation)
		$sFileExtension = StringRight($sDocSavePath, $iLength - $iDotLocation + 1)
	Else
		Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	EndIf

	If $sFileExtension = $sDocSavePath Then ;  If no file extension identified, append .odt extension and return.
		$sDocSavePath = $sDocSavePath & ".odt"
		Return SetError($__LOW_STATUS_SUCCESS, 3, "writer8")
	Else
		$sFileExtension = StringLower(StringStripWS($sFileExtension, $STR_STRIPALL))
	EndIf

	$sFilterName = $msSaveAsFilters[$sFileExtension]

	If IsString($sFilterName) Then Return SetError($__LOW_STATUS_SUCCESS, 1, $sFilterName)

	If $bIncludeExportFilters Then $sFilterName = $msExportFilters[$sFileExtension]

	If IsString($sFilterName) Then Return SetError($__LOW_STATUS_SUCCESS, 2, $sFilterName)

	$sDocSavePath = StringReplace($sDocSavePath, $sFileExtension, ".odt") ; If No results, replace with ODT extension.

	Return SetError($__LOW_STATUS_SUCCESS, 3, "writer8")
EndFunc   ;==>__LOWriter_FilterNameGet

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_FindFormatAddSetting
; Description ...: Add or Update a setting in a Find Format Array.
; Syntax ........: __LOWriter_FindFormatAddSetting(Byref $aArray, $tSetting)
; Parameters ....: $aArray              - [in/out] an array of structs. A Find Format Array of Settings to Search. Array will be directly modified.
;                  $tSetting            - a struct value. A Libre Office Structure setting Object.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $aArray not an Array.
;				   @Error 1 @Extended 2 Return 0 = $tSetting not an Object.
;				   @Error 1 @Extended 3 Return 0 = $sSettingName not a String.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Setting was successfully updated or added.
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

	If Not IsArray($atArray) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($tSetting) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	$sSettingName = $tSetting.Name()
	If Not IsString($sSettingName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	For $i = 0 To UBound($atArray) - 1
		If $atArray[$i].Name() = $sSettingName Then
			$atArray[$i].Value = $tSetting.Value()
			$bFound = True
			ExitLoop
		EndIf

		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
	Next

	If ($bFound = False) Then
		ReDim $atArray[UBound($atArray) + 1]
		$atArray[UBound($atArray) - 1] = $tSetting
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_FindFormatAddSetting

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_FindFormatDeleteSetting
; Description ...: Delete a setting from a Find Format Array.
; Syntax ........: __LOWriter_FindFormatDeleteSetting(Byref $aArray, $sSettingName)
; Parameters ....: $aArray              - [in/out] an array of structs. A Find Format Array of Settings to Search. Array will be directly modified.
;                  $sSettingName        - a string value. The setting name to search and delete.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $aArray not an Array
;				   @Error 1 @Extended 2 Return 0 = $sSettingName not a String.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Setting was either not found or was successfully deleted.
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

	If Not IsArray($atArray) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sSettingName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	For $i = 0 To UBound($atArray) - 1
		If $atArray[$i].Name() <> $sSettingName Then
			$atArray[$iCount] = $atArray[$i]
			$iCount += 1
		EndIf
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
	Next
	ReDim $atArray[$iCount]
	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_FindFormatDeleteSetting

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_FindFormatRetrieveSetting
; Description ...: Retrieve a specific setting from a Find Format Array of Settings.
; Syntax ........: __LOWriter_FindFormatRetrieveSetting(Byref $aArray, $sSettingName)
; Parameters ....: $aArray              - [in/out] an array of structs. A Find Format Array of Settings to Search. Array will not be modified.
;                  $sSettingName        - a string value. The Setting name to search for.
; Return values .: Success: Object or 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $aArray not an Array.
;				   @Error 1 @Extended 2 Return 0 = $sSettingName not a String.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Search was successful, but setting was not found.
;				   @Error 0 @Extended 1 Return Object = Success. Setting found, returning requested setting Object.
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

	If Not IsArray($atArray) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sSettingName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	For $i = 0 To UBound($atArray) - 1
		If $atArray[$i].Name() = $sSettingName Then Return SetError($__LOW_STATUS_SUCCESS, 1, $atArray[$i])
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
	Next

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_FindFormatRetrieveSetting

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_FooterBorder
; Description ...: Header Border Setting Internal function.
; Syntax ........: __LOWriter_FooterBorder(Byref $oObj, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
; Parameters ....: $oObj                - [in/out] an object. Footer Object.
;                  $bWid                - a boolean value. If True the calling function is for setting Border Line Width.
;                  $bSty                - a boolean value. If True the calling function is for setting Border Line Style.
;                  $bCol                - a boolean value. If True the calling function is for setting Border Line Color.
;                  $iTop                - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iBottom             - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iLeft               - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iRight              - an integer value. See Border Style, Width, and Color functions for possible values.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oObj Variable not Object type variable.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style/Color when Border width not set.
;				   @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border style/Color when Border width not set.
;				   @Error 4 @Extended 3 Return 0 = Cannot set Left Border style/Color when Border width not set.
;				   @Error 4 @Extended 4 Return 0 = Cannot set Right Border style/Color when Border width not set.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with all other parameters set to Null keyword, and $bWid, or $bSty, or $bCol set to true to get the corresponding current settings.
;					All distance values are set in Micrometers. Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_FooterBorder(ByRef $oObj, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aiBorder[4]
	Local $tBL2

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If (($bWid + $bSty + $bCol) <> 1) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0) ; If more than one Boolean is true = error

	If __LOWriter_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then
		If $bWid Then
			__LOWriter_ArrayFill($aiBorder, $oObj.FooterTopBorder.LineWidth(), $oObj.FooterBottomBorder.LineWidth(), _
					$oObj.FooterLeftBorder.LineWidth(), $oObj.FooterRightBorder.LineWidth())
		ElseIf $bSty Then
			__LOWriter_ArrayFill($aiBorder, $oObj.FooterTopBorder.LineStyle(), $oObj.FooterBottomBorder.LineStyle(), _
					$oObj.FooterLeftBorder.LineStyle(), $oObj.FooterRightBorder.LineStyle())
		ElseIf $bCol Then
			__LOWriter_ArrayFill($aiBorder, $oObj.FooterTopBorder.Color(), $oObj.FooterBottomBorder.Color(), $oObj.FooterLeftBorder.Color(), _
					$oObj.FooterRightBorder.Color())
		EndIf
		Return SetError($__LOW_STATUS_SUCCESS, 1, $aiBorder)
	EndIf

	$tBL2 = __LOWriter_CreateStruct("com.sun.star.table.BorderLine2")
	If Not IsObj($tBL2) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If $iTop <> Null Then
		If Not $bWid And ($oObj.FooterTopBorder.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 1, 0) ; If Width not set, cant set color or style.
		; Top Line
		$tBL2.LineWidth = ($bWid) ? $iTop : $oObj.FooterTopBorder.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iTop : $oObj.FooterTopBorder.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iTop : $oObj.FooterTopBorder.Color() ; copy Color over to new size structure
		$oObj.FooterTopBorder = $tBL2
	EndIf

	If $iBottom <> Null Then
		If Not $bWid And ($oObj.FooterBottomBorder.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 2, 0) ; If Width not set, cant set color or style.
		; Bottom Line
		$tBL2.LineWidth = ($bWid) ? $iBottom : $oObj.FooterBottomBorder.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iBottom : $oObj.FooterBottomBorder.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iBottom : $oObj.FooterBottomBorder.Color() ; copy Color over to new size structure
		$oObj.FooterBottomBorder = $tBL2
	EndIf

	If $iLeft <> Null Then
		If Not $bWid And ($oObj.FooterLeftBorder.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 3, 0) ; If Width not set, cant set color or style.
		; Left Line
		$tBL2.LineWidth = ($bWid) ? $iLeft : $oObj.FooterLeftBorder.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iLeft : $oObj.FooterLeftBorder.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iLeft : $oObj.FooterLeftBorder.Color() ; copy Color over to new size structure
		$oObj.FooterLeftBorder = $tBL2
	EndIf

	If $iRight <> Null Then
		If Not $bWid And ($oObj.FooterRightBorder.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 4, 0) ; If Width not set, cant set color or style.
		; Right Line
		$tBL2.LineWidth = ($bWid) ? $iRight : $oObj.FooterRightBorder.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iRight : $oObj.FooterRightBorder.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iRight : $oObj.FooterRightBorder.Color() ; copy Color over to new size structure
		$oObj.FooterRightBorder = $tBL2
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_FooterBorder

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ObjRelativeSize
; Description ...: Calculate appropriate values to set Frame, Frame Style or Image Width or Height, when using relative values.
; Syntax ........: __LOWriter_ObjRelativeSize(Byref $oDoc, Byref $oObj[, $bRelativeWidth = False[, $bRelativeHeight = False]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oObj                - [in/out] an object. A Frame or Frame Style object returned by previous _LOWriter_FrameStyleCreate, _LOWriter_FrameCreate, _LOWriter_FrameStyleGetObj, _LOWriter_FrameGetObjByName, or _LOWriter_FrameGetObjByCursor function. Can also be an Image object returned by a previous _LOWriter_ImageInsert, or _LOWriter_ImageGetObjByName function.
;                  $bRelativeWidth      - [optional] a boolean value. Default is False. If True, modify Width based on relative Width percentage.
;                  $bRelativeHeight     - [optional] a boolean value. Default is False. If True, modify Height based on relative Height percentage.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oObj not an Object.
;				   @Error 1 @Extended 3 Return 0 = $bRelativeWidth not a boolean.
;				   @Error 1 @Extended 4 Return 0 = $bRelativeHeight not a boolean.
;				   @Error 1 @Extended 5 Return 0 = $bRelativeHeight and $bRelativeWidth both set to False.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Retrieving PageStyle Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function isn't totally necessary, because when setting Relative Width/Height, for a Frame/Frame
;					style, the frame is still appropriately set to the correct percentage. However, the L.O. U.I. does not
;					show the percentage unless you set a width value for the frame or frame style based on the Page width.
;					For Frame Styles, If you notice in L.O. when you set the relative value, while the Viewcursor is in one
;					PageStyle, and then move the cursor to another type of page style, the percentage changes. So when I am
;					modifying a FrameStyle obtain the ViewCursor, retrieve what PageStyle it is currently in, and calculate the
;					Width/Height values based on that sizing. Or when modifying a Frame, I obtain its anchor, and retrieve the
;					page style name, and get the page size settings for setting Frame Width/Height. However, is makes no
;					material difference, as the frame still is set to the correct width/height regardless.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ObjRelativeSize(ByRef $oDoc, ByRef $oObj, $bRelativeWidth = False, $bRelativeHeight = False)
	Local $iPageWidth, $iPageHeight, $iObjWidth, $iObjHeight
	Local $oPageStyle

	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", "__LOWriter_InternalComErrorHandler")
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bRelativeWidth) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bRelativeHeight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If (($bRelativeHeight = False) And ($bRelativeWidth = False)) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	If ($oObj.supportsService("com.sun.star.text.TextFrame")) Or ($oObj.supportsService("com.sun.star.text.TextGraphicObject")) Then
		$oPageStyle = $oDoc.StyleFamilies().getByName("PageStyles").getByName($oObj.Anchor.PageStyleName())
	Else
		$oPageStyle = $oDoc.StyleFamilies().getByName("PageStyles").getByName($oDoc.CurrentController.getViewCursor().PageStyleName())
	EndIf

	If Not IsObj($oPageStyle) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($bRelativeWidth = True) Then
		$iPageWidth = $oPageStyle.Width() ; Retrieve total PageStyle width
		$iPageWidth = $iPageWidth - $oPageStyle.RightMargin()
		$iPageWidth = $iPageWidth - $oPageStyle.LeftMargin() ; Minus off both margins.

		$iObjWidth = $iPageWidth * ($oObj.RelativeWidth() / 100) ; Times Page width minus margins by relative width percentage.

		$oObj.Width = $iObjWidth

	EndIf

	If ($bRelativeHeight = True) Then
		$iPageHeight = $oPageStyle.Height() ; Retrieve total PageStyle Height
		$iPageHeight = $iPageHeight - $oPageStyle.TopMargin()
		$iPageHeight = $iPageHeight - $oPageStyle.BottomMargin() ; Minus off both margins.

		$iObjHeight = $iPageHeight * ($oObj.RelativeHeight() / 100) ; Times Page Height minus margins by relative Height percentage.

		$oObj.Height = $iObjHeight
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_ObjRelativeSize

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_GetPrinterSetting
; Description ...: Internal function for retrieving Printer settings.
; Syntax ........: __LOWriter_GetPrinterSetting(Byref $oDoc, $sSetting)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sSetting            - a string value. The setting Name.
; Return values .: Success: Variable Value.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Array of Printer setting objects.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Requested setting not found.
;				   --Success--
;				   @Error 0 @Extended 0 Return Variable = Success. The requested setting's value.
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
	If Not IsArray($aoPrinterProperties) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	For $i = 0 To UBound($aoPrinterProperties) - 1
		If (($aoPrinterProperties[$i].Name()) = $sSetting) Then Return SetError($__LOW_STATUS_SUCCESS, 0, $aoPrinterProperties[$i].Value())
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
	Next

	Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0) ; No Matches
EndFunc   ;==>__LOWriter_GetPrinterSetting

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_GradientNameInsert
; Description ...: Create and insert a new Gradient name.
; Syntax ........: __LOWriter_GradientNameInsert(Byref $oDoc, $tGradient[, $sGradientName = "Gradient "])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $tGradient           - a dll struct value. A Gradient Structure to copy settings from.
;                  $sGradientName       - [optional] a string value. Default is "Gradient ". The Gradient name to create.
; Return values .: Success: String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $tGradient not an Object.
;				   @Error 1 @Extended 3 Return 0 = $sGradientName not a string.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.drawing.TransparencyGradientTable" Object.
;				   @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.awt.Gradient" structure.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error creating Gradient Name.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. A new Gradient name was created. Returning the new name as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If The Gradient name is blank, I need to create a new name and apply it. I think I could re-use
;					an old one without problems, but I'm not sure, so to be safe, I will create a new one. If there are no names
;					that have been already created, then I need to create and apply one before the transparency gradient will
;					be displayed. Else if a preset Gradient is called, I need to create its name before it can be used.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($tGradient) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sGradientName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	$oGradTable = $oDoc.createInstance("com.sun.star.drawing.GradientTable")
	If Not IsObj($oGradTable) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($sGradientName = "Gradient ") Then
		While $oGradTable.hasByName($sGradientName & $iCount)
			$iCount += 1
			Sleep((IsInt($iCount / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
		WEnd
		$sGradientName = $sGradientName & $iCount
	EndIf

	$tNewGradient = __LOWriter_CreateStruct("com.sun.star.awt.Gradient")
	If Not IsObj($tNewGradient) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

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
	EndWith

	If Not $oGradTable.hasByName($sGradientName) Then
		$oGradTable.insertByName($sGradientName, $tNewGradient)
		If Not ($oGradTable.hasByName($sGradientName)) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, $sGradientName)
EndFunc   ;==>__LOWriter_GradientNameInsert

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_GradientPresets
; Description ...: Set Page background Gradient to preset settings.
; Syntax ........: __LOWriter_GradientPresets(Byref $oDoc, Byref $oObject, Byref $tGradient, $sGradientName[, $bFooter = False[, $bHeader = False]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;				   $oObject             - [in/out] an object. The Object to modify the Gradient settings for.
;                  $tGradient           - [in/out] an object. The Fill Gradient Object to modify the Gradient settings for.
;                  $sGradientName       - a string value. The Gradient Preset name to apply.
;                  $bFooter             - [optional] a boolean value. Default is False. If True, settings are being set for footer Fill Gradient. If both are false, settings are for The Page itself.
;                  $bHeader             - [optional] a boolean value. Default is False. If True, settings are being set for Header Fill Gradient. If both are false, settings are for The Page itself.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to create Gradient name.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success. The Style Gradient settings were successfully set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_GradientPresets(ByRef $oDoc, ByRef $oObject, ByRef $tGradient, $sGradientName, $bFooter = False, $bHeader = False)

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
			EndWith

		Case $LOW_GRAD_NAME_BLANK_W_GRAY

			With $tGradient
				.Style = $LOW_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 900
				.Border = 75
				.StartColor = $LOW_COLOR_WHITE
				.EndColor = 14540253
				.StartIntensity = 100
				.EndIntensity = 100
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
			EndWith

		Case $LOW_GRAD_NAME_MIDNIGHT

			With $tGradient
				.Style = $LOW_GRAD_TYPE_LINEAR
				.StepCount = 0
				.XOffset = 0
				.YOffset = 0
				.Angle = 0
				.Border = 0
				.StartColor = $LOW_COLOR_BLACK
				.EndColor = 2777241
				.StartIntensity = 100
				.EndIntensity = 100
			EndWith

		Case $LOW_GRAD_NAME_DEEP_OCEAN

			With $tGradient
				.Style = $LOW_GRAD_TYPE_RADIAL
				.StepCount = 0
				.XOffset = 50
				.YOffset = 50
				.Angle = 0
				.Border = 0
				.StartColor = $LOW_COLOR_BLACK
				.EndColor = 7512015
				.StartIntensity = 100
				.EndIntensity = 100
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
				.EndColor = $LOW_COLOR_WHITE
				.StartIntensity = 100
				.EndIntensity = 100
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
			EndWith

		Case $LOW_GRAD_NAME_MAHOGANY

			With $tGradient
				.Style = $LOW_GRAD_TYPE_SQUARE
				.StepCount = 0
				.XOffset = 50
				.YOffset = 50
				.Angle = 450
				.Border = 0
				.StartColor = $LOW_COLOR_BLACK
				.EndColor = 9250846
				.StartIntensity = 100
				.EndIntensity = 100
			EndWith

		Case Else ;Custom Gradient Name
			__LOWriter_GradientNameInsert($oDoc, $tGradient, $sGradientName)
			If (@error > 0) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

			If $bFooter Then
				$oObject.FooterFillGradientName = $sGradientName

			ElseIf $bHeader Then
				$oObject.HeaderFillGradientName = $sGradientName

			Else
				$oObject.FillGradientName = $sGradientName
			EndIf
			Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
	EndSwitch

	__LOWriter_GradientNameInsert($oDoc, $tGradient, $sGradientName)
	If (@error > 0) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

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

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_GradientPresets

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_HeaderBorder
; Description ...: Header Border Setting Internal function.
; Syntax ........: __LOWriter_HeaderBorder(Byref $oObj, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
; Parameters ....: $oObj                - [in/out] an object. A Header object.
;                  $bWid                - a boolean value. If True the calling function is for setting Border Line Width.
;                  $bSty                - a boolean value. If True the calling function is for setting Border Line Style.
;                  $bCol                - a boolean value. If True the calling function is for setting Border Line Color.
;                  $iTop                - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iBottom             - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iLeft               - an integer value. See Border Style, Width, and Color functions for possible values.
;                  $iRight              - an integer value. See Border Style, Width, and Color functions for possible values.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oObj Variable not Object type variable.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style/Color when Border width not set.
;				   @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border style/Color when Border width not set.
;				   @Error 4 @Extended 3 Return 0 = Cannot set Left Border style/Color when Border width not set.
;				   @Error 4 @Extended 4 Return 0 = Cannot set Right Border style/Color when Border width not set.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with all other parameters set to Null keyword, and $bWid, or $bSty, or $bCol set to true to get the corresponding current settings.
;					All distance values are set in MicroMeters. Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_HeaderBorder(ByRef $oObj, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tBL2
	Local $aiBorder[4]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If (($bWid + $bSty + $bCol) <> 1) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0) ; If more than one Boolean is true = error

	If __LOWriter_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then
		If $bWid Then
			__LOWriter_ArrayFill($aiBorder, $oObj.HeaderTopBorder.LineWidth(), $oObj.HeaderBottomBorder.LineWidth(), _
					$oObj.HeaderLeftBorder.LineWidth(), $oObj.HeaderRightBorder.LineWidth())
		ElseIf $bSty Then
			__LOWriter_ArrayFill($aiBorder, $oObj.HeaderTopBorder.LineStyle(), $oObj.HeaderBottomBorder.LineStyle(), _
					$oObj.HeaderLeftBorder.LineStyle(), $oObj.HeaderRightBorder.LineStyle())
		ElseIf $bCol Then
			__LOWriter_ArrayFill($aiBorder, $oObj.HeaderTopBorder.Color(), $oObj.HeaderBottomBorder.Color(), $oObj.HeaderLeftBorder.Color(), _
					$oObj.HeaderRightBorder.Color())
		EndIf
		Return SetError($__LOW_STATUS_SUCCESS, 1, $aiBorder)
	EndIf

	$tBL2 = __LOWriter_CreateStruct("com.sun.star.table.BorderLine2")
	If Not IsObj($tBL2) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If $iTop <> Null Then
		If Not $bWid And ($oObj.HeaderTopBorder.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 1, 0) ; If Width not set, cant set color or style.
		; Top Line
		$tBL2.LineWidth = ($bWid) ? $iTop : $oObj.HeaderTopBorder.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iTop : $oObj.HeaderTopBorder.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iTop : $oObj.HeaderTopBorder.Color() ; copy Color over to new size structure
		$oObj.HeaderTopBorder = $tBL2
	EndIf

	If $iBottom <> Null Then
		If Not $bWid And ($oObj.HeaderBottomBorder.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 2, 0) ; If Width not set, cant set color or style.
		; Bottom Line
		$tBL2.LineWidth = ($bWid) ? $iBottom : $oObj.HeaderBottomBorder.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iBottom : $oObj.HeaderBottomBorder.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iBottom : $oObj.HeaderBottomBorder.Color() ; copy Color over to new size structure
		$oObj.HeaderBottomBorder = $tBL2
	EndIf

	If $iLeft <> Null Then
		If Not $bWid And ($oObj.HeaderLeftBorder.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 3, 0) ; If Width not set, cant set color or style.
		; Left Line
		$tBL2.LineWidth = ($bWid) ? $iLeft : $oObj.HeaderLeftBorder.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iLeft : $oObj.HeaderLeftBorder.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iLeft : $oObj.HeaderLeftBorder.Color() ; copy Color over to new size structure
		$oObj.HeaderLeftBorder = $tBL2
	EndIf

	If $iRight <> Null Then
		If Not $bWid And ($oObj.HeaderRightBorder.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 4, 0) ; If Width not set, cant set color or style.
		; Right Line
		$tBL2.LineWidth = ($bWid) ? $iRight : $oObj.HeaderRightBorder.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iRight : $oObj.HeaderRightBorder.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iRight : $oObj.HeaderRightBorder.Color() ; copy Color over to new size structure
		$oObj.HeaderRightBorder = $tBL2
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_HeaderBorder

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ImageGetSuggestedSize
; Description ...: Return a suggested image width/height based on an image's original size.
; Syntax ........: __LOWriter_ImageGetSuggestedSize(ByRef $oGraphic, $oPageStyle)
; Parameters ....: $oGraphic            - [in/out] an object. A graphic Object returned from a queryGraphicDescriptor call.
;                  $oPageStyle          - an object. A Page Style object returned by a previous _LOWriter_PageStyleGetObj function.
; Return values .: Success: Structure.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oGraphic not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oPageStyle not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error calculating Width and Height.
;				   --Success--
;				   @Error 0 @Extended 0 Return Structure. = Successfully calculated suggested Width and Height, returning size Structure.
; Author ........: Andrew Pitonyak ("Useful Macro Information For OpenOffice.org", Page 62, listing 5.28)
; Modified ......: donnyh13, converted code from L.O. Basic to AutoIt. Added a max W/H based on current page size.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ImageGetSuggestedSize($oGraphic, $oPageStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", "__LOWriter_InternalComErrorHandler")
	#forceref $oCOM_ErrorHandler

	Local $oSize
	Local $iMaxH, $iMaxW

	If Not IsObj($oGraphic) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPageStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	;Retrieve the Current PageStyle's height minus top/bottom margins
	$iMaxH = Int($oPageStyle.Height() - $oPageStyle.LeftMargin() - $oPageStyle.RightMargin())
	If ($iMaxH = 0) Then $iMaxH = 9.5 * 2540 ; If error or is equal to 0, then set to 9.5 Inches in Micrometers

	;Retrieve the Current PageStyle's width minus left/right margins
	$iMaxW = Int($oPageStyle.Width() - $oPageStyle.TopMargin() - $oPageStyle.BottomMargin())
	If ($iMaxW = 0) Then $iMaxW = 6.75 * 2540 ; If error or is equal to 0, then set to 6.75 Inches in Micrometers.

	$oSize = $oGraphic.Size100thMM()

	If ($oSize.Height = 0) Or ($oSize.Width = 0) Then
		;2540 Micrometers per Inch, 1440 TWIPS per inch
		$oSize.Height = $oGraphic.SizePixel.Height * 2540 * _WinAPI_TwipsPerPixelY() / 1440
		$oSize.Width = $oGraphic.SizePixel.Width * 2540 * _WinAPI_TwipsPerPixelX() / 1440
	EndIf

	If ($oSize.Height = 0) Or ($oSize.Width = 0) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

	If ($oSize.Width() > $iMaxW) Then
		$oSize.Height = $oSize.Height * $iMaxW / $oSize.Width()
		$oSize.Width = $iMaxW
	EndIf

	If ($oSize.Height() > $iMaxH) Then
		$oSize.Width = $oSize.Width() * $iMaxH / $oSize.Height
		$oSize.Height = $iMaxH
	EndIf

	Return $oSize
EndFunc   ;==>__LOWriter_ImageGetSuggestedSize

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Internal_CursorGetDataType
; Description ...: Get what type of Text data the cursor object is currently in. Internal version of CursorGetType.
; Syntax ........: __LOWriter_Internal_CursorGetDataType(Byref $oDoc, Byref $oCursor[, $ReturnObject = False])
; Parameters ....: $oDoc                 - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor              - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions.
;                  $bReturnObject        - [optional] a boolean value. Default is False. Whether to return an object used for creating a Text Object etc.
; Return values .:Success: Object or Integer.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $ReturnObject not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $oCursor is a Table Cursor, or a View Cursor with table cells selected. Can't get data type from a Table Cursor.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Footnotes Object for document.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Endnotes Object for document.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error retrieving TextFrame Object.
;				   @Error 3 @Extended 2 Return 0 = Error retrieving TextCell Object.
;				   @Error 3 @Extended 3 Return 0 = Unable to identify Foot/EndNote.
;				   @Error 3 @Extended 4 Return 0 = Cursor in unknown DataType
;				   --Success--
;				   @Error 0 @Extended Integer Return Object = Success, If $bReturnObject is True, returns an object used for creating a Text Object, @Extended is set to one of the constants, $LOW_CURDATA_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 0 @Extended 0 Return Integer  = Success, If $bReturnObject is False, Return value will be one of constants, $LOW_CURDATA_* as defined in LibreOfficeWriter_Constants.au3.
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

	Local $oEndNotes, $oFootNotes, $oFootEndNote, $oReturnObject
	Local $iLWFootEndNote = 0
	Local $bFound = False
	Local $sNoteRefID

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bReturnObject) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If (($oCursor.ImplementationName()) = "SwXTextTableCursor") Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0) ; Can't get data type from Table Cursor.

	Switch $oCursor.Text.getImplementationName()
		Case "SwXBodyText"
			$oReturnObject = $oDoc
			Return ($bReturnObject) ? SetError($__LOW_STATUS_SUCCESS, $LOW_CURDATA_BODY_TEXT, $oReturnObject) : SetError($__LOW_STATUS_SUCCESS, 0, $LOW_CURDATA_BODY_TEXT)
		Case "SwXTextFrame"
			$oReturnObject = $oDoc.TextFrames.getByName($oCursor.TextFrame.Name)
			If Not IsObj($oReturnObject) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
			Return ($bReturnObject) ? SetError($__LOW_STATUS_SUCCESS, $LOW_CURDATA_FRAME, $oReturnObject) : SetError($__LOW_STATUS_SUCCESS, 0, $LOW_CURDATA_FRAME)
		Case "SwXCell"
			$oReturnObject = $oDoc.TextTables.getByName($oCursor.TextTable.Name)
			If Not IsObj($oReturnObject) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0)
			Return ($bReturnObject) ? SetError($__LOW_STATUS_SUCCESS, $LOW_CURDATA_CELL, $oReturnObject) : SetError($__LOW_STATUS_SUCCESS, 0, $LOW_CURDATA_CELL)

		Case "SwXHeadFootText"
			$oReturnObject = $oCursor
			Return ($bReturnObject) ? SetError($__LOW_STATUS_SUCCESS, $LOW_CURDATA_HEADER_FOOTER, $oReturnObject) : SetError($__LOW_STATUS_SUCCESS, 0, $LOW_CURDATA_HEADER_FOOTER)

		Case "SwXFootnote"
			$oFootNotes = $oDoc.getFootnotes()
			If Not IsObj($oFootNotes) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
			$sNoteRefID = $oCursor.Text.ReferenceId()

			If $oFootNotes.hasElements() Then
				For $i = 0 To $oFootNotes.getCount() - 1
					$oFootEndNote = $oFootNotes.getByIndex($i)

					If ($oFootEndNote.ReferenceId() = $sNoteRefID) Then

						If ($oFootEndNote.Anchor.String() = $oCursor.Text.Anchor.String()) And _
								(($oDoc.Text.compareRegionEnds($oCursor.Text.Anchor, $oFootEndNote.Text.Anchor)) = 0) Then
							$bFound = True
							$iLWFootEndNote = $LOW_CURDATA_FOOTNOTE
							ExitLoop
						EndIf
					EndIf
				Next
			EndIf

			If ($bFound = False) Then
				$oEndNotes = $oDoc.getEndnotes()
				If Not IsObj($oEndNotes) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

				If $oEndNotes.hasElements() Then
					For $i = 0 To $oEndNotes.getCount() - 1
						$oFootEndNote = $oEndNotes.getByIndex($i)

						If ($oFootEndNote.ReferenceId() = $sNoteRefID) Then

							If ($oFootEndNote.Anchor.String() = $oCursor.Text.Anchor.String()) And _
									(($oDoc.Text.compareRegionEnds($oCursor.Text.Anchor, $oFootEndNote.Text.Anchor)) = 0) Then
								$bFound = True
								$iLWFootEndNote = $LOW_CURDATA_ENDNOTE
								ExitLoop
							EndIf
						EndIf
					Next
				EndIf
			EndIf

			If ($bFound = True) And ($iLWFootEndNote <> 0) Then
				$oReturnObject = $oFootEndNote
				Return ($bReturnObject) ? SetError($__LOW_STATUS_SUCCESS, $iLWFootEndNote, $oReturnObject) : SetError($__LOW_STATUS_SUCCESS, 0, $iLWFootEndNote)
			EndIf
			Return SetError($__LOW_STATUS_PROCESSING_ERROR, 3, 0) ; no matches
		Case Else
			Return SetError($__LOW_STATUS_PROCESSING_ERROR, 4, 0) ; unknown data type.
	EndSwitch
EndFunc   ;==>__LOWriter_Internal_CursorGetDataType

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_Internal_CursorGetType
; Description ...: Get what type of Text data the cursor object is currently in. Internal version of CursorGetType.
; Syntax ........: __LOWriter_Internal_CursorGetType(Byref $oCursor)
; Parameters ....: $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions.
; Return values .:Success: Integer.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCursor variable not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Unknown Cursor type.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer  = Success, Return value will be one of the constants, $LOW_CURTYPE_* as defined in LibreOfficeWriter_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Returns what type of cursor the input Object is, such as a TableCursor, Text Cursor or a ViewCursor.
;					Can also be a Paragraph or Text Portion.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_Internal_CursorGetType(ByRef $oCursor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	Switch $oCursor.getImplementationName()
		Case "SwXTextViewCursor"
			Return SetError($__LOW_STATUS_SUCCESS, 0, $LOW_CURTYPE_VIEW_CURSOR)
		Case "SwXTextTableCursor"
			Return SetError($__LOW_STATUS_SUCCESS, 0, $LOW_CURTYPE_TABLE_CURSOR)
		Case "SwXTextCursor"
			Return SetError($__LOW_STATUS_SUCCESS, 0, $LOW_CURTYPE_TEXT_CURSOR)
		Case "SwXParagraph"
			Return SetError($__LOW_STATUS_SUCCESS, 0, $LOW_CURTYPE_PARAGRAPH)
		Case "SwXTextPortion"
			Return SetError($__LOW_STATUS_SUCCESS, 0, $LOW_CURTYPE_TEXT_PORTION)
		Case Else
			Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0) ; unknown Cursor type.
	EndSwitch
EndFunc   ;==>__LOWriter_Internal_CursorGetType

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_InternalComErrorHandler
; Description ...: ComError Handler
; Syntax ........: __LOWriter_InternalComErrorHandler(Byref $oComError)
; Parameters ....: $oComError           - [in/out] an object. The Com Error Object passed by Autoit.Error.
; Return values .: None
; Author ........: mLipok
; Modified ......: donnyh13 - Added parameters option. Also added MsgBox & ConsoleWrite option.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_InternalComErrorHandler(ByRef $oComError)
	; If not defined ComError_UserFunction then this function does nothing.
	; In that case you can only check @error / @extended after suspect functions.
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
; Name ..........: __LOWriter_IntIsBetween
; Description ...: Test whether an input is an Integer and is between two Numbers.
; Syntax ........: __LOWriter_IntIsBetween($iTest, $nMin, $nMax[, $snNot = ""[, $snIncl = Default]])
; Parameters ....: $iTest               - an integer value. The Value to test.
;                  $nMin                - a general number value. The minimum $iTest can be.
;                  $nMax                - a general number value. The maximum $iTest can be.
;                  $snNot               - [optional] a string value. Default is "". Can be a single number, or a String of numbers separated by ":". Defines numbers inside the min/max range that are not allowed.
;                  $snIncl              - [optional] a string value. Default is Default. Can be a single number, or a String of numbers separated by ":". Defines numbers Outside the min/max range that are allowed.
; Return values .: Success: Boolean
;				   Failure: False
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = If the input is between Min and Max or is an allowed number, and not one of the disallowed numbers, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_IntIsBetween($iTest, $nMin, $nMax, $snNot = "", $snIncl = Default)
	Local $bMatch = False
	Local $anNot, $anIncl

	If Not IsInt($iTest) Then Return False
	If (@NumParams = 3) Then Return (($iTest < $nMin) Or ($iTest > $nMax)) ? False : True

	If ($snNot <> "") Then
		If IsString($snNot) And StringInStr($snNot, ":") Then
			$anNot = StringSplit($snNot, ":")
			For $i = 1 To $anNot[0]
				If ($anNot[$i] = $iTest) Then Return False
			Next
		Else
			If ($iTest = $snNot) Then Return False
		EndIf
	EndIf

	If (($iTest >= $nMin) And ($iTest <= $nMax)) Then Return True

	If IsString($snIncl) And StringInStr($snIncl, ":") Then
		$anIncl = StringSplit($snIncl, ":")
		For $j = 1 To $anIncl[0]
			$bMatch = ($anIncl[$j] = $iTest) ? True : False
			If $bMatch Then ExitLoop
		Next
	ElseIf IsNumber($snIncl) Then
		$bMatch = ($iTest = $snIncl) ? True : False
	EndIf

	Return $bMatch
EndFunc   ;==>__LOWriter_IntIsBetween

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_IsCellRange
; Description ...: Check whether a Cell Object is a single cell or a CellRange.
; Syntax ........: __LOWriter_IsCellRange(Byref $oCell)
; Parameters ....: $oCell               - [in/out] an object. A Table Cell or Cell Range Object returned from _LOWriter_TableGetCellObjByCursor, _LOWriter_TableGetCellObjByName, or
; Return values .: Success: Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oTable variable not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean: If the cell object is a Cell Range, True is returned. Else False.
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

	If Not IsObj($oCell) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	Return ($oCell.supportsService("com.sun.star.text.CellRange")) ? SetError($__LOW_STATUS_SUCCESS, 0, True) : SetError($__LOW_STATUS_SUCCESS, 0, False)
EndFunc   ;==>__LOWriter_IsCellRange

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_IsTableInDoc
; Description ...: Check if Table is inserted in a Document or has only been created and not inserted.
; Syntax ........: __LOWriter_IsTableInDoc(Byref $oTable)
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned from _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName functions.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oTable Variable not Object type variable.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error retrieving Table cell names.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = If True, Table is inserted into the document, If false Table has been created with TableCreate but not inserted.
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

	If Not IsObj($oTable) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	$aTableNames = $oTable.getCellNames()
	If Not IsArray($aTableNames) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

	Return (UBound($aTableNames)) ? True : False ; If 0 elements = False = not in doc.
EndFunc   ;==>__LOWriter_IsTableInDoc

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_NumIsBetween
; Description ...: Test whether an input is a Number and is between two Numbers.
; Syntax ........: __LOWriter_NumIsBetween($nTest, $nMin, $nMax[, $snNot = ""[, $snIncl = Default]])
; Parameters ....: $nTest               - a general number value. The Value to test.
;                  $nMin                - a general number value. The minimum $iTest can be.
;                  $nMax                - a general number value. The maximum $iTest can be.
;                  $snNot               - [optional] a string value. Default is "". Can be a single number, or a String of numbers separated by ":". Defines numbers inside the min/max range that are not allowed.
;                  $snIncl              - [optional] a string value. Default is Default. Can be a single number, or a String of numbers separated by ":". Defines numbers Outside the min/max range that are allowed.
; Return values .: Success: Boolean
;				   Failure: False
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = If the input is between Min and Max or is an allowed number, and not one of the disallowed numbers, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_NumIsBetween($nTest, $nMin, $nMax, $snNot = "", $snIncl = Default)
	Local $bMatch = False
	Local $anNot, $anIncl

	If Not IsNumber($nTest) Then Return SetError($__LOW_STATUS_SUCCESS, 0, False)
	If (@NumParams = 3) Then Return (($nTest < $nMin) Or ($nTest > $nMax)) ? SetError($__LOW_STATUS_SUCCESS, 0, False) : SetError($__LOW_STATUS_SUCCESS, 0, True)

	If ($snNot <> "") Then
		If IsString($snNot) And StringInStr($snNot, ":") Then
			$anNot = StringSplit($snNot, ":")
			For $i = 1 To $anNot[0]
				If ($anNot[$i] = $nTest) Then Return SetError($__LOW_STATUS_SUCCESS, 0, False)
			Next
		Else
			If ($nTest = $snNot) Then Return SetError($__LOW_STATUS_SUCCESS, 0, False)
		EndIf
	EndIf

	If (($nTest >= $nMin) And ($nTest <= $nMax)) Then Return SetError($__LOW_STATUS_SUCCESS, 0, True)

	If IsString($snIncl) And StringInStr($snIncl, ":") Then
		$anIncl = StringSplit($snIncl, ":")
		For $j = 1 To $anIncl[0]
			$bMatch = ($anIncl[$j] = $nTest) ? True : False
			If $bMatch Then ExitLoop
		Next
	ElseIf IsNumber($snIncl) Then
		$bMatch = ($nTest = $snIncl) ? True : False
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, $bMatch)
EndFunc   ;==>__LOWriter_NumIsBetween

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_NumStyleCreateScript
; Description ...: Part of the Numbering Style Modification workaround, creates a Macro in a document.
; Syntax ........: __LOWriter_NumStyleCreateScript(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Standard Macro Library.
;				   @Error 2 @Extended 2 Return 0 = Document already contains the Macro.
;				   @Error 2 @Extended 3 Return 0 = Error creating Macro in Document.
;				   @Error 2 @Extended 4 Return 0 = Error retrieving Script Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Function successfully created the Macro in Document. Returning Script Object.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	; Retrieving the BasicLibrary.Standard Object fails when using a newly opened document, I found a workaround by updating the
	; following setting.
	$oDoc.BasicLibraries.VBACompatibilityMode = $oDoc.BasicLibraries.VBACompatibilityMode()

	$oStandardLibrary = $oDoc.BasicLibraries.Standard()
	If Not IsObj($oStandardLibrary) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	If $oStandardLibrary.hasByName("NumStyleModifierDH13") Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	$oStandardLibrary.insertByName("NumStyleModifierDH13", $sNumStyleScript)
	If Not $oStandardLibrary.hasByName("NumStyleModifierDH13") Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)

	$oScript = $oDoc.getScriptProvider().getScript("vnd.sun.star.script:Standard.NumStyleModifierDH13.ReplaceByIndex?language=Basic&location=document")
	If Not IsObj($oScript) Then Return SetError($__LOW_STATUS_INIT_ERROR, 4, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, $oScript)
EndFunc   ;==>__LOWriter_NumStyleCreateScript

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_NumStyleDeleteScript
; Description ...: Part of the Numbering Style Modification workaround, deletes a Macro in a document.
; Syntax ........: __LOWriter_NumStyleDeleteScript(Byref $oDoc)
; Parameters ....: $oDoc                - [in/out] an object.  A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Standard Macro Library.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error deleting Macro.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Function successfully deleted the Macro in Document.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	; Retrieving the BasicLibrary.Standard Object fails when using a newly opened document, I found a workaround by updating the
	; following setting.
	$oDoc.BasicLibraries.VBACompatibilityMode = $oDoc.BasicLibraries.VBACompatibilityMode()

	$oStandardLibrary = $oDoc.BasicLibraries.Standard()
	If Not IsObj($oStandardLibrary) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	If $oStandardLibrary.hasByName("NumStyleModifierDH13") Then $oStandardLibrary.removeByName("NumStyleModifierDH13")

	If $oStandardLibrary.hasByName("NumStyleModifierDH13") Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_NumStyleDeleteScript

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_NumStyleInitiateDocument
; Description ...: Part of the work around method for modifying Numbering Style settings.
; Syntax ........: __LOWriter_NumStyleInitiateDocument()
; Parameters ....: None
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.ServiceManager" Object.
;				   @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.frame.Desktop" Object.
;				   @Error 2 @Extended 3 Return 0 = Error Creating document.
;				   @Error 2 @Extended 4 Return 0 = Error retrieving standard Macro Library Object from Document.
;				   @Error 2 @Extended 5 Return 0 = Error creating NumStyleModifierDH13 Module in document.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting Hidden
;				   |								2 = Error setting MacroExecutionMode
;				   |								4 = Error setting ReadOnly
;				   --Success--
;				   @Error 0 @Extended 0 Return Doc Object. = Success. The Numbering Style Modification Document was successfully created.
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

	Local Const $iMacroExecMode_ALWAYS_EXECUTE_NO_WARN = 4, $iURLFrameCreate = 8 ;frame will be created if not found
	Local $iError = 0
	Local $oNumStyleDoc, $oServiceManager, $oDesktop
	Local $atProperties[3]
	Local $vProperty

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	$vProperty = __LOWriter_SetPropertyValue("Hidden", True)
	If @error Then $iError = BitOR($iError, 1)
	If Not BitAND($iError, 1) Then $atProperties[0] = $vProperty

	$vProperty = __LOWriter_SetPropertyValue("MacroExecutionMode", $iMacroExecMode_ALWAYS_EXECUTE_NO_WARN)
	If @error Then $iError = BitOR($iError, 2)
	If Not BitAND($iError, 2) Then $atProperties[1] = $vProperty

	$vProperty = __LOWriter_SetPropertyValue("ReadOnly", True)
	If @error Then $iError = BitOR($iError, 4)
	If Not BitAND($iError, 4) Then $atProperties[2] = $vProperty

	$oNumStyleDoc = $oDesktop.loadComponentFromURL("private:factory/swriter", "_blank", $iURLFrameCreate, $atProperties)
	If Not IsObj($oNumStyleDoc) Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)

	__LOWriter_NumStyleCreateScript($oNumStyleDoc)
	If (@error > 0) Then Return SetError($__LOW_STATUS_INIT_ERROR, 5, 0)

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, $oNumStyleDoc) : SetError($__LOW_STATUS_SUCCESS, 1, $oNumStyleDoc)
EndFunc   ;==>__LOWriter_NumStyleInitiateDocument

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_NumStyleListFormat
; Description ...: Creates a string or Array for modifying List Format Number Style setting.
; Syntax ........: __LOWriter_NumStyleListFormat(Byref $oNumRules, $iLevel, $iSubLevels[, $sPrefix = Null[, $sSuffix = Null]])
; Parameters ....: $oNumRules           - [in/out] an object. The Numbering Rules object retrieved from a Numbering Style.
;                  $iLevel              - an integer value. The Level to create the ListFormat string for
;                  $iSubLevels          - an integer value. The number of levels to go up from $iLevel.
;                  $sPrefix             - [optional] a string value. Default is Null. If Null, retrieves the current Prefix, else use the input prefix.
;                  $sSuffix             - [optional] a string value. Default is Null. If Null, retrieves the current Suffix, else use the input Suffix.
; Return values .: Success: Array or String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oNumRules not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iLevel not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iSubLevels not an Integer.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. A String used for modifying ListFormat Numbering Style setting.
;				   @Error 0 @Extended 1 Return Array = Success. An Array of List format strings used for modifying all levels of ListFormat Numbering Style setting.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_NumStyleListFormat(ByRef $oNumRules, $iLevel, $iSubLevels, $sPrefix = Null, $sSuffix = Null)
	Local $sListFormat = "", $sSeperator = "."
	Local $iBeginLevel, $iEndLevel
	Local $aListFormats[10]

	If Not IsObj($oNumRules) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iLevel) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iSubLevels) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	$iBeginLevel = ($iLevel = -1) ? 9 : $iLevel ; If Level = -1 (all levels) then set the begin level at 9 (the last level), else at the called level.
	$iEndLevel = ($iLevel = -1) ? 0 : ($iLevel - $iSubLevels + 1) ; If Level = -1 (all levels) then set the end level at 0 (the first level), else
	; at the called level - any Sublevels.

	If ($iLevel = -1) Then ;  If Level = -1 (all levels) cycle through them all, Applying their respective Prefix/Suffix
		For $i = $iBeginLevel To $iEndLevel Step -1
			$sPrefix = ($sPrefix = Null) ? __LOWriter_NumStyleRetrieve($oNumRules, $i, "Prefix") : $sPrefix
			$sSuffix = ($sSuffix = Null) ? __LOWriter_NumStyleRetrieve($oNumRules, $i, "Suffix") : $sSuffix
			$aListFormats[$i] = $sPrefix & "%" & ($i + 1) & "%" & $sSuffix
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next

	Else ;Else if I'm modifying a specific level, retrieve its prefix/Suffix
		$sPrefix = ($sPrefix = Null) ? __LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "Prefix") : $sPrefix
		$sSuffix = ($sSuffix = Null) ? __LOWriter_NumStyleRetrieve($oNumRules, $iLevel, "Suffix") : $sSuffix

		For $i = $iBeginLevel To $iEndLevel Step -1 ;Cycle Through the levels if any Sub levels are set.
			If ($i = $iEndLevel) Then $sSeperator = ""
			$sListFormat = $sSeperator & "%" & ($i + 1) & "%" & $sListFormat
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next
		$sListFormat = $sPrefix & $sListFormat & $sSuffix

	EndIf

	Return ($iLevel = -1) ? SetError($__LOW_STATUS_SUCCESS, 0, $aListFormats) : SetError($__LOW_STATUS_SUCCESS, 1, $sListFormat)

EndFunc   ;==>__LOWriter_NumStyleListFormat

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_NumStyleModify
; Description ...: Internal function for modifying Numbering Style settings.
; Syntax ........: __LOWriter_NumStyleModify(Byref $oDoc, Byref $oNumRules, $iLevel, $avSettings)
; Parameters ....: $oDoc                - [in/out] an object.  A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function, to modify NumberingRules for.
;                  $oNumRules           - [in/out] an object. The Numbering Rules object retrieved from a Numbering Style.
;                  $iLevel              - an integer value. The Numbering Style level to modify. 0-9 or -1 for all.
;                  $avSettings          - an array of variants. Array containing Numbering Style settings to set.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oNumRules not an Object.
;				   @Error 1 @Extended 3 Return 0 = $iLevel not between -1 and 9 to indicate correct level.
;				   @Error 1 @Extended 4 Return 0 = $avSettings not an array.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error opening new document, and inserting ReplaceByIndex Script.
;				   @Error 2 @Extended 2 Return 0 = Error retrieving "Standard.NumStyleModifierDH13.ReplaceByIndex" Macro in new document.
;				   @Error 2 @Extended 3 Return 0 = Error retrieving Numbering Rules level.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error deleting ReplaceByIndex Macro from Document.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully set the requested settings.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This works, but only with a work-around method, see inside this function for a description of why a work-around method is necessary.
;					When a lot of settings are set, especially for all levels, this function can be a bit slow.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_NumStyleModify(ByRef $oDoc, ByRef $oNumRules, $iLevel, $avSettings)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atNumLevel
	Local $iGetLevel, $iParentNumber, $iListFormatIndex, $iEndLevel
	Local $oNumStyleDoc, $oScript
	Local $aDummyArray[0], $avParamArray[3]
	Local $sSettingName, $sListFormat
	Local $vSettingValue
	Local $bNumDocOpen = False

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oNumRules) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not __LOWriter_IntIsBetween($iLevel, -1, 9) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsArray($avSettings) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)

	$iEndLevel = ($iLevel = -1) ? 9 : 0 ;  replace setting value for all levels at once or for one level

	$oScript = __LOWriter_NumStyleCreateScript($oDoc) ; Create my modification Script.

	If Not IsObj($oScript) Then ; If creating my Mod. Script fails, open a new document and create a script in there.
		$oNumStyleDoc = __LOWriter_NumStyleInitiateDocument()
		If Not IsObj($oNumStyleDoc) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
		$oScript = $oNumStyleDoc.getScriptProvider().getScript("vnd.sun.star.script:Standard.NumStyleModifierDH13.ReplaceByIndex?language=Basic&location=document")
		If Not IsObj($oScript) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)
		$bNumDocOpen = True
	EndIf

	For $k = 0 To UBound($avSettings) - 1 ;Cycle through the settings to be set.
		$sSettingName = $avSettings[$k][0]
		$vSettingValue = $avSettings[$k][1]

		$iGetLevel = ($iLevel = -1) ? 0 : $iLevel

		For $j = 0 To $iEndLevel ;Set the settings for each level.
			$atNumLevel = $oNumRules.getByIndex($iGetLevel)
			If Not IsArray($atNumLevel) Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)

			For $i = 0 To UBound($atNumLevel) - 1 ;Cycle through Array of Num Style settings and modify the settings.
				If ($atNumLevel[$i].Name() = $sSettingName) Then

					Switch $sSettingName
						Case "ListFormat"
							$atNumLevel[$i].Value = ($iLevel = -1) ? $vSettingValue[$j] : $vSettingValue

						Case "Prefix"
							__LOWriter_NumStyleRetrieve($oNumRules, $iGetLevel, "ListFormat") ; Test is ListFormat exists, if so modify prefix/Suffix via new method.
							If (@error = 0) Then
								$iListFormatIndex = @extended ;Record Index location of "ListFormat"
								; Get number of Sublevels (ParentNumbering) for formatting ListFormat.
								$iParentNumber = __LOWriter_NumStyleRetrieve($oNumRules, $iGetLevel, "ParentNumbering")
								$sListFormat = __LOWriter_NumStyleListFormat($oNumRules, $iGetLevel, $iParentNumber, $vSettingValue, Null) ; Add prefix to ListFormat
								$atNumLevel[$iListFormatIndex].Value = $sListFormat
							EndIf
							$atNumLevel[$i].Value = $vSettingValue ;Set Literal Prefix Value, which wont work if "ListFormat" is present.

						Case "Suffix"
							__LOWriter_NumStyleRetrieve($oNumRules, $iGetLevel, "ListFormat") ; Test is ListFormat exists, if so modify prefix/Suffix via new method.
							If (@error = 0) Then
								$iListFormatIndex = @extended ;Record Index location of "ListFormat"
								; Get number of Sublevels (ParentNumbering) for formatting ListFormat.
								$iParentNumber = __LOWriter_NumStyleRetrieve($oNumRules, $iGetLevel, "ParentNumbering")
								$sListFormat = __LOWriter_NumStyleListFormat($oNumRules, $iGetLevel, $iParentNumber, Null, $vSettingValue) ; Add suffix to ListFormat
								$atNumLevel[$iListFormatIndex].Value = $sListFormat
							EndIf
							$atNumLevel[$i].Value = $vSettingValue ;Set Literal Suffix Value, which wont work if "ListFormat" is present.

						Case Else
							$atNumLevel[$i].Value = $vSettingValue
					EndSwitch

					; $oNumRules.replaceByIndex($iGetLevel, $atNumLevel);This should work but doesn't -- It would seem that the Array passed by
					; Autoit is not recognized as an appropriate array(or Sequence) by LibreOffice, or perhaps as variable type "Any", which is
					; what LibreOfficereplace by index is expecting, and consequently causes a com.sun.star.lang.IllegalArgumentException COM error.

					$avParamArray[0] = $oNumRules
					$avParamArray[1] = $iGetLevel
					$avParamArray[2] = $atNumLevel

					$oNumRules = $oScript.Invoke($avParamArray, $aDummyArray, $aDummyArray)
					ExitLoop
				EndIf

				Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
			Next

			$iGetLevel += 1

		Next

	Next

	If ($bNumDocOpen = True) Then
		$oNumStyleDoc.Close(True)
	Else
		__LOWriter_NumStyleDeleteScript($oDoc)
		If @error Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
	EndIf

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_NumStyleModify

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_NumStyleRetrieve
; Description ...: Internal function for Retrieving Numbering Style settings.
; Syntax ........: __LOWriter_NumStyleRetrieve(Byref $oNumRules, $iLevel, $sSettingName)
; Parameters ....: $oNumRules           - [in/out] an object. The Numbering Rules object retrieved from a Numbering Style.
;                  $iLevel              - an integer value. The Numbering Style level to modify. 0-9.
;                  $sSettingName        - a string value. The Numbering Style Setting name to modify.
; Return values .: Success: Variable
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oNumRules not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iLevel not between 0 and 9 to indicate correct level.
;				   @Error 1 @Extended 3 Return 0 = $sSettingName not a String.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error retrieving specified Numbering Level.
;				   @Error 3 @Extended 2 Return 0 = Requested setting not found.
;				   --Success--
;				   @Error 0 @Extended ? Return Variable = Success. Successfully retrieved requested Setting Value, returning value in Return value, and index location in @Extended.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_NumStyleRetrieve(ByRef $oNumRules, $iLevel, $sSettingName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atNumLevel
	Local $iGetLevel = $iLevel

	If Not IsObj($oNumRules) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_IntIsBetween($iLevel, 0, 9) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sSettingName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	$atNumLevel = $oNumRules.getByIndex($iGetLevel)
	If Not IsArray($atNumLevel) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

	For $i = 0 To UBound($atNumLevel) - 1
		If ($atNumLevel[$i].Name() = $sSettingName) Then Return SetError($__LOW_STATUS_SUCCESS, $i, $atNumLevel[$i].Value())
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
	Next
	Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0)
EndFunc   ;==>__LOWriter_NumStyleRetrieve

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_PageStyleNameToggle
; Description ...: Toggle from Page Style Display Name to Internal Name for error checking and setting retrieval.
; Syntax ........: __LOWriter_PageStyleNameToggle(Byref $sPageStyle[, $bReverse = False])
; Parameters ....: $sPageStyle          - a string value. The PageStyle Name to Toggle.
;                  $bReverse            - [optional] a boolean value. Default is False. If True Reverse toggles the Page Style Name.
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $sPageStyle not a String.
;				   @Error 1 @Extended 2 Return 0 = $bReverse not a Boolean.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. Page Style Name successfully toggled. Returning changed name as a string.
;				   @Error 0 @Extended 1 Return String = Success. Page Style Name successfully reverse toggled. Returning changed name as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_PageStyleNameToggle($sPageStyle, $bReverse = False)
	If Not IsString($sPageStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReverse) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If ($bReverse = False) Then
		$sPageStyle = ($sPageStyle = "Default Page Style") ? "Standard" : $sPageStyle
		Return SetError($__LOW_STATUS_SUCCESS, 0, $sPageStyle)
	Else
		$sPageStyle = ($sPageStyle = "Standard") ? "Default Page Style" : $sPageStyle
		Return SetError($__LOW_STATUS_SUCCESS, 1, $sPageStyle)
	EndIf
EndFunc   ;==>__LOWriter_PageStyleNameToggle

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParAlignment
; Description ...: Set and Retrieve Alignment settings.
; Syntax ........: __LOWriter_ParAlignment(Byref $oObj, $iHorAlign, $iVertAlign, $iLastLineAlign, $bExpandSingleWord, $bSnapToGrid, $iTxtDirection)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iHorAlign           - an integer value (0-3). The Horizontal alignment of the paragraph. See Constants, $LOW_PAR_ALIGN_HOR_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $iVertAlign          - an integer value (0-4). The Vertical alignment of the paragraph. See Constants, $LOW_PAR_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLastLineAlign      - an integer value (0-3). Specify the alignment for the last line in the paragraph. See Constants, $LOW_PAR_LAST_LINE_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $bExpandSingleWord   - a boolean value. If the last line of a justified paragraph consists of one word, the word is stretched to the width of the paragraph.
;                  $bSnapToGrid         - a boolean value. If True, Aligns the paragraph to a text grid (if one is active).
;                  $iTxtDirection       - an integer value (0-5). The Text Writing Direction. See Constants, $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3. [Libre Office Default is 4]
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iHorAlign not an integer, less than 0 or greater than 3. See Constants, $LOW_PAR_ALIGN_HOR_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 5 Return 0 = $iVertAlign not an integer, less than 0 or more than 4. See Constants, $LOW_PAR_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 6 Return 0 = $iLastLineAlign not an integer, less than 0 or more than 3. See Constants, $LOW_PAR_LAST_LINE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 7 Return 0 = $bExpandSingleWord not a Boolean.
;				   @Error 1 @Extended 8 Return 0 = $bSnapToGrid not a Boolean.
;				   @Error 1 @Extended 9 Return 0 = $iTxtDirection not an Integer, less than 0 or greater than 5, See Constants, $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iHorAlign
;				   |								2 = Error setting $iVertAlign
;				   |								4 = Error setting $iLastLineALign
;				   |								8 = Error setting $bExpandSIngleWord
;				   |								16 = Error setting $bSnapToGrid
;				   |								32 = Error setting $iTxtDirection
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iHorAlign must be set to $LOW_PAR_ALIGN_HOR_JUSTIFIED(2) before you can set $iLastLineAlign, and
;					$iLastLineAlign must be set to $LOW_PAR_LAST_LINE_JUSTIFIED(2) before $bExpandSingleWord can be set.
;					Note: $iTxtDirection constants 2,3, and 5 may not be available depending on your language settings.
;					Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParAlignment(ByRef $oObj, $iHorAlign, $iVertAlign, $iLastLineAlign, $bExpandSingleWord, $bSnapToGrid, $iTxtDirection)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avAlignment[6]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If __LOWriter_VarsAreNull($iHorAlign, $iVertAlign, $iLastLineAlign, $bExpandSingleWord, $bSnapToGrid, $iTxtDirection) Then
		__LOWriter_ArrayFill($avAlignment, $oObj.ParaAdjust(), $oObj.ParaVertAlignment(), $oObj.ParaLastLineAdjust(), $oObj.ParaExpandSingleWord(), _
				$oObj.SnapToGrid(), $oObj.WritingMode())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avAlignment)
	EndIf

	If ($iHorAlign <> Null) Then
		If Not __LOWriter_IntIsBetween($iHorAlign, $LOW_PAR_ALIGN_HOR_LEFT, $LOW_PAR_ALIGN_HOR_CENTER) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oObj.ParaAdjust = $iHorAlign
		$iError = ($oObj.ParaAdjust() = $iHorAlign) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iVertAlign <> Null) Then
		If Not __LOWriter_IntIsBetween($iVertAlign, $LOW_PAR_ALIGN_VERT_AUTO, $LOW_PAR_ALIGN_VERT_BOTTOM) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oObj.ParaVertAlignment = $iVertAlign
		$iError = ($oObj.ParaVertAlignment() = $iVertAlign) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iLastLineAlign <> Null) Then
		If Not __LOWriter_IntIsBetween($iLastLineAlign, $LOW_PAR_LAST_LINE_JUSTIFIED, $LOW_PAR_LAST_LINE_CENTER, "", $LOW_PAR_LAST_LINE_START) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oObj.ParaLastLineAdjust = $iLastLineAlign
		$iError = ($oObj.ParaLastLineAdjust() = $iLastLineAlign) ? $iError : BitOR($iError, 4)
	EndIf

	If ($bExpandSingleWord <> Null) Then
		If Not IsBool($bExpandSingleWord) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oObj.ParaExpandSingleWord = $bExpandSingleWord
		$iError = ($oObj.ParaExpandSingleWord() = $bExpandSingleWord) ? $iError : BitOR($iError, 8)
	EndIf

	If ($bSnapToGrid <> Null) Then
		If Not IsBool($bSnapToGrid) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$oObj.SnapToGrid = $bSnapToGrid
		$iError = ($oObj.SnapToGrid() = $bSnapToGrid) ? $iError : BitOR($iError, 16)
	EndIf

	If ($iTxtDirection <> Null) Then
		If Not __LOWriter_IntIsBetween($iTxtDirection, $LOW_TXT_DIR_LR_TB, $LOW_TXT_DIR_BT_LR) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		$oObj.WritingMode = $iTxtDirection
		$iError = ($oObj.WritingMode() = $iTxtDirection) ? $iError : BitOR($iError, 32)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_ParAlignment

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParBackColor
; Description ...: Set or Retrieve background color settings.
; Syntax ........: __LOWriter_ParBackColor(Byref $oObj, $iBackColor, $bBackTransparent)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iBackColor          - an integer value (-1-16777215). The color to make the background. Set in Long integer format. Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3. Set to $LOW_COLOR_OFF(-1) for "None".
;                  $bBackTransparent    - a boolean value. Whether the background color is transparent or not. True = visible.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iBackColor not an integer, less than -1 or greater than 16777215.
;				   @Error 1 @Extended 5 Return 0 = $bBackTransparent not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iBackColor
;				   |								2 = Error setting $bBackTransparent
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParBackColor(ByRef $oObj, $iBackColor, $bBackTransparent)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avColor[2]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If __LOWriter_VarsAreNull($iBackColor, $bBackTransparent) Then
		__LOWriter_ArrayFill($avColor, $oObj.ParaBackColor(), $oObj.ParaBackTransparent())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avColor)
	EndIf

	If ($iBackColor <> Null) Then
		If Not __LOWriter_IntIsBetween($iBackColor, $LOW_COLOR_OFF, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oObj.ParaBackColor = $iBackColor
		$iError = ($oObj.ParaBackColor() = $iBackColor) ? $iError : BitOR($iError, 1)
	EndIf

	If ($bBackTransparent <> Null) Then
		If Not IsBool($bBackTransparent) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oObj.ParaBackTransparent = $bBackTransparent
		$iError = ($oObj.ParaBackTransparent() = $bBackTransparent) ? $iError : BitOR($iError, 2)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_ParBackColor

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParBorderPadding
; Description ...: Set or retrieve the Border Padding (spacing between the Paragraph and border) settings.
; Syntax ........: __LOWriter_ParBorderPadding(Byref $oObj, $iAll, $iTop, $iBottom, $iLeft, $iRight)
; Parameters ....: $oObj                - [in/out] an object. A ParagraphStyle object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $iAll                - an integer value. Set all four padding distances to one distance in Micrometers (uM).
;                  $iTop                - an integer value. Set the Top Distance between the Border and Paragraph in Micrometers(uM).
;                  $iBottom             - an integer value. Set the Bottom Distance between the Border and Paragraph in Micrometers(uM).
;                  $iLeft               - an integer value. Set the Left Distance between the Border and Paragraph in Micrometers(uM).
;                  $iRight              - an integer value. Set the Right Distance between the Border and Paragraph in Micrometers(uM).
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 3 Return 0 = Passed Object to internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iAll not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iTop not an Integer.
;				   @Error 1 @Extended 6 Return 0 = $iBottom not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $Left not an Integer.
;				   @Error 1 @Extended 8 Return 0 = $iRight not an Integer.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iAll border distance
;				   |								2 = Error setting $iTop border distance
;				   |								4 = Error setting $iBottom border distance
;				   |								8 = Error setting $iLeft border distance
;				   |								16 = Error setting $iRight border distance
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParBorderPadding(ByRef $oObj, $iAll, $iTop, $iBottom, $iLeft, $iRight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiBPadding[5]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If __LOWriter_VarsAreNull($iAll, $iTop, $iBottom, $iLeft, $iRight) Then
		__LOWriter_ArrayFill($aiBPadding, $oObj.BorderDistance(), $oObj.TopBorderDistance(), $oObj.BottomBorderDistance(), _
				$oObj.LeftBorderDistance(), $oObj.RightBorderDistance())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $aiBPadding)
	EndIf

	If ($iAll <> Null) Then
		If Not __LOWriter_IntIsBetween($iAll, 0, $iAll) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oObj.BorderDistance = $iAll
		$iError = (__LOWriter_IntIsBetween($oObj.BorderDistance(), $iAll - 1, $iAll + 1)) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iTop <> Null) Then
		If Not __LOWriter_IntIsBetween($iTop, 0, $iTop) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oObj.TopBorderDistance = $iTop
		$iError = (__LOWriter_IntIsBetween($oObj.TopBorderDistance(), $iTop - 1, $iTop + 1)) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iBottom <> Null) Then
		If Not __LOWriter_IntIsBetween($iBottom, 0, $iBottom) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oObj.BottomBorderDistance = $iBottom
		$iError = (__LOWriter_IntIsBetween($oObj.BottomBorderDistance(), $iBottom - 1, $iBottom + 1)) ? $iError : BitOR($iError, 4)
	EndIf

	If ($iLeft <> Null) Then
		If Not __LOWriter_IntIsBetween($iLeft, 0, $iLeft) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oObj.LeftBorderDistance = $iLeft
		$iError = (__LOWriter_IntIsBetween($oObj.LeftBorderDistance(), $iLeft - 1, $iLeft + 1)) ? $iError : BitOR($iError, 8)
	EndIf

	If ($iRight <> Null) Then
		If Not __LOWriter_IntIsBetween($iRight, 0, $iRight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$oObj.RightBorderDistance = $iRight
		$iError = (__LOWriter_IntIsBetween($oObj.RightBorderDistance(), $iRight - 1, $iRight + 1)) ? $iError : BitOR($iError, 16)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_ParBorderPadding

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParDropCaps
; Description ...: Set or Retrieve DropCaps settings
; Syntax ........: __LOWriter_ParDropCaps(Byref $oObj, $iNumChar, $iLines, $iSpcTxt, $bWholeWord, $sCharStyle)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iNumChar            - an integer value. The number of characters to make into DropCaps. Min is 0, max is 9.
;                  $iLines              - an integer value. The number of lines to drop down, min is 0, max is 9, cannot be 1.
;                  $iSpcTxt             - an integer value. The distance between the drop cap and the following text.
;                  $bWholeWord          - a boolean value. Whether to DropCap the whole first word. (Nullifys $iNumChars.)
;                  $sCharStyle          - a string value. The character style to use for the DropCaps. See Remarks.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 5 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 6 Return 0 = $iNumChar not an integer, less than 0 or greater than 9.
;				   @Error 1 @Extended 7 Return 0 = $iLines not an Integer, less than 0, equal to 1 or greater than 9
;				   @Error 1 @Extended 8 Return 0 = $iSpaceTxt not an Integer, or less than 0.
;				   @Error 1 @Extended 9 Return 0 = $bWholeWord not a Boolean.
;				   @Error 1 @Extended 10 Return 0 = $sCharStyle not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving DropCap Format Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iNumChar
;				   |								2 = Error setting $iLines
;				   |								4 = Error setting $iSpcTxt
;				   |								8 = Error setting $bWholeWord
;				   |								16 = Error setting $sCharStyle
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Set $iNumChars, $iLines, $iSpcTxt to 0 to disable DropCaps.
;					I am unable to find a way to set Drop Caps character style to "None" as is available in the User Interface.
;					When it is set to "None" Libre returns a blank string ("") but setting it to a blank string throws a COM
;					error/Exception, even when attempting to set it to Libre's own return value without any in-between
;					variables, in case I was mistaken as to it being a blank string, but this still caused a COM error. So
;					consequently, you cannot set Character Style to "None", but you can still disable Drop Caps as noted above.
;				Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
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

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	$tDCFrmt = $oObj.DropCapFormat()
	If Not IsObj($tDCFrmt) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iNumChar, $iLines, $iSpcTxt, $bWholeWord, $sCharStyle) Then
		__LOWriter_ArrayFill($avDropCaps, $tDCFrmt.Count(), $tDCFrmt.Lines(), $tDCFrmt.Distance(), $oObj.DropCapWholeWord(), _
				__LOWriter_CharStyleNameToggle($oObj.DropCapCharStyleName(), True))
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avDropCaps)
	EndIf

	If Not __LOWriter_VarsAreNull($iNumChar, $iLines, $iSpcTxt) Then
		If ($iNumChar <> Null) Then
			If Not __LOWriter_IntIsBetween($iNumChar, 0, 9) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
			$tDCFrmt.Count = $iNumChar
		EndIf

		If ($iLines <> Null) Then
			If Not __LOWriter_IntIsBetween($iLines, 0, 9, 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
			$tDCFrmt.Lines = $iLines
		EndIf

		If ($iSpcTxt <> Null) Then
			If Not __LOWriter_IntIsBetween($iSpcTxt, 0, $iSpcTxt) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
			$tDCFrmt.Distance = $iSpcTxt
		EndIf

		$oObj.DropCapFormat = $tDCFrmt
		$iError = ($iNumChar = Null) ? $iError : ($tDCFrmt.Count() = $iNumChar) ? $iError : BitOR($iError, 1)
		$iError = ($iLines = Null) ? $iError : ($tDCFrmt.Lines() = $iLines) ? $iError : BitOR($iError, 2)
		$iError = ($iSpcTxt = Null) ? $iError : (__LOWriter_IntIsBetween($tDCFrmt.Distance(), $iSpcTxt - 1, $iSpcTxt + 1)) ? $iError : BitOR($iError, 4)
	EndIf

	If ($bWholeWord <> Null) Then
		If Not IsBool($bWholeWord) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		$oObj.DropCapWholeWord = $bWholeWord
		$iError = ($oObj.DropCapWholeWord() = $bWholeWord) ? $iError : BitOR($iError, 8)
	EndIf

	If ($sCharStyle <> Null) Then
		If Not IsString($sCharStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 10, 0)
		$sCharStyle = __LOWriter_CharStyleNameToggle($sCharStyle)
		$oObj.DropCapCharStyleName = $sCharStyle
		$iError = ($oObj.DropCapCharStyleName() = $sCharStyle) ? $iError : BitOR($iError, 16)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_ParDropCaps

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParHasTabStop
; Description ...: Check whether a Paragraph has a requested TabStop created already.
; Syntax ........: __LOWriter_ParHasTabStop(Byref $oObj, $iTabStop)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iTabStop            - an integer value. The Tab Stop to look for.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iTabStop not an Integer.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve ParaTabStops Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean  = True if Paragraph has the requested TabStop. Else False.
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

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iTabStop) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$atTabStops = $oObj.ParaTabStops()
	If Not IsArray($atTabStops) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	For $i = 0 To UBound($atTabStops) - 1
		If ($atTabStops[$i].Position() = $iTabStop) Then Return SetError($__LOW_STATUS_SUCCESS, 0, True)
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
	Next

	Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 0, False)
EndFunc   ;==>__LOWriter_ParHasTabStop

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParHyphenation
; Description ...: Set or Retrieve Hyphenation settings.
; Syntax ........: __LOWriter_ParHyphenation(Byref $oObj, $bAutoHyphen, $bHyphenNoCaps, $iMaxHyphens, $iMinLeadingChar, $iMinTrailingChar)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $bAutoHyphen         - a boolean value. Whether  automatic hyphenation is applied.
;                  $bHyphenNoCaps       - a boolean value.  Setting to true will disable hyphenation of words written in CAPS for this paragraph. Libre 6.4 and up.
;                  $iMaxHyphens         - an integer value. The maximum number of consecutive hyphens. Min 0, Max 99.
;                  $iMinLeadingChar     - an integer value. Specifies the minimum number of characters to remain before the hyphen character (when hyphenation is applied). Min 2, max 9.
;                  $iMinTrailingChar    - an integer value. Specifies the minimum number of characters to remain after the hyphen character (when hyphenation is applied). Min 2, max 9.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bAutoHyphen not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bHyphenNoCaps not  a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iMaxHyphens not an Integer, less than 0, or greater than 99.
;				   @Error 1 @Extended 7 Return 0 = $iMinLeadingChar not an Integer, less than 2 or greater than 9.
;				   @Error 1 @Extended 8 Return 0 = $iMinTrailingChar not an Integer, less than 2 or greater than 9.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bAutoHyphen
;				   |								2 = Error setting $bHyphenNoCaps
;				   |								4 = Error setting $iMaxHyphens
;				   |								8 = Error setting $iMinLeadingChar
;				   |								16 = Error setting $iMinTrailingChar
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 6.4.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 or 5 Element Array with values in order of function parameters. If the current Libre Office Version is below 6.4, then the Array returned will contain 4 elements because $bHyphenNoCaps is not available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $bAutoHyphen set to True for the rest of the settings to be activated, but they will be still
;					successfully set regardless.
;					Call this function with only the Object parameter and all other parameters set to Null keyword, to
;					get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParHyphenation(ByRef $oObj, $bAutoHyphen, $bHyphenNoCaps, $iMaxHyphens, $iMinLeadingChar, $iMinTrailingChar)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avHyphenation[4]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If __LOWriter_VarsAreNull($bAutoHyphen, $bHyphenNoCaps, $iMaxHyphens, $iMinLeadingChar, $iMinTrailingChar) Then
		If __LOWriter_VersionCheck(6.4) Then
			__LOWriter_ArrayFill($avHyphenation, $oObj.ParaIsHyphenation(), $oObj.ParaHyphenationNoCaps(), $oObj.ParaHyphenationMaxHyphens(), _
					$oObj.ParaHyphenationMaxLeadingChars(), $oObj.ParaHyphenationMaxTrailingChars())
		Else
			__LOWriter_ArrayFill($avHyphenation, $oObj.ParaIsHyphenation(), $oObj.ParaHyphenationMaxHyphens(), _
					$oObj.ParaHyphenationMaxLeadingChars(), $oObj.ParaHyphenationMaxTrailingChars())
		EndIf

		Return SetError($__LOW_STATUS_SUCCESS, 1, $avHyphenation)
	EndIf

	If ($bAutoHyphen <> Null) Then
		If Not IsBool($bAutoHyphen) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oObj.ParaIsHyphenation = $bAutoHyphen
		$iError = ($oObj.ParaIsHyphenation = $bAutoHyphen) ? $iError : BitOR($iError, 1)
	EndIf

	If ($bHyphenNoCaps <> Null) Then
		If Not IsBool($bHyphenNoCaps) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		If Not __LOWriter_VersionCheck(6.4) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
		$oObj.ParaHyphenationNoCaps = $bHyphenNoCaps
		$iError = ($oObj.ParaHyphenationNoCaps = $bHyphenNoCaps) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iMaxHyphens <> Null) Then
		If Not __LOWriter_IntIsBetween($iMaxHyphens, 0, 99) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oObj.ParaHyphenationMaxHyphens = $iMaxHyphens
		$iError = ($oObj.ParaHyphenationMaxHyphens = $iMaxHyphens) ? $iError : BitOR($iError, 4)
	EndIf

	If ($iMinLeadingChar <> Null) Then
		If Not __LOWriter_IntIsBetween($iMinLeadingChar, 2, 9) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oObj.ParaHyphenationMaxLeadingChars = $iMinLeadingChar
		$iError = ($oObj.ParaHyphenationMaxLeadingChars = $iMinLeadingChar) ? $iError : BitOR($iError, 8)
	EndIf

	If ($iMinTrailingChar <> Null) Then
		If Not __LOWriter_IntIsBetween($iMinTrailingChar, 2, 9) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$oObj.ParaHyphenationMaxTrailingChars = $iMinTrailingChar
		$iError = ($oObj.ParaHyphenationMaxTrailingChars = $iMinTrailingChar) ? $iError : BitOR($iError, 16)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_ParHyphenation

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParIndent
; Description ...: Set or Retrieve Indent settings.
; Syntax ........: __LOWriter_ParIndent(Byref $oObj, $iBeforeTxt, $iAfterTxt, $iFirstLine, $bAutoFirstLine)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iBeforeTxt          - an integer value. The amount of space that you want to indent the paragraph from the page margin. If you want the paragraph to extend into the page margin, enter a negative number. Set in MicroMeters(uM) Min. -9998989, Max.17094
;                  $iAfterTxt           - an integer value. The amount of space that you want to indent the paragraph from the page margin. If you want the paragraph to extend into the page margin, enter a negative number. Set in MicroMeters(uM) Min. -9998989, Max.17094
;                  $iFirstLine          - an integer value. Indentation distance of the first line of a paragraph. Set in MicroMeters(uM) Min. -57785, Max.17094.
;                  $bAutoFirstLine      - a boolean value. Whether the first line should be indented automatically.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iBeforeText not an integer, less than -9998989 or more than 17094 uM.
;				   @Error 1 @Extended 5 Return 0 = $iAfterText not an integer, less than -9998989 or more than 17094 uM.
;				   @Error 1 @Extended 6 Return 0 = $iFirstLine not an integer, less than -57785 or more than 17094 uM.
;				   @Error 1 @Extended 7 Return 0 = $bAutoFirstLine not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iBeforeTxt
;				   |								2 = Error setting $iAfterTxt
;				   |								4 = Error setting $iFirstLine
;				   |								8 = Error setting $bAutoFirstLine
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iFirstLine Indent cannot be set if $bAutoFirstLine is set to True.
;					Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParIndent(ByRef $oObj, $iBeforeTxt, $iAfterTxt, $iFirstLine, $bAutoFirstLine)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avIndent[4]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If __LOWriter_VarsAreNull($iBeforeTxt, $iAfterTxt, $iFirstLine, $bAutoFirstLine) Then
		__LOWriter_ArrayFill($avIndent, $oObj.ParaLeftMargin(), $oObj.ParaRightMargin(), $oObj.ParaFirstLineIndent(), $oObj.ParaIsAutoFirstLineIndent())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avIndent)
	EndIf

	; Min: -9998989;Max: 17094
	If ($iBeforeTxt <> Null) Then
		If Not __LOWriter_IntIsBetween($iBeforeTxt, -9998989, 17094) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oObj.ParaLeftMargin = $iBeforeTxt
		$iError = (__LOWriter_NumIsBetween(($oObj.ParaLeftMargin()), ($iBeforeTxt - 1), ($iBeforeTxt + 1))) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iAfterTxt <> Null) Then
		If Not __LOWriter_IntIsBetween($iAfterTxt, -9998989, 17094) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oObj.ParaRightMargin = $iAfterTxt
		$iError = (__LOWriter_NumIsBetween(($oObj.ParaRightMargin()), ($iAfterTxt - 1), ($iAfterTxt + 1))) ? $iError : BitOR($iError, 2)
	EndIf

	; max 17094; min;-57785
	If ($iFirstLine <> Null) Then
		If Not __LOWriter_IntIsBetween($iFirstLine, -57785, 17094) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oObj.ParaFirstLineIndent = $iFirstLine
		$iError = (__LOWriter_NumIsBetween(($oObj.ParaFirstLineIndent()), ($iFirstLine - 1), ($iFirstLine + 1))) ? $iError : BitOR($iError, 4)
	EndIf

	If ($bAutoFirstLine <> Null) Then
		If Not IsBool($bAutoFirstLine) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oObj.ParaIsAutoFirstLineIndent = $bAutoFirstLine
		$iError = ($oObj.ParaIsAutoFirstLineIndent() = $bAutoFirstLine) ? $iError : BitOR($iError, 8)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_ParIndent

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParOutLineAndList
; Description ...: Set and Retrieve the Outline and List settings.
; Syntax ........: __LOWriter_ParOutLineAndList(Byref $oObj, $iOutline, $sNumStyle, $bParLineCount, $iLineCountVal)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iOutline            - an integer value (0-10). The Outline Level, see Constants, $LOW_OUTLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sNumStyle           - a string value. Specifies the name of the style for the Paragraph numbering. Set to "" for None.
;                  $bParLineCount       - a boolean value. Whether the paragraph is included in the line numbering.
;                  $iLineCountVal       - an integer value. The start value for numbering if a new numbering starts at this paragraph. Set to 0 for no line numbering restart.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 5 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 6 Return 0 = $iOutline not an integer, less than 0 or greater than 10. See constants, $LOW_OUTLINE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 7 Return 0 = $sNumStyle not a String.
;				   @Error 1 @Extended 8 Return 0 = $bParLineCount not a Boolean.
;				   @Error 1 @Extended 9 Return 0 = $iLineCountVal Not an Integer or less than 0.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iOutline
;				   |								2 = Error setting $sNumStyle
;				   |								4 = Error setting $bParLineCount
;				   |								8 = Error setting $iLineCountVal
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParOutLineAndList(ByRef $oObj, $iOutline, $sNumStyle, $bParLineCount, $iLineCountVal)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avOutlineNList[4]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	If __LOWriter_VarsAreNull($iOutline, $sNumStyle, $bParLineCount, $iLineCountVal) Then
		__LOWriter_ArrayFill($avOutlineNList, $oObj.OutlineLevel(), $oObj.NumberingStyleName(), $oObj.ParaLineNumberCount(), _
				$oObj.ParaLineNumberStartValue())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avOutlineNList)
	EndIf

	If ($iOutline <> Null) Then
		If Not __LOWriter_IntIsBetween($iOutline, $LOW_OUTLINE_BODY, $LOW_OUTLINE_LEVEL_10) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oObj.OutlineLevel = $iOutline
		$iError = ($oObj.OutlineLevel = $iOutline) ? $iError : BitOR($iError, 1)
	EndIf

	If ($sNumStyle <> Null) Then
		If Not IsString($sNumStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oObj.NumberingStyleName = $sNumStyle
		$iError = ($oObj.NumberingStyleName = $sNumStyle) ? $iError : BitOR($iError, 2)
	EndIf

	If ($bParLineCount <> Null) Then
		If Not IsBool($bParLineCount) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$oObj.ParaLineNumberCount = $bParLineCount
		$iError = ($oObj.ParaLineNumberCount = $bParLineCount) ? $iError : BitOR($iError, 4)
	EndIf

	If ($iLineCountVal <> Null) Then
		If Not __LOWriter_IntIsBetween($iLineCountVal, 0, $iLineCountVal) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		$oObj.ParaLineNumberStartValue = $iLineCountVal
		$iError = ($oObj.ParaLineNumberStartValue = $iLineCountVal) ? $iError : BitOR($iError, 8)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_ParOutLineAndList

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParPageBreak
; Description ...: Set or Retrieve Page Break Settings.
; Syntax ........: __LOWriter_ParPageBreak(Byref $oObj, $iBreakType, $iPgNumOffSet, $sPageStyle)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iBreakType          - an integer value (0-6). The Page Break Type. See Constants, $LOW_BREAK_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iPgNumOffSet        - an integer value. If a page break property is set at a paragraph, this property contains the new value for the page number.
;                  $sPageStyle          - a string value. Creates a page break before the paragraph it belongs to and assigns the value as the name of the new page style to use. Note: If you set this parameter, to remove the page break setting you must set this to "".
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 5 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 6 Return 0 = $iBreakType not an integer, less than 0 or greater than 6. See Constants, $LOW_BREAK_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 7 Return 0 = $iPgNumOffSet not an Integer or less than 0.
;				   @Error 1 @Extended 8 Return 0 = $sPageStyle not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iBreakType
;				   |								2 = Error setting $iPgNumOffSet
;				   |								4 = Error setting $sPageStyle
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Break Type must be set before PageStyle will be able to be set, and page style needs set before
;					$iPgNumOffSet can be set.
;					Libre doesn't directly show in its User interface options for Break type constants #3 and #6 (Column both)
;						and (Page both), but  doesn't throw an error when being set to either one, so they are included here,
;						though I'm not sure if they will work correctly.
;					Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParPageBreak(ByRef $oObj, $iBreakType, $iPgNumOffSet, $sPageStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avPageBreak[3]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	If __LOWriter_VarsAreNull($iBreakType, $iPgNumOffSet, $sPageStyle) Then
		__LOWriter_ArrayFill($avPageBreak, $oObj.BreakType(), $oObj.PageNumberOffset(), $oObj.PageDescName())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avPageBreak)
	EndIf

	If ($iBreakType <> Null) Then
		If Not __LOWriter_IntIsBetween($iBreakType, $LOW_BREAK_NONE, $LOW_BREAK_PAGE_BOTH) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oObj.BreakType = $iBreakType
		$iError = ($oObj.BreakType = $iBreakType) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iPgNumOffSet <> Null) Then
		If Not __LOWriter_IntIsBetween($iPgNumOffSet, 0, $iPgNumOffSet) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oObj.PageNumberOffset = $iPgNumOffSet
		$iError = ($oObj.PageNumberOffset = $iPgNumOffSet) ? $iError : BitOR($iError, 2)
	EndIf

	If ($sPageStyle <> Null) Then
		If Not IsString($sPageStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$oObj.PageDescName = $sPageStyle
		$iError = ($oObj.PageDescName = $sPageStyle) ? $iError : BitOR($iError, 4)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_ParPageBreak

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParShadow
; Description ...: Set or Retrieve the Shadow settings for a Paragraph.
; Syntax ........: __LOWriter_ParShadow(Byref $oObj, $iWidth, $iColor, $bTransparent, $iLocation)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iWidth              - an integer value. The width of the shadow set in Micrometers.
;                  $iColor              - an integer value (0-16777215). The color of the shadow, set in Long Integer format. Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bTransparent        - a boolean value. Whether or not the shadow is transparent.
;                  $iLocation           - an integer value (0-4). The location of the shadow compared to the paragraph. See Constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iWidth not an integer or less than 0.
;				   @Error 1 @Extended 5 Return 0 = $iColor not an integer, less than 0 or greater than 16777215.
;				   @Error 1 @Extended 6 Return 0 = $bTransparent not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $iLocation not an Integer, less than 0 or greater than 4. See Constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Shadow Format Object.
;				   @Error 2 @Extended 2 Return 0 = Error retrieving Shadow Format Object for Error Checking.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iWidth
;				   |								2 = Error setting $iColor
;				   |								4 = Error setting $bTransparent
;				   |								8 = Error setting $iLocation
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
;					Note: LibreOffice may change the shadow width +/- a Micrometer.
; Related .......: _LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong,  _LOWriter_ConvertFromMicrometer,
;					_LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParShadow(ByRef $oObj, $iWidth, $iColor, $bTransparent, $iLocation)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $tShdwFrmt
	Local $avShadow[4]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	$tShdwFrmt = $oObj.ParaShadowFormat()
	If Not IsObj($tShdwFrmt) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iWidth, $iColor, $bTransparent, $iLocation) Then
		__LOWriter_ArrayFill($avShadow, $tShdwFrmt.ShadowWidth(), $tShdwFrmt.Color(), $tShdwFrmt.IsTransparent(), $tShdwFrmt.Location())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avShadow)
	EndIf

	If ($iWidth <> Null) Then
		If Not __LOWriter_IntIsBetween($iWidth, 0, $iWidth) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$tShdwFrmt.ShadowWidth = $iWidth
	EndIf

	If ($iColor <> Null) Then
		If Not __LOWriter_IntIsBetween($iColor, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$tShdwFrmt.Color = $iColor
	EndIf

	If ($bTransparent <> Null) Then
		If Not IsBool($bTransparent) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$tShdwFrmt.IsTransparent = $bTransparent
	EndIf

	If ($iLocation <> Null) Then
		If Not __LOWriter_IntIsBetween($iLocation, $LOW_SHADOW_NONE, $LOW_SHADOW_BOTTOM_RIGHT) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$tShdwFrmt.Location = $iLocation
	EndIf

	$oObj.ParaShadowFormat = $tShdwFrmt
	; Error Checking
	$tShdwFrmt = $oObj.ParaShadowFormat()
	If Not IsObj($tShdwFrmt) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	$iError = ($iWidth = Null) ? $iError : ($tShdwFrmt.ShadowWidth() = $iWidth) ? $iError : BitOR($iError, 1)
	$iError = ($iColor = Null) ? $iError : ($tShdwFrmt.Color() = $iColor) ? $iError : BitOR($iError, 2)
	$iError = ($bTransparent = Null) ? $iError : ($tShdwFrmt.IsTransparent() = $bTransparent) ? $iError : BitOR($iError, 4)
	$iError = ($iLocation = Null) ? $iError : ($tShdwFrmt.Location() = $iLocation) ? $iError : BitOR($iError, 8)

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_ParShadow

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParSpace
; Description ...: Set and Retrieve Line Spacing settings.
; Syntax ........: __LOWriter_ParSpace(Byref $oObj, $iAbovePar, $iBelowPar, $bAddSpace, $iLineSpcMode, $iLineSpcHeight, $bPageLineSpc)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iAbovePar           - an integer value. The Space above a paragraph, in Micrometers. Min 0 Micrometers (uM) Max 10,008 uM.
;                  $iBelowPar           - an integer value. The Space Below a paragraph, in Micrometers. Min 0 Micrometers (uM) Max 10,008 uM.
;                  $bAddSpace           - a boolean value. If true, the top and bottom margins of the paragraph should not be applied when the previous and next paragraphs have the same style. Libre Office Version 3.6 and Up.
;                  $iLineSpcMode        - an integer value (0-3). The type of the line spacing of a paragraph. See Constants, $LOW_LINE_SPC_MODE_* as defined in LibreOfficeWriter_Constants.au3, also notice min and max values for each.
;                  $iLineSpcHeight      - an integer value. This value specifies the height in regard to Mode. See Remarks.
;                  $bPageLineSpc        - a boolean value. Determines if the register mode is applied to a paragraph. See Remarks.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iAbovePar not an integer, less than 0 or more than 10008 uM.
;				   @Error 1 @Extended 5 Return 0 = $iBelowPar not an integer, less than 0 or more than 10008 uM.
;				   @Error 1 @Extended 6 Return 0 = $bAddSpc not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $iLineSpcMode Not an integer, less than 0 or greater than 3. See Constants, $LOW_LINE_SPC_MODE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 8 Return 0 = $iLineSpcHeight not an Integer.
;				   @Error 1 @Extended 9 Return 0 = $iLineSpcMode set to 0(Proportional) and $iLineSpcHeight less than 6(%) or greater than 65535(%).
;				   @Error 1 @Extended 10 Return 0 = $iLineSpcMode set to 1 or 2(Minimum, or Leading) and $iLineSpcHeight less than 0 uM or greater than 10008 uM
;				   @Error 1 @Extended 11 Return 0 = $iLineSpcMode set to 3(Fixed) and $iLineSpcHeight less than 51 uM or greater than 10008 uM.
;				   @Error 1 @Extended 12 Return 0 = $bPageLineSpc not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving ParaLineSpacing Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iAbovePar
;				   |								2 = Error setting $iBelowPar
;				   |								4 = Error setting $bAddSpace
;				   |								8 = Error setting $iLineSpcMode
;				   |								16 = Error setting $iLineSpcHeight
;				   |								32 = Error setting bPageLineSpc
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $bPageLineSpc(Register mode) is only used if the register mode property of the page style is switched
;						on. $bPageLineSpc(Register Mode) Aligns the baseline of each line of text to a vertical document grid,
;						so that each line is the same height.
;					Note: The settings in Libre Office, (Single,1.15, 1.5, Double,) Use the Proportional mode, and are just
;						varying percentages. e.g Single = 100, 1.15 = 115%, 1.5 = 150%, Double = 200%.
;					$iLineSpcHeight depends on the $iLineSpcMode used, see constants for accepted Input values.
;					Note: $iAbovePar, $iBelowPar, $iLineSpcHeight may change +/- 1 MicroMeter once set.
;					Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParSpace(ByRef $oObj, $iAbovePar, $iBelowPar, $bAddSpace, $iLineSpcMode, $iLineSpcHeight, $bPageLineSpc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tLine
	Local $iError = 0
	Local $avSpacing[5]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If __LOWriter_VarsAreNull($iAbovePar, $iBelowPar, $bAddSpace, $iLineSpcMode, $iLineSpcHeight, $bPageLineSpc) Then

		If __LOWriter_VersionCheck(3.6) Then
			__LOWriter_ArrayFill($avSpacing, $oObj.ParaTopMargin(), $oObj.ParaBottomMargin(), $oObj.ParaContextMargin(), _
					$oObj.ParaLineSpacing.Mode(), $oObj.ParaLineSpacing.Height(), $oObj.ParaRegisterModeActive())
		Else
			__LOWriter_ArrayFill($avSpacing, $oObj.ParaTopMargin(), $oObj.ParaBottomMargin(), $oObj.ParaLineSpacing.Mode(), $oObj.ParaLineSpacing.Height(), _
					$oObj.ParaRegisterModeActive())
		EndIf
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avSpacing)
	EndIf

	If ($iAbovePar <> Null) Then
		If Not __LOWriter_IntIsBetween($iAbovePar, 0, 10008) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oObj.ParaTopMargin = $iAbovePar
		$iError = (__LOWriter_NumIsBetween(($oObj.ParaTopMargin()), ($iAbovePar - 1), ($iAbovePar + 1))) ? $iError : BitOR($iError, 1)
	EndIf

	If ($iBelowPar <> Null) Then
		If Not __LOWriter_IntIsBetween($iBelowPar, 0, 10008) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oObj.ParaBottomMargin = $iBelowPar
		$iError = (__LOWriter_NumIsBetween(($oObj.ParaBottomMargin()), ($iBelowPar - 1), ($iBelowPar + 1))) ? $iError : BitOR($iError, 2)
	EndIf

	If ($bAddSpace <> Null) Then
		If Not IsBool($bAddSpace) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		If Not __LOWriter_VersionCheck(3.6) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
		$oObj.ParaContextMargin = $bAddSpace
		$iError = ($oObj.ParaContextMargin = $bAddSpace) ? $iError : BitOR($iError, 4)
	EndIf

	If ($iLineSpcMode <> Null) Then
		If Not __LOWriter_IntIsBetween($iLineSpcMode, $LOW_LINE_SPC_MODE_PROP, $LOW_LINE_SPC_MODE_FIX) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$tLine = $oObj.ParaLineSpacing()
		If Not IsObj($tLine) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
		$tLine.Mode = $iLineSpcMode
		$oObj.ParaLineSpacing = $tLine
		$iError = ($oObj.ParaLineSpacing.Mode() = $iLineSpcMode) ? $iError : BitOR($iError, 8)
	EndIf

	If ($iLineSpcHeight <> Null) Then
		If Not IsInt($iLineSpcHeight) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$tLine = $oObj.ParaLineSpacing()
		If Not IsObj($tLine) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

		Switch $tLine.Mode()
			Case $LOW_LINE_SPC_MODE_PROP ;Proportional
				If Not __LOWriter_IntIsBetween($iLineSpcHeight, 6, 65535) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0) ; Min setting on Proportional is 6%
			Case $LOW_LINE_SPC_MODE_MIN, $LOW_LINE_SPC_MODE_LEADING ;Minimum and Leading Modes
				If Not __LOWriter_IntIsBetween($iLineSpcHeight, 0, 10008) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 10, 0)
			Case $LOW_LINE_SPC_MODE_FIX ;Fixed Line Spacing Mode
				If Not __LOWriter_IntIsBetween($iLineSpcHeight, 51, 10008) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 11, 0) ; Min spacing is 51 when Fixed Mode
		EndSwitch
		$tLine.Height = $iLineSpcHeight
		$oObj.ParaLineSpacing = $tLine
		$iError = (__LOWriter_NumIsBetween(($oObj.ParaLineSpacing.Height()), ($iLineSpcHeight - 1), ($iLineSpcHeight + 1))) ? $iError : BitOR($iError, 16)
	EndIf

	If ($bPageLineSpc <> Null) Then
		If Not IsBool($bPageLineSpc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 12, 0)
		$oObj.ParaRegisterModeActive = $bPageLineSpc
		$iError = ($oObj.ParaRegisterModeActive() = $bPageLineSpc) ? $iError : BitOR($iError, 32)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_ParSpace

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParStyleNameToggle
; Description ...: Toggle from Par Style Display Name to Internal Name for error checking, or setting retrieval.
; Syntax ........: __LOWriter_ParStyleNameToggle(Byref $sParStyle[, $bReverse = False])
; Parameters ....: $sParStyle           - a string value. The ParStyle Name to Toggle.
;                  $bReverse            - [optional] a boolean value. Default is False. If True, Reverse toggles the name.
; Return values .: Success: String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $sParStyle not a String.
;				   @Error 1 @Extended 2 Return 0 = $bReverse not a Boolean.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. Paragraph Style Name was Successfully toggled. Returning changed name as a string.
;				   @Error 0 @Extended 1 Return String = Success. Paragraph Style Name was Successfully reverse toggled. Returning changed name as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParStyleNameToggle($sParStyle, $bReverse = False)
	If Not IsString($sParStyle) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bReverse) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	If ($bReverse = False) Then
		$sParStyle = ($sParStyle = "Default Paragraph Style") ? "Standard" : $sParStyle
		$sParStyle = ($sParStyle = "Complimentary Close") ? "Salutation" : $sParStyle
		Return SetError($__LOW_STATUS_SUCCESS, 0, $sParStyle)
	Else
		$sParStyle = ($sParStyle = "Standard") ? "Default Paragraph Style" : $sParStyle
		$sParStyle = ($sParStyle = "Salutation") ? "Complimentary Close" : $sParStyle
		Return SetError($__LOW_STATUS_SUCCESS, 1, $sParStyle)
	EndIf
EndFunc   ;==>__LOWriter_ParStyleNameToggle

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParTabStopCreate
; Description ...: Create a new TabStop for a Paragraph.
; Syntax ........: __LOWriter_ParTabStopCreate(Byref $oObj, $iPosition, $iAlignment, $iFillChar, $iDecChar)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iPosition           - an integer value. The TabStop position/length to set the new TabStop to. Set in Micrometers (uM). See Remarks.
;                  $iAlignment          - an integer value (0-4). The position of where the end of a Tab is aligned to compared to the text. See Constants, $LOW_TAB_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iFillChar           - an integer value. The Asc (see autoit function) value of any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
;                  $iDecChar            - an integer value. Enter a character(in Asc Value(See Autoit Function)) that you want the decimal tab to use as a decimal separator. Can only be set if $iAlignment is set to $LOW_TAB_ALIGN_DECIMAL.
; Return values .: Success: Integer.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 5 Return 0 = Passed Object to internal function not an Object.
;				   @Error 1 @Extended 6 Return 0 = $iFillChar not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iAlignment not an Integer, less than 0 or greater than 4. See Constants, $LOW_TAB_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 8 Return 0 = $iDecChar not an Integer.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving ParaTabStops Array Object.
;				   @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.style.TabStop" Object.
;				   @Error 2 @Extended 3 Return 0 = Error retrieving list of TabStop Positions.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to identify the new Tabstop once inserted. in $iPosition.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return Integer = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iPosition
;				   |								2 = Error setting $iFillChar
;				   |								4 = Error setting $iAlignment
;				   |								8 = Error setting $iDecChar
;				   |						Note: $iNewTabStop position is still returned as even though some settings weren't successfully set, the new TabStop was still created.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success. Settings were successfully set. New TabStop position is returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iPosition once set can vary +/- 1 uM. To ensure you can identify the tabstop to modify it again,
;						This function returns the new TabStop position in @Extended when $iPosition is set, return value will
;						be set to 2. See Return Values.
;					Note: Since $iPosition can fluctuate +/- 1 uM when it is inserted into LibreOffice, it is possible to
;						accidentally overwrite an already existing TabStop.
;					Note: $iFillChar, Libre's Default value, "None" is in reality a space character which is Asc value 32.
;						The other values offered by Libre are: Period (ASC 46), Dash (ASC 45) and Underscore (ASC 95). You can
;						also enter a custom ASC value. See ASC Autoit Func. and "ASCII Character Codes" in the Autoit help file.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer
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

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	$atTabStops = $oObj.ParaTabStops()
	If Not IsArray($atTabStops) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	$tTabStruct = __LOWriter_CreateStruct("com.sun.star.style.TabStop")
	If @error Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	$tTabStruct.Position = $iPosition
	$tTabStruct.FillChar = 32
	; If set to 0 Libre sets fill character to Null instead of setting to None. 32 = None.(Space character)
	$tTabStruct.Alignment = 0
	$tTabStruct.DecimalChar = 0

	If ($iFillChar <> Null) Then
		If Not IsInt($iFillChar) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$tTabStruct.FillChar = ($iFillChar = 0) ? 32 : $iFillChar
	EndIf

	If ($iAlignment <> Null) Then
		If Not __LOWriter_IntIsBetween($iAlignment, $LOW_TAB_ALIGN_LEFT, $LOW_TAB_ALIGN_DEFAULT) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$tTabStruct.Alignment = $iAlignment
	EndIf

	If ($iDecChar <> Null) Then
		If Not IsInt($iDecChar) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$tTabStruct.DecimalChar = $iDecChar
	EndIf

	If ($atTabStops[0].Alignment() = $LOW_TAB_ALIGN_DEFAULT) And (UBound($atTabStops) = 1) Then ; if inserting a  Tabstop for the first time, overwrite the "Default blank TabStop.
		$atTabStops[0] = $tTabStruct
		$oObj.ParaTabStops = $atTabStops ;insert the new TabStop
		$atNewTabStops = $oObj.ParaTabStops()
		$tFoundTabStop = $atNewTabStops[0]
		$iNewPosition = $tFoundTabStop.Position()
	Else
		__LOWriter_AddTo1DArray($atTabStops, $tTabStruct)

		$aiTabList = __LOWriter_ParTabStopList($oObj) ; Get a list of existing tabstops to compare with
		If Not IsArray($aiTabList) Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)
		__LOWriter_AddTo1DArray($aiTabList, 0) ; Add a dummy to make Array sizes equal.

		$oObj.ParaTabStops = $atTabStops ;insert the new TabStop

		$atNewTabStops = $oObj.ParaTabStops() ; now retrieve a new list to find the final Tab Stop position.
		For $i = 0 To UBound($atNewTabStops) - 1
			If ($atNewTabStops[$i].Position()) <> $aiTabList[$i] Then
				$iNewPosition = $atNewTabStops[$i].Position()
				$tFoundTabStop = $atNewTabStops[$i]
				$bFound = True
				ExitLoop
			EndIf
		Next

		If Not $bFound Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0) ; didn't find the new TabStop
	EndIf

	$iError = (__LOWriter_NumIsBetween(($tFoundTabStop.Position()), ($iPosition - 1), ($iPosition + 1))) ? $iError : BitOR($iError, 1)
	$iError = ($iFillChar = Null) ? $iError : ($tFoundTabStop.FillChar = $iFillChar) ? $iError : BitOR($iError, 2)
	$iError = ($iAlignment = Null) ? $iError : ($tFoundTabStop.Alignment = $iAlignment) ? $iError : BitOR($iError, 4)
	$iError = ($iDecChar = Null) ? $iError : ($tFoundTabStop.DecimalChar = $iDecChar) ? $iError : BitOR($iError, 8)

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, $iNewPosition) : SetError($__LOW_STATUS_SUCCESS, 0, $iNewPosition)
EndFunc   ;==>__LOWriter_ParTabStopCreate

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParTabStopDelete
; Description ...: Delete a TabStop from a Paragraph
; Syntax ........: __LOWriter_ParTabStopDelete(Byref $oObj, $iTabStop)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iTabStop            - an integer value. The Tab position of the TabStop to modify. See Remarks.
; Return values .: Success: Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 6 Return 0 = Passed Object to internal function not an Object.
;				   @Error 1 @Extended 7 Return 0 = Passed Document Object to internal function not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving ParaTabStops Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to identify and delete TabStop in Paragraph.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = Returns true if TabStop was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iTabStop refers to the position, or essential the "length" of a TabStop from the edge of a page
;						margin. This is the only reliable way to identify a Tabstop to be able to interact with it, as there
;						 can only be one of a certain length per document.
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

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)

	$atOldTabStops = $oObj.ParaTabStops()
	ReDim $atNewTabStops[UBound($atOldTabStops)]
	If Not IsArray($atOldTabStops) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If (UBound($atOldTabStops) = 1) Then
		$oDefaults = $oDoc.createInstance("com.sun.star.text.Defaults")
		$tTabStruct = $atOldTabStops[0]
		$tTabStruct.Alignment = $LOW_TAB_ALIGN_DEFAULT
		$tTabStruct.Position = $oDefaults.TabStopDistance()
		$tTabStruct.FillChar = 32 ;Space
		$tTabStruct.DecimalChar = 46 ;Period
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
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next
	EndIf
	ReDim $atNewTabStops[(($bDeleted) ? (UBound($atNewTabStops) - 1) : UBound($atNewTabStops))]

	$oObj.ParaTabStops = $atNewTabStops

	Return ($bDeleted) ? SetError($__LOW_STATUS_SUCCESS, 0, $bDeleted) : SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
EndFunc   ;==>__LOWriter_ParTabStopDelete

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParTabStopList
; Description ...: Retrieve a List of TabStops available in a Paragraph.
; Syntax ........: __LOWriter_ParTabStopList(Byref $oObj)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
; Return values .: Success: Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving ParaTabStops Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. An Array of TabStops. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParTabStopList(ByRef $oObj)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $atTabStops[0]
	Local $aiTabList[0]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	$atTabStops = $oObj.ParaTabStops()
	If Not IsArray($atTabStops) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	ReDim $aiTabList[UBound($atTabStops)]

	For $i = 0 To UBound($atTabStops) - 1
		$aiTabList[$i] = $atTabStops[$i].Position()
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
	Next
	Return SetError($__LOW_STATUS_SUCCESS, $i, $aiTabList)
EndFunc   ;==>__LOWriter_ParTabStopList

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParTabStopMod
; Description ...: Modify or retrieve the properties of an existing TabStop.
; Syntax ........: __LOWriter_ParTabStopMod(Byref $oObj, $iTabStop, $iPosition, $iFillChar, $iAlignment, $iDecChar)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $iTabStop            - an integer value. The Tab position of the TabStop to modify. See Remarks.
;                  $iPosition           - an integer value. The New position to set the input position to. Set in Micrometers (uM). See Remarks.
;                  $iFillChar           - an integer value. The Asc (see autoit function) value of any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
;                  $iAlignment          - an integer value. The position of where the end of a Tab is aligned to compared to the text. See Constants, $LOW_TAB_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iDecChar            - an integer value. Enter a character(in Asc Value(See Autoit Function)) that you want the decimal tab to use as a decimal separator. Can only be set if $iAlignment is set to $LOW_TAB_ALIGN_DECIMAL.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 5 Return 0 = Passed Object to internal function not an Object.
;				   @Error 1 @Extended 6 Return 0 = $iPosition not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iFillChar not an Integer.
;				   @Error 1 @Extended 8 Return 0 = $iAlignment not an Integer, less than 0 or greater than 4. See Constants, $LOW_TAB_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 9 Return 0 = $iDecChar not an Integer.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving ParaTabStops Object.
;				   @Error 2 @Extended 2 Return 0 = Error retrieving Requested TabStop Object.
;				   @Error 2 @Extended 3 Return 0 = Error retrieving list of TabStop Positions.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Paragraph style already contains a TabStop at the length/Position specified in $iPosition.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iPosition
;				   |								2 = Error setting $iFillChar
;				   |								4 = Error setting $iAlignment
;				   |								8 = Error setting $iDecChar
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
;				   @Error 0 @Extended ? Return 2 = Success. Settings were successfully set. New TabStop position is returned in @Extended.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iTabStop refers to the position, or essential the "length" of a TabStop from the edge of a page margin.
;						This is the only reliable way to identify a Tabstop to be able to interact with it, as there can only be
;						one of a certain length per document.
;					Note: $iPosition once set can vary +/- 1 uM. To ensure you can identify the tabstop to modify it again,
;						This function returns the new TabStop position in @Extended when $iPosition is set, return value will
;						be set to 2. See Return Values.
;					Note: Since $iPosition can fluctuate +/- 1 uM when it is inserted into LibreOffice, it is possible to
;						accidentally overwrite an already existing TabStop.
;					Note: $iFillChar, Libre's Default value, "None" is in reality a space character which is Asc value 32.
;						The other values offered by Libre are: Period (ASC 46), Dash (ASC 45) and Underscore (ASC 95). You can
;						also enter a custom ASC value. See ASC Autoit Func. and "ASCII Character Codes" in the Autoit help file.
;					 Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer
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

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
	$atTabStops = $oObj.ParaTabStops()
	If Not IsArray($atTabStops) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	For $i = 0 To UBound($atTabStops) - 1
		$tTabStruct = ($atTabStops[$i].Position() = $iTabStop) ? $atTabStops[$i] : Null
		If IsObj($tTabStruct) Then ExitLoop
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
	Next
	If Not IsObj($tTabStruct) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($iPosition, $iFillChar, $iAlignment, $iDecChar) Then
		__LOWriter_ArrayFill($aiTabSettings, $tTabStruct.Position(), $tTabStruct.FillChar(), $tTabStruct.Alignment(), $tTabStruct.DecimalChar())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $aiTabSettings)
	EndIf

	If ($iPosition <> Null) Then
		If Not IsInt($iPosition) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		If __LOWriter_ParHasTabStop($oObj, $iPosition) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)
		$tTabStruct.Position = $iPosition
		$iError = ($tTabStruct.Position() = $iPosition) ? $iError : BitOR($iError, 1)
		$bNewPosition = True
	EndIf

	If ($iFillChar <> Null) Then
		If Not IsInt($iFillChar) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$tTabStruct.FillChar = $iFillChar
		$tTabStruct.FillChar = ($tTabStruct.FillChar() = 0) ? 32 : $tTabStruct.FillChar()
		$iError = ($tTabStruct.FillChar = $iFillChar) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iAlignment <> Null) Then
		If Not __LOWriter_IntIsBetween($iAlignment, $LOW_TAB_ALIGN_LEFT, $LOW_TAB_ALIGN_DEFAULT) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 8, 0)
		$tTabStruct.Alignment = $iAlignment
		$iError = ($tTabStruct.Alignment = $iAlignment) ? $iError : BitOR($iError, 4)
	EndIf

	If ($iDecChar <> Null) Then
		If Not IsInt($iDecChar) And ($iDecChar <> Null) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 9, 0)
		$tTabStruct.DecimalChar = $iDecChar
		$iError = ($tTabStruct.DecimalChar = $iDecChar) ? $iError : BitOR($iError, 8)
	EndIf

	$atTabStops[$i] = $tTabStruct

	If $bNewPosition Then
		$aiTabList = __LOWriter_ParTabStopList($oObj)
		If Not IsArray($aiTabList) Then Return SetError($__LOW_STATUS_INIT_ERROR, 3, 0)
	EndIf

	$oObj.ParaTabStops = $atTabStops

	If $bNewPosition Then
		$atNewTabStops = $oObj.ParaTabStops()
		For $j = 0 To UBound($atNewTabStops) - 1
			If ($atNewTabStops[$j].Position()) <> $aiTabList[$j] Then
				$iNewPosition = $atNewTabStops[$j].Position()
				ExitLoop
			EndIf
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
		Next
		Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, $iNewPosition, 2)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_ParTabStopMod

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ParTxtFlowOpt
; Description ...: Set and Retrieve Text Flow settings.
; Syntax ........: __LOWriter_ParTxtFlowOpt(Byref $oObj, $bParSplit, $bKeepTogether, $iParOrphans, $iParWidows)
; Parameters ....: $oObj                - [in/out] an object. Paragraph Style Object or a Cursor or Paragraph Object.
;                  $bParSplit           - a boolean value. False prevents the paragraph from getting split into two pages or columns
;                  $bKeepTogether       - a boolean value. True prevents page or column breaks between this and the following paragraph
;                  $iParOrphans         - an integer value. Specifies the minimum number of lines of the paragraph that have to be at bottom of a page if the paragraph is spread over more than one page. Min is 0 (disabled), and cannot be 1. Max is 9.
;                  $iParWidows          - an integer value. Specifies the minimum number of lines of the paragraph that have to be at top of a page if the paragraph is spread over more than one page. Min is 0 (disabled), and cannot be 1. Max is 9.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bParSplit not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bKeepTogether not  a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iParOrphans not an Integer, less than 0, equal to 1, or greater than 9.
;				   @Error 1 @Extended 7 Return 0 = $iParWidows not an Integer, less than 0, equal to 1, or greater than 9.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bParSplit
;				   |								2 = Error setting $bKeepTogether
;				   |								4 = Error setting $iParOrphans
;				   |								8 = Error setting $iParWidows
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:Note: If you do not set ParSplit to True, the rest of the settings will still show to have been set but will
;					not become active until $bParSplit is set to true.
;					Call this function with only the Object parameter and all other parameters set to Null keyword, to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_ParTxtFlowOpt(ByRef $oObj, $bParSplit, $bKeepTogether, $iParOrphans, $iParWidows)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avTxtFlowOpt[4]

	If Not IsObj($oObj) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)

	If __LOWriter_VarsAreNull($bParSplit, $bKeepTogether, $iParOrphans, $iParWidows) Then
		__LOWriter_ArrayFill($avTxtFlowOpt, $oObj.ParaSplit(), $oObj.ParaKeepTogether(), $oObj.ParaOrphans(), $oObj.ParaWidows())
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avTxtFlowOpt)
	EndIf

	If ($bParSplit <> Null) Then
		If Not IsBool($bParSplit) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
		$oObj.ParaSplit = $bParSplit
		$iError = ($oObj.ParaSplit = $bParSplit) ? $iError : BitOR($iError, 1)
	EndIf

	If ($bKeepTogether <> Null) Then
		If Not IsBool($bKeepTogether) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)
		$oObj.ParaKeepTogether = $bKeepTogether
		$iError = ($oObj.ParaKeepTogether = $bKeepTogether) ? $iError : BitOR($iError, 2)
	EndIf

	If ($iParOrphans <> Null) Then
		If Not __LOWriter_IntIsBetween($iParOrphans, 0, 9, 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 6, 0)
		$oObj.ParaOrphans = $iParOrphans
		$iError = ($oObj.ParaOrphans = $iParOrphans) ? $iError : BitOR($iError, 4)
	EndIf

	If ($iParWidows <> Null) Then
		If Not __LOWriter_IntIsBetween($iParWidows, 0, 9, 1) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 7, 0)
		$oObj.ParaWidows = $iParWidows
		$iError = ($oObj.ParaWidows = $iParWidows) ? $iError : BitOR($iError, 8)
	EndIf

	Return ($iError > 0) ? SetError($__LOW_STATUS_PROP_SETTING_ERROR, $iError, 0) : SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_ParTxtFlowOpt

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_RegExpConvert
; Description ...: Convert a Libre Office Regular Expression for use in AutoIt RegExpReplace.
; Syntax ........: __LOWriter_RegExpConvert(Byref $sRegExpString)
; Parameters ....: $sRegExpString       - [in/out] a string value. The L.O. Regular Expression string. String will be directly modified.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $sRegExpString not a String.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. String was successfully converted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_RegExpConvert(ByRef $sRegExpString)
	Local $iPos1 = 0, $iPos2, $iReplacements
	Local $sRegExpStringTemp = $sRegExpString, $sBackSlashFlag = "~*#!DH13!#*~"
	Local $STR_NOCASESENSE = 0

	If Not IsString($sRegExpString) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$sRegExpStringTemp = StringReplace($sRegExpStringTemp, "\\", $sBackSlashFlag) ; Temporarily replace all double backslashes with a personal pattern to replace them later.
	$iReplacements = @extended ; Capture number of replacements
	Do
		$iPos1 = StringInStr($sRegExpStringTemp, "&", $STR_NOCASESENSE, 1, $iPos1 + 1) ; test for & found in string, starting at last found position + 1.
		If ($iPos1 > 0) Then ;  if there is a find, begin testing for backslash.
			$iPos2 = StringInStr($sRegExpStringTemp, "\", $STR_NOCASESENSE, 1, $iPos1 - 1, 1) ; Test for a backslash, if there is one, then the & is not to be replaced.
			; If there is no backslash, then replace the & with Autoit's accepted back reference, $0
			If ($iPos2 = 0) Then $sRegExpStringTemp = StringLeft($sRegExpStringTemp, $iPos1 - 1) & "${0}" & StringMid($sRegExpStringTemp, $iPos1 + 1)
		EndIf
	Until $iPos1 = 0

	$sRegExpStringTemp = StringReplace($sRegExpStringTemp, "\n", @CR) ; Replace L.O. keyword for New Par. line for Autoit Carriage return.
	$sRegExpStringTemp = StringReplace($sRegExpStringTemp, "\t", @TAB) ; Replace L.O. keyword for Tab for Autoit Tab.

	If ($iReplacements > 0) Then $sRegExpStringTemp = StringReplace($sRegExpStringTemp, $sBackSlashFlag, "\\") ; Replace the Flag with literal double Backslashes.

	$sRegExpString = $sRegExpStringTemp ; Update the Replacement String.
	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_RegExpConvert

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_SetPropertyValue
; Description ...: Creates a property value struct object.
; Syntax ........: __LOWriter_SetPropertyValue($sName, $vValue)
; Parameters ....: $sName               - a string value. Property name.
;                  $vValue              - a variant value. Property value.
; Return values .:Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = Property $sName Value was not a string
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Properties Object failed to be created
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Property Object Returned
; Author ........: Leagnus, GMK
; Modified ......: donnyh13 - added CreateStruct function. Modified variable names.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_SetPropertyValue($sName, $vValue)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tProperties

	If Not IsString($sName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	$tProperties = __LOWriter_CreateStruct("com.sun.star.beans.PropertyValue")
	If @error Or Not IsObj($tProperties) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	$tProperties.Name = $sName
	$tProperties.Value = $vValue

	Return SetError($__LOW_STATUS_SUCCESS, 0, $tProperties)
EndFunc   ;==>__LOWriter_SetPropertyValue

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_TableBorder
; Description ...: Set or Retrieve Table Border settings -- internal function. Libre Office 3.6 and Up.
; Syntax ........: __LOWriter_TableBorder(Byref $oTable, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight, $iVert, $iHori)
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned from _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName functions.
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
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oTable Variable not Object type variable.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;				   @Error 2 @Extended 2 Return 0 = Error retrieving Object "TableBorder2".
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style/Color when Border width not set.
;				   @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Style/Color when Border width not set.
;				   @Error 4 @Extended 3 Return 0 = Cannot set Left Border Style/Color when Border width not set.
;				   @Error 4 @Extended 4 Return 0 = Cannot set Right Border Style/Color when Border width not set.
;				   @Error 4 @Extended 5 Return 0 = Cannot set Vertical Border Style/Color when Border width not set.
;				   @Error 4 @Extended 6 Return 0 = Cannot set Horizontal Border Style/Color when Border width not set.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the Table Object, and either $bWid, $bSty, or $bCol set to true, with all other parameters set to Null keyword, to get the current settings.
;					Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_TableBorder(ByRef $oTable, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight, $iVert, $iHori)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avBorder[6]
	Local $tBL2, $tTB2

	If Not __LOWriter_VersionCheck(3.6) Then Return SetError($__LOW_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oTable) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If (($bWid + $bSty + $bCol) <> 1) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then

		If $bWid Then
			__LOWriter_ArrayFill($avBorder, $oTable.TableBorder2.TopLine.LineWidth(), $oTable.TableBorder2.BottomLine.LineWidth(), _
					$oTable.TableBorder2.LeftLine.LineWidth(), $oTable.TableBorder2.RightLine.LineWidth(), $oTable.TableBorder2.VerticalLine.LineWidth(), _
					$oTable.TableBorder2.HorizontalLine.LineWidth())
		ElseIf $bSty Then
			__LOWriter_ArrayFill($avBorder, $oTable.TableBorder2.TopLine.LineStyle(), $oTable.TableBorder2.BottomLine.LineStyle(), _
					$oTable.TableBorder2.LeftLine.LineStyle(), $oTable.TableBorder2.RightLine.LineStyle(), $oTable.TableBorder2.VerticalLine.LineStyle(), _
					$oTable.TableBorder2.HorizontalLine.LineStyle())
		ElseIf $bCol Then
			__LOWriter_ArrayFill($avBorder, $oTable.TableBorder2.TopLine.Color(), $oTable.TableBorder2.BottomLine.Color(), _
					$oTable.TableBorder2.LeftLine.Color(), $oTable.TableBorder2.RightLine.Color(), $oTable.TableBorder2.VerticalLine.Color(), _
					$oTable.TableBorder2.HorizontalLine.Color())
		EndIf
		Return SetError($__LOW_STATUS_SUCCESS, 1, $avBorder)
	EndIf

	$tBL2 = __LOWriter_CreateStruct("com.sun.star.table.BorderLine2")
	$tTB2 = $oTable.TableBorder2
	If Not IsObj($tBL2) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)
	If Not IsObj($tTB2) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	If $iTop <> Null Then
		If Not $bWid And ($tTB2.TopLine.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 1, 0) ; If Width not set, cant set color or style.
		; Top Line
		$tBL2.LineWidth = ($bWid) ? $iTop : $tTB2.TopLine.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iTop : $tTB2.TopLine.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iTop : $tTB2.TopLine.Color() ; copy Color over to new size structure
		$tTB2.TopLine = $tBL2
	EndIf

	If $iBottom <> Null Then
		If Not $bWid And ($tTB2.BottomLine.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 2, 0) ; If Width not set, cant set color or style.
		; Bottom Line
		$tBL2.LineWidth = ($bWid) ? $iBottom : $tTB2.BottomLine.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iBottom : $tTB2.BottomLine.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iBottom : $tTB2.BottomLine.Color() ; copy Color over to new size structure
		$tTB2.BottomLine = $tBL2
	EndIf

	If $iLeft <> Null Then
		If Not $bWid And ($tTB2.LeftLine.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 3, 0) ; If Width not set, cant set color or style.
		; Left Line
		$tBL2.LineWidth = ($bWid) ? $iLeft : $tTB2.LeftLine.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iLeft : $tTB2.LeftLine.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iLeft : $tTB2.LeftLine.Color() ; copy Color over to new size structure
		$tTB2.LeftLine = $tBL2
	EndIf

	If $iRight <> Null Then
		If Not $bWid And ($tTB2.RightLine.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 4, 0) ; If Width not set, cant set color or style.
		; Right Line
		$tBL2.LineWidth = ($bWid) ? $iRight : $tTB2.RightLine.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iRight : $tTB2.RightLine.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iRight : $tTB2.RightLine.Color() ; copy Color over to new size structure
		$tTB2.RightLine = $tBL2
	EndIf

	If $iVert <> Null Then
		If Not $bWid And ($tTB2.VerticalLine.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 5, 0) ; If Width not set, cant set color or style.
		; Vertical Line
		$tBL2.LineWidth = ($bWid) ? $iVert : $tTB2.VerticalLine.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iVert : $tTB2.VerticalLine.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iVert : $tTB2.VerticalLine.Color() ; copy Color over to new size structure
		$tTB2.VerticalLine = $tBL2
	EndIf

	If $iHori <> Null Then
		If Not $bWid And ($tTB2.HorizontalLine.LineWidth() = 0) Then Return SetError($__LOW_STATUS_PROP_SETTING_ERROR, 6, 0) ; If Width not set, cant set color or style.
		; Horizontal Line
		$tBL2.LineWidth = ($bWid) ? $iHori : $tTB2.HorizontalLine.LineWidth() ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? $iHori : $tTB2.HorizontalLine.LineStyle() ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? $iHori : $tTB2.HorizontalLine.Color() ; copy Color over to new size structure
		$tTB2.HorizontalLine = $tBL2
	EndIf

	$oTable.TableBorder2 = $tTB2

	Return SetError($__LOW_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOWriter_TableBorder

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_TableCursorMove
; Description ...: Text-TableCursor related movements.
; Syntax ........: __LOWriter_TableCursorMove(Byref $oCursor, $iMove, $iCount[, $bSelect = False])
; Parameters ....: $oCursor             - [in/out] an object. A TableCursor Object returned from _LOWriter_TableCreateCursor function.
;                  $iMove               - an Integer value. The movement command. See remarks and Constants.
;                  $iCount              - an integer value. Number of movements to make.
;                  $bSelect             - [optional] a boolean value. Default is False. Whether to select data during this cursor movement.
; Return values .: Success: Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iMove not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iMove mismatch with Cursor type. See Cursor Type/Move Type Constants.
;				   @Error 1 @Extended 4 Return 0 = $iCount not an integer or is a negative.
;				   @Error 1 @Extended 5 Return 0 = $bSelect not a Boolean.
;				   --Processing Errors--
;				   @Error 3 @Extended 2 Return 0 = Error processing cursor move.
;				   --Success--
;				   @Error 0 @Extended ? Return Boolean = Success, Cursor object movement was processed successfully.
;				   +				Returns True if the full count of movements were successful, else false if none or only partially successful.
;				   +				@Extended set to number of successful movements. Or Page Number for "gotoPage" command. See Remarks
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iMove may be set to any of the following constants depending on the Cursor type you are intending to move.
;					 Only some movements accept movement amounts (such as "goRight" 2) etc. Also only some accept creating/
;					extending a selection of text/ data. They will be specified below. To Clear /Unselect a current selection,
;					you can input a move such as "goRight", 0, False.
; Cursor Movement Constants:
;					#Cursor Movement Constants which accept number of Moves and Selecting:
;					-ViewCursor
;						$LOW_VIEWCUR_GO_DOWN, Move the cursor Down by n lines.
;						$LOW_VIEWCUR_GO_UP, Move the cursor Up by n lines.
;						$LOW_VIEWCUR_GO_LEFT, Move the cursor left by n characters.
;						$LOW_VIEWCUR_GO_RIGHT, Move the cursor right by n characters.
;					-TextCursor
;						$LOW_TEXTCUR_GO_LEFT,Move the cursor left by n characters.
;						$LOW_TEXTCUR_GO_RIGHT, Move the cursor right by n characters.
;						$LOW_TEXTCUR_GOTO_NEXT_WORD, Move to the start of the next word.
;						$LOW_TEXTCUR_GOTO_PREV_WORD, Move to the end of the previous word.
;						$LOW_TEXTCUR_GOTO_NEXT_SENTENCE,Move to the start of the next sentence.
;						$LOW_TEXTCUR_GOTO_PREV_SENTENCE, Move to the end of the previous sentence.
;						$LOW_TEXTCUR_GOTO_NEXT_PARAGRAPH, Move to the start of the next paragraph.
;						$LOW_TEXTCUR_GOTO_PREV_PARAGRAPH, Move to the End of the previous paragraph.
;					-TableCursor
;						$LOW_TABLECUR_GO_LEFT, Move the cursor left/right n cells.
;						$LOW_TABLECUR_GO_RIGHT, Move the cursor left/right n cells.
;						$LOW_TABLECUR_GO_UP,  Move the cursor up/down n cells.
;						$LOW_TABLECUR_GO_DOWN, Move the cursor up/down n cells.
;					#Cursor Movements which accept number of Moves Only:
;					-ViewCursor
;						$LOW_VIEWCUR_JUMP_TO_NEXT_PAGE, Move the cursor to the Next page.
;						$LOW_VIEWCUR_JUMP_TO_PREV_PAGE, Move the cursor to the previous page.
;						$LOW_VIEWCUR_SCREEN_DOWN, Scroll the view forward by one visible page.
;						$LOW_VIEWCUR_SCREEN_UP, Scroll the view back by one visible page.
;					#Cursor Movements which accept Selecting Only:
;					-ViewCursor
;						$LOW_VIEWCUR_GOTO_END_OF_LINE, Move the cursor to the end of the current line.
;						$LOW_VIEWCUR_GOTO_START_OF_LINE, Move the cursor to the start of the current line.
;						$LOW_VIEWCUR_GOTO_START, Move the cursor to the start of the document or Table.
;						$LOW_VIEWCUR_GOTO_END, Move the cursor to the end of the document or Table.
;					-TextCursor
;						$LOW_TEXTCUR_GOTO_START, Move the cursor to the start of the text.
;						$LOW_TEXTCUR_GOTO_END, Move the cursor to the end of the text.
;						$LOW_TEXTCUR_GOTO_END_OF_WORD, Move to the end of the current word.
;						$LOW_TEXTCUR_GOTO_START_OF_WORD, Move to the start of the current word.
;						$LOW_TEXTCUR_GOTO_END_OF_SENTENCE, Move to the end of the current sentence.
;						$LOW_TEXTCUR_GOTO_START_OF_SENTENCE, Move to the start of the current sentence.
;						$LOW_TEXTCUR_GOTO_END_OF_PARAGRAPH, Move to the end of the current paragraph.
;						$LOW_TEXTCUR_GOTO_START_OF_PARAGRAPH, Move to the start of the current paragraph.
;					-TableCursor
;						$LOW_TABLECUR_GOTO_START, Move the cursor to the top left cell.
;						$LOW_TABLECUR_GOTO_END,  Move the cursor to the bottom right cell.
;					#Cursor Movements which accept nothing and are done once per call:
;					-ViewCursor
;						$LOW_VIEWCUR_JUMP_TO_FIRST_PAGE, Move the cursor to the first page.
;						$LOW_VIEWCUR_JUMP_TO_LAST_PAGE, Move the cursor to the Last page.
;						$LOW_VIEWCUR_JUMP_TO_END_OF_PAGE, Move the cursor to the end of the current page.
;						$LOW_VIEWCUR_JUMP_TO_START_OF_PAGE, Move the cursor to the start of the current page.
;					-TextCursor
;						$LOW_TEXTCUR_COLLAPSE_TO_START,
;						$LOW_TEXTCUR_COLLAPSE_TO_END (Collapses the current selection and moves the cursor  to start or End of selection.
;					#Misc. Cursor Movements:
;					-ViewCursor
;						$LOW_VIEWCUR_JUMP_TO_PAGE (accepts page number to jump to in $iCount, Returns what page was successfully jumped to.
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

	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iMove) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If ($iMove >= UBound($asMoves)) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iCount) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bSelect) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	Switch $iMove
		Case $LOW_TABLECUR_GO_LEFT, $LOW_TABLECUR_GO_RIGHT, $LOW_TABLECUR_GO_UP, $LOW_TABLECUR_GO_DOWN
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "(" & $iCount & "," & $bSelect & ")")
			$iCounted = ($bMoved) ? $iCount : 0
			Return SetError($__LOW_STATUS_SUCCESS, $iCounted, $bMoved)

		Case $LOW_TABLECUR_GOTO_START, $LOW_TABLECUR_GOTO_END
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "(" & $bSelect & ")")
			$iCounted = ($bMoved) ? 1 : 0
			Return SetError($__LOW_STATUS_SUCCESS, $iCounted, $bMoved)
		Case Else
			Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0)
	EndSwitch
EndFunc   ;==>__LOWriter_TableCursorMove

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_TableHasCellName
; Description ...: Check whether the Table contains a Cell by the requested name.
; Syntax ........: __LOWriter_TableHasCellName(Byref $oTable, Byref $sCellName)
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned from _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName functions.
;                  $sCellName           - [in/out] a string value. The requested cell name.
; Return values .: Success: Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oTable variable not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sCellName variable not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Cell Names.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean: If  the table contains the requested Cell Name, True is returned. Else False.
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

	If Not IsObj($oTable) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sCellName) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$aCellNames = $oTable.getCellNames()
	If Not IsArray($aCellNames) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	For $i = 0 To UBound($aCellNames) - 1
		If StringInStr($aCellNames[$i], $sCellName) Then Return SetError($__LOW_STATUS_SUCCESS, 0, True)
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
	Next

	Return SetError($__LOW_STATUS_SUCCESS, 0, False) ; Cell not found
EndFunc   ;==>__LOWriter_TableHasCellName

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_TableHasColumnRange
; Description ...: Check if Table contains the requested Column.
; Syntax ........: __LOWriter_TableHasColumnRange(Byref $oTable, Byref $iColumn)
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned from _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName functions.
;                  $iColumn             - [in/out] an integer value. The requested Column.
; Return values .: Success: Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oTable variable not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iColumn variable not an Integer.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean: If True, the table contains the requested Column. Else False.
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

	If Not IsObj($oTable) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iColumn) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	Return ($iColumn <= ($oTable.getColumns.getCount() - 1)) ? True : False
EndFunc   ;==>__LOWriter_TableHasColumnRange

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_TableHasRowRange
; Description ...: Check if Table contains the requested row.
; Syntax ........: __LOWriter_TableHasRowRange(Byref $oTable, Byref $iRow)
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned from _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName functions.
;                  $iRow                - [in/out] an integer value. The requested row.
; Return values .: Success: Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oTable variable not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iRow variable not an Integer.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean: If True, the table contains the requested row. Else False.
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

	If Not IsObj($oTable) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iRow) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	Return ($iRow <= ($oTable.getRows.getCount() - 1)) ? True : False
EndFunc   ;==>__LOWriter_TableHasRowRange

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_TableRowSplitToggle
; Description ...: Set or Retrieve Table Row split setting for an entire Table.
; Syntax ........: __LOWriter_TableRowSplitToggle(Byref $oTable[, $bSplitRows = Null])
; Parameters ....: $oTable              - [in/out] an object. A Table Object returned from _LOWriter_TableCreate, _LOWriter_TableGetObjByCursor, or _LOWriter_TableGetObjByName functions.
;                  $bSplitRows          - [optional] a boolean value. Default is Null. If True, the content in a Table row is allowed to split at page splits, else if False, Content is not allowed to split across pages.
; Return values .: Success: Integer or Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oTable not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bSplitRows not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Table's Row count.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve first Row's current split row setting.
;				   --Success--
;				   @Error 0 @Extended 0 Return 0 = Success. All optional parameters were set to Null, Table Rows have multiple SplitRow settings, returning 0 to indicate this.
;				   @Error 0 @Extended 1 Return Boolean = Success. All optional parameters were set to Null, returning current split row setting as a Boolean.
;				   @Error 0 @Extended 2 Return 1 = Success. Setting was successfully set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to
;					get the current settings.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_TableRowSplitToggle(ByRef $oTable, $bSplitRows = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iRows
	Local $bSplitRowTest

	If Not IsObj($oTable) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)

	$iRows = $oTable.getRows.getCount()
	If Not IsInt($iRows) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	If ($bSplitRows = Null) Then ; Retrieve Split Rows Setting

		; Retrieve the First Row's Split Row setting.
		$bSplitRowTest = $oTable.getRows.getByIndex(0).IsSplitAllowed()
		If Not IsBool($bSplitRowTest) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

		For $i = 1 To $iRows - 1

			If $bSplitRowTest <> ($oTable.getRows.getByIndex($i).IsSplitAllowed()) Then Return SetError($__LOW_STATUS_SUCCESS, 0, 0) ; Table Rows have mixed settings, return 0.
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
		Next
		Return SetError($__LOW_STATUS_SUCCESS, 1, $bSplitRowTest) ; All Table Rows are set the same as the first Row, return that setting.

	Else ;Set Split Rows Setting
		If Not IsBool($bSplitRows) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

		For $i = 0 To $iRows - 1
			$oTable.getRows.getByIndex($i).IsSplitAllowed = $bSplitRows
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
		Next
		Return SetError($__LOW_STATUS_SUCCESS, 2, 1)
	EndIf

EndFunc   ;==>__LOWriter_TableRowSplitToggle

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_TextCursorMove
; Description ...: For TextCursor related movements.
; Syntax ........: __LOWriter_TextCursorMove(Byref $oCursor, $iMove, $iCount[, $bSelect = False])
; Parameters ....: $oCursor             - [in/out] an object. A TextCursor Object returned from any TextCursor creation functions.
;                  $iMove               - an Integer value. The movement command. See remarks and Constants.
;                  $iCount              - an integer value. Number of movements to make.
;                  $bSelect             - [optional] a boolean value. Default is False. Whether to select data during this cursor movement.
; Return values .: Success: Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iMove not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iMove mismatch with Cursor type. See Cursor Type/Move Type Constants.
;				   @Error 1 @Extended 4 Return 0 = $iCount not an integer or is a negative.
;				   @Error 1 @Extended 5 Return 0 = $bSelect not a Boolean.
;				   --Processing Errors--
;				   @Error 3 @Extended 2 Return 0 = Error processing cursor move.
;				   --Success--
;				   @Error 0 @Extended ? Return Boolean = Success, Cursor object movement was processed successfully.
;				   +				Returns True if the full count of movements were successful, else false if none or only partially successful.
;				   +				@Extended set to number of successful movements. Or Page Number for "gotoPage" command. See Remarks
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iMove may be set to any of the following constants depending on the Cursor type you are intending to move.
;					 Only some movements accept movement amounts (such as "goRight" 2) etc. Also only some accept creating/
;					extending a selection of text/ data. They will be specified below. To Clear /Unselect a current selection,
;					you can input a move such as "goRight", 0, False.
; Cursor Movement Constants:
;					#Cursor Movement Constants which accept number of Moves and Selecting:
;					-ViewCursor
;						$LOW_VIEWCUR_GO_DOWN, Move the cursor Down by n lines.
;						$LOW_VIEWCUR_GO_UP, Move the cursor Up by n lines.
;						$LOW_VIEWCUR_GO_LEFT, Move the cursor left by n characters.
;						$LOW_VIEWCUR_GO_RIGHT, Move the cursor right by n characters.
;					-TextCursor
;						$LOW_TEXTCUR_GO_LEFT,Move the cursor left by n characters.
;						$LOW_TEXTCUR_GO_RIGHT, Move the cursor right by n characters.
;						$LOW_TEXTCUR_GOTO_NEXT_WORD, Move to the start of the next word.
;						$LOW_TEXTCUR_GOTO_PREV_WORD, Move to the end of the previous word.
;						$LOW_TEXTCUR_GOTO_NEXT_SENTENCE,Move to the start of the next sentence.
;						$LOW_TEXTCUR_GOTO_PREV_SENTENCE, Move to the end of the previous sentence.
;						$LOW_TEXTCUR_GOTO_NEXT_PARAGRAPH, Move to the start of the next paragraph.
;						$LOW_TEXTCUR_GOTO_PREV_PARAGRAPH, Move to the End of the previous paragraph.
;					-TableCursor
;						$LOW_TABLECUR_GO_LEFT, Move the cursor left/right n cells.
;						$LOW_TABLECUR_GO_RIGHT, Move the cursor left/right n cells.
;						$LOW_TABLECUR_GO_UP,  Move the cursor up/down n cells.
;						$LOW_TABLECUR_GO_DOWN, Move the cursor up/down n cells.
;					#Cursor Movements which accept number of Moves Only:
;					-ViewCursor
;						$LOW_VIEWCUR_JUMP_TO_NEXT_PAGE, Move the cursor to the Next page.
;						$LOW_VIEWCUR_JUMP_TO_PREV_PAGE, Move the cursor to the previous page.
;						$LOW_VIEWCUR_SCREEN_DOWN, Scroll the view forward by one visible page.
;						$LOW_VIEWCUR_SCREEN_UP, Scroll the view back by one visible page.
;					#Cursor Movements which accept Selecting Only:
;					-ViewCursor
;						$LOW_VIEWCUR_GOTO_END_OF_LINE, Move the cursor to the end of the current line.
;						$LOW_VIEWCUR_GOTO_START_OF_LINE, Move the cursor to the start of the current line.
;						$LOW_VIEWCUR_GOTO_START, Move the cursor to the start of the document or Table.
;						$LOW_VIEWCUR_GOTO_END, Move the cursor to the end of the document or Table.
;					-TextCursor
;						$LOW_TEXTCUR_GOTO_START, Move the cursor to the start of the text.
;						$LOW_TEXTCUR_GOTO_END, Move the cursor to the end of the text.
;						$LOW_TEXTCUR_GOTO_END_OF_WORD, Move to the end of the current word.
;						$LOW_TEXTCUR_GOTO_START_OF_WORD, Move to the start of the current word.
;						$LOW_TEXTCUR_GOTO_END_OF_SENTENCE, Move to the end of the current sentence.
;						$LOW_TEXTCUR_GOTO_START_OF_SENTENCE, Move to the start of the current sentence.
;						$LOW_TEXTCUR_GOTO_END_OF_PARAGRAPH, Move to the end of the current paragraph.
;						$LOW_TEXTCUR_GOTO_START_OF_PARAGRAPH, Move to the start of the current paragraph.
;					-TableCursor
;						$LOW_TABLECUR_GOTO_START, Move the cursor to the top left cell.
;						$LOW_TABLECUR_GOTO_END,  Move the cursor to the bottom right cell.
;					#Cursor Movements which accept nothing and are done once per call:
;					-ViewCursor
;						$LOW_VIEWCUR_JUMP_TO_FIRST_PAGE, Move the cursor to the first page.
;						$LOW_VIEWCUR_JUMP_TO_LAST_PAGE, Move the cursor to the Last page.
;						$LOW_VIEWCUR_JUMP_TO_END_OF_PAGE, Move the cursor to the end of the current page.
;						$LOW_VIEWCUR_JUMP_TO_START_OF_PAGE, Move the cursor to the start of the current page.
;					-TextCursor
;						$LOW_TEXTCUR_COLLAPSE_TO_START,
;						$LOW_TEXTCUR_COLLAPSE_TO_END (Collapses the current selection and moves the cursor  to start or End of selection.
;					#Misc. Cursor Movements:
;					-ViewCursor
;						$LOW_VIEWCUR_JUMP_TO_PAGE (accepts page number to jump to in $iCount, Returns what page was successfully jumped to.
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

	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iMove) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If ($iMove >= UBound($asMoves)) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iCount) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bSelect) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	Switch $iMove
		Case $LOW_TEXTCUR_GO_LEFT, $LOW_TEXTCUR_GO_RIGHT
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "(" & $iCount & "," & $bSelect & ")")
			$iCounted = ($bMoved) ? $iCount : 0
			Return SetError($__LOW_STATUS_SUCCESS, $iCounted, $bMoved)

		Case $LOW_TEXTCUR_GOTO_NEXT_WORD, $LOW_TEXTCUR_GOTO_PREV_WORD, $LOW_TEXTCUR_GOTO_NEXT_SENTENCE, $LOW_TEXTCUR_GOTO_PREV_SENTENCE, _
				$LOW_TEXTCUR_GOTO_NEXT_PARAGRAPH, $LOW_TEXTCUR_GOTO_PREV_PARAGRAPH
			Do
				$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "(" & $bSelect & ")")
				$iCounted += ($bMoved) ? 1 : 0
				Sleep((IsInt($iCounted / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
			Until ($iCounted >= $iCount) Or ($bMoved = False)
			Return SetError($__LOW_STATUS_SUCCESS, $iCounted, $bMoved)

		Case $LOW_TEXTCUR_GOTO_START, $LOW_TEXTCUR_GOTO_END, $LOW_TEXTCUR_GOTO_END_OF_WORD, $LOW_TEXTCUR_GOTO_START_OF_WORD, _
				$LOW_TEXTCUR_GOTO_END_OF_SENTENCE, $LOW_TEXTCUR_GOTO_START_OF_SENTENCE, $LOW_TEXTCUR_GOTO_END_OF_PARAGRAPH, _
				$LOW_TEXTCUR_GOTO_START_OF_PARAGRAPH
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "(" & $bSelect & ")")
			$iCounted = ($bMoved) ? 1 : 0
			Return SetError($__LOW_STATUS_SUCCESS, $iCounted, $bMoved)

		Case $LOW_TEXTCUR_COLLAPSE_TO_START, $LOW_TEXTCUR_COLLAPSE_TO_END
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "()")
			$iCounted = ($bMoved) ? 1 : 0
			Return SetError($__LOW_STATUS_SUCCESS, $iCounted, $bMoved)
		Case Else
			Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0)
	EndSwitch
EndFunc   ;==>__LOWriter_TextCursorMove

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_TransparencyGradientConvert
; Description ...: Convert a Transparency Gradient percentage value to a color value or from a color value to a percentage.
; Syntax ........: __LOWriter_TransparencyGradientConvert([$iPercentToLong = Null[, $iLongToPercent = Null]])
; Parameters ....: $iPercentToLong      - [optional] an integer value. Default is Null. The percentage to convert to Long color integer value.
;                  $iLongToPercent      - [optional] an integer value. Default is Null. The Long color integer value to convert to percentage.
; Return values .: Success: Integer.
;					Failure: Null and sets the @Error and @Extended flags to non-zero.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return Null = No values called in parameters.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success. The requested Integer value to Long color format.
;				   @Error 0 @Extended 1 Return Integer = Success. The requested Integer value from Long color format.
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
		$iReturn = (255 * ($iPercentToLong / 100)) ; Change percentage to decimal and times by White color (255 RGB)
		$iReturn = _LOWriter_ConvertColorToLong(Int($iReturn), Int($iReturn), Int($iReturn))
		Return SetError($__LOW_STATUS_SUCCESS, 0, $iReturn)
	ElseIf ($iLongToPercent <> Null) Then
		$iReturn = _LOWriter_ConvertColorFromLong(Null, $iLongToPercent)
		$iReturn = Int((($iReturn[0] / 255) * 100) + .50) ; All return color values will be the same, so use only one. Add . 50 to round up if applicable.
		Return SetError($__LOW_STATUS_SUCCESS, 1, $iReturn)
	Else
		Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, Null)
	EndIf

EndFunc   ;==>__LOWriter_TransparencyGradientConvert

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_TransparencyGradientNameInsert
; Description ...: Create and insert a new Transparency Gradient name.
; Syntax ........: __LOWriter_TransparencyGradientNameInsert(Byref $oDoc, $tTGradient)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $tTGradient          - a dll struct value. A Gradient Structure to copy settings from.
; Return values .: Success: String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $tTGradient not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.drawing.TransparencyGradientTable" Object.
;				   @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.awt.Gradient" structure.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error creating Transparency Gradient Name.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. A new transparency Gradient name was created. Returning the new name as a string.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If The Transparency Gradient name is blank, I need to create a new name and apply it. I think I could re-use
;					an old one without problems, but I'm not sure, so to be safe, I will create a new one. If there are no names
;					that have been already created, then I need to create and apply one before the transparency gradient will
;					be displayed.
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

	If Not IsObj($oDoc) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($tTGradient) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	$oTGradTable = $oDoc.createInstance("com.sun.star.drawing.TransparencyGradientTable")
	If Not IsObj($oTGradTable) Then Return SetError($__LOW_STATUS_INIT_ERROR, 1, 0)

	While $oTGradTable.hasByName("Transparency " & $iCount)
		$iCount += 1
		Sleep((IsInt($iCount / $__LOWCONST_SLEEP_DIV)) ? 10 : 0)
	WEnd

	$tNewTGradient = __LOWriter_CreateStruct("com.sun.star.awt.Gradient")
	If Not IsObj($tNewTGradient) Then Return SetError($__LOW_STATUS_INIT_ERROR, 2, 0)

	; Copy the settings over from the input Style Gradient to my new one. This may not be necessary? But just in case.
	With $tNewTGradient
		.Style = $tTGradient.Style()
		.XOffset = $tTGradient.XOffset()
		.YOffset = $tTGradient.YOffset()
		.Angle = $tTGradient.Angle()
		.Border = $tTGradient.Border()
		.StartColor = $tTGradient.StartColor()
		.EndColor = $tTGradient.EndColor()
	EndWith

	$oTGradTable.insertByName("Transparency " & $iCount, $tNewTGradient)
	If Not ($oTGradTable.hasByName("Transparency " & $iCount)) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LOW_STATUS_SUCCESS, 0, "Transparency " & $iCount)
EndFunc   ;==>__LOWriter_TransparencyGradientNameInsert

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_UnitConvert
; Description ...: For converting measurement units.
; Syntax ........: __LOWriter_UnitConvert($nValue, $sReturnType)
; Parameters ....: $nValue              - a general number value. The Number to be converted.
;                  $iReturnType         - a Integer value. Determines conversion type. See Constants, $__LOWCONST_CONVERT_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .:Success: Integer or Number.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $nValue is Not a Number.
;				   @Error 1 @Extended 2 Return 0 = $iReturnType is not a Integer.
;				   @Error 1 @Extended 3 Return 0 = $iReturnType does not match predefined input.
;				   --Success--
;				   @Error 0 @Extended 1 Return Number =  Returns Number converted from TWIPS to Centimeters.
;				   @Error 0 @Extended 2 Return Number =  Returns Number converted from TWIPS to Inches.
;				   @Error 0 @Extended 3 Return Integer =  Returns Number converted from Millimeters to uM (Micrometers).
;				   @Error 0 @Extended 4 Return Number =  Returns Number converted from Micrometers to MM
;				   @Error 0 @Extended 5 Return Integer =  Returns Number converted from Centimeters To uM
;				   @Error 0 @Extended 6 Return Number =  Returns Number converted from um (Micrometers) To CM
;				   @Error 0 @Extended 7 Return Integer =  Returns Number converted from Inches to uM(Micrometers).
;				   @Error 0 @Extended 8 Return Number =  Returns Number converted from uM(Micrometers) to Inches.
;				   @Error 0 @Extended 9 Return Integer =  Returns Number converted from TWIPS to uM(Micrometers).
;				   @Error 0 @Extended 10 Return Integer = Returns Number converted from Point to uM(Micrometers).
;				   @Error 0 @Extended 11 Return Number = Returns Number converted from uM(Micrometers) to Point.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_UnitConvert($nValue, $iReturnType)
	Local $iUM, $iMM, $iCM, $iInch

	If Not IsNumber($nValue) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iReturnType) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)

	Switch $iReturnType

		Case $__LOWCONST_CONVERT_TWIPS_CM ;TWIPS TO CM
			; 1 TWIP = 1/20 of a point, 1 Point = 1/72 of an Inch.
			$iInch = ($nValue / 20 / 72)
			; 1 Inch = 2.54 CM
			$iCM = Round(Round($iInch * 2.54, 3), 2)
			Return SetError($__LOW_STATUS_SUCCESS, 1, Number($iCM))

		Case $__LOWCONST_CONVERT_TWIPS_INCH ;TWIPS to Inch
			; 1 TWIP = 1/20 of a point, 1 Point = 1/72 of an Inch.
			$iInch = ($nValue / 20 / 72)
			$iInch = Round(Round($iInch, 3), 2)
			Return SetError($__LOW_STATUS_SUCCESS, 2, Number($iInch))

		Case $__LOWCONST_CONVERT_MM_UM ;Millimeter to Micrometer
			$iUM = ($nValue * 100)
			$iUM = Round(Round($iUM, 1))
			Return SetError($__LOW_STATUS_SUCCESS, 3, Number($iUM))

		Case $__LOWCONST_CONVERT_UM_MM ;Micrometer to Millimeter
			$iMM = ($nValue / 100)
			$iMM = Round(Round($iMM, 3), 2)
			Return SetError($__LOW_STATUS_SUCCESS, 4, Number($iMM))

		Case $__LOWCONST_CONVERT_CM_UM ;Centimeter to Micrometer
			$iUM = ($nValue * 1000)
			$iUM = Round(Round($iUM, 1))
			Return SetError($__LOW_STATUS_SUCCESS, 5, Int($iUM))

		Case $__LOWCONST_CONVERT_UM_CM ;Micrometer to Centimeter
			$iCM = ($nValue / 1000)
			$iCM = Round(Round($iCM, 3), 2)
			Return SetError($__LOW_STATUS_SUCCESS, 6, Number($iCM))

		Case $__LOWCONST_CONVERT_INCH_UM ;Inch to Micrometer
			; 1 Inch - 2.54 Cm; Micrometer = 1/1000 CM
			$iUM = ($nValue * 2.54) * 1000 ; + .0055
			$iUM = Round(Round($iUM, 1))
			Return SetError($__LOW_STATUS_SUCCESS, 7, Int($iUM))

		Case $__LOWCONST_CONVERT_UM_INCH ;Micrometer to Inch
			; 1 Inch - 2.54 Cm; Micrometer = 1/1000 CM
			$iInch = ($nValue / 1000) / 2.54 ; + .0055
			$iInch = Round(Round($iInch, 3), 2)
			Return SetError($__LOW_STATUS_SUCCESS, 8, $iInch)

		Case $__LOWCONST_CONVERT_TWIPS_UM ;TWIPS to MicroMeter
			; 1 TWIP = 1/20 of a point, 1 Point = 1/72 of an Inch.
			$iInch = (($nValue / 20) / 72)
			$iInch = Round(Round($iInch, 3), 2)
			; 1 Inch - 25.4 MM; Micrometer = 1/100 MM
			$iUM = Round($iInch * 25.4 * 100)
			Return SetError($__LOW_STATUS_SUCCESS, 9, Int($iUM))

		Case $__LOWCONST_CONVERT_PT_UM
			; 1 pt = 35 uM
			Return ($nValue = 0) ? SetError($__LOW_STATUS_SUCCESS, 10, 0) : SetError($__LOW_STATUS_SUCCESS, 10, Round(($nValue * 35.2778)))

		Case $__LOWCONST_CONVERT_UM_PT
			Return ($nValue = 0) ? SetError($__LOW_STATUS_SUCCESS, 11, 0) : SetError($__LOW_STATUS_SUCCESS, 11, Round(($nValue / 35.2778), 2))

		Case Else
			Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	EndSwitch
EndFunc   ;==>__LOWriter_UnitConvert

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_VarsAreDefault
; Description ...: Tests whether all input parameters are equal to Default keyword.
; Syntax ........: __LOWriter_VarsAreDefault($vVar1[, $vVar2 = Default[, $vVar3 = Default[, $vVar4 = Default[, $vVar5 = Default[, $vVar6 = Default[, $vVar7 = Default[, $vVar8 = Default]]]]]]])
; Parameters ....: $vVar1               - a variant value.
;                  $vVar2               - [optional] a variant value. Default is Default.
;                  $vVar3               - [optional] a variant value. Default is Default.
;                  $vVar4               - [optional] a variant value. Default is Default.
;                  $vVar5               - [optional] a variant value. Default is Default.
;                  $vVar6               - [optional] a variant value. Default is Default.
;                  $vVar7               - [optional] a variant value. Default is Default.
;                  $vVar8               - [optional] a variant value. Default is Default.
; Return values .: Success: Boolean
;				   Failure: False
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = If All parameters are Equal to Default, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_VarsAreDefault($vVar1, $vVar2 = Default, $vVar3 = Default, $vVar4 = Default, $vVar5 = Default, $vVar6 = Default, $vVar7 = Default, $vVar8 = Default)
	Local $bAllDefault1, $bAllDefault2
	$bAllDefault1 = (($vVar1 = Default) And ($vVar2 = Default) And ($vVar3 = Default) And ($vVar4 = Default)) ? True : False
	$bAllDefault2 = (($vVar5 = Default) And ($vVar6 = Default) And ($vVar7 = Default) And ($vVar8 = Default)) ? True : False
	Return ($bAllDefault1 And $bAllDefault2) ? True : False
EndFunc   ;==>__LOWriter_VarsAreDefault

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_VarsAreNull
; Description ...: Tests whether all input parameters are equal to Null keyword.
; Syntax ........: __LOWriter_VarsAreNull($vVar1[, $vVar2 = Null[, $vVar3 = Null[, $vVar4 = Null[, $vVar5 = Null[, $vVar6 = Null[, $vVar7 = Null[, $vVar8 = Null[, $vVar9 = Null[, $vVar10 = Null[, $vVar11 = Null[, $vVar12 = Null]]]]]]]]]]])
; Parameters ....: $vVar1               - a variant value.
;                  $vVar2               - [optional] a variant value. Default is Null.
;                  $vVar3               - [optional] a variant value. Default is Null.
;                  $vVar4               - [optional] a variant value. Default is Null.
;                  $vVar5               - [optional] a variant value. Default is Null.
;                  $vVar6               - [optional] a variant value. Default is Null.
;                  $vVar7               - [optional] a variant value. Default is Null.
;                  $vVar8               - [optional] a variant value. Default is Null.
;                  $vVar9               - [optional] a variant value. Default is Null.
;                  $vVar10              - [optional] a variant value. Default is Null.
;                  $vVar11              - [optional] a variant value. Default is Null.
;                  $vVar12              - [optional] a variant value. Default is Null.
; Return values .: Success: Boolean
;				   Failure: False
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = If All parameters are Equal to Null, True is returned. Else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_VarsAreNull($vVar1, $vVar2 = Null, $vVar3 = Null, $vVar4 = Null, $vVar5 = Null, $vVar6 = Null, $vVar7 = Null, $vVar8 = Null, $vVar9 = Null, $vVar10 = Null, $vVar11 = Null, $vVar12 = Null)
	Local $bAllNull1, $bAllNull2, $bAllNull3
	$bAllNull1 = (($vVar1 = Null) And ($vVar2 = Null) And ($vVar3 = Null) And ($vVar4 = Null)) ? True : False
	If (@NumParams <= 4) Then Return ($bAllNull1) ? True : False
	$bAllNull2 = (($vVar5 = Null) And ($vVar6 = Null) And ($vVar7 = Null) And ($vVar8 = Null)) ? True : False
	If (@NumParams <= 8) Then Return ($bAllNull1 And $bAllNull2) ? True : False
	$bAllNull3 = (($vVar9 = Null) And ($vVar10 = Null) And ($vVar11 = Null) And ($vVar12 = Null)) ? True : False
	Return ($bAllNull1 And $bAllNull2 And $bAllNull3) ? True : False
EndFunc   ;==>__LOWriter_VarsAreNull

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_VersionCheck
; Description ...: Test if the currently installed LibreOffice version is high enough to support a certain function.
; Syntax ........: __LOWriter_VersionCheck($fRequiredVersion)
; Parameters ....: $fRequiredVersion            - a floating point value. The version of LibreOffice required.
; Return values .: Success: Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $fRequiredVersion not a Number.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error retrieving Current LO Version.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = Success. If the Current LO version is higher than or equal to the required version, then the return is True, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_VersionCheck($fRequiredVersion)
	Local Static $sCurrentVersion = _LOWriter_VersionGet(True, False)
	If (@error > 0) Then Return SetError($__LOW_STATUS_PROCESSING_ERROR, 1, False)
	Local Static $fCurrentVersion = Number($sCurrentVersion)

	If Not IsNumber($fRequiredVersion) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, False)

	Return SetError($__LOW_STATUS_SUCCESS, 1, ($fCurrentVersion >= $fRequiredVersion) ? True : False)
EndFunc   ;==>__LOWriter_VersionCheck

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_ViewCursorMove
; Description ...: For ViewCursor related movements.
; Syntax ........: __LOWriter_ViewCursorMove(Byref $oCursor, $iMove, $iCount[, $bSelect = False])
; Parameters ....: $oCursor             - [in/out] an object. A ViewCursor Object returned from _LOWriter_DocGetViewCursor function.
;                  $iMove               - an integer value. The movement command. See remarks and Constants.
;                  $iCount              - an integer value. Number of movements to make.
;                  $bSelect             - [optional] a boolean value. Default is False. Whether to select data during this cursor movement.
; Return values .: Success: Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iMove not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iMove mismatch with Cursor type. See Cursor Type/Move Type Constants.
;				   @Error 1 @Extended 4 Return 0 = $iCount not an integer or is a negative.
;				   @Error 1 @Extended 5 Return 0 = $bSelect not a Boolean.
;				   --Processing Errors--
;				   @Error 3 @Extended 2 Return 0 = Error processing cursor move.
;				   --Success--
;				   @Error 0 @Extended ? Return Boolean = Success, Cursor object movement was processed successfully.
;				   +				Returns True if the full count of movements were successful, else false if none or only partially successful.
;				   +				@Extended set to number of successful movements. Or Page Number for "gotoPage" command. See Remarks
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iMove may be set to any of the following constants depending on the Cursor type you are intending to move.
;					 Only some movements accept movement amounts (such as "goRight" 2) etc. Also only some accept creating/
;					extending a selection of text/ data. They will be specified below. To Clear /Unselect a current selection,
;					you can input a move such as "goRight", 0, False.
; Cursor Movement Constants:
;					#Cursor Movement Constants which accept number of Moves and Selecting:
;					-ViewCursor
;						$LOW_VIEWCUR_GO_DOWN, Move the cursor Down by n lines.
;						$LOW_VIEWCUR_GO_UP, Move the cursor Up by n lines.
;						$LOW_VIEWCUR_GO_LEFT, Move the cursor left by n characters.
;						$LOW_VIEWCUR_GO_RIGHT, Move the cursor right by n characters.
;					-TextCursor
;						$LOW_TEXTCUR_GO_LEFT,Move the cursor left by n characters.
;						$LOW_TEXTCUR_GO_RIGHT, Move the cursor right by n characters.
;						$LOW_TEXTCUR_GOTO_NEXT_WORD, Move to the start of the next word.
;						$LOW_TEXTCUR_GOTO_PREV_WORD, Move to the end of the previous word.
;						$LOW_TEXTCUR_GOTO_NEXT_SENTENCE,Move to the start of the next sentence.
;						$LOW_TEXTCUR_GOTO_PREV_SENTENCE, Move to the end of the previous sentence.
;						$LOW_TEXTCUR_GOTO_NEXT_PARAGRAPH, Move to the start of the next paragraph.
;						$LOW_TEXTCUR_GOTO_PREV_PARAGRAPH, Move to the End of the previous paragraph.
;					-TableCursor
;						$LOW_TABLECUR_GO_LEFT, Move the cursor left/right n cells.
;						$LOW_TABLECUR_GO_RIGHT, Move the cursor left/right n cells.
;						$LOW_TABLECUR_GO_UP,  Move the cursor up/down n cells.
;						$LOW_TABLECUR_GO_DOWN, Move the cursor up/down n cells.
;					#Cursor Movements which accept number of Moves Only:
;					-ViewCursor
;						$LOW_VIEWCUR_JUMP_TO_NEXT_PAGE, Move the cursor to the Next page.
;						$LOW_VIEWCUR_JUMP_TO_PREV_PAGE, Move the cursor to the previous page.
;						$LOW_VIEWCUR_SCREEN_DOWN, Scroll the view forward by one visible page.
;						$LOW_VIEWCUR_SCREEN_UP, Scroll the view back by one visible page.
;					#Cursor Movements which accept Selecting Only:
;					-ViewCursor
;						$LOW_VIEWCUR_GOTO_END_OF_LINE, Move the cursor to the end of the current line.
;						$LOW_VIEWCUR_GOTO_START_OF_LINE, Move the cursor to the start of the current line.
;						$LOW_VIEWCUR_GOTO_START, Move the cursor to the start of the document or Table.
;						$LOW_VIEWCUR_GOTO_END, Move the cursor to the end of the document or Table.
;					-TextCursor
;						$LOW_TEXTCUR_GOTO_START, Move the cursor to the start of the text.
;						$LOW_TEXTCUR_GOTO_END, Move the cursor to the end of the text.
;						$LOW_TEXTCUR_GOTO_END_OF_WORD, Move to the end of the current word.
;						$LOW_TEXTCUR_GOTO_START_OF_WORD, Move to the start of the current word.
;						$LOW_TEXTCUR_GOTO_END_OF_SENTENCE, Move to the end of the current sentence.
;						$LOW_TEXTCUR_GOTO_START_OF_SENTENCE, Move to the start of the current sentence.
;						$LOW_TEXTCUR_GOTO_END_OF_PARAGRAPH, Move to the end of the current paragraph.
;						$LOW_TEXTCUR_GOTO_START_OF_PARAGRAPH, Move to the start of the current paragraph.
;					-TableCursor
;						$LOW_TABLECUR_GOTO_START, Move the cursor to the top left cell.
;						$LOW_TABLECUR_GOTO_END,  Move the cursor to the bottom right cell.
;					#Cursor Movements which accept nothing and are done once per call:
;					-ViewCursor
;						$LOW_VIEWCUR_JUMP_TO_FIRST_PAGE, Move the cursor to the first page.
;						$LOW_VIEWCUR_JUMP_TO_LAST_PAGE, Move the cursor to the Last page.
;						$LOW_VIEWCUR_JUMP_TO_END_OF_PAGE, Move the cursor to the end of the current page.
;						$LOW_VIEWCUR_JUMP_TO_START_OF_PAGE, Move the cursor to the start of the current page.
;					-TextCursor
;						$LOW_TEXTCUR_COLLAPSE_TO_START,
;						$LOW_TEXTCUR_COLLAPSE_TO_END (Collapses the current selection and moves the cursor  to start or End of selection.
;					#Misc. Cursor Movements:
;					-ViewCursor
;						$LOW_VIEWCUR_JUMP_TO_PAGE (accepts page number to jump to in $iCount, Returns what page was successfully jumped to.
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

	If Not IsObj($oCursor) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iMove) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 2, 0)
	If ($iMove >= UBound($asMoves)) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iCount) Or ($iCount < 0) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bSelect) Then Return SetError($__LOW_STATUS_INPUT_ERROR, 5, 0)

	Switch $iMove
		Case $LOW_VIEWCUR_GO_DOWN, $LOW_VIEWCUR_GO_UP, $LOW_VIEWCUR_GO_LEFT, $LOW_VIEWCUR_GO_RIGHT
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "(" & $iCount & "," & $bSelect & ")")
			$iCounted = ($bMoved) ? $iCount : 0
			Return SetError($__LOW_STATUS_SUCCESS, $iCounted, $bMoved)

		Case $LOW_VIEWCUR_GOTO_END_OF_LINE, $LOW_VIEWCUR_GOTO_START_OF_LINE, $LOW_VIEWCUR_GOTO_START, $LOW_VIEWCUR_GOTO_END
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "(" & $bSelect & ")")
			$iCounted = ($bMoved) ? 1 : 0
			Return SetError($__LOW_STATUS_SUCCESS, $iCounted, $bMoved)

		Case $LOW_VIEWCUR_JUMP_TO_PAGE
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "(" & $iCount & ")")
			Return SetError($__LOW_STATUS_SUCCESS, $oCursor.getPage(), $bMoved)

		Case $LOW_VIEWCUR_JUMP_TO_NEXT_PAGE, $LOW_VIEWCUR_JUMP_TO_PREV_PAGE, $LOW_VIEWCUR_SCREEN_DOWN, $LOW_VIEWCUR_SCREEN_UP
			Do
				$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "()")
				$iCounted += ($bMoved) ? 1 : 0
				Sleep((IsInt($iCounted / $__LOWCONST_SLEEP_DIV) ? 10 : 0))
			Until ($iCounted >= $iCount) Or ($bMoved = False)
			Return SetError($__LOW_STATUS_SUCCESS, $iCounted, $bMoved)

		Case $LOW_VIEWCUR_JUMP_TO_FIRST_PAGE, $LOW_VIEWCUR_JUMP_TO_LAST_PAGE, $LOW_VIEWCUR_JUMP_TO_END_OF_PAGE, _
				$LOW_VIEWCUR_JUMP_TO_START_OF_PAGE
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "()")
			$iCounted = ($bMoved) ? 1 : 0
			Return SetError($__LOW_STATUS_SUCCESS, $iCounted, $bMoved)
		Case Else
			Return SetError($__LOW_STATUS_PROCESSING_ERROR, 2, 0)
	EndSwitch
EndFunc   ;==>__LOWriter_ViewCursorMove
