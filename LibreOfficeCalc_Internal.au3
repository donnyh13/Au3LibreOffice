#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Calc
#include "LibreOfficeCalc_Constants.au3"
#include "LibreOfficeCalc_Helper.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Various functions for internal data processing, data retrieval, retrieving and applying settings for LibreOffice UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __LOCalc_CellAddressIsSame
; __LOCalc_CellBackColor
; __LOCalc_CellBorder
; __LOCalc_CellBorderPadding
; __LOCalc_CellEffect
; __LOCalc_CellFont
; __LOCalc_CellFontColor
; __LOCalc_CellNumberFormat
; __LOCalc_CellOverLine
; __LOCalc_CellProtection
; __LOCalc_CellShadow
; __LOCalc_CellStrikeOut
; __LOCalc_CellStyleBorder
; __LOCalc_CellTextAlign
; __LOCalc_CellTextOrient
; __LOCalc_CellTextProperties
; __LOCalc_CellUnderLine
; __LOCalc_CharPosition
; __LOCalc_CharSpacing
; __LOCalc_CommentAreaShadowModify
; __LOCalc_CommentArrowStyleName
; __LOCalc_CommentGetObjByCell
; __LOCalc_CommentLineStyleName
; __LOCalc_FieldGetObj
; __LOCalc_FieldTypeServices
; __LOCalc_FilterNameGet
; __LOCalc_Internal_CursorGetType
; __LOCalc_InternalComErrorHandler
; __LOCalc_NamedRangeGetScopeObj
; __LOCalc_PageStyleBorder
; __LOCalc_PageStyleFooterBorder
; __LOCalc_PageStyleHeaderBorder
; __LOCalc_RangeAddressIsSame
; __LOCalc_SheetCursorMove
; __LOCalc_TextCursorMove
; __LOCalc_TransparencyGradientConvert
; __LOCalc_TransparencyGradientNameInsert
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CellAddressIsSame
; Description ...: Compare two Cell Addresses to see if they are the same.
; Syntax ........: __LOCalc_CellAddressIsSame(ByRef $tCellAddr1, ByRef $tCellAddr2)
; Parameters ....: $tCellAddr1          - a dll struct value. The first Cell Address Structure to compare.
;                  $tCellAddr2          - a dll struct value. The second Cell Address Structure to compare.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $tCellAddr1 not an Object.
;                  @Error 1 @Extended 2 Return 0 = $tCellAddr2 not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If the Cell Addresses are identical, True is returned, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CellAddressIsSame($tCellAddr1, $tCellAddr2)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($tCellAddr1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($tCellAddr2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If ($tCellAddr1.Sheet() = $tCellAddr2.Sheet()) And _
			($tCellAddr1.Column() = $tCellAddr2.Column()) And _
			($tCellAddr1.Row() = $tCellAddr2.Row()) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>__LOCalc_CellAddressIsSame

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CellBackColor
; Description ...: Internal function to Set or Retrieve the background color setting for a Cell, Cell Range, or Cell Style.
; Syntax ........: __LOCalc_CellBackColor(ByRef $oObj, $iBackColor, $bBackTransparent)
; Parameters ....: $oObj                - [in/out] an object. A Cell, Cell Range or Cell Style Object returned from an applicable function.
;                  $iBackColor          - an integer value (-1-16777215). The background color. Set in Long integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Set to $LO_COLOR_OFF(-1), to turn Background color off.
;                  $bBackTransparent    - a boolean value. If True, the background color is transparent.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  @Error 1 @Extended 5 Return 0 = $bBackTransparent not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iBackColor
;                  |                               2 = Error setting $bBackTransparent
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CellBackColor(ByRef $oObj, $iBackColor, $bBackTransparent)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avColor[2]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($iBackColor, $bBackTransparent) Then
		__LO_ArrayFill($avColor, $oObj.CellBackColor(), $oObj.IsCellBackgroundTransparent())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avColor)
	EndIf

	If ($iBackColor <> Null) Then
		If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.CellBackColor = $iBackColor
		$iError = ($oObj.CellBackColor() = $iBackColor) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bBackTransparent <> Null) Then
		If Not IsBool($bBackTransparent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.IsCellBackgroundTransparent = $bBackTransparent
		$iError = ($oObj.IsCellBackgroundTransparent() = $bBackTransparent) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOCalc_CellBackColor

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CellBorder
; Description ...: Internal function to Set and Retrieve the Cell, or Cell Range Border Line Width, Style, and Color. Libre Office Version 3.6 and Up.
; Syntax ........: __LOCalc_CellBorder(ByRef $oRange, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight, $iVert, $iHori, $iTLBRDiag, $iBLTRDiag)
; Parameters ....: $oRange              - [in/out] an object. A Cell Range or Cell object returned by a previous _LOCalc_RangeGetCellByName, _LOCalc_RangeGetCellByPosition, _LOCalc_RangeColumnGetObjByPosition, _LOCalc_RangeColumnGetObjByName, _LOcalc_RangeRowGetObjByPosition, _LOCalc_SheetGetObjByName, or _LOCalc_SheetGetActive function.
;                  $bWid                - a boolean value. If True, Border Width is being modified. Only one can be True at once.
;                  $bSty                - a boolean value. If True, Border Style is being modified. Only one can be True at once.
;                  $bCol                - a boolean value. If True, Border Color is being modified. Only one can be True at once.
;                  $iTop                - an integer value. Modifies the top border line settings. See Width, Style or Color functions for values.
;                  $iBottom             - an integer value. Modifies the bottom border line settings. See Width, Style or Color functions for values.
;                  $iLeft               - an integer value. Modifies the left border line settings. See Width, Style or Color functions for values.
;                  $iRight              - an integer value. Modifies the right border line settings. See Width, Style or Color functions for values.
;                  $iVert               - an integer value. Modifies the vertical border line settings. See Width, Style or Color functions for values.
;                  $iHori               - an integer value. Modifies the horizontal border line settings. See Width, Style or Color functions for values.
;                  $iTLBRDiag           - an integer value. Modifies the top-left to bottom-right diagonal border line settings. See Width, Style or Color functions for values.
;                  $iBLTRDiag           - an integer value. Modifies the bottom-left to top-right diagonal border line settings. See Width, Style or Color functions for values.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 11 Return 0 = Variable passed to internal function not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  @Error 3 @Extended 2 Return 0 = Error Retrieving TableBorder2 Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style/Color when Top Border width not set.
;                  @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Style/Color when Bottom Border width not set.
;                  @Error 4 @Extended 3 Return 0 = Cannot set Left Border Style/Color when Left Border width not set.
;                  @Error 4 @Extended 4 Return 0 = Cannot set Right Border Style/Color when Right Border width not set.
;                  @Error 4 @Extended 5 Return 0 = Cannot set Vertical Border Style/Color when Vertical Border width not set.
;                  @Error 4 @Extended 6 Return 0 = Cannot set Horizontal Border Style/Color when Horizontal Border width not set.
;                  @Error 4 @Extended 7 Return 0 = Cannot set Top-Left to Bottom-Right Diagonal Border Style/Color when Top-Left to Bottom-Right Diagonal Border width not set.
;                  @Error 4 @Extended 8 Return 0 = Cannot set Bottom-Left to Top-Right Diagonal Border Style/Color when Bottom-Left to Top-Right Diagonal Border width not set.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 8 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CellBorder(ByRef $oRange, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight, $iVert, $iHori, $iTLBRDiag, $iBLTRDiag)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avBorder[8]
	Local $tBL2, $tTB2

	If Not __LO_VersionCheck(3.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)
	If (($bWid + $bSty + $bCol) <> 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight, $iVert, $iHori, $iTLBRDiag, $iBLTRDiag) Then
		If $bWid Then
			__LO_ArrayFill($avBorder, $oRange.TableBorder2.TopLine.LineWidth(), $oRange.TableBorder2.BottomLine.LineWidth(), _
					$oRange.TableBorder2.LeftLine.LineWidth(), $oRange.TableBorder2.RightLine.LineWidth(), $oRange.TableBorder2.VerticalLine.LineWidth(), _
					$oRange.TableBorder2.HorizontalLine.LineWidth(), $oRange.DiagonalTLBR2.LineWidth(), $oRange.DiagonalBLTR2.LineWidth())

		ElseIf $bSty Then
			__LO_ArrayFill($avBorder, $oRange.TableBorder2.TopLine.LineStyle(), $oRange.TableBorder2.BottomLine.LineStyle(), _
					$oRange.TableBorder2.LeftLine.LineStyle(), $oRange.TableBorder2.RightLine.LineStyle(), $oRange.TableBorder2.VerticalLine.LineStyle(), _
					$oRange.TableBorder2.HorizontalLine.LineStyle(), $oRange.DiagonalTLBR2.LineStyle(), $oRange.DiagonalBLTR2.LineStyle())

		ElseIf $bCol Then
			__LO_ArrayFill($avBorder, $oRange.TableBorder2.TopLine.Color(), $oRange.TableBorder2.BottomLine.Color(), _
					$oRange.TableBorder2.LeftLine.Color(), $oRange.TableBorder2.RightLine.Color(), $oRange.TableBorder2.VerticalLine.Color(), _
					$oRange.TableBorder2.HorizontalLine.Color(), $oRange.DiagonalTLBR2.Color(), $oRange.DiagonalBLTR2.Color())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avBorder)
	EndIf

	$tBL2 = __LO_CreateStruct("com.sun.star.table.BorderLine2")
	If Not IsObj($tBL2) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$tTB2 = $oRange.TableBorder2
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
		;~ 		$tBL2.LineWidth = ($bWid) ? ($iHori) : ($tTB2.HorizontalLine.LineWidth()) ; copy Line Width over to new size structure

		; I have to use OuterLineWidth instead of LineWidth because LineWidth doesn't set for some reason for Horizontal
		$tBL2.OuterLineWidth = ($bWid) ? ($iHori) : ($tTB2.HorizontalLine.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iHori) : ($tTB2.HorizontalLine.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iHori) : ($tTB2.HorizontalLine.Color()) ; copy Color over to new size structure
		$tTB2.HorizontalLine = $tBL2
	EndIf

	$oRange.TableBorder2 = $tTB2

	If $iTLBRDiag <> Null Then
		If Not $bWid And ($oRange.DiagonalTLBR2.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 7, 0) ; If Width not set, cant set color or style.

		; Top-Left to Bottom Right Diagonal Line
		$tBL2.LineWidth = ($bWid) ? ($iTLBRDiag) : ($oRange.DiagonalTLBR2.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iTLBRDiag) : ($oRange.DiagonalTLBR2.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iTLBRDiag) : ($oRange.DiagonalTLBR2.Color()) ; copy Color over to new size structure
		$oRange.DiagonalTLBR2 = $tBL2
	EndIf

	If $iBLTRDiag <> Null Then
		If Not $bWid And ($oRange.DiagonalBLTR2.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 8, 0) ; If Width not set, cant set color or style.

		; Bottom-Left to Top-Right Diagonal Line
		$tBL2.LineWidth = ($bWid) ? ($iBLTRDiag) : ($oRange.DiagonalBLTR2.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iBLTRDiag) : ($oRange.DiagonalBLTR2.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iBLTRDiag) : ($oRange.DiagonalBLTR2.Color()) ; copy Color over to new size structure
		$oRange.DiagonalBLTR2 = $tBL2
	EndIf
EndFunc   ;==>__LOCalc_CellBorder

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CellBorderPadding
; Description ...: Internal function to Set or retrieve the Cell, Cell Range, or Cell Style Border Padding settings.
; Syntax ........: __LOCalc_CellBorderPadding(ByRef $oObj, $iAll, $iTop, $iBottom, $iLeft, $iRight)
; Parameters ....: $oObj                - [in/out] an object. A Cell, Cell Range or Cell Style Object returned from an applicable function.
;                  $iAll                - an integer value. Set all four padding distances to one distance in Micrometers (uM).
;                  $iTop                - an integer value. Set the Top Distance between the Border and Cell contents, in Micrometers(uM).
;                  $iBottom             - an integer value. Set the Bottom Distance between the Border and Cell contents, in Micrometers(uM).
;                  $iLeft               - an integer value. Set the Left Distance between the Border and Cell contents, in Micrometers(uM).
;                  $iRight              - an integer value. Set the Right Distance between the Border and Cell contents, in Micrometers(uM).
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iAll not an Integer, or less than 0.
;                  @Error 1 @Extended 5 Return 0 = $iTop not an Integer, or less than 0.
;                  @Error 1 @Extended 6 Return 0 = $iBottom not an Integer, or less than 0.
;                  @Error 1 @Extended 7 Return 0 = $iLeft not an Integer, or less than 0.
;                  @Error 1 @Extended 8 Return 0 = $iRight not an Integer, or less than 0.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  $iAll returns an integer value if all (Top, Bottom, Left, Right) padding values are equal, else Null is returned.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CellBorderPadding(ByRef $oObj, $iAll, $iTop, $iBottom, $iLeft, $iRight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiBPadding[5]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($iAll, $iTop, $iBottom, $iLeft, $iRight) Then ; Return Top Margin value for $iAll
		__LO_ArrayFill($aiBPadding, (($oObj.ParaTopMargin() = $oObj.ParaBottomMargin()) And ($oObj.ParaLeftMargin() = $oObj.ParaRightMargin()) And ($oObj.ParaBottomMargin() = $oObj.ParaLeftMargin())) ? ($oObj.ParaBottomMargin()) : (Null), _
				$oObj.ParaTopMargin(), $oObj.ParaBottomMargin(), $oObj.ParaLeftMargin(), $oObj.ParaRightMargin())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiBPadding)
	EndIf

	If ($iAll <> Null) Then
		If Not __LO_IntIsBetween($iAll, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.ParaTopMargin = $iAll
		$oObj.ParaBottomMargin = $iAll
		$oObj.ParaLeftMargin = $iAll
		$oObj.ParaRightMargin = $iAll
		$iError = ($iTop <> Null) ? ($iError) : (__LO_IntIsBetween($oObj.ParaTopMargin(), $iAll - 1, $iAll + 1)) ? ($iError) : (BitOR($iError, 1))
		$iError = ($iBottom <> Null) ? ($iError) : (__LO_IntIsBetween($oObj.ParaBottomMargin(), $iAll - 1, $iAll + 1)) ? ($iError) : (BitOR($iError, 2))
		$iError = ($iLeft <> Null) ? ($iError) : (__LO_IntIsBetween($oObj.ParaLeftMargin(), $iAll - 1, $iAll + 1)) ? ($iError) : (BitOR($iError, 4))
		$iError = ($iRight <> Null) ? ($iError) : (__LO_IntIsBetween($oObj.ParaRightMargin(), $iAll - 1, $iAll + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iTop <> Null) Then
		If Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.ParaTopMargin = $iTop
		$iError = (__LO_IntIsBetween($oObj.ParaTopMargin(), $iTop - 1, $iTop + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iBottom <> Null) Then
		If Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.ParaBottomMargin = $iBottom
		$iError = (__LO_IntIsBetween($oObj.ParaBottomMargin(), $iBottom - 1, $iBottom + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iLeft <> Null) Then
		If Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oObj.ParaLeftMargin = $iLeft
		$iError = (__LO_IntIsBetween($oObj.ParaLeftMargin(), $iLeft - 1, $iLeft + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iRight <> Null) Then
		If Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oObj.ParaRightMargin = $iRight
		$iError = (__LO_IntIsBetween($oObj.ParaRightMargin(), $iRight - 1, $iRight + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOCalc_CellBorderPadding

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CellEffect
; Description ...: Internal function to Set or Retrieve the Font Effect settings for a Cell, Cell Range, or Cell Style.
; Syntax ........: __LOCalc_CellEffect(ByRef $oObj, $iRelief, $bOutline, $bShadow)
; Parameters ....: $oObj                - [in/out] an object. A Cell, Cell Range or Cell Style Object returned from an applicable function.
;                  $iRelief             - an integer value (0-2). The Character Relief style. See Constants, $LOC_RELIEF_* as defined in LibreOfficeCalc_Constants.au3.
;                  $bOutline            - a boolean value. If True, the characters have an outline around the outside.
;                  $bShadow             - a boolean value. If True, the characters have a shadow.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iRelief not an Integer, less than 0 or greater than 2. See Constants, $LOC_RELIEF_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $bOutline not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bShadow not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iRelief
;                  |                               2 = Error setting $bOutline
;                  |                               4 = Error setting $bShadow
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CellEffect(ByRef $oObj, $iRelief, $bOutline, $bShadow)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avEffect[3]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($iRelief, $bOutline, $bShadow) Then
		__LO_ArrayFill($avEffect, $oObj.CharRelief(), $oObj.CharContoured(), $oObj.CharShadowed())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avEffect)
	EndIf

	If ($iRelief <> Null) Then
		If Not __LO_IntIsBetween($iRelief, $LOC_RELIEF_NONE, $LOC_RELIEF_ENGRAVED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.CharRelief = $iRelief
		$iError = ($oObj.CharRelief() = $iRelief) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bOutline <> Null) Then
		If Not IsBool($bOutline) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.CharContoured = $bOutline
		$iError = ($oObj.CharContoured() = $bOutline) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bShadow <> Null) Then
		If Not IsBool($bShadow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.CharShadowed = $bShadow
		$iError = ($oObj.CharShadowed() = $bShadow) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOCalc_CellEffect

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CellFont
; Description ...: Internal function to Set and Retrieve the Font Settings for a Cell, Cell Range, or Cell Style.
; Syntax ........: __LOCalc_CellFont(ByRef $oObj, $sFontName, $nFontSize, $iPosture, $iWeight)
; Parameters ....: $oObj                - [in/out] an object. A Cell, Cell Range or Cell Style Object returned from an applicable function.
;                  $sFontName           - a string value. The Font Name to use.
;                  $nFontSize           - a general number value. The new Font size.
;                  $iPosture            - an integer value (0-5). The Font Italic setting. See Constants, $LOC_POSTURE_* as defined in LibreOfficeCalc_Constants.au3. Also see remarks.
;                  $iWeight             - an integer value (0, 50-200). The Font Bold settings see Constants, $LOC_WEIGHT_* as defined in LibreOfficeCalc_Constants.au3. Also see remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 4 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 5 Return 0 = $sFontName not a String.
;                  @Error 1 @Extended 6 Return 0 = $nFontSize not a number.
;                  @Error 1 @Extended 7 Return 0 = $iPosture not an Integer, less than 0, or greater than 5. See Constants, $LOC_POSTURE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $iWeight not an Integer, less than 50 but not equal to 0, or greater than 200. See Constants, $LOC_WEIGHT_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sFontName
;                  |                               2 = Error setting $nFontSize
;                  |                               4 = Error setting $iPosture
;                  |                               8 = Error setting $iWeight
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Not every font accepts Bold and Italic settings, and not all settings for bold and Italic are accepted, such as oblique, ultra Bold etc.
;                  Libre Calc accepts only the predefined weight values, any other values are changed automatically to an acceptable value, which could trigger a settings error.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CellFont(ByRef $oObj, $sFontName, $nFontSize, $iPosture, $iWeight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
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
		If Not __LO_IntIsBetween($iPosture, $LOC_POSTURE_NONE, $LOC_POSTURE_ITALIC) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oObj.CharPosture = $iPosture
		$iError = ($oObj.CharPosture() = $iPosture) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iWeight <> Null) Then
		If Not __LO_IntIsBetween($iWeight, $LOC_WEIGHT_THIN, $LOC_WEIGHT_BLACK, "", $LOC_WEIGHT_DONT_KNOW) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$oObj.CharWeight = $iWeight
		$iError = ($oObj.CharWeight() = $iWeight) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOCalc_CellFont

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CellFontColor
; Description ...: Internal function to Set or Retrieve the Font Color for a Cell, Cell Range, or Cell Style.
; Syntax ........: __LOCalc_CellFontColor(ByRef $oObj, $iFontColor)
; Parameters ....: $oObj                - [in/out] an object. A Cell, Cell Range or Cell Style Object returned from an applicable function.
;                  $iFontColor          - an integer value (-1-16777215). The Color value in Long Integer format to make the font, can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Set to $LO_COLOR_OFF(-1) for Auto color.
; Return values .: Success: 1 or Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iFontColor not an Integer, less than 0, or greater than 16777215.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iFontColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning current Font Color as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Though Transparency is present on the Font Effects page in the UI, there is (as best as I can find) no setting for it available to read and modify. And further, it seems even in L.O. the setting does not affect the font's transparency, though it may change the color value.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CellFontColor(ByRef $oObj, $iFontColor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($iFontColor) Then

		Return SetError($__LO_STATUS_SUCCESS, 1, $oObj.CharColor())
	EndIf

	If ($iFontColor <> Null) Then
		If Not __LO_IntIsBetween($iFontColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.CharColor = $iFontColor
		$iError = ($oObj.CharColor() = $iFontColor) ? ($iError) : (BitOR($iError, 1))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOCalc_CellFontColor

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CellNumberFormat
; Description ...: Internal function to Set or Retrieve Cell, Cell Range, or Cell Style Number Format settings.
; Syntax ........: __LOCalc_CellNumberFormat(ByRef $oDoc, ByRef $oObj, $iFormatKey)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $oObj                - [in/out] an object. A Cell, Cell Range or Cell Style Object returned from an applicable function.
;                  $iFormatKey          - an integer value. A Format Key from a previous _LOCalc_FormatKeyCreate or _LOCalc_FormatKeysGetList function.
; Return values .: Success: 1 or Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 4 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 5 Return 0 = $iFormatKey not an Integer.
;                  @Error 1 @Extended 6 Return 0 = Format Key called in $iFormatKey not found in document.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iFormatKey
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current setting as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CellNumberFormat(ByRef $oDoc, ByRef $oObj, $iFormatKey)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If __LO_VarsAreNull($iFormatKey) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oObj.NumberFormat())

	If Not IsInt($iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not _LOCalc_FormatKeyExists($oDoc, $iFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oObj.NumberFormat = $iFormatKey
	$iError = ($oObj.NumberFormat() = $iFormatKey) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOCalc_CellNumberFormat

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CellOverLine
; Description ...: Internal function to Set and retrieve the OverLine settings for a Cell, Cell Range, or Cell Style.
; Syntax ........: __LOCalc_CellOverLine(ByRef $oObj, $bWordOnly, $iOverLineStyle, $bOLHasColor, $iOLColor)
; Parameters ....: $oObj                - [in/out] an object. A Cell, Cell Range or Cell Style Object returned from an applicable function.
;                  $bWordOnly           - a boolean value. If true, white spaces are not Overlined.
;                  $iOverLineStyle      - an integer value (0-18). The style of the Overline line, see constants, $LOC_UNDERLINE_* as defined in LibreOfficeCalc_Constants.au3. See Remarks.
;                  $bOLHasColor         - a boolean value. If True, the Overline is colored, must be set to true in order to set the Overline color.
;                  $iOLColor            - an integer value (-1-16777215). The Overline color, set in Long integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Set to $LO_COLOR_OFF(-1) for automatic color mode.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iOverLineStyle not an Integer, less than 0, or greater than 18. See constants, $LOC_UNDERLINE_* as defined in LibreOfficeCalc_Constants.au3. See Remarks.
;                  @Error 1 @Extended 6 Return 0 = $bOLHasColor not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $iOLColor not an Integer, less than -1, or greater than 16777215.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bWordOnly
;                  |                               2 = Error setting $iOverLineStyle
;                  |                               4 = Error setting $bOLHasColor
;                  |                               8 = Error setting $iOLColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Overline line style uses the same constants as underline style.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CellOverLine(ByRef $oObj, $bWordOnly, $iOverLineStyle, $bOLHasColor, $iOLColor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
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
		If Not __LO_IntIsBetween($iOverLineStyle, $LOC_UNDERLINE_NONE, $LOC_UNDERLINE_BOLD_WAVE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

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
EndFunc   ;==>__LOCalc_CellOverLine

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CellProtection
; Description ...: Internal function to Set or Retrieve Cell, Cell Range, or Cell Style protection settings.
; Syntax ........: __LOCalc_CellProtection(ByRef $oObj, $bHideAll, $bProtected, $bHideFormula, $bHideWhenPrint)
; Parameters ....: $oObj                - [in/out] an object. A Cell, Cell Range or Cell Style Object returned from an applicable function.
;                  $bHideAll            - a boolean value. If True, Hides formulas and contents of the cell.
;                  $bProtected          - a boolean value. If True, Prevents the cell from being modified.
;                  $bHideFormula        - a boolean value. If True, Hides formulas in the cell.
;                  $bHideWhenPrint      - a boolean value. If True, the cell is kept from being printed.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bHideAll not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bProtected not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bHideFormula not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bHideWhenPrint not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Cell Protection Structure.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bHideAll
;                  |                               2 = Error setting $bProtected
;                  |                               4 = Error setting $bHideFormula
;                  |                               8 = Error setting $bHideWhenPrint
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Cell protection only takes effect if you also protect the sheet.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CellProtection(ByRef $oObj, $bHideAll, $bProtected, $bHideFormula, $bHideWhenPrint)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $abProtection[4]
	Local $tCellProtection

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tCellProtection = $oObj.CellProtection()
	If Not IsObj($tCellProtection) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($bHideAll, $bProtected, $bHideFormula, $bHideWhenPrint) Then
		__LO_ArrayFill($abProtection, $tCellProtection.IsHidden(), $tCellProtection.IsLocked(), $tCellProtection.IsFormulaHidden(), $tCellProtection.IsPrintHidden())

		Return SetError($__LO_STATUS_SUCCESS, 1, $abProtection)
	EndIf

	If ($bHideAll <> Null) Then
		If Not IsBool($bHideAll) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tCellProtection.IsHidden = $bHideAll
	EndIf

	If ($bProtected <> Null) Then
		If Not IsBool($bProtected) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$tCellProtection.IsLocked = $bProtected
	EndIf

	If ($bHideFormula <> Null) Then
		If Not IsBool($bHideFormula) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tCellProtection.IsFormulaHidden = $bHideFormula
	EndIf

	If ($bHideWhenPrint <> Null) Then
		If Not IsBool($bHideWhenPrint) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tCellProtection.IsPrintHidden = $bHideWhenPrint
	EndIf

	$oObj.CellProtection = $tCellProtection

	$iError = ($bHideAll = Null) ? ($iError) : ($oObj.CellProtection.IsHidden() = $bHideAll) ? ($iError) : (BitOR($iError, 1))
	$iError = ($bProtected = Null) ? ($iError) : ($oObj.CellProtection.IsLocked() = $bProtected) ? ($iError) : (BitOR($iError, 2))
	$iError = ($bHideFormula = Null) ? ($iError) : ($oObj.CellProtection.IsFormulaHidden() = $bHideFormula) ? ($iError) : (BitOR($iError, 4))
	$iError = ($bHideWhenPrint = Null) ? ($iError) : ($oObj.CellProtection.IsPrintHidden() = $bHideWhenPrint) ? ($iError) : (BitOR($iError, 8))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOCalc_CellProtection

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CellShadow
; Description ...: Internal function to Set or Retrieve the Shadow settings for a Cell, Cell Range, or Cell style.
; Syntax ........: __LOCalc_CellShadow(ByRef $oObj, $iWidth, $iColor, $bTransparent, $iLocation)
; Parameters ....: $oObj                - [in/out] an object. A Cell, Cell Range or Cell Style Object returned from an applicable function.
;                  $iWidth              - an integer value (0-5009). The shadow width, set in Micrometers.
;                  $iColor              - an integer value (0-16777215). The color of the shadow, set in Long Integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $bTransparent        - a boolean value. If True, the shadow is transparent.
;                  $iLocation           - an integer value (0-4). The location of the shadow compared to the Cell. See Constants, $LOC_SHADOW_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iWidth not an Integer, less than 0, or greater than 5009.
;                  @Error 1 @Extended 5 Return 0 = $iColor not an Integer, less than 0, or greater than 16,777,215.
;                  @Error 1 @Extended 6 Return 0 = $bTransparent not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $iLocation not an Integer, less than 0, or greater than 4. See Constants, $LOC_SHADOW_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Shadow Format Structure.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iWidth
;                  |                               2 = Error setting $iColor
;                  |                               4 = Error setting $bTransparent
;                  |                               8 = Error setting $iLocation
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CellShadow(ByRef $oObj, $iWidth, $iColor, $bTransparent, $iLocation)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $tShdwFrmt
	Local $avShadow[4]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$tShdwFrmt = $oObj.ShadowFormat()
	If Not IsObj($tShdwFrmt) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($iWidth, $iColor, $bTransparent, $iLocation) Then
		__LO_ArrayFill($avShadow, $tShdwFrmt.ShadowWidth(), $tShdwFrmt.Color(), $tShdwFrmt.IsTransparent(), $tShdwFrmt.Location())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avShadow)
	EndIf

	If ($iWidth <> Null) Then
		If Not __LO_IntIsBetween($iWidth, 0, 5009) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

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
		If Not __LO_IntIsBetween($iLocation, $LOC_SHADOW_NONE, $LOC_SHADOW_BOTTOM_RIGHT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tShdwFrmt.Location = $iLocation
	EndIf

	$oObj.ShadowFormat = $tShdwFrmt

	$iError = ($iWidth = Null) ? ($iError) : ((__LO_IntIsBetween($oObj.ShadowFormat.ShadowWidth(), $iWidth - 1, $iWidth + 1)) ? ($iError) : (BitOR($iError, 1)))
	$iError = ($iColor = Null) ? ($iError) : (($oObj.ShadowFormat.Color() = $iColor) ? ($iError) : (BitOR($iError, 2)))
	$iError = ($bTransparent = Null) ? ($iError) : (($oObj.ShadowFormat.IsTransparent() = $bTransparent) ? ($iError) : (BitOR($iError, 4)))
	$iError = ($iLocation = Null) ? ($iError) : (($oObj.ShadowFormat.Location() = $iLocation) ? ($iError) : (BitOR($iError, 8)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOCalc_CellShadow

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CellStrikeOut
; Description ...: Internal function to Set or Retrieve the Strikeout settings for a Cell, Cell Range, or Cell Style.
; Syntax ........: __LOCalc_CellStrikeOut(ByRef $oObj, $bWordOnly, $bStrikeOut, $iStrikeLineStyle)
; Parameters ....: $oObj                - [in/out] an object. A Cell, Cell Range or Cell Style Object returned from an applicable function.
;                  $bWordOnly           - a boolean value. If True, strike out is applied to words only, skipping whitespaces.
;                  $bStrikeOut          - a boolean value. If True, strikeout is applied to characters.
;                  $iStrikeLineStyle    - an integer value (0-6). The Strikeout Line Style, see constants, $LOC_STRIKEOUT_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bStrikeOut not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iStrikeLineStyle not an Integer, less than 0 or greater than 6. See constants, $LOC_STRIKEOUT_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bWordOnly
;                  |                               2 = Error setting $bStrikeOut
;                  |                               4 = Error setting $iStrikeLineStyle
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CellStrikeOut(ByRef $oObj, $bWordOnly, $bStrikeOut, $iStrikeLineStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
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
		If Not __LO_IntIsBetween($iStrikeLineStyle, $LOC_STRIKEOUT_NONE, $LOC_STRIKEOUT_X) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.CharStrikeout = $iStrikeLineStyle
		$iError = ($oObj.CharStrikeout() = $iStrikeLineStyle) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOCalc_CellStrikeOut

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CellStyleBorder
; Description ...: Internal function to Set and Retrieve the Cell Style Border Line Width, Style, and Color. Libre Office Version 3.6 and Up.
; Syntax ........: __LOCalc_CellStyleBorder(ByRef $oCellStyle, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight, $iTLBRDiag, $iBLTRDiag)
; Parameters ....: $oCellStyle          - [in/out] an object. A Cell Style object returned by a previous _LOCalc_CellStyleCreate, or _LOCalc_CellStyleGetObj function.
;                  $bWid                - a boolean value. If True, Border Width is being modified. Only one can be True at once.
;                  $bSty                - a boolean value. If True, Border Style is being modified. Only one can be True at once.
;                  $bCol                - a boolean value. If True, Border Color is being modified. Only one can be True at once.
;                  $iTop                - an integer value. Modifies the top border line settings. See Width, Style or Color functions for values.
;                  $iBottom             - an integer value. Modifies the bottom border line settings. See Width, Style or Color functions for values.
;                  $iLeft               - an integer value. Modifies the left border line settings. See Width, Style or Color functions for values.
;                  $iRight              - an integer value. Modifies the right border line settings. See Width, Style or Color functions for values.
;                  $iTLBRDiag           - an integer value. Modifies the top-left to bottom-right diagonal border line settings. See Width, Style or Color functions for values.
;                  $iBLTRDiag           - an integer value. Modifies the bottom-left to top-right diagonal border line settings. See Width, Style or Color functions for values.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 9 Return 0 = Variable passed to internal function not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style/Color when Top Border width not set.
;                  @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Style/Color when Bottom Border width not set.
;                  @Error 4 @Extended 3 Return 0 = Cannot set Left Border Style/Color when Left Border width not set.
;                  @Error 4 @Extended 4 Return 0 = Cannot set Right Border Style/Color when Right Border width not set.
;                  @Error 4 @Extended 5 Return 0 = Cannot set Top-Left to Bottom-Right Diagonal Border Style/Color when Top-Left to Bottom-Right Diagonal Border width not set.
;                  @Error 4 @Extended 6 Return 0 = Cannot set Bottom-Left to Top-Right Diagonal Border Style/Color when Bottom-Left to Top-Right Diagonal Border width not set.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CellStyleBorder(ByRef $oCellStyle, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight, $iTLBRDiag, $iBLTRDiag)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avBorder[6]
	Local $tBL2

	If Not __LO_VersionCheck(3.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oCellStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
	If (($bWid + $bSty + $bCol) <> 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; If more than one Boolean is true = error

	If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight, $iTLBRDiag, $iBLTRDiag) Then
		If $bWid Then
			__LO_ArrayFill($avBorder, $oCellStyle.TopBorder2.LineWidth(), $oCellStyle.BottomBorder2.LineWidth(), $oCellStyle.LeftBorder2.LineWidth(), $oCellStyle.RightBorder2.LineWidth(), _
					$oCellStyle.DiagonalTLBR2.LineWidth(), $oCellStyle.DiagonalBLTR2.LineWidth())

		ElseIf $bSty Then
			__LO_ArrayFill($avBorder, $oCellStyle.TopBorder2.LineStyle(), $oCellStyle.BottomBorder2.LineStyle(), $oCellStyle.LeftBorder2.LineStyle(), $oCellStyle.RightBorder2.LineStyle(), _
					$oCellStyle.DiagonalTLBR2.LineStyle(), $oCellStyle.DiagonalBLTR2.LineStyle())

		ElseIf $bCol Then
			__LO_ArrayFill($avBorder, $oCellStyle.TopBorder2.Color(), $oCellStyle.BottomBorder2.Color(), $oCellStyle.LeftBorder2.Color(), $oCellStyle.RightBorder2.Color(), _
					$oCellStyle.DiagonalTLBR2.Color(), $oCellStyle.DiagonalBLTR2.Color())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avBorder)
	EndIf

	$tBL2 = __LO_CreateStruct("com.sun.star.table.BorderLine2")
	If Not IsObj($tBL2) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If $iTop <> Null Then
		If Not $bWid And ($oCellStyle.TopBorder2.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0) ; If Width not set, cant set color or style.

		; Top Line
		$tBL2.LineWidth = ($bWid) ? ($iTop) : ($oCellStyle.TopBorder2.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iTop) : ($oCellStyle.TopBorder2.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iTop) : ($oCellStyle.TopBorder2.Color()) ; copy Color over to new size structure
		$oCellStyle.TopBorder2 = $tBL2
	EndIf

	If $iBottom <> Null Then
		If Not $bWid And ($oCellStyle.BottomBorder2.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 2, 0) ; If Width not set, cant set color or style.

		; Bottom Line
		$tBL2.LineWidth = ($bWid) ? ($iBottom) : ($oCellStyle.BottomBorder2.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iBottom) : ($oCellStyle.BottomBorder2.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iBottom) : ($oCellStyle.BottomBorder2.Color()) ; copy Color over to new size structure
		$oCellStyle.BottomBorder2 = $tBL2
	EndIf

	If $iLeft <> Null Then
		If Not $bWid And ($oCellStyle.LeftBorder2.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 3, 0) ; If Width not set, cant set color or style.

		; Left Line
		$tBL2.LineWidth = ($bWid) ? ($iLeft) : ($oCellStyle.LeftBorder2.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iLeft) : ($oCellStyle.LeftBorder2.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iLeft) : ($oCellStyle.LeftBorder2.Color()) ; copy Color over to new size structure
		$oCellStyle.LeftBorder2 = $tBL2
	EndIf

	If $iRight <> Null Then
		If Not $bWid And ($oCellStyle.RightBorder2.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 4, 0) ; If Width not set, cant set color or style.

		; Right Line
		$tBL2.LineWidth = ($bWid) ? ($iRight) : ($oCellStyle.RightBorder2.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iRight) : ($oCellStyle.RightBorder2.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iRight) : ($oCellStyle.RightBorder2.Color()) ; copy Color over to new size structure
		$oCellStyle.RightBorder2 = $tBL2
	EndIf

	If $iTLBRDiag <> Null Then
		If Not $bWid And ($oCellStyle.DiagonalTLBR2.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 5, 0) ; If Width not set, cant set color or style.

		; Right Line
		$tBL2.LineWidth = ($bWid) ? ($iTLBRDiag) : ($oCellStyle.DiagonalTLBR2.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iTLBRDiag) : ($oCellStyle.DiagonalTLBR2.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iTLBRDiag) : ($oCellStyle.DiagonalTLBR2.Color()) ; copy Color over to new size structure
		$oCellStyle.DiagonalTLBR2 = $tBL2
	EndIf

	If $iBLTRDiag <> Null Then
		If Not $bWid And ($oCellStyle.DiagonalBLTR2.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 6, 0) ; If Width not set, cant set color or style.

		; Right Line
		$tBL2.LineWidth = ($bWid) ? ($iBLTRDiag) : ($oCellStyle.DiagonalBLTR2.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iBLTRDiag) : ($oCellStyle.DiagonalBLTR2.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iBLTRDiag) : ($oCellStyle.DiagonalBLTR2.Color()) ; copy Color over to new size structure
		$oCellStyle.DiagonalBLTR2 = $tBL2
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOCalc_CellStyleBorder

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CellTextAlign
; Description ...: Internal function to Set and Retrieve text Alignment settings for a Cell, Cell Range, or Cell style.
; Syntax ........: __LOCalc_CellTextAlign(ByRef $oObj, $iHoriAlign, $iVertAlign, $iIndent)
; Parameters ....: $oObj                - [in/out] an object. A Cell, Cell Range or Cell Style Object returned from an applicable function.
;                  $iHoriAlign          - an integer value (0-6). The Horizontal alignment of the text. See Constants, $LOC_CELL_ALIGN_HORI_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iVertAlign          - an integer value (0-5). The Vertical alignment of the text. See Constants, $LOC_CELL_ALIGN_VERT_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iIndent             - an integer value. The amount of indentation from the left side of the cell, in micrometers.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iHoriAlign not an Integer, less than 0, or greater than 6. See Constants, $LOC_CELL_ALIGN_HORI_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $iVertAlign not an Integer, less than 0, or greater than 5. See Constants, $LOC_CELL_ALIGN_VERT_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $iIndent not an Integer.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iHoriAlign
;                  |                               2 = Error setting $iVertAlign
;                  |                               4 = Error setting $iIndent
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CellTextAlign(ByRef $oObj, $iHoriAlign, $iVertAlign, $iIndent)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local Const $iAlignNoDistribute = 0, $iAlignDistribute = 1
	Local $aiAlign[3]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($iHoriAlign, $iVertAlign, $iIndent) Then
		__LO_ArrayFill($aiAlign, $oObj.HoriJustify(), $oObj.VertJustify(), $oObj.ParaIndent())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiAlign)
	EndIf

	If ($iHoriAlign <> Null) Then
		If Not __LO_IntIsBetween($iHoriAlign, $LOC_CELL_ALIGN_HORI_DEFAULT, $LOC_CELL_ALIGN_HORI_DISTRIBUTED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		; $LOC_CELL_ALIGN_HORI_DISTRIBUTED Isn't a real setting, it is a combination of Filled (Block) and an undocumented setting called "HoriJustifyMethod" set to 1, instead of 0.

		If ($iHoriAlign = $LOC_CELL_ALIGN_HORI_DISTRIBUTED) Then
			$oObj.HoriJustifyMethod = $iAlignDistribute
			$oObj.HoriJustify = $LOC_CELL_ALIGN_HORI_FILLED
			$iError = (($oObj.HoriJustify() = $LOC_CELL_ALIGN_HORI_FILLED) And ($oObj.HoriJustifyMethod() = $iAlignDistribute)) ? ($iError) : (BitOR($iError, 1))

		Else
			$oObj.HoriJustifyMethod = $iAlignNoDistribute
			$oObj.HoriJustify = $iHoriAlign
			$iError = ($oObj.HoriJustify() = $iHoriAlign) ? ($iError) : (BitOR($iError, 1))
		EndIf
	EndIf

	If ($iVertAlign <> Null) Then
		If Not __LO_IntIsBetween($iVertAlign, $LOC_CELL_ALIGN_VERT_DEFAULT, $LOC_CELL_ALIGN_VERT_DISTRIBUTED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		; $LOC_CELL_ALIGN_VERT_DISTRIBUTED Isn't a real setting, it is a combination of Filled (Block) and an undocumented setting called "VertJustifyMethod" set to 1, instead of 0.

		If ($iVertAlign = $LOC_CELL_ALIGN_VERT_DISTRIBUTED) Then
			$oObj.VertJustifyMethod = $iAlignDistribute
			$oObj.VertJustify = $LOC_CELL_ALIGN_VERT_JUSTIFIED
			$iError = (($oObj.VertJustify() = $LOC_CELL_ALIGN_VERT_JUSTIFIED) And ($oObj.VertJustifyMethod() = $iAlignDistribute)) ? ($iError) : (BitOR($iError, 2))

		Else
			$oObj.VertJustifyMethod = $iAlignNoDistribute
			$oObj.VertJustify = $iVertAlign
			$iError = ($oObj.VertJustify() = $iVertAlign) ? ($iError) : (BitOR($iError, 2))
		EndIf
	EndIf

	If ($iIndent <> Null) Then
		If Not IsInt($iIndent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.ParaIndent = $iIndent
		$iError = (__LO_IntIsBetween($oObj.ParaIndent(), $iIndent - 1, $iIndent + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOCalc_CellTextAlign

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CellTextOrient
; Description ...: Internal function to Set or Retrieve Text Orientation settings for a Cell, Cell Range, or Cell Style.
; Syntax ........: __LOCalc_CellTextOrient(ByRef $oObj, $iRotate, $iReference, $bVerticalStack, $bAsianLayout)
; Parameters ....: $oObj                - [in/out] an object. A Cell, Cell Range or Cell Style Object returned from an applicable function.
;                  $iRotate             - an integer value (0-359). The rotation angle of the text.
;                  $iReference          - an integer value (0,1,3). The cell edge from which to write the rotated text. See Constants $LOC_CELL_ROTATE_REF_* as defined in LibreOfficeCalc_Constants.au3.
;                  $bVerticalStack      - a boolean value. If True, Aligns text vertically. Only available after you enable support for Asian languages in Libre Office settings.
;                  $bAsianLayout        - a boolean value. If True, Aligns Asian characters one below the other. Only available after you enable support for Asian languages in Libre Office settings, and enable vertical text.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iRotate not an Integer, less than 0, or greater than 359.
;                  @Error 1 @Extended 5 Return 0 = $iReference not an Integer, less than 0, or greater than 1 but not equal to 3. See Constants $LOC_CELL_ROTATE_REF_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 6 Return 0 = $bVerticalStack not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $bAsianLayout not a Boolean.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iRotate
;                  |                               2 = Error setting $iReference
;                  |                               4 = Error setting $bVerticalStack
;                  |                               8 = Error setting $bAsianLayout
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CellTextOrient(ByRef $oObj, $iRotate, $iReference, $bVerticalStack, $bAsianLayout)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local Const $__iIsNotStacked = 0, $__iIsStacked = 3
	Local $avOrient[4]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($iRotate, $iReference, $bVerticalStack, $bAsianLayout) Then
		__LO_ArrayFill($avOrient, Int($oObj.RotateAngle() / 100), $oObj.RotateReference(), (($oObj.Orientation() = $__iIsStacked) ? (True) : (False)), $oObj.AsianVerticalMode())
		; Rotate Angle is in 100ths of degrees.
		; When Vertical Stack is True, Orientation is set to 3, when false, it is set to 0.

		Return SetError($__LO_STATUS_SUCCESS, 1, $avOrient)
	EndIf

	If ($iRotate <> Null) Then
		If Not __LO_IntIsBetween($iRotate, 0, 359) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.RotateAngle = Int($iRotate * 100) ; Rotate Angle is in 100ths of degrees.
		$iError = ($oObj.RotateAngle = Int($iRotate * 100)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iReference <> Null) Then
		If Not __LO_IntIsBetween($iReference, $LOC_CELL_ALIGN_VERT_DEFAULT, $LOC_CELL_ALIGN_VERT_TOP, "", $LOC_CELL_ALIGN_VERT_BOTTOM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.RotateReference = $iReference
		$iError = ($oObj.RotateReference() = $iReference) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bVerticalStack <> Null) Then
		If Not IsBool($bVerticalStack) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		; According to Libre Office IDL Vertical Stack (Orientation set to 3) is only taken into account when RotateAngle is set to 0.
		If ($bVerticalStack = True) Then
			$oObj.RotateAngle = 0
			$oObj.Orientation = $__iIsStacked
			$iError = ($oObj.Orientation() = $__iIsStacked) ? ($iError) : (BitOR($iError, 4))

		Else
			$oObj.Orientation = $__iIsNotStacked
			$iError = ($oObj.Orientation() = $__iIsNotStacked) ? ($iError) : (BitOR($iError, 4))
		EndIf
	EndIf

	If ($bAsianLayout <> Null) Then
		If Not IsBool($bAsianLayout) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oObj.AsianVerticalMode = $bAsianLayout
		$iError = ($oObj.AsianVerticalMode() = $bAsianLayout) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOCalc_CellTextOrient

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CellTextProperties
; Description ...: Internal function to Set or Retrieve Text property settings for a Cell, Cell Range, or Cell Style.
; Syntax ........: __LOCalc_CellTextProperties(ByRef $oObj, $bAutoWrapText, $bHyphen, $bShrinkToFit, $iTextDirection)
; Parameters ....: $oObj                - [in/out] an object. A Cell, Cell Range or Cell Style Object returned from an applicable function.
;                  $bAutoWrapText       - a boolean value. If True, Wraps text onto another line at the cell border.
;                  $bHyphen             - a boolean value. If True, Enables word hyphenation for text wrapping to the next line.
;                  $bShrinkToFit        - a boolean value. If True, Reduces the apparent size of the font so that the contents of the cell fit into the current cell width.
;                  $iTextDirection      - an integer value (0,1,4). The Text Writing Direction. See Constants, $LOC_TXT_DIR_* as defined in LibreOfficeCalc_Constants.au3. [Libre Office Default is 4]
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bAutoWrapText not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bHyphen not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bShrinkToFitnot a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $iTextDirection not an Integer, less than 0, or greater than 1 but not equal to 4. See Constants, $LOC_TXT_DIR_* as defined in LibreOfficeCalc_Constants.au3. [Libre Office Default is 4]
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bAutoWrapText
;                  |                               2 = Error setting $bHyphen
;                  |                               4 = Error setting $bShrinkToFit
;                  |                               8 = Error setting $iTextDirection
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CellTextProperties(ByRef $oObj, $bAutoWrapText, $bHyphen, $bShrinkToFit, $iTextDirection)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avTextProp[4]

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LO_VarsAreNull($bAutoWrapText, $bHyphen, $bShrinkToFit, $iTextDirection) Then
		__LO_ArrayFill($avTextProp, $oObj.IsTextWrapped(), $oObj.ParaIsHyphenation(), $oObj.ShrinkToFit(), $oObj.WritingMode())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avTextProp)
	EndIf

	If ($bAutoWrapText <> Null) Then
		If Not IsBool($bAutoWrapText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oObj.IsTextWrapped = $bAutoWrapText
		$iError = ($oObj.IsTextWrapped = $bAutoWrapText) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($bHyphen <> Null) Then
		If Not IsBool($bHyphen) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oObj.ParaIsHyphenation = $bHyphen
		$iError = ($oObj.ParaIsHyphenation() = $bHyphen) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bShrinkToFit <> Null) Then
		If Not IsBool($bShrinkToFit) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$oObj.ShrinkToFit = $bShrinkToFit
		$iError = ($oObj.ShrinkToFit() = $bShrinkToFit) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iTextDirection <> Null) Then
		If Not __LO_IntIsBetween($iTextDirection, $LOC_TXT_DIR_LR, $LOC_TXT_DIR_RL, "", $LOC_TXT_DIR_CONTEXT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$oObj.WritingMode = $iTextDirection
		$iError = ($oObj.WritingMode() = $iTextDirection) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>__LOCalc_CellTextProperties

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CellUnderLine
; Description ...: Internal function to Set and retrieve the Underline settings for a Cell, Cell Range, or Cell Style.
; Syntax ........: __LOCalc_CellUnderLine(ByRef $oObj, $bWordOnly, $iUnderLineStyle, $bULHasColor, $iULColor)
; Parameters ....: $oObj                - [in/out] an object. A Cell, Cell Range or Cell Style Object returned from an applicable function.
;                  $bWordOnly           - a boolean value. If true, white spaces are not underlined.
;                  $iUnderLineStyle     - an integer value (0-18). The Underline line style, see constants, $LOC_UNDERLINE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $bULHasColor         - a boolean value. If True, the underline is colored, must be set to true in order to set the underline color.
;                  $iULColor            - an integer value (-1-16777215). The color of the underline, set in Long integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Set to $LO_COLOR_OFF(-1) for automatic color mode.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iUnderLineStyle not an Integer, less than 0, or greater than 18. See constants, $LOC_UNDERLINE_* as defined in LibreOfficeCalc_Constants.au3. See Remarks.
;                  @Error 1 @Extended 6 Return 0 = $bULHasColor not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $iULColor not an Integer, less than -1, or greater than 16777215.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bWordOnly
;                  |                               2 = Error setting $iUnderLineStyle
;                  |                               4 = Error setting $bULHasColor
;                  |                               8 = Error setting $iULColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CellUnderLine(ByRef $oObj, $bWordOnly, $iUnderLineStyle, $bULHasColor, $iULColor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
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
		If Not __LO_IntIsBetween($iUnderLineStyle, $LOC_UNDERLINE_NONE, $LOC_UNDERLINE_BOLD_WAVE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

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
EndFunc   ;==>__LOCalc_CellUnderLine

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CharPosition
; Description ...: Set and retrieve settings related to Sub/Super Script and relative size.
; Syntax ........: __LOCalc_CharPosition(ByRef $oObj, $bAutoSuper, $iSuperScript, $bAutoSub, $iSubScript, $iRelativeSize)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.style.CharacterProperties".
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
;                  The way LibreOffice is set up Super/Subscript are set in the same setting, Super is a positive number from 1 to 100 (percentage), Subscript is a negative number set to 1 to 100 percentage. For the user's convenience this function accepts both positive and negative numbers for Subscript, if a positive number is called for Subscript, it is automatically set to a negative.
;                  Automatic Superscript has a integer value of 14000, Auto Subscript has a integer value of -14000. There is no settable setting of Automatic Super/Sub Script, though one exists, it is read-only in LibreOffice, consequently I have made two separate parameters to be able to determine if the user wants to automatically set Superscript or Subscript.
;                  If you set both Auto Superscript to True and Auto Subscript to True, or $iSuperScript to an integer and $iSubScript to an integer, Subscript will be set as it is the last in the line to be set in this function, and thus will over-write any Superscript settings.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CharPosition(ByRef $oObj, $bAutoSuper, $iSuperScript, $bAutoSub, $iSubScript, $iRelativeSize)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
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
EndFunc   ;==>__LOCalc_CharPosition

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CharSpacing
; Description ...: Set and retrieve the spacing between characters (Kerning).
; Syntax ........: __LOCalc_CharSpacing(ByRef $oObj, $bAutoKerning, $nKerning)
; Parameters ....: $oObj                - [in/out] an object. An Object that supports "com.sun.star.style.CharacterProperties".
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
;                  When setting Kerning values in LibreOffice, the measurement is listed in Pt (Printer's Points) in the User Display, however the internal setting is measured in Micrometers. They will be automatically converted from Points to Micrometers and back for retrieval of settings.
;                  The acceptable values for $nKerning are from -2 Pt to 928.8 Pt.
;                  The values can be directly converted easily, however, for an unknown reason to myself, LibreOffice begins counting backwards and in negative Micrometers internally from 928.9 up to 1000 Pt (Max setting).
;                  For example, 928.8Pt is the last correct value, which equals 32766 uM (Micrometers), after this LibreOffice reports the following: ;928.9 Pt = -32766 uM; 929 Pt = -32763 uM; 929.1 = -32759; 1000 pt = -30258. Attempting to set Libre's kerning value to anything over 32768 uM causes a COM exception, and attempting to set the kerning to any of these negative numbers sets the User viewable kerning value to -2.0 Pt. For these reasons the max settable kerning is -2.0 Pt to 928.8 Pt.
; Related .......: _LO_ConvertFromMicrometer, _LO_ConvertToMicrometer
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CharSpacing(ByRef $oObj, $bAutoKerning, $nKerning)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
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
EndFunc   ;==>__LOCalc_CharSpacing

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CommentAreaShadowModify
; Description ...: Internal function for setting or retrieving Comment Shadow Location and Distance settings.
; Syntax ........: __LOCalc_CommentAreaShadowModify($oAnnotationShape[, $iLocation = Null[, $iDistance = Null]])
; Parameters ....: $oAnnotationShape    - an object. A Annotation Shape Object retrieved from a Comment.
;                  $iLocation           - [optional] an integer value. Default is Null. The Location of the Shadow, must be one of the Constants, $LOC_COMMENT_SHADOW_* as defined in LibreOfficeCalc_Constants.au3..
;                  $iDistance           - [optional] an integer value. Default is Null. The distance of the Shadow from the Comment box, set in Micrometers.
; Return values .: Success: 1 or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oAnnotationShape not an Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iLocation
;                  |                               2 = Error setting $iDistance
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Successfully set the settings.
;                  @Error 0 @Extended ? Return Integer = Success. $iLocation and $iDistance set to Null, returning current Values. Return will be current distance, and @Extended will be the current Location.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CommentAreaShadowModify($oAnnotationShape, $iLocation = Null, $iDistance = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $bReturn = False, $bModifyLocation = True
	Local $iError = 1

	If Not IsObj($oAnnotationShape) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iLocation, $iDistance) Then $bReturn = True

	If ($iLocation = Null) Then ; Determine current location)
		$bModifyLocation = False
		$iError = 2
		Select
			Case (($oAnnotationShape.ShadowXDistance() < 0) And ($oAnnotationShape.ShadowYDistance() < 0)) ; Top Left.
				$iLocation = $LOC_COMMENT_ANCHOR_TOP_LEFT

			Case (($oAnnotationShape.ShadowXDistance() = 0) And ($oAnnotationShape.ShadowYDistance() < 0)) ; Top Center
				$iLocation = $LOC_COMMENT_ANCHOR_TOP_CENTER

			Case (($oAnnotationShape.ShadowXDistance() > 0) And ($oAnnotationShape.ShadowYDistance() < 0)) ; Top Right
				$iLocation = $LOC_COMMENT_ANCHOR_TOP_RIGHT

			Case (($oAnnotationShape.ShadowXDistance() < 0) And ($oAnnotationShape.ShadowYDistance() = 0)) ; Middle Left
				$iLocation = $LOC_COMMENT_ANCHOR_MIDDLE_LEFT

			Case (($oAnnotationShape.ShadowXDistance() = 0) And ($oAnnotationShape.ShadowYDistance() = 0)) ; Middle Center
				$iLocation = $LOC_COMMENT_ANCHOR_MIDDLE_CENTER

			Case (($oAnnotationShape.ShadowXDistance() > 0) And ($oAnnotationShape.ShadowYDistance() = 0)) ; Middle Right
				$iLocation = $LOC_COMMENT_ANCHOR_MIDDLE_RIGHT

			Case (($oAnnotationShape.ShadowXDistance() < 0) And ($oAnnotationShape.ShadowYDistance() > 0)) ; Bottom Left
				$iLocation = $LOC_COMMENT_ANCHOR_BOTTOM_LEFT

			Case (($oAnnotationShape.ShadowXDistance() = 0) And ($oAnnotationShape.ShadowYDistance() > 0)) ; Bottom Center
				$iLocation = $LOC_COMMENT_ANCHOR_BOTTOM_CENTER

			Case (($oAnnotationShape.ShadowXDistance() > 0) And ($oAnnotationShape.ShadowYDistance() > 0)) ; Bottom Right
				$iLocation = $LOC_COMMENT_ANCHOR_BOTTOM_RIGHT
		EndSelect
	EndIf

	If ($iDistance = Null) Then
		; Retrieve the current Distance setting
		If ($oAnnotationShape.ShadowXDistance() <> 0) Then
			$iDistance = $oAnnotationShape.ShadowXDistance()

		ElseIf ($oAnnotationShape.ShadowYDistance() <> 0) Then
			$iDistance = $oAnnotationShape.ShadowYDistance()

		Else
			$iDistance = 0
		EndIf

		If $bModifyLocation And ($iDistance = 0) Then $iDistance = 100 ; Set a non 0 value so location can be set.

		; If negative, make it positive for easier processing.
		$iDistance = ($iDistance < 0) ? ($iDistance * -1) : ($iDistance)
	EndIf

	If $bReturn Then Return SetError($__LO_STATUS_SUCCESS, $iLocation, $iDistance)

	Switch $iLocation
		Case $LOC_COMMENT_SHADOW_TOP_LEFT
			$oAnnotationShape.ShadowXDistance = ($iDistance * -1)
			$oAnnotationShape.ShadowYDistance = ($iDistance * -1)

			Return (($oAnnotationShape.ShadowXDistance() = ($iDistance * -1)) And ($oAnnotationShape.ShadowYDistance() = ($iDistance * -1))) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))

		Case $LOC_COMMENT_SHADOW_TOP_CENTER
			$oAnnotationShape.ShadowXDistance = 0
			$oAnnotationShape.ShadowYDistance = ($iDistance * -1)

			Return (($oAnnotationShape.ShadowXDistance() = 0) And ($oAnnotationShape.ShadowYDistance() = ($iDistance * -1))) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))

		Case $LOC_COMMENT_SHADOW_TOP_RIGHT
			$oAnnotationShape.ShadowXDistance = $iDistance
			$oAnnotationShape.ShadowYDistance = ($iDistance * -1)

			Return (($oAnnotationShape.ShadowXDistance() = $iDistance) And ($oAnnotationShape.ShadowYDistance() = ($iDistance * -1))) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))

		Case $LOC_COMMENT_SHADOW_MIDDLE_LEFT
			$oAnnotationShape.ShadowXDistance = ($iDistance * -1)
			$oAnnotationShape.ShadowYDistance = 0

			Return (($oAnnotationShape.ShadowXDistance() = ($iDistance * -1)) And ($oAnnotationShape.ShadowYDistance() = 0)) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))

		Case $LOC_COMMENT_SHADOW_MIDDLE_CENTER
			$oAnnotationShape.ShadowXDistance = ($bModifyLocation) ? (0) : ($iDistance)
			$oAnnotationShape.ShadowYDistance = ($bModifyLocation) ? (0) : ($iDistance)

			Return (($oAnnotationShape.ShadowXDistance() = (($bModifyLocation) ? (0) : ($iDistance))) And ($oAnnotationShape.ShadowYDistance() = (($bModifyLocation) ? (0) : ($iDistance)))) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))

		Case $LOC_COMMENT_SHADOW_MIDDLE_RIGHT
			$oAnnotationShape.ShadowXDistance = $iDistance
			$oAnnotationShape.ShadowYDistance = 0

			Return (($oAnnotationShape.ShadowXDistance() = $iDistance) And ($oAnnotationShape.ShadowYDistance() = 0)) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))

		Case $LOC_COMMENT_SHADOW_BOTTOM_LEFT
			$oAnnotationShape.ShadowXDistance = ($iDistance * -1)
			$oAnnotationShape.ShadowYDistance = $iDistance

			Return (($oAnnotationShape.ShadowXDistance() = ($iDistance * -1)) And ($oAnnotationShape.ShadowYDistance() = $iDistance)) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))

		Case $LOC_COMMENT_SHADOW_BOTTOM_CENTER
			$oAnnotationShape.ShadowXDistance = 0
			$oAnnotationShape.ShadowYDistance = $iDistance

			Return (($oAnnotationShape.ShadowXDistance() = 0) And ($oAnnotationShape.ShadowYDistance() = $iDistance)) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))

		Case $LOC_COMMENT_SHADOW_BOTTOM_RIGHT
			$oAnnotationShape.ShadowXDistance = $iDistance
			$oAnnotationShape.ShadowYDistance = $iDistance

			Return (($oAnnotationShape.ShadowXDistance() = $iDistance) And ($oAnnotationShape.ShadowYDistance() = $iDistance)) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0))
	EndSwitch
EndFunc   ;==>__LOCalc_CommentAreaShadowModify

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CommentArrowStyleName
; Description ...: Convert a Arrow head Constant to the corresponding name or reverse.
; Syntax ........: __LOCalc_CommentArrowStyleName([$iArrowStyle = Null[, $sArrowStyle = Null]])
; Parameters ....: $iArrowStyle         - [optional] an integer value (0-32). Default is Null. The Arrow Style Constant to convert to its corresponding name. See $LOC_COMMENT_LINE_ARROW_TYPE_* as defined in LibreOfficeCalc_Constants.au3
;                  $sArrowStyle         - [optional] a string value. Default is Null. The Arrow Style Name to convert to the corresponding constant if found.
; Return values .: Success: String or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iArrowStyle not set to Null, not an Integer, less than 0, or greater than Arrow type constants. See $LOC_COMMENT_LINE_ARROW_TYPE_* as defined in LibreOfficeCalc_Constants.au3
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
Func __LOCalc_CommentArrowStyleName($iArrowStyle = Null, $sArrowStyle = Null)
	Local $asArrowStyles[33]

	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_NONE] = ""
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_ARROW_SHORT] = "Arrow short"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_CONCAVE_SHORT] = "Concave short"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_ARROW] = "Arrow"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_TRIANGLE] = "Triangle"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_CONCAVE] = "Concave"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_ARROW_LARGE] = "Arrow large"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_CIRCLE] = "Circle"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_SQUARE] = "Square"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_SQUARE_45] = "Square 45"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_DIAMOND] = "Diamond"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_HALF_CIRCLE] = "Half Circle"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_DIMENSIONAL_LINES] = "Dimension Lines"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_DIMENSIONAL_LINE_ARROW] = "Dimension Line Arrow"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_DIMENSION_LINE] = "Dimension Line"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_LINE_SHORT] = "Line short"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_LINE] = "Line"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_TRIANGLE_UNFILLED] = "Triangle unfilled"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_DIAMOND_UNFILLED] = "Diamond unfilled"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_CIRCLE_UNFILLED] = "Circle unfilled"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_SQUARE_45_UNFILLED] = "Square 45 unfilled"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_SQUARE_UNFILLED] = "Square unfilled"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_HALF_CIRCLE_UNFILLED] = "Half Circle unfilled"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_HALF_ARROW_LEFT] = "Half Arrow left"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_HALF_ARROW_RIGHT] = "Half Arrow right"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_REVERSED_ARROW] = "Reversed Arrow"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_DOUBLE_ARROW] = "Double Arrow"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_CF_ONE] = "CF One"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_CF_ONLY_ONE] = "CF Only One"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_CF_MANY] = "CF Many"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_CF_MANY_ONE] = "CF Many One"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_CF_ZERO_ONE] = "CF Zero One"
	$asArrowStyles[$LOC_COMMENT_LINE_ARROW_TYPE_CF_ZERO_MANY] = "CF Zero Many"

	If ($iArrowStyle <> Null) Then
		If Not __LO_IntIsBetween($iArrowStyle, 0, UBound($asArrowStyles) - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 0, $asArrowStyles[$iArrowStyle]) ; Return the requested Arrow Style name.

	ElseIf ($sArrowStyle <> Null) Then
		If Not IsString($sArrowStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		For $i = 0 To UBound($asArrowStyles) - 1
			If ($asArrowStyles[$i] = $sArrowStyle) Then Return SetError($__LO_STATUS_SUCCESS, 1, $i) ; Return the array element where the matching Arrow Style was found.

			Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
		Next

		Return SetError($__LO_STATUS_SUCCESS, 2, $sArrowStyle) ; If no matches, just return the name, as it could be a custom value.

	Else

		Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; No values called.
	EndIf
EndFunc   ;==>__LOCalc_CommentArrowStyleName

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CommentGetObjByCell
; Description ...: Internal function for getting a Comment Object by Cell.
; Syntax ........: __LOCalc_CommentGetObjByCell(ByRef $oCell[, $bReturnIndex = False])
; Parameters ....: $oCell               - [in/out] an object. A Cell object returned by a previous _LOCalc_RangeGetCellByName, or _LOCalc_RangeGetCellByPosition function.
;                  $bReturnIndex        - [optional] a boolean value. Default is False. If True, the Comment's index number is returned instead of its Object.
; Return values .: Success: Integer or Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCell not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCell not a Cell Object.
;                  @Error 1 @Extended 3 Return 0 = $bReturnIndex not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Annotations Object.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Cell Address.
;                  @Error 3 @Extended 3 Return 0 = Failed to find comment for specified cell.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success. $bReturnIndex set to True, returning Comment's Index number.
;                  @Error 0 @Extended ? Return Object = Success. Returning Comment's Object. @Extended set to Comment's Index number.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_CommentGetObjByCell(ByRef $oCell, $bReturnIndex = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tAddress
	Local $oAnnotations, $oAnnotation

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oCell.SupportsService("com.sun.star.sheet.SheetCell") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bReturnIndex) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oAnnotations = $oCell.Spreadsheet.Annotations()
	If Not IsObj($oAnnotations) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$tAddress = $oCell.CellAddress()
	If Not IsObj($tAddress) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	For $i = 0 To $oAnnotations.Count() - 1
		$oAnnotation = $oAnnotations.getByIndex($i)

		If __LOCalc_CellAddressIsSame($tAddress, $oAnnotation.Position()) Then
			If $bReturnIndex Then Return SetError($__LO_STATUS_SUCCESS, 1, $i)

			Return SetError($__LO_STATUS_SUCCESS, $i, $oAnnotation)
		EndIf

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
EndFunc   ;==>__LOCalc_CommentGetObjByCell

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_CommentLineStyleName
; Description ...: Convert a Line Style Constant to the corresponding name or reverse.
; Syntax ........: __LOCalc_CommentLineStyleName([$iLineStyle = Null[, $sLineStyle = Null]])
; Parameters ....: $iLineStyle          - [optional] an integer value. Default is Null. The Line Style Constant to convert to its corresponding name. See $LOC_COMMENT_LINE_STYLE_* as defined in LibreOfficeCalc_Constants.au3
;                  $sLineStyle          - [optional] a string value. Default is Null. The Line Style Name to convert to the corresponding constant if found.
; Return values .: Success: String or Integer
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iLineStyle not set to Null, not an Integer, less than 0, or greater than Line Style constants. See $LOC_COMMENT_LINE_STYLE_* as defined in LibreOfficeCalc_Constants.au3
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
Func __LOCalc_CommentLineStyleName($iLineStyle = Null, $sLineStyle = Null)
	Local $asLineStyles[32]

	; $LOC_COMMENT_LINE_STYLE_NONE, $LOC_COMMENT_LINE_STYLE_CONTINUOUS, don't have a name, so to keep things symmetrical I created my own, but those two won't be used.
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_NONE] = "NONE"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_CONTINUOUS] = "CONTINUOUS"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_DOT] = "Dot"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_DOT_ROUNDED] = "Dot (Rounded)"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_LONG_DOT] = "Long Dot"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_LONG_DOT_ROUNDED] = "Long Dot (Rounded)"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_DASH] = "Dash"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_DASH_ROUNDED] = "Dash (Rounded)"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_LONG_DASH] = "Long Dash"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_LONG_DASH_ROUNDED] = "Long Dash (Rounded)"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_DOUBLE_DASH] = "Double Dash"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_DOUBLE_DASH_ROUNDED] = "Double Dash (Rounded)"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_DASH_DOT] = "Dash Dot"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_DASH_DOT_ROUNDED] = "Dash Dot (Rounded)"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_LONG_DASH_DOT] = "Long Dash Dot"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_LONG_DASH_DOT_ROUNDED] = "Long Dash Dot (Rounded)"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_DOUBLE_DASH_DOT] = "Double Dash Dot"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_DOUBLE_DASH_DOT_ROUNDED] = "Double Dash Dot (Rounded)"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_DASH_DOT_DOT] = "Dash Dot Dot"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_DASH_DOT_DOT_ROUNDED] = "Dash Dot Dot (Rounded)"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_DOUBLE_DASH_DOT_DOT] = "Double Dash Dot Dot"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_DOUBLE_DASH_DOT_DOT_ROUNDED] = "Double Dash Dot Dot (Rounded)"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_ULTRAFINE_DOTTED] = "Ultrafine Dotted (var)"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_FINE_DOTTED] = "Fine Dotted"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_ULTRAFINE_DASHED] = "Ultrafine Dashed"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_FINE_DASHED] = "Fine Dashed"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_DASHED] = "Dashed (var)"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_LINE_STYLE_9] = "Line Style 9"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_3_DASHES_3_DOTS] = "3 Dashes 3 Dots (var)"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_ULTRAFINE_2_DOTS_3_DASHES] = "Ultrafine 2 Dots 3 Dashes"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_2_DOTS_1_DASH] = "2 Dots 1 Dash"
	$asLineStyles[$LOC_COMMENT_LINE_STYLE_LINE_WITH_FINE_DOTS] = "Line with Fine Dots"

	If ($iLineStyle <> Null) Then
		If Not __LO_IntIsBetween($iLineStyle, 0, UBound($asLineStyles) - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 0, $asLineStyles[$iLineStyle]) ; Return the requested Line Style name.

	ElseIf ($sLineStyle <> Null) Then
		If Not IsString($sLineStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		For $i = 0 To UBound($asLineStyles) - 1
			If ($asLineStyles[$i] = $sLineStyle) Then Return SetError($__LO_STATUS_SUCCESS, 1, $i) ; Return the array element where the matching Line Style was found.

			Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
		Next

		Return SetError($__LO_STATUS_SUCCESS, 2, $sLineStyle) ; If no matches, just return the name, as it could be a custom value.

	Else

		Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0) ; No values called.
	EndIf
EndFunc   ;==>__LOCalc_CommentLineStyleName

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_FieldGetObj
; Description ...: Retrieve the Field's Object after insertion.
; Syntax ........: __LOCalc_FieldGetObj(ByRef $oTextCursor[, $iType = $LOC_FIELD_TYPE_ALL])
; Parameters ....: $oTextCursor         - [in/out] an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
;                  $iType               - [optional] an integer value. Default is $LOC_FIELD_TYPE_ALL. The Type of field to search for.
; Return values .: Success: Map
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTextCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iType not an Integer, less than 1, or greater than 255. (The total of all Constants added together.) See Constants, $LOC_FIELD_TYPE_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create enumeration of paragraphs in Cell.
;                  @Error 2 @Extended 2 Return 0 = Failed to create enumeration of Text Portions in Paragraph.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify requested Field Types.
;                  @Error 3 @Extended 2 Return 0 = Failed to retrieve Text Fields Object.
;                  @Error 3 @Extended 3 Return 0 = Failed to retrieve total Fields count.
;                  @Error 3 @Extended 4 Return 0 = Failed to retrieve Text Field Object.
;                  @Error 3 @Extended 5 Return 0 = Number of identified fields is greater than number of expected fields.
;                  @Error 3 @Extended 6 Return 0 = Failed to retrieve Alternate Text Field Object.
;                  @Error 3 @Extended 7 Return 0 = Failed to identify newly created Field.
;                  --Success--
;                  @Error 0 @Extended 0 Return Map = Success. Returning newly inserted Field's Object inside of a map.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: After inserting a Field, the Object is not usable for modifying the field later on, so I retrieve it again after insertion.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_FieldGetObj(ByRef $oTextCursor, $iType = $LOC_FIELD_TYPE_ALL)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avFieldTypes[0][0]
	Local $oParEnum, $oPar, $oTextEnum, $oTextPortion, $oTextField, $oInternalCursor = $oTextCursor, $oFields, $oField
	Local $iTotalFields = 0, $iTotalFound = 0
	Local $mFieldObj[]

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iType, $LOC_FIELD_TYPE_ALL, 255) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	; When a Text Cursor has been used to insert Strings previous to inserting or looking for a Field, the fields sometimes are not able to be identified.
	; The workaround I figured out was to create the Text Cursor again before enumerating the fields. I only create the text cursor again if the Text Cursor is in a Cell, not a header.
	If ($oTextCursor.Text.SupportsService("com.sun.star.sheet.SheetCell")) Then
		$oInternalCursor = $oTextCursor.Text.Spreadsheet.getCellByPosition($oTextCursor.Text.RangeAddress.StartColumn(), $oTextCursor.Text.RangeAddress.StartRow()).Text.createTextCursorByRange($oTextCursor)
	EndIf

	$avFieldTypes = __LOCalc_FieldTypeServices($iType)
	If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oFields = $oInternalCursor.Text.TextFields()
	If Not IsObj($oFields) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$iTotalFields = $oFields.Count()
	If Not IsInt($iTotalFields) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$oParEnum = $oInternalCursor.getText().createEnumeration()
	If Not IsObj($oParEnum) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	While $oParEnum.hasMoreElements()
		$oPar = $oParEnum.nextElement()

		$oTextEnum = $oPar.createEnumeration()
		If Not IsObj($oTextEnum) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		While $oTextEnum.hasMoreElements()
			$oTextPortion = $oTextEnum.nextElement()

			If ($oTextPortion.TextPortionType = "TextField") Then
				$oTextField = $oTextPortion.TextField()
				If Not IsObj($oTextField) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
				If ($iTotalFound >= $iTotalFields) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

				For $i = 0 To UBound($avFieldTypes) - 1
					If $oTextField.supportsService($avFieldTypes[$i][1]) And ($oInternalCursor.compareRegionEnds($oInternalCursor, $oTextField.Anchor.End()) = 0) Then
						$oField = $oFields.getByIndex($iTotalFound)
						If Not IsObj($oField) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

						$mFieldObj.EnumFieldObj = $oTextField
						$mFieldObj.FieldObj = $oField

						Return SetError($__LO_STATUS_SUCCESS, 0, $mFieldObj)
					EndIf
					Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
				Next

				$iTotalFound += 1
			EndIf
		WEnd
	WEnd

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)
EndFunc   ;==>__LOCalc_FieldGetObj

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_FieldTypeServices
; Description ...: Retrieve an Array of Supported Service Names and Integer Constants to search for Fields.
; Syntax ........: __LOCalc_FieldTypeServices($iFieldType)
; Parameters ....: $iFieldType          - an integer value. The Constant Field type.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iFieldType not an Integer.
;                  --Success--
;                  @Error 0 @Extended 0 Return Array = Success. $iFieldType set to All, returning full regular Field Service list String Array.
;                  @Error 0 @Extended 1 Return Array = Success. $iFieldType BitOr'd together, determining which flags are called from the Array. Returning Field Service String list Array.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_FieldTypeServices($iFieldType)
	Local $avFieldTypes[7][2] = [[$LOC_FIELD_TYPE_DATE_TIME, "com.sun.star.text.TextField.DateTime"], [$LOC_FIELD_TYPE_DOC_TITLE, "com.sun.star.text.TextField.docinfo.Title"], _
			[$LOC_FIELD_TYPE_FILE_NAME, "com.sun.star.text.TextField.FileName"], [$LOC_FIELD_TYPE_PAGE_NUM, "com.sun.star.text.TextField.PageNumber"], _
			[$LOC_FIELD_TYPE_PAGE_COUNT, "com.sun.star.text.TextField.PageCount"], [$LOC_FIELD_TYPE_SHEET_NAME, "com.sun.star.text.TextField.SheetName"], _
			[$LOC_FIELD_TYPE_URL, "com.sun.star.text.TextField.URL"]]

	Local $avFieldResults[UBound($avFieldTypes)][2]
	Local $iCount = 0

	If Not IsInt($iFieldType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If (BitAND($iFieldType, $LOC_FIELD_TYPE_ALL)) Then Return SetError($__LO_STATUS_SUCCESS, 0, $avFieldTypes)

	For $i = 0 To UBound($avFieldTypes) - 1
		If BitAND($avFieldTypes[$i][0], $iFieldType) Then
			$avFieldResults[$iCount][0] = $avFieldTypes[$i][0]
			$avFieldResults[$iCount][1] = $avFieldTypes[$i][1]
			$iCount += 1
		EndIf
		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
	Next

	ReDim $avFieldResults[$iCount][2]

	Return SetError($__LO_STATUS_SUCCESS, 1, $avFieldResults)
EndFunc   ;==>__LOCalc_FieldTypeServices

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_FilterNameGet
; Description ...: Retrieves the correct L.O. Filtername for use in SaveAs and Export.
; Syntax ........: __LOCalc_FilterNameGet(ByRef $sDocSavePath[, $bExportFilters = False])
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
;                  @Error 0 @Extended 3 Return String = FilterName not found for given file extension, defaulting to .ods file format and updating save path accordingly.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Searches a predefined list of extensions stored in an array. Not all FilterNames are listed.
;                  For finding your own FilterNames, see convertfilters.html found in L.O. Install Folder: LibreOffice\help\en-US\text\shared\guide
;                  Or See: "OOME_3_0", "OpenOffice.org Macros Explained OOME Third Edition" by Andrew D. Pitonyak, which has a handy Macro for listing all FilterNames, found on page 284 of the above book in the ODT format.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_FilterNameGet(ByRef $sDocSavePath, $bExportFilters = False)
	Local $iLength, $iSlashLocation, $iDotLocation
	Local Const $STR_NOCASESENSE = 0, $STR_STRIPALL = 8
	Local $sFileExtension, $sFilterName
	Local $msSaveAsFilters[], $msExportFilters[]

	If Not IsString($sDocSavePath) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bExportFilters) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$iLength = StringLen($sDocSavePath)

	$msSaveAsFilters[".csv"] = "Text - txt - csv (StarCalc)"
	$msSaveAsFilters[".dbf"] = "dBase"
	$msSaveAsFilters[".dif"] = "DIF"
	$msSaveAsFilters[".et"] = "MS Excel 97"
	$msSaveAsFilters[".ett"] = "MS Excel 97 Vorlage/Template"
	$msSaveAsFilters[".fods"] = "OpenDocument Spreadsheet Flat XML"
	$msSaveAsFilters[".htm"] = "HTML (StarCalc)"
	$msSaveAsFilters[".html"] = "HTML (StarCalc)"
	$msSaveAsFilters[".ods"] = "calc8"
	$msSaveAsFilters[".ots"] = "calc8_template"
	$msSaveAsFilters[".slk"] = "SYLK"
	$msSaveAsFilters[".sylk"] = "SYLK"
	$msSaveAsFilters[".tab"] = "Text - txt - csv (StarCalc)"
	$msSaveAsFilters[".tsv"] = "Text - txt - csv (StarCalc)"
	$msSaveAsFilters[".txt"] = "Text - txt - csv (StarCalc)"
	$msSaveAsFilters[".uof"] = "UOF spreadsheet"
	$msSaveAsFilters[".uos"] = "UOF spreadsheet"
	$msSaveAsFilters[".xhtml"] = "HTML (StarCalc)"
	$msSaveAsFilters[".xlc"] = "MS Excel 97"
	$msSaveAsFilters[".xlk"] = "MS Excel 97"
	$msSaveAsFilters[".xlm"] = "MS Excel 97"
	$msSaveAsFilters[".xls"] = "MS Excel 97"
	$msSaveAsFilters[".xlsm"] = "Calc MS Excel 2007 VBA XML"
	$msSaveAsFilters[".xlsx"] = "Calc MS Excel 2007 XML"
	$msSaveAsFilters[".xlt"] = "MS Excel 97 Vorlage/Template"
	$msSaveAsFilters[".xltm"] = "Calc MS Excel 2007 XML Template"
	$msSaveAsFilters[".xltx"] = "Calc MS Excel 2007 XML Template"
	$msSaveAsFilters[".xlw"] = "MS Excel 97"
	$msSaveAsFilters[".xml"] = "OpenDocument Spreadsheet Flat XML"

	If $bExportFilters Then
		$msExportFilters[".jfif"] = "calc_jpg_Export"
		$msExportFilters[".jif"] = "calc_jpg_Export"
		$msExportFilters[".jpe"] = "calc_jpg_Export"
		$msExportFilters[".jpeg"] = "calc_jpg_Export"
		$msExportFilters[".jpg"] = "calc_jpg_Export"
		$msExportFilters[".pdf"] = "calc_pdf_Export"
		$msExportFilters[".png"] = "calc_png_Export"
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

		Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	EndIf

	If $sFileExtension = $sDocSavePath Then ;  If no file extension identified, append .ods extension and return.
		$sDocSavePath = $sDocSavePath & ".ods"

		Return SetError($__LO_STATUS_SUCCESS, 3, "calc8")

	Else
		$sFileExtension = StringLower(StringStripWS($sFileExtension, $STR_STRIPALL))
	EndIf

	$sFilterName = $msSaveAsFilters[$sFileExtension]

	If IsString($sFilterName) Then Return SetError($__LO_STATUS_SUCCESS, 1, $sFilterName)

	If $bExportFilters Then $sFilterName = $msExportFilters[$sFileExtension]

	If IsString($sFilterName) Then Return SetError($__LO_STATUS_SUCCESS, 2, $sFilterName)

	$sDocSavePath = StringReplace($sDocSavePath, $sFileExtension, ".ods") ; If No results, replace with ODS extension.

	Return SetError($__LO_STATUS_SUCCESS, 3, "calc8")
EndFunc   ;==>__LOCalc_FilterNameGet

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_Internal_CursorGetType
; Description ...: Get what type of cursor the object is.
; Syntax ........: __LOCalc_Internal_CursorGetType(ByRef $oCursor)
; Parameters ....: $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions.
; Return values .: Success: Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Unknown Cursor type.
;                  --Success--
;                  @Error 0 @Extended 0 Return Integer = Success, Return value will be one of the constants, $LOC_CURTYPE_* as defined in LibreOfficeCalc_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Returns what type of cursor the input Object is, such as a Text Cursor or a Sheet Cursor. Can also be a Paragraph or Text Portion.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_Internal_CursorGetType(ByRef $oCursor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Switch $oCursor.getImplementationName()
		Case "SvxUnoTextCursor"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOC_CURTYPE_TEXT_CURSOR)

		Case "ScCellCursorObj"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOC_CURTYPE_SHEET_CURSOR)

		Case "SvxUnoTextContent"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOC_CURTYPE_PARAGRAPH)

		Case "SvxUnoTextRange"

			Return SetError($__LO_STATUS_SUCCESS, 0, $LOC_CURTYPE_TEXT_PORTION)

		Case Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; unknown Cursor type.
	EndSwitch
EndFunc   ;==>__LOCalc_Internal_CursorGetType

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_InternalComErrorHandler
; Description ...: ComError Handler
; Syntax ........: __LOCalc_InternalComErrorHandler(ByRef $oComError)
; Parameters ....: $oComError           - [in/out] an object. The Com Error Object passed by Autoit.Error.
; Return values .: None
; Author ........: mLipok
; Modified ......: donnyh13 - Added parameters option. Also added MsgBox & ConsoleWrite options.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_InternalComErrorHandler(ByRef $oComError)
	; If not defined ComError_UserFunction then this function does nothing, in which case you can only check @error / @extended after suspect functions.
	Local $avUserFunction = _LOCalc_ComError_UserFunction(Default)
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
EndFunc   ;==>__LOCalc_InternalComErrorHandler

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_NamedRangeGetScopeObj
; Description ...: Retrieve the Scope Object that contains a particular Named Range.
; Syntax ........: __LOCalc_NamedRangeGetScopeObj(ByRef $oDoc, $sName, $iTokenIndex, $sContent)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $sName               - a string value. The name of the Named Range to retrieve the scope object for.
;                  $iTokenIndex         - an integer value. The Token Index of the Named Range to retrieve the scope object for.
;                  $sContent            - a string value. The Content of the Named Range to retrieve the scope object for.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sName not a String.
;                  @Error 1 @Extended 3 Return 0 = $iTokenIndex not an Integer.
;                  @Error 1 @Extended 4 Return 0 = $sContent not a String.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify Scope Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Returning Scope object (Doc or Sheet) that contains the Named Range.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_NamedRangeGetScopeObj(ByRef $oDoc, $sName, $iTokenIndex, $sContent)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iTokenIndex) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sContent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If ($oDoc.NamedRanges.Count() >= $iTokenIndex) Then
		$oObj = $oDoc.NamedRanges.getByIndex($iTokenIndex - 1)
		If ($oObj.Name() == $sName) And ($oObj.Content = $sContent) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oDoc)
	EndIf

	For $i = 0 To $oDoc.Sheets.Count() - 1
		If ($oDoc.Sheets.getByIndex($i).NamedRanges.Count() >= $iTokenIndex) Then
			$oObj = $oDoc.Sheets.getByIndex($i).NamedRanges.getByIndex($iTokenIndex - 1)
			If ($oObj.Name() == $sName) And ($oObj.Content = $sContent) Then Return SetError($__LO_STATUS_SUCCESS, 2, $oDoc.Sheets.getByIndex($i))
		EndIf

		Sleep((IsInt($i / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
EndFunc   ;==>__LOCalc_NamedRangeGetScopeObj

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_PageStyleBorder
; Description ...: Internal function to Set and Retrieve the Page Style Border Line Width, Style, and Color. Libre Office Version 3.6 and Up.
; Syntax ........: __LOCalc_PageStyleBorder(ByRef $oPageStyle, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $bWid                - a boolean value. If True, Border Width is being modified. Only one can be True at once.
;                  $bSty                - a boolean value. If True, Border Style is being modified. Only one can be True at once.
;                  $bCol                - a boolean value. If True, Border Color is being modified. Only one can be True at once.
;                  $iTop                - an integer value. Modifies the top border line settings. See Width, Style or Color functions for values.
;                  $iBottom             - an integer value. Modifies the bottom border line settings. See Width, Style or Color functions for values.
;                  $iLeft               - an integer value. Modifies the left border line settings. See Width, Style or Color functions for values.
;                  $iRight              - an integer value. Modifies the right border line settings. See Width, Style or Color functions for values.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 7 Return 0 = Variable passed to internal function not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style/Color when Top Border width not set.
;                  @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Style/Color when Bottom Border width not set.
;                  @Error 4 @Extended 3 Return 0 = Cannot set Left Border Style/Color when Left Border width not set.
;                  @Error 4 @Extended 4 Return 0 = Cannot set Right Border Style/Color when Right Border width not set.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_PageStyleBorder(ByRef $oPageStyle, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avBorder[4]
	Local $tBL2

	If Not __LO_VersionCheck(3.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If (($bWid + $bSty + $bCol) <> 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; If more than one Boolean is true = error

	If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then
		If $bWid Then
			__LO_ArrayFill($avBorder, $oPageStyle.TopBorder.LineWidth(), $oPageStyle.BottomBorder.LineWidth(), $oPageStyle.LeftBorder.LineWidth(), $oPageStyle.RightBorder.LineWidth())

		ElseIf $bSty Then
			__LO_ArrayFill($avBorder, $oPageStyle.TopBorder.LineStyle(), $oPageStyle.BottomBorder.LineStyle(), $oPageStyle.LeftBorder.LineStyle(), $oPageStyle.RightBorder.LineStyle())

		ElseIf $bCol Then
			__LO_ArrayFill($avBorder, $oPageStyle.TopBorder.Color(), $oPageStyle.BottomBorder.Color(), $oPageStyle.LeftBorder.Color(), $oPageStyle.RightBorder.Color())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avBorder)
	EndIf

	$tBL2 = __LO_CreateStruct("com.sun.star.table.BorderLine2")
	If Not IsObj($tBL2) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If $iTop <> Null Then
		If Not $bWid And ($oPageStyle.TopBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0) ; If Width not set, cant set color or style.

		; Top Line
		$tBL2.LineWidth = ($bWid) ? ($iTop) : ($oPageStyle.TopBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iTop) : ($oPageStyle.TopBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iTop) : ($oPageStyle.TopBorder.Color()) ; copy Color over to new size structure
		$oPageStyle.TopBorder = $tBL2
	EndIf

	If $iBottom <> Null Then
		If Not $bWid And ($oPageStyle.BottomBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 2, 0) ; If Width not set, cant set color or style.

		; Bottom Line
		$tBL2.LineWidth = ($bWid) ? ($iBottom) : ($oPageStyle.BottomBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iBottom) : ($oPageStyle.BottomBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iBottom) : ($oPageStyle.BottomBorder.Color()) ; copy Color over to new size structure
		$oPageStyle.BottomBorder = $tBL2
	EndIf

	If $iLeft <> Null Then
		If Not $bWid And ($oPageStyle.LeftBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 3, 0) ; If Width not set, cant set color or style.

		; Left Line
		$tBL2.LineWidth = ($bWid) ? ($iLeft) : ($oPageStyle.LeftBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iLeft) : ($oPageStyle.LeftBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iLeft) : ($oPageStyle.LeftBorder.Color()) ; copy Color over to new size structure
		$oPageStyle.LeftBorder = $tBL2
	EndIf

	If $iRight <> Null Then
		If Not $bWid And ($oPageStyle.RightBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 4, 0) ; If Width not set, cant set color or style.

		; Right Line
		$tBL2.LineWidth = ($bWid) ? ($iRight) : ($oPageStyle.RightBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iRight) : ($oPageStyle.RightBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iRight) : ($oPageStyle.RightBorder.Color()) ; copy Color over to new size structure
		$oPageStyle.RightBorder = $tBL2
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOCalc_PageStyleBorder

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_PageStyleFooterBorder
; Description ...: Internal function to Set and Retrieve the Page Style Footer Border Line Width, Style, and Color. Libre Office Version 3.6 and Up.
; Syntax ........: __LOCalc_PageStyleFooterBorder(ByRef $oPageStyle, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $bWid                - a boolean value. If True, Border Width is being modified. Only one can be True at once.
;                  $bSty                - a boolean value. If True, Border Style is being modified. Only one can be True at once.
;                  $bCol                - a boolean value. If True, Border Color is being modified. Only one can be True at once.
;                  $iTop                - an integer value. Modifies the top border line settings. See Width, Style or Color functions for values.
;                  $iBottom             - an integer value. Modifies the bottom border line settings. See Width, Style or Color functions for values.
;                  $iLeft               - an integer value. Modifies the left border line settings. See Width, Style or Color functions for values.
;                  $iRight              - an integer value. Modifies the right border line settings. See Width, Style or Color functions for values.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 7 Return 0 = Variable passed to internal function not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style/Color when Top Border width not set.
;                  @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Style/Color when Bottom Border width not set.
;                  @Error 4 @Extended 3 Return 0 = Cannot set Left Border Style/Color when Left Border width not set.
;                  @Error 4 @Extended 4 Return 0 = Cannot set Right Border Style/Color when Right Border width not set.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_PageStyleFooterBorder(ByRef $oPageStyle, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avBorder[4]
	Local $tBL2

	If Not __LO_VersionCheck(3.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If (($bWid + $bSty + $bCol) <> 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; If more than one Boolean is true = error

	If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then
		If $bWid Then
			__LO_ArrayFill($avBorder, $oPageStyle.FooterTopBorder.LineWidth(), $oPageStyle.FooterBottomBorder.LineWidth(), $oPageStyle.FooterLeftBorder.LineWidth(), $oPageStyle.FooterRightBorder.LineWidth())

		ElseIf $bSty Then
			__LO_ArrayFill($avBorder, $oPageStyle.FooterTopBorder.LineStyle(), $oPageStyle.FooterBottomBorder.LineStyle(), $oPageStyle.FooterLeftBorder.LineStyle(), $oPageStyle.FooterRightBorder.LineStyle())

		ElseIf $bCol Then
			__LO_ArrayFill($avBorder, $oPageStyle.FooterTopBorder.Color(), $oPageStyle.FooterBottomBorder.Color(), $oPageStyle.FooterLeftBorder.Color(), $oPageStyle.FooterRightBorder.Color())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avBorder)
	EndIf

	$tBL2 = __LO_CreateStruct("com.sun.star.table.BorderLine2")
	If Not IsObj($tBL2) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If $iTop <> Null Then
		If Not $bWid And ($oPageStyle.FooterTopBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0) ; If Width not set, cant set color or style.

		; Top Line
		$tBL2.LineWidth = ($bWid) ? ($iTop) : ($oPageStyle.FooterTopBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iTop) : ($oPageStyle.FooterTopBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iTop) : ($oPageStyle.FooterTopBorder.Color()) ; copy Color over to new size structure
		$oPageStyle.FooterTopBorder = $tBL2
	EndIf

	If $iBottom <> Null Then
		If Not $bWid And ($oPageStyle.FooterBottomBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 2, 0) ; If Width not set, cant set color or style.

		; Bottom Line
		$tBL2.LineWidth = ($bWid) ? ($iBottom) : ($oPageStyle.FooterBottomBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iBottom) : ($oPageStyle.FooterBottomBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iBottom) : ($oPageStyle.FooterBottomBorder.Color()) ; copy Color over to new size structure
		$oPageStyle.FooterBottomBorder = $tBL2
	EndIf

	If $iLeft <> Null Then
		If Not $bWid And ($oPageStyle.FooterLeftBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 3, 0) ; If Width not set, cant set color or style.

		; Left Line
		$tBL2.LineWidth = ($bWid) ? ($iLeft) : ($oPageStyle.FooterLeftBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iLeft) : ($oPageStyle.FooterLeftBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iLeft) : ($oPageStyle.FooterLeftBorder.Color()) ; copy Color over to new size structure
		$oPageStyle.FooterLeftBorder = $tBL2
	EndIf

	If $iRight <> Null Then
		If Not $bWid And ($oPageStyle.FooterRightBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 4, 0) ; If Width not set, cant set color or style.

		; Right Line
		$tBL2.LineWidth = ($bWid) ? ($iRight) : ($oPageStyle.FooterRightBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iRight) : ($oPageStyle.FooterRightBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iRight) : ($oPageStyle.FooterRightBorder.Color()) ; copy Color over to new size structure
		$oPageStyle.FooterRightBorder = $tBL2
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOCalc_PageStyleFooterBorder

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_PageStyleHeaderBorder
; Description ...: Internal function to Set and Retrieve the Page Style Header Border Line Width, Style, and Color. Libre Office Version 3.6 and Up.
; Syntax ........: __LOCalc_PageStyleHeaderBorder(ByRef $oPageStyle, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
; Parameters ....: $oPageStyle          - [in/out] an object. A Page Style object returned by a previous _LOCalc_PageStyleCreate, or _LOCalc_PageStyleGetObj function.
;                  $bWid                - a boolean value. If True, Border Width is being modified. Only one can be True at once.
;                  $bSty                - a boolean value. If True, Border Style is being modified. Only one can be True at once.
;                  $bCol                - a boolean value. If True, Border Color is being modified. Only one can be True at once.
;                  $iTop                - an integer value. Modifies the top border line settings. See Width, Style or Color functions for values.
;                  $iBottom             - an integer value. Modifies the bottom border line settings. See Width, Style or Color functions for values.
;                  $iLeft               - an integer value. Modifies the left border line settings. See Width, Style or Color functions for values.
;                  $iRight              - an integer value. Modifies the right border line settings. See Width, Style or Color functions for values.
; Return values .: Success: 1 or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 9 Return 0 = Variable passed to internal function not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;                  --Property Setting Errors--
;                  @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style/Color when Top Border width not set.
;                  @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Style/Color when Bottom Border width not set.
;                  @Error 4 @Extended 3 Return 0 = Cannot set Left Border Style/Color when Left Border width not set.
;                  @Error 4 @Extended 4 Return 0 = Cannot set Right Border Style/Color when Right Border width not set.
;                  --Version Related Errors--
;                  @Error 6 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_PageStyleHeaderBorder(ByRef $oPageStyle, $bWid, $bSty, $bCol, $iTop, $iBottom, $iLeft, $iRight)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avBorder[4]
	Local $tBL2

	If Not __LO_VersionCheck(3.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
	If Not IsObj($oPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
	If (($bWid + $bSty + $bCol) <> 1) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; If more than one Boolean is true = error

	If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then
		If $bWid Then
			__LO_ArrayFill($avBorder, $oPageStyle.HeaderTopBorder.LineWidth(), $oPageStyle.HeaderBottomBorder.LineWidth(), $oPageStyle.HeaderLeftBorder.LineWidth(), $oPageStyle.HeaderRightBorder.LineWidth())

		ElseIf $bSty Then
			__LO_ArrayFill($avBorder, $oPageStyle.HeaderTopBorder.LineStyle(), $oPageStyle.HeaderBottomBorder.LineStyle(), $oPageStyle.HeaderLeftBorder.LineStyle(), $oPageStyle.HeaderRightBorder.LineStyle())

		ElseIf $bCol Then
			__LO_ArrayFill($avBorder, $oPageStyle.HeaderTopBorder.Color(), $oPageStyle.HeaderBottomBorder.Color(), $oPageStyle.HeaderLeftBorder.Color(), $oPageStyle.HeaderRightBorder.Color())
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 1, $avBorder)
	EndIf

	$tBL2 = __LO_CreateStruct("com.sun.star.table.BorderLine2")
	If Not IsObj($tBL2) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If $iTop <> Null Then
		If Not $bWid And ($oPageStyle.HeaderTopBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0) ; If Width not set, cant set color or style.

		; Top Line
		$tBL2.LineWidth = ($bWid) ? ($iTop) : ($oPageStyle.HeaderTopBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iTop) : ($oPageStyle.HeaderTopBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iTop) : ($oPageStyle.HeaderTopBorder.Color()) ; copy Color over to new size structure
		$oPageStyle.HeaderTopBorder = $tBL2
	EndIf

	If $iBottom <> Null Then
		If Not $bWid And ($oPageStyle.HeaderBottomBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 2, 0) ; If Width not set, cant set color or style.

		; Bottom Line
		$tBL2.LineWidth = ($bWid) ? ($iBottom) : ($oPageStyle.HeaderBottomBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iBottom) : ($oPageStyle.HeaderBottomBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iBottom) : ($oPageStyle.HeaderBottomBorder.Color()) ; copy Color over to new size structure
		$oPageStyle.HeaderBottomBorder = $tBL2
	EndIf

	If $iLeft <> Null Then
		If Not $bWid And ($oPageStyle.HeaderLeftBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 3, 0) ; If Width not set, cant set color or style.

		; Left Line
		$tBL2.LineWidth = ($bWid) ? ($iLeft) : ($oPageStyle.HeaderLeftBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iLeft) : ($oPageStyle.HeaderLeftBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iLeft) : ($oPageStyle.HeaderLeftBorder.Color()) ; copy Color over to new size structure
		$oPageStyle.HeaderLeftBorder = $tBL2
	EndIf

	If $iRight <> Null Then
		If Not $bWid And ($oPageStyle.HeaderRightBorder.LineWidth() = 0) Then Return SetError($__LO_STATUS_PROP_SETTING_ERROR, 4, 0) ; If Width not set, cant set color or style.

		; Right Line
		$tBL2.LineWidth = ($bWid) ? ($iRight) : ($oPageStyle.HeaderRightBorder.LineWidth()) ; copy Line Width over to new size structure
		$tBL2.LineStyle = ($bSty) ? ($iRight) : ($oPageStyle.HeaderRightBorder.LineStyle()) ; copy Line style over to new size structure
		$tBL2.Color = ($bCol) ? ($iRight) : ($oPageStyle.HeaderRightBorder.Color()) ; copy Color over to new size structure
		$oPageStyle.HeaderRightBorder = $tBL2
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>__LOCalc_PageStyleHeaderBorder

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_RangeAddressIsSame
; Description ...: Compare two Range Addresses to see if they are the same.
; Syntax ........: __LOCalc_RangeAddressIsSame(ByRef $tRange1, ByRef $tRange2)
; Parameters ....: $tRange1             - a dll struct value. The first Range Address Structure to compare.
;                  $tRange2             - a dll struct value. The second Range Address Structure to compare.
; Return values .: Success: Boolean
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $tRange1 not an Object.
;                  @Error 1 @Extended 2 Return 0 = $tRange2 not an Object.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. If the Range Addresses are identical, True is returned, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_RangeAddressIsSame($tRange1, $tRange2)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($tRange1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($tRange2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If ($tRange1.Sheet() = $tRange2.Sheet()) And _
			($tRange1.StartColumn() = $tRange2.StartColumn()) And _
			($tRange1.StartRow() = $tRange2.StartRow()) And _
			($tRange1.EndColumn() = $tRange2.EndColumn()) And _
			($tRange1.EndRow() = $tRange2.EndRow()) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>__LOCalc_RangeAddressIsSame

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_SheetCursorMove
; Description ...: For Sheet Cursor related movements.
; Syntax ........: __LOCalc_SheetCursorMove(ByRef $oCursor, $iMove, $iColumns, $iRows, $iCount, $bSelect)
; Parameters ....: $oCursor             - [in/out] an object. A Sheet Cursor Object returned from any Sheet Cursor creation functions.
;                  $iMove               - an Integer value. The movement command constant. See remarks and Constants, $LOC_SHEETCUR* as defined in LibreOfficeCalc_Constants.au3.
;                  $iColumns            - an integer value. The Number of Columns either to contain in the Range, or to move, depending on the called move command.
;                  $iRows               - an integer value. The Number of Rows either to contain in the Range, or to move, depending on the called move command.
;                  $iCount              - an integer value. Number of movements to make.
;                  $bSelect             - [optional] a boolean value. Default is False. If True, select data during this cursor movement.
; Return values .: Success: 1.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iMove not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iMove less than 0 or greater than highest move Constant. See Constants, $LOC_SHEETCUR* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iColumns not an integer.
;                  @Error 1 @Extended 5 Return 0 = $iRows not an integer.
;                  @Error 1 @Extended 6 Return 0 = $iCount not an integer or is a negative.
;                  @Error 1 @Extended 7 Return 0 = $bSelect not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 2 Return 0 = Error processing cursor move.
;                  --Success--
;                  @Error 0 @Extended ? Return 1 = Success, Cursor object movement was processed successfully.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only some movements accept Column and Row Values, creating/ extending a selection of cells, etc. They will be specified below.
;                  #Cursor Movement Constants which accept Column and Row values:
;                  $LOC_SHEETCUR_COLLAPSE_TO_SIZE,
;                  $LOC_SHEETCUR_GOTO_OFFSET
;                  #Cursor Movements which accept Selecting Only:
;                  $LOC_SHEETCUR_GOTO_USED_AREA_START,
;                  $LOC_SHEETCUR_GOTO_USED_AREA_END
;                  #Cursor Movements which accept nothing and are done once per call:
;                  $LOC_SHEETCUR_COLLAPSE_TO_CURRENT_ARRAY,
;                  $LOC_SHEETCUR_COLLAPSE_TO_CURRENT_REGION,
;                  $LOC_SHEETCUR_COLLAPSE_TO_MERGED_AREA,
;                  $LOC_SHEETCUR_EXPAND_TO_ENTIRE_COLUMN,
;                  $LOC_SHEETCUR_EXPAND_TO_ENTIRE_ROW,
;                  $LOC_SHEETCUR_GOTO_START,
;                  $LOC_SHEETCUR_GOTO_END
;                  #Cursor Movements which accept only number of moves ($iCount):
;                  $LOC_SHEETCUR_GOTO_NEXT,
;                  $LOC_SHEETCUR_GOTO_PREV
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_SheetCursorMove(ByRef $oCursor, $iMove, $iColumns, $iRows, $iCount, $bSelect)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCounted = 0
	Local $asMoves[13]

	$asMoves[$LOC_SHEETCUR_COLLAPSE_TO_CURRENT_ARRAY] = "collapseToCurrentArray"
	$asMoves[$LOC_SHEETCUR_COLLAPSE_TO_CURRENT_REGION] = "collapseToCurrentRegion"
	$asMoves[$LOC_SHEETCUR_COLLAPSE_TO_MERGED_AREA] = "collapseToMergedArea"
	$asMoves[$LOC_SHEETCUR_COLLAPSE_TO_SIZE] = "collapseToSize"
	$asMoves[$LOC_SHEETCUR_EXPAND_TO_ENTIRE_COLUMN] = "expandToEntireColumns"
	$asMoves[$LOC_SHEETCUR_EXPAND_TO_ENTIRE_ROW] = "expandToEntireRows"
	$asMoves[$LOC_SHEETCUR_GOTO_OFFSET] = "gotoOffset"
	$asMoves[$LOC_SHEETCUR_GOTO_START] = "gotoStart"
	$asMoves[$LOC_SHEETCUR_GOTO_END] = "gotoEnd"
	$asMoves[$LOC_SHEETCUR_GOTO_NEXT] = "gotoNext"
	$asMoves[$LOC_SHEETCUR_GOTO_PREV] = "gotoPrevious"
	$asMoves[$LOC_SHEETCUR_GOTO_USED_AREA_START] = "gotoStartOfUsedArea"
	$asMoves[$LOC_SHEETCUR_GOTO_USED_AREA_END] = "gotoEndOfUsedArea"

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iMove) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iMove, 0, UBound($asMoves) - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iColumns) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsInt($iRows) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not IsBool($bSelect) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	Switch $iMove
		Case $LOC_SHEETCUR_COLLAPSE_TO_SIZE, $LOC_SHEETCUR_GOTO_OFFSET
			Execute("$oCursor." & $asMoves[$iMove] & "(" & $iColumns & "," & $iRows & ")")

			Return SetError($__LO_STATUS_SUCCESS, 1, 1)

		Case $LOC_SHEETCUR_GOTO_NEXT, $LOC_SHEETCUR_GOTO_PREV
			Do
				Execute("$oCursor." & $asMoves[$iMove] & "()")
				$iCounted += 1

				Sleep((IsInt($iCounted / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
			Until ($iCounted >= $iCount)

			Return SetError($__LO_STATUS_SUCCESS, $iCounted, 1)

		Case $LOC_SHEETCUR_GOTO_USED_AREA_START, $LOC_SHEETCUR_GOTO_USED_AREA_END
			Execute("$oCursor." & $asMoves[$iMove] & "(" & $bSelect & ")")

			Return SetError($__LO_STATUS_SUCCESS, 1, 1)

		Case $LOC_SHEETCUR_COLLAPSE_TO_CURRENT_ARRAY, $LOC_SHEETCUR_COLLAPSE_TO_CURRENT_REGION, $LOC_SHEETCUR_COLLAPSE_TO_MERGED_AREA, _
				$LOC_SHEETCUR_EXPAND_TO_ENTIRE_COLUMN, $LOC_SHEETCUR_EXPAND_TO_ENTIRE_ROW, $LOC_SHEETCUR_GOTO_START, $LOC_SHEETCUR_GOTO_END
			Execute("$oCursor." & $asMoves[$iMove] & "()")

			Return SetError($__LO_STATUS_SUCCESS, 1, 1)

		Case Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	EndSwitch
EndFunc   ;==>__LOCalc_SheetCursorMove

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_TextCursorMove
; Description ...: For Text Cursor related movements.
; Syntax ........: __LOCalc_TextCursorMove(ByRef $oCursor, $iMove, $iCount[, $bSelect = False])
; Parameters ....: $oCursor             - [in/out] an object. A Text Cursor Object returned from any Text Cursor creation functions.
;                  $iMove               - an Integer value. The movement command constant. See remarks and Constants, $LOC_TEXTCUR_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iCount              - an integer value. Number of movements to make.
;                  $bSelect             - [optional] a boolean value. Default is False. If True, select data during this cursor movement.
; Return values .: Success: Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iMove not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iMove less than 0 or greater than highest move Constant. See Constants, $LOC_TEXTCUR_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iCount not an integer or is a negative.
;                  @Error 1 @Extended 5 Return 0 = $bSelect not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 2 Return 0 = Error processing cursor move.
;                  --Success--
;                  @Error 0 @Extended ? Return Boolean = Success, Cursor object movement was processed successfully. Returns True if the full count of movements were successful, else false if none or only partially successful. @Extended set to number of successful movements. Or Page Number for "gotoPage" command. See Remarks
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only some movements accept movement amounts and selecting (such as $LOC_TEXTCUR_GO_RIGHT 2, True) etc. Also only some accept creating/ extending a selection of text/ data. They will be specified below.
;                  To Clear /Unselect a current selection, you can input a move such as $LOC_TEXTCUR_GO_RIGHT, 0, False.
;                  #Cursor Movement Constants which accept number of Moves and Selecting:
;                  $LOC_TEXTCUR_GO_LEFT, Move the cursor left by n characters.
;                  $LOC_TEXTCUR_GO_RIGHT, Move the cursor right by n characters.
;                  #Cursor Movements which accept Selecting Only:
;                  $LOC_TEXTCUR_GOTO_START, Move the cursor to the start of the text.
;                  $LOC_TEXTCUR_GOTO_END, Move the cursor to the end of the text.
;                  #Cursor Movements which accept nothing and are done once per call:
;                  $LOC_TEXTCUR_COLLAPSE_TO_START,
;                  $LOC_TEXTCUR_COLLAPSE_TO_END (Collapses the current selection and moves the cursor to start or End of selection.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOCalc_TextCursorMove(ByRef $oCursor, $iMove, $iCount, $bSelect = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCounted = 0
	Local $bMoved = False
	Local $asMoves[6]

	$asMoves[$LOC_TEXTCUR_COLLAPSE_TO_START] = "collapseToStart"
	$asMoves[$LOC_TEXTCUR_COLLAPSE_TO_END] = "collapseToEnd"
	$asMoves[$LOC_TEXTCUR_GO_LEFT] = "goLeft"
	$asMoves[$LOC_TEXTCUR_GO_RIGHT] = "goRight"
	$asMoves[$LOC_TEXTCUR_GOTO_START] = "gotoStart"
	$asMoves[$LOC_TEXTCUR_GOTO_END] = "gotoEnd"

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iMove) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iMove, 0, UBound($asMoves) - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iCount) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bSelect) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	Switch $iMove
		Case $LOC_TEXTCUR_GO_LEFT, $LOC_TEXTCUR_GO_RIGHT
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "(" & $iCount & "," & $bSelect & ")")
			$iCounted = ($bMoved) ? ($iCount) : (0)

			Return SetError($__LO_STATUS_SUCCESS, $iCounted, $bMoved)

		Case $LOC_TEXTCUR_GOTO_START, $LOC_TEXTCUR_GOTO_END
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "(" & $bSelect & ")")
			$iCounted = ($bMoved) ? (1) : (0)

			Return SetError($__LO_STATUS_SUCCESS, $iCounted, $bMoved)

		Case $LOC_TEXTCUR_COLLAPSE_TO_START, $LOC_TEXTCUR_COLLAPSE_TO_END
			$bMoved = Execute("$oCursor." & $asMoves[$iMove] & "()")
			$iCounted = ($bMoved) ? (1) : (0)

			Return SetError($__LO_STATUS_SUCCESS, $iCounted, $bMoved)

		Case Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	EndSwitch
EndFunc   ;==>__LOCalc_TextCursorMove

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_TransparencyGradientConvert
; Description ...: Convert a Transparency Gradient percentage value to a color value or from a color value to a percentage.
; Syntax ........: __LOCalc_TransparencyGradientConvert([$iPercentToLong = Null[, $iLongToPercent = Null]])
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
Func __LOCalc_TransparencyGradientConvert($iPercentToLong = Null, $iLongToPercent = Null)
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
EndFunc   ;==>__LOCalc_TransparencyGradientConvert

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOCalc_TransparencyGradientNameInsert
; Description ...: Create and insert a new Transparency Gradient name.
; Syntax ........: __LOCalc_TransparencyGradientNameInsert(ByRef $oDoc, $tTGradient)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by a previous _LOCalc_DocOpen, _LOCalc_DocConnect, or _LOCalc_DocCreate function.
;                  $tTGradient          - a dll struct value. A Gradient Structure to copy settings from.
; Return values .: Success: String.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;                  @Error 1 @Extended 2 Return 0 = $tTGradient not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.drawing.TransparencyGradientTable" Object.
;                  @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.awt.Gradient" structure.
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
Func __LOCalc_TransparencyGradientNameInsert(ByRef $oDoc, $tTGradient)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
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
		Sleep((IsInt($iCount / $__LOCCONST_SLEEP_DIV)) ? (10) : (0))
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
EndFunc   ;==>__LOCalc_TransparencyGradientNameInsert
