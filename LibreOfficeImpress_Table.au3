#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Impress
#include "LibreOfficeImpress_Internal.au3"
#include "LibreOfficeImpress_Constants.au3"

; Other includes for Impress

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, and Deleting, etc. Impress Tables.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
; Note...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOImpress_TableBackColor
; _LOImpress_TableBackFillStyle
; _LOImpress_TableBackGradient
; _LOImpress_TableBackGradientMulticolor
; _LOImpress_TableBorderColor
; _LOImpress_TableBorderPadding
; _LOImpress_TableBorderStyle
; _LOImpress_TableBorderWidth
; _LOImpress_TableCellBackColor
; _LOImpress_TableCellBackFillStyle
; _LOImpress_TableCellBackGradient
; _LOImpress_TableCellBackGradientMulticolor
; _LOImpress_TableCellBorderColor
; _LOImpress_TableCellBorderPadding
; _LOImpress_TableCellBorderStyle
; _LOImpress_TableCellBorderWidth
; _LOImpress_TableCellCharEffect
; _LOImpress_TableCellCharFont
; _LOImpress_TableCellCharFontColor
; _LOImpress_TableCellCharOverLine
; _LOImpress_TableCellCharPosition
; _LOImpress_TableCellCharScaling
; _LOImpress_TableCellCharSpacing
; _LOImpress_TableCellCharStrikeOut
; _LOImpress_TableCellCharUnderLine
; _LOImpress_TableCellCreateTextCursor
; _LOImpress_TableCellGetObjByPosition
; _LOImpress_TableCellParAlignment
; _LOImpress_TableCellParIndent
; _LOImpress_TableCellParSpacing
; _LOImpress_TableCellParTabStopCreate
; _LOImpress_TableCellParTabStopDelete
; _LOImpress_TableCellParTabStopMod
; _LOImpress_TableCellParTabStopsGetList
; _LOImpress_TableCellString
; _LOImpress_TableCharEffect
; _LOImpress_TableCharFont
; _LOImpress_TableCharFontColor
; _LOImpress_TableCharOverLine
; _LOImpress_TableCharStrikeOut
; _LOImpress_TableCharUnderLine
; _LOImpress_TableColumnDelete
; _LOImpress_TableColumnGetCount
; _LOImpress_TableColumnInsert
; _LOImpress_TableInsert
; _LOImpress_TableRowDelete
; _LOImpress_TableRowGetCount
; _LOImpress_TableRowInsert
; _LOImpress_TableShadow
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableBackColor
; Description ...: Set or Retrieve the Background color of a Table.
; Syntax ........: _LOImpress_TableBackColor(ByRef $oTable[, $iBackColor = Null])
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
;                  $iBackColor          - [optional] (-1-16777215) Default is Null. The Table background color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for no background color.
; Return values .: Success: 1 or Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current setting as an Integer
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  @Error: 1, @Extended: 2 = $iBackColor not an Integer, less than -1 or greater than 16777215.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Cell Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve current Background color.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iBackColor
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  Tables require that the properties be set individually for each Cell, therefore this function cycles through each cell and sets the value, and may be slower for large tables.
;                  When retrieving the current property values for a table, if all of the cells in the Table do not have the same value, Null is returned for that property value.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOImpress_TableBackFillStyle, _LOImpress_TableBackGradient, _LOImpress_TableCellBackColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableBackColor(ByRef $oTable, $iBackColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCurColor, $iTempColor, $iError = 0
	Local $oCell

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	; If $iBackColor is Null, and Fill Style is set to solid, then return current color value, else return LO_COLOR_OFF.
	If __LO_VarsAreNull($iBackColor) Then
		For $iRow = 0 To $oTable.Model.RowCount() - 1
			For $iCol = 0 To $oTable.Model.ColumnCount() - 1
				$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
				If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

				If Not IsInt($iCurColor) Then     ; Retrieve the value once the first time, to use to test against the rest.
					If ($oCell.FillStyle() = $LOI_AREA_FILL_STYLE_SOLID) Then     ; If FillStyle is set to solid, then retrieve current color value, else return $LO_COLOR_OFF (Probably a Gradient is used or otherwise).
						$iCurColor = $oCell.FillColor()
						If Not IsInt($iCurColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

					Else
						$iCurColor = $LO_COLOR_OFF
					EndIf

				Else
					If ($oCell.FillStyle() = $LOI_AREA_FILL_STYLE_SOLID) Then     ; If FillStyle is set to solid, then retrieve current color value, else return $LO_COLOR_OFF (Probably a Gradient is used or otherwise).
						$iTempColor = $oCell.FillColor()
						If Not IsInt($iTempColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

					Else
						$iTempColor = $LO_COLOR_OFF
					EndIf

					If ($iTempColor <> $iCurColor) Then    ; Cycle through the Table and retrieve the current value for each cell, if it isn't the same, change the return to Null.
						$iCurColor = Null
						ExitLoop 2
					EndIf
				EndIf
			Next
		Next

		Return SetError($__LO_STATUS_SUCCESS, 1, $iCurColor)
	EndIf

	If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	For $iCol = 0 To $oTable.Model.ColumnCount() - 1
		For $iRow = 0 To $oTable.Model.RowCount() - 1
			$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
			If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			If ($iBackColor = $LO_COLOR_OFF) Then
				$oCell.FillStyle = $LOI_AREA_FILL_STYLE_OFF
				If ($oCell.PropertySetInfo.hasPropertyByName("FillUseSlideBackground")) Then $oCell.FillUseSlideBackground = False
				$iError = ($oCell.FillStyle() = $LOI_AREA_FILL_STYLE_OFF) ? ($iError) : (BitOR($iError, 1))

			Else
				$oCell.FillStyle = $LOI_AREA_FILL_STYLE_SOLID
				If ($oCell.PropertySetInfo.hasPropertyByName("FillUseSlideBackground")) Then $oCell.FillUseSlideBackground = False
				$oCell.FillColor = $iBackColor
				$iError = ($oCell.FillColor() = $iBackColor) ? ($iError) : (BitOR($iError, 1))
			EndIf
		Next
	Next

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_TableBackColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableBackFillStyle
; Description ...: Retrieve what kind of background fill is active, if any.
; Syntax ........: _LOImpress_TableBackFillStyle(ByRef $oTable)
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Returning current background fill style. Return will be one of the constants $LOI_AREA_FILL_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Cell Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve current fill style value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function is to help determine if a Gradient background, or a solid color background is currently active.
;                  This is useful because, if a Gradient is active, the solid color value is still present, and thus it would not be possible to determine which function should be used to retrieve the current values for, whether the Color function, or the Gradient function.
;                  When retrieving the current property values for a table, if all of the cells in the Table do not have the same value, Null is returned for that property value.
; Related .......: _LOImpress_TableBackColor, _LOImpress_TableBackGradient, _LOImpress_TableCellBackFillStyle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableBackFillStyle(ByRef $oTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFillStyle
	Local $oCell

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	For $iRow = 0 To $oTable.Model.RowCount() - 1
		For $iCol = 0 To $oTable.Model.ColumnCount() - 1
			$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
			If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			If Not IsInt($iFillStyle) Then         ; Retrieve the value once the first time, to use to test against the rest.
				$iFillStyle = $oCell.FillStyle()
				If Not IsInt($iFillStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			Else
				If ($oCell.FillStyle() <> $iFillStyle) Then      ; Cycle through the Table and retrieve the current value for each cell, if it isn't the same, change the return to Null.
					$iFillStyle = Null
					ExitLoop 2
				EndIf
			EndIf
		Next
	Next

	Return SetError($__LO_STATUS_SUCCESS, 0, $iFillStyle)
EndFunc   ;==>_LOImpress_TableBackFillStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableBackGradient
; Description ...: Set or Retrieve the settings for Table Background color Gradient.
; Syntax ........: _LOImpress_TableBackGradient(ByRef $oTable[, $sGradientName = Null[, $iType = Null[, $iIncrement = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iFromColor = Null[, $iToColor = Null[, $iFromIntense = Null[, $iToIntense = Null]]]]]]]]]]])
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
;                  $sGradientName       - [optional] Default is Null. A Preset Gradient Name. See remarks. See constants, $LOI_GRAD_NAME_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iType               - [optional] (-1-5) Default is Null. The gradient type to apply. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iIncrement          - [optional] (0, 3-256) Default is Null. The number of steps of color change. 0 = Automatic.
;                  $iXCenter            - [optional] (0-100) Default is Null. The horizontal offset for the gradient, where 0% corresponds to the current horizontal location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] (0-100) Default is Null. The vertical offset for the gradient, where 0% corresponds to the current vertical location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" Setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] (0-359) Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] (0-100) Default is Null. The amount by which to adjust the transparent area of the gradient. Set in percentage.
;                  $iFromColor          - [optional] (0-16777215) Default is Null. A color for the beginning point of the gradient, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iToColor            - [optional] (0-16777215) Default is Null. A color for the endpoint of the gradient, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iFromIntense        - [optional] (0-100) Default is Null. Enter the intensity for the color in the "From Color", where 0% corresponds to black, and 100 % to the selected color.
;                  $iToIntense          - [optional] (0-100) Default is Null. Enter the intensity for the color in the "To Color", where 0% corresponds to black, and 100 % to the selected color.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 11 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  @Error: 1, @Extended: 2 = $sGradientName not a String.
;                  @Error: 1, @Extended: 3 = $iType not an Integer, less than -1 or greater than 5. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 4 = $iIncrement not an Integer, less than 3, but not 0, or greater than 256.
;                  @Error: 1, @Extended: 5 = $iXCenter not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 6 = $iYCenter not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 7 = $iAngle not an Integer, less than 0 or greater than 359.
;                  @Error: 1, @Extended: 8 = $iTransitionStart not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 9 = $iFromColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 10 = $iToColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 11 = $iFromIntense not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 12 = $iToIntense not an Integer, less than 0 or greater than 100.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Parent Document Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Cell Object.
;                  @Error: 3, @Extended: 3 = Failed to retrieve "FillGradient" Object.
;                  @Error: 3, @Extended: 4 = Failed to retrieve ColorStops Array.
;                  @Error: 3, @Extended: 5 = Error creating Gradient Name.
;                  @Error: 3, @Extended: 6 = Error setting Gradient Name.
;                  @Error: 3, @Extended: 7 = Error retrieving Gradient Name.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sGradientName
;                  |                               2 = Error setting $iType
;                  |                               4 = Error setting $iIncrement
;                  |                               8 = Error setting $iXCenter
;                  |                               16 = Error setting $iYCenter
;                  |                               32 = Error setting $iAngle
;                  |                               64 = Error setting $iTransitionStart
;                  |                               128 = Error setting $iFromColor
;                  |                               256 = Error setting $iToColor
;                  |                               512 = Error setting $iFromIntense
;                  |                               1024 = Error setting $iToIntense
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Gradient Name has no use other than for applying a pre-existing preset gradient.
;                  Tables require that the properties be set individually for each Cell, therefore this function cycles through each cell and sets the value, and may be slower for large tables.
;                  When retrieving the current property values for a table, if all of the cells in the Table do not have the same value, Null is returned for all the property values.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOImpress_TableBackFillStyle, _LOImpress_TableBackGradientMulticolor, _LOImpress_TableCellBackGradient
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableBackGradient(ByRef $oTable, $sGradientName = Null, $iType = Null, $iIncrement = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iFromColor = Null, $iToColor = Null, $iFromIntense = Null, $iToIntense = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $oDoc, $oCell
	Local $avGradient[11], $avTemp[11]
	Local $atColorStop
	Local $sGradName

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oDoc = $oTable.Parent.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sGradientName, $iType, $iIncrement, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iFromColor, $iToColor, $iFromIntense, $iToIntense) Then
		For $iRow = 0 To $oTable.Model.RowCount() - 1
			For $iCol = 0 To $oTable.Model.ColumnCount() - 1
				$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
				If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

				$tStyleGradient = $oCell.FillGradient()
				If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

				If ($iRow = 0) And ($iCol = 0) Then     ; Retrieve the value once the first time, to use to test against the rest.
					__LO_ArrayFill($avGradient, $oCell.FillGradientName(), $tStyleGradient.Style(), _
							$oCell.FillGradientStepCount(), $tStyleGradient.XOffset(), $tStyleGradient.YOffset(), Int($tStyleGradient.Angle() / 10), _
							$tStyleGradient.Border(), $tStyleGradient.StartColor(), $tStyleGradient.EndColor(), $tStyleGradient.StartIntensity(), _
							$tStyleGradient.EndIntensity())     ; Angle is set in thousands

				Else
					__LO_ArrayFill($avTemp, $oCell.FillGradientName(), $tStyleGradient.Style(), _
							$oCell.FillGradientStepCount(), $tStyleGradient.XOffset(), $tStyleGradient.YOffset(), Int($tStyleGradient.Angle() / 10), _
							$tStyleGradient.Border(), $tStyleGradient.StartColor(), $tStyleGradient.EndColor(), $tStyleGradient.StartIntensity(), _
							$tStyleGradient.EndIntensity())     ; Angle is set in thousands

					For $i = 0 To UBound($avGradient) - 1
						If ($avGradient[$i] <> $avTemp[$i]) Then
							For $j = 0 To UBound($avGradient) - 1
								$avGradient[$j] = Null
							Next
							; Cycle through the Table and retrieve the current value for each cell, if it isn't the same, change the return to Null.
							; If one is different, all should be Null since they're all related, exit the loops.
							ExitLoop 3
						EndIf

						Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
					Next
				EndIf
			Next
		Next

		Return SetError($__LO_STATUS_SUCCESS, 1, $avGradient)
	EndIf

	For $iRow = 0 To $oTable.Model.RowCount() - 1
		For $iCol = 0 To $oTable.Model.ColumnCount() - 1
			$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
			If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$tStyleGradient = $oCell.FillGradient()
			If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			If IsString($sGradName) And ($sGradName <> "") Then
				If ($oCell.FillStyle() <> $LOI_AREA_FILL_STYLE_GRADIENT) Then $oCell.FillStyle = $LOI_AREA_FILL_STYLE_GRADIENT

				If ($iIncrement <> Null) Then
					$oCell.FillGradientStepCount = $iIncrement         ; I still have to set this, since I am skipping the function part.

					$iError = ($oCell.FillGradientStepCount() = $iIncrement) ? ($iError) : (BitOR($iError, 4))
				EndIf

				$oCell.FillGradientName = $sGradName
				If ($oCell.FillGradientName <> $sGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

				If ($sGradientName <> Null) Then         ; If Gradient name is set, just check to make sure it was set correctly for the rest of the cells.
					$iError = ($oCell.FillGradientName() = $sGradientName) ? ($iError) : (BitOR($iError, 1))
				EndIf

			Else
				If ($oCell.FillStyle() <> $LOI_AREA_FILL_STYLE_GRADIENT) Then $oCell.FillStyle = $LOI_AREA_FILL_STYLE_GRADIENT

				If ($sGradientName <> Null) Then
					If Not IsString($sGradientName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

					__LOImpress_GradientPresets($oDoc, $oCell, $tStyleGradient, $sGradientName)
					$iError = ($oCell.FillGradientName() = $sGradientName) ? ($iError) : (BitOR($iError, 1))
				EndIf

				If ($iType <> Null) Then
					If ($iType = $LOI_GRAD_TYPE_OFF) Then ; Turn Off Gradient
						$oCell.FillStyle = $LOI_AREA_FILL_STYLE_OFF
						$oCell.FillGradientName = ""

						ContinueLoop
					EndIf

					If Not __LO_IntIsBetween($iType, $LOI_GRAD_TYPE_LINEAR, $LOI_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

					$tStyleGradient.Style = $iType
				EndIf

				If ($iIncrement <> Null) Then
					If Not __LO_IntIsBetween($iIncrement, 3, 256, "", 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

					$oCell.FillGradientStepCount = $iIncrement
					$tStyleGradient.StepCount = $iIncrement ; Must set both of these in order for it to take effect.
					$iError = ($oCell.FillGradientStepCount() = $iIncrement) ? ($iError) : (BitOR($iError, 4))
				EndIf

				If ($iXCenter <> Null) Then
					If Not __LO_IntIsBetween($iXCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

					$tStyleGradient.XOffset = $iXCenter
				EndIf

				If ($iYCenter <> Null) Then
					If Not __LO_IntIsBetween($iYCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

					$tStyleGradient.YOffset = $iYCenter
				EndIf

				If ($iAngle <> Null) Then
					If Not __LO_IntIsBetween($iAngle, 0, 359) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

					$tStyleGradient.Angle = Int($iAngle * 10) ; Angle is set in thousands
				EndIf

				If ($iTransitionStart <> Null) Then
					If Not __LO_IntIsBetween($iTransitionStart, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

					$tStyleGradient.Border = $iTransitionStart
				EndIf

				If ($iFromColor <> Null) Then
					If Not __LO_IntIsBetween($iFromColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

					$tStyleGradient.StartColor = $iFromColor

					If __LO_VersionCheck(7.6) Then
						$atColorStop = $tStyleGradient.ColorStops()
						If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

						$tColorStop = $atColorStop[0] ; StopOffset 0 is the "From Color" Value.

						$tStopColor = $tColorStop.StopColor()

						$tStopColor.Red = (BitAND(BitShift($iFromColor, 16), 0xff) / 255)
						$tStopColor.Green = (BitAND(BitShift($iFromColor, 8), 0xff) / 255)
						$tStopColor.Blue = (BitAND($iFromColor, 0xff) / 255)

						$tColorStop.StopColor = $tStopColor

						$atColorStop[0] = $tColorStop

						$tStyleGradient.ColorStops = $atColorStop
					EndIf
				EndIf

				If ($iToColor <> Null) Then
					If Not __LO_IntIsBetween($iToColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

					$tStyleGradient.EndColor = $iToColor

					If __LO_VersionCheck(7.6) Then
						$atColorStop = $tStyleGradient.ColorStops()
						If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

						$tColorStop = $atColorStop[UBound($atColorStop) - 1] ; Last StopOffset is the "To Color" Value.

						$tStopColor = $tColorStop.StopColor()

						$tStopColor.Red = (BitAND(BitShift($iToColor, 16), 0xff) / 255)
						$tStopColor.Green = (BitAND(BitShift($iToColor, 8), 0xff) / 255)
						$tStopColor.Blue = (BitAND($iToColor, 0xff) / 255)

						$tColorStop.StopColor = $tStopColor

						$atColorStop[UBound($atColorStop) - 1] = $tColorStop

						$tStyleGradient.ColorStops = $atColorStop
					EndIf
				EndIf

				If ($iFromIntense <> Null) Then
					If Not __LO_IntIsBetween($iFromIntense, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

					$tStyleGradient.StartIntensity = $iFromIntense
				EndIf

				If ($iToIntense <> Null) Then
					If Not __LO_IntIsBetween($iToIntense, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

					$tStyleGradient.EndIntensity = $iToIntense
				EndIf

				If ($oCell.FillGradientName() = "") Or __LOImpress_GradientIsModified($tStyleGradient, $oCell.FillGradientName()) Then
					$sGradName = __LOImpress_GradientNameInsert($oDoc, $tStyleGradient)
					If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)

					$oCell.FillGradientName = $sGradName
					If ($oCell.FillGradientName <> $sGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)
				EndIf

				$oCell.FillGradient = $tStyleGradient

				; If Gradient is not turned off, then set rest of the Table cells to same Gradient name.
				$sGradName = $oCell.FillGradientName()
				If Not IsString($sGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 7, 0)
			EndIf

			; Error checking
			$iError = (__LO_VarsAreNull($iType)) ? ($iError) : (($oCell.FillGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 2)))
			$iError = (__LO_VarsAreNull($iXCenter)) ? ($iError) : (($oCell.FillGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 8)))
			$iError = (__LO_VarsAreNull($iYCenter)) ? ($iError) : (($oCell.FillGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 16)))
			$iError = (__LO_VarsAreNull($iAngle)) ? ($iError) : ((Int($oCell.FillGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 32)))
			$iError = (__LO_VarsAreNull($iTransitionStart)) ? ($iError) : (($oCell.FillGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 64)))
			$iError = (__LO_VarsAreNull($iFromColor)) ? ($iError) : (($oCell.FillGradient.StartColor() = $iFromColor) ? ($iError) : (BitOR($iError, 128)))
			$iError = (__LO_VarsAreNull($iToColor)) ? ($iError) : (($oCell.FillGradient.EndColor() = $iToColor) ? ($iError) : (BitOR($iError, 256)))
			$iError = (__LO_VarsAreNull($iFromIntense)) ? ($iError) : (($oCell.FillGradient.StartIntensity() = $iFromIntense) ? ($iError) : (BitOR($iError, 512)))
			$iError = (__LO_VarsAreNull($iToIntense)) ? ($iError) : (($oCell.FillGradient.EndIntensity() = $iToIntense) ? ($iError) : (BitOR($iError, 1024)))
		Next
	Next

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_TableBackGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableBackGradientMulticolor
; Description ...: Set or Retrieve a Table's Multicolor Gradient settings.
; Syntax ........: _LOImpress_TableBackGradientMulticolor(ByRef $oTable[, $avColorStops = Null])
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
;                  $avColorStops        - [optional] Default is Null. A Two column array of Colors and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: ?, Return: Array = Success. All optional parameters were called with Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
;                  Failure: 0 or Integer and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  @Error: 1, @Extended: 2 = $avColorStops not an Array, or does not contain two columns.
;                  @Error: 1, @Extended: 3 = $avColorStops contains less than two rows.
;                  @Error: 1, @Extended: 4 = ColorStop offset not a number, less than 0 or greater than 1.0. Returning problem element index.
;                  @Error: 1, @Extended: 5 = ColorStop color not an Integer, less than 0 or greater than 16777215. Returning problem element index.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create com.sun.star.awt.ColorStop Struct.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Cell Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve FillGradient Struct.
;                  @Error: 3, @Extended: 3 = Failed to retrieve ColorStops Array.
;                  @Error: 3, @Extended: 4 = Failed to retrieve StopColor Struct.
;                  @Error: 3, @Extended: 5 = Error retrieving Gradient Name.
;                  @Error: 3, @Extended: 6 = Error setting Gradient Name.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $avColorStops
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current version less than 7.6.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Starting with version 7.6 LibreOffice introduced an option to have multiple color stops in a Gradient rather than just a beginning and an ending color, but as of yet, the option is not available in the User Interface. However it has been made available in the API.
;                  The returned array will contain two columns, the first column will contain the ColorStop offset values, a number between 0 and 1.0. The second column will contain an Integer, the color value, as a RGB Color Integer.
;                  $avColorStops expects an array as described above.
;                  ColorStop offsets are sorted in ascending order, you can have more than one of the same value. There must be a minimum of two ColorStops. The first and last ColorStop offsets do not need to have an offset value of 0 and 1 respectively.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  Tables require that the properties be set individually for each Cell, therefore this function cycles through each cell and sets the value, and may be slower for large tables.
;                  When retrieving the current property values for a table, if all of the cells in the Table do not have the same value, a single row 2 columned array with Null values is returned.
; Related .......: _LO_GradientMulticolorAdd, _LO_GradientMulticolorDelete, _LO_GradientMulticolorModify, _LOImpress_TableBackGradient, _LOImpress_TableCellBackGradientMulticolor, _LOImpress_ShapeAreaTransparencyGradientMulti
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableBackGradientMulticolor(ByRef $oTable, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sGradName
	Local $oCell
	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $atColorStops[0]
	Local $avNewColorStops[0][2], $avTemp[0][2]
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	If __LO_VarsAreNull($avColorStops) Then
		For $iCol = 0 To $oTable.Model.ColumnCount() - 1
			For $iRow = 0 To $oTable.Model.RowCount() - 1
				$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
				If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

				$tStyleGradient = $oCell.FillGradient()
				If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

				$atColorStops = $tStyleGradient.ColorStops()
				If Not IsArray($atColorStops) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

				If ($iRow = 0) And ($iCol = 0) Then
					ReDim $avNewColorStops[UBound($atColorStops)][2]

					For $i = 0 To UBound($atColorStops) - 1
						$avNewColorStops[$i][0] = $atColorStops[$i].StopOffset()
						$tStopColor = $atColorStops[$i].StopColor()
						If Not IsObj($tStopColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

						$avNewColorStops[$i][1] = Int(BitShift(($tStopColor.Red() * 255), -16) + BitShift(($tStopColor.Green() * 255), -8) + ($tStopColor.Blue() * 255))     ; RGB to Long
						Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
					Next

				Else
					ReDim $avTemp[UBound($atColorStops)][2]

					For $i = 0 To UBound($atColorStops) - 1
						$avTemp[$i][0] = $atColorStops[$i].StopOffset()
						$tStopColor = $atColorStops[$i].StopColor()
						If Not IsObj($tStopColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

						$avTemp[$i][1] = Int(BitShift(($tStopColor.Red() * 255), -16) + BitShift(($tStopColor.Green() * 255), -8) + ($tStopColor.Blue() * 255))     ; RGB to Long
						Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
					Next

					If (UBound($avNewColorStops) <> UBound($avTemp)) Then ; If one batch of color stops isn't the same size, just Null the array and exit as it's obviously different.
						ReDim $avNewColorStops[1][2]
						$avNewColorStops[0][0] = Null
						$avNewColorStops[0][1] = Null
						ExitLoop 2
					EndIf

					For $i = 0 To UBound($avNewColorStops) - 1
						If ($avNewColorStops[$i][0] <> $avTemp[$i][0]) Or ($avNewColorStops[$i][1] <> $avTemp[$i][1]) Then
							ReDim $avNewColorStops[1][2]
							$avNewColorStops[0][0] = Null
							$avNewColorStops[0][1] = Null
							ExitLoop 3
						EndIf
					Next
				EndIf
			Next
		Next

		Return SetError($__LO_STATUS_SUCCESS, UBound($avNewColorStops), $avNewColorStops)
	EndIf

	If Not IsArray($avColorStops) Or (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (UBound($avColorStops) < 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	ReDim $atColorStops[UBound($avColorStops)]

	For $i = 0 To UBound($avColorStops) - 1
		$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
		If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tStopColor = $tColorStop.StopColor()
		If Not IsObj($tStopColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
		If Not __LO_NumIsBetween($avColorStops[$i][0], 0, 1.0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, $i)

		$tColorStop.StopOffset = $avColorStops[$i][0]

		If Not __LO_IntIsBetween($avColorStops[$i][1], $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)

		$tStopColor.Red = (BitAND(BitShift($avColorStops[$i][1], 16), 0xff) / 255)
		$tStopColor.Green = (BitAND(BitShift($avColorStops[$i][1], 8), 0xff) / 255)
		$tStopColor.Blue = (BitAND($avColorStops[$i][1], 0xff) / 255)

		$tColorStop.StopColor = $tStopColor

		$atColorStops[$i] = $tColorStop

		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	For $iRow = 0 To $oTable.Model.RowCount() - 1
		For $iCol = 0 To $oTable.Model.ColumnCount() - 1
			$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
			If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			If IsString($sGradName) And ($sGradName <> "") Then
				If ($oCell.FillStyle() <> $LOI_AREA_FILL_STYLE_GRADIENT) Then $oCell.FillStyle = $LOI_AREA_FILL_STYLE_GRADIENT

				$oCell.FillGradientName = $sGradName
				If ($oCell.FillGradientName() <> $sGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 6, 0)

				$iError = (UBound($avColorStops) = UBound($oCell.FillGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

			Else
				If ($oCell.FillStyle() <> $LOI_AREA_FILL_STYLE_GRADIENT) Then $oCell.FillStyle = $LOI_AREA_FILL_STYLE_GRADIENT

				$tStyleGradient = $oCell.FillGradient()
				If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

				$tStyleGradient.ColorStops = $atColorStops
				$oCell.FillGradient = $tStyleGradient

				$iError = (UBound($avColorStops) = UBound($oCell.FillGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

				$sGradName = $oCell.FillGradientName()
				If Not IsString($sGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 5, 0)
			EndIf
		Next
	Next

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_TableBackGradientMulticolor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableBorderColor
; Description ...: Set or Retrieve the Table Border Line Color. L.O. 3.6+.
; Syntax ........: _LOImpress_TableBorderColor(ByRef $oTable[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $iVert = Null[, $iHori = Null]]]]]])
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
;                  $iTop                - [optional] (0-16777215) Default is Null. The Top Border Line Color of the Table, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBottom             - [optional] (0-16777215) Default is Null. The Bottom Border Line Color of the Table, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLeft               - [optional] (0-16777215) Default is Null. The Left Border Line Color of the Table, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iRight              - [optional] (0-16777215) Default is Null. The Right Border Line Color of the Table, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iVert               - [optional] (0-16777215) Default is Null. The Vertical Border Line Color of the Table, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iHori               - [optional] (0-16777215) Default is Null. The Horizontal Border Line Color of the Table, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  @Error: 1, @Extended: 2 = $iTop not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 3 = $iBottom not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 4 = $iLeft not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 5 = $iRight not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 6 = $iVert not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 7 = $iHori not an Integer, less than 0 or greater than 16777215.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error Creating "com.sun.star.table.BorderLine2" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Cell Object.
;                  @Error: 3, @Extended: 3 = Failed to retrieve current Top border value.
;                  @Error: 3, @Extended: 4 = Failed to retrieve current Bottom border value.
;                  @Error: 3, @Extended: 5 = Failed to retrieve current Left border value.
;                  @Error: 3, @Extended: 6 = Failed to retrieve current Right border value.
;                  @Error: 3, @Extended: 7 = Failed to retrieve current Vertical border value.
;                  @Error: 3, @Extended: 8 = Failed to retrieve current Horizontal border value.
;                  @Error: 3, @Extended: 9 = Cannot set Top Border Style/Color when Top Border width not set.
;                  @Error: 3, @Extended: 10 = Cannot set Bottom Border Style/Color when Bottom Border width not set.
;                  @Error: 3, @Extended: 11 = Cannot set Left Border Style/Color when Left Border width not set.
;                  @Error: 3, @Extended: 12 = Cannot set Right Border Style/Color when Right Border width not set.
;                  @Error: 3, @Extended: 13 = Cannot set Vertical Border Style/Color when Vertical Border width not set.
;                  @Error: 3, @Extended: 14 = Cannot set Horizontal Border Style/Color when Horizontal Border width not set.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  |                               16 = Error setting $iVert
;                  |                               32 = Error setting $iHori
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 3.6.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Tables require that the properties be set individually for each Cell, therefore this function cycles through each cell and sets the value, and may be slower for large tables.
;                  When retrieving the current property values for a table, if all of the cells in the Table do not have the same value, Null is returned for that property value.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOImpress_TableBorderWidth, _LOImpress_TableBorderStyle, _LOImpress_TableBorderPadding, _LOImpress_TableCellBorderColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableBorderColor(ByRef $oTable, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $iVert = Null, $iHori = Null)
	Local $vReturn

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iVert <> Null) And Not __LO_IntIsBetween($iVert, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If ($iHori <> Null) And Not __LO_IntIsBetween($iHori, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$vReturn = __LOImpress_TableBorder($oTable, False, False, True, $iTop, $iBottom, $iLeft, $iRight, $iVert, $iHori)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableBorderPadding
; Description ...: Set or retrieve the Table Border Padding settings.
; Syntax ........: _LOImpress_TableBorderPadding(ByRef $oTable[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
;                  $iTop                - [optional] Default is Null. The Top Distance between the Border and Table contents in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] Default is Null. The Bottom Distance between the Border and Table contents in Hundredths of a Millimeter (HMM).
;                  $iLeft               - [optional] Default is Null. The Left Distance between the Border and Table contents in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] Default is Null. The Right Distance between the Border and Table contents in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  @Error: 1, @Extended: 2 = $iTop not an Integer.
;                  @Error: 1, @Extended: 3 = $iBottom not an Integer.
;                  @Error: 1, @Extended: 4 = $Left not an Integer.
;                  @Error: 1, @Extended: 5 = $iRight not an Integer.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Cell Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Tables require that the properties be set individually for each Cell, therefore this function cycles through each cell and sets the value, and may be slower for large tables.
;                  When retrieving the current property values for a table, if all of the cells in the Table do not have the same value, Null is returned for that property value.
; Related .......: _LO_UnitConvert, _LOImpress_TableBorderWidth, _LOImpress_TableBorderStyle, _LOImpress_TableBorderColor, _LOImpress_TableCellBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableBorderPadding(ByRef $oTable, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $oCell
	Local $iTempCount = 0
	Local $aiBPadding[4], $aiTemp[0]

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then
		For $iRow = 0 To $oTable.Model.RowCount() - 1
			For $iCol = 0 To $oTable.Model.ColumnCount() - 1
				$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
				If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

				If ($iRow = 0) And ($iCol = 0) Then     ; Retrieve the value once the first time, to use to test against the rest.
					__LO_ArrayFill($aiBPadding, $oCell.TextUpperDistance(), $oCell.TextLowerDistance(), $oCell.TextLeftDistance(), $oCell.TextRightDistance())

				Else
					__LO_ArrayFill($aiTemp, $oCell.TextUpperDistance(), $oCell.TextLowerDistance(), $oCell.TextLeftDistance(), $oCell.TextRightDistance())

					$iTempCount = 0

					; Cycle through the Table and retrieve the current value for each cell, if it isn't the same, change the return to Null.
					For $i = 0 To UBound($aiBPadding) - 1
						If ($aiBPadding[$i] <> $aiTemp[$i]) Then
							$aiBPadding[$i] = Null
							$iTempCount += 1

						ElseIf ($aiBPadding[$i] = Null) Then
							$iTempCount += 1
						EndIf

						If ($iTempCount = UBound($aiBPadding)) Then ExitLoop 3 ; Exit the loops if all values are already nulled.

						Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
					Next
				EndIf
			Next
		Next

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiBPadding)
	EndIf

	For $iRow = 0 To $oTable.Model.RowCount() - 1
		For $iCol = 0 To $oTable.Model.ColumnCount() - 1
			$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
			If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			If ($iTop <> Null) Then
				If Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

				$oCell.TextUpperDistance = $iTop
				$iError = (__LO_IntIsBetween($oCell.TextUpperDistance(), $iTop - 1, $iTop + 1)) ? ($iError) : (BitOR($iError, 1))
			EndIf

			If ($iBottom <> Null) Then
				If Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

				$oCell.TextLowerDistance = $iBottom
				$iError = (__LO_IntIsBetween($oCell.TextLowerDistance(), $iBottom - 1, $iBottom + 1)) ? ($iError) : (BitOR($iError, 2))
			EndIf

			If ($iLeft <> Null) Then
				If Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

				$oCell.TextLeftDistance = $iLeft
				$iError = (__LO_IntIsBetween($oCell.TextLeftDistance(), $iLeft - 1, $iLeft + 1)) ? ($iError) : (BitOR($iError, 4))
			EndIf

			If ($iRight <> Null) Then
				If Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

				$oCell.TextRightDistance = $iRight
				$iError = (__LO_IntIsBetween($oCell.TextRightDistance(), $iRight - 1, $iRight + 1)) ? ($iError) : (BitOR($iError, 8))
			EndIf
		Next
	Next

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_TableBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableBorderStyle
; Description ...: Set or Retrieve the Table Border Line style. L.O. 3.6+.
; Syntax ........: _LOImpress_TableBorderStyle(ByRef $oTable[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $iVert = Null[, $iHori = Null]]]]]])
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
;                  $iTop                - [optional] (0x7FFF,0-17) Default is Null. The Top Border Line Style of the Table. See Constants, $LOI_SHAPE_BORDER_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iBottom             - [optional] (0x7FFF,0-17) Default is Null. The Bottom Border Line Style of the Table. See Constants, $LOI_SHAPE_BORDER_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iLeft               - [optional] (0x7FFF,0-17) Default is Null. The Left Border Line Style of the Table. See Constants, $LOI_SHAPE_BORDER_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iRight              - [optional] (0x7FFF,0-17) Default is Null. The Right Border Line Style of the Table. See Constants, $LOI_SHAPE_BORDER_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iVert               - [optional] (0x7FFF,0-17) Default is Null. The internal Vertical Border Line Styles of the Table. See Constants, $LOI_SHAPE_BORDER_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iHori               - [optional] (0x7FFF,0-17) Default is Null. The internal Horizontal Border Line Styles of the Table. See Constants, $LOI_SHAPE_BORDER_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  @Error: 1, @Extended: 2 = $iTop not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See Constants, $LOI_SHAPE_BORDER_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iBottom not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See Constants, $LOI_SHAPE_BORDER_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 4 = $iLeft not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See Constants, $LOI_SHAPE_BORDER_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 5 = $iRight not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See Constants, $LOI_SHAPE_BORDER_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 6 = $iVert not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See Constants, $LOI_SHAPE_BORDER_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 7 = $iHori not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See Constants, $LOI_SHAPE_BORDER_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error Creating "com.sun.star.table.BorderLine2" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Cell Object.
;                  @Error: 3, @Extended: 3 = Failed to retrieve current Top border value.
;                  @Error: 3, @Extended: 4 = Failed to retrieve current Bottom border value.
;                  @Error: 3, @Extended: 5 = Failed to retrieve current Left border value.
;                  @Error: 3, @Extended: 6 = Failed to retrieve current Right border value.
;                  @Error: 3, @Extended: 7 = Failed to retrieve current Vertical border value.
;                  @Error: 3, @Extended: 8 = Failed to retrieve current Horizontal border value.
;                  @Error: 3, @Extended: 9 = Cannot set Top Border Style/Color when Top Border width not set.
;                  @Error: 3, @Extended: 10 = Cannot set Bottom Border Style/Color when Bottom Border width not set.
;                  @Error: 3, @Extended: 11 = Cannot set Left Border Style/Color when Left Border width not set.
;                  @Error: 3, @Extended: 12 = Cannot set Right Border Style/Color when Right Border width not set.
;                  @Error: 3, @Extended: 13 = Cannot set Vertical Border Style/Color when Vertical Border width not set.
;                  @Error: 3, @Extended: 14 = Cannot set Horizontal Border Style/Color when Horizontal Border width not set.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  |                               16 = Error setting $iVert
;                  |                               32 = Error setting $iHori
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 3.6.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Tables require that the properties be set individually for each Cell, therefore this function cycles through each cell and sets the value, and may be slower for large tables.
;                  When retrieving the current property values for a table, if all of the cells in the Table do not have the same value, Null is returned for that property value.
; Related .......: _LOImpress_TableBorderWidth, _LOImpress_TableBorderColor, _LOImpress_TableBorderPadding, _LOImpress_TableCellBorderStyle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableBorderStyle(ByRef $oTable, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $iVert = Null, $iHori = Null)
	Local $vReturn

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LOI_SHAPE_BORDER_STYLE_SOLID, $LOI_SHAPE_BORDER_STYLE_DASH_DOT_DOT, "", $LOI_SHAPE_BORDER_STYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LOI_SHAPE_BORDER_STYLE_SOLID, $LOI_SHAPE_BORDER_STYLE_DASH_DOT_DOT, "", $LOI_SHAPE_BORDER_STYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LOI_SHAPE_BORDER_STYLE_SOLID, $LOI_SHAPE_BORDER_STYLE_DASH_DOT_DOT, "", $LOI_SHAPE_BORDER_STYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LOI_SHAPE_BORDER_STYLE_SOLID, $LOI_SHAPE_BORDER_STYLE_DASH_DOT_DOT, "", $LOI_SHAPE_BORDER_STYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iVert <> Null) And Not __LO_IntIsBetween($iVert, $LOI_SHAPE_BORDER_STYLE_SOLID, $LOI_SHAPE_BORDER_STYLE_DASH_DOT_DOT, "", $LOI_SHAPE_BORDER_STYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If ($iHori <> Null) And Not __LO_IntIsBetween($iHori, $LOI_SHAPE_BORDER_STYLE_SOLID, $LOI_SHAPE_BORDER_STYLE_DASH_DOT_DOT, "", $LOI_SHAPE_BORDER_STYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$vReturn = __LOImpress_TableBorder($oTable, False, True, False, $iTop, $iBottom, $iLeft, $iRight, $iVert, $iHori)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableBorderWidth
; Description ...: Set or Retrieve the Table Border Line Width. L.O. 3.6+.
; Syntax ........: _LOImpress_TableBorderWidth(ByRef $oTable[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $iVert = Null[, $iHori = Null]]]]]])
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
;                  $iTop                - [optional] Default is Null. The Top Border Line width of the Table in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOI_SHAPE_BORDER_WIDTH_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iBottom             - [optional] Default is Null. The Bottom Border Line Width of the Table in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOI_SHAPE_BORDER_WIDTH_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iLeft               - [optional] Default is Null. The Left Border Line width of the Table in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOI_SHAPE_BORDER_WIDTH_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iRight              - [optional] Default is Null. The Right Border Line Width of the Table in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOI_SHAPE_BORDER_WIDTH_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iVert               - [optional] Default is Null. The Internal Vertical Border Line width of the Table in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOI_SHAPE_BORDER_WIDTH_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iHori               - [optional] Default is Null. The Internal Horizontal Border Line width of the Table in Hundredths of a Millimeter (HMM). Can be a custom value, or one of the constants, $LOI_SHAPE_BORDER_WIDTH_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  @Error: 1, @Extended: 2 = $iTop not an Integer, or less than 0.
;                  @Error: 1, @Extended: 3 = $iBottom not an Integer, or less than 0.
;                  @Error: 1, @Extended: 4 = $iLeft not an Integer, or less than 0.
;                  @Error: 1, @Extended: 5 = $iRight not an Integer, or less than 0.
;                  @Error: 1, @Extended: 6 = $iVert not an Integer, or less than 0.
;                  @Error: 1, @Extended: 7 = $iHori not an Integer, or less than 0.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error Creating "com.sun.star.table.BorderLine2" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Cell Object.
;                  @Error: 3, @Extended: 3 = Failed to retrieve current Top border value.
;                  @Error: 3, @Extended: 4 = Failed to retrieve current Bottom border value.
;                  @Error: 3, @Extended: 5 = Failed to retrieve current Left border value.
;                  @Error: 3, @Extended: 6 = Failed to retrieve current Right border value.
;                  @Error: 3, @Extended: 7 = Failed to retrieve current Vertical border value.
;                  @Error: 3, @Extended: 8 = Failed to retrieve current Horizontal border value.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  |                               16 = Error setting $iVert
;                  |                               32 = Error setting $iHori
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 3.6.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To "Turn Off" Borders, set them to 0
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Tables require that the properties be set individually for each Cell, therefore this function cycles through each cell and sets the value, and may be slower for large tables.
;                  When retrieving the current property values for a table, if all of the cells in the Table do not have the same value, Null is returned for that property value.
; Related .......: _LO_UnitConvert, _LOImpress_TableBorderStyle, _LOImpress_TableBorderColor, _LOImpress_TableBorderPadding, _LOImpress_TableCellBorderWidth
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableBorderWidth(ByRef $oTable, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $iVert = Null, $iHori = Null)
	Local $vReturn

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iVert <> Null) And Not __LO_IntIsBetween($iVert, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If ($iHori <> Null) And Not __LO_IntIsBetween($iHori, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$vReturn = __LOImpress_TableBorder($oTable, True, False, False, $iTop, $iBottom, $iLeft, $iRight, $iVert, $iHori)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellBackColor
; Description ...: Set or Retrieve the Background color of a Cell.
; Syntax ........: _LOImpress_TableCellBackColor(ByRef $oCell[, $iBackColor = Null])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $iBackColor          - [optional] (-1-16777215) Default is Null. The Cell background color as a RGB Color Integer. Call with $LO_COLOR_OFF(-1) to disable Background color. Can also be one of the constants $LO_COLOR_* as defined in LibreOffice_Constants.au3
; Return values .: Success: 1 or Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current setting as an Integer
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  @Error: 1, @Extended: 2 = $iBackColor not an Integer, set less than -1 or greater than 16777215.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Background color.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iBackColor
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOImpress_TableCellBackFillStyle, _LOImpress_TableCellBackGradient, _LOImpress_TableBackColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellBackColor(ByRef $oCell, $iBackColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $iCurColor

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iBackColor) Then
		If ($oCell.FillStyle() = $LOI_AREA_FILL_STYLE_SOLID) Then ; If FillStyle is set to solid, then retrieve current color value, else return $LO_COLOR_OFF (Probably a Gradient is used or otherwise).
			$iCurColor = __LOImpress_ColorRemoveAlpha($oCell.FillColor())
			If Not IsInt($iCurColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Else
			$iCurColor = $LO_COLOR_OFF
		EndIf

		Return SetError($__LO_STATUS_SUCCESS, 0, $iCurColor)
	EndIf

	If Not __LO_IntIsBetween($iBackColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If ($iBackColor = $LO_COLOR_OFF) Then
		$oCell.FillStyle = $LOI_AREA_FILL_STYLE_OFF
		If ($oCell.PropertySetInfo.hasPropertyByName("FillUseSlideBackground")) Then $oCell.FillUseSlideBackground = False

	Else
		$oCell.FillStyle = $LOI_AREA_FILL_STYLE_SOLID
		If ($oCell.PropertySetInfo.hasPropertyByName("FillUseSlideBackground")) Then $oCell.FillUseSlideBackground = False
		$oCell.FillColor = $iBackColor
		$iError = ($oCell.FillColor() = $iBackColor) ? ($iError) : (BitOR($iError, 1))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_TableCellBackColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellBackFillStyle
; Description ...: Retrieve what kind of background fill is active, if any.
; Syntax ........: _LOImpress_TableCellBackFillStyle(ByRef $oCell)
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Returning current background fill style. Return will be one of the constants $LOI_AREA_FILL_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current fill style value for Cell.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function is to help determine if a Gradient background, or a solid color background is currently active.
;                  This is useful because, if a Gradient is active, the solid color value is still present, and thus it would not be possible to determine which function should be used to retrieve the current values for, whether the Color function, or the Gradient function.
; Related .......: _LOImpress_TableCellBackColor, _LOImpress_TableCellBackGradient, _LOImpress_TableBackFillStyle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellBackFillStyle(ByRef $oCell)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iFillStyle

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iFillStyle = $oCell.FillStyle()
	If Not IsInt($iFillStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iFillStyle)
EndFunc   ;==>_LOImpress_TableCellBackFillStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellBackGradient
; Description ...: Set or retrieve the settings for Cell Background color Gradient.
; Syntax ........: _LOImpress_TableCellBackGradient(ByRef $oDoc, ByRef $oCell[, $sGradientName = Null[, $iType = Null[, $iIncrement = Null[, $iXCenter = Null[, $iYCenter = Null[, $iAngle = Null[, $iTransitionStart = Null[, $iFromColor = Null[, $iToColor = Null[, $iFromIntense = Null[, $iToIntense = Null]]]]]]]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $sGradientName       - [optional] Default is Null. A Preset Gradient Name. See remarks. See constants, $LOI_GRAD_NAME_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iType               - [optional] (-1-5) Default is Null. The gradient type to apply. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iIncrement          - [optional] (0, 3-256) Default is Null. The number of steps of color change. 0 = Automatic.
;                  $iXCenter            - [optional] (0-100) Default is Null. The horizontal offset for the gradient, where 0% corresponds to the current horizontal location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iYCenter            - [optional] (0-100) Default is Null. The vertical offset for the gradient, where 0% corresponds to the current vertical location of the endpoint color in the gradient. The endpoint color is the color that is selected in the "To Color" Setting. Set in percentage. $iType must be other than "Linear", or "Axial".
;                  $iAngle              - [optional] (0-359) Default is Null. The rotation angle for the gradient. Set in degrees. $iType must be other than "Radial".
;                  $iTransitionStart    - [optional] (0-100) Default is Null. The amount by which to adjust the transparent area of the gradient. Set in percentage.
;                  $iFromColor          - [optional] (0-16777215) Default is Null. A color for the beginning point of the gradient, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iToColor            - [optional] (0-16777215) Default is Null. A color for the endpoint of the gradient, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iFromIntense        - [optional] (0-100) Default is Null. Enter the intensity for the color in the "From Color", where 0% corresponds to black, and 100 % to the selected color.
;                  $iToIntense          - [optional] (0-100) Default is Null. Enter the intensity for the color in the "To Color", where 0% corresponds to black, and 100 % to the selected color.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings have been successfully set.
;                  @Error: 0, @Extended: 0, Return: 2 = Success. Gradient has been successfully turned off.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 11 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oCell not an Object.
;                  @Error: 1, @Extended: 3 = $sGradientName not a String.
;                  @Error: 1, @Extended: 4 = $iType not an Integer, less than -1 or greater than 5. See Constants, $LOI_GRAD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 5 = $iIncrement not an Integer, less than 3, but not 0, or greater than 256.
;                  @Error: 1, @Extended: 6 = $iXCenter not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 7 = $iYCenter not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 8 = $iAngle not an Integer, less than 0 or greater than 359.
;                  @Error: 1, @Extended: 9 = $iTransitionStart not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 10 = $iFromColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 11 = $iToColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 12 = $iFromIntense not an Integer, less than 0 or greater than 100.
;                  @Error: 1, @Extended: 13 = $iToIntense not an Integer, less than 0 or greater than 100.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving "FillGradient" Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve ColorStops Array.
;                  @Error: 3, @Extended: 3 = Error creating Gradient Name.
;                  @Error: 3, @Extended: 4 = Error setting Gradient Name.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $sGradientName
;                  |                               2 = Error setting $iType
;                  |                               4 = Error setting $iIncrement
;                  |                               8 = Error setting $iXCenter
;                  |                               16 = Error setting $iYCenter
;                  |                               32 = Error setting $iAngle
;                  |                               64 = Error setting $iTransitionStart
;                  |                               128 = Error setting $iFromColor
;                  |                               256 = Error setting $iToColor
;                  |                               512 = Error setting $iFromIntense
;                  |                               1024 = Error setting $iToIntense
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Gradient Name has no use other than for applying a pre-existing preset gradient.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOImpress_TableCellBackFillStyle, _LOImpress_TableCellBackGradientMulticolor, _LOImpress_TableBackGradient
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellBackGradient(ByRef $oDoc, ByRef $oCell, $sGradientName = Null, $iType = Null, $iIncrement = Null, $iXCenter = Null, $iYCenter = Null, $iAngle = Null, $iTransitionStart = Null, $iFromColor = Null, $iToColor = Null, $iFromIntense = Null, $iToIntense = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $avGradient[11]
	Local $sGradName
	Local $atColorStop[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$tStyleGradient = $oCell.FillGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($sGradientName, $iType, $iIncrement, $iXCenter, $iYCenter, $iAngle, $iTransitionStart, $iFromColor, $iToColor, $iFromIntense, $iToIntense) Then
		__LO_ArrayFill($avGradient, $oCell.FillGradientName(), $tStyleGradient.Style(), _
				$oCell.FillGradientStepCount(), $tStyleGradient.XOffset(), $tStyleGradient.YOffset(), Int($tStyleGradient.Angle() / 10), _
				$tStyleGradient.Border(), $tStyleGradient.StartColor(), $tStyleGradient.EndColor(), $tStyleGradient.StartIntensity(), _
				$tStyleGradient.EndIntensity()) ; Angle is set in thousands

		Return SetError($__LO_STATUS_SUCCESS, 1, $avGradient)
	EndIf

	If ($oCell.FillStyle() <> $LOI_AREA_FILL_STYLE_GRADIENT) Then $oCell.FillStyle = $LOI_AREA_FILL_STYLE_GRADIENT

	If ($sGradientName <> Null) Then
		If Not IsString($sGradientName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		__LOImpress_GradientPresets($oDoc, $oCell, $tStyleGradient, $sGradientName)
		$iError = ($oCell.FillGradientName() = $sGradientName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iType <> Null) Then
		If ($iType = $LOI_GRAD_TYPE_OFF) Then ; Turn Off Gradient
			$oCell.FillStyle = $LOI_AREA_FILL_STYLE_OFF
			$oCell.FillGradientName = ""

			Return SetError($__LO_STATUS_SUCCESS, 0, 2)
		EndIf

		If Not __LO_IntIsBetween($iType, $LOI_GRAD_TYPE_LINEAR, $LOI_GRAD_TYPE_RECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$tStyleGradient.Style = $iType
	EndIf

	If ($iIncrement <> Null) Then
		If Not __LO_IntIsBetween($iIncrement, 3, 256, "", 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oCell.FillGradientStepCount = $iIncrement
		$tStyleGradient.StepCount = $iIncrement ; Must set both of these in order for it to take effect.
		$iError = ($oCell.FillGradientStepCount() = $iIncrement) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iXCenter <> Null) Then
		If Not __LO_IntIsBetween($iXCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

		$tStyleGradient.XOffset = $iXCenter
	EndIf

	If ($iYCenter <> Null) Then
		If Not __LO_IntIsBetween($iYCenter, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

		$tStyleGradient.YOffset = $iYCenter
	EndIf

	If ($iAngle <> Null) Then
		If Not __LO_IntIsBetween($iAngle, 0, 359) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

		$tStyleGradient.Angle = Int($iAngle * 10) ; Angle is set in thousands
	EndIf

	If ($iTransitionStart <> Null) Then
		If Not __LO_IntIsBetween($iTransitionStart, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)

		$tStyleGradient.Border = $iTransitionStart
	EndIf

	If ($iFromColor <> Null) Then
		If Not __LO_IntIsBetween($iFromColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)

		$tStyleGradient.StartColor = $iFromColor

		If __LO_VersionCheck(7.6) Then
			$atColorStop = $tStyleGradient.ColorStops()
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$tColorStop = $atColorStop[0] ; StopOffset 0 is the "From Color" Value.

			$tStopColor = $tColorStop.StopColor()

			$tStopColor.Red = (BitAND(BitShift($iFromColor, 16), 0xff) / 255)
			$tStopColor.Green = (BitAND(BitShift($iFromColor, 8), 0xff) / 255)
			$tStopColor.Blue = (BitAND($iFromColor, 0xff) / 255)

			$tColorStop.StopColor = $tStopColor

			$atColorStop[0] = $tColorStop

			$tStyleGradient.ColorStops = $atColorStop
		EndIf
	EndIf

	If ($iToColor <> Null) Then
		If Not __LO_IntIsBetween($iToColor, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)

		$tStyleGradient.EndColor = $iToColor

		If __LO_VersionCheck(7.6) Then
			$atColorStop = $tStyleGradient.ColorStops()
			If Not IsArray($atColorStop) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			$tColorStop = $atColorStop[UBound($atColorStop) - 1] ; Last StopOffset is the "To Color" Value.

			$tStopColor = $tColorStop.StopColor()

			$tStopColor.Red = (BitAND(BitShift($iToColor, 16), 0xff) / 255)
			$tStopColor.Green = (BitAND(BitShift($iToColor, 8), 0xff) / 255)
			$tStopColor.Blue = (BitAND($iToColor, 0xff) / 255)

			$tColorStop.StopColor = $tStopColor

			$atColorStop[UBound($atColorStop) - 1] = $tColorStop

			$tStyleGradient.ColorStops = $atColorStop
		EndIf
	EndIf

	If ($iFromIntense <> Null) Then
		If Not __LO_IntIsBetween($iFromIntense, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 12, 0)

		$tStyleGradient.StartIntensity = $iFromIntense
	EndIf

	If ($iToIntense <> Null) Then
		If Not __LO_IntIsBetween($iToIntense, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 13, 0)

		$tStyleGradient.EndIntensity = $iToIntense
	EndIf

	If ($oCell.FillGradientName() = "") Or __LOImpress_GradientIsModified($tStyleGradient, $oCell.FillGradientName()) Then
		$sGradName = __LOImpress_GradientNameInsert($oDoc, $tStyleGradient)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

		$oCell.FillGradientName = $sGradName
		If ($oCell.FillGradientName <> $sGradName) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)
	EndIf

	$oCell.FillGradient = $tStyleGradient

	; Error checking
	$iError = (__LO_VarsAreNull($iType)) ? ($iError) : (($oCell.FillGradient.Style() = $iType) ? ($iError) : (BitOR($iError, 2)))
	$iError = (__LO_VarsAreNull($iXCenter)) ? ($iError) : (($oCell.FillGradient.XOffset() = $iXCenter) ? ($iError) : (BitOR($iError, 8)))
	$iError = (__LO_VarsAreNull($iYCenter)) ? ($iError) : (($oCell.FillGradient.YOffset() = $iYCenter) ? ($iError) : (BitOR($iError, 16)))
	$iError = (__LO_VarsAreNull($iAngle)) ? ($iError) : ((Int($oCell.FillGradient.Angle() / 10) = $iAngle) ? ($iError) : (BitOR($iError, 32)))
	$iError = (__LO_VarsAreNull($iTransitionStart)) ? ($iError) : (($oCell.FillGradient.Border() = $iTransitionStart) ? ($iError) : (BitOR($iError, 64)))
	$iError = (__LO_VarsAreNull($iFromColor)) ? ($iError) : (($oCell.FillGradient.StartColor() = $iFromColor) ? ($iError) : (BitOR($iError, 128)))
	$iError = (__LO_VarsAreNull($iToColor)) ? ($iError) : (($oCell.FillGradient.EndColor() = $iToColor) ? ($iError) : (BitOR($iError, 256)))
	$iError = (__LO_VarsAreNull($iFromIntense)) ? ($iError) : (($oCell.FillGradient.StartIntensity() = $iFromIntense) ? ($iError) : (BitOR($iError, 512)))
	$iError = (__LO_VarsAreNull($iToIntense)) ? ($iError) : (($oCell.FillGradient.EndIntensity() = $iToIntense) ? ($iError) : (BitOR($iError, 1024)))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_TableCellBackGradient

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellBackGradientMulticolor
; Description ...: Set or Retrieve a Cell's Multicolor Gradient settings.
; Syntax ........: _LOImpress_TableCellBackGradientMulticolor(ByRef $oCell[, $avColorStops = Null])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $avColorStops        - [optional] Default is Null. A Two column array of Colors and ColorStop offsets. See remarks.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: ?, Return: Array = Success. All optional parameters were called with Null, returning current Array of ColorStops. See remarks. @Extended set to number of ColorStops returned.
;                  Failure: 0 or Integer and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  @Error: 1, @Extended: 2 = $avColorStops not an Array, or does not contain two columns.
;                  @Error: 1, @Extended: 3 = $avColorStops contains less than two rows.
;                  @Error: 1, @Extended: 4 = ColorStop offset not a number, less than 0 or greater than 1.0. Returning problem element index.
;                  @Error: 1, @Extended: 5 = ColorStop color not an Integer, less than 0 or greater than 16777215. Returning problem element index.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create com.sun.star.awt.ColorStop Struct.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve FillGradient Struct.
;                  @Error: 3, @Extended: 2 = Failed to retrieve ColorStops Array.
;                  @Error: 3, @Extended: 3 = Failed to retrieve StopColor Struct.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $avColorStops
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current version less than 7.6.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Starting with version 7.6 LibreOffice introduced an option to have multiple color stops in a Gradient rather than just a beginning and an ending color, but as of yet, the option is not available in the User Interface. However it has been made available in the API.
;                  The returned array will contain two columns, the first column will contain the ColorStop offset values, a number between 0 and 1.0. The second column will contain an Integer, the color value, as a RGB Color Integer.
;                  $avColorStops expects an array as described above.
;                  ColorStop offsets are sorted in ascending order, you can have more than one of the same value. There must be a minimum of two ColorStops. The first and last ColorStop offsets do not need to have an offset value of 0 and 1 respectively.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
; Related .......: _LO_GradientMulticolorAdd, _LO_GradientMulticolorDelete, _LO_GradientMulticolorModify, _LOImpress_TableCellBackGradient, _LOImpress_ShapeAreaTransparencyGradientMulti, _LOImpress_TableBackGradientMulticolor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellBackGradientMulticolor(ByRef $oCell, $avColorStops = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $tStyleGradient, $tColorStop, $tStopColor
	Local $iError = 0
	Local $atColorStops[0]
	Local $avNewColorStops[0][2]
	Local Const $__UBOUND_COLUMNS = 2

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_VersionCheck(7.6) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

	$tStyleGradient = $oCell.FillGradient()
	If Not IsObj($tStyleGradient) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	If __LO_VarsAreNull($avColorStops) Then
		$atColorStops = $tStyleGradient.ColorStops()
		If Not IsArray($atColorStops) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		ReDim $avNewColorStops[UBound($atColorStops)][2]

		For $i = 0 To UBound($atColorStops) - 1
			$avNewColorStops[$i][0] = $atColorStops[$i].StopOffset()
			$tStopColor = $atColorStops[$i].StopColor()
			If Not IsObj($tStopColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			$avNewColorStops[$i][1] = Int(BitShift(($tStopColor.Red() * 255), -16) + BitShift(($tStopColor.Green() * 255), -8) + ($tStopColor.Blue() * 255)) ; RGB to Long
			Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
		Next

		Return SetError($__LO_STATUS_SUCCESS, UBound($avNewColorStops), $avNewColorStops)
	EndIf

	If Not IsArray($avColorStops) Or (UBound($avColorStops, $__UBOUND_COLUMNS) <> 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (UBound($avColorStops) < 2) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	ReDim $atColorStops[UBound($avColorStops)]

	For $i = 0 To UBound($avColorStops) - 1
		$tColorStop = __LO_CreateStruct("com.sun.star.awt.ColorStop")
		If Not IsObj($tColorStop) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$tStopColor = $tColorStop.StopColor()
		If Not IsObj($tStopColor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)
		If Not __LO_NumIsBetween($avColorStops[$i][0], 0, 1.0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, $i)

		$tColorStop.StopOffset = $avColorStops[$i][0]

		If Not __LO_IntIsBetween($avColorStops[$i][1], $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, $i)

		$tStopColor.Red = (BitAND(BitShift($avColorStops[$i][1], 16), 0xff) / 255)
		$tStopColor.Green = (BitAND(BitShift($avColorStops[$i][1], 8), 0xff) / 255)
		$tStopColor.Blue = (BitAND($avColorStops[$i][1], 0xff) / 255)

		$tColorStop.StopColor = $tStopColor

		$atColorStops[$i] = $tColorStop

		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	$tStyleGradient.ColorStops = $atColorStops
	$oCell.FillGradient = $tStyleGradient

	$iError = (UBound($avColorStops) = UBound($oCell.FillGradient.ColorStops())) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_TableCellBackGradientMulticolor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellBorderColor
; Description ...: Set or retrieve the Cell Border Line Color. L.O. 3.4+.
; Syntax ........: _LOImpress_TableCellBorderColor(ByRef $oCell[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $iTop                - [optional] (0-16777215) Default is Null. The Top Border Line Color of the Cell, as a RGB Color Integer. A custom value or one of the constants $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iBottom             - [optional] (0-16777215) Default is Null. The Bottom Border Line Color of the Cell, as a RGB Color Integer. A custom value or one of the constants $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iLeft               - [optional] (0-16777215) Default is Null. The Left Border Line Color of the Cell, as a RGB Color Integer. A custom value or one of the constants $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iRight              - [optional] (0-16777215) Default is Null. The Right Border Line Color of the Cell, as a RGB Color Integer. A custom value or one of the constants $LO_COLOR_* as defined in LibreOffice_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  @Error: 1, @Extended: 2 = $iTop not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 3 = $iBottom not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 4 = $iLeft not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 5 = $iRight not an Integer, less than 0 or greater than 16777215.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error Creating "com.sun.star.table.BorderLine2" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  @Error: 3, @Extended: 2 = Cannot set Top Border Style/Color when Top Border width not set.
;                  @Error: 3, @Extended: 3 = Cannot set Bottom Border style/Color when Bottom Border width not set.
;                  @Error: 3, @Extended: 4 = Cannot set Left Border style/Color when Left Border width not set.
;                  @Error: 3, @Extended: 5 = Cannot set Right Border style/Color when Right Border width not set.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 3.4.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOImpress_TableCellBorderWidth, _LOImpress_TableCellBorderStyle, _LOImpress_TableCellBorderPadding, _LOImpress_TableBorderColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellBorderColor(ByRef $oCell, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LO_COLOR_BLACK, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$vReturn = __LOImpress_TableCellBorder($oCell, False, False, True, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableCellBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellBorderPadding
; Description ...: Set or retrieve the Cell Border Padding settings.
; Syntax ........: _LOImpress_TableCellBorderPadding(ByRef $oCell[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $iTop                - [optional] Default is Null. The Top Distance between the Border and Cell text in Hundredths of a Millimeter (HMM).
;                  $iBottom             - [optional] Default is Null. The Bottom Distance between the Border and Cell text in Hundredths of a Millimeter (HMM).
;                  $iLeft               - [optional] Default is Null. The Left Distance between the Border and Cell text in Hundredths of a Millimeter (HMM).
;                  $iRight              - [optional] Default is Null. The Right Distance between the Border and Cell text in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell an Object.
;                  @Error: 1, @Extended: 2 = $iTop not an Integer.
;                  @Error: 1, @Extended: 3 = $iBottom not an Integer.
;                  @Error: 1, @Extended: 4 = $Left not an Integer.
;                  @Error: 1, @Extended: 5 = $iRight not an Integer.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop.
;                  |                               2 = Error setting $iBottom.
;                  |                               4 = Error setting $iLeft.
;                  |                               8 = Error setting $iRight.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LO_UnitConvert, _LOImpress_TableCellBorderColor, _LOImpress_TableCellBorderStyle, _LOImpress_TableCellBorderWidth, _LOImpress_TableBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellBorderPadding(ByRef $oCell, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiBPadding[4]

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then
		__LO_ArrayFill($aiBPadding, $oCell.TextUpperDistance(), $oCell.TextLowerDistance(), $oCell.TextLeftDistance(), $oCell.TextRightDistance())

		Return SetError($__LO_STATUS_SUCCESS, 1, $aiBPadding)
	EndIf

	If ($iTop <> Null) Then
		If Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oCell.TextUpperDistance = $iTop
		$iError = (__LO_IntIsBetween($oCell.TextUpperDistance(), $iTop - 1, $iTop + 1)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iBottom <> Null) Then
		If Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oCell.TextLowerDistance = $iBottom
		$iError = (__LO_IntIsBetween($oCell.TextLowerDistance(), $iBottom - 1, $iBottom + 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iLeft <> Null) Then
		If Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oCell.TextLeftDistance = $iLeft
		$iError = (__LO_IntIsBetween($oCell.TextLeftDistance(), $iLeft - 1, $iLeft + 1)) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iRight <> Null) Then
		If Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

		$oCell.TextRightDistance = $iRight
		$iError = (__LO_IntIsBetween($oCell.TextRightDistance(), $iRight - 1, $iRight + 1)) ? ($iError) : (BitOR($iError, 8))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_TableCellBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellBorderStyle
; Description ...: Set or Retrieve the Cell Border Line style. L.O. 3.4+.
; Syntax ........: _LOImpress_TableCellBorderStyle(ByRef $oCell[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $iTop                - [optional] (0x7FFF,0-17) Default is Null. The Top Border Line Style of the Cell. See Constants, $LOI_SHAPE_BORDER_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iBottom             - [optional] (0x7FFF,0-17) Default is Null. The Bottom Border Line Style of the Cell. See Constants, $LOI_SHAPE_BORDER_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iLeft               - [optional] (0x7FFF,0-17) Default is Null. The Left Border Line Style of the Cell. See Constants, $LOI_SHAPE_BORDER_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iRight              - [optional] (0x7FFF,0-17) Default is Null. The Right Border Line Style of the Cell. See Constants, $LOI_SHAPE_BORDER_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  @Error: 1, @Extended: 2 = $iTop not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See Constants, $LOI_SHAPE_BORDER_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iBottom not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See Constants, $LOI_SHAPE_BORDER_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 4 = $iLeft not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See Constants, $LOI_SHAPE_BORDER_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 5 = $iRight not an Integer, less than 0 or greater than 17, but not equal to 0x7FFF. See Constants, $LOI_SHAPE_BORDER_STYLE_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error Creating "com.sun.star.table.BorderLine2" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  @Error: 3, @Extended: 2 = Cannot set Top Border Style/Color when Top Border width not set.
;                  @Error: 3, @Extended: 3 = Cannot set Bottom Border style/Color when Bottom Border width not set.
;                  @Error: 3, @Extended: 4 = Cannot set Left Border style/Color when Left Border width not set.
;                  @Error: 3, @Extended: 5 = Cannot set Right Border style/Color when Right Border width not set.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 3.4.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOImpress_TableCellBorderWidth, _LOImpress_TableCellBorderColor, _LOImpress_TableCellBorderPadding, _LOImpress_TableBorderStyle
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellBorderStyle(ByRef $oCell, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, $LOI_SHAPE_BORDER_STYLE_SOLID, $LOI_SHAPE_BORDER_STYLE_DASH_DOT_DOT, "", $LOI_SHAPE_BORDER_STYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, $LOI_SHAPE_BORDER_STYLE_SOLID, $LOI_SHAPE_BORDER_STYLE_DASH_DOT_DOT, "", $LOI_SHAPE_BORDER_STYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, $LOI_SHAPE_BORDER_STYLE_SOLID, $LOI_SHAPE_BORDER_STYLE_DASH_DOT_DOT, "", $LOI_SHAPE_BORDER_STYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, $LOI_SHAPE_BORDER_STYLE_SOLID, $LOI_SHAPE_BORDER_STYLE_DASH_DOT_DOT, "", $LOI_SHAPE_BORDER_STYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$vReturn = __LOImpress_TableCellBorder($oCell, False, True, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableCellBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellBorderWidth
; Description ...: Set or Retrieve the Cell Border Line Width. L.O. 3.4+.
; Syntax ........: _LOImpress_TableCellBorderWidth(ByRef $oCell[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $iTop                - [optional] Default is Null. The Top Border Line width of the Cell in Hundredths of a Millimeter (HMM). Can be a custom value or one of the constants, $LOI_SHAPE_BORDER_WIDTH_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iBottom             - [optional] Default is Null. The Bottom Border Line Width of the Cell in Hundredths of a Millimeter (HMM). Can be a custom value or one of the constants, $LOI_SHAPE_BORDER_WIDTH_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iLeft               - [optional] Default is Null. The Left Border Line width of the Cell in Hundredths of a Millimeter (HMM). Can be a custom value or one of the constants, $LOI_SHAPE_BORDER_WIDTH_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iRight              - [optional] Default is Null. The Right Border Line Width of the Cell in Hundredths of a Millimeter (HMM). Can be a custom value or one of the constants, $LOI_SHAPE_BORDER_WIDTH_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  @Error: 1, @Extended: 2 = $iTop not an Integer, or less than 0.
;                  @Error: 1, @Extended: 3 = $iBottom not an Integer, or less than 0.
;                  @Error: 1, @Extended: 4 = $iLeft not an Integer, or less than 0.
;                  @Error: 1, @Extended: 5 = $iRight not an Integer, or less than 0.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error Creating "com.sun.star.table.BorderLine2" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Internal command error. More than one parameter called with True. UDF Must be fixed.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iTop
;                  |                               2 = Error setting $iBottom
;                  |                               4 = Error setting $iLeft
;                  |                               8 = Error setting $iRight
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 3.4.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To "Turn Off" Borders, set them to 0
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LO_UnitConvert, _LOImpress_TableCellBorderStyle, _LOImpress_TableCellBorderColor, _LOImpress_TableCellBorderPadding, _LOImpress_TableBorderWidth
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellBorderWidth(ByRef $oCell, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($iTop <> Null) And Not __LO_IntIsBetween($iTop, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iBottom <> Null) And Not __LO_IntIsBetween($iBottom, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iLeft <> Null) And Not __LO_IntIsBetween($iLeft, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iRight <> Null) And Not __LO_IntIsBetween($iRight, 0) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$vReturn = __LOImpress_TableCellBorder($oCell, True, False, False, $iTop, $iBottom, $iLeft, $iRight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableCellBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellCharEffect
; Description ...: Set or Retrieve the Font Effect settings for a Table cell.
; Syntax ........: _LOImpress_TableCellCharEffect(ByRef $oCell[, $iCase = Null[, $iRelief = Null[, $bOutline = Null[, $bShadow = Null]]]])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $iCase               - [optional] (0-4) Default is Null. The Character Case Style. See Constants, $LOI_CHAR_CASEMAP_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iRelief             - [optional] (0-2) Default is Null. The Character Relief style. See Constants, $LOI_CHAR_RELIEF_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bOutline            - [optional] Default is Null. If True, the characters have an outline around the outside.
;                  $bShadow             - [optional] Default is Null. If True, the characters have a shadow.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  @Error: 1, @Extended: 2 = $iCase not an Integer, less than 0 or greater than 4. See Constants, $LOI_CHAR_CASEMAP_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iRelief not an Integer, less than 0 or greater than 2. See Constants, $LOI_CHAR_RELIEF_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 4 = $bOutline not a Boolean.
;                  @Error: 1, @Extended: 5 = $bShadow not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iCase
;                  |                               2 = Error setting $iRelief
;                  |                               4 = Error setting $bOutline
;                  |                               8 = Error setting $bShadow
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOImpress_TableCellCharOverLine, _LOImpress_TableCellCharStrikeOut, _LOImpress_TableCellCharUnderLine, _LOImpress_TableCharEffect
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellCharEffect(ByRef $oCell, $iCase = Null, $iRelief = Null, $bOutline = Null, $bShadow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharEffect($oCell, $iCase, $iRelief, $bOutline, $bShadow)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableCellCharEffect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellCharFont
; Description ...: Set and Retrieve the Font Settings for a Table cell.
; Syntax ........: _LOImpress_TableCellCharFont(ByRef $oCell[, $sFontName = Null[, $nFontSize = Null[, $iPosture = Null[, $iWeight = Null]]]])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $sFontName           - [optional] Default is Null. The Font Name to use.
;                  $nFontSize           - [optional] Default is Null. The new Font size.
;                  $iPosture            - [optional] (0-5) Default is Null. The Font Italic setting. See Constants, $LOI_CHAR_POSTURE_* as defined in LibreOfficeImpress_Constants.au3. Also see remarks.
;                  $iWeight             - [optional] (0, 50-200) Default is Null. The Font Bold settings see Constants, $LOI_CHAR_WEIGHT_* as defined in LibreOfficeImpress_Constants.au3. Also see remarks.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  @Error: 1, @Extended: 2 = $sFontName not a String.
;                  @Error: 1, @Extended: 3 = Font called in $sFontName not available.
;                  @Error: 1, @Extended: 4 = $nFontSize not a number.
;                  @Error: 1, @Extended: 5 = $iPosture not an Integer, less than 0 or greater than 5. See Constants, $LOI_CHAR_POSTURE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 6 = $iWeight not an Integer, less than 50 but not equal to 0, or greater than 200. See Constants, $LOI_CHAR_WEIGHT_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sFontName
;                  |                               2 = Error setting $nFontSize
;                  |                               4 = Error setting $iPosture
;                  |                               8 = Error setting $iWeight
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Not every font accepts Bold and Italic settings, and not all settings for bold and Italic are accepted, such as oblique, ultra Bold etc.
;                  LibreOffice accepts only the predefined weight values, any other values are changed automatically to an acceptable value, which could trigger a settings error.
; Related .......: _LOImpress_TableCellCharFontColor, _LOImpress_TableCharFont, _LOImpress_FontsGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellCharFont(ByRef $oCell, $sFontName = Null, $nFontSize = Null, $iPosture = Null, $iWeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharFont($oCell, $sFontName, $nFontSize, $iPosture, $iWeight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableCellCharFont

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellCharFontColor
; Description ...: Set or retrieve the font color, transparency and highlighting values for a Table cell.
; Syntax ........: _LOImpress_TableCellCharFontColor(ByRef $oCell[, $iFontColor = Null[, $iTransparency = Null[, $iHighlight = Null]]])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $iFontColor          - [optional] (-1-16777215) Default is Null. The font Color value, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for Auto color.
;                  $iTransparency       - [optional] (0-100) Default is Null. Transparency percentage. 0 is visible, 100 is invisible. Available for LibreOffice 7.0 and up.
;                  $iHighlight          - [optional] (-1-16777215) Default is Null. The highlight Color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for No color.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters. If The current LibreOffice version is below 7.0 the $iTransparency parameter will return a Null value.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  @Error: 1, @Extended: 2 = $iFontColor not an Integer, less than -1 or greater than 16777215.
;                  @Error: 1, @Extended: 3 = $iTransparency not an Integer, less than 0 or greater than 100%.
;                  @Error: 1, @Extended: 4 = $iHighlight not an Integer, less than -1 or greater than 16777215.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve old Transparency value.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $FontColor
;                  |                               2 = Error setting $iTransparency.
;                  |                               4 = Error setting $iHighlight
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 7.0.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOImpress_TableCellCharFont, _LOImpress_TableCharFontColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellCharFontColor(ByRef $oCell, $iFontColor = Null, $iTransparency = Null, $iHighlight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharFontColor($oCell, $iFontColor, $iTransparency, $iHighlight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableCellCharFontColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellCharOverLine
; Description ...: Set and retrieve the OverLine settings for a Table cell.
; Syntax ........: _LOImpress_TableCellCharOverLine(ByRef $oCell[, $iOverLineStyle = Null[, $iOLColor = Null[, $bWordOnly = Null]]])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $iOverLineStyle      - [optional] (0-18) Default is Null. The style of the Overline line, see constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  $iOLColor            - [optional] (-1-16777215) Default is Null. The Overline color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for automatic color mode.
;                  $bWordOnly           - [optional] Default is Null. If True, white spaces are not Overlined.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  @Error: 1, @Extended: 2 = $iOverLineStyle not an Integer, less than 0 or greater than 18. See constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  @Error: 1, @Extended: 3 = $iOLColor not an Integer, less than -1 or greater than 16777215.
;                  @Error: 1, @Extended: 4 = $bWordOnly not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iOverLineStyle
;                  |                               2 = Error setting $iOLColor
;                  |                               4 = Error setting $bWordOnly
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Overline line style uses the same constants as underline style.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOImpress_TableCellCharEffect, _LOImpress_TableCellCharStrikeOut, _LOImpress_TableCellCharUnderLine, _LOImpress_TableCharOverLine
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellCharOverLine(ByRef $oCell, $iOverLineStyle = Null, $iOLColor = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharOverLine($oCell, $iOverLineStyle, $iOLColor, $bWordOnly)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableCellCharOverLine

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellCharPosition
; Description ...: Set and retrieve settings related to Sub/Super Script and relative size for a Table cell.
; Syntax ........: _LOImpress_TableCellCharPosition(ByRef $oCell[, $iSuperScript = Null[, $iSubScript = Null[, $iRelativeSize = Null]]])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $iSuperScript        - [optional] (-1-100) Default is Null. The Superscript percentage value. Call with -1 for Automatic SuperScript. See Remarks.
;                  $iSubScript          - [optional] (-1-100) Default is Null. Subscript percentage value. Call with -1 for Automatic SubScript. See Remarks.
;                  $iRelativeSize       - [optional] (1-100) Default is Null. The size percentage relative to current font size.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  @Error: 1, @Extended: 2 = $oCell does not support Character properties.
;                  @Error: 1, @Extended: 3 = $iSuperScript not an Integer, less than 0 or greater than 100, but not 14000.
;                  @Error: 1, @Extended: 4 = $iSubScript not an Integer, less than -100 or greater than 100, but not 14000 or -14000.
;                  @Error: 1, @Extended: 5 = $iRelativeSize not an Integer, less than 1 or greater than 100.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iSuperScript
;                  |                               2 = Error setting $iSubScript
;                  |                               4 = Error setting $iRelativeSize.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Set either $iSubScript or $iSuperScript to 0 to return it to Normal setting.
;                  The way LibreOffice is set up Super/Subscript are set in the same setting, Superscript is a positive number from 1 to 100 (percentage), Subscript is a negative number set to -1 to -100 percentage. For the user's convenience this function automatically converts the positive numbers to negative, and back when setting or retrieving subscript values.
;                  Automatic Superscript has an Integer value of 14000, Auto Subscript has a Integer value of -14000. Being that there is no settable setting of Automatic Super/Sub Script, it has been chosen to use -1 to indicate an automatic Sub/SuperScript value.
;                  If you set both $iSuperScript and $iSubScript to -1 (Automatic), or both $iSuperScript and $iSubScript to any value, Subscript will be the result, as it is the last in the function to be set, and thus will overwrite any Superscript values.
; Related .......: _LOImpress_TableCellParAlignment, _LOImpress_TableCellParIndent, _LOImpress_TableCellParSpacing
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellCharPosition(ByRef $oCell, $iSuperScript = Null, $iSubScript = Null, $iRelativeSize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharPosition($oCell, $iSuperScript, $iSubScript, $iRelativeSize)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableCellCharPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellCharScaling
; Description ...: Set or retrieve the character Scale settings for a Table cell.
; Syntax ........: _LOImpress_TableCellCharScaling(ByRef $oCell[, $iScaleWidth = Null])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $iScaleWidth         - [optional] (1-100) Default is Null. The percentage to horizontally stretch or compress the text. 100 is normal sizing.
; Return values .: Success: 1 or Integer.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Integer = Success. All optional parameters were called with Null, returning current Scale Width value as an Integer.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  @Error: 1, @Extended: 2 = $iScaleWidth not an Integer or less than 1% or greater than 100%.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Scale width.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iScaleWidth
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Fit to line seems to be unavailable in the API, and does not seem to work in LibreOffice anyway.
; Related .......: _LOImpress_TableCellCharSpacing
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellCharScaling(ByRef $oCell, $iScaleWidth = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharScaling($oCell, $iScaleWidth)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableCellCharScaling

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellCharSpacing
; Description ...: Set and retrieve the spacing between characters (Kerning) for a Table cell.
; Syntax ........: _LOImpress_TableCellCharSpacing(ByRef $oCell[, $bAutoKerning = Null[, $nKerning = Null]])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $bAutoKerning        - [optional] Default is Null. If True, applies a spacing in between certain pairs of characters.
;                  $nKerning            - [optional] (-928.8-928.8) Default is Null. The kerning value of the characters. See Remarks. Values are in Printer's Points as set in the LibreOffice UI.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  @Error: 1, @Extended: 2 = $bAutoKerning not a Boolean.
;                  @Error: 1, @Extended: 3 = $nKerning not a number, less than -928.8 or greater than 928.8 Points.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bAutoKerning
;                  |                               2 = Error setting $nKerning.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  When setting Kerning values in LibreOffice, the measurement is listed in Pt (Printer's Points) in the User Display, however the internal setting is measured in Hundredths of a Millimeter (HMM). They will be automatically converted from Points to Hundredths of a Millimeter and back for retrieval of settings.
;                  The acceptable values are from -2 Pt to 928.8 Pt. The values can be directly converted easily, however, for an unknown reason to myself, LibreOffice begins counting backwards and in negative Hundredths of a Millimeter internally from 928.9 up to 1000 Pt (Max setting).
;                  For example, 928.8Pt is the last correct value, which equals 32766 Hundredths of a Millimeter (HMM), after this LibreOffice reports the following: 928.9 Pt = -32766 HMM; 929 Pt = -32763 HMM; 929.1 = -32759; 1000 pt = -30258. Attempting to set Libre's kerning value to anything over 32768 HMM causes a COM exception, and attempting to set the kerning to any of these negative numbers sets the User viewable kerning value to -2.0 Pt. For these reasons the max settable kerning is -2.0 Pt to 928.8 Pt.
; Related .......: _LOImpress_TableCellCharScaling, _LOImpress_TableCellParSpacing
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellCharSpacing(ByRef $oCell, $bAutoKerning = Null, $nKerning = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharSpacing($oCell, $bAutoKerning, $nKerning)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableCellCharSpacing

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellCharStrikeOut
; Description ...: Set or Retrieve the Strikeout settings for a Table cell.
; Syntax ........: _LOImpress_TableCellCharStrikeOut(ByRef $oCell[, $iStrikeLineStyle = Null[, $bWordOnly = Null]])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $iStrikeLineStyle    - [optional] (0-6) Default is Null. The Strikeout Line Style, see constants, $LOI_CHAR_STRIKEOUT_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bWordOnly           - [optional] Default is Null. If True, strike out is applied to words only, skipping whitespaces.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  @Error: 1, @Extended: 2 = $iStrikeLineStyle not an Integer, less than 0 or greater than 6. See constants, $LOI_CHAR_STRIKEOUT_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $bWordOnly not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iStrikeLineStyle
;                  |                               2 = Error setting $bWordOnly
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOImpress_TableCellCharEffect, _LOImpress_TableCellCharOverLine, _LOImpress_TableCellCharUnderLine, _LOImpress_TableCharStrikeOut
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellCharStrikeOut(ByRef $oCell, $iStrikeLineStyle = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharStrikeOut($oCell, $iStrikeLineStyle, $bWordOnly)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableCellCharStrikeOut

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellCharUnderLine
; Description ...: Set or retrieve Underline settings for a Table cell.
; Syntax ........: _LOImpress_TableCellCharUnderLine(ByRef $oCell[, $iUnderLineStyle = Null[, $iULColor = Null[, $bWordOnly = Null]]])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $iUnderLineStyle     - [optional] (0-18) Default is Null. The Underline line style, see constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iULColor            - [optional] (-1-16777215) Default is Null. The underline color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for automatic color mode.
;                  $bWordOnly           - [optional] Default is Null. If True, white spaces are not underlined.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell an Object.
;                  @Error: 1, @Extended: 2 = $iUnderLineStyle not an Integer, less than 0 or greater than 18. See constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iULColor not an Integer, less than -1 or greater than 16777215.
;                  @Error: 1, @Extended: 4 = $bWordOnly not a Boolean.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iUnderLineStyle
;                  |                               2 = Error setting $iULColor
;                  |                               4 = Error setting $bWordOnly
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOImpress_TableCellCharEffect, _LOImpress_TableCellCharOverLine, _LOImpress_TableCellCharStrikeOut, _LOImpress_TableCharUnderLine
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellCharUnderLine(ByRef $oCell, $iUnderLineStyle = Null, $iULColor = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_CharUnderLine($oCell, $iUnderLineStyle, $iULColor, $bWordOnly)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableCellCharUnderLine

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellCreateTextCursor
; Description ...: Create a Text Cursor in a particular cell for inserting text etc.
; Syntax ........: _LOImpress_TableCellCreateTextCursor(ByRef $oCell)
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning a Text Cursor Object located in the specified Cell.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a TextCursor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_CursorInsertString, _LOImpress_TableCellString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellCreateTextCursor(ByRef $oCell)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextCursor

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oTextCursor = $oCell.Text.createTextCursor()
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextCursor)
EndFunc   ;==>_LOImpress_TableCellCreateTextCursor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellGetObjByPosition
; Description ...: Retrieve a Cell's object by position.
; Syntax ........: _LOImpress_TableCellGetObjByPosition(ByRef $oTable, $iColumn, $iRow)
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
;                  $iColumn             - The column index of the desired cell. 0 based.
;                  $iRow                - The row index of the desired cell. 0 based.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning the Cell object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  @Error: 1, @Extended: 2 = $iColumn not an Integer, less than 0 or greater than number of Columns contained in the table.
;                  @Error: 1, @Extended: 3 = $iRow not an Integer, less than 0 or greater than number of Rows contained in the table.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Cell Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function can fail with complex Tables. Complex tables are tables that contain cells that have been split or joined.
;                  Rows and Columns in a Table are 0 based, meaning they start their count at 0. The first cell is column 0 row 0.
; Related .......: _LOImpress_TableColumnGetCount, _LOImpress_TableRowGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellGetObjByPosition(ByRef $oTable, $iColumn, $iRow)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCell

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iColumn, 0, ($oTable.Model.ColumnCount() - 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iRow, 0, ($oTable.Model.RowCount() - 1)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oCell = $oTable.Model.getCellByPosition($iColumn, $iRow)
	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oCell)
EndFunc   ;==>_LOImpress_TableCellGetObjByPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellParAlignment
; Description ...: Set and Retrieve Paragraph Alignment settings for a Table cell.
; Syntax ........: _LOImpress_TableCellParAlignment(ByRef $oCell[, $iHorAlign = Null[, $iLastLineAlign = Null[, $iTxtDirection = Null]]])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $iHorAlign           - [optional] (0-3) Default is Null. The Horizontal alignment of the paragraph. See Constants, $LOI_PAR_ALIGN_HOR_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  $iLastLineAlign      - [optional] (0-3) Default is Null. Specify the alignment for the last line in the paragraph. See Constants, $LOI_PAR_LAST_LINE_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  $iTxtDirection       - [optional] (0-5) Default is Null. The Text Writing Direction. See Constants, $LOI_PAR_TXT_DIR_* as defined in LibreOfficeImpress_Constants.au3. [LibreOffice Default is 4]
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  @Error: 1, @Extended: 2 = $iHorAlign not an Integer, less than 0 or greater than 3. See Constants, $LOI_PAR_ALIGN_HOR_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iLastLineAlign not an Integer, less than 0 or greater than 3. See Constants, $LOI_PAR_LAST_LINE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 4 = $iTxtDirection not an Integer, less than 0 or greater than 5. See Constants, $LOI_PAR_TXT_DIR_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iHorAlign
;                  |                               2 = Error setting $iLastLineALign
;                  |                               4 = Error setting $iTxtDirection
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iHorAlign must be set to $LOI_PAR_ALIGN_HOR_JUSTIFIED(2) before you can set $iLastLineAlign.
;                  $iTxtDirection constants 2,3, and 5 may not be available depending on your language settings.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Expand single word, Snap to grid, and Vertical align (Text-To-Text), seem to be unavailable in the API, and do not seem to work in LibreOffice.
; Related .......: _LOImpress_TableCellParIndent, _LOImpress_TableCellParSpacing, _LOImpress_TableCellCharPosition
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellParAlignment(ByRef $oCell, $iHorAlign = Null, $iLastLineAlign = Null, $iTxtDirection = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParAlignment($oCell, $iHorAlign, $iLastLineAlign, $iTxtDirection)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableCellParAlignment

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellParIndent
; Description ...: Set or Retrieve Paragraph Indent settings for a Table cell.
; Syntax ........: _LOImpress_TableCellParIndent(ByRef $oCell[, $iBeforeTxt = Null[, $iAfterTxt = Null[, $iFirstLine = Null]]])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $iBeforeTxt          - [optional] (0-1162202) Default is Null. The amount of space that you want to indent the paragraph from the page margin. Set in Hundredths of a Millimeter (HMM).
;                  $iAfterTxt           - [optional] (0-1162202) Default is Null. The amount of space that you want to indent the paragraph from the page margin. Set in Hundredths of a Millimeter (HMM)
;                  $iFirstLine          - [optional] (0-1162202) Default is Null. Indentation distance of the first line of a paragraph. Set in Hundredths of a Millimeter (HMM).
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  @Error: 1, @Extended: 2 = $iBeforeText not an Integer, less than 0 or greater than 1162202.
;                  @Error: 1, @Extended: 3 = $iAfterText not an Integer, less than 0 or greater than 1162202.
;                  @Error: 1, @Extended: 4 = $iFirstLine not an Integer, less than 0 or greater than 1162202.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iBeforeTxt
;                  |                               2 = Error setting $iAfterTxt
;                  |                               4 = Error setting $iFirstLine
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Auto indent first line does not seem to work in LibreOffice, and seems to be not available in the API.
; Related .......: _LO_UnitConvert, _LOImpress_TableCellParAlignment, _LOImpress_TableCellParSpacing, _LOImpress_TableCellCharPosition
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellParIndent(ByRef $oCell, $iBeforeTxt = Null, $iAfterTxt = Null, $iFirstLine = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParIndent($oCell, $iBeforeTxt, $iAfterTxt, $iFirstLine)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableCellParIndent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellParSpacing
; Description ...: Set or Retrieve Line Spacing settings for a Table cell.
; Syntax ........: _LOImpress_TableCellParSpacing(ByRef $oCell[, $iAbovePar = Null[, $iBelowPar = Null[, $iLineSpcMode = Null[, $iLineSpcHeight = Null]]]])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $iAbovePar           - [optional] (0-100000) Default is Null. The Space above a paragraph, in Hundredths of a Millimeter (HMM).
;                  $iBelowPar           - [optional] (0-100000) Default is Null. The Space Below a paragraph, in Hundredths of a Millimeter (HMM).
;                  $iLineSpcMode        - [optional] (0-3) Default is Null. The line spacing type of the paragraph. See Constants, $LOI_PAR_LINE_SPC_MODE_* as defined in LibreOfficeImpress_Constants.au3, also notice min and max values for each.
;                  $iLineSpcHeight      - [optional] Default is Null. This value specifies the height in regard to Mode. See Remarks.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  @Error: 1, @Extended: 2 = $iAbovePar not an Integer, less than 0 or greater than 100000.
;                  @Error: 1, @Extended: 3 = $iBelowPar not an Integer, less than 0 or greater than 100000.
;                  @Error: 1, @Extended: 4 = $iLineSpcMode not an Integer, less than 0 or greater than 3. See Constants, $LOI_PAR_LINE_SPC_MODE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 5 = $iLineSpcHeight not an Integer.
;                  @Error: 1, @Extended: 6 = $iLineSpcMode set to 0(Proportional) and $iLineSpcHeight less than 6(%) or greater than 65535(%).
;                  @Error: 1, @Extended: 7 = $iLineSpcMode set to 1 or 2(Minimum, or Leading) and $iLineSpcHeight less than 0 or greater than 100000.
;                  @Error: 1, @Extended: 8 = $iLineSpcMode set to 3(Fixed) and $iLineSpcHeight less than 51 or greater than 100000.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving ParaLineSpacing Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iAbovePar
;                  |                               2 = Error setting $iBelowPar
;                  |                               4 = Error setting $iLineSpcMode
;                  |                               8 = Error setting $iLineSpcHeight
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The settings in LibreOffice, (Single, 1.15, 1.5, Double), Use the Proportional mode, and are just varying percentages. e.g Single = 100, 1.15 = 115%, 1.5 = 150%, Double = 200%.
;                  $iLineSpcHeight depends on the $iLineSpcMode used, see constants for accepted Input values.
;                  $iAbovePar, $iBelowPar, $iLineSpcHeight may change +/- a Hundredth of a Millimeter (HMM) once set.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  The "Do not add space between paragraphs as the same style" setting seems to be not available to set or retrieve in the API, and seems to do nothing in LibreOffice anyway.
; Related .......: _LO_UnitConvert, _LOImpress_TableCellParAlignment, _LOImpress_TableCellParIndent, _LOImpress_TableCellCharSpacing
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellParSpacing(ByRef $oCell, $iAbovePar = Null, $iBelowPar = Null, $iLineSpcMode = Null, $iLineSpcHeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParSpacing($oCell, $iAbovePar, $iBelowPar, $iLineSpcMode, $iLineSpcHeight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableCellParSpacing

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellParTabStopCreate
; Description ...: Create a new TabStop for a Table cell.
; Syntax ........: _LOImpress_TableCellParTabStopCreate(ByRef $oCell, $iPosition[, $iAlignment = Null[, $iDecChar = Null[, $iFillChar = Null]]])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $iPosition           - The TabStop position to set the new TabStop to. Set in Hundredths of a Millimeter (HMM). See Remarks.
;                  $iAlignment          - [optional] (0-4) Default is Null. The position of where the end of a Tab is aligned to compared to the text. See Constants, $LOI_PAR_TAB_ALIGN_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iDecChar            - [optional] Default is Null. Enter a character(in Asc Value(See AutoIt Asc Function)) that you want the decimal tab to use as a decimal separator. Can only be set if $iAlignment is set to $LOI_PAR_TAB_ALIGN_DECIMAL.
;                  $iFillChar           - [optional] Default is Null. The Asc (see AutoIt function) value of any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
; Return values .: Success: Integer.
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Settings were successfully set. New TabStop position is returned.
;                  Failure: 0 or Integer and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  @Error: 1, @Extended: 2 = $iPosition not an Integer.
;                  @Error: 1, @Extended: 3 = Tab Stop position called in $iPosition already exists in this Paragraph.
;                  @Error: 1, @Extended: 4 = $iAlignment not an Integer, less than 0 or greater than 4. See Constants, $LOI_PAR_TAB_ALIGN_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 5 = $iDecChar not an Integer.
;                  @Error: 1, @Extended: 6 = $iFillChar not an Integer.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error creating "com.sun.star.style.TabStop" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving ParaTabStops Array Object.
;                  @Error: 3, @Extended: 2 = Error retrieving list of TabStop Positions.
;                  @Error: 3, @Extended: 3 = Failed to identify the new Tabstop once inserted.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iPosition
;                  |                               2 = Error setting $iAlignment
;                  |                               4 = Error setting $iDecChar
;                  |                               8 = Error setting $iFillChar
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iPosition once set can vary +/- a Hundredth of a Millimeter (HMM). To ensure you can identify the tabstop to modify it again, This function returns the new TabStop position.
;                  Since $iPosition can fluctuate +/- a Hundredth of a Millimeter (HMM) when it is inserted into LibreOffice, it is possible to accidentally overwrite an already existing TabStop.
;                  $iFillChar, Libre's Default value, "None" is in reality a space character which is Asc value 32. The other values offered by Libre are: Period (ASC 46), Dash (ASC 45) and Underscore (ASC 95). You can also enter a custom ASC value. See ASC AutoIt Func. and "ASCII Character Codes" in the AutoIt help file.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  $iNewTabStop position is still returned as even though some settings weren't successfully set, the new TabStop was still created.
; Related .......: _LO_UnitConvert, _LOImpress_TableCellParTabStopDelete, _LOImpress_TableCellParTabStopMod, _LOImpress_TableCellParTabStopsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellParTabStopCreate(ByRef $oCell, $iPosition, $iAlignment = Null, $iDecChar = Null, $iFillChar = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParTabStopCreate($oCell, $iPosition, $iAlignment, $iDecChar, $iFillChar)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableCellParTabStopCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellParTabStopDelete
; Description ...: Delete a TabStop from a Table cell.
; Syntax ........: _LOImpress_TableCellParTabStopDelete(ByRef $oCell, $iTabStop)
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $iTabStop            - The Tab position of the TabStop to modify. See Remarks.
; Return values .: Success: Boolean.
;                  @Error: 0, @Extended: 0, Return: Boolean = Returning True if TabStop was successfully deleted, else False.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  @Error: 1, @Extended: 2 = $iTabStop not an Integer.
;                  @Error: 1, @Extended: 3 = $iTabStop not found.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving ParaTabStops Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iTabStop refers to the position, or essential the "length" of a TabStop from the edge of a page margin. This is the only reliable way to identify a Tabstop to be able to interact with it, as there can only be one of a certain length per paragraph.
; Related .......: _LOImpress_TableCellParTabStopCreate, _LOImpress_TableCellParTabStopsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellParTabStopDelete(ByRef $oCell, $iTabStop)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParTabStopDelete($oCell, $iTabStop)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableCellParTabStopDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellParTabStopMod
; Description ...: Modify or retrieve the properties of an existing TabStop in a Table cell.
; Syntax ........: _LOImpress_TableCellParTabStopMod(ByRef $oCell, $iTabStop[, $iPosition = Null[, $iAlignment = Null[, $iDecChar = Null[, $iFillChar = Null]]]])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $iTabStop            - The Tab position of the TabStop to modify. See Remarks.
;                  $iPosition           - [optional] Default is Null. The New position to set the input position to. Set in Hundredths of a Millimeter (HMM). See Remarks.
;                  $iAlignment          - [optional] (0-4) Default is Null. The position of where the end of a Tab is aligned to compared to the text. See Constants, $LOI_PAR_TAB_ALIGN_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iDecChar            - [optional] Default is Null. Enter a character(in Asc Value(See AutoIt Asc Function)) that you want the decimal tab to use as a decimal separator. Can only be set if $iAlignment is set to $LOI_PAR_TAB_ALIGN_DECIMAL.
;                  $iFillChar           - [optional] Default is Null. The Asc (see AutoIt function) value of any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
; Return values .: Success: Integer or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  @Error: 0, @Extended: ?, Return: 2 = Success. Settings were successfully set. New TabStop position is returned in @Extended.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  @Error: 1, @Extended: 2 = $iTabStop not an Integer.
;                  @Error: 1, @Extended: 3 = TabStop called in $iTabStop not found.
;                  @Error: 1, @Extended: 4 = $iPosition not an Integer.
;                  @Error: 1, @Extended: 5 = $iAlignment not an Integer, less than 0 or greater than 4. See Constants, $LOI_PAR_TAB_ALIGN_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 6 = $iDecChar not an Integer.
;                  @Error: 1, @Extended: 7 = $iFillChar not an Integer.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving ParaTabStops Object.
;                  @Error: 3, @Extended: 2 = Error retrieving Requested TabStop Object.
;                  @Error: 3, @Extended: 3 = Paragraph already contains a TabStop at the length/Position specified in $iPosition.
;                  @Error: 3, @Extended: 4 = Error retrieving list of TabStop Positions.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iPosition
;                  |                               2 = Error setting $iAlignment
;                  |                               4 = Error setting $iDecChar
;                  |                               8 = Error setting $iFillChar
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iTabStop refers to the position, or essential the "length" of a TabStop from the edge of a page margin. This is the only reliable way to identify a Tabstop to be able to interact with it, as there can only be one of a certain length per Paragraph.
;                  $iPosition once set can vary +/- a Hundredth of a Millimeter (HMM). To ensure you can identify the tabstop to modify it again, This function returns the new TabStop position in @Extended when $iPosition is set, return value will be set to 2. See Return Values.
;                  Since $iPosition can fluctuate +/- a Hundredth of a Millimeter (HMM) when it is inserted into LibreOffice, it is possible to accidentally overwrite an already existing TabStop.
;                  $iFillChar, Libre's Default value, "None" is in reality a space character which is Asc value 32. The other values offered by Libre are: Period (ASC 46), Dash (ASC 45) and Underscore (ASC 95). You can also enter a custom ASC value. See ASC AutoIt Func. and "ASCII Character Codes" in the AutoIt help file.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LO_UnitConvert, _LOImpress_TableCellParTabStopCreate, _LOImpress_TableCellParTabStopsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellParTabStopMod(ByRef $oCell, $iTabStop, $iPosition = Null, $iAlignment = Null, $iDecChar = Null, $iFillChar = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParTabStopMod($oCell, $iTabStop, $iPosition, $iAlignment, $iDecChar, $iFillChar)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableCellParTabStopMod

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellParTabStopsGetList
; Description ...: Retrieve an array of TabStops available in a Table cell.
; Syntax ........: _LOImpress_TableCellParTabStopsGetList(ByRef $oCell)
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
; Return values .: Success: Array.
;                  @Error: 0, @Extended: ?, Return: Array = Success. An Array of TabStops. @Extended set to number of results.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving ParaTabStops Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_TableCellParTabStopCreate, _LOImpress_TableCellParTabStopDelete, _LOImpress_TableCellParTabStopMod
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellParTabStopsGetList(ByRef $oCell)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ParTabStopsGetList($oCell)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableCellParTabStopsGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCellString
; Description ...: Set or retrieve the current string of a cell.
; Syntax ........: _LOImpress_TableCellString(ByRef $oCell[, $sString = Null])
; Parameters ....: $oCell               - A Table Cell object returned by a previous _LOImpress_TableCellGetObjByPosition function.
;                  $sString             - [optional] Default is Null. The String of text to set the cell to.
; Return values .: Success: 1 or String.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: String = Success. All optional parameters were called with Null, returning current cell string.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCell not an Object.
;                  @Error: 1, @Extended: 2 = $sString not a String.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current cell string.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sString
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Setting the String will overwrite any existing data in the cell.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To prevent accidental and unwanted newlines, @CRLF is automatically replaced with @CR to match LibreOffice's newline style.
; Related .......: _LOImpress_TableCellCreateTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCellString(ByRef $oCell, $sString = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sCurString
	Local $iError = 0

	If Not IsObj($oCell) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($sString) Then
		$sCurString = $oCell.getString()
		If Not IsString($sCurString) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		Return SetError($__LO_STATUS_SUCCESS, 1, $sCurString)
	EndIf

	If Not IsString($sString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	; Exchange CRLF for CR to prevent errors.
	$sString = StringRegExpReplace($sString, @CRLF, @CR)

	$oCell.setString($sString)
	$iError = ($oCell.getString() = $sString) ? ($iError) : (BitOR($iError, 1))

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_TableCellString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCharEffect
; Description ...: Set or Retrieve the Font Effect settings for a Table.
; Syntax ........: _LOImpress_TableCharEffect(ByRef $oTable[, $iCase = Null[, $iRelief = Null[, $bOutline = Null[, $bShadow = Null]]]])
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
;                  $iCase               - [optional] (0-4) Default is Null. The Character Case Style. See Constants, $LOI_CHAR_CASEMAP_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iRelief             - [optional] (0-2) Default is Null. The Character Relief style. See Constants, $LOI_CHAR_RELIEF_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bOutline            - [optional] Default is Null. If True, the characters have an outline around the outside.
;                  $bShadow             - [optional] Default is Null. If True, the characters have a shadow.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  @Error: 1, @Extended: 2 = $iCase not an Integer, less than 0 or greater than 4. See Constants, $LOI_CHAR_CASEMAP_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iRelief not an Integer, less than 0 or greater than 2. See Constants, $LOI_CHAR_RELIEF_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 4 = $bOutline not a Boolean.
;                  @Error: 1, @Extended: 5 = $bShadow not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Cell Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iCase
;                  |                               2 = Error setting $iRelief
;                  |                               4 = Error setting $bOutline
;                  |                               8 = Error setting $bShadow
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Tables require that the properties be set individually for each Cell, therefore this function cycles through each cell and sets the value, and may be slower for large tables.
;                  When retrieving the current property values for a table, if all of the cells in the Table do not have the same value, Null is returned for that property value.
; Related .......: _LOImpress_TableCellCharEffect, _LOImpress_TableCharOverLine, _LOImpress_TableCharStrikeOut, _LOImpress_TableCharUnderLine
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCharEffect(ByRef $oTable, $iCase = Null, $iRelief = Null, $bOutline = Null, $bShadow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCell
	Local $iError = 0, $iTempCount = 0
	Local $avEffect[4], $avTemp[4]

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iCase, $iRelief, $bOutline, $bShadow) Then
		For $iRow = 0 To $oTable.Model.RowCount() - 1
			For $iCol = 0 To $oTable.Model.ColumnCount() - 1
				$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
				If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

				If ($iRow = 0) And ($iCol = 0) Then     ; Retrieve the value once the first time, to use to test against the rest.
					__LO_ArrayFill($avEffect, $oCell.CharCaseMap(), $oCell.CharRelief(), $oCell.CharContoured(), $oCell.CharShadowed())

				Else
					__LO_ArrayFill($avTemp, $oCell.CharCaseMap(), $oCell.CharRelief(), $oCell.CharContoured(), $oCell.CharShadowed())

					$iTempCount = 0

					; Cycle through the Table and retrieve the current value for each cell, if it isn't the same, change the return to Null.
					For $i = 0 To UBound($avEffect) - 1
						If ($avEffect[$i] <> $avTemp[$i]) Then
							$avEffect[$i] = Null
							$iTempCount += 1

						ElseIf ($avEffect[$i] = Null) Then
							$iTempCount += 1
						EndIf

						If ($iTempCount = UBound($avEffect)) Then ExitLoop 3 ; Exit the loops if all values are already nulled.

						Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
					Next
				EndIf
			Next
		Next

		Return SetError($__LO_STATUS_SUCCESS, 1, $avEffect)
	EndIf

	For $iCol = 0 To $oTable.Model.ColumnCount() - 1
		For $iRow = 0 To $oTable.Model.RowCount() - 1
			$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
			If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			If ($iCase <> Null) Then
				If Not __LO_IntIsBetween($iCase, $LOI_CHAR_CASEMAP_NONE, $LOI_CHAR_CASEMAP_SM_CAPS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

				$oCell.CharCaseMap = $iCase
				$iError = ($oCell.CharCaseMap() = $iCase) ? ($iError) : (BitOR($iError, 1))
			EndIf

			If ($iRelief <> Null) Then
				If Not __LO_IntIsBetween($iRelief, $LOI_CHAR_RELIEF_NONE, $LOI_CHAR_RELIEF_ENGRAVED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

				$oCell.CharRelief = $iRelief
				$iError = ($oCell.CharRelief() = $iRelief) ? ($iError) : (BitOR($iError, 2))
			EndIf

			If ($bOutline <> Null) Then
				If Not IsBool($bOutline) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

				$oCell.CharContoured = $bOutline
				$iError = ($oCell.CharContoured() = $bOutline) ? ($iError) : (BitOR($iError, 4))
			EndIf

			If ($bShadow <> Null) Then
				If Not IsBool($bShadow) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

				$oCell.CharShadowed = $bShadow
				$iError = ($oCell.CharShadowed() = $bShadow) ? ($iError) : (BitOR($iError, 8))
			EndIf
		Next
	Next

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_TableCharEffect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCharFont
; Description ...: Set and Retrieve the Font Settings for a Table.
; Syntax ........: _LOImpress_TableCharFont(ByRef $oTable[, $sFontName = Null[, $nFontSize = Null[, $iPosture = Null[, $iWeight = Null]]]])
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
;                  $sFontName           - [optional] Default is Null. The Font Name to use.
;                  $nFontSize           - [optional] Default is Null. The new Font size.
;                  $iPosture            - [optional] (0-5) Default is Null. The Font Italic setting. See Constants, $LOI_CHAR_POSTURE_* as defined in LibreOfficeImpress_Constants.au3. Also see remarks.
;                  $iWeight             - [optional] (0, 50-200) Default is Null. The Font Bold settings see Constants, $LOI_CHAR_WEIGHT_* as defined in LibreOfficeImpress_Constants.au3. Also see remarks.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 4 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  @Error: 1, @Extended: 2 = $sFontName not a String.
;                  @Error: 1, @Extended: 3 = Font called in $sFontName not available.
;                  @Error: 1, @Extended: 4 = $nFontSize not a number.
;                  @Error: 1, @Extended: 5 = $iPosture not an Integer, less than 0 or greater than 5. See Constants, $LOI_CHAR_POSTURE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 6 = $iWeight not an Integer, less than 50 but not equal to 0, or greater than 200. See Constants, $LOI_CHAR_WEIGHT_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Cell Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sFontName
;                  |                               2 = Error setting $nFontSize
;                  |                               4 = Error setting $iPosture
;                  |                               8 = Error setting $iWeight
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Not every font accepts Bold and Italic settings, and not all settings for bold and Italic are accepted, such as oblique, ultra Bold etc.
;                  LibreOffice accepts only the predefined weight values, any other values are changed automatically to an acceptable value, which could trigger a settings error.
;                  Tables require that the properties be set individually for each Cell, therefore this function cycles through each cell and sets the value, and may be slower for large tables.
;                  When retrieving the current property values for a table, if all of the cells in the Table do not have the same value, Null is returned for that property value.
; Related .......: _LOImpress_FontsGetNames, _LOImpress_TableCharFontColor, _LOImpress_TableCellCharFont
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCharFont(ByRef $oTable, $sFontName = Null, $nFontSize = Null, $iPosture = Null, $iWeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCell
	Local $iError = 0, $iTempCount = 0
	Local $avFont[4], $avTemp[4]

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($sFontName, $nFontSize, $iPosture, $iWeight) Then
		For $iRow = 0 To $oTable.Model.RowCount() - 1
			For $iCol = 0 To $oTable.Model.ColumnCount() - 1
				$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
				If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

				If ($iRow = 0) And ($iCol = 0) Then     ; Retrieve the value once the first time, to use to test against the rest.
					__LO_ArrayFill($avFont, $oCell.CharFontName(), $oCell.CharHeight(), $oCell.CharPosture(), $oCell.CharWeight())

				Else
					__LO_ArrayFill($avTemp, $oCell.CharFontName(), $oCell.CharHeight(), $oCell.CharPosture(), $oCell.CharWeight())

					$iTempCount = 0

					; Cycle through the Table and retrieve the current value for each cell, if it isn't the same, change the return to Null.
					For $i = 0 To UBound($avFont) - 1
						If ($avFont[$i] <> $avTemp[$i]) Then
							$avFont[$i] = Null
							$iTempCount += 1

						ElseIf ($avFont[$i] = Null) Then
							$iTempCount += 1
						EndIf

						If ($iTempCount = UBound($avFont)) Then ExitLoop 3 ; Exit the loops if all values are already nulled.

						Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
					Next
				EndIf
			Next
		Next

		Return SetError($__LO_STATUS_SUCCESS, 1, $avFont)
	EndIf

	If ($sFontName <> Null) Then ; Error check for font outside of the loop to prevent unneeded delay.
		If Not IsString($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		If Not _LOImpress_FontExists($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	EndIf

	For $iCol = 0 To $oTable.Model.ColumnCount() - 1
		For $iRow = 0 To $oTable.Model.RowCount() - 1
			$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
			If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			If ($sFontName <> Null) Then
				$oCell.CharFontName = $sFontName
				$iError = ($oCell.CharFontName() = $sFontName) ? ($iError) : (BitOR($iError, 1))
			EndIf

			If ($nFontSize <> Null) Then
				If Not IsNumber($nFontSize) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

				$oCell.CharHeight = $nFontSize
				$iError = ($oCell.CharHeight() = $nFontSize) ? ($iError) : (BitOR($iError, 2))
			EndIf

			If ($iPosture <> Null) Then
				If Not __LO_IntIsBetween($iPosture, $LOI_CHAR_POSTURE_NONE, $LOI_CHAR_POSTURE_ITALIC) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

				$oCell.CharPosture = $iPosture
				$iError = ($oCell.CharPosture() = $iPosture) ? ($iError) : (BitOR($iError, 4))
			EndIf

			If ($iWeight <> Null) Then
				If Not __LO_IntIsBetween($iWeight, $LOI_CHAR_WEIGHT_THIN, $LOI_CHAR_WEIGHT_BLACK, "", $LOI_CHAR_WEIGHT_DONT_KNOW) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

				$oCell.CharWeight = $iWeight
				$iError = ($oCell.CharWeight() = $iWeight) ? ($iError) : (BitOR($iError, 8))
			EndIf
		Next
	Next

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_TableCharFont

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCharFontColor
; Description ...: Set or retrieve the font color, transparency and highlighting values for a Table.
; Syntax ........: _LOImpress_TableCharFontColor(ByRef $oTable[, $iFontColor = Null[, $iTransparency = Null[, $iHighlight = Null]]])
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
;                  $iFontColor          - [optional] (-1-16777215) Default is Null. The font Color value, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for Auto color.
;                  $iTransparency       - [optional] (0-100) Default is Null. Transparency percentage. 0 is visible, 100 is invisible. Available for LibreOffice 7.0 and up.
;                  $iHighlight          - [optional] (-1-16777215) Default is Null. The highlight Color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for No color.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters. If The current LibreOffice version is below 7.0 the $iTransparency parameter will return a Null value.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  @Error: 1, @Extended: 2 = $iFontColor not an Integer, less than -1 or greater than 16777215.
;                  @Error: 1, @Extended: 3 = $iTransparency not an Integer, less than 0 or greater than 100%.
;                  @Error: 1, @Extended: 4 = $iHighlight not an Integer, less than -1 or greater than 16777215.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Cell Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve old Transparency value.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $FontColor
;                  |                               2 = Error setting $iTransparency.
;                  |                               4 = Error setting $iHighlight
;                  --Version Related Errors--
;                  @Error: 6, @Extended: 1 = Current LibreOffice version lower than 7.0.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Tables require that the properties be set individually for each Cell, therefore this function cycles through each cell and sets the value, and may be slower for large tables.
;                  When retrieving the current property values for a table, if all of the cells in the Table do not have the same value, Null is returned for that property value.
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LOImpress_TableCharFont, _LOImpress_TableCellCharFontColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCharFontColor(ByRef $oTable, $iFontColor = Null, $iTransparency = Null, $iHighlight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCell
	Local $iError = 0, $iTempCount = 0, $iOldTransparency
	Local $avColor[3], $avTemp[3]

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iFontColor, $iTransparency, $iHighlight) Then
		For $iRow = 0 To $oTable.Model.RowCount() - 1
			For $iCol = 0 To $oTable.Model.ColumnCount() - 1
				$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
				If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

				If ($iRow = 0) And ($iCol = 0) Then     ; Retrieve the value once the first time, to use to test against the rest.
					If __LO_VersionCheck(7.0) Then
						__LO_ArrayFill($avColor, __LOImpress_ColorRemoveAlpha($oCell.CharColor()), $oCell.CharTransparence(), $oCell.CharBackColor())

					Else
						__LO_ArrayFill($avColor, __LOImpress_ColorRemoveAlpha($oCell.CharColor()), Null, $oCell.CharBackColor())
					EndIf

				Else
					If __LO_VersionCheck(7.0) Then
						__LO_ArrayFill($avTemp, __LOImpress_ColorRemoveAlpha($oCell.CharColor()), $oCell.CharTransparence(), $oCell.CharBackColor())

					Else
						__LO_ArrayFill($avTemp, __LOImpress_ColorRemoveAlpha($oCell.CharColor()), Null, $oCell.CharBackColor())
					EndIf

					$iTempCount = 0

					; Cycle through the Table and retrieve the current value for each cell, if it isn't the same, change the return to Null.
					For $i = 0 To UBound($avColor) - 1
						If ($avColor[$i] <> $avTemp[$i]) Then
							$avColor[$i] = Null
							$iTempCount += 1

						ElseIf ($avColor[$i] = Null) Then
							$iTempCount += 1
						EndIf

						If ($iTempCount = UBound($avColor)) Then ExitLoop 3 ; Exit the loops if all values are already nulled.

						Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
					Next
				EndIf
			Next
		Next

		Return SetError($__LO_STATUS_SUCCESS, 1, $avColor)
	EndIf

	For $iCol = 0 To $oTable.Model.ColumnCount() - 1
		For $iRow = 0 To $oTable.Model.RowCount() - 1
			$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
			If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			If ($iFontColor <> Null) Then
				If Not __LO_IntIsBetween($iFontColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

				If __LO_VersionCheck(7.0) Then
					$iOldTransparency = $oCell.CharTransparence()
					If Not IsInt($iOldTransparency) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
				EndIf

				$oCell.CharColor = $iFontColor
				$iError = ($oCell.CharColor() = $iFontColor) ? ($iError) : (BitOR($iError, 1))

				If IsInt($iOldTransparency) Then $oCell.CharTransparence = $iOldTransparency
			EndIf

			If ($iTransparency <> Null) Then
				If Not __LO_IntIsBetween($iTransparency, 0, 100) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
				If Not __LO_VersionCheck(7.0) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)

				$oCell.CharTransparence = $iTransparency
				$iError = ($oCell.CharTransparence() = $iTransparency) ? ($iError) : (BitOR($iError, 2))
			EndIf

			If ($iHighlight <> Null) Then
				If Not __LO_IntIsBetween($iHighlight, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

				; CharHighlight; same as CharBackColor---Libre seems to use back color for highlighting however, so using that for setting.
;~ 		If Not __LO_VersionCheck(4.2) Then Return SetError($__LO_STATUS_VER_ERROR, 2, 0)
;~ 		$oCell.CharHighlight = $iHighlight ;-- keeping old method in case.
;~ 		$iError = ($oCell.CharHighlight() = $iHighlight) ? ($iError) : (BitOR($iError, 4)
				$oCell.CharBackColor = $iHighlight
				$iError = ($oCell.CharBackColor() = $iHighlight) ? ($iError) : (BitOR($iError, 4))
			EndIf
		Next
	Next

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_TableCharFontColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCharOverLine
; Description ...: Set and retrieve the OverLine settings for a Table.
; Syntax ........: _LOImpress_TableCharOverLine(ByRef $oTable[, $iOverLineStyle = Null[, $iOLColor = Null[, $bWordOnly = Null]]])
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
;                  $iOverLineStyle      - [optional] (0-18) Default is Null. The style of the Overline line, see constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  $iOLColor            - [optional] (-1-16777215) Default is Null. The Overline color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for automatic color mode.
;                  $bWordOnly           - [optional] Default is Null. If True, white spaces are not Overlined.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  @Error: 1, @Extended: 2 = $iOverLineStyle not an Integer, less than 0 or greater than 18. See constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3. See Remarks.
;                  @Error: 1, @Extended: 3 = $iOLColor not an Integer, less than -1 or greater than 16777215.
;                  @Error: 1, @Extended: 4 = $bWordOnly not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Cell Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iOverLineStyle
;                  |                               2 = Error setting $iOLColor
;                  |                               4 = Error setting $bWordOnly
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Overline line style uses the same constants as underline style.
;                  To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Tables require that the properties be set individually for each Cell, therefore this function cycles through each cell and sets the value, and may be slower for large tables.
;                  When retrieving the current property values for a table, if all of the cells in the Table do not have the same value, Null is returned for that property value.
; Related .......: _LOImpress_TableCellCharOverLine, _LOImpress_TableCharEffect, _LOImpress_TableCharStrikeOut, _LOImpress_TableCharUnderLine
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCharOverLine(ByRef $oTable, $iOverLineStyle = Null, $iOLColor = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCell
	Local $iError = 0, $iTempCount = 0
	Local $avOverLine[3], $avTemp[3]

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iOverLineStyle, $iOLColor, $bWordOnly) Then
		For $iRow = 0 To $oTable.Model.RowCount() - 1
			For $iCol = 0 To $oTable.Model.ColumnCount() - 1
				$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
				If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

				If ($iRow = 0) And ($iCol = 0) Then     ; Retrieve the value once the first time, to use to test against the rest.
					__LO_ArrayFill($avOverLine, $oCell.CharOverline(), $oCell.CharOverlineColor(), $oCell.CharWordMode())

				Else
					__LO_ArrayFill($avTemp, $oCell.CharOverline(), $oCell.CharOverlineColor(), $oCell.CharWordMode())

					$iTempCount = 0

					; Cycle through the Table and retrieve the current value for each cell, if it isn't the same, change the return to Null.
					For $i = 0 To UBound($avOverLine) - 1
						If ($avOverLine[$i] <> $avTemp[$i]) Then
							$avOverLine[$i] = Null
							$iTempCount += 1

						ElseIf ($avOverLine[$i] = Null) Then
							$iTempCount += 1
						EndIf

						If ($iTempCount = UBound($avOverLine)) Then ExitLoop 3 ; Exit the loops if all values are already nulled.

						Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
					Next
				EndIf
			Next
		Next

		Return SetError($__LO_STATUS_SUCCESS, 1, $avOverLine)
	EndIf

	For $iCol = 0 To $oTable.Model.ColumnCount() - 1
		For $iRow = 0 To $oTable.Model.RowCount() - 1
			$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
			If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			If ($iOverLineStyle <> Null) Then
				If Not __LO_IntIsBetween($iOverLineStyle, $LOI_CHAR_UNDERLINE_NONE, $LOI_CHAR_UNDERLINE_BOLD_WAVE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

				$oCell.CharOverline = $iOverLineStyle
				$iError = ($oCell.CharOverline() = $iOverLineStyle) ? ($iError) : (BitOR($iError, 1))
			EndIf

			If ($iOLColor <> Null) Then
				If Not __LO_IntIsBetween($iOLColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

				If ($iOLColor = $LO_COLOR_OFF) Then
					If ($oCell.CharOverlineHasColor() = True) Then $oCell.CharOverlineHasColor = False
					$oCell.CharOverlineColor = $iOLColor

				Else
					If ($oCell.CharOverlineHasColor() = False) Then $oCell.CharOverlineHasColor = True
					$oCell.CharOverlineColor = $iOLColor
				EndIf

				$oCell.CharOverlineColor = $iOLColor
				$iError = ($oCell.CharOverlineColor() = $iOLColor) ? ($iError) : (BitOR($iError, 2))
			EndIf

			If ($bWordOnly <> Null) Then
				If Not IsBool($bWordOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

				$oCell.CharWordMode = $bWordOnly
				$iError = ($oCell.CharWordMode() = $bWordOnly) ? ($iError) : (BitOR($iError, 4))
			EndIf
		Next
	Next

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_TableCharOverLine

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCharStrikeOut
; Description ...: Set or Retrieve the Strikeout settings for a Table.
; Syntax ........: _LOImpress_TableCharStrikeOut(ByRef $oTable[, $iStrikeLineStyle = Null[, $bWordOnly = Null]])
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
;                  $iStrikeLineStyle    - [optional] (0-6) Default is Null. The Strikeout Line Style, see constants, $LOI_CHAR_STRIKEOUT_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bWordOnly           - [optional] Default is Null. If True, strike out is applied to words only, skipping whitespaces.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  @Error: 1, @Extended: 2 = $iStrikeLineStyle not an Integer, less than 0 or greater than 6. See constants, $LOI_CHAR_STRIKEOUT_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $bWordOnly not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Cell Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iStrikeLineStyle
;                  |                               2 = Error setting $bWordOnly
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Tables require that the properties be set individually for each Cell, therefore this function cycles through each cell and sets the value, and may be slower for large tables.
;                  When retrieving the current property values for a table, if all of the cells in the Table do not have the same value, Null is returned for that property value.
; Related .......: _LOImpress_TableCellCharStrikeOut, _LOImpress_TableCharEffect, _LOImpress_TableCharOverLine, _LOImpress_TableCharUnderLine
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCharStrikeOut(ByRef $oTable, $iStrikeLineStyle = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCell
	Local $iError = 0, $iTempCount = 0
	Local $avStrikeOut[2], $avTemp[2]

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iStrikeLineStyle, $bWordOnly) Then
		For $iRow = 0 To $oTable.Model.RowCount() - 1
			For $iCol = 0 To $oTable.Model.ColumnCount() - 1
				$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
				If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

				If ($iRow = 0) And ($iCol = 0) Then     ; Retrieve the value once the first time, to use to test against the rest.
					__LO_ArrayFill($avStrikeOut, $oCell.CharStrikeout(), $oCell.CharWordMode())

				Else
					__LO_ArrayFill($avTemp, $oCell.CharStrikeout(), $oCell.CharWordMode())

					$iTempCount = 0

					; Cycle through the Table and retrieve the current value for each cell, if it isn't the same, change the return to Null.
					For $i = 0 To UBound($avStrikeOut) - 1
						If ($avStrikeOut[$i] <> $avTemp[$i]) Then
							$avStrikeOut[$i] = Null
							$iTempCount += 1

						ElseIf ($avStrikeOut[$i] = Null) Then
							$iTempCount += 1
						EndIf

						If ($iTempCount = UBound($avStrikeOut)) Then ExitLoop 3 ; Exit the loops if all values are already nulled.

						Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
					Next
				EndIf
			Next
		Next

		Return SetError($__LO_STATUS_SUCCESS, 1, $avStrikeOut)
	EndIf

	For $iCol = 0 To $oTable.Model.ColumnCount() - 1
		For $iRow = 0 To $oTable.Model.RowCount() - 1
			$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
			If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			If ($iStrikeLineStyle <> Null) Then
				If Not __LO_IntIsBetween($iStrikeLineStyle, $LOI_CHAR_STRIKEOUT_NONE, $LOI_CHAR_STRIKEOUT_X) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

				$oCell.CharStrikeout = $iStrikeLineStyle
				$iError = ($oCell.CharStrikeout() = $iStrikeLineStyle) ? ($iError) : (BitOR($iError, 1))
			EndIf

			If ($bWordOnly <> Null) Then
				If Not IsBool($bWordOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

				$oCell.CharWordMode = $bWordOnly
				$iError = ($oCell.CharWordMode() = $bWordOnly) ? ($iError) : (BitOR($iError, 2))
			EndIf
		Next
	Next

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_TableCharStrikeOut

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableCharUnderLine
; Description ...: Set and retrieve the Underline settings for a Table.
; Syntax ........: _LOImpress_TableCharUnderLine(ByRef $oTable[, $iUnderLineStyle = Null[, $iULColor = Null[, $bWordOnly = Null]]])
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
;                  $iUnderLineStyle     - [optional] (0-18) Default is Null. The Underline line style, see constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iULColor            - [optional] (-1-16777215) Default is Null. The underline color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Call with $LO_COLOR_OFF(-1) for automatic color mode.
;                  $bWordOnly           - [optional] Default is Null. If True, white spaces are not underlined.
; Return values .: Success: 1 or Array
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable an Object.
;                  @Error: 1, @Extended: 2 = $iUnderLineStyle not an Integer, less than 0 or greater than 18. See constants, $LOI_CHAR_UNDERLINE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $iULColor not an Integer, less than -1 or greater than 16777215.
;                  @Error: 1, @Extended: 4 = $bWordOnly not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Cell Object.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iUnderLineStyle
;                  |                               2 = Error setting $iULColor
;                  |                               4 = Error setting $bWordOnly
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  Tables require that the properties be set individually for each Cell, therefore this function cycles through each cell and sets the value, and may be slower for large tables.
;                  When retrieving the current property values for a table, if all of the cells in the Table do not have the same value, Null is returned for that property value.
; Related .......: _LOImpress_TableCharEffect, _LOImpress_TableCharOverLine, _LOImpress_TableCharStrikeOut, _LOImpress_TableCellCharUnderLine
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableCharUnderLine(ByRef $oTable, $iUnderLineStyle = Null, $iULColor = Null, $bWordOnly = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCell
	Local $iError = 0, $iTempCount = 0
	Local $avUnderLine[3], $avTemp[3]

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($iUnderLineStyle, $iULColor, $bWordOnly) Then
		For $iRow = 0 To $oTable.Model.RowCount() - 1
			For $iCol = 0 To $oTable.Model.ColumnCount() - 1
				$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
				If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

				If ($iRow = 0) And ($iCol = 0) Then     ; Retrieve the value once the first time, to use to test against the rest.
					__LO_ArrayFill($avUnderLine, $oCell.CharUnderline(), $oCell.CharUnderlineColor(), $oCell.CharWordMode())

				Else
					__LO_ArrayFill($avTemp, $oCell.CharUnderline(), $oCell.CharUnderlineColor(), $oCell.CharWordMode())

					$iTempCount = 0

					; Cycle through the Table and retrieve the current value for each cell, if it isn't the same, change the return to Null.
					For $i = 0 To UBound($avUnderLine) - 1
						If ($avUnderLine[$i] <> $avTemp[$i]) Then
							$avUnderLine[$i] = Null
							$iTempCount += 1

						ElseIf ($avUnderLine[$i] = Null) Then
							$iTempCount += 1
						EndIf

						If ($iTempCount = UBound($avUnderLine)) Then ExitLoop 3 ; Exit the loops if all values are already nulled.

						Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
					Next
				EndIf
			Next
		Next

		Return SetError($__LO_STATUS_SUCCESS, 1, $avUnderLine)
	EndIf

	For $iCol = 0 To $oTable.Model.ColumnCount() - 1
		For $iRow = 0 To $oTable.Model.RowCount() - 1
			$oCell = $oTable.Model.getCellByPosition($iCol, $iRow)
			If Not IsObj($oCell) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

			If ($iUnderLineStyle <> Null) Then
				If Not __LO_IntIsBetween($iUnderLineStyle, $LOI_CHAR_UNDERLINE_NONE, $LOI_CHAR_UNDERLINE_BOLD_WAVE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

				$oCell.CharUnderline = $iUnderLineStyle
				$iError = ($oCell.CharUnderline() = $iUnderLineStyle) ? ($iError) : (BitOR($iError, 1))
			EndIf

			If ($iULColor <> Null) Then
				If Not __LO_IntIsBetween($iULColor, $LO_COLOR_OFF, $LO_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

				If ($iULColor = $LO_COLOR_OFF) Then
					If ($oCell.CharUnderlineHasColor() = True) Then $oCell.CharUnderlineHasColor = False
					$oCell.CharUnderlineColor = $iULColor

				Else
					If ($oCell.CharUnderlineHasColor() = False) Then $oCell.CharUnderlineHasColor = True
					$oCell.CharUnderlineColor = $iULColor
				EndIf

				$iError = ($oCell.CharUnderlineColor() = $iULColor) ? ($iError) : (BitOR($iError, 2))
			EndIf

			If ($bWordOnly <> Null) Then
				If Not IsBool($bWordOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

				$oCell.CharWordMode = $bWordOnly
				$iError = ($oCell.CharWordMode() = $bWordOnly) ? ($iError) : (BitOR($iError, 4))
			EndIf
		Next
	Next

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_TableCharUnderLine

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableColumnDelete
; Description ...: Delete a column from a Table.
; Syntax ........: _LOImpress_TableColumnDelete(ByRef $oTable, $iColumn[, $iCount = 1])
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
;                  $iColumn             - The Column index to begin deleting from. 0 based.
;                  $iCount              - [optional] Default is 1. The number of columns to delete starting at the column called in $iColumn and moving right.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: ?, Return: 1 = Success. Full amount of columns deleted. @Extended set to total columns deleted.
;                  @Error: 0, @Extended: ?, Return: 2 = Success. $iCount greater than amount of columns contained in Table; deleted all columns from $iColumn over. @Extended set to total columns deleted.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  @Error: 1, @Extended: 2 = $iColumn not an Integer, less than 0 or greater than number of columns in the Table.
;                  @Error: 1, @Extended: 3 = $iCount not an Integer, or less than 1.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve count of columns contained in table.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: LibreOffice counts columns and Rows starting at 0. So to delete the first column in a Table you would call $iColumn with 0.
;                  If you attempt to delete more columns than are present all columns from $iColumn over will be deleted.
;                  If you delete all columns starting from column 0, the entire Table is deleted.
; Related .......: _LOImpress_TableColumnGetCount, _LOImpress_TableColumnInsert, _LOImpress_TableRowDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableColumnDelete(ByRef $oTable, $iColumn, $iCount = 1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iColumnCount, $iReturn = 0

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iColumnCount = $oTable.Model.ColumnCount()
	If Not IsInt($iColumnCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iColumn, 0, $iColumnCount - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iCount, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$iCount = ($iCount > ($iColumnCount - $iColumn)) ? ($iColumnCount - $iColumn) : ($iCount) ; See if I can delete the full amount, else delete as many as possible.
	$iReturn = ($iCount > ($iColumnCount - $iColumn)) ? (2) : (1) ; Return 1 if full amount deleted else 2 if only partial.
	$oTable.Model.getColumns.removeByIndex($iColumn, $iCount)

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $iReturn)
EndFunc   ;==>_LOImpress_TableColumnDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableColumnGetCount
; Description ...: Retrieve the number of Columns in a table.
; Syntax ........: _LOImpress_TableColumnGetCount(ByRef $oTable)
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Returning Column Count as an Integer.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Column count.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_TableColumnDelete, _LOImpress_TableColumnInsert, _LOImpress_TableRowGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableColumnGetCount(ByRef $oTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iColumnCount

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iColumnCount = $oTable.Model.ColumnCount()
	If Not IsInt($iColumnCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iColumnCount)
EndFunc   ;==>_LOImpress_TableColumnGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableColumnInsert
; Description ...: Insert column(s) into a Text Table
; Syntax ........: _LOImpress_TableColumnInsert(ByRef $oTable[, $iCount = 1[, $iColumn = Null]])
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
;                  $iCount              - [optional] Default is 1. The number of columns to insert.
;                  $iColumn             - [optional] Default is Null. The column index to insert columns after. 0 based. See Remarks.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully inserted the number of desired columns.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  @Error: 1, @Extended: 2 = $iCount not an Integer, or less than 1.
;                  @Error: 1, @Extended: 3 = $iColumn not an Integer, less than 0 or greater than number of columns contained in the Table.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve count of columns.
;                  @Error: 3, @Extended: 2 = Failed to insert columns.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call $iColumn with Null to insert the column(s) at the end (right-hand side) of the Table.
;                  LibreOffice counts the Table columns/Rows starting at 0. The columns are placed behind the desired column when inserted.
;                  To insert a column at the left most of the Table you would call $iColumn to 0. To insert columns at the Right of a table you would call $iColumn to one higher than the last column. e.g. a Table containing 3 columns, would be numbered as follows: 0(first-Column), 1(second-Column), 2(third-Column), to insert columns at the very Right of the columns, you would call $iColumn to 3.
; Related .......: _LOImpress_TableColumnDelete, _LOImpress_TableColumnGetCount, _LOImpress_TableRowInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableColumnInsert(ByRef $oTable, $iCount = 1, $iColumn = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iColumnCount

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iCount, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$iColumnCount = $oTable.Model.ColumnCount()
	If Not IsInt($iColumnCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; Requested column out of bounds.
	If ($iColumn <> Null) And Not __LO_IntIsBetween($iColumn, 0, $iColumnCount) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$iColumn = ($iColumn = Null) ? ($iColumnCount) : ($iColumn) ; Insert the new Columns at the end of the Table if $iColumn isn't defined.

	$oTable.Model.getColumns.insertByIndex($iColumn, $iCount)
	If ($oTable.Model.ColumnCount() <> ($iColumnCount + $iCount)) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_TableColumnInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableInsert
; Description ...: Create and Insert a Table into a Slide.
; Syntax ........: _LOImpress_TableInsert(ByRef $oSlide, $iWidth, $iHeight[, $iRows = 2[, $iColumns = 2[, $iX = -1[, $iY = -1]]]])
; Parameters ....: $oSlide              - A Slide object returned by a previous _LOImpress_SlideAdd, _LOImpress_SlideGetObjByIndex, _LOImpress_SlideGetObjByName, or _LOImpress_SlideCopy function.
;                  $iWidth              - The Table's Width in Hundredths of a Millimeter (HMM).
;                  $iHeight             - The Table's Height in Hundredths of a Millimeter (HMM).
;                  $iRows               - [optional] (1-75) Default is 2. The number of Rows.
;                  $iColumns            - [optional] (1-75) Default is 2. The number of Columns.
;                  $iX                  - [optional] Default is -1. The X position from the top-left of the page, in Hundredths of a Millimeter (HMM). Call with -1 to center the table horizontally.
;                  $iY                  - [optional] Default is -1. The Y position from the top-left of the page, in Hundredths of a Millimeter (HMM). Call with -1 to center the table vertically.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Inserted a new Table. Returning its Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oSlide not an Object.
;                  @Error: 1, @Extended: 2 = $iWidth not an Integer.
;                  @Error: 1, @Extended: 3 = $iHeight not an Integer.
;                  @Error: 1, @Extended: 4 = $iRows not an Integer, less than 1 or greater than 75.
;                  @Error: 1, @Extended: 5 = $iColumns not an Integer, less than 1 or greater than 75.
;                  @Error: 1, @Extended: 6 = $iX not an Integer.
;                  @Error: 1, @Extended: 7 = $iY not an Integer.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a "com.sun.star.drawing.TableShape" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve parent Document Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Default Table Style.
;                  @Error: 3, @Extended: 3 = Failed to retrieve Position Structure.
;                  @Error: 3, @Extended: 4 = Failed to retrieve Size Structure.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_ShapeDelete, _LOImpress_ShapeImageInsert, _LOImpress_ShapeTextBoxInsert, _LOImpress_DrawShapeInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableInsert(ByRef $oSlide, $iWidth, $iHeight, $iRows = 2, $iColumns = 2, $iX = -1, $iY = -1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShape, $oDoc, $oStyle
	Local $tSize, $tPos

	If Not IsObj($oSlide) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iWidth) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iHeight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_IntIsBetween($iRows, 1, 75) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not __LO_IntIsBetween($iColumns, 1, 75) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsInt($iX) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not IsInt($iY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	$oDoc = $oSlide.MasterPage.Forms.Parent()
	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oShape = $oDoc.createInstance("com.sun.star.drawing.TableShape")
	If Not IsObj($oShape) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oStyle = $oDoc.StyleFamilies.getByName("table").getByName("default")
	If Not IsObj($oStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	; Apply Default Table Style, to match L.O.
	$oShape.TableTemplate = $oStyle

	$oSlide.add($oShape)

	$oShape.Model.Rows.insertByIndex(0, ($iRows - 1)) ; Minus one to account for 1 Column/Row being present upon creation.
	$oShape.Model.Columns.insertByIndex(0, ($iColumns - 1)) ; Minus one to account for 1 Column/Row being present upon creation.

	$tPos = $oShape.Position()
	If Not IsObj($tPos) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	$tPos.X = ($iX = -1) ? (Int(($oSlide.Width() - $iWidth) / 2)) : ($iX)
	$tPos.Y = ($iY = -1) ? (Int(($oSlide.Height() - $iHeight) / 2)) : ($iY)

	$oShape.Position = $tPos

	$tSize = $oShape.Size()
	If Not IsObj($tSize) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

	$tSize.Width = $iWidth
	$tSize.Height = $iHeight

	$oShape.Size = $tSize

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShape)
EndFunc   ;==>_LOImpress_TableInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableRowDelete
; Description ...: Delete a row from a Table.
; Syntax ........: _LOImpress_TableRowDelete(ByRef $oTable, $iRow[, $iCount = 1])
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
;                  $iRow                - The row index to begin deleting from. 0 based.
;                  $iCount              - [optional] Default is 1. The number of rows to delete starting at $iRow and moving down.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: ?, Return: 1 = Success. Full amount of Rows deleted. @Extended set to total rows deleted.
;                  @Error: 0, @Extended: ?, Return: 2 = Success. $iCount greater than amount of rows contained in Table; deleted all rows from $iRow over. @Extended set to total rows deleted.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  @Error: 1, @Extended: 2 = $iRow not an Integer, less than 0 or greater than number of rows in the Table.
;                  @Error: 1, @Extended: 3 = $iCount not an Integer, or less than 1.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve count of rows contained in table.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: LibreOffice counts Rows starting at 0. So to delete the first Row in a Table you would set $iRow to 0.
;                  If you attempt to delete more rows than are present, all rows from $iRow over will be deleted.
;                  If you delete all Rows starting from Row 0, the entire Table is deleted.
; Related .......: _LOImpress_TableRowGetCount, _LOImpress_TableRowInsert, _LOImpress_TableColumnDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableRowDelete(ByRef $oTable, $iRow, $iCount = 1)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iRowCount, $iReturn = 0

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iRowCount = $oTable.Model.RowCount()
	If Not IsInt($iRowCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iRow, 0, $iRowCount - 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iCount, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$iCount = ($iCount > ($iRowCount - $iRow)) ? ($iRowCount - $iRow) : ($iCount) ; See if I can delete the full amount, else delete as many as possible.
	$iReturn = ($iCount > ($iRowCount - $iRow)) ? (2) : (1) ; Return 1 if full amount deleted else 2 if only partial.
	$oTable.Model.getRows.removeByIndex($iRow, $iCount)

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $iReturn)
EndFunc   ;==>_LOImpress_TableRowDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableRowGetCount
; Description ...: Retrieve the number of Rows in a table.
; Syntax ........: _LOImpress_TableRowGetCount(ByRef $oTable)
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Returning Row Count as an Integer.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve Row count.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_TableRowDelete, _LOImpress_TableRowInsert, _LOImpress_TableColumnGetCount
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableRowGetCount(ByRef $oTable)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iRowCount

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iRowCount = $oTable.Model.RowCount()
	If Not IsInt($iRowCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iRowCount)
EndFunc   ;==>_LOImpress_TableRowGetCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableRowInsert
; Description ...: Insert a row into a Table
; Syntax ........: _LOImpress_TableRowInsert(ByRef $oTable[, $iCount = 1[, $iRow = Null]])
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
;                  $iCount              - [optional] Default is 1. The number of rows to insert.
;                  $iRow                - [optional] Default is Null. The row index to insert rows after. See Remarks.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully inserted the number of desired rows.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  @Error: 1, @Extended: 2 = $iCount not an Integer, or less than 1.
;                  @Error: 1, @Extended: 3 = $iRow not an Integer, less than 0 or greater than number of rows contained in the Table.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve count of rows.
;                  @Error: 3, @Extended: 2 = Failed to insert rows.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call $iRow with Null to insert the row(s) at the bottom of the Table.
;                  LibreOffice counts the Table Rows starting at 0. The Rows are placed above the desired Row when inserted.
;                  To insert a Row at the top most of the Table call $iRow with 0.
;                  To insert rows at the bottom of a table you would call $iRow with one higher than the last row. e.g. a Table containing 3 rows, would be numbered as follows: 0(first-row), 1(second-row), 2(third-row), to insert rows at the very bottom of the rows, call $iRow with 3.
; Related .......: _LOImpress_TableRowDelete, _LOImpress_TableRowGetCount, _LOImpress_TableColumnInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableRowInsert(ByRef $oTable, $iCount = 1, $iRow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iRowCount

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iCount, 1) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($iRow <> Null) And Not __LO_IntIsBetween($iRow, 0, $oTable.Model.RowCount()) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$iRowCount = $oTable.Model.RowCount()
	If Not IsInt($iRowCount) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iRow = ($iRow = Null) ? ($iRowCount) : ($iRow)
	$oTable.Model.getRows.insertByIndex($iRow, $iCount)
	If ($oTable.Model.RowCount() <> ($iRowCount + $iCount)) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_TableRowInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_TableShadow
; Description ...: Set or Retrieve the shadow settings for a Table.
; Syntax ........: _LOImpress_TableShadow(ByRef $oTable[, $bShadow = Null[, $iLocation = Null[, $iColor = Null[, $iDistance = Null[, $iBlur = Null[, $iTransparency = Null]]]]]])
; Parameters ....: $oTable              - A Table Shape object returned by a previous _LOImpress_TableInsert, or _LOImpress_ShapesGetList function.
;                  $bShadow             - [optional] Default is Null. If True, a Shadow is present for the Shape.
;                  $iLocation           - [optional] (0-8) Default is Null. The Location of the Shadow, must be one of the Constants, $LOI_SHAPE_SHADOW_LOCATION_* as defined in LibreOfficeImpress_Constants.au3.
;                  $iColor              - [optional] (0-16777215) Default is Null. The Shadow color, as a RGB Color Integer. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3.
;                  $iDistance           - [optional] Default is Null. The distance of the Shadow from the Shape's edges, set in Hundredths of a Millimeter (HMM).
;                  $iBlur               - [optional] (0-150) Default is Null. The amount of blur applied to the Shadow, set in Printer's Points.
;                  $iTransparency       - [optional] (0-100) Default is Null. The percentage of Shadow transparency. 100% means completely transparent.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 6 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTable not an Object.
;                  @Error: 1, @Extended: 2 = $bShadow not a Boolean.
;                  @Error: 1, @Extended: 3 = $iLocation not an Integer, less than 0 or greater than 8. See Constants, $LOI_SHAPE_SHADOW_LOCATION_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 4 = $iColor not an Integer, less than 0 or greater than 16777215.
;                  @Error: 1, @Extended: 5 = $iDistance not an Integer, or less than 0.
;                  @Error: 1, @Extended: 6 = $iBlur not an Integer, less than 0 or greater than 150 Printer's Points.
;                  @Error: 1, @Extended: 7 = $iTransparency not an Integer, less than 0 or greater than 100.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve current Distance and Location Values.
;                  @Error: 3, @Extended: 2 = Failed to modify Location property.
;                  @Error: 3, @Extended: 3 = Failed to modify Distance property.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bShadow
;                  |                               2 = Error setting $iLocation
;                  |                               4 = Error setting $iColor
;                  |                               8 = Error setting $iDistance
;                  |                               16 = Error setting $iBlur
;                  |                               32 = Error setting $iTransparency
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  LibreOffice may change the shadow distance +/- a Hundredth of a Millimeter (HMM).
; Related .......: _LO_ConvertColorFromLong, _LO_ConvertColorToLong, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_TableShadow(ByRef $oTable, $bShadow = Null, $iLocation = Null, $iColor = Null, $iDistance = Null, $iBlur = Null, $iTransparency = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oTable) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$vReturn = __LOImpress_ShapeAreaShadow($oTable, $bShadow, $iLocation, $iColor, $iDistance, $iBlur, $iTransparency)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOImpress_TableShadow
