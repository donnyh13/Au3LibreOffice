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
; Description ...: Provides basic functionality through AutoIt for inserting or manipulating Impress Fields.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOImpress_FieldAuthorInsert
; _LOImpress_FieldAuthorModify
; _LOImpress_FieldCurrentDisplayGet
; _LOImpress_FieldDateTimeInsert
; _LOImpress_FieldDateTimeModify
; _LOImpress_FieldDelete
; _LOImpress_FieldFileNameInsert
; _LOImpress_FieldFileNameModify
; _LOImpress_FieldGetAnchor
; _LOImpress_FieldHyperlinkInsert
; _LOImpress_FieldHyperlinkModify
; _LOImpress_FieldsGetList
; _LOImpress_FieldSlideCountInsert
; _LOImpress_FieldSlideNumberInsert
; _LOImpress_FieldSlideTitleInsert
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_FieldAuthorInsert
; Description ...: Insert an Author field.
; Syntax ........: _LOImpress_FieldAuthorInsert(ByRef $oDoc, ByRef $oTextCursor[, $bIsFixed = False[, $sAuthor = ""[, $iFormat = $LOI_FIELD_AUTH_NAME_FULL[, $bOverwrite = False]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $bIsFixed            - [optional] Default is False. If True, the field value is fixed at the time of insertion.
;                  $sAuthor             - [optional] Default is "". If $bIsFixed is True, the Author name to display.
;                  $iFormat             - [optional] (0-3) Default is $LOI_FIELD_AUTH_NAME_FULL. The format to display the Author. See Constants, $LOI_FIELD_AUTH_NAME_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bOverwrite          - [optional] Default is False. If True, any content selected by the Cursor is overwritten.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Successfully inserted the field, returning its Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 3 = $bIsFixed not a Boolean.
;                  @Error: 1, @Extended: 4 = $sAuthor not a String.
;                  @Error: 1, @Extended: 5 = $iFormat not an Integer, less than 0 or greater than 3. See Constants, $LOI_FIELD_AUTH_NAME_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 6 = $bOverwrite not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to Create a "com.sun.star.text.TextField.Author" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to identify and retrieve Field object after insertion.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Setting $iFormat while the field is fixed, seems to do nothing.
; Related .......: _LOImpress_FieldAuthorModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_FieldAuthorInsert(ByRef $oDoc, ByRef $oTextCursor, $bIsFixed = False, $sAuthor = "", $iFormat = $LOI_FIELD_AUTH_NAME_FULL, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextField, $oTextFieldReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sAuthor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not __LO_IntIsBetween($iFormat, $LOI_FIELD_AUTH_NAME_FULL, $LOI_FIELD_AUTH_NAME_INITIALS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oTextField = $oDoc.createInstance("com.sun.star.text.TextField.Author")
	If Not IsObj($oTextField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	With $oTextField
		.IsFixed = $bIsFixed
		If ($sAuthor <> "") Then .Content = $sAuthor
		.AuthorFormat = $iFormat
	EndWith

	$oTextCursor.Text.insertTextContent($oTextCursor, $oTextField, $bOverwrite)

	; Have to retrieve the Field's Object again, otherwise the Field Object seems invalid once inserted (Can't be used for modifying the field etc.).
	$oTextFieldReturn = __LOImpress_FieldGetObj($oTextCursor, $LOI_FIELD_TYPE_AUTHOR)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextFieldReturn)
EndFunc   ;==>_LOImpress_FieldAuthorInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_FieldAuthorModify
; Description ...: Set or Retrieve the settings of a Author field.
; Syntax ........: _LOImpress_FieldAuthorModify(ByRef $oAuthorField[, $bIsFixed = Null[, $sAuthor = Null[, $iFormat = Null]]])
; Parameters ....: $oAuthorField        - An Author Field Object returned by a previous _LOImpress_FieldFileNameInsert or _LOImpress_FieldsGetList function.
;                  $bIsFixed            - [optional] Default is Null. If True, the field value is fixed at the time of insertion.
;                  $sAuthor             - [optional] Default is Null. If $bIsFixed is True, the Author name to display.
;                  $iFormat             - [optional] (0-3) Default is Null. The format to display the Author name. See Constants, $LOI_FIELD_AUTH_NAME_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oAuthorField not an Object.
;                  @Error: 1, @Extended: 2 = $bIsFixed not a Boolean.
;                  @Error: 1, @Extended: 3 = $sAuthor not a String.
;                  @Error: 1, @Extended: 4 = $iFormat not an Integer, less than 0 or greater than 3. See Constants, $LOI_FIELD_AUTH_NAME_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bIsFixed
;                  |                               2 = Error setting $sAuthor
;                  |                               4 = Error setting $iFormat
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOImpress_FieldAuthorInsert
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LOImpress_FieldAuthorModify(ByRef $oAuthorField, $bIsFixed = Null, $sAuthor = Null, $iFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avField[3]

	If Not IsObj($oAuthorField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bIsFixed, $sAuthor, $iFormat) Then
		__LO_ArrayFill($avField, $oAuthorField.IsFixed(), $oAuthorField.Content(), $oAuthorField.AuthorFormat())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avField)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oAuthorField.IsFixed = $bIsFixed
		$iError = ($oAuthorField.IsFixed() = $bIsFixed) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oAuthorField.Content = $sAuthor
		$iError = ($oAuthorField.Content() = $sAuthor) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iFormat <> Null) Then
		If Not __LO_IntIsBetween($iFormat, $LOI_FIELD_AUTH_NAME_FULL, $LOI_FIELD_AUTH_NAME_INITIALS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oAuthorField.AuthorFormat = $iFormat
		$iError = ($oAuthorField.AuthorFormat() = $iFormat) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_FieldAuthorModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_FieldCurrentDisplayGet
; Description ...: Retrieve the current data displayed by a field.
; Syntax ........: _LOImpress_FieldCurrentDisplayGet(ByRef $oField)
; Parameters ....: $oField              - A Field Object as returned from a previous insert, or _LOImpress_FieldsGetList function.
; Return values .: Success: String
;                  @Error: 0, @Extended: 0, Return: String = Success. Returning current Field display content in String format.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oField not an Object.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a TextCursor.
;                  @Error: 2, @Extended: 2 = Failed to create enumeration of paragraphs.
;                  @Error: 2, @Extended: 3 = Failed to create enumeration of Text Portions in Paragraph.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to identify Field's Text portion object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve Field's current display.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Both Slide Title and Slide Number fields may return "<slide-name>" or "<number>" respectively instead of their current display value. I don't know why.
; Related .......: _LOImpress_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_FieldCurrentDisplayGet(ByRef $oField)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sPresentation
	Local $oTextCursor, $oParEnum, $oPar, $oTextEnum, $oTextPortion, $oFieldTextPortion

	If Not IsObj($oField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	; Generally the Field Object does not have a method to call to get the current displayed value.
	; In order to obtain the currently displayed value, you have to enumerate the containing shape's text Paragraphs and text portions,
	; until I find the text portion containing the field. Then I can just use the method getString to retrieve the currently displayed value of the field.
	$oTextCursor = $oField.Anchor.Text.createTextCursor()
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oParEnum = $oField.Anchor.getText().createEnumeration()
	If Not IsObj($oParEnum) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	While $oParEnum.hasMoreElements()
		$oPar = $oParEnum.nextElement()

		$oTextEnum = $oPar.createEnumeration()
		If Not IsObj($oTextEnum) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

		While $oTextEnum.hasMoreElements()
			$oTextPortion = $oTextEnum.nextElement()

			If ($oTextPortion.TextPortionType = "TextField") Then
				If ($oTextCursor.compareRegionEnds($oTextPortion.End(), $oField.Anchor.End()) = 0) Then
					$oFieldTextPortion = $oTextPortion
					ExitLoop
				EndIf
			EndIf
		WEnd
	WEnd

	If Not IsObj($oFieldTextPortion) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$sPresentation = $oFieldTextPortion.getString()
	If Not IsString($sPresentation) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sPresentation)
EndFunc   ;==>_LOImpress_FieldCurrentDisplayGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_FieldDateTimeInsert
; Description ...: Insert a Date or Time Field.
; Syntax ........: _LOImpress_FieldDateTimeInsert(ByRef $oDoc, ByRef $oTextCursor[, $bIsDate = True[, $bIsFixed = False[, $tDateTime = Null[, $iFormat = $LOI_FIELD_DATE_FMT_STANDARD_SHORT[, $bOverwrite = False]]]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $bIsDate             - [optional] Default is True. If True, the inserted Field will be a Date Field, if False, the Field will be a Time Field.
;                  $bIsFixed            - [optional] Default is False. If True, the field value is fixed at the time of insertion.
;                  $tDateTime           - [optional] Default is Null. If $bIsFixed is True, The date or time to display for the comment, created previously by _LOImpress_DateStructCreate. If left as Null, the current date or time is used.
;                  $iFormat             - [optional] (2-9) Default is $LOI_FIELD_DATE_FMT_STANDARD_SHORT. The format to display the date or time in. See Constants, $LOI_FIELD_TIME_FMT_* or $LOI_FIELD_DATE_FMT_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bOverwrite          - [optional] Default is False. If True, any content selected by the Cursor is overwritten.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Successfully inserted the field, returning its Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 3 = $bIsDate not a Boolean.
;                  @Error: 1, @Extended: 4 = $bIsFixed not a Boolean.
;                  @Error: 1, @Extended: 5 = $tDateTime not an Object.
;                  @Error: 1, @Extended: 6 = $bIsDate is True and $iFormat not an Integer, less than 2 or greater than 9. See Constants, $LOI_FIELD_DATE_FMT_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 7 = $bIsDate is False and $iFormat not an Integer, less than 2 or greater than 8. See Constants, $LOI_FIELD_TIME_FMT_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 8 = $bOverwrite not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to Create a "com.sun.star.text.TextField.DateTime" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to identify and retrieve Field object after insertion.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If the Field is a Time field, $LOI_FIELD_DATE_FMT_STANDARD_SHORT is the equivalent of $LOI_FIELD_TIME_FMT_STANDARD.
; Related .......: _LOImpress_FieldDateTimeModify, _LOImpress_FieldDelete, _LOImpress_DateStructCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_FieldDateTimeInsert(ByRef $oDoc, ByRef $oTextCursor, $bIsDate = True, $bIsFixed = False, $tDateTime = Null, $iFormat = $LOI_FIELD_DATE_FMT_STANDARD_SHORT, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextField, $oTextFieldReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bIsDate) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($tDateTime <> Null) And Not IsObj($tDateTime) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($bIsDate And Not __LO_IntIsBetween($iFormat, $LOI_FIELD_DATE_FMT_STANDARD_SHORT, $LOI_FIELD_DATE_FMT_DOW_MMMM_DD_YYYY)) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If Not $bIsDate And Not __LO_IntIsBetween($iFormat, $LOI_FIELD_TIME_FMT_STANDARD, $LOI_FIELD_TIME_FMT_12H_HMS_MS_AMPM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)

	$oTextField = $oDoc.createInstance("com.sun.star.text.TextField.DateTime")
	If Not IsObj($oTextField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	With $oTextField
		.IsDate = $bIsDate
		.IsFixed = $bIsFixed
		If IsObj($tDateTime) Then .DateTime = $tDateTime
		.NumberFormat = $iFormat
	EndWith

	$oTextCursor.Text.insertTextContent($oTextCursor, $oTextField, $bOverwrite)

	; Have to retrieve the Field's Object again, otherwise the Field Object seems invalid once inserted (Can't be used for modifying the field etc.).
	$oTextFieldReturn = __LOImpress_FieldGetObj($oTextCursor, $LOI_FIELD_TYPE_DATE_TIME)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextFieldReturn)
EndFunc   ;==>_LOImpress_FieldDateTimeInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_FieldDateTimeModify
; Description ...: Set or Retrieve the settings of a Date/Time field.
; Syntax ........: _LOImpress_FieldDateTimeModify(ByRef $oDateTimeField[, $bIsFixed = Null[, $tDateTime = Null[, $iFormat = Null]]])
; Parameters ....: $oDateTimeField      - A Date/Time Field Object returned by a previous _LOImpress_FieldDateTimeInsert or _LOImpress_FieldsGetList function.
;                  $bIsFixed            - [optional] Default is Null. If True, the field value is fixed at the time of insertion.
;                  $tDateTime           - [optional] Default is Null. If $bIsFixed is True, The date or time to display for the comment, created previously by _LOImpress_DateStructCreate. If left as Null, the current date or time is used.
;                  $iFormat             - [optional] (2-9) Default is Null. The format to display the date or time in. See Constants, $LOI_FIELD_TIME_FMT_* or $LOI_FIELD_DATE_FMT_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current Date Field settings in a 3 Element Array with values in order of function parameters. @Extended is set to 1.
;                  @Error: 0, @Extended: 2, Return: Array = Success. All optional parameters were called with Null, returning current Time Field settings in a 3 Element Array with values in order of function parameters. @Extended is set to 2.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDateTimeField not an Object.
;                  @Error: 1, @Extended: 2 = $bIsFixed not a Boolean.
;                  @Error: 1, @Extended: 3 = $tDateTime not an Object.
;                  @Error: 1, @Extended: 4 = Field is a Date and $iFormat not an Integer, less than 2 or greater than 9. See Constants, $LOI_FIELD_DATE_FMT_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 5 = Field is a Time and $iFormat not an Integer, less than 2 or greater than 8. See Constants, $LOI_FIELD_TIME_FMT_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bIsFixed
;                  |                               2 = Error setting $tDateTime
;                  |                               4 = Error setting $iFormat
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
;                  To retrieve the current date/time's values pass the returned structure from to function to _LOImpress_DateStructModify.
; Related .......: _LOImpress_FieldDateTimeInsert, _LOImpress_DateStructModify, _LOImpress_FieldCurrentDisplayGet
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LOImpress_FieldDateTimeModify(ByRef $oDateTimeField, $bIsFixed = Null, $tDateTime = Null, $iFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iExtended
	Local $avField[3]

	If Not IsObj($oDateTimeField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bIsFixed, $tDateTime, $iFormat) Then
		__LO_ArrayFill($avField, $oDateTimeField.IsFixed(), $oDateTimeField.DateTime(), $oDateTimeField.NumberFormat())

		$iExtended = ($oDateTimeField.IsDate() = True) ? (1) : (2) ; If the Field is a Date, set Extended to 1, else 2 if it is a Time Field.

		Return SetError($__LO_STATUS_SUCCESS, $iExtended, $avField)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oDateTimeField.IsFixed = $bIsFixed
		$iError = ($oDateTimeField.IsFixed() = $bIsFixed) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($tDateTime <> Null) Then
		If Not IsObj($tDateTime) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oDateTimeField.DateTime = $tDateTime
		; If not comparing a Date, I will be comparing a Time, so reverse the IsDate value.
		$iError = (__LOImpress_DateStructCompare($oDateTimeField.DateTime(), $tDateTime, $oDateTimeField.IsDate(), ($oDateTimeField.IsDate() = True) ? (False) : (True))) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iFormat <> Null) Then
		If $oDateTimeField.IsDate() Then
			If Not __LO_IntIsBetween($iFormat, $LOI_FIELD_DATE_FMT_STANDARD_SHORT, $LOI_FIELD_DATE_FMT_DOW_MMMM_DD_YYYY) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		Else
			If Not __LO_IntIsBetween($iFormat, $LOI_FIELD_TIME_FMT_STANDARD, $LOI_FIELD_TIME_FMT_12H_HMS_MS_AMPM) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		EndIf

		$oDateTimeField.NumberFormat = $iFormat
		$iError = ($oDateTimeField.NumberFormat() = $iFormat) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_FieldDateTimeModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_FieldDelete
; Description ...: Delete a Field from a Document.
; Syntax ........: _LOImpress_FieldDelete(ByRef $oField)
; Parameters ....: $oField              - A Field Object as returned from a previous insert, or _LOImpress_FieldsGetList function.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Successfully deleted the field.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oField not an Object.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a TextCursor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_FieldDelete(ByRef $oField)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCursor

	If Not IsObj($oField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	; For some reason the only way to delete a field in Impress is the create a TextCursor with the field selected, and then overwriting it with an empty string.
	; The normal method ($oField.Anchor.Text.removeTextContent($oField)), doesn't throw an error, but it also does nothing at all.
	$oCursor = $oField.Anchor.Text.createTextCursorByRange($oField.Anchor())
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oCursor.Text.insertString($oCursor, "", True)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOImpress_FieldDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_FieldFileNameInsert
; Description ...: Insert a File Name field.
; Syntax ........: _LOImpress_FieldFileNameInsert(ByRef $oDoc, ByRef $oTextCursor[, $bIsFixed = False[, $iFormat = $LOI_FIELD_FILENAME_FULL_PATH[, $bOverwrite = False]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $bIsFixed            - [optional] Default is False. If True, the field value is fixed at the time of insertion.
;                  $iFormat             - [optional] (0-3) Default is $LOI_FIELD_FILENAME_FULL_PATH. The format to display the File name/path. See Constants, $LOI_FIELD_FILENAME_* as defined in LibreOfficeImpress_Constants.au3.
;                  $bOverwrite          - [optional] Default is False. If True, any content selected by the Cursor is overwritten.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Successfully inserted the field, returning its Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 3 = $bIsFixed not a Boolean.
;                  @Error: 1, @Extended: 4 = $iFormat not an Integer, less than 0 or greater than 3. See Constants, $LOI_FIELD_FILENAME_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 5 = $bOverwrite not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to Create a "com.sun.star.text.TextField.FileName" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to identify and retrieve Field object after insertion.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_FieldSlideTitleInsert, _LOImpress_FieldFileNameModify, _LOImpress_FieldDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_FieldFileNameInsert(ByRef $oDoc, ByRef $oTextCursor, $bIsFixed = False, $iFormat = $LOI_FIELD_FILENAME_FULL_PATH, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextField, $oTextFieldReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LO_IntIsBetween($iFormat, $LOI_FIELD_FILENAME_FULL_PATH, $LOI_FIELD_FILENAME_NAME_AND_EXT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oTextField = $oDoc.createInstance("com.sun.star.text.TextField.FileName")
	If Not IsObj($oTextField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oTextField.IsFixed = $bIsFixed
	$oTextField.FileFormat = $iFormat

	$oTextCursor.Text.insertTextContent($oTextCursor, $oTextField, $bOverwrite)

	; Have to retrieve the Field's Object again, otherwise the Field Object seems invalid once inserted (Can't be used for modifying the field etc.).
	$oTextFieldReturn = __LOImpress_FieldGetObj($oTextCursor, $LOI_FIELD_TYPE_FILE_NAME)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextFieldReturn)
EndFunc   ;==>_LOImpress_FieldFileNameInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_FieldFileNameModify
; Description ...: Set or Retrieve the settings of a File Name field.
; Syntax ........: _LOImpress_FieldFileNameModify(ByRef $oFileNameField[, $bIsFixed = Null[, $iFormat = Null]])
; Parameters ....: $oFileNameField      - A File Name Field Object returned by a previous _LOImpress_FieldFileNameInsert or _LOImpress_FieldsGetList function.
;                  $bIsFixed            - [optional] Default is Null. If True, the field value is fixed at the time of insertion.
;                  $iFormat             - [optional] (0-3) Default is Null. The format to display the File name/path. See Constants, $LOI_FIELD_FILENAME_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oFileNameField not an Object.
;                  @Error: 1, @Extended: 2 = $bIsFixed not a Boolean.
;                  @Error: 1, @Extended: 3 = $iFormat not an Integer, less than 0 or greater than 3. See Constants, $LOI_FIELD_FILENAME_* as defined in LibreOfficeImpress_Constants.au3.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bIsFixed
;                  |                               2 = Error setting $iFormat
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOImpress_FieldFileNameInsert
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _LOImpress_FieldFileNameModify(ByRef $oFileNameField, $bIsFixed = Null, $iFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avField[2]

	If Not IsObj($oFileNameField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($bIsFixed, $iFormat) Then
		__LO_ArrayFill($avField, $oFileNameField.IsFixed(), $oFileNameField.FileFormat())

		Return SetError($__LO_STATUS_SUCCESS, 1, $avField)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oFileNameField.IsFixed = $bIsFixed
		$iError = ($oFileNameField.IsDate() = $bIsFixed) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iFormat <> Null) Then
		If Not __LO_IntIsBetween($iFormat, $LOI_FIELD_FILENAME_FULL_PATH, $LOI_FIELD_FILENAME_NAME_AND_EXT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oFileNameField.FileFormat = $iFormat
		$iError = ($oFileNameField.FileFormat() = $iFormat) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_FieldFileNameModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_FieldGetAnchor
; Description ...: Retrieve the Anchor Cursor Object for a Field.
; Syntax ........: _LOImpress_FieldGetAnchor(ByRef $oField)
; Parameters ....: $oField              - A Field Object as returned from a previous insert, or _LOImpress_FieldsGetList function.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Returning requested Field Anchor Cursor Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oField not an Object.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to retrieve Field anchor Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_FieldsGetList, _LOImpress_CursorInsertString, _LOImpress_CursorMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_FieldGetAnchor(ByRef $oField)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFieldAnchor

	If Not IsObj($oField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oFieldAnchor = $oField.Anchor.Text.createTextCursorByRange($oField.Anchor())
	If Not IsObj($oFieldAnchor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oFieldAnchor)
EndFunc   ;==>_LOImpress_FieldGetAnchor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_FieldHyperlinkInsert
; Description ...: Insert a Hyperlink field.
; Syntax ........: _LOImpress_FieldHyperlinkInsert(ByRef $oDoc, ByRef $oTextCursor, $sURL[, $sText = ""[, $sTargetFrame = ""[, $bOverwrite = False]]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $sURL                - The URL/Hyperlink Address.
;                  $sText               - [optional] Default is "". The Text to display instead of the URL. "" means the URL itself is displayed.
;                  $sTargetFrame        - [optional] Default is "". Enter the name of the frame that you want the linked file to open in. Pass an empty string to skip.
;                  $bOverwrite          - [optional] Default is False. If True, any content selected by the Cursor is overwritten.
; Return values .: Success: Map
;                  @Error: 0, @Extended: 0, Return: Object = Success. Successfully inserted the field, returning its Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 3 = $sURL not a String.
;                  @Error: 1, @Extended: 4 = $sText not a String.
;                  @Error: 1, @Extended: 5 = $sTargetFrame not a String.
;                  @Error: 1, @Extended: 6 = $bOverwrite not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to Create a "com.sun.star.text.TextField.URL" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to identify and retrieve Field object after insertion.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_FieldHyperlinkModify, _LOImpress_FieldDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_FieldHyperlinkInsert(ByRef $oDoc, ByRef $oTextCursor, $sURL, $sText = "", $sTargetFrame = "", $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextField, $oTextFieldReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsString($sTargetFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oTextField = $oDoc.createInstance("com.sun.star.text.TextField.URL")
	If Not IsObj($oTextField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	With $oTextField
		.URL = $sURL
		.Representation = $sText
		.TargetFrame = $sTargetFrame
	EndWith

	$oTextCursor.Text.insertTextContent($oTextCursor, $oTextField, $bOverwrite)

	; Have to retrieve the Field's Object again, otherwise the Field Object seems invalid once inserted (Can't be used for modifying the field etc.).
	$oTextFieldReturn = __LOImpress_FieldGetObj($oTextCursor, $LOI_FIELD_TYPE_URL)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextFieldReturn)
EndFunc   ;==>_LOImpress_FieldHyperlinkInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_FieldHyperlinkModify
; Description ...: Set or Retrieve the settings of a Hyperlink/URL field.
; Syntax ........: _LOImpress_FieldHyperlinkModify(ByRef $oHyperlinkField[, $sURL = Null[, $sText = Null[, $sTargetFrame = Null]]])
; Parameters ....: $oHyperlinkField     - A Hyperlink/URL Field Object returned by a previous _LOImpress_FieldHyperlinkInsert or _LOImpress_FieldsGetList function.
;                  $sURL                - [optional] Default is Null. The URL/Hyperlink Address.
;                  $sText               - [optional] Default is Null. The Text to display instead of the URL. "" means the URL itself is displayed.
;                  $sTargetFrame        - [optional] Default is Null. Enter the name of the frame that you want the linked file to open in. Pass an empty string to skip.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 3 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oHyperlinkField not an Object.
;                  @Error: 1, @Extended: 2 = $sURL not a String.
;                  @Error: 1, @Extended: 3 = $sText not a String.
;                  @Error: 1, @Extended: 4 = $sTargetFrame not a String.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sURL
;                  |                               2 = Error setting $sText
;                  |                               4 = Error setting $sTargetFrame
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To retrieve the current value(s): Omit all optional parameters, or pass Null for each parameter.
;                  To skip parameters: Pass the Null keyword to any optional parameter.
; Related .......: _LOImpress_FieldHyperlinkInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_FieldHyperlinkModify(ByRef $oHyperlinkField, $sURL = Null, $sText = Null, $sTargetFrame = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $asField[3]

	If Not IsObj($oHyperlinkField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LO_VarsAreNull($sURL, $sText, $sTargetFrame) Then
		__LO_ArrayFill($asField, $oHyperlinkField.URL(), $oHyperlinkField.Representation(), $oHyperlinkField.TargetFrame())

		Return SetError($__LO_STATUS_SUCCESS, 1, $asField)
	EndIf

	If ($sURL <> Null) Then
		If Not IsString($sURL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

		$oHyperlinkField.URL = $sURL
		$iError = ($oHyperlinkField.URL() = $sURL) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sText <> Null) Then
		If Not IsString($sText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oHyperlinkField.Representation = $sText
		$iError = ($oHyperlinkField.Representation() = $sText) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($sTargetFrame <> Null) Then
		If Not IsString($sTargetFrame) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oHyperlinkField.TargetFrame = $sTargetFrame
		$iError = ($oHyperlinkField.TargetFrame() = $sTargetFrame) ? ($iError) : (BitOR($iError, 4))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOImpress_FieldHyperlinkModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_FieldsGetList
; Description ...: Retrieve an Array of Field Objects present in a Shape.
; Syntax ........: _LOImpress_FieldsGetList(ByRef $oTextCursor[, $iType = $LOI_FIELD_TYPE_ALL[, $bFieldTypeNum = True]])
; Parameters ....: $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $iType               - [optional] (1-127) Default is $LOI_FIELD_TYPE_ALL. The type of Field to search for. See Constants, $LOI_FIELD_TYPE_* as defined in LibreOfficeImpress_Constants.au3. Can be BitOr'd together.
;                  $bFieldTypeNum       - [optional] Default is True. If True, adds a column to the array that has the Field Type Constant Integer for that particular Field, to assist in identifying the Field type. See Constants, $LOI_FIELD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
; Return values .: Success: Array
;                  @Error: 0, @Extended: ?, Return: Array = Success. Returning Array of Text Field Objects with @Extended set to number of results. See Remarks for Array sizing.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 2 = $iType not an Integer, less than 1 or greater than 127. (The total of all Constants added together.) See Constants, $LOI_FIELD_TYPE_* as defined in LibreOfficeImpress_Constants.au3.
;                  @Error: 1, @Extended: 3 = $bFieldTypeNum not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create a TextCursor.
;                  @Error: 2, @Extended: 2 = Failed to create enumeration of paragraphs.
;                  @Error: 2, @Extended: 3 = Failed to create enumeration of Text Portions in Paragraph.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve parent slide Object.
;                  @Error: 3, @Extended: 2 = Failed to retrieve containing Shape Object.
;                  @Error: 3, @Extended: 3 = Failed to identify requested Field Types.
;                  @Error: 3, @Extended: 4 = Failed to retrieve Text Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Array can vary in the number of columns, if $bFieldTypeNum is called with False, the Array will be a single column. If $bFieldTypeNum is called with True, a column will be added to the array. First column will always be the Field's Object.
;                  Setting $bFieldTypeNum to True will add a Field type Number column, matching the constants, $LOI_FIELD_TYPE_* as defined in LibreOfficeImpress_Constants.au3 for the found Field.
;                  This function may fail to identify Fields if text has been inserted recently using the same Cursor.
; Related .......: _LOImpress_FieldDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_FieldsGetList(ByRef $oTextCursor, $iType = $LOI_FIELD_TYPE_ALL, $bFieldTypeNum = True)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avFieldTypes[0][0]
	Local $oParEnum, $oPar, $oTextEnum, $oTextPortion, $oTextField, $oInternalCursor, $oDrawPage, $oShape
	Local $iCount = 0
	Local $avTextFields[1]

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LO_IntIsBetween($iType, $LOI_FIELD_TYPE_AUTHOR, $LOI_FIELD_TYPE_ALL) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bFieldTypeNum) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	; When a Text Cursor has been used to insert Strings previous to inserting or looking for a Field, the fields sometimes are not able to be identified.
	; The workaround I figured out was to create the Text Cursor again before enumerating the fields.
	; To do this I have to retrieve the shape Object again, then create a textcursor using the new Object. The parent of the shape is the drawpage (Slide), I
	; then cycle through all shapes in the slide to identify which one the current textcursor is in. Once found, I create a new cursor.
	$oDrawPage = $oTextCursor.Text.getParent()
	If Not IsObj($oDrawPage) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	For $i = 0 To $oDrawPage.Count() - 1
		$oShape = $oDrawPage.getByIndex($i)
		If Not IsObj($oShape) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		If ($oShape.Text() = $oTextCursor.Text()) Then
			$oInternalCursor = $oShape.Text.createTextCursorByRange($oTextCursor)
			ExitLoop
		EndIf

		Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
	Next

	If Not IsObj($oInternalCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$avFieldTypes = __LOImpress_FieldTypeServices($iType)
	If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	If $bFieldTypeNum Then ReDim $avTextFields[1][2]

	$oParEnum = $oInternalCursor.getText().createEnumeration()
	If Not IsObj($oParEnum) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	While $oParEnum.hasMoreElements()
		$oPar = $oParEnum.nextElement()

		$oTextEnum = $oPar.createEnumeration()
		If Not IsObj($oTextEnum) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

		While $oTextEnum.hasMoreElements()
			$oTextPortion = $oTextEnum.nextElement()

			If ($oTextPortion.TextPortionType = "TextField") Then
				$oTextField = $oTextPortion.TextField()
				If Not IsObj($oTextField) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

				For $i = 0 To UBound($avFieldTypes) - 1
					If $oTextField.supportsService($avFieldTypes[$i][1]) Then
						If $bFieldTypeNum Then
							$avTextFields[$iCount][0] = $oTextField
							$avTextFields[$iCount][1] = $avFieldTypes[$i][0]
							$iCount += 1
							If ($iCount = UBound($avTextFields)) Then ReDim $avTextFields[$iCount * 2][2]

						Else
							$avTextFields[$iCount] = $oTextField
							$iCount += 1
							If ($iCount = UBound($avTextFields)) Then ReDim $avTextFields[$iCount * 2]
						EndIf

						ExitLoop
					EndIf
					Sleep((IsInt($i / $__LOICONST_SLEEP_DIV) ? (10) : (0)))
				Next
			EndIf
		WEnd
	WEnd

	If $bFieldTypeNum Then
		ReDim $avTextFields[$iCount][2]

	Else
		ReDim $avTextFields[$iCount]
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $avTextFields)
EndFunc   ;==>_LOImpress_FieldsGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_FieldSlideCountInsert
; Description ...: Insert a total Slide Count Field.
; Syntax ........: _LOImpress_FieldSlideCountInsert(ByRef $oDoc, ByRef $oTextCursor[, $bOverwrite = False])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $bOverwrite          - [optional] Default is False. If True, any content selected by the Cursor is overwritten.
; Return values .: Success: Map
;                  @Error: 0, @Extended: 0, Return: Object = Success. Successfully inserted the field, returning its Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 3 = $bOverwrite not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to Create a "com.sun.star.text.TextField.PageCount" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to identify and retrieve Field object after insertion.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_FieldSlideNumberInsert, _LOImpress_FieldDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_FieldSlideCountInsert(ByRef $oDoc, ByRef $oTextCursor, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextField, $oTextFieldReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oTextField = $oDoc.createInstance("com.sun.star.text.TextField.PageCount")
	If Not IsObj($oTextField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oTextCursor.Text.insertTextContent($oTextCursor, $oTextField, $bOverwrite)

	; Have to retrieve the Field's Object again, otherwise the Field Object seems invalid once inserted (Can't be used for modifying the field etc.).
	$oTextFieldReturn = __LOImpress_FieldGetObj($oTextCursor, $LOI_FIELD_TYPE_SLIDE_COUNT)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextFieldReturn)
EndFunc   ;==>_LOImpress_FieldSlideCountInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_FieldSlideNumberInsert
; Description ...: Insert a Slide Number Field.
; Syntax ........: _LOImpress_FieldSlideNumberInsert(ByRef $oDoc, ByRef $oTextCursor[, $bOverwrite = False])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $bOverwrite          - [optional] Default is False. If True, any content selected by the Cursor is overwritten.
; Return values .: Success: Map
;                  @Error: 0, @Extended: 0, Return: Object = Success. Successfully inserted the field, returning its Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 3 = $bOverwrite not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to Create a "com.sun.star.text.TextField.PageNumber" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to identify and retrieve Field object after insertion.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_FieldSlideCountInsert, _LOImpress_FieldDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_FieldSlideNumberInsert(ByRef $oDoc, ByRef $oTextCursor, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextField, $oTextFieldReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oTextField = $oDoc.createInstance("com.sun.star.text.TextField.PageNumber")
	If Not IsObj($oTextField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oTextCursor.Text.insertTextContent($oTextCursor, $oTextField, $bOverwrite)

	; Have to retrieve the Field's Object again, otherwise the Field Object seems invalid once inserted (Can't be used for modifying the field etc.).
	$oTextFieldReturn = __LOImpress_FieldGetObj($oTextCursor, $LOI_FIELD_TYPE_SLIDE_NUM)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextFieldReturn)
EndFunc   ;==>_LOImpress_FieldSlideNumberInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOImpress_FieldSlideTitleInsert
; Description ...: Insert a Slide Title Field.
; Syntax ........: _LOImpress_FieldSlideTitleInsert(ByRef $oDoc, ByRef $oTextCursor[, $bOverwrite = False])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOImpress_DocOpen, _LOImpress_DocConnect, or _LOImpress_DocCreate function.
;                  $oTextCursor         - A Text Cursor Object returned by a previous _LOImpress_ShapeCreateTextCursor function.
;                  $bOverwrite          - [optional] Default is False. If True, any content selected by the Cursor is overwritten.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Successfully inserted the field, returning its Object.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oTextCursor not an Object.
;                  @Error: 1, @Extended: 3 = $bOverwrite not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to Create a "com.sun.star.text.TextField.PageName" Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to identify and retrieve Field object after insertion.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOImpress_FieldFileNameInsert, _LOImpress_FieldAuthorInsert, _LOImpress_FieldDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOImpress_FieldSlideTitleInsert(ByRef $oDoc, ByRef $oTextCursor, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOImpress_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextField, $oTextFieldReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oTextField = $oDoc.createInstance("com.sun.star.text.TextField.PageName")
	If Not IsObj($oTextField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oTextCursor.Text.insertTextContent($oTextCursor, $oTextField, $bOverwrite)

	; Have to retrieve the Field's Object again, otherwise the Field Object seems invalid once inserted (Can't be used for modifying the field etc.).
	$oTextFieldReturn = __LOImpress_FieldGetObj($oTextCursor, $LOI_FIELD_TYPE_SLIDE_TITLE)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTextFieldReturn)
EndFunc   ;==>_LOImpress_FieldSlideTitleInsert
